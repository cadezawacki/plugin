# Domain 5: Direct File I/O — Round 2 Production Audit

File audited: `/home/user/plugin/src/ExcelPerfToolkit/DirectFileIO.cs` (457 lines).

Round 2 went deeper into the CSV parser state machine, chunk-boundary semantics, encoding plumbing, resource lifetime, and UDF bridging. Round 1's two confirmed bugs (lastWasCR not reset on some branches; quotes mid-unquoted-field silently entering quote mode) are re-examined here and expanded into a complete enumeration of every branch where `lastWasCR` is mishandled, with concrete failing inputs.

## CSV Parser State Transition Table

For each input branch (top-level conditions in the `for (var i ...)` loop body, lines 93–158), the columns describe what changes:

| # | Branch (preconditions)                                       | Code lines | inQuotes after | lastWasCR after                          | fieldBuffer effect          | rowBuffer effect            | Per-RFC-4180 correctness |
|---|--------------------------------------------------------------|------------|----------------|------------------------------------------|-----------------------------|-----------------------------|--------------------------|
| 1 | inQuotes; ch=='"'; i+1<read; next=='"'                       | 101–108    | true           | unchanged                                | append `"`                  | none                        | OK                       |
| 2 | inQuotes; ch=='"'; i+1<read; next!='"'                       | 109–110    | **false**      | unchanged                                | none                        | none                        | OK (close-quote)         |
| 3 | inQuotes; ch=='"'; i+1==read; reader.Peek()=='"'             | 115–121    | true           | unchanged                                | append `"`                  | none                        | OK; sync Peek+Read used  |
| 4 | inQuotes; ch=='"'; i+1==read; Peek!='"' (incl. -1 EOF)       | 115,122    | **false**      | unchanged                                | none                        | none                        | OK                       |
| 5 | inQuotes; ch!='"'                                            | 125–126    | true           | unchanged (correct: CR inside quote is content) | append ch              | none                        | OK                       |
| 6 | !inQuotes; ch=='"' (opening quote)                           | 128–132    | **true**       | **NOT RESET — bug; see F1**              | none                        | none                        | OK iff fieldBuffer empty; otherwise see F2 |
| 7 | !inQuotes; ch==delimiter                                     | 133–137    | false          | **NOT RESET — bug; see F3**              | flushed → row               | append field text           | end-of-field              |
| 8 | !inQuotes; ch=='\r'                                          | 138–144    | false          | true (correct)                           | flushed → row               | flushed → rows              | line terminator           |
| 9 | !inQuotes; ch=='\n'; lastWasCR                               | 147–151    | false          | false (correct)                          | none                        | none                        | swallow LF-after-CR       |
| 10| !inQuotes; ch=='\n'; !lastWasCR                              | 145,152–154| false          | unchanged (was false; OK)                | flushed → row               | flushed → rows              | line terminator           |
| 11| !inQuotes; other char                                        | 156–157    | false          | false (set)                              | append ch                   | none                        | OK                        |

Branches where `lastWasCR` is *left untouched* and that follow a CR-terminated row: 2, 4, 6, 7. Branches 2 and 4 only fire while `inQuotes` is true — but `inQuotes` becomes true via branch 6, and branch 6 itself can fire while `lastWasCR` is still true from a previous CR. This is the root cause for findings F1 and F3.

## Test-input trace results

Below, each input is walked character by character; output column shows what the current parser actually emits.

| Input                              | Expected (per RFC 4180)              | Actual parser output                  | Correct? |
|------------------------------------|--------------------------------------|---------------------------------------|----------|
| `"a","b","c"\n`                    | `[["a","b","c"]]`                    | `[["a","b","c"]]`                     | yes      |
| `a,b\r\nc,d\r\n`                   | `[["a","b"],["c","d"]]`              | `[["a","b"],["c","d"]]`               | yes      |
| `a,b\rc,d\r`                       | `[["a","b"],["c","d"]]`              | `[["a","b"],["c","d"]]`               | yes      |
| `a,b\nc,d\n`                       | `[["a","b"],["c","d"]]`              | `[["a","b"],["c","d"]]`               | yes      |
| `\r\nfoo\n`                        | parser-specific; spec ambiguous       | `[[""],["foo"]]` (treats leading CR as empty row) | acceptable |
| `a,"b\nc",d\n`                     | `[["a","b\nc","d"]]`                 | `[["a","b\nc","d"]]`                  | yes      |
| `a,"b""c",d\n`                     | `[["a","b\"c","d"]]`                 | `[["a","b\"c","d"]]`                  | yes      |
| `a,"b\r\nc",d\r\n`                 | `[["a","b\r\nc","d"]]`               | `[["a","b\r\nc","d"]]`                | yes      |
| `a,b\r,c`                          | `[["a","b"],["","c"]]`               | `[["a","b"],["","c"]]`                | yes      |
| `"a"\r\n"b"\r\n`                   | `[["a"],["b"]]`                      | `[["a"],["b"]]`                       | yes      |
| `a\r"b"\nc\n` (CR-row, LF-row, F1) | `[["a"],["b"],["c"]]`                | `[["a"],["bc"]]`                      | **NO**   |
| `a\r"b"\n` (degenerate of above)   | `[["a"],["b"]]`                      | `[["a"],["b"]]` (by EOF luck)         | yes by luck |
| `a,b\r,\nc,d\n` (CR, then comma, then LF) | `[["a","b"],["",""],["c","d"]]` | `[["a","b"],["","c","d"]]`            | **NO**   |
| `a, "b" ,c\n` (round-1 confirmed)  | `[["a"," \"b\" ","c"]]` or quoted-stripped per impl | `[["a"," b ","c"]]` — quotes silently consumed | **NO** |

Two new failures emerge in round 2: F1 (`a\r"b"\nc\n`) and F3 (`a,b\r,\nc,d\n`). Both are direct consequences of branches 6 and 7 not clearing `lastWasCR`.

## Findings table

| ID  | Severity | Location (file:line)                          | Category                  | One-line failure                                                                 | Confidence |
|-----|----------|-----------------------------------------------|---------------------------|----------------------------------------------------------------------------------|------------|
| F1  | Critical | DirectFileIO.cs:128–132                        | parser/state-machine      | Opening quote does not reset lastWasCR; CR followed by a quoted field eats the next LF and concatenates fields | 0.97 |
| F2  | High     | DirectFileIO.cs:128–132                        | parser/correctness        | Quote in middle of unquoted field silently enters quote mode (round-1 confirmed) | 0.99 |
| F3  | High     | DirectFileIO.cs:133–137                        | parser/state-machine      | Delimiter branch does not reset lastWasCR; `\r,\n` sequence eats the LF and merges the empty row into the next | 0.93 |
| F4  | High     | DirectFileIO.cs:115–123                        | sync-over-async/blocking  | `reader.Peek()` and `reader.Read()` block the awaiting thread when the StreamReader buffer is dry, defeating async discipline | 0.85 |
| F5  | High     | DirectFileIO.cs:422–441                        | encoding/error-handling   | ResolveEncoding catches only `ArgumentException`; `NotSupportedException` from missing CodePagesEncodingProvider escapes for legacy codepages | 0.9 |
| F6  | High     | DirectFileIO.cs:74, 173                        | memory/unbounded          | `rows` list and final `object[rows.Count, maxCols]` array unbounded; large CSVs OOM the Excel process with no backpressure or row cap | 0.95 |
| F7  | Med      | DirectFileIO.cs:371–386 vs 237                 | encoding/defaults         | `WriteDelimitedUdf` default goes through `ResolveEncoding` which returns `Encoding.UTF8` (emits BOM), bypassing `WriteDelimitedAsync`'s `??=` no-BOM default. Docs say "no BOM"; users get a BOM. | 0.97 |
| F8  | Med      | DirectFileIO.cs:82, 162–166                    | parser/error-swallowing   | EOF reached while `inQuotes==true` is silently accepted: trailing field flushed as if quote had closed; no error surfaced | 0.95 |
| F9  | Med      | DirectFileIO.cs:404–420                        | input-validation          | `ResolveDelimiter` accepts `"`, `\r`, or `\n` as delimiter, breaking parser invariants (quote / line-terminator conflict) | 0.95 |
| F10 | Med      | DirectFileIO.cs:419                            | input-validation/unicode  | Delimiter parsed as `s[0]` — high surrogate of a supplementary char silently used as delimiter, mismatches any low surrogate paired char | 0.85 |
| F11 | Med      | DirectFileIO.cs:189                            | numeric-coercion          | `Marshaling.TryToDouble` likely matches "Infinity"/"NaN"/"-Infinity" under `NumberStyles.Float`, returning IEEE infinities/NaNs that Excel renders as `#NUM!` and that compare badly | 0.7 |
| F12 | Med      | DirectFileIO.cs:65–72, 89                      | data-corruption-edges     | `FileShare.ReadWrite \| FileShare.Delete` on read permits concurrent truncate/append/delete; produces torn reads, mis-parsed rows, or premature EOF mid-quote | 0.9 |
| F13 | Low      | DirectFileIO.cs:72                             | encoding/asymmetry        | `detectEncodingFromByteOrderMarks: true` silently overrides caller-supplied encoding when a BOM is present; caller's explicit "windows-1252" ignored if file starts with UTF-8 BOM | 0.95 |
| F14 | Low      | DirectFileIO.cs:71, 245                        | OS hint                   | `FileOptions.SequentialScan` is a read-only hint on Windows; passing it on a write FileStream is benign noise but indicates copy-paste, not deliberate tuning | 0.95 |
| F15 | Low      | DirectFileIO.cs:171–183                        | memory/copy               | Final two-pass copy from `List<object[]>` into `object[rows.Count, maxCols]` doubles peak memory for the duration of the copy; for a 100M-cell result this is hundreds of MB of duplicate retention | 0.9 |
| F16 | Low      | DirectFileIO.cs:352–356, 397–401              | error-handling/diagnostics | Generic `Exception` catch returns `#VALUE!` and traces only `ex.Message`; root cause (e.g., `IOException` with HRESULT, `PathTooLongException`, encoding fault) erased for the user, no stack trace traced | 0.95 |
| F17 | Low      | DirectFileIO.cs:404–420                        | parser/contract           | `ResolveDelimiter` accepts the literal two-character string `"\\t"` and maps to tab. Undocumented and brittle — a tab via Alt+0009 or a literal tab character work, but `\t` parsed by C# is two chars `\` + `t` and would only match the explicit escape | 0.85 |
| F18 | Low      | DirectFileIO.cs:268                            | api-surface               | `StreamWriter.WriteLineAsync(StringBuilder, CancellationToken)` exists in .NET 5+ and is correct; however the per-row `await` adds two state-machine allocations per row in the hot loop. For multi-million-row writes this is a measurable allocation tax. | 0.6 |
| F19 | Low      | DirectFileIO.cs:443–455                        | input-coercion            | `ResolveBool` accepts only `"1"` and case-insensitive `"TRUE"` as truthy strings; `"yes"`, `"true "`, `"1.0"` (double) silently fall through to default. Behavior may surprise users; not a bug per se | 0.9 |

## Per-finding detail

---

### F1 — Critical — Opening-quote branch does not reset `lastWasCR`; CR followed by a quoted field swallows the following LF and concatenates fields

Location: `DirectFileIO.cs:128–132`

```
if (ch == '"')
{
    inQuotes = true;
    continue;
}
```

Concrete failure scenario: A CSV produced by a Mac/Windows mixed pipeline that uses CRLF and contains quoted fields, e.g. `a\r"b"\nc\n`.

Evidence (full trace, starting state `inQuotes=false`, `lastWasCR=false`, `fieldBuffer=""`):

1. `'a'` — branch 11. fieldBuffer="a", lastWasCR=false.
2. `'\r'` — branch 8. PushField("a") → rowBuffer=["a"]. PushRow → rows=[["a"]], rowBuffer cleared. lastWasCR=**true**.
3. `'"'` — branch 6. inQuotes=true. **lastWasCR remains true.**
4. `'b'` — branch 5 (in quotes). fieldBuffer="b". inQuotes still true.
5. `'"'` — branch 2 (assuming non-boundary). inQuotes=false. **lastWasCR still true.**
6. `'\n'` — branch 9 (because lastWasCR is true). lastWasCR=false. **No PushField, no PushRow.**
7. `'c'` — branch 11. fieldBuffer="bc" (appended onto the un-flushed "b"). lastWasCR=false.
8. `'\n'` — branch 10. PushField("bc") → rowBuffer=["bc"]. PushRow → rows=[["a"],["bc"]].

Expected: `[["a"],["b"],["c"]]`. Actual: `[["a"],["bc"]]`. Silent data corruption with no error surfaced.

Proposed surgical fix: set `lastWasCR = false;` as the first statement of the opening-quote branch (between line 130 and 131). The branch already guarantees `ch == '"'` which is not CR, so resetting is correct.

Confidence: 0.97. The trace is mechanical and reproducible from any input shaped `<chars>\r"<quoted>"\n<more>`.

---

### F2 — High — Quote mid-unquoted-field silently enters quote mode (round-1 confirmed)

Location: `DirectFileIO.cs:128–132`

Concrete failure scenario: Field `a, "b" ,c` (with the spaces) — the leading space before `"` makes the field unquoted at the moment the `"` is consumed, so the parser is in "non-quoted field" state with fieldBuffer containing `" "`. The `"` toggles inQuotes=true without flagging an error and without escaping; subsequent content goes into quote mode.

Evidence (trace of `a, "b" ,c\n`):
1. `'a'` → fieldBuffer="a", lastWasCR=false.
2. `','` → PushField("a") → rowBuffer=["a"].
3. `' '` → branch 11. fieldBuffer=" ".
4. `'"'` → branch 6. inQuotes=true. The leading space is preserved in fieldBuffer.
5. `'b'` → fieldBuffer=" b".
6. `'"'` → branch 2 (next is space). inQuotes=false.
7. `' '` → fieldBuffer=" b ".
8. `','` → PushField(" b ") → rowBuffer=["a"," b "].
9. `'c','\n'` → standard.

Result: `[["a"," b ","c"]]`. The quotes are consumed but no doubled-quote escaping rule applies inside, so any `""` inside this "embedded" quote run would be interpreted as an escape that, under RFC 4180, should never apply when the field is not opened with `"` as its first char.

Proposed surgical fix: Only enter quote mode if `fieldBuffer.Length == 0`. Otherwise treat `"` literally (or raise a parse error). Two-line change at line 128:

```
if (ch == '"' && fieldBuffer.Length == 0)
{
    inQuotes = true;
    lastWasCR = false;
    continue;
}
// fall through to default; the '"' becomes a literal character of an unquoted field
```

Confidence: 0.99. Round-1 already confirmed.

---

### F3 — High — Delimiter branch does not reset `lastWasCR`; `\r,\n` swallows the LF and merges rows

Location: `DirectFileIO.cs:133–137`

```
if (ch == delimiter)
{
    PushField(fieldBuffer, rowBuffer, coerceNumeric);
    continue;
}
```

Concrete failure scenario: CSV with a stray CRLF *after* a CR-terminated row, e.g. `a,b\r,\nc,d\n` (which can be produced by a buggy upstream writer that emits CR followed by a row starting with the comma of an empty leading field followed by LF).

Evidence (trace):
1. `'a'`, `','`, `'b'` → rowBuffer=["a"], fieldBuffer="b".
2. `'\r'` → PushField("b"), PushRow → rows=[["a","b"]]. lastWasCR=true.
3. `','` → branch 7. PushField("") → rowBuffer=[""]. **lastWasCR stays true.**
4. `'\n'` → branch 9 (lastWasCR=true). Swallowed. lastWasCR=false. **No PushRow.**
5. `'c'` → fieldBuffer="c".
6. `','` → PushField("c") → rowBuffer=["","c"].
7. `'d','\n'` → PushField("d"), PushRow → rows=[["a","b"],["","c","d"]].

Expected per "CR is a row terminator, LF is a row terminator, CRLF is one row terminator": `[["a","b"],["",""],["c","d"]]`. Actual: `[["a","b"],["","c","d"]]`. The empty row vanished and its empty-field count is merged into row 3.

Proposed surgical fix: identical pattern as F1 — set `lastWasCR = false;` at the top of the delimiter branch (line 135). Same fix should be applied to branches 2, 4, and 6 (all non-CR branches that fire while we are physically outside a CR-terminated row).

A cleaner refactor: hoist `lastWasCR = (ch == '\r');` to the end of every branch — or, equivalently, set `lastWasCR = false;` at the top of the for-body and only re-set true in the CR branch. The current approach of trying to clear it lazily in each branch is exactly what creates these bugs.

Confidence: 0.93. The interleaving is rare in clean files but trivially reproducible with the input above.

---

### F4 — High — `reader.Peek()` / `reader.Read()` at chunk boundary block the awaiting thread

Location: `DirectFileIO.cs:115–123`

Code:
```
var next = reader.Peek();
if (next == '"')
{
    reader.Read();
    fieldBuffer.Append('"');
    continue;
}
```

Concrete failure scenario: A 64 KiB chunk of a CSV ends exactly on a `"` while `inQuotes==true`. The underlying `FileStream` buffer (64 KiB) is empty. `StreamReader.Peek()` requests one more decoded char from the stream — this synchronously calls `FileStream.Read(byte[], int, int)` under the hood, which on a `FileOptions.Asynchronous`-opened FileStream goes through a synchronous wrapper that **blocks the calling thread**. If the file is on a slow disk (network share, hung SMB, encrypted drive doing key unwrap), the thread blocks for seconds; if the disk is unresponsive, the thread blocks until the OS times out.

Because the surrounding code awaits `reader.ReadAsync(...)` everywhere else (line 88), the contract of "all I/O is asynchronous" stated in the file's preamble (lines 19–20) is broken at chunk boundaries. On a `TaskScheduler` with a small thread pool (Excel's internal pools, or any constrained host), this can stall other work, escalating to thread-pool starvation under load.

Evidence: `StreamReader.Peek()` source (.NET runtime) checks the internal char buffer; if empty, calls `ReadBuffer()` synchronously which calls `_stream.Read(...)`. There is no async overload of `Peek` and no way to make this respect cancellation either.

Proposed surgical fix: Treat the chunk-boundary `"` as "ambiguous, defer decision". Save a `pendingCloseQuote = true` flag and process the next char in the *next* outer-loop iteration after the next `ReadAsync` completes. On entry to the inner for-loop, if `pendingCloseQuote` is set, inspect `charBuffer[0]` (or treat EOF as close-quote). This removes the only two synchronous I/O calls in the read path.

Confidence: 0.85 — the *correctness* is fine; the production-grade concern is thread blocking on slow disks. Excel UDF wrappers run on worker threads, so a blocked thread doesn't lock Excel's UI, but pool starvation under concurrent calls is real.

---

### F5 — High — `ResolveEncoding` catches only `ArgumentException`; `NotSupportedException` for legacy codepages escapes

Location: `DirectFileIO.cs:422–441`

Code:
```
try
{
    return Encoding.GetEncoding(s);
}
catch (ArgumentException)
{
    return Encoding.UTF8;
}
```

Concrete failure scenario: User passes `"windows-1252"` (the very example named in the XML doc on line 329). On .NET 5+, the default encoding provider supports only Unicode encodings and ASCII; code pages require `Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)`. If `System.Text.Encoding.CodePages` is not added as a dependency and registered, `Encoding.GetEncoding("windows-1252")` throws `NotSupportedException("No data is available for encoding 1252...")`, **not** `ArgumentException`.

`NotSupportedException` is not caught by the `catch (ArgumentException)` filter, so it propagates up to `ReadDelimitedUdf` / `WriteDelimitedUdf` where the catch-all turns it into `#VALUE!`. The user sees a generic error for a documented argument value.

Evidence: .NET runtime — `EncodingTable.GetCodePageFromName` and friends throw `NotSupportedException` for code pages not registered; this is explicitly documented in the `Encoding.GetEncoding` remarks.

Proposed surgical fix: widen the catch to `(ArgumentException or NotSupportedException)` (C# 9 pattern):

```
catch (Exception ex) when (ex is ArgumentException || ex is NotSupportedException)
{
    return Encoding.UTF8;
}
```

Optionally call `Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)` once in `ToolkitLifetime` startup so the documented `windows-1252` example actually works.

Confidence: 0.9.

---

### F6 — High — Unbounded `rows` accumulation and final 2-D allocation OOMs Excel on large files

Location: `DirectFileIO.cs:74` (`new List<object[]>(capacity: 256)`) and `DirectFileIO.cs:173` (`new object[rows.Count, maxCols]`).

Concrete failure scenario: User points `EPT.READCSV` at a 4 GiB CSV with 60M rows × 8 cols. The parser:
1. Accumulates 60M `object[]` references in the rows list (~480 MB just for refs).
2. Accumulates 480M `object` references (~3.8 GB on 64-bit).
3. At line 173 allocates `new object[60_000_000, 8]` = 480M entries × 8 bytes = 3.84 GB *additional* — and Excel's process now holds both copies during the final copy loop (line 174–181) before the source list goes out of scope.

Excel will OOM long before the result reaches the caller. There is no row cap, no streaming-to-cells boundary, no warning. The XML doc on line 21 says "works incrementally over arbitrarily large files" — that is false for the API's actual contract (single allocated `object[,]`).

Worse: the final 2-D array allocation `new object[rows.Count, maxCols]` can throw `OverflowException` when `rows.Count * maxCols` exceeds `Array.MaxLength` (~2.1B). The exception is caught by the catch-all and returned as `#VALUE!` with no indication of size.

Proposed surgical fix: Add a configurable cap (e.g. 1M rows × 16K cols, matching Excel's worksheet limits which are 1,048,576 × 16,384) and short-circuit with a clear error if exceeded. Cap should be checked *during* parsing, not after:

```
if (rows.Count >= MaxRows) throw new InvalidDataException($"CSV exceeds {MaxRows} rows");
```

Confidence: 0.95. Trivially reproducible with any large CSV.

---

### F7 — Med — `WriteDelimitedUdf` writes a BOM despite docs saying "no BOM"

Location: `DirectFileIO.cs:371–386` (UDF) and `DirectFileIO.cs:237` (async core).

Concrete failure scenario: User calls `=EPT.WRITECSV("out.csv", A1:Z1000)` with no `encoding` argument. The UDF wrapper does:

```
var enc = ResolveEncoding(encoding);                            // line 380
WriteDelimitedAsync(path, block, delim, enc, "\r\n", ...);      // line 381
```

`ResolveEncoding` with a blank `encoding` argument returns `Encoding.UTF8` at line 426 — this is the *static* `Encoding.UTF8` singleton, which emits a BOM by default (`new UTF8Encoding(true)` semantics).

`WriteDelimitedAsync` line 237: `encoding ??= new UTF8Encoding(false);`. But `encoding` is already non-null (`Encoding.UTF8`), so the no-BOM fallback never fires.

Result: every CSV written by `EPT.WRITECSV` with no explicit encoding starts with the three bytes `EF BB BF`. The XML doc on line 375 says "Defaults to UTF-8 (no BOM)." Downstream tooling (legacy Excel, many shell scripts, AWK, sed) that doesn't strip BOMs gets a corrupted first cell of `﻿header1`.

Evidence: `Encoding.UTF8` returns the singleton whose `Preamble` is `[0xEF, 0xBB, 0xBF]`. `StreamWriter` writes the preamble on first write.

Proposed surgical fix: change `ResolveEncoding` line 426 to return the no-BOM instance:

```
if (Marshaling.IsBlankOrError(encoding))
{
    return new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
}
```

…or, better, cache a single static no-BOM `UTF8Encoding` instance and return that from both `ResolveEncoding` and `WriteDelimitedAsync`'s `??=`. Note this also affects `ReadDelimitedUdf`, but BOM-on-read is harmless because `detectEncodingFromByteOrderMarks: true` strips it.

Confidence: 0.97. Trivially testable by hex-dumping the output of a no-encoding call.

---

### F8 — Med — Unterminated quoted field at EOF silently flushed as if quote closed

Location: `DirectFileIO.cs:82` (`inQuotes` flag) and `DirectFileIO.cs:162–166` (trailing flush).

Concrete failure scenario: CSV file truncated mid-quote (e.g. a download that failed at byte 1.5 GB of a 2 GB file). Last bytes are `,"some text with no clos`. Parser ends with `inQuotes==true` and fieldBuffer containing the partial text.

The trailing-flush block at lines 162–166 checks `fieldBuffer.Length > 0 || rowBuffer.Count > 0` but **ignores `inQuotes`**:

```
if (fieldBuffer.Length > 0 || rowBuffer.Count > 0)
{
    PushField(fieldBuffer, rowBuffer, coerceNumeric);
    PushRow(rowBuffer, rows, ref maxCols);
}
```

The partial text is pushed as if the quote had closed normally; no error surfaces. The user sees a normal-looking result with a truncated last cell and no warning that the file was corrupt.

For a finance pipeline (this is an Excel add-in) silent data corruption on a truncated source file is a Critical-leaning Med — the only reason it's not Critical is that the truncated field is at least visibly truncated; structural row count is plausibly correct.

Proposed surgical fix: at the trailing flush:

```
if (inQuotes)
{
    throw new InvalidDataException("CSV ended inside a quoted field; file is truncated or malformed.");
}
```

Confidence: 0.95.

---

### F9 — Med — `ResolveDelimiter` does not validate against `"`, `\r`, or `\n`

Location: `DirectFileIO.cs:404–420`

Concrete failure scenario: User passes `=EPT.READCSV("a.csv", "\"")`. `ResolveDelimiter` returns `'"'`. In the parser:

- Line 128 (`if (ch == '"')`) is checked **before** line 133 (`if (ch == delimiter)`). So every `"` enters quote mode; the delimiter branch never fires; fields are never separated.

- Similarly with `'\r'` as delimiter: line 133 fires first and treats `\r` as field separator, never as line terminator; the CSV is parsed as a single line with thousands of fields.

Both produce silently wrong results, not errors.

Proposed surgical fix: in `ResolveDelimiter`, reject and throw `ArgumentException` for `'"'`, `'\r'`, `'\n'` after resolving:

```
var c = ...; // existing logic
if (c == '"' || c == '\r' || c == '\n')
    throw new ArgumentException("Delimiter cannot be quote or newline.", nameof(delimiter));
return c;
```

The UDF catch-all turns this into `#VALUE!`, which is the correct user-facing outcome.

Confidence: 0.95.

---

### F10 — Med — Delimiter `s[0]` slices off the high surrogate of an emoji or supplementary character

Location: `DirectFileIO.cs:419`

```
return s[0];
```

Concrete failure scenario: User passes a supplementary-plane character as delimiter (e.g. some unusual finance separator). `s` contains two UTF-16 code units (high + low surrogate). `s[0]` returns only the high surrogate (`\uD83D` for the emoji `😀`). The parser then matches char-by-char on the stream; if the source file contains any other supplementary character starting with the same high surrogate, those bytes are misidentified as the delimiter. Or, more likely, no character in the file matches `\uD83D` alone (because high surrogates only appear in valid pairs), so the delimiter never matches and the entire file becomes one field.

Proposed surgical fix: reject multi-code-unit delimiters explicitly:

```
if (s.Length != 1)
    throw new ArgumentException("Delimiter must be a single BMP character.", nameof(delimiter));
return s[0];
```

(The two-character `\\t` escape branch on line 415 is the only legitimate multi-char input, handled separately.)

Confidence: 0.85. Real risk only if users genuinely pass supplementary delimiters, which is rare.

---

### F11 — Med — Numeric coercion likely accepts `Infinity` / `NaN`

Location: `DirectFileIO.cs:189`

```
if (coerceNumeric && text.Length > 0 && Marshaling.TryToDouble(text, out var d))
{
    row.Add(d);
}
```

`Marshaling.TryToDouble` is out of scope, but the standard implementation pattern `double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d)` returns true for `"Infinity"`, `"-Infinity"`, `"NaN"`. These values become IEEE non-finite doubles, which `BulkTransfer.WriteBlock` will hand to Excel; Excel renders `+inf` and `NaN` as `#NUM!` cells, but they may equal `0` in some COM marshalers and produce silently wrong values when compared.

Worse: a field containing the literal string `"Infinity"` (which might be the name of a company, a band, a movie) is silently converted to a number when the user expected a string.

Evidence: `double.TryParse` accepts `"Infinity"` and `"NaN"` under `NumberStyles.Float` (the default).

Proposed surgical fix: after `TryToDouble`, reject non-finite results:

```
if (coerceNumeric && text.Length > 0 && Marshaling.TryToDouble(text, out var d) && !double.IsNaN(d) && !double.IsInfinity(d))
{
    row.Add(d);
}
else
{
    row.Add(text);
}
```

Confidence: 0.7. Requires verifying `TryToDouble`'s actual `NumberStyles` use; if it already excludes `AllowExponent` and the special tokens, the concern is moot.

---

### F12 — Med — `FileShare.ReadWrite | FileShare.Delete` permits torn reads mid-stream

Location: `DirectFileIO.cs:65–72`

```
new FileStream(
    path,
    FileMode.Open,
    FileAccess.Read,
    FileShare.ReadWrite | FileShare.Delete,
    ...);
```

Concrete failure scenario: User reads a CSV that is being actively appended to (a log file). With `FileShare.ReadWrite`, another writer can extend or truncate the file while we read. If the writer truncates between two of our `ReadAsync` calls, the second call may return 0 (EOF) prematurely — we report partial data. If the writer overwrites a region we've already read, we miss the update silently (acceptable for snapshot semantics) but if it overwrites a region we *haven't* yet read, we mix old and new bytes mid-quote and the parser misinterprets the stream.

`FileShare.Delete` permits another process to mark the file for deletion while we hold it. On Windows, the file isn't actually removed until the last handle closes — so this is mostly OK — but on Linux/.NET-Linux this is straightforward unlink, and any cached pages may or may not reflect post-unlink writes.

Proposed surgical fix: For read paths default to `FileShare.Read` (other readers OK, no writers). Expose a `shareMode` parameter for callers that explicitly want to read live-appended files; document the torn-read risk in that case.

Confidence: 0.9.

---

### F13 — Low — `detectEncodingFromByteOrderMarks: true` silently overrides explicit encoding

Location: `DirectFileIO.cs:72`

```
using var reader = new StreamReader(stream, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: DefaultBufferSize);
```

Concrete failure scenario: User passes `encoding: "windows-1252"` for a file that starts with `0xEF 0xBB 0xBF`. `StreamReader` ignores the supplied encoding and uses UTF-8. If the file is *actually* Windows-1252 with three leading garbage bytes that happen to match the UTF-8 BOM, mojibake follows.

This is an intentional StreamReader behavior and arguably correct — but the API contract should document it. Currently the XML doc at line 329 says "Optional encoding name" without noting that BOM detection overrides it.

Proposed surgical fix: document the behavior in the XML doc for both `ReadDelimitedAsync` and `ReadDelimitedUdf`. No code change required unless we want stricter behavior, in which case set `detectEncodingFromByteOrderMarks: false` when caller supplied a non-default encoding.

Confidence: 0.95.

---

### F14 — Low — `FileOptions.SequentialScan` on a write FileStream is a no-op

Location: `DirectFileIO.cs:245`

```
options: FileOptions.Asynchronous | FileOptions.SequentialScan);
```

`FILE_FLAG_SEQUENTIAL_SCAN` is documented by Microsoft as a hint that *reads* will be sequential. On a write-only handle it is ignored by the OS. The flag here is benign noise but signals copy-paste from the read path rather than deliberate tuning.

Proposed surgical fix: drop `FileOptions.SequentialScan` on line 245. Keep `FileOptions.Asynchronous`.

Confidence: 0.95.

---

### F15 — Low — Two-pass copy doubles peak memory during materialization

Location: `DirectFileIO.cs:171–183`

```
var result = new object[rows.Count, maxCols];
for (var r = 0; r < rows.Count; r++)
{
    var row = rows[r];
    for (var c = 0; c < maxCols; c++)
    {
        result[r, c] = c < row.Length ? row[c] : ExcelEmpty.Value;
    }
}
return result;
```

For the duration of the copy, both the `rows` list (with all its inner `object[]` rows) and the final `object[,]` are alive — peak memory is ~2x the result size. For a 100M-cell result this is hundreds of MB extra. After return, `rows` becomes garbage but GC may not reclaim immediately.

Proposed surgical fix: clear each inner row after copying it to free incrementally:

```
for (var r = 0; r < rows.Count; r++)
{
    var row = rows[r];
    for (var c = 0; c < maxCols; c++)
        result[r, c] = c < row.Length ? row[c] : ExcelEmpty.Value;
    rows[r] = null!; // allow GC of the inner array as we go
}
```

Confidence: 0.9. Real concern only at very large sizes; F6 (cap rows) addresses the worse case.

---

### F16 — Low — Generic catch returns `#VALUE!` with only `ex.Message`, erasing stack and HRESULT

Location: `DirectFileIO.cs:352–356`, `DirectFileIO.cs:397–401`

```
catch (Exception ex)
{
    TraceSource.TraceEvent(TraceEventType.Warning, 4, "EPT.READCSV failed: {0}", ex.Message);
    return ExcelError.ExcelErrorValue;
}
```

Concrete failure scenario: User reports `#VALUE!`. Operator looks at the trace and sees only "EPT.READCSV failed: Access to the path 'C:\foo' is denied.". The exception type, stack trace, inner exception, and HRESULT are lost. For diagnosing intermittent network-share faults (which surface as `IOException` with various NTSTATUS / HRESULT codes), this is the difference between a 5-minute fix and a 5-day investigation.

Proposed surgical fix: trace the full exception, not just `ex.Message`:

```
TraceSource.TraceEvent(TraceEventType.Warning, 4, "EPT.READCSV failed: {0}", ex);
```

`TraceEvent` with an `Exception` object calls `ToString()` which includes type, message, stack, and inner exceptions. Same for the WRITECSV path.

Confidence: 0.95.

---

### F17 — Low — `ResolveDelimiter`'s `"\\t"` parsing is brittle / undocumented

Location: `DirectFileIO.cs:415`

```
if (string.Equals(s, "\\t", StringComparison.Ordinal) || s == "\t")
{
    return '\t';
}
```

The XML doc on lines 328 and 374 reads `Pass "\t" for TSV`. In Excel formula syntax, `"\t"` is the two-character string `\` + `t`, not the C# escape `'\t'`. So users actually have to pass the literal two-character string `\t`, which works via the first leg of the OR. Users who paste a real tab character (`\t` from clipboard) also work via the second leg. Confusing but functional.

The brittleness: the code only matches exactly `"\\t"` (two chars). If a user types `"\\T"` (capital T) or `" \t"` (with surrounding whitespace) it falls through to `s[0]` which returns `'\\'`, silently using backslash as the delimiter.

Proposed surgical fix: case-insensitive match and trim, or document the exact expected string.

Confidence: 0.85.

---

### F18 — Low — Per-row `await WriteLineAsync` adds two state-machine allocations per row

Location: `DirectFileIO.cs:268`

```
await writer.WriteLineAsync(sb, cancellationToken).ConfigureAwait(false);
```

Each await creates an `AsyncStateMachineBox` (for the for-loop method) and a `ConfiguredTaskAwaitable.ConfiguredTaskAwaiter`. For a 10M-row write that's ~20M allocations on top of the StringBuilder churn. Most awaits complete synchronously (buffer not full), so the state machines are short-lived gen-0 garbage — but they still bump the allocator and GC.

Proposed surgical fix: optionally batch multiple rows into one big StringBuilder before awaiting, e.g. flush every 1024 rows. This trades latency for throughput; for the write path latency is irrelevant.

Confidence: 0.6 — this is a microbench-level concern, not a correctness bug.

---

### F19 — Low — `ResolveBool` accepts only `"TRUE"` / `"1"`

Location: `DirectFileIO.cs:443–455`

```
return value switch
{
    bool b => b,
    double d => d != 0d,
    string s => s.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || s == "1",
    _ => defaultValue,
};
```

Users typing `"true"` work (case-insensitive); `"yes"`, `"y"`, `"on"`, `"1.0"` do not and silently fall through to the default. Excel's own boolean coercion is more permissive. Document or extend.

Confidence: 0.9. Behavior decision rather than a bug.

---

## Rejected findings

The following candidates were considered and rejected because the concrete failure scenario could not be traced to a real interleaving or input.

- **R1** — *"`reader.Peek()` could return a multi-char surrogate and miss a `"`."* `Peek` returns one UTF-16 code unit; the `"` character is in the BMP. No issue.
- **R2** — *"`PushRow` after PushField with empty fieldBuffer creates phantom empty rows."* PushRow's `if (rowBuffer.Count == 0) return;` early-out at line 201–204 means a totally empty rowBuffer is silently skipped. Correct behavior; not a bug.
- **R3** — *"`fieldBuffer` is not cleared on entering quote mode and may retain content from a prior field."* Reviewed: PushField clears it on every field flush; the only path that reaches the opening-quote branch without an intervening flush is F2 (mid-field quote), which is filed separately.
- **R4** — *"`encoding ??= Encoding.UTF8` at line 63 races with another thread mutating the parameter."* Parameters are local; no race possible.
- **R5** — *"`TraceSource` may be null if `ToolkitLifetime.CreateTraceSource` returned null."* Out of scope (ToolkitLifetime is in another file). Note for cross-cutting review.
- **R6** — *"`block.GetLength(0)` * `block.GetLength(1)` overflow at line 252–253."* Both are int, no multiplication performed; iteration is per-dimension. No overflow.
- **R7** — *"Final array allocation `new object[rows.Count, maxCols]` at line 173 may overflow for `rows.Count * maxCols > int.MaxValue`."* This is real but folded into F6 since the underlying issue is unbounded accumulation; the surgical fix for F6 (cap rows) also prevents the overflow.
- **R8** — *"`reader.Read()` at line 118 might be called when Peek returned -1."* Inspecting the branch: Read() is only called inside `if (next == '"')`, so Peek returned `'"'` (34), not -1. Safe.
- **R9** — *"`await using var stream` followed by `using var reader` mixes async-disposable and disposable in a way that may not flush async writes on exception."* Read path is read-only; StreamReader's Dispose does not flush. Safe.
- **R10** — *"`WriteLineAsync(StringBuilder, CancellationToken)` doesn't exist in .NET 8."* Verified: introduced in .NET 5, present in .NET 8. Honors `NewLine` and respects the CancellationToken. No issue.
- **R11** — *"Sync-over-async at `ReadDelimitedUdf` line 337–340 deadlocks on Excel's main thread."* Audited every `await` inside `ReadDelimitedAsync` (line 88) and `WriteDelimitedAsync` (lines 258, 268, 270): all use `ConfigureAwait(false)`. Plus the top-level `.ConfigureAwait(false).GetAwaiter().GetResult()` does not capture any SynchronizationContext. UDFs run on worker threads (per the file's own doc lines 27–29). No deadlock path traceable.

Forbidden-territory observations (noted, not filed): `Marshaling.TryToDouble` and `Marshaling.ToStringSafe` semantics are referenced from this file but live in other files; concerns about `Infinity`/`NaN` acceptance are filed as F11 against the *call site* (which is in scope) rather than the implementation (which is not). `ToolkitLifetime.ShutdownToken` and `ToolkitLifetime.CreateTraceSource` similarly out of scope.
