# Domain 7 audit (round 3): Text / Regex / Series / Date / Distance block utilities

Scope: `TextUtilities.cs`, `RegexUtilities.cs`, `SeriesUtilities.cs`, `DateUtilities.cs`,
`DistanceUtilities.cs` (all under `/home/user/plugin/src/ExcelPerfToolkit/`).
Context read (no findings reported in): `Marshaling.cs`, `VectorizedKernels.cs`, `readme.md`,
`AddIn.cs`, generated `ExcelPerfToolkit-AddIn.dna`, `ExcelPerfToolkit.csproj`.

All "Empirically verified" items were reproduced on .NET 8.0.28 x64 (`/tmp/probe7`).
Registration context (verified from the build): the project references only
`ExcelDna.AddIn`/`ExcelDna.Integration` 1.8.0; the generated `.dna` is a plain
`<ExternalLibrary ... ExplicitExports="false">` — Excel-DNA's **built-in** registration is in
effect, with no `ExcelDna.Registration` parameter-conversion pipeline. Under built-in
registration, C# default parameter values are **not** applied; a missing argument arrives as
`default(T)` for value-typed parameters (`bool` → `false`, `int` → `0`, `double` → `0.0`) and
`null` for `string`.

## Summary table

| ID | Severity | One-line |
|----|----------|----------|
| BUG-HIGH-1 | HIGH | `EPT.TEMPLATEFILL` documented default `has_header_row = TRUE` never applies; omitting the argument yields `false` → headers rendered as data, all `{name}` tokens left literal |
| BUG-HIGH-2 | HIGH | Euclidean distance via `a·a + b·b − 2a·b` suffers catastrophic cancellation: distinct points at large magnitude silently return distance 0 (verified: true 1.0 → computed 0.0) |
| BUG-MEDIUM-3 | MEDIUM | `EPT.REPEAT` with `count ≥ 2^31` silently returns `""` for even-length cells (fp→int wraps to `int.MinValue`, defeats the 32767 guard) instead of the documented throw |
| BUG-MEDIUM-4 | MEDIUM | `EPT.REGEXEXTRACTALL` / `EPT.REGEXSPLIT` spill width is unbounded → multi-GB / >2^31-element `object[,]` allocation (OOM inside the Excel process) for patterns like `""` |
| BUG-MEDIUM-5 | MEDIUM | `EPT.WORKDAYADD` with huge `days`: out-of-int-range values hit `Math.Abs(int.MinValue)` → `OverflowException`; in-range huge values walk ~2.9M days then throw — either way the *whole call* fails instead of the contractual per-cell error |
| BUG-MEDIUM-6 | MEDIUM | `EPT.CAMELCASE` silently deletes astral-plane letters (CJK Ext-B ideographs, math alphanumerics, emoji) and mis-cases the following character (verified `"x𠀀y zz"` → `"xYZz"`) |
| BUG-LOW-7 | LOW | `EPT.REVERSE` emits ill-formed UTF-16 (reversed surrogate pairs / reordered combining marks), worse than the doc's "not preserved" |
| BUG-LOW-8 | LOW | Pad character `s[0]` takes the high surrogate of an astral pad char → cells padded with lone `U+D83D`-style units (ill-formed strings) |
| BUG-LOW-9 | LOW | `REGEXMATCH`/`REGEXCOUNT`/`REGEXEXTRACT` silently coerce input **error** cells to `FALSE`/`0`/`""`, masking upstream `#N/A`/`#REF!` |
| BUG-LOW-10 | LOW | Cosine self-distance diagonal returns ±1–2·10⁻¹⁶ instead of exactly 0 in ~47% of rows (verified `1 − s/(√s·√s) ≠ 0` for 470,569/1,000,000 samples) |
| BUG-LOW-11 | LOW | `EPT.DISTANCE` treats a *scalar error* passed as `matrix_b` as "omitted" → silently computes self-distance of `matrix_a` instead of propagating the error |
| BUG-LOW-12 | LOW | `EPT.TEMPLATEFILL` has no 32,767-char output cap; a rendered row exceeding Excel's cell limit is silently truncated/failed at the marshaling boundary, violating the file's per-cell-error discipline |
| MEM-1 | — | `REGEXCOUNT` allocates `Match`/`MatchCollection` per cell: measured 33.6 MB per 100k-cell call → 0 MB with `Regex.Count` |
| MEM-2 | — | `EPT.DISTANCE` with `matrix_b` omitted flattens the same matrix twice (duplicate `double[]` + duplicate per-cell parse) |
| MEM-3 | — | Text transforms allocate `StringBuilder` + buffer + string per cell: measured 18.2 MB → 6.8 MB (2.7×) with `string.Create` for `EPT.PROPER` over 100k cells |
| MEM-4 | — | `EPT.REVERSE` `ToCharArray` + `new string` doubles allocations: 13.7 MB → 6.8 MB (2.0×) per 100k cells |
| MEM-5 | — | `EPT.OUTLIERS`/`EPT.QUANTILES`: full-capacity `List<double>` + `ToArray` copy + `deviations` array ≈ 2.4 MB avoidable transient per 100k-cell call |
| PERF-1 | — | `WORKDAYADD` O(days) `DateTime`-walk: 16 ns/step measured; int-serial walk = 3.1×, closed-form week jumps ≈ 50× at year-scale offsets |
| PERF-2 | — | Per-call interpreted `Regex`: `RegexOptions.Compiled` measured 4.9× faster (212 ms → 44 ms per 100k cells); needs a cache to amortize the 15.6 ms compile |
| PERF-3 | — | `EPT.DISTANCE` self-comparison computes both triangles of a symmetric matrix → 1.95× wasted kernel work + double boxing |
| PERF-4 | — | Manhattan/Chebyshev inner loop is scalar with a per-element branch → ~3–4× available via `Vector.Abs`/`Vector.Max` + loop splitting |
| PERF-5 | — | `EPT.OUTLIERS` re-runs `TryToDouble` (string re-parse) on every cell in the flagging pass → ~1.5–2× whole-call on string-heavy blocks |
| ARCH-1 | — | Process-wide bounded compiled-Regex cache; the file's "static cache would violate MTR safety" premise is factually wrong |
| ARCH-2 | — | O(1) workday engine (cumulative weekday table + sorted holiday array) shared per call |
| ARCH-3 | — | Centralized spill-shape governance (grid limits enforced before allocation) for all spilling UDFs |
| ARCH-4 | — | Span/`string.Create`-based, Rune-aware text-transform layer replacing per-cell StringBuilder lambdas |

---

## Pass 1 — Deep bug scan

### BUG-HIGH-1
**File:** /home/user/plugin/src/ExcelPerfToolkit/TextUtilities.cs:172-216 (parameter at 176)
**Category:** logic_error
**Description:** `TemplateFill` declares `bool hasHeaderRow = true` and documents "TRUE (default) treats row 0 as header names". The build uses Excel-DNA 1.8.0 built-in registration (verified: no `ExcelDna.Registration` package; generated `.dna` is a plain `ExternalLibrary`), which ignores C# default parameter values: an omitted argument is `xltypeMissing` and is converted to `default(bool)` = **false**. The documented default can never occur.
**Trigger condition:** `=EPT.TEMPLATEFILL("{Region}: {Sales}", A1:B3)` with `A1:B3` = `{"Region","Sales"; "East",100; "West",200}`, third argument omitted (the documented common case).
**Trace:**
1. Excel passes `xltypeMissing` for arg 3 → Excel-DNA built-in wrapper converts to `false` (C# default `= true` is metadata only; no registration processing consults it).
2. Line 185 `if (hasHeaderRow)` → false branch (lines 199-203): `names = null`, `dataStart = 0`, `outRows = 3`.
3. Row 0 (the header row) is rendered as data: output row 0 = `RenderTemplate(...)` over `{"Region","Sales"}`.
4. In `TryResolveToken` (355-368): `int.TryParse("Region")` fails; `names is null` → returns false → line 337 appends the literal `{Region}`.
5. Result: 3 rows, every one the literal string `"{Region}: {Sales}"`. Expected: 2 rows `"East: 100"`, `"West: 200"`.
**Fix:** take the flag as `object` and decode the sentinel — the same pattern already used for `padChar` (line 69) and `holidays` (DateUtilities.cs:37):
```csharp
// before
[ExcelArgument(Name = "has_header_row", ...)] bool hasHeaderRow = true)
{
    ...
    if (hasHeaderRow)
// after
[ExcelArgument(Name = "has_header_row", ...)] object hasHeaderRow)
{
    var useHeader = hasHeaderRow switch
    {
        bool b => b,
        _ when Marshaling.IsBlankOrError(hasHeaderRow) => true,           // omitted -> documented default
        _ => !Marshaling.TryToDouble(hasHeaderRow, out var d) || d != 0d, // numeric truthiness
    };
    if (useHeader)
```
(Sweep note: all other optional parameters in the five files were checked and are safe under built-in registration: `string` defaults arrive as `null` and are normalized by `IsNullOrWhiteSpace`/`IsNullOrEmpty` (`FillForward`, `Outliers`, `WorkdayAdd`, `Distance`); `bool ignoreCase = false`, `int groupIndex = 0`, `double threshold = 0` coincide with `default(T)`.)

### BUG-HIGH-2
**File:** /home/user/plugin/src/ExcelPerfToolkit/DistanceUtilities.cs:65-79 (cancellation at 75-76)
**Category:** data_corruption (silent numeric corruption)
**Description:** Euclidean distance is computed as `sqrt(selfA + selfB − 2·dot)`. When `‖a−b‖ ≲ 1.5·10⁻⁸·‖a‖`, the subtraction cancels *all* significant bits and the clamp at line 76 turns the residue into 0: distinct vectors report distance exactly 0 (or off by ~2× in the partial-cancellation band). The direct-difference formulation has ulp-level error.
**Trigger condition:** Any data with large common magnitude and small differences — Unix-ms timestamps (~1.7·10¹²), UTM/state-plane coordinates (~10⁵–10⁷), account balances in cents. `=EPT.DISTANCE({100000000,0;100000001,0})`.
**Trace (empirically verified, .NET 8 x64):**
1. a = (1e8, 0), b = (1e8+1, 0). `selfA = 1e16`; `selfB = (1e8+1)² = 10000000200000001` → rounds to a double with ulp 2 at that scale; `dot = 1e16 + 1e8` (exactly representable).
2. Line 75: `sq = selfA[i] + selfB[j] − 2·dot` = `2.00000002e16 − 2.00000002e16` computed at ulp(2e16) = 4 → result **0** (true value 1).
3. Line 76: `Math.Sqrt(sq > 0 ? sq : 0)` → **0.0**. Direct formula: `sqrt((a₀−b₀)²) = 1.0`. Second case (1e8,1) vs (1e8,2) likewise: identity → 0.0, direct → 1.0.
4. The result block silently contains 0 where the true distance is 1 → dedup/nearest-neighbor logic built on the sheet collapses distinct rows.
**Fix:** keep the fast identity but detect cancellation and recompute that pair directly (the slow path only fires for near-duplicate pairs, preserving the SIMD win):
```csharp
// before
var sq = selfA[i] + selfB[j] - (2d * dot);
result[i, j] = Math.Sqrt(sq > 0d ? sq : 0d);
// after
var sq = selfA[i] + selfB[j] - (2d * dot);
if (sq < 1e-8 * (selfA[i] + selfB[j]))     // >~26 bits cancelled: identity result is noise
{
    var rowB = flatB.AsSpan(j * dim, dim);
    sq = 0d;
    for (var k = 0; k < dim; k++) { var diff = rowA[k] - rowB[k]; sq += diff * diff; }
}
result[i, j] = Math.Sqrt(sq > 0d ? sq : 0d);
```
(A SIMD `DiffDot` kernel would keep even the slow path vectorized; see ARCH-4 seam.)

### BUG-MEDIUM-3
**File:** /home/user/plugin/src/ExcelPerfToolkit/TextUtilities.cs:115-140 (conversion at 121, guard at 129, capacity at 133)
**Category:** data_corruption (silent wrong output) + logic_error (int overflow)
**Description:** `count` is validated only as "non-negative integer-valued double" (117-120) and then cast `var n = (int)count` (121). On .NET 8 x64 an out-of-`int`-range double converts to **`int.MinValue`** regardless of sign (empirically verified: `(int)4294967296d = -2147483648`). The guard `(long)n * s.Length > MaxCellChars` (129) then sees a large *negative* product and passes; the StringBuilder capacity `n * s.Length` (133, unchecked `int`) wraps to 0 for even-length strings; the loop `i < n` never runs → `""` is returned silently.
**Trigger condition:** `=EPT.REPEAT(A1:A10, B1)` where B1 holds any value ≥ 2,147,483,648 — e.g. a Unix-millisecond timestamp (1.7·10¹²) picked up from the wrong column. Documented behavior: throw → `#VALUE!`.
**Trace (empirically verified, full code path replayed):** `count = 4294967296` → passes validation (`valid=True`) → `n = -2147483648` → `(long)n*2 = -4294967296` → `limitThrows=False` → capacity `unchecked(int.MinValue*2) = 0` → loop skipped → result `""` (len 0), no exception. For odd `s.Length` the capacity is negative → `ArgumentOutOfRangeException` (verified) → whole-call `#VALUE!` — so a mixed block half-empties and half-errors.
**Fix:** any `count > 32767` must throw anyway for non-blank cells (`s.Length ≥ 1` ⇒ `n·len ≥ n`), so reject it before the cast:
```csharp
// before
if (double.IsNaN(count) || count < 0d || count != Math.Truncate(count))
// after
if (double.IsNaN(count) || count < 0d || count != Math.Truncate(count) || count > MaxCellChars)
```
(Also makes the per-cell limit check at 129-132 reachable only for in-range `n`, eliminating both overflow paths.)

### BUG-MEDIUM-4
**File:** /home/user/plugin/src/ExcelPerfToolkit/RegexUtilities.cs:263-283 (`Spill`), reached from 126-172 and 181-219
**Category:** memory_leak (unbounded transient allocation / OOM in host)
**Description:** Spill width = widest row's match/part count, with no cap. `new object[rows, maxCols]` (273) can demand billions of elements; the CLR throws `OutOfMemoryException` for >2³¹ total elements — an exception the codebase itself classifies as critical-do-not-swallow (`VectorizedKernels.IsCritical`), thrown here on an MTR recalc thread inside Excel. Sub-limit cases still allocate gigabytes of `""` padding. Excel additionally cannot spill wider than 16,384 columns, so everything beyond that is guaranteed-wasted memory.
**Trigger condition:** `=EPT.REGEXSPLIT(A1:A100000, "")` — empty pattern matches at every position; a 32,767-char cell splits into **32,769** parts (empirically verified). Same for `EPT.REGEXEXTRACTALL(...; pattern with zero-width match)`.
**Trace:**
1. 100,000 rows × one 32,767-char cell each → `rx.Split` (204) yields 32,769 parts per row; `perRow[r].Length = 32769`.
2. `Spill` (263): `maxCols = 32769`; line 273 `new object[100000, 32769]` → 3.28·10⁹ elements → exceeds CLR 2D-array element limit → `OutOfMemoryException` after the per-row arrays (already ~26 GB of string references + padding) have been built.
3. Even 10,000 rows: 3.3·10⁸ cells × (8 B ref + interned `""`) ≈ 2.6 GB of result block + per-row arrays before Excel rejects the >16,384-wide spill anyway.
**Fix:** enforce grid/CLR limits before allocating, in `Spill`:
```csharp
// before
var result = new object[rows, maxCols];
// after
const int MaxSpillCols = 16_384;                       // Excel grid width
if (maxCols > MaxSpillCols)
    throw new ArgumentException($"Result would be {maxCols} columns wide; Excel allows at most {MaxSpillCols}.");
if ((long)rows * maxCols > int.MaxValue)               // same guard as DistanceUtilities.cs:45
    throw new ArgumentException($"Result {rows}x{maxCols} exceeds Int32.MaxValue cells.");
var result = new object[rows, maxCols];
```

### BUG-MEDIUM-5
**File:** /home/user/plugin/src/ExcelPerfToolkit/DateUtilities.cs:106-121
**Category:** logic_error (contract violation + unbounded loop)
**Description:** Two compounding defects in `ComputeOne`. (a) `var remaining = (int)Math.Truncate(dayCountRaw)` (106): any |days| ≥ 2³¹ converts to `int.MinValue` on .NET 8 x64 (verified) → `step` becomes −1 even for positive inputs (111) → `Math.Abs(int.MinValue)` (112) throws `OverflowException` (verified). (b) For huge in-range counts the `while` walks day-by-day until `date.AddDays(step)` (115) exceeds `DateTime.MaxValue` and throws `ArgumentOutOfRangeException` (verified) after up to ~2.9M iterations (2026→9999 = 2,912,283 days, verified). Neither exception is caught in `ComputeOne` (the `try` at 97-104 only wraps `FromOADate`), so **one bad cell fails the entire block** — the doc promises `#VALUE!` "in that cell".
**Trigger condition:** `=EPT.WORKDAYADD(A1:A100, B1)` where B1 = `3e9` (case a) or `5e6` (case b — also freezes recalc: 100 cells × 2.9M steps × 16 ns ≈ 4.7 s before failing; 100k cells × days=5000 ≈ 11 s of legitimate-looking but unbounded walking).
**Trace (case a):** `dayCountRaw = 3e9` → `remaining = int.MinValue` → `remaining == 0`? no → `step = remaining > 0 ? 1 : -1` = **−1** (wrong direction) → `Math.Abs(int.MinValue)` → `OverflowException` → propagates out of `WorkdayAdd` → Excel-DNA unhandled handler → whole call `#VALUE!` (every cell of the result, including valid rows).
**Fix:** range-check before the cast and trap the walk, returning per-cell errors. Excel's serial range is 0..2,958,465 (9999-12-31), so any |days| above that can never complete:
```csharp
// before
var remaining = (int)Math.Truncate(dayCountRaw);
...
while (remaining > 0) { date = date.AddDays(step); ... }
// after
const double MaxSerial = 2_958_465d;
var truncated = Math.Truncate(dayCountRaw);
if (Math.Abs(truncated) > MaxSerial) return ExcelError.ExcelErrorNum;
var remaining = (int)truncated;
...
try
{
    while (remaining > 0) { date = date.AddDays(step); if (!IsNonWorking(date, weekend, holidays)) remaining--; }
}
catch (ArgumentOutOfRangeException) { return ExcelError.ExcelErrorNum; }   // walked past year 9999 / year 1
```

### BUG-MEDIUM-6
**File:** /home/user/plugin/src/ExcelPerfToolkit/TextUtilities.cs:277-307 (`ToCamel`), 257-275 (`ToProper`)
**Category:** data_corruption
**Description:** `ToCamel` classifies UTF-16 code units with `char.IsLetterOrDigit`, which is `false` for both halves of a surrogate pair. Astral-plane letters — CJK Extension B+ ideographs (used in real Chinese personal/place names), mathematical alphanumerics, emoji — are treated as separators and **silently deleted**, and the next BMP letter is wrongly treated as a word start. `ToProper` has the milder form: astral letters survive but reset `newWord` (264-272), so the following letter is wrongly capitalized.
**Trigger condition:** `=EPT.CAMELCASE(A1)` with A1 = `"x𠀀y zz"` (U+20000, CJK Ext-B).
**Trace (empirically verified):** `char.IsLetterOrDigit('\uD840')` = false → both surrogate halves skipped by the separator loop (284-287) → `'y'` begins a "new word" → output `"xYZz"`: the ideograph is gone and `y` is wrongly upper-cased. Also verified: `"𝒜lpha beta"` → `"lphaBeta"` (leading astral letter deleted).
**Fix:** iterate runes, not chars:
```csharp
// before (ToCamel inner loops, char-based)
while (i < s.Length && !char.IsLetterOrDigit(s[i])) i++;
...
sb.Append(first ? char.ToUpperInvariant(ch) : char.ToLowerInvariant(ch));
// after (rune-based)
foreach (var rune in s.EnumerateRunes())
{
    if (!Rune.IsLetterOrDigit(rune)) { first = true; continue; }
    sb.Append(first ? Rune.ToUpperInvariant(rune) : Rune.ToLowerInvariant(rune));
    first = false;
}
// camel-lowering of the first rune via Rune.DecodeFromUtf16(sb-prefix) instead of sb[0]
```
Apply the same `Rune.IsLetter` classification to `ToProper`'s `newWord` logic.

### BUG-LOW-7
**File:** /home/user/plugin/src/ExcelPerfToolkit/TextUtilities.cs:148-159
**Category:** data_corruption (documented, but understated)
**Description:** Code-unit `Array.Reverse` (157) turns every surrogate pair into a *reversed* pair — an ill-formed UTF-16 string, not merely "characters not preserved" as the doc (143-145) claims. Ill-formed BSTRs survive the COM hop but become `U+FFFD` on any UTF-8 conversion (clipboard, CSV export, Power Query), and combining sequences ("e" + U+0301) reorder so the accent attaches to the wrong base.
**Trigger condition:** `=EPT.REVERSE(A1)` with A1 = `"ab🙂"` → result is `'\uDE42','\uD83D',"ba"` — lone-surrogate-ordered garbage rather than `"🙂ba"`.
**Trace:** `"ab🙂"` = `['a','b','\uD83D','\uDE42']` → `Array.Reverse` → `['\uDE42','\uD83D','b','a']` — low surrogate now precedes high surrogate; invalid Unicode.
**Fix:** after `Array.Reverse`, re-swap surrogate pairs in one linear pass (or reverse by `StringInfo` text elements if combining-mark fidelity is wanted):
```csharp
Array.Reverse(chars);
for (var k = 0; k < chars.Length - 1; k++)
    if (char.IsLowSurrogate(chars[k]) && char.IsHighSurrogate(chars[k + 1]))
    { (chars[k], chars[k + 1]) = (chars[k + 1], chars[k]); k++; }
```

### BUG-LOW-8
**File:** /home/user/plugin/src/ExcelPerfToolkit/TextUtilities.cs:247-255 (`PadCharOf`, used at 73/89)
**Category:** data_corruption
**Description:** `s[0]` takes the first UTF-16 code unit of the pad-char argument. For an astral pad char the result is a lone high surrogate (empirically verified: `"🙂"[0] = U+D83D`), so every padded cell is filled with ill-formed units.
**Trigger condition:** `=EPT.PADLEFT(A1:A10, 10, "🙂")`.
**Trace:** `PadCharOf("🙂", ' ')` → `s.Length > 0` → returns `'\uD83D'` → `PadLeft(width, '\uD83D')` fills with lone high surrogates.
**Fix:** reject what cannot be represented as a single `char`:
```csharp
var s = Marshaling.ToStringSafe(padChar);
if (s.Length == 0) return fallback;
if (char.IsSurrogate(s[0]))
    throw new ArgumentException("pad_char must be a single basic-plane character.");
return s[0];
```

### BUG-LOW-9
**File:** /home/user/plugin/src/ExcelPerfToolkit/RegexUtilities.cs:245-261 (`MapText` line 255-257), call sites 40, 67, 100
**Category:** logic_error (silent error masking)
**Description:** `Marshaling.IsBlankOrError(cell)` lumps `ExcelError` with blanks, so an input `#N/A`/`#REF!` cell becomes `FALSE` (REGEXMATCH), `0` (REGEXCOUNT) or `""` (REGEXEXTRACT). The docs only promise this for *blank* cells. Downstream logic (`COUNTIF(range,TRUE)`, sums of counts) silently absorbs broken upstream formulas.
**Trigger condition:** A1 = `=1/0` (`#DIV/0!`); `=EPT.REGEXCOUNT(A1, "\d")` → `0`, indistinguishable from "no digits".
**Trace:** `block[0,0] = ExcelError.ExcelErrorDiv0` → `IsBlankOrError` → true → `result[0,0] = blankDefault (0d)`.
**Fix:**
```csharp
// before
result[r, c] = Marshaling.IsBlankOrError(cell) ? blankDefault : cellFn(...);
// after
result[r, c] = cell is ExcelError err ? err
             : Marshaling.IsBlankOrError(cell) ? blankDefault
             : cellFn(Marshaling.ToStringSafe(cell), r, c);
```

### BUG-LOW-10
**File:** /home/user/plugin/src/ExcelPerfToolkit/DistanceUtilities.cs:90-99
**Category:** logic_error (FP residue on the symmetric diagonal)
**Description:** For the self-comparison case, `dot == selfA[i]` exactly (identical computation), but `denom = √s·√s ≠ s` for roughly half of all magnitudes, so the diagonal of `EPT.DISTANCE(A,,"cosine")` is ±1–2·10⁻¹⁶ instead of 0. Empirically verified: `1 − s/(√s·√s) ≠ 0` for 470,569 of 1,000,000 random `s`. Sheet-level checks like `=result=0` or `MATCH(0, row, 0)` fail.
**Trigger condition:** `=EPT.DISTANCE(A1:C5,, "cosine")` — diagonal cells show `-2.22E-16`.
**Trace:** i == j, `bBlock == matrixA` → `dot = selfA[i] = s`; `denom = Math.Sqrt(s) * Math.Sqrt(s) = s·(1 ± 2⁻⁵²)`; `1 − s/denom = ∓1.1·10⁻¹⁶`.
**Fix:** short-circuit the diagonal in the self case:
```csharp
var self = ReferenceEquals(matrixA, bBlock);
...
for (var j = 0; j < rowsB; j++)
{
    if (self && i == j) { result[i, j] = 0d; continue; }
    ...
}
```

### BUG-LOW-11
**File:** /home/user/plugin/src/ExcelPerfToolkit/DistanceUtilities.cs:36
**Category:** logic_error (silent error masking)
**Description:** `Marshaling.IsBlankOrError(matrixB) ? matrixA : ...` treats a scalar **error** second argument the same as "omitted": the function silently computes the self-distance of `matrix_a` instead of propagating the error.
**Trigger condition:** `=EPT.DISTANCE(A1:C10, F1)` where F1 evaluates to `#REF!` (e.g. a deleted-range reference) → a plausible-looking rowsA×rowsA matrix appears instead of an error.
**Trace:** `matrixB = ExcelError.ExcelErrorRef` → `IsBlankOrError` → true → `bBlock = matrixA` → full self-distance computed and returned.
**Fix:**
```csharp
if (matrixB is ExcelError bErr) return Marshaling.ErrorBlock(bErr);
var bBlock = Marshaling.IsBlankOrError(matrixB) ? matrixA : Marshaling.AsArray2D(matrixB);
```

### BUG-LOW-12
**File:** /home/user/plugin/src/ExcelPerfToolkit/TextUtilities.cs:309-353 (`RenderTemplate`), 209-214
**Category:** logic_error (missing limit enforcement)
**Description:** Rendered rows have no `MaxCellChars` check; the file otherwise enforces the 32,767 limit everywhere (24, 110, 129, 240). A template substituting several near-limit cells produces a >32,767-char string that cannot exist in a cell — Excel-DNA's XLOPER12 string marshaling cannot represent it, so the cell is silently truncated/failed at the boundary instead of the file's documented per-cell error discipline (doc lines 17-19).
**Trigger condition:** `=EPT.TEMPLATEFILL("{0}{1}", A1:B2, FALSE)` with two 20,000-char cells → 40,000-char render.
**Trace:** `RenderTemplate` appends `Marshaling.ToStringSafe(data[row,col])` (333) unbounded → `result[r,0]` = 40,000-char string → exceeds Excel's cell capacity at marshal time.
**Fix:** after rendering: `var s = sb.ToString(); result[r, 0] = s.Length > MaxCellChars ? ExcelError.ExcelErrorValue : s;` (matches the class-doc contract "per-cell limits embed an ExcelError").

---

## Pass 2 — Memory optimization

Representative call: 100,000 cells, average ~20-char strings, one regex match per cell.
Measurements via `GC.GetAllocatedBytesForCurrentThread()` on .NET 8.0.28 x64.

### MEM-1
**File:** /home/user/plugin/src/ExcelPerfToolkit/RegexUtilities.cs:71
**Current cost:** `rx.Matches(s).Count` materializes a `MatchCollection` + one `Match` (with group/capture arrays) per match. **Measured: 33.6 MB allocated per 100k-cell call** (1 match/cell, ~352 B/cell).
**Optimization:** .NET 7+ instance method `Regex.Count(string)` runs the scan without materializing `Match` objects (verified available on .NET 8).
**Expected reduction:** 33.6 MB → **0 MB** of match-machinery garbage (only the boxed `double` result remains, 2.4 MB); also removes the proportional Gen0 pressure during MTR recalc.
**Before → after:** `return (double)rx.Matches(s).Count;` → `return (double)rx.Count(s);` (note: `Regex.Count` honors the instance `MatchTimeout`, keep the existing catch).

### MEM-2
**File:** /home/user/plugin/src/ExcelPerfToolkit/DistanceUtilities.cs:39-40
**Current cost:** When `matrix_b` is omitted, `bBlock == matrixA` (line 36) yet line 40 runs `Flatten(bBlock, ...)` again: a duplicate `rows·dim·8 B` array **plus** a duplicate `TryToDouble` pass over every cell. For 5,000×200 self-distance: 8 MB duplicate buffer + 1M redundant cell conversions (string cells re-parse).
**Optimization:** reuse the flat buffer exactly as `SelfDots` already reuses `selfA` (line 68).
**Expected reduction:** −8 MB and −50% of flatten CPU for the self case (the most common call shape).
**Before → after:**
```csharp
var flatA = Flatten(matrixA, out var rowsA, out var dim);
var flatB = Flatten(bBlock, out var rowsB, out var dimB);
// →
var flatA = Flatten(matrixA, out var rowsA, out var dim);
double[] flatB; int rowsB, dimB;
if (ReferenceEquals(bBlock, matrixA)) { flatB = flatA; rowsB = rowsA; dimB = dim; }
else flatB = Flatten(bBlock, out rowsB, out dimB);
```

### MEM-3
**File:** /home/user/plugin/src/ExcelPerfToolkit/TextUtilities.cs:257-275 (`ToProper`), 277-307 (`ToCamel`), 122-139 (`Repeat`)
**Current cost:** one `StringBuilder` (≈48 B object) + its `char[]` buffer + the final string per cell — 3 allocations where 1 suffices. **Measured for `ToProper` over 100k×22-char cells: 18.2 MB and 45.6 ms.**
**Optimization:** `ToProper` is length-preserving → `string.Create(s.Length, s, ...)` writes the single result allocation in place. **Measured: 6.8 MB / 32.2 ms — 2.7× less garbage, 1.4× faster.** `Repeat`'s output length is known (`n·s.Length`) → same treatment. (`ToCamel` shrinks: render into a `stackalloc`/pooled span, then `new string(span)`.) Same span technique applies to the Pad* numeric path (ToStringSafe string + PadLeft string = 2 strings/cell → `TryFormat` into the padded buffer = 1).
**Expected reduction:** ~11.4 MB per 100k-cell `EPT.PROPER` call; similar ratios for CAMELCASE/REPEAT.
**Before → after (ToProper):**
```csharp
var sb = new StringBuilder(s.Length); ... return sb.ToString();
// →
return string.Create(s.Length, s, static (dst, src) => { /* same newWord loop writing dst[i] */ });
```

### MEM-4
**File:** /home/user/plugin/src/ExcelPerfToolkit/TextUtilities.cs:156-158
**Current cost:** `ToCharArray()` (copy 1) + `new string(chars)` (copy 2). **Measured: 13.7 MB per 100k-cell call.**
**Optimization:** `string.Create` with reverse copy: **measured 6.8 MB (2.0× reduction), 17.4 vs 19.7 ms.**
**Before → after:**
```csharp
var chars = Marshaling.ToStringSafe(c).ToCharArray(); Array.Reverse(chars); return new string(chars);
// →
return string.Create(s.Length, s, static (dst, src) => { for (var i = 0; i < dst.Length; i++) dst[i] = src[src.Length - 1 - i]; /* + surrogate fix-up from BUG-LOW-7 */ });
```

### MEM-5
**File:** /home/user/plugin/src/ExcelPerfToolkit/SeriesUtilities.cs:118 & 224 (`new List<double>(rows * cols)`), 141/168/235 (`values.ToArray()`), 171 (`deviations`)
**Current cost:** per 100k-cell call: 800 KB list backing array (allocated at full capacity even for text-heavy blocks) + 800 KB `ToArray` copy + (MAD) 800 KB deviations = up to 2.4 MB transient.
**Optimization:** sort in place over the list's buffer: `var span = CollectionsMarshal.AsSpan(values); span.Sort();` and pass the span to `PercentileSorted` (signature → `ReadOnlySpan<double>`); reuse the same buffer for MAD deviations (overwrite then re-sort).
**Expected reduction:** −1.6 MB (`iqr`/`zscore`) to −2.4 MB (`mad`) per call; zero extra copies.
**Before → after:** `var sorted = values.ToArray(); Array.Sort(sorted);` → `var sorted = CollectionsMarshal.AsSpan(values); sorted.Sort();`

---

## Pass 3 — CPU / throughput

### PERF-1
**File:** /home/user/plugin/src/ExcelPerfToolkit/DateUtilities.cs:113-120 + 124-133 — **hot path: yes** (per cell × per day)
**Current cost:** O(walked days) per cell with `DateTime.AddDays` + `DayOfWeek` + (holidays present) `date.ToOADate()` *every step*. **Measured: 16 ns/step; 10k cells × +260 workdays (364 walked days each) = 59.2 ms.** Cost scales linearly with |days|; 100k cells × 260 workdays ≈ 0.6 s, ×2600 ≈ 6 s.
**Optimization (two stages, both verified equivalent on the probe):**
1. Walk in pure `int` serial space — `serial++; maskIndex = maskIndex==6 ? 0 : maskIndex+1; holidays.Contains(serial)` — eliminating `AddDays`/`DayOfWeek`/`ToOADate` entirely (serial 0 = 1899-12-30 = Saturday → start `maskIndex = (serial + 5) % 7`). **Measured: 19.3 ms, 3.1×, bit-identical results.**
2. Closed-form full-week jumps: with W working days per week (already counted in `ParseWeekendMask`), jump `7·(remaining/W)` days at once, walk the ≤6-day remainder, then adjust for holidays in the covered interval via a sorted `int[]` + two binary searches (re-extend for holidays landed on). Walked steps drop from `days·7/W` to ≤ 7 + O(#holiday hits): ~**50×** on top of stage 1 at year-scale day counts.
**Expected speedup:** 3.1× (stage 1, trivial) → ~50–150× (stage 2) on WORKDAYADD-heavy sheets.
**Before → after (stage 1):**
```csharp
date = date.AddDays(step);
if (!IsNonWorking(date, weekend, holidays)) remaining--;
// →
serial += step; maskIndex = (maskIndex + (step > 0 ? 1 : 6)) % 7;
if (!(weekend[maskIndex] || holidays.Contains(serial))) remaining--;
```

### PERF-2
**File:** /home/user/plugin/src/ExcelPerfToolkit/RegexUtilities.cs:223-243 (`Build`) — **hot path: yes** (regex engine runs per cell)
**Current cost:** interpreted regex per cell. **Measured (email-ish pattern, 100k ~60-char cells): interpreted 212.3 ms vs compiled 43.7 ms — 4.9×.** Compile cost measured at 15.6 ms (vs 1.4 ms interpreted ctor), paid once.
**Optimization:** add `RegexOptions.Compiled`, amortized via the ARCH-1 cache (or, minimally, when `rows*cols >= 4096` so the 15.6 ms compile is repaid ≥10×).
**Expected speedup:** ~4–5× on all five regex UDFs for medium/large blocks; break-even ≈ 10k cells single-call, ≈ first recalc with a cache.
**Before → after:** `return new Regex(pattern, options, MatchTimeout);` → `return RegexCache.Get(pattern, options | RegexOptions.Compiled, MatchTimeout);` (see ARCH-1).

### PERF-3
**File:** /home/user/plugin/src/ExcelPerfToolkit/DistanceUtilities.cs:69-79, 86-101, 116-140 — **hot path: yes** (O(rowsA·rowsB·dim))
**Current cost:** when `matrix_b` is omitted the matrix is symmetric for all four metrics, yet `[i,j]` and `[j,i]` are both computed — 2× the dot products/element loops and 2× the result boxing (24 B/cell). 3,000-row self-distance, dim 50: 9M cells, ~450M FLOPs and 216 MB of boxes where half suffices.
**Optimization:** in the self case iterate `j >= i`, set the diagonal directly (0 — also fixes BUG-LOW-10), and mirror the *same boxed object*: `var boxed = (object)v; result[i,j] = boxed; result[j,i] = boxed;`.
**Expected speedup:** **1.95×** kernel time and **2×** fewer box allocations for self-distance (the default call shape).
**Before → after:** `for (var j = 0; j < rowsB; j++)` → `for (var j = self ? i : 0; j < rowsB; j++)` + mirror write.

### PERF-4
**File:** /home/user/plugin/src/ExcelPerfToolkit/DistanceUtilities.cs:116-140 — **hot path: yes**
**Current cost:** Manhattan/Chebyshev run a scalar `k`-loop with a per-element `if (manhattan)` branch (128-135) — no SIMD despite the file's stated design ("routed through the AVX2/FMA kernels") applying only to euclidean/cosine.
**Optimization:** split into two specialized loops (removes the per-element branch) and vectorize: `acc += Vector.Abs(va - vb)` / `vmax = Vector.Max(vmax, Vector.Abs(va - vb))` with scalar tails, mirroring `VectorizedKernels.DotProductPortable`.
**Expected speedup:** ~3–4× for dim ≥ 16 on AVX2 (4 doubles/lane), ~1.1–1.2× from branch removal alone for small dims.
**Before → after:** per-element `if (manhattan) acc += diff; else if (diff > acc) acc = diff;` → two dedicated kernels selected once per call.

### PERF-5
**File:** /home/user/plugin/src/ExcelPerfToolkit/SeriesUtilities.cs:119-128 vs 194-202 — **hot path: yes** (two full passes per call)
**Current cost:** `Outliers` calls `Marshaling.TryToDouble` on every cell twice — once to collect (123), once to flag (199). For numeric-as-text cells each call is a double `TryParse` (~50–150 ns) vs ~2 ns for a double unbox: on a 100k-cell text-numeric block the second pass redoes ~10–15 ms of parsing, roughly doubling non-sort time.
**Optimization:** during the first pass record per-cell results into flat `double[] vals` + `bool[] ok` (900 KB transient for 100k cells); the flagging pass reads the arrays.
**Expected speedup:** ~10× on the flagging pass for string-heavy blocks; **~1.5–2× whole-call**; also makes the two passes consistent if cell contents could differ (they cannot here, but it removes the duplicated classification logic).
**Before → after:** `result[r, c] = Marshaling.TryToDouble(block[r, c], out var d) && isOutlier(d);` → `result[r, c] = ok[idx] && isOutlier(vals[idx]); idx++;`

---

## Pass 4 — Architectural wins

### ARCH-1
**Scope:** RegexUtilities.cs (all five UDFs; header doc lines 11-14).
**Current pattern:** a fresh interpreted `Regex` per UDF call, justified by "never a shared static cache, which would violate MTR safety". That premise is incorrect: `Regex` instances are documented thread-safe for all matching operations; only the *cache container* needs to be concurrent.
**Proposed pattern:** `static readonly ConcurrentDictionary<(string Pattern, RegexOptions Options), Regex>` with `RegexOptions.Compiled`, bounded (e.g. clear-at-256 or an LRU ring) to cap JIT'd-IL retention from volatile patterns; `GetOrAdd` per call.
**Impact estimate:** 4.9× matching throughput (measured, PERF-2) on every recalc after the first; eliminates per-recalc construction (1.4 ms interpreted / 15.6 ms compiled per distinct pattern × every dependent formula re-evaluation).
**Effort:** S (≈30 LOC + doc correction). **Risk:** low — Regex thread safety is a documented BCL contract; bound the cache to avoid unbounded compiled-regex retention (the only real leak vector).

### ARCH-2
**Scope:** DateUtilities.cs (`ComputeOne`, `IsNonWorking`, `ParseHolidays`).
**Current pattern:** day-by-day walk per cell; holidays in a `HashSet<int>` probed per step.
**Proposed pattern:** per-call immutable plan: (a) `int[7]` prefix-sums of working days per weekday offset → O(1) "add N working days ignoring holidays" arithmetic; (b) holidays filtered to non-weekend days, sorted `int[]` → binary-search count in the jumped interval, iterate the (rare) extension. Plan built once in `WorkdayAdd` (it already parses mask/holidays once), shared across the broadcast loop.
**Impact estimate:** ~50–150× per cell at year-scale day counts (PERF-1 stage 2); turns the BUG-MEDIUM-5 unbounded walk into closed-form range checks for free.
**Effort:** M (~120 LOC + property tests vs the walking implementation across mask×holiday×negative-days grid). **Risk:** medium — WORKDAY.INTL edge semantics (start-on-holiday, negative counts) must be locked by tests before swapping.

### ARCH-3
**Scope:** RegexUtilities.cs `Spill` today; any future spilling UDF (the readme advertises more rounds).
**Current pattern:** each spilling function shapes its own `object[,]` with no grid-limit awareness; `DistanceUtilities` invented its own `int.MaxValue` guard (line 45-48) while `RegexUtilities` has none (BUG-MEDIUM-4).
**Proposed pattern:** one `Marshaling.AllocateSpill(rows, cols)` helper that validates rows ≤ 1,048,576, cols ≤ 16,384, total ≤ `int.MaxValue`, and throws a uniform descriptive `ArgumentException` (→ `#VALUE!`) *before* allocation.
**Impact estimate:** converts a host-process OOM (potential Excel crash, unsaved-work loss) into a clean per-call error; removes duplicated guard logic.
**Effort:** S. **Risk:** low.

### ARCH-4
**Scope:** TextUtilities.cs (all per-cell transforms) + the euclidean fallback in DistanceUtilities (BUG-HIGH-2).
**Current pattern:** delegate-per-cell `MapCells` with StringBuilder-based transforms (3 allocations/cell), char-level Unicode handling (BUG-MEDIUM-6, BUG-LOW-7/8).
**Proposed pattern:** a small span-kernel layer: each transform is a `(ReadOnlySpan<char> src, Span<char> dst)` or `string.Create`-based function, Rune-aware where classification/casing matters; `MapCells` stays as the block driver. Add a `DiffSquaredDot(ReadOnlySpan<double>, ReadOnlySpan<double>)` kernel next to `DotProduct` so the euclidean cancellation fallback stays vectorized.
**Impact estimate:** 2.0–2.7× allocation reduction and 1.3–1.4× CPU on text blocks (measured, MEM-3/4); fixes the Unicode-corruption class structurally rather than per-function.
**Effort:** M. **Risk:** low (pure functions; golden-file tests over BMP/astral/combining inputs).

---

## Seam notes (out-of-scope files, one line each)

- **Marshaling.cs `TryToDouble` (114-132):** violates its own "finite double" contract for strings — `"NaN"`, `"Infinity"`, `"1e999"` all parse `true` with non-finite results (empirically verified), letting NaN/∞ leak into `EPT.QUANTILES`/`EPT.OUTLIERS`/`EPT.DISTANCE`/`EPT.WORKDAYADD` from a single text cell (NaN-poisoned sorts, `#NUM!` spills); fix: re-check `double.IsFinite(result)` on the string path.
- **VectorizedKernels.cs `DotProductAvx2` (341-342):** loads lanes via `Vector256.Create(a[i],…,a[i+3])` (4 scalar loads + inserts) while the comment claims `Vector256.LoadUnsafe` / `vmovupd` — measurable dot-product throughput is being left on the table for `EPT.DISTANCE`'s inner loop.

---

## Rejection appendix (considered and rejected)

1. **Turkish-I / culture-sensitive casing in ToProper/ToCamel/TitleCase** — rejected: all casing is `char.ToUpperInvariant`/`ToLowerInvariant`/`InvariantCulture.TextInfo` (TextUtilities.cs:45, 265, 297, 304); no `ToUpper()`/`ToLower()` culture-sensitive calls exist in scope. Verified `char.ToLowerInvariant('İ')` returns U+0130 unchanged (no 1:1 invariant mapping) — no corruption, merely a no-op.
2. **TitleCase acronym claim wrong** — rejected: empirically verified `ToTitleCase("NASA and esa")` → `"NASA And Esa"`, `"the QUICK bRown fOx"` → `"The QUICK Brown Fox"`; the doc at TextUtilities.cs:39-40 is accurate.
3. **PercentileSorted ≠ PERCENTILE.INC** — rejected: `rank = p·(n−1)` linear interpolation is exactly Excel's inclusive definition; verified `P({1,2,3,4},0.75)=3.25` and `P({15,20,35,40,50},0.4)=29` match Excel's documented results; `p=1` and single-element edges correct (SeriesUtilities.cs:286-301).
4. **zscore σ=0 / MAD=0 divide-by-zero** — rejected: both explicitly guarded with all-FALSE fallbacks (SeriesUtilities.cs:156-159, 179-181); IQR=0 with genuine outliers still flags correctly (fences collapse to [q1,q1]).
5. **All-weekend mask infinite loop** — rejected: `ParseWeekendMask` counts working days and throws when zero (DateUtilities.cs:162-165); holidays are a finite set so the walk always terminates for valid masks (modulo BUG-MEDIUM-5's range issues).
6. **Serial-60 / 1900-leap-day off-by-one in WORKDAYADD** — rejected after verification: `FromOADate(60)` = 1900-02-28 round-trips to 60, and `FromOADate(1)` = 1899-12-31 **Sunday** coincides with Excel's `WEEKDAY(1)=Sunday` labeling (both shifted from proleptic reality identically), so weekend classification and serial arithmetic agree with Excel's own WORKDAY behavior across the phantom-day region; the phantom serial 60 is counted as a day exactly as Excel itself does.
7. **Fractional-OADate holiday mismatch** — rejected: both sides floor to whole-day ints — `ParseHolidays` at DateUtilities.cs:185, `IsNonWorking` at 132 (and `date` is midnight-normalized at 99, so `ToOADate()` is integral).
8. **Regex static-cache race / any MTR race** — rejected: no static *mutable* state exists in any in-scope file; `TraceSource` is thread-safe; per-call `Regex` is correct (only slow — see PERF-2/ARCH-1, where the *stated rationale* is flagged instead).
9. **`rows*cols` int overflow in `new List<double>(rows * cols)` (SeriesUtilities.cs:118, 224)** — rejected: overflow needs >2³¹ cells in one `object[,]`, which the CLR/Excel-DNA cannot construct; `DistanceUtilities.Flatten` (158-161) guards the same product explicitly where it feeds an allocation size.
10. **Lazy `MatchCollection` timeout escaping the try/catch** — rejected: enumeration is forced inside the `try` at every site (`.Count` at RegexUtilities.cs:71, `foreach` at 156, `Split` at 204), so `RegexMatchTimeoutException` is always caught per cell/row.
11. **FillForward direction semantics inverted** — rejected: traced all four loops (SeriesUtilities.cs:39-78); `down` carries top→bottom per column, `up` bottom→top, `right`/`left` per row, matching the doc; errors deliberately count as carriable values (doc says *blank* cells are filled), leading blanks stay blank via `last ?? ExcelEmpty.Value`.
12. **`group_index` bounds / missing-group crash** — rejected: `groupIndex >= m.Groups.Count` and `!Groups[i].Success` both checked (RegexUtilities.cs:105, 158).
13. **Booleans / numeric-looking strings counted as data in Outliers/Quantiles/Distance** — rejected as toolkit-wide design: `Marshaling.TryToDouble` documents recognizing booleans and numeric strings; behavior is consistent across all block functions (a deliberate divergence from native Excel aggregate coercion).
14. **Text-date holidays / text start dates silently ignored** — rejected: the contract (DateUtilities.cs:24-29) says "range of date serials"; coercing text dates would import locale parsing problems the toolkit deliberately avoids.
15. **WORKDAYADD negative result serials** (backward walks past 1899-12-30 return negative OADate where native WORKDAY gives `#NUM!`) — rejected: numerically correct serial arithmetic; purely a native-compat cosmetic divergence on inputs Excel can't even format.
16. **`Repeat(x, 0)` returns `""` not blank** — rejected: matches native `REPT(x,0)`.
17. **`string.PadLeft` splitting surrogates / wrong width math** — rejected: PadLeft only prepends, never splits existing pairs; `WidthOf` counts UTF-16 units exactly as Excel's `LEN` does, so width semantics agree with the host. (The pad *char* itself is the only hazard — BUG-LOW-8.)
18. **TemplateFill duplicate headers (first-wins, 191) / numeric token shadowing a header literally named "3"** — rejected: deterministic, documented-enough precedence ("index or header name"), negligible real-world surface.
19. **`Quantiles` NaN-poisoned sort** — rejected as an in-scope bug: unreachable when `TryToDouble` honors its finite contract (doubles that are NaN/∞ are filtered at the case at Marshaling.cs:85); reachable only through the string seam — covered by the Marshaling seam note.
20. **Cosine ∞/NaN for ~1e154+ magnitude rows** (selfdot overflows to +∞ → denom ∞ → distance NaN → `#NUM!`) — rejected: requires magnitudes Excel sheets don't produce except via the same Marshaling seam; with finite inputs ≤1e150 the math is safe.
21. **`Distance` `dim == 0`** — rejected: unreachable from Excel (minimum 1×1 range) and explicitly handled anyway (DistanceUtilities.cs:51-61).
22. **MapCells/MapText delegate dispatch overhead** — rejected as a finding: ~2-3 ns/cell against µs-scale per-cell string/regex work (<1% of any measured loop); not worth the API churn outside ARCH-4.
23. **`MapCells` re-boxing pass-through numbers** — rejected: pass-through returns the original boxed reference (`c ?? ExcelEmpty.Value`), no new boxes.
24. **`(int)totalWidth`, `(int)Math.Floor(rank)` overflow** — rejected: `WidthOf` rejects >32,767 before the cast (TextUtilities.cs:240-244); `rank ≤ n−1 < 2³¹` by construction.
25. **`Repeat` throws whole-call instead of embedding per-cell error** (TextUtilities.cs:129-132 vs class doc 17-19) — rejected as a separate bug: the function-level doc (110) explicitly promises a throw; noted as a doc inconsistency only, and the fix for BUG-MEDIUM-3 keeps that contract.
26. **`Outliers` negative `threshold` silently mapped to the default** (SeriesUtilities.cs:146, 155, 178: `threshold > 0d ? threshold : default`) — rejected: arguably intentional ("0 (default) uses the per-method default"); a negative threshold has no sensible meaning and the chosen behavior is benign and deterministic.
27. **Euclidean diagonal residue (self case)** — rejected: unlike cosine (BUG-LOW-10), `sq = s + s − 2s = 0` exactly; diagonal is exact.
