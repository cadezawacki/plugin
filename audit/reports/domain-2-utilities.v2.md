# Domain 2: Developer Utilities — Round 2 Audit (Fresh Eyes)

Scope: `/home/user/plugin/src/ExcelPerfToolkit/DeveloperUtilities.cs` only.
Project setting confirmed: `ExcelPerfToolkit.csproj` does NOT define
`<CheckForOverflowUnderflow>` — multiplication is unchecked.
Marshaling helpers consulted for inferred behavior of `TryToDouble`/`ToStringSafe`/`IsBlankOrError` only (not audited).

## Findings table

| ID  | Sev      | Location                                     | Category                   | Confidence |
| --- | -------- | -------------------------------------------- | -------------------------- | ---------- |
| U-01| Critical | DeveloperUtilities.cs:418                    | ReDoS / blocking call      | 0.95       |
| U-02| High     | DeveloperUtilities.cs:864-867, 902-906       | Hash collision (cell separator) | 0.90  |
| U-03| High     | DeveloperUtilities.cs:185                    | Key collision (RemoveDuplicateRows) | 0.85 |
| U-04| High     | DeveloperUtilities.cs:466                    | Integer overflow → IndexOutOfRange | 0.90 |
| U-05| High     | DeveloperUtilities.cs:502                    | Integer overflow → IndexOutOfRange | 0.90 |
| U-06| High     | DeveloperUtilities.cs:757-792                | Silent misbehavior on scalar descendingFlags | 0.85 |
| U-07| Med      | DeveloperUtilities.cs:125, 819               | "NaN"/"Infinity" parse asymmetry; double NaN never compares-as-number | 0.85 |
| U-08| Med      | DeveloperUtilities.cs:153-178                | Doc/code mismatch — uses ToStringSafe, not CellEquality | 0.95 |
| U-09| Med      | DeveloperUtilities.cs:865, 903-908           | Per-cell byte[] / `ToArray` allocations | 0.90 |
| U-10| Med      | DeveloperUtilities.cs:711-737                | FlattenIntIndexes silent double→int truncation/wrap | 0.80 |
| U-11| Med      | DeveloperUtilities.cs:413-425                | Regex contract mismatch with Replace path | 0.70 |
| U-12| Low      | DeveloperUtilities.cs:209-216                | Cache-unfriendly Transpose loop order | 0.70 |
| U-13| Low      | DeveloperUtilities.cs:674-681                | ContainsKey + indexer instead of TryAdd | 0.95 |
| U-14| Low      | DeveloperUtilities.cs:543, 715, 768          | List capacity uses unchecked multiplication | 0.65 |
| U-15| Low      | DeveloperUtilities.cs:585-625                | UniqueCount: redundant `samples` dictionary | 0.95 |
| U-16| Low      | DeveloperUtilities.cs:262-286                | FillDown propagates NaN/error sentinels as-if-data | 0.55 |

---

## U-01 — Critical — ReDoS via user-supplied regex with no timeout or NonBacktracking

`ID | severity | location | category | scenario | evidence | fix | confidence`
`U-01 | Critical | DeveloperUtilities.cs:418 | ReDoS / blocking call | An attacker (or unwitting user) supplies a regex with catastrophic backtracking. The UDF blocks the Excel calc thread forever. | rx = new Regex(find, RegexOptions.CultureInvariant); ... rx.Replace(s, replace); — no MatchTimeout, no NonBacktracking. .NET default Regex engine is backtracking. | Use either RegexOptions.NonBacktracking (.NET 7+, available on net8.0) or pass an explicit MatchTimeout (e.g. TimeSpan.FromSeconds(2)) via the Regex(pattern, options, matchTimeout) constructor. NonBacktracking is preferred since it guarantees linear time. | 0.95`

`find` is user-controlled. Pattern `^(a+)+$` against input `aaaaaaaaaaaaaaaaaaab` exhibits exponential backtracking. Excel calc engine has no preemption — the worksheet hangs.

The XML-doc comment claims "Regex compilation is cached locally to the call to avoid a shared static cache, which would violate MTR safety guidance" but says nothing about ReDoS protection.

Surgical fix:
```csharp
rx = new Regex(find, RegexOptions.CultureInvariant | RegexOptions.NonBacktracking);
// or
rx = new Regex(find, RegexOptions.CultureInvariant, TimeSpan.FromSeconds(2));
```
The second form throws `RegexMatchTimeoutException` per `Replace`, which would need translation to `ExcelError.ExcelErrorValue` or rethrow.

---

## U-02 — High — HashBlock/Sha256Block separator collision (real ambiguity, not theoretical)

`U-02 | High | DeveloperUtilities.cs:864-867, 902-906 | Hash collision (cell separator) | Two semantically distinct blocks produce identical digests. Defeats the function's stated purpose ("change detection"). | The 0x1F (Unit Separator) byte is appended after every cell's UTF-8. A cell containing the character U+001F encodes to byte 0x1F in UTF-8 (since U+001F is in the ASCII range). Cells ["AB", ""] and ["A", "B"] both encode to: 'A' 0x1F 'B' 0x1F 0x1F | 'A' 0x1F 'B' 0x1F 0x1F — identical hasher input. Same for 0x1E (Record Separator) and the row boundary. | Encode the cell length before the bytes (e.g. hash a 4-byte little-endian length prefix followed by the UTF-8). Length-prefixing is unambiguous regardless of cell contents. | 0.90`

Concrete trace at line 864-867:
```
text = Marshaling.ToStringSafe(block[r, c]);   // can legitimately contain U+001F
bytes = Encoding.UTF8.GetBytes(text);          // U+001F → 0x1F
hasher.Append(bytes);
hasher.Append(separator);                      // 0x1F
```

While Excel rarely stores control characters via normal entry, they routinely appear via copy/paste from web sources, CSV imports, RTF, and PDF-extracted text. The function is documented as suitable for "stable hash...change detection" — the documented contract is broken on a wide and realistic input class.

Same identical issue at 902-906 for `Sha256Block`. Per-cell 0x1F and per-row 0x1E both collide with the same byte values inside cell strings.

---

## U-03 — High — RemoveDuplicateRows key separator collision

`U-03 | High | DeveloperUtilities.cs:185 | Key collision (RemoveDuplicateRows) | Two distinct rows are silently coalesced. Data loss. | sb.Append(Marshaling.ToStringSafe(block[row, c])).Append('\x1f'); — uses a single  character to separate fields. A cell containing U+001F creates an ambiguous join. Row ["AB", ""] and row ["A", "B"] both produce key "A\x1FB\x1F\x1F". The second row is dropped as a "duplicate". | Length-prefix each field, or use a separator that cannot appear in any string (impossible) — best fix is to prepend each field's char count, e.g. sb.Append(field.Length).Append(':').Append(field).Append('\x1f'). | 0.85`

The XML doc on this function (line 142-143) claims equality is "determined by cell-by-cell comparison using `Marshaling.CellEquality`" but the code path does no such thing — it concatenates `ToStringSafe` outputs. See also U-08 for the doc/code mismatch.

---

## U-04 — High — StackColumns integer overflow → silent zero-size array → IndexOutOfRange

`U-04 | High | DeveloperUtilities.cs:466 | Integer overflow → IndexOutOfRange | A worksheet-sized block crashes with a misleading IndexOutOfRangeException instead of a clean ArgumentException. | var result = new object[rows * cols, 1]; — int multiplication is unchecked (no <CheckForOverflowUnderflow>). For rows=1,048,576 (Excel max) and cols=16,384 (Excel max), the product is 17,179,869,184. Mod 2^32 = 0, then narrowed to int = 0. result = new object[0, 1]. The first iteration result[idx++, 0] = block[r, c]; throws IndexOutOfRangeException with idx=0. | Use checked: var total = checked((long)rows * cols); if (total > int.MaxValue) throw new ArgumentException(...); — or simply checked((int)((long)rows * cols)). | 0.90`

Other overflow-into-positive cases also produce a too-small array and crash partway through with `IndexOutOfRangeException` rather than a clear error. Either result is functionally a denial-of-service against legitimate large inputs.

---

## U-05 — High — Unpivot integer overflow on `dataRows * valueColumns`

`U-05 | High | DeveloperUtilities.cs:502 | Integer overflow → IndexOutOfRange | Worksheet-sized unpivot operation crashes part-way through with IndexOutOfRangeException. | var outRows = dataRows * valueColumns; — unchecked int multiplication. dataRows=(rows-1), valueColumns=(cols-keyColumns). For rows≈1M, valueColumns≈16K the product overflows. If it wraps to a small positive value the allocation succeeds and the nested write loop overshoots. If it wraps to negative, new object[negative, outCols] throws OverflowException. | Compute the product as long, check against int.MaxValue, throw a clear ArgumentException("Unpivot output exceeds 2^31 cells"). | 0.90`

---

## U-06 — High — `SortBlock` silently ignores a scalar `descendingFlags` argument

`U-06 | High | DeveloperUtilities.cs:757-792 | Silent misbehavior on scalar descendingFlags | User passes a single TRUE for descendingFlags expecting descending sort; receives ascending sort with no error. | descendingFlags is declared as object. Excel-DNA delivers a single-cell argument as the boxed scalar (e.g. bool true), NOT as object[,]. The only branch that handles flags is `if (descendingFlags is object[,] flagsBlock)` at line 764. For any non-array input — bool, double, string — control falls through to `desc = new bool[keys.Length]` (all false). The user's intent is silently lost. | Add a scalar branch: switch on descendingFlags is bool/double/string and replicate to all keys; or document that descendingFlags MUST be a range. Best: handle scalars explicitly. | 0.85`

Trace:
- Line 757: parameter type is `object` (not `object[,]`).
- Line 764: `if (descendingFlags is object[,] flagsBlock)` — false for `bool` and `double` and `string`.
- Line 791: `desc = new bool[keys.Length];` — all ascending.

This is the same Excel-DNA marshaling quirk that the rest of the file mostly avoids (note `keyColumns` IS typed `object[,]` which causes Excel-DNA to wrap scalars into 1x1 arrays — but `descendingFlags` is typed `object`, breaking the pattern).

---

## U-07 — Med — "NaN"/"Infinity" text coerces to NaN/Infinity double; numeric paths then reject it

`U-07 | Med | DeveloperUtilities.cs:125, 819 | Inconsistent NaN handling | A worksheet author lays out CoerceNumeric over a column containing the string "NaN" or "Infinity"; downstream UDFs (Sort, BlockLookup numeric compare, etc.) then exhibit surprising ordering or fall back to text. | Marshaling.TryToDouble's string branch uses double.TryParse(..., NumberStyles.Float | AllowThousands, InvariantCulture). Per .NET docs, double.TryParse with InvariantCulture accepts "NaN", "Infinity", "-Infinity" using the culture's NumberFormatInfo symbols (works regardless of NumberStyles flags). So "NaN" → double.NaN. But Marshaling.TryToDouble for an already-boxed double returns FALSE if double.IsNaN(d) (line 85 of Marshaling.cs). So CoerceNumeric turns "NaN" into a double that subsequent numeric operations refuse to recognize as a number. In SortBlock (line 819) the (double NaN, double NaN) pair fails TryToDouble && TryToDouble, falls into ordinal string compare of "NaN" vs "NaN". | Either reject "NaN"/"Infinity" in the string branch of TryToDouble, or accept double.NaN in the double branch. Pick one. The cleaner fix is reject special symbols in CoerceNumeric — leave the original cell intact. | 0.85`

This is a real "corruption edge" — a value goes through one transformation and becomes opaque to subsequent transformations.

---

## U-08 — Med — XML doc claims `RemoveDuplicateRows` uses `CellEquality`, code uses `ToStringSafe`

`U-08 | Med | DeveloperUtilities.cs:153-178 | Documentation/code mismatch | Caller relies on the documented contract — believes 1.0 and "1" are compared via CellEquality (which has a numeric path) — and is surprised when the function actually uses stringified comparison. | Lines 142-143: "Equality is determined by cell-by-cell comparison using Marshaling.CellEquality." Lines 162, 180-188: BuildRowKey uses Marshaling.ToStringSafe, not CellEquality.GetHashCode/Equals. ToStringSafe is type-erasing: ToStringSafe(1.0) == "1" and ToStringSafe("1") == "1" collide as duplicates; CellEquality would treat double 1.0 vs string "1" by its TryToDouble path with numeric equality. The user-observable contract differs from the documented one. | Either (a) update XML doc to match implementation, or (b) refactor to use CellEquality (HashSet<object> with Marshaling.CellEquality as IEqualityComparer<object?>). Option (b) is the documented contract and is what callers will rely on; it also avoids the U-03 separator collision. | 0.95`

---

## U-09 — Med — Per-cell `byte[]` and `Span.ToArray()` allocations in hash paths

`U-09 | Med | DeveloperUtilities.cs:865, 903-908 | Excessive allocations under load | HashBlock and Sha256Block allocate 1+ throwaway byte arrays per cell. For a 1M-cell block: ~1M byte[] allocations for HashBlock and ~3M for Sha256Block (text bytes + separator.ToArray() per cell + rowSeparator.ToArray() per row). At Excel-recalc cadence this creates Gen0 pressure and tens of milliseconds of GC time. | Line 865 (HashBlock): `var bytes = Encoding.UTF8.GetBytes(text);` — fresh array per cell. Line 905-906 (Sha256Block): `var sep = separator.ToArray();` inside the inner cell loop — fresh 1-byte array per cell. Same 0x1F separator allocated 1M times. Line 908: rowSeparator.ToArray() per row. | (a) Reuse a pooled byte buffer via ArrayPool<byte>.Shared and Encoding.UTF8.GetBytes(string, byte[]) or GetBytes(ReadOnlySpan<char>, Span<byte>). (b) For Sha256Block specifically, allocate the separator byte[] once outside the loop and reuse — TransformBlock takes byte[] not Span, so the static-allocation pattern is the right one. The current `stackalloc + ToArray() inside loop` is the worst of both worlds. | 0.90`

---

## U-10 — Med — `FlattenIntIndexes` truncates large doubles to int without range check

`U-10 | Med | DeveloperUtilities.cs:711-737 | Silent double→int truncation | A user passes a double value that does not fit in int (e.g. 5e9 from a formula); FlattenIntIndexes silently produces a wrapped int via the C# narrowing conversion. The caller then either: (a) trips the bounds check at BlockLookup line 668 (safe — clean ArgumentOutOfRangeException), or (b) for SortBlock at line 815 — `var col = keys[k]` then `block[a, col]` — the bounds check is the array indexer itself, which throws IndexOutOfRangeException for out-of-range. Either way the wrapped value never reaches an unsafe access — but the error message hides the root cause (truncation). | Line 729: list.Add((int)d); — C# (int)(double) on an out-of-range value produces an unspecified result on x86/x64 (per ECMA-335 — typically int.MinValue on overflow). No range check before cast. Line 725 already checks TryToDouble; just needs an extra range check. | if (d < int.MinValue \|\| d > int.MaxValue \|\| double.IsNaN(d)) throw new ArgumentException("Index out of int range."); list.Add((int)d); | 0.80`

Round 1's claim of "unsafe access" is rejected — the wrapped value is always bounds-checked by the immediate caller (BlockLookup line 668, SortBlock indexer at 816). But the error message is poor and the silent wrap is a usability finding. Severity: Med.

---

## U-11 — Med — Regex path uses CultureInvariant; non-regex path uses Ordinal — inconsistent contract

`U-11 | Med | DeveloperUtilities.cs:413-425 | Inconsistent culture handling between regex and non-regex paths | A user toggles useRegex on the same data and observes different match counts for case-folded text. | Line 435: non-regex branch uses `s.Replace(find, replace, StringComparison.Ordinal)` — byte-for-byte, no culture. Line 418: regex branch uses `RegexOptions.CultureInvariant` — culture-aware-but-invariant case folding for character classes like `\w`. These are not the same; ordinal is stricter. For Turkish-I or ß/ss the two paths diverge. | Pick one. If the intent is "no culture surprises" (most likely for an Excel UDF), use RegexOptions.None and let the default culture-insensitive matching apply, or use RegexOptions.CultureInvariant on both sides (but s.Replace doesn't take a CultureInfo for ordinal — that's fine). Document the chosen contract. | 0.70`

---

## U-12 — Low — `Transpose` writes column-major into a row-major 2D array

`U-12 | Low | DeveloperUtilities.cs:209-216 | Cache-unfriendly access pattern | Large transposes (16K x 16K) thrash the L1/L2 cache. Wall time can be 5-10x slower than blocked transpose for blocks that don't fit in cache. | Inner loop iterates c=0..cols and writes result[c, r]. .NET's object[,] is stored in row-major order. So consecutive writes step by cols * sizeof(object) bytes — one full row apart — guaranteeing a cache miss on every write for any block wider than L1/cols. | Cache-block the transpose (tile size ~64 elements) — or swap loop order so reads stride cache lines and accept that writes will stride. For object[,] either direction strides on one side, so this is a write-stride vs read-stride tradeoff. Tiled transpose (Bx B blocks) wins for large blocks. | 0.70`

Low severity because correctness is fine and the file's primary contract is "avoid COM boundary crossings", not in-process throughput.

---

## U-13 — Low — `BlockLookup` uses `ContainsKey + indexer assign` instead of `TryAdd`

`U-13 | Low | DeveloperUtilities.cs:674-681 | Micro-inefficiency | Per lookup-table row, two hash lookups instead of one. For a 1M-row lookup, 1M extra hash computations. | for (var r = 0; r < lookupRows; r++) { var key = ...; if (!index.ContainsKey(key)) { index[key] = r; } } — ContainsKey is O(1) hash; indexer set is another O(1) hash + insert. | for (var r = 0; r < lookupRows; r++) { index.TryAdd(Marshaling.ToStringSafe(lookup[r, 0]), r); } — same first-wins semantic, single hash op. | 0.95`

Pure cleanup — not data-correctness.

---

## U-14 — Low — `List<T>` capacity hints use unchecked multiplication

`U-14 | Low | DeveloperUtilities.cs:543, 715, 768 | Defensive crash on overflow | Worksheet-sized inputs to Unique, FlattenIntIndexes, or SortBlock's flag block crash with ArgumentOutOfRangeException ("'capacity' must be non-negative") instead of a clean error. | Line 543: new List<object>(Math.Min(rows * cols, 1024)) — rows * cols overflows for very large inputs; Math.Min(negative, 1024) = negative; new List<object>(negative) throws. Line 715: new List<int>(rows * cols) — same. Line 768: new List<bool>(fRows * fCols) — same. | Use long for the product, clamp: var hint = (int)Math.Min((long)rows * cols, 1024); | 0.65`

Not data corruption — defensive crash. Low.

---

## U-15 — Low — `UniqueCount` keeps a redundant `samples` dictionary mirroring `order`

`U-15 | Low | DeveloperUtilities.cs:585-625 | Wasted memory | For a block with K distinct values, samples stores K (string, object) pairs that duplicate information already in order (the strings) and reachable via "the first non-blank cell whose ToStringSafe(...) == key". | Line 591: var samples = new Dictionary<string, object>(StringComparer.Ordinal). Line 610: samples[key] = v. Line 621: result[i, 0] = samples[order[i]]. The order list could just be List<(string key, object sample)> or two parallel lists, dropping the dict. | Replace samples dict with a parallel List<object> firstSeenValues that grows in lockstep with order. Removes one dict + saves rehashes. | 0.95`

---

## U-16 — Low — `FillDown` treats `double.NaN` / scalar errors as legitimate fill values

`U-16 | Low | DeveloperUtilities.cs:262-286 | Edge-case data-quality risk | After CoerceNumeric (see U-07) lands a double.NaN, FillDown propagates that NaN down the rest of the column. | Line 281: `last = v;` updated whenever IsCellBlank(v) is false. IsCellBlank does not classify NaN or non-blank ExcelError sentinels as blank (correctly), but the user may expect "fill with the value above" to skip degenerate values. | Document that NaN/error sentinels are treated as data. Optionally add an opt-in argument to skip them. | 0.55`

Speculative on intent. Including because the file's "developer utilities" framing suggests data cleaning, and silent NaN propagation is a classic data-cleaning footgun.

---

## Rejected findings

The following concerns were considered and rejected:

- **"Array.Sort is unstable" / SortBlock is unstable** — REJECTED. Line 832 tiebreaks on `a.CompareTo(b)` (the original row index from `Enumerable.Range(0, rows).ToArray()`). Since indices are unique, equal-key rows are ordered by original index, which is the definition of stable. The Round-1 finding was wrong.

- **"FlattenIntIndexes (int)d wraps and reaches unsafe access"** — REJECTED for unsafe access. Traced both callers: `BlockLookup` line 668 has `(uint)idx >= (uint)lookupCols` bounds check; `SortBlock` line 816 uses `block[a, col]` whose array indexer throws on out-of-range. Neither path performs unsafe access. Downgraded to U-10 (Med — bad error UX, not memory safety).

- **"Sha256Block: TransformFinalBlock(Array.Empty<byte>(), 0, 0) doesn't finalize"** — REJECTED. Per .NET API docs, `HashAlgorithm.TransformFinalBlock(inputBuffer, inputOffset, inputCount)` correctly finalizes the hash even when inputCount is 0, populating the `Hash` property. Calling it once after all `TransformBlock` calls is the documented usage. Verified semantics: `TransformFinalBlock(Array.Empty<byte>(), 0, 0)` is the standard "finalize with no remaining bytes" idiom.

- **"FindReplace's per-call Regex compilation is a perf bug"** — REJECTED for severity. The XML doc explicitly states "Regex compilation is cached locally to the call to avoid a shared static cache, which would violate MTR safety guidance." This is a deliberate design choice for MTR-safety. Trade-off acknowledged in the code. Not a defect.

- **"SplitColumn keeps empty entries from consecutive delimiters"** — REJECTED. `s.Split(delimiter, StringSplitOptions.None)` is the documented contract (one column per delimiter-separated field). `"a,,b".Split(",")` → `["a","","b"]` is the standard behavior; users wanting `RemoveEmptyEntries` can pre-process or this becomes a feature request, not a defect.

- **"Unpivot rows == 1 returns object[0, outCols] which Excel renders as #VALUE!"** — REJECTED as a finding; this is the correct behavior. Header-only input has no data rows to unpivot. Excel-DNA marshals zero-length arrays to `#VALUE!` per documented Excel-DNA behavior. The user gets a clear signal that there's nothing to return. If the file's authors wanted to special-case this, they would; the current behavior is defensible.

- **"BlockLookup with 0 lookup rows: every key misses → #N/A"** — REJECTED. Confirmed correct behavior; matches the documented left-join semantic with no matches.

- **"BlockLookup line 678-681 is TOCTOU"** — REJECTED. Confirmed: the dict is a local variable, never escapes, no concurrency. Not a race. (See U-13 for the micro-opt.)

- **"FillBlanks with fillValue == ExcelError"** — REJECTED. The function takes any `object` and the contract doesn't restrict types; passing an error sentinel as the fill value is the user's choice. The documentation could be tightened but this isn't a defect.

- **"JoinColumns uses a single StringBuilder across rows — capacity reset cost"** — REJECTED. `StringBuilder.Clear()` keeps the underlying buffer; this is the recommended reuse pattern. The function is correctly optimized.

- **"RemoveDuplicateRows pre-sizes seen with default capacity"** — REJECTED. `new HashSet<string>(StringComparer.Ordinal)` lets the set grow naturally. For very-large inputs this causes a few extra rehashes; not a correctness issue, and the file already provides capacity hints elsewhere only where the size is known up front.

- **"Unique returns 1x1 ExcelEmpty for empty inputs — Excel-DNA marshaling"** — REJECTED. Confirmed correct: `new object[1, 1] { { ExcelEmpty.Value } }` is the idiomatic "single empty cell" return for Excel-DNA UDFs; renders as a blank cell on the grid.

- **"TrimBlock returns input string when it's already trimmed — could intern or memoize"** — REJECTED. Micro-opt with no clear benefit; trim is cheap, and strings in Excel are generally short and non-repeating.

- **"CoerceNumeric uses double.TryParse with both InvariantCulture and CurrentCulture (in Marshaling.TryToDouble) — culture-dependent parsing"** — Not a finding in *this* file; the parsing lives in `Marshaling.TryToDouble` which is outside scope. Noting only.

### Out-of-scope notes (Marshaling.cs, etc. — not filed as findings)

- `Marshaling.TryToDouble` (line 73-115) accepts "NaN"/"Infinity" via `double.TryParse + InvariantCulture` but rejects already-boxed `double.NaN`/Infinity in its `double` branch. That asymmetry is the root cause of U-07 surfacing in this file. Fix belongs in `Marshaling.cs`.
- `Marshaling.ToStringSafe` (line 124-142) collapses `1.0` (double) and `"1"` (string) to the same string; consumers (RemoveDuplicateRows, Unique, UniqueCount, HashBlock, Sha256Block, BlockLookup) inherit that conflation. Whether that's a bug or feature depends on intent.
