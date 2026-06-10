# Round-3 changelog

Mapping of confirmed findings → files changed → one-line fix description.
All fixes are surgical; release builds for both x86 and x64 are clean with
zero warnings after every commit. Finding IDs refer to
`audit/reports/domain-{6,7,8}-*.v3.md`; the verification record is
`audit/SYNTHESIS.v3.md`.

## Cross-domain

| Finding | Files | Fix |
| --- | --- | --- |
| **CD-1** (upgrades round-2 deferrals U-07/F11) | `Marshaling.cs` | `TryToDouble`'s string path re-checks `double.IsFinite` after both parses: `"NaN"`, `"Infinity"`, `"1e999"` no longer leak non-finite doubles into sort-based consumers (QUANTILES, OUTLIERS, DISTANCE). |

## High

| Finding | Files | Fix |
| --- | --- | --- |
| **D6 BUG-HIGH-001** date criteria degrade to string compare | `ConditionalAggregates.cs` | `Criterion.Parse` falls back to `DateTime.TryParse` (CurrentCulture, digit-gated): `">1/1/2025"` compares serials numerically; time-only literals become day fractions. |
| **D7 BUG-HIGH-1** TEMPLATEFILL default never applies | `TextUtilities.cs` | `has_header_row` taken as `object` + `ResolveBool(…, true)` — Excel-DNA built-in registration ignores C# parameter defaults. |
| **D7 BUG-HIGH-2** Euclidean catastrophic cancellation | `DistanceUtilities.cs` | Pairs where `sq <= 1e-8·(a²+b²)` are recomputed by direct differences (verified: (1e8) vs (1e8+1) returned 0 instead of 1). |
| **D8 BUG-HIGH-1** one bad file kills READFOLDER | `FileSystemUtilities.cs` | Per-file try/catch (IO/UAC/Json/InvalidData): skipped with a trace entry; the rest of the folder still loads. |
| **D8 BUG-HIGH-2** ACL-denied subdir kills recursion | `FileSystemUtilities.cs` | `EnumerationOptions { IgnoreInaccessible = true }` replaces the legacy `SearchOption` overload. |
| **D8 BUG-HIGH-3** READJSON fails on UTF-8 BOM | `JsonUtilities.cs` | `JsonDocument.ParseAsync(FileStream)` replaces `Parse(ReadAllBytes)` — BOM tolerated, LOH byte[] gone (MEM-2). |

## Medium

| Finding | Files | Fix |
| --- | --- | --- |
| **D6 BUG-MEDIUM-002** `<>` never matches blanks | `ConditionalAggregates.cs` | Blank branch returns `_op == Op.NotEqual` (native COUNTIFS counts blanks for `"<>x"`). |
| **D6 BUG-MEDIUM-003** bools conflated with numbers | `ConditionalAggregates.cs` | `IsNumericCell` excludes `bool`; bool criteria match bool cells only (dedicated criterion kind). |
| **D6 BUG-MEDIUM-004** text criteria match numbers | `ConditionalAggregates.cs` | Text/wildcard criteria match `string` cells only; non-text satisfies only `<>` (COUNTIF(rng,"*") no longer counts numbers). |
| **D6 BUG-MEDIUM-005** XLOOKUPB cross-type key collisions | `LookupBoost.cs` | Typed `CellKey` (number/text/bool/error/blank) dictionary replaces `ToStringSafe` canonicalization; error lookup values propagate per cell. |
| **D6 BUG-MEDIUM-006** errors in matched values silently dropped | `ConditionalAggregates.cs` | Numeric aggregates, PERCENTILEIFS, WAVG(IFS), WMEDIAN, WSTDEV, SUMPRODUCTIFS propagate an `ExcelError` found in a matched cell. |
| **D6 BUG-MEDIUM-007** match_mode trap | `LookupBoost.cs`, `readme.md`, `docs/usage.md` | Documented bool form kept; XLOOKUP's `-1`/`0` accepted; other numeric modes rejected loudly instead of silently running exact. |
| **D7 BUG-MEDIUM-3** REPEAT silent `""` via int wrap | `TextUtilities.cs` | `count > 32767` rejected before the int cast. |
| **D7 BUG-MEDIUM-4** unbounded regex spill | `RegexUtilities.cs` | `Spill` caps: ≤16,384 columns, ≤Int32.MaxValue cells, checked before allocation. |
| **D7 BUG-MEDIUM-5** WORKDAYADD overflow/whole-call failure | `DateUtilities.cs` | Day counts beyond the serial span → per-cell `#NUM!`; the walk runs in bounded int-serial space (also PERF-1: 3.1× measured, bit-identical). |
| **D7 BUG-MEDIUM-6** CAMELCASE deletes astral letters | `TextUtilities.cs` | `ToProper`/`ToCamel` iterate `Rune`s; astral letters are cased, not dropped. |
| **D8 BUG-MEDIUM-4** junction/symlink cycles | `FileSystemUtilities.cs` | `AttributesToSkip = FileAttributes.ReparsePoint` (verified: looped tree returns 1 file, not 42). |
| **D8 BUG-MEDIUM-5** int64 > 2^53 silently rounded | `JsonUtilities.cs` | `NumberToCell`: int64 within 2^53 → exact double; bigger or non-finite → raw text (digits preserved). |
| **D8 BUG-MEDIUM-6** PARSEJSON defaults ignored | `JsonUtilities.cs` | `path`/`has_header_row` taken as `object` + `PathOf`/`ResolveBool`, matching the sibling file UDFs. |
| **D8 BUG-MEDIUM-7** disk-cache truncate-in-place | `CacheUtilities.cs` | Temp file + atomic `File.Move(overwrite)`; unparseable entries are a self-healing miss (file deleted), not a permanent `#VALUE!`. |
| **D8 BUG-MEDIUM-8 (case 3)** empty file pollutes header union | `FileSystemUtilities.cs` | 1×1 blank blocks skipped in both concat modes. (JSON object/scalar header plumbing deferred — see DEFERRED.) |
| **D8 BUG-MEDIUM-9** dead watcher after fatal FSW error | `WatchFeeds.cs` | Re-arm loop with 2 s backoff: bump on death, quiet retries, bump on regain. First-arm failure still propagates (documented `#N/A`). |

## Low

| Finding | Files | Fix |
| --- | --- | --- |
| **D6 BUG-LOW-008/009** GROUPBY key collisions / policy mismatch | `ConditionalAggregates.cs` | `AppendCellKey`: type-tagged, length-prefixed (injective) segments; grouping and `distinct` both case-insensitive, both type-tagged. |
| **D6 BUG-LOW-010** count-only shape validation | `ConditionalAggregates.cs` | `BuildMask(rows, cols, pairs)` requires criteria ranges to match the value range's shape. |
| **D6 BUG-LOW-011** nondeterministic approximate dup | `LookupBoost.cs` | Sort tiebreak by original index: last occurrence wins, deterministically. |
| **D6 BUG-LOW-012** error criteria never match | `ConditionalAggregates.cs` | `=`/`<>` criteria spelled as error text (or an error criterion object) match error cells. |
| **D6 BUG-LOW-013** multi-cell criteria silently match nothing | `ConditionalAggregates.cs` | 1×1 blocks unwrapped; larger blocks rejected with a clear message. |
| **D7 BUG-LOW-7** REVERSE emits ill-formed UTF-16 | `TextUtilities.cs` | Surrogate pairs re-swapped after reversal (single-allocation `string.Create`). |
| **D7 BUG-LOW-8** lone-surrogate pad char | `TextUtilities.cs` | Surrogate pad chars rejected. |
| **D7 BUG-LOW-9** regex UDFs mask input errors | `RegexUtilities.cs` | Input `ExcelError` cells pass through all five UDFs. |
| **D7 BUG-LOW-10** cosine diagonal ±1e-16 | `DistanceUtilities.cs` | Self-comparison diagonal written as exact 0. |
| **D7 BUG-LOW-11** error matrix_b treated as omitted | `DistanceUtilities.cs` | Scalar-error `matrix_b` propagates via `ErrorBlock`. |
| **D7 BUG-LOW-12** TEMPLATEFILL no cell cap | `TextUtilities.cs` | Renders > 32,767 chars embed per-cell `#VALUE!`. |
| **D8 BUG-LOW-10** watch on relative path = process CWD | `WatchFeeds.cs` | `Path.IsPathFullyQualified` required; loud error otherwise. |
| **D8 BUG-LOW-11** NDJSON failures report line 0 | `JsonUtilities.cs` | Real file line number wrapped into the rethrown `JsonException`. |
| **D8 BUG-LOW-12** unvalidated `(ExcelError)` cast | `CacheUtilities.cs` | `Enum.IsDefined` guard; undefined payloads fall back to raw text. |
| **D8 BUG-LOW-13** duplicate headers drop data | `FileSystemUtilities.cs` | Occurrence-aware union: duplicate same-name columns stay separate output columns. |

## Memory / CPU (applied)

| Finding | Files | Fix |
| --- | --- | --- |
| D6 MEM-1/2/3, PERF-1/2 | `ConditionalAggregates.cs` | Mask + aggregate paths iterate `object[,]` in place (no flattened copies of criteria/value ranges: −1.6..2.4 MB per 100k-row call); fast `is double` mask loop for plain numeric criteria; `matched` list no longer pre-sized to the full range. |
| D6 MEM-4, PERF-8 | `ConditionalAggregates.cs` | GROUPBY key-row arrays allocated per group (was per row: 3.2 MB garbage/100k rows); operation parsed once per call. |
| D6 PERF-3/4/7, MEM-5 | `LookupBoost.cs` | Sortedness pre-check (47.6 ms → 0.21 ms measured on presorted 100k); `TryAdd` single-probe insert; ≤4 probes use a linear scan (~60× for the copied-down shape); typed keys remove ~12.6 MB of string churn per 100k×100k call. |
| D6 PERF-5 (partial), PERF-6 | `ConditionalAggregates.cs` | Median/percentile sort in place via `CollectionsMarshal.AsSpan` (no `ToArray` copy); WSTDEV collects typed pairs once (4 → 2 conversions/row). |
| D7 MEM-1 | `RegexUtilities.cs` | `Regex.Count` replaces `Matches().Count`: 33.6 MB → ~0 of match machinery per 100k cells (measured). |
| D7 PERF-2 | `RegexUtilities.cs` | `RegexOptions.Compiled` for blocks ≥ 16,384 cells (4.9× matching measured; compile cost amortized within the call). |
| D7 MEM-2/PERF-3 | `DistanceUtilities.cs` | Self case: one flatten, one triangle computed and mirrored (1.95× kernel, 2× fewer boxes), branch-split manhattan/chebyshev loops. |
| D7 MEM-3/4 | `TextUtilities.cs` | Shared per-call StringBuilder for PROPER/CAMELCASE/REPEAT; REVERSE via `string.Create` (13.7 → 6.8 MB per 100k cells measured). |
| D7 MEM-5, PERF-5 | `SeriesUtilities.cs` | OUTLIERS parses each cell once (vals/ok arrays) and reuses the numeric buffer for MAD deviations; QUANTILES sorts the list backing store in place. |
| D7 PERF-1 | `DateUtilities.cs` | Int-serial workday walk (16 → ~5 ns/step measured, bit-identical results). |
| D8 MEM-1, PERF-5 | `JsonUtilities.cs` | NDJSON / array-of-objects rows are column-indexed `object[]`s (~600 → ~104 B/row at 10 keys); table fill is a bounds check + array read instead of a string-hash probe per cell. |
| D8 MEM-2 | `JsonUtilities.cs`, `CacheUtilities.cs` | Stream-based `JsonDocument.ParseAsync`: no whole-file LOH `byte[]` per read. |
| D8 MEM-3, PERF-1 | `JsonUtilities.cs` | JSONPATH compiles the path once per call (was ~100 B + ~300 ns re-tokenization per cell). |
| D8 MEM-4, PERF-3 | `FileSystemUtilities.cs` | Concat releases consumed blocks; header mapping via dictionary (was `List.IndexOf` per column per file). |
| D8 PERF-2 | `FileSystemUtilities.cs` | Bounded-parallel file reads (gate = min(8, cores)), deterministic output order. |
| D8 PERF-4 | `FileSystemUtilities.cs` | FILEINFO stats once via `FileInfo` (3 → 1 syscalls for files). |
| D8 PERF-6 | `CacheUtilities.cs` | Batch eviction (64 oldest per scan); cache directory ensured only on writes. |

## Build verification

After each domain's fixes, a release build was re-run; final state on the
audit host (`-p:EnableWindowsTargeting=true`, XLL packing included):

```
dotnet build -c Release -p:Platform=x64 -> 0 Warning(s), 0 Error(s)
dotnet build -c Release -p:Platform=x86 -> 0 Warning(s), 0 Error(s)
```

## Behavioral notes (intentional changes)

- The *IFS family now mirrors native Excel more closely: blanks count for
  `<>`, text criteria no longer match numbers/bools, logicals are not
  numbers, matched error cells propagate, criteria ranges must match the
  value range's shape. Workbooks that relied on the previous (divergent)
  semantics will see different — native-consistent — results.
- `EPT.XLOOKUPB` exact match is now type-aware (numeric 5 ≠ text "5") and
  propagates error lookup values; `match_mode` accepts XLOOKUP's `0`/`-1`.
- `EPT.READFOLDER` skips unreadable files (with a trace) instead of failing
  the whole call, skips reparse points, and ignores inaccessible
  subdirectories.
- `EPT.READJSON`/`EPT.READNDJSON`/`EPT.JSONPATH`/`EPT.PARSEJSON` preserve
  >2^53 integers as text instead of silently rounding them.
- `EPT.WATCHFILE`/`EPT.WATCHFOLDER` require fully qualified paths and
  re-arm automatically after fatal watcher errors.
- Disk-cache entries are written atomically; a corrupt entry reads as a
  miss and is deleted.
