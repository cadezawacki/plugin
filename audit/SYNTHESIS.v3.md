# Synthesis — round 3

Orchestrator reading of `audit/reports/domain-{6,7,8}-*.v3.md`. Every finding
below was independently re-verified by re-reading the cited source line and,
where the claim rests on BCL behavior, re-running the probe on .NET 8.0.422
(`/tmp/probe-orch`, `/tmp/probe6..8`). Findings that did not survive scrutiny
are in the rejected section.

## Verification highlights (orchestrator probes)

- `JsonDocument.Parse(byte[] with EF BB BF)` → `JsonException` (BOM bug real);
  `JsonDocument.ParseAsync(Stream)` succeeds on identical bytes.
- Euclidean-by-identity on (1e8) vs (1e8+1) rows → **0.0** where the true
  distance is 1.732 (100 % silent error).
- `(int)1e300 == int.MinValue`; `Math.Abs(int.MinValue)` throws
  `OverflowException` (REPEAT / WORKDAYADD overflow paths real).
- `string.Compare("43997", "1/1/2025", OrdinalIgnoreCase) > 0` (date-criteria
  degradation real).
- Excel-DNA 1.8 built-in registration (no `ExcelDna.Registration` in the
  generated `.dna`) ignores C# default parameter values → omitted `bool`
  arrives as `false` (TEMPLATEFILL / PARSEJSON default bug real).
- `double.TryParse("1e999"/"NaN"/"Infinity")` succeeds with non-finite results
  on .NET 8 → `Marshaling.TryToDouble` violates its own "finite" contract on
  the string path (upgrades round-2 deferral U-07/F11 to a code fix now that
  sort-based consumers exist: QUANTILES/OUTLIERS/DISTANCE).

## Confirmed — fix now (cross-domain)

| ID | Location | Fix |
| --- | --- | --- |
| **CD-1** (U-07/F11 upgrade) | `Marshaling.cs:127-132` | Re-check `double.IsFinite(result)` after both string parse passes. |

## Confirmed — fix now (domain 6: ConditionalAggregates, LookupBoost)

| ID | Sev | Fix |
| --- | --- | --- |
| BUG-HIGH-001 | HIGH | Date-literal criteria: `DateTime.TryParse(CurrentCulture)` fallback to a numeric (serial) criterion. |
| BUG-MEDIUM-002 | MED | Blank cells satisfy `<>` criteria (`return _op == Op.NotEqual` in the blank branch). |
| BUG-MEDIUM-003 | MED | Bools excluded from `IsNumericCell`; bool criteria match bool cells only. |
| BUG-MEDIUM-004 | MED | Text/wildcard criteria match string cells only; non-text satisfies only `<>`. |
| BUG-MEDIUM-005 | MED | XLOOKUPB: type-tagged `CellKey` (number/text/bool/error/blank) replaces string canonicalization; error lookup values propagate. |
| BUG-MEDIUM-006 | MED | Errors in matched value cells propagate from numeric aggregates (and weighted/sumproduct/percentile paths), as native. |
| BUG-MEDIUM-007 | MED | `match_mode`: keep documented bool form; add XLOOKUP `-1` (approximate) / `0` (exact); other numerics rejected loudly. |
| BUG-LOW-008/009 | LOW | GROUPBY/distinct keys: type-tagged + length-prefixed segments (injective), case-insensitive for text in both. |
| BUG-LOW-010 | LOW | Criteria ranges validated by shape, not just count (BuildMask takes rows×cols). |
| BUG-LOW-011 | LOW | Approximate-lookup sort gets an index tiebreak (deterministic last-occurrence). |
| BUG-LOW-012 | LOW | `=`/`<>` criteria spelled as error text (`"#N/A"`) match error cells. |
| BUG-LOW-013 | LOW | Multi-cell criteria argument: 1×1 unwrapped; larger → loud `#VALUE!`. |
| MEM-1/2/3, PERF-1/2 | — | Mask + aggregate paths iterate `object[,]` in place (no flatten copies); fast all-double mask loop; `matched` list default capacity. |
| MEM-4 | — | GROUPBY key-row array allocated only on first-seen keys. |
| PERF-3 | — | ApproximateLookup: O(M) sortedness pre-check before sorting. |
| PERF-4 | — | `TryAdd` instead of ContainsKey+indexer (with CellKey dictionary). |
| PERF-5 (partial) | — | Median/Percentile sort in place via `CollectionsMarshal.AsSpan` (no `ToArray` copy). Quickselect deferred. |
| PERF-6 | — | WSTDEV collects typed pairs once (no 4× re-conversion). |
| PERF-7 | — | ≤4 probes → linear scan instead of full index/sort. |
| PERF-8 | — | GROUPBY normalizes the operation string once per call. |

## Confirmed — fix now (domain 7: Text/Regex/Series/Date/Distance)

| ID | Sev | Fix |
| --- | --- | --- |
| BUG-HIGH-1 | HIGH | TEMPLATEFILL `has_header_row` → `object` + ResolveBool(default TRUE) (built-in registration ignores C# defaults). |
| BUG-HIGH-2 | HIGH | Euclidean: detect cancellation (`sq <= 1e-8·(a²+b²)`) and recompute that pair by direct differences. |
| BUG-MEDIUM-3 | MED | REPEAT rejects `count > 32767` before the int cast (kills the int.MinValue wrap → silent `""`). |
| BUG-MEDIUM-4 | MED | Spill caps: ≤16,384 columns and ≤Int32.MaxValue cells before allocation. |
| BUG-MEDIUM-5 | MED | WORKDAYADD: serial-range guard (±2,958,465) + bounded int-serial walk; per-cell `#NUM!` instead of whole-call failure. |
| BUG-MEDIUM-6 | MED | PROPER/CAMELCASE rune-based word/casing logic (astral letters no longer dropped/miscased). |
| BUG-LOW-7 | LOW | REVERSE re-swaps surrogate pairs after reversal (no more ill-formed UTF-16). |
| BUG-LOW-8 | LOW | Pad char that is a surrogate is rejected loudly. |
| BUG-LOW-9 | LOW | REGEXMATCH/COUNT/EXTRACT(/ALL/SPLIT) propagate input error cells instead of coercing to FALSE/0/"". |
| BUG-LOW-10 | LOW | Cosine self-diagonal exactly 0 (with the symmetry fast path). |
| BUG-LOW-11 | LOW | DISTANCE propagates a scalar-error `matrix_b` instead of treating it as omitted. |
| BUG-LOW-12 | LOW | TEMPLATEFILL embeds per-cell `#VALUE!` for >32,767-char renders. |
| MEM-1 | — | REGEXCOUNT uses `Regex.Count` (no Match materialization; measured 33.6 MB→~0/100k cells). |
| MEM-2 | — | DISTANCE self case reuses the flattened buffer (no duplicate flatten). |
| MEM-3 (partial) | — | PROPER/CAMELCASE/REPEAT share one StringBuilder per call. |
| MEM-4 | — | REVERSE via `string.Create` (one allocation per cell, surrogate fix-up included). |
| MEM-5 | — | OUTLIERS/QUANTILES sort in place (`CollectionsMarshal.AsSpan`); MAD reuses the buffer. |
| PERF-1 (stage 1) | — | WORKDAYADD walks int serials + mod-7 mask index (measured 3.1×, bit-identical). Closed-form jumps deferred. |
| PERF-2 | — | `RegexOptions.Compiled` when the block has ≥16,384 cells (measured 4.9× matching). Cross-call cache deferred. |
| PERF-3 | — | DISTANCE self case computes one triangle and mirrors the boxed value (1.95×). |
| PERF-5 | — | OUTLIERS parses each cell once (vals/ok arrays), not twice. |

## Confirmed — fix now (domain 8: JSON/FS/Watch/Cache)

| ID | Sev | Fix |
| --- | --- | --- |
| BUG-HIGH-1 | HIGH | READFOLDER: per-file fault isolation (skip + trace unreadable/malformed files). |
| BUG-HIGH-2 | HIGH | Enumeration via `EnumerationOptions { IgnoreInaccessible = true }` (ACL-denied subdir no longer kills the call). |
| BUG-HIGH-3 | HIGH | READJSON parses via `JsonDocument.ParseAsync(Stream)` (UTF-8 BOM tolerated; LOH byte[] gone — MEM-2). Same change in the disk-cache reader. |
| BUG-MEDIUM-4 | MED | `AttributesToSkip = ReparsePoint` (no junction/symlink cycles). |
| BUG-MEDIUM-5 | MED | JSON numbers: int64 ≤ 2^53 exact; bigger/out-of-range preserved as raw text (no silent low-digit corruption / Infinity). |
| BUG-MEDIUM-6 | MED | PARSEJSON `path`/`has_header_row` → `object` + PathOf/ResolveBool (defaults actually apply). |
| BUG-MEDIUM-7 | MED | Disk cache: write to temp + atomic `File.Move(overwrite)`; unparseable entry = miss (self-healing, file deleted). |
| BUG-MEDIUM-8 (case 3) | MED | ConcatByHeader skips 1×1 blank blocks (empty files no longer inject a phantom "" column). Object/scalar-JSON header plumbing deferred (ARCH-1). |
| BUG-MEDIUM-9 | MED | Watch feeds re-arm after fatal watcher errors (bounded 2 s backoff, single bump on regain); initial missing-directory still surfaces as before. |
| BUG-LOW-10 | LOW | WATCHFILE/WATCHFOLDER reject non-fully-qualified paths loudly (no more watching the process CWD). |
| BUG-LOW-11 | LOW | READNDJSON reports the failing line number. |
| BUG-LOW-12 | LOW | Disk-cache `$err` values validated via `Enum.IsDefined`. |
| BUG-LOW-13 | LOW | ConcatByHeader keeps duplicate same-name columns per file (occurrence-aware union; no silent data drop). |
| MEM-1/PERF-5 | — | NDJSON/array-of-objects rows stored as column-indexed `object[]` (no per-row Dictionary; ~6× less transient). |
| MEM-3/PERF-1 | — | JSONPATH path compiled once per call. |
| MEM-4 | — | Concat releases consumed per-file blocks. |
| PERF-2 | — | READFOLDER reads files with bounded parallelism (gate 8), order preserved. |
| PERF-3 | — | Header mapping via dictionary, not `List.IndexOf`. |
| PERF-4 | — | FILEINFO stats once via `FileInfo`, falls back to `DirectoryInfo`. |
| PERF-6 | — | Batch eviction (64 oldest per scan); `DiskPath` only ensures the directory on writes. |

## Rejected / no action (orchestrator concurrence with agent appendices)

The three rejection appendices (28 + 27 + 23 entries) were spot-checked; the
orchestrator specifically re-confirmed and concurs with: no MTR races in any
new file (no shared mutable state beyond thread-safe TraceSource /
ConcurrentDictionary); `params` trailing-missing handling; PercentileInc ==
PERCENTILE.INC (probe-verified by domain 7); FloorIndex boundaries; serial-60
WORKDAY agreement with Excel's phantom-day handling (probe-verified);
JsonDocument lifetime discipline (no element escapes its `using` scope);
sync-over-async bridges all `ConfigureAwait(false)` (accepted round-2
pattern); SHA-256 disk-cache filenames (no traversal); `WMedian` tie
behavior; `Memoize` aliasing.

Notable judgment calls:

- **D6 BUG-MEDIUM-007**: the inverted `match_mode` is *documented*, so the
  documented bool behavior is kept; the fix adds the XLOOKUP numeric forms and
  rejects unsupported ones loudly. Docs updated.
- **D8 BUG-MEDIUM-9**: full self-healing (including pre-creation watch) was
  trimmed to "re-arm after fatal errors"; initially-missing directories keep
  the documented immediate-failure behavior.
- **Numeric-looking text in OUTLIERS/QUANTILES/DISTANCE** (vs the *IFS
  family's genuine-numeric rule): both agents rejected this as documented
  toolkit-wide `TryToDouble` semantics; concur, no change.

## Deferred (architectural) — see DEFERRED.v3.md

D6 ARCH-1..4 (batched *IFS, index cache, shared CellKey unification,
streaming accumulators; quickselect), D7 ARCH-1..4 (regex cache, closed-form
workday engine, central spill governance, span text kernels; SIMD
manhattan/chebyshev), D8 ARCH-1..4 (full READFOLDER resilience surface,
shared atomic-write helper for WRITEJSON/WRITECSV, cache governance,
self-healing pre-creation watches), READFOLDER object/scalar-JSON header
plumbing (D8 BUG-MEDIUM-8 cases 1-2), D8 BUG-LOW-14 cache growth bounds.

## Fix order

1. `Marshaling.cs` (CD-1) — contract fix first; all domains consume it.
2. `ConditionalAggregates.cs` + `LookupBoost.cs` (domain 6).
3. `TextUtilities.cs`, `RegexUtilities.cs`, `SeriesUtilities.cs`,
   `DateUtilities.cs`, `DistanceUtilities.cs` (domain 7).
4. `JsonUtilities.cs`, `FileSystemUtilities.cs`, `WatchFeeds.cs`,
   `CacheUtilities.cs` (domain 8).
5. Doc updates (readme/usage for XLOOKUPB match modes).

Release x64 build after each step; x64 + x86 after the last.
