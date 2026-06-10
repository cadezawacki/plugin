# Audit plan — round 3

Rounds 1–2 audited and fixed the original 9-file core (commit `ce17fee`).
PR #3 then merged ~4,100 LOC of new, unaudited surface: conditional
aggregates, boosted lookups, text/regex/series/date/distance utilities,
JSON, filesystem, watch feeds, and result caches. Round 3 audits exactly
that new surface with the same methodology that worked in round 2 —
parallel domain agents, mandatory rejection appendices, independent
orchestrator re-verification, surgical fixes with a clean build after
each domain.

## Scope — files never audited before

| File | LOC | Notes |
| --- | --- | --- |
| `ConditionalAggregates.cs` | 1093 | *IFS family, weighted stats, GROUPBY |
| `JsonUtilities.cs` | 621 | JSONPATH/PARSEJSON + async file I/O |
| `CacheUtilities.cs` | 394 | session + disk caches (stateful) |
| `TextUtilities.cs` | 369 | case/pad/repeat/reverse/templatefill |
| `FileSystemUtilities.cs` | 361 | FILEINFO/READFOLDER/WATCH* |
| `SeriesUtilities.cs` | 302 | fillforward/outliers/quantiles |
| `RegexUtilities.cs` | 284 | block regex match/count/extract/split |
| `LookupBoost.cs` | 211 | XLOOKUPB hash/binary index |
| `DateUtilities.cs` | 191 | WORKDAYADD |
| `DistanceUtilities.cs` | 173 | pairwise distance matrices |
| `WatchFeeds.cs` | 103 | FileSystemWatcher RTD feeds |

Out of scope (already audited; seams only): `Marshaling.cs`,
`BulkTransfer.cs`, `AddIn.cs`, `DeveloperUtilities.cs`,
`ParallelUtilities.cs`, `VectorizedKernels.cs`, `RtdServer.cs`,
`ToolkitLifetime.cs`, `DirectFileIO.cs`.

## Domain partition

### Domain 6 — Conditional aggregation & lookups (~1,304 LOC)
**Files:** `ConditionalAggregates.cs`, `LookupBoost.cs`
Hot paths: criteria-mask construction, per-aggregate accumulation,
hash/sorted index build and probe. All `IsThreadSafe = true` — MTR
concurrency contract applies.

### Domain 7 — Block utilities: text, regex, series, dates, distance (~1,319 LOC)
**Files:** `TextUtilities.cs`, `RegexUtilities.cs`, `SeriesUtilities.cs`,
`DateUtilities.cs`, `DistanceUtilities.cs`
Hot paths: per-cell string transforms, regex over blocks, quantile sort,
workday loops, distance kernels routed through SIMD dot products.

### Domain 8 — JSON, filesystem, watch feeds, caches (~1,479 LOC)
**Files:** `JsonUtilities.cs`, `FileSystemUtilities.cs`, `WatchFeeds.cs`,
`CacheUtilities.cs`
Hot paths: JsonDocument traversal, folder ingest concatenation,
FileSystemWatcher event handling, cache dictionary access from
concurrent recalc threads, disk-cache file format round-trip.

## Cross-domain seams (annotate, do not chase)

- `Marshaling.TryToDouble` / `ToStringSafe` / `CellEquality` — Domain 1
  (round 2) owns the contract; consumers must honor it.
- `ToolkitLifetime.ShutdownToken` — consumed by Domain 8 async paths.
- `RtdServer.Feed` abstract contract (`Start`/`StopCore`/`LatestValue`)
  — `WatchFeeds.cs` implements it; the server side is already audited.
- Sync-over-async UDF bridging (`.GetAwaiter().GetResult()`) — pattern
  accepted in round 2 for `DirectFileIO`; flag only *new* deviations.
- `IsThreadSafe = true` + stateful caches — `CacheUtilities` is
  registered `IsThreadSafe = false`; verify the registration matches the
  implementation's actual thread-safety.

## Method — four passes per domain

1. **DEEP BUG SCAN** — correctness, races, leaks, silent corruption.
   Output `BUG-{severity}-{n}` with file:line, category, trigger, fix.
2. **MEMORY OPTIMIZATION** — allocations, copies, retention.
   Output `MEM-{n}` with current cost, change, expected reduction.
3. **CPU/THROUGHPUT** — hot-path speedups. Output `PERF-{n}` with
   hot-path flag, current cost, expected multiplier.
4. **ARCHITECTURAL WINS** — output `ARCH-{n}` with scope, impact,
   effort, risk.

Per round-2 rules: exact failure traces with citable file:line; trace
BCL contracts (not summaries); mandatory rejection appendix; the
orchestrator independently re-verifies every finding before any fix and
drops anything it cannot reproduce from source.

## Outputs

- `audit/reports/domain-6-aggregates.v3.md`
- `audit/reports/domain-7-blockutils.v3.md`
- `audit/reports/domain-8-json-fs-cache.v3.md`
- `audit/SYNTHESIS.v3.md` — verified findings, rejections, fix order.
- `audit/CHANGELOG.v3.md` — applied fixes.
- `audit/DEFERRED.v3.md` — confirmed but out-of-charter items.

## Acceptance criteria

- Release build x64 + x86, `0 Warning(s), 0 Error(s)` after every
  domain's fix commit (verified with `-p:EnableWindowsTargeting=true`
  on the Linux audit host; packing included).
- No public-surface behavioral change without a finding justifying it.
- Every fix maps 1:1 to a verified finding ID.
