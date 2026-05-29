# Deferred — round 2

Findings whose fix would require an architectural change beyond surgical
scope. Each is real and confirmed; each is listed here rather than in the
synthesis-to-fix track because applying the fix would rewrite a kernel,
introduce a new API surface, or otherwise extend beyond the audit charter.

## F6 (concurrency) — `ParallelRowReduce` allocates `[rows, cols]` to return one column

**Location:** `ParallelUtilities.cs:429-444`

**Problem:** The UDF allocates an `expanded` double-matrix of shape
`[rows, cols]` only to read column 0 into the result. For large inputs this
is gigabytes of throwaway memory.

**Why deferred:** A surgical fix would either:
1. Rewrite the helper signature to accept a different output shape (changes
   the `RowTransform` delegate contract).
2. Switch from `ParallelBatchTransform` to a `Parallel.For` per-row loop with
   its own buffer, duplicating the partitioned-buffer logic.

Either path is a real refactor of a public API. The current code is correct
(produces the right answer); it just over-allocates. Recommend a follow-up
ticket to add a `ParallelRowReduceCore(double[,], int, RowReducer)` helper
that writes directly to the single-column output.

## F6 (file I/O) — Unbounded CSV read OOMs Excel

**Location:** `DirectFileIO.cs:74, 173`

**Problem:** `ReadDelimitedAsync` keeps every row in a `List<object[]>` and
finally allocates `new object[rows.Count, maxCols]`. A multi-GB CSV will OOM
the Excel process without any backpressure.

**Why deferred:** Excel's UDF model is fundamentally one-bulk-array-return.
Streaming a CSV into a worksheet without buffering the full result would
require either:
1. A new RTD-style streaming UDF that pushes rows as they arrive (large
   surface change).
2. A documented hard row cap with a clear `#NUM!`-style error above it
   (user-facing behavior change).

Recommend documenting the limit in the readme/usage docs and offering
option 2 as a follow-up.

## RTDv2-013 — Per-tick O(N) scan over all topics

**Location:** `RtdServer.cs:118-139`

**Problem:** The throttle timer walks every subscribed topic every 250 ms,
checking each `Feed.LatestValue` regardless of whether the feed produced a
new value. At 50k topics this is 200k boxed `Equals` calls per second on
idle.

**Why deferred:** A per-feed dirty-set would require:
1. Each `Feed` to track a monotonic version counter on `LatestValue` writes
   (cheap, but invasive across every feed implementation).
2. `FlushTick` to walk a single dirty-flag set instead of `_topics`.

Worthwhile optimization at scale; not a correctness issue. Recommend a
follow-up after the critical/high fixes settle.

## RTDv2-015 — `Stop()` doesn't join the producer Task

**Location:** `RtdServer.cs:355-371` (now `StopLocked` post-fix)

**Problem:** `Stop()` cancels the producer's CTS but does not `await` the
Task. The Task may continue running briefly while a new Subscribe creates
a fresh producer for the same Feed instance. Two producers (the orphan +
the new one) can write `LatestValue` until the orphan observes
cancellation.

**Why deferred:** A correct fix requires either making `Stop` async (changes
the public contract; callers don't expect `await`) or blocking the calling
thread on `_producer.Wait()` (risky under lock; could deadlock if called
from the producer's own continuation). The orphan window is small (typically
< 1 throttle period, ~250 ms) and the orphan writes are to a soon-to-be-GC'd
Feed instance — no Excel visibility.

Recommend a follow-up that splits `Stop` into a synchronous "request stop"
and an awaitable "join" pair.

## V009 — `TryToDouble` doesn't cover `uint`/`short`/`ushort`/`byte`/`sbyte`/`ulong`

**Location:** `Marshaling.cs:73-115`

**Problem:** These integer types fall through to `default` and return false
even though they trivially convert. Excel-DNA never produces them on its
own, but other .NET callers wiring through `AsArray2D` can.

**Why deferred:** Adding six cases is mechanical but the surface keeps
growing. A cleaner fix is `IConvertible.ToDouble(InvariantCulture)` for
unknown numeric types — but that adds a runtime type check and broadens the
contract. The existing fallback is conservative (returns false rather than
silently lossy). Recommend evaluating whether to broaden the contract in a
separate PR.

## V010 — `BoxDoubleMatrix` doesn't guard NaN/Infinity going to Excel

**Location:** `Marshaling.cs:255-269` area

**Problem:** A `double.NaN` or `±Infinity` written via
`ExcelReference.SetValue` is rendered by Excel as `#NUM!` silently. Callers
that meant a real number get a visible error with no log line.

**Why deferred:** Filtering NaN/Inf at marshaling-out time would change
behavior for code that intentionally writes those values (e.g. statistical
outputs). A safer approach is a separate `WriteBlockSafe` variant that maps
NaN/Inf to `ExcelError.ExcelErrorNum` explicitly. Cosmetic.

## Documentation-only items

These are real but resolved by clarifying docs rather than code changes:

- **V012** Lossy `SetValue` error — diagnostic only.
- **V014** AutoClose-without-AutoOpen leaks one CTS — one-time, until
  process exit.
- **V015** ErrorToText returns `#ERR` for future error sentinels.
- **V016** float-to-double round-trip via ToString.
- **V017** Trace contains raw input strings.
- **U-07** "NaN"/"Infinity" text coerces to a non-finite double.
- **U-08** RemoveDuplicateRows doc claims CellEquality, code uses
  ToStringSafe.
- **U-11** Regex (CultureInvariant) vs non-regex (Ordinal) asymmetry.
- **U-12** Transpose cache thrashing.
- **U-13** ContainsKey+indexer vs TryAdd.
- **U-14** List capacity uses unchecked multiplication (capacity-only; no
  correctness impact).
- **U-15** UniqueCount has a redundant `samples` dict.
- **U-16** FillDown propagates NaN.
- **F10 (file I/O)** Delimiter as a high surrogate.
- **F11 (file I/O)** "Infinity"/"NaN" coerced to a non-finite double.
- **F13–F19 (file I/O)** BOM detection asymmetry, FileOptions hint on
  write, two-pass copy, generic exception trace, `\\t` literal, per-row
  await allocations, ResolveBool's narrow truthy set.
- **RTDv2-010, -012, -014, -019–025** RTD low-severity items: stale
  Topic.UpdateValue traces, Excel-DNA queue backpressure, registration on
  disposed CTS edge, infinite producer restart from a buggy custom feed,
  boxing per Equals, malformed sine/counter spec swallow, no initial value
  in ClockFeed, SineFeed wall-clock skew, RandomFeed unnecessary lock.

These are tracked in the per-domain reports and intentionally not patched in
this round.
