# Synthesis — round 2

Orchestrator reading of all five `audit/reports/domain-*.v2.md` reports.
Independent verification done by re-reading the cited source and, where
useful, running the .NET BCL against the input in question to confirm
behavior (e.g. V001's `double.TryParse` on `"1,5"`).

## Verification process

For each finding the orchestrator:

1. Re-opened the file at the cited line.
2. Confirmed the failure scenario is reachable from real input on real
   .NET 8 / Excel-DNA semantics.
3. For numeric / parsing claims, ran the BCL call against the documented
   input in a throwaway console (`/tmp/v001`).
4. Downgraded or rejected anything not surviving scrutiny.

## Confirmed findings — Critical

| ID | Location | Root cause | Blast radius | Chosen fix |
| --- | --- | --- | --- | --- |
| **V001** | `Marshaling.cs:104-110` | `double.TryParse` with `NumberStyles.AllowThousands` against `InvariantCulture` interprets `,` as a thousands separator. `"1,5"` (German/French/Spanish decimal) parses to `15.0`. **Empirically verified** in `/tmp/v001`. | Every UDF that coerces strings to doubles: SIMD kernels, RTD producers consuming text, `CoerceNumeric`, `PolyEval` coefficients, lookup keys parsed as numbers. Silent 10x error throughout the toolkit on non-English Excel data. | Drop `AllowThousands` from the invariant pass: invariant numbers don't carry group separators. Keep it for the current-culture fallback. |
| **U-01** | `DeveloperUtilities.cs:418` | `new Regex(find, RegexOptions.CultureInvariant)` with no `MatchTimeout` and no `NonBacktracking`. User-supplied pattern can be a ReDoS pathological case. | Hangs the Excel calc thread on any pathological pattern. The thread is the main calc thread for `EPT.FINDREPLACE` (`IsThreadSafe = false`), so the entire workbook hangs. | Add `TimeSpan.FromSeconds(1)` `MatchTimeout` to the `Regex` constructor; catch `RegexMatchTimeoutException` and surface `#VALUE!` with a logged reason. |
| **RTDv2-001 / RTD-001** | `RtdServer.cs:73-79` | `Timer.Dispose()` returns immediately; `FlushTick` may be mid-iteration on a ThreadPool thread when `_topics.Clear()` and `FeedManager.Shutdown()` run. `Topic.UpdateValue` on a torn-down topic crosses the COM boundary unpredictably. | Add-in unload races a running tick → potential Excel-DNA marshaler exceptions, log noise, and (worst case) a Topic reference used after Excel-DNA freed it. | Use `Timer.Dispose(WaitHandle)` to join the in-flight callback before clearing `_topics`. |
| **RTDv2-002 / RTD-002** | `RtdServer.cs:122-138` | The outer `try/catch` wraps the entire `foreach`. One topic whose `UpdateValue` throws aborts the rest of the tick. | Any deterministic-throw topic permanently starves every other topic of updates after the first tick. Production impact: a single bad cell silently freezes the RTD pipeline. | Move the `try/catch` inside the loop body, scoped to one topic's `UpdateValue + LastPushed` pair. |
| **F1 (file I/O)** | `DirectFileIO.cs:128-132` | Opening-quote branch sets `inQuotes = true; continue;` without resetting `lastWasCR`. After a `\r`, a quoted field leaves `lastWasCR=true`; the next `\n` is swallowed as the second half of a phantom CRLF. **Traced** for `a\r"b"\nc\n` → emits `[["a"],["bc"]]` instead of 3 rows. | Any CSV with `\r`-terminated lines that contain quoted fields. Data corruption: rows merge, fields concatenate. | Add `lastWasCR = false;` to the opening-quote branch. Also apply the same fix to the delimiter branch (F3) and reset `lastWasCR` in the default/append branch already (line 156) — verify the four state branches all clear it. |

## Confirmed findings — High

| ID | Location | Root cause | Chosen fix |
| --- | --- | --- | --- |
| **V002** | `Marshaling.cs:101-103` | `DateTime.ToOADate()` throws `OverflowException` for years 1-99 (excluding `MinValue`). **Empirically verified.** `TryToDouble` contract is "best-effort, no throws". | Wrap in `try/catch (OverflowException)`. |
| **V003** | `BulkTransfer.cs:46-55` | `XlCall.Excel(xlfEvaluate)` throws `XlCallException` on certain non-success returns. `ResolveRange` lets it escape; C# callers expect `ArgumentException`. | Catch `XlCallException`, throw `ArgumentException` with inner. |
| **V004** | `AddIn.cs:62-73` | `xlfGetWorkspace` can return `ExcelError` boxed in `object`; `v?.ToString()` then logs the boxed enum's text as "Excel version". | Type-check for `ExcelError` before stringifying. |
| **V006** | `Marshaling.cs:104-110` | `double.TryParse` is O(n) over string length, with no cap; multi-MB cell strings cause runaway parsing. | Length cap: bail out for strings longer than 64 chars before parsing. (Real numbers don't need more.) |
| **V007** | `BulkTransfer.cs:120-131` | Multi-cell anchor's `RowLast/ColumnLast` are silently ignored; the write expands past the anchor without warning. | If anchor is multi-cell, require it to match the block; else require single-cell. |
| **F1 (concurrency, kernels)** | `VectorizedKernels.cs:70` | `new double[rows * cols]` uses unchecked int multiply. 1M x 16K wraps. | Compute the product as `long`, throw `ArgumentException` if it exceeds `int.MaxValue / sizeof(double)`. |
| **F2 (concurrency)** | `VectorizedKernels.cs:89` | Same overflow in `BoxFlatDoubles`'s length check. | Same long-cast guard. |
| **F3 (concurrency)** | `VectorizedKernels.cs:637,641,649` | `MatrixMultiply` int math for `m * k`, `k * n`, `m * n`. | Long-cast bounds check at entry. |
| **F4 (concurrency)** | `VectorizedKernels.cs:582-589` | `L2Normalize`: if `sumSq` overflows to `+Inf`, `norm = sqrt(Inf) = Inf`; dividing by Inf yields zeros while the function returns Inf. Silent corruption. | After computing `norm`, treat non-finite norm same as zero norm: clear destination and return 0. |
| **F5 (concurrency)** | `ParallelUtilities.cs:466-467` | `Dot`: `na = ra*ca`, `nb = rb*cb` int overflow; `na != nb` check may pass under matched overflow. | Long-cast before comparing; reject ranges larger than `int.MaxValue`. |
| **F6 (concurrency)** | `ParallelUtilities.cs:433-444` | `ParallelRowReduce` allocates `[rows, cols]` (potentially GB) just to read column 0. | Allocate `[rows, 1]` only; project the body to write to a single-cell output. (Real refactor; tag DEFERRED if non-surgical.) |
| **U-02** | `DeveloperUtilities.cs:864-867, 902-906` | Hash separator bytes `0x1E`/`0x1F` may appear inside cell content, producing collisions. | Either escape the bytes when they appear in the text, or use a length-prefix encoding ahead of each cell. |
| **U-03** | `DeveloperUtilities.cs:185` | `BuildRowKey` uses `\x1F` as field separator; same collision risk. | Replace `Append('\x1F')` with `Append(length).Append('|').Append(text)` style or use a length-prefix prefix. |
| **U-04** | `DeveloperUtilities.cs:466` | `StackColumns`: `new object[rows * cols, 1]` int overflow. | Long-cast guard. |
| **U-05** | `DeveloperUtilities.cs:502` | `Unpivot`: `outRows = dataRows * valueColumns` int overflow. | Long-cast guard. |
| **U-06** | `DeveloperUtilities.cs:757-792` | `descendingFlags` typed `object`; only `object[,]` branch implemented. Scalar `TRUE` from Excel marshals as `bool`, silently ignored. | Accept `bool`/`double`/`string` scalar in addition to `object[,]`; broadcast to `keys.Length`. |
| **F2 (file I/O)** | `DirectFileIO.cs:128-132` | A quote in the middle of an unquoted field silently turns on quote mode. | Treat `"` inside an unquoted field as a literal character (append, not state-change), OR raise an error. RFC 4180 says don't allow it; the safer surgical fix is to append as literal. |
| **F3 (file I/O)** | `DirectFileIO.cs:133-137` | Delimiter branch doesn't reset `lastWasCR`. `\r,\n` eats the LF. | Same one-liner as F1: reset `lastWasCR = false` on entry. |
| **F4 (file I/O)** | `DirectFileIO.cs:115-123` | `reader.Peek()` and `reader.Read()` block the awaiting thread at chunk boundary. | Replace with an explicit lookahead char carried across reads; never call sync `Peek`/`Read`. (Mild refactor.) |
| **F5 (file I/O)** | `DirectFileIO.cs:422-441` | `ResolveEncoding` catches only `ArgumentException`; legacy codepages throw `NotSupportedException` (missing `CodePagesEncodingProvider`). | Catch both, fall back to UTF-8. |
| **F6 (file I/O)** | `DirectFileIO.cs:74, 173` | Unbounded `List<object[]> rows` and final `object[rows.Count, maxCols]`. | Document the contract: bulk read holds full file. Hard cap optional; defer until requested. |
| **RTDv2-003** | `RtdServer.cs:326-350` | Subscribe/Unsubscribe mutate `_subscribers` outside `_gate`. | Move dict mutations inside the lock. |
| **RTDv2-004** | `RtdServer.cs:355-371` | `Stop()` disposes `_cts` while producer awaits `Task.Delay(token)`; possible `ObjectDisposedException`. | Cancel only; let GC / finalizer dispose after producer observes cancellation. (Or: await `_producer` then dispose, but that requires async Stop. Surgical fix: skip the dispose.) |
| **RTDv2-005** | `RtdServer.cs:93-94` | `ConnectData` overwrites prior reg for same `TopicId` without unsubscribing the old reg. | Check `TryGetValue`; if a prior reg exists, `Unsubscribe` it before adding the new one. |
| **RTDv2-006** | `RtdServer.cs:232-243` | `f.Stop()` throw skips `_feeds.Clear()`. | Wrap foreach in try/finally; clear in finally. |
| **RTDv2-007** | `ToolkitLifetime.cs:54-67` | `Shutdown` cancels but never disposes the CTS. | Dispose in `Shutdown` (cancel first, then dispose; both wrapped against `ObjectDisposedException`). |
| **RTDv2-008** | `RtdServer.cs:333-335` | If `ToolkitLifetime.ShutdownToken` is already cancelled when `Subscribe` runs, the linked CTS fires synchronously; `Task.Run(action, ct)` with cancelled ct never invokes the action; producer is born complete. | Before `Task.Run`, check `ToolkitLifetime.ShutdownToken.IsCancellationRequested` and return without creating the task. Document that subscribe-after-shutdown is a no-op. |

## Confirmed findings — Medium

| ID | Location | Notes / Fix |
| --- | --- | --- |
| V005 | `Marshaling.cs:140` | `Convert.ToString(object, IFormatProvider)` ignores the provider for non-`IConvertible` types. Patch to call `IConvertible.ToString(provider)` explicitly when the value is convertible; fall back to `value.ToString() ?? ""`. |
| V008 | `Marshaling.cs:343-354` | All `ExcelError` variants hash to `typeof(ExcelError).GetHashCode()`. Fix: include the enum value in the hash. |
| V011 | `BulkTransfer.cs:31-56` | Sheet name with `]`/`[` reaches `xlfEvaluate` unsanitized. Patch: reject or escape these characters in `sheetName`. |
| V013 | `AddIn.cs:55-60` | Non-`Exception` payload to `OnUnhandledException` is dropped silently. Patch: log `exceptionObject?.ToString()` instead. |
| F5 | `ParallelUtilities.cs:466-467` | See above (already in High). |
| F7 (concurrency) | `VectorizedKernels.cs` UDF wrappers | Catch-all `Exception` masks `OutOfMemoryException`. Patch: re-throw `OutOfMemoryException` / `StackOverflowException` / `AccessViolationException`. |
| U-07 | `DeveloperUtilities.cs:125,819` | `"Infinity"` / `"NaN"` text coerces to non-finite double, asymmetric. Fix: reject non-finite in string path of `TryToDouble`. |
| U-08 | `DeveloperUtilities.cs:153-178` | `RemoveDuplicateRows` doc claims `CellEquality`, code uses `ToStringSafe`. Update doc. |
| U-09 | `DeveloperUtilities.cs:865,903-908` | Per-cell `byte[]` allocs in hash paths. Patch: reuse a one-byte buffer outside the loop. |
| U-10 | `DeveloperUtilities.cs:711-737` | `FlattenIntIndexes` silently wraps large doubles to negative ints. Callers bounds-check, so it produces a confusing message rather than an unsafe access. Patch: bounds-check `d` against `int.MinValue/MaxValue` before the cast, throw with the original value in the message. |
| U-11 | `DeveloperUtilities.cs:413-425` | Regex path uses `CultureInvariant`, non-regex uses `Ordinal` — document the asymmetry. |
| F7 (file I/O) | `DirectFileIO.cs:371-386 vs 237` | UTF-8 BOM asymmetry. `ResolveEncoding(null/blank)` returns `Encoding.UTF8` (BOM-emitting); `WriteDelimitedAsync`'s `??=` path uses `new UTF8Encoding(false)` (no BOM). Fix: `ResolveEncoding` returns the no-BOM `UTF8Encoding` for the default case. |
| F8 (file I/O) | `DirectFileIO.cs:82,162-166` | EOF inside a quoted field silently flushes. Patch: throw or surface `#VALUE!` when EOF with `inQuotes==true`. |
| F9 (file I/O) | `DirectFileIO.cs:404-420` | `ResolveDelimiter` accepts `"`, `\r`, `\n`. Fix: reject these. |
| F10 (file I/O) | `DirectFileIO.cs:419` | Delimiter as a high surrogate is broken. Low priority. |
| F12 (file I/O) | `DirectFileIO.cs:65-72,89` | `FileShare.ReadWrite | FileShare.Delete` on read allows torn reads. Fix: `FileShare.Read` (allow concurrent readers only). |
| RTDv2-011 | `RtdServer.cs:118-139` | `FlushTick` not reentrancy-guarded. Patch: `Interlocked.CompareExchange` gate. |
| RTDv2-013 | `RtdServer.cs:118-139` | O(n) per-tick scan over all topics regardless of dirtiness. Patch (optional): maintain a per-feed dirty flag. Tag DEFERRED. |
| RTDv2-015 | `RtdServer.cs:355-371` | `Stop()` doesn't await `_producer`. Patch: minor — store the producer task and let callers await if they need; not blocking for current AutoClose path. |
| RTDv2-017 | `RtdServer.cs:91` | Unbounded `_feeds` growth. Patch: cap `_feeds.Count`; reject new specs beyond cap. |

## Confirmed findings — Low

V009, V010, V012, V014, V015, V016, V017, U-12, U-13, U-14, U-15, U-16, F8 (conc), F9 (conc), F10 (conc), F11 (conc), F13 (file), F14 (file), F15 (file), F16 (file), F17 (file), F18 (file), F19 (file), RTDv2-010, RTDv2-012, RTDv2-014, RTDv2-019, RTDv2-020, RTDv2-021, RTDv2-022, RTDv2-023, RTDv2-024, RTDv2-025 — all real but low-impact; documented in the per-domain reports.

## Rejected findings (with reasoning)

Round 1 false positives that survived into round-2 reports as candidates and that I now re-reject:

- **`DotProductAvx2` OOB** (round-1 P003). Trace: `Length=5, width=4`, condition `i <= a.Length - width` = `i <= 1`. Iterations: i=0 reads a[0..3] OK; i=4 exits. Tail handles i=4. **No OOB.** Round-2 concurrency agent correctly noted this.
- **`Array.Sort` claimed unstable** (round-1 U007). The comparer at `DeveloperUtilities.cs:832` returns `a.CompareTo(b)` as the final tiebreaker — original index. This makes the sort effectively stable. Confirmed correct.
- **`(double)decimal.MaxValue → Infinity`** (round-1 B004). `decimal.MaxValue ≈ 7.9e28`, well within `double.MaxValue ≈ 1.8e308`. No Infinity. Confirmed by re-check of .NET docs.
- **`MatrixMultiply` per-column `bCol` allocation** (round-1 P006). `var bCol = new double[k]` is at line 657, OUTSIDE the j-loop. Allocated once, reused.
- **`CellEquality` NaN inconsistency** (round-1 B008). `TryToDouble` filters NaN (rejects as not-finite); the comparer never sees a NaN double from this codebase.
- **`IsBlankOrError` "contract ambiguity"** (round-1 B006). Documented design choice; not a defect.
- **`ToolkitLifetime.ShutdownToken` race with `Reset`** (round-1 RTD-007). Both lock `Gate` — serialized. No race.
- **`TopicRegistration.LastPushed` torn write** (round-1 RTD-013). Reference assignment is atomic on .NET; only flush thread writes. Safe.
- **`Volatile.Read/Write` on LatestValue** (round-1 RTD-009 / RTDv2 self-rejected). Correct for ref assignment.
- **`SineFeed._epoch` race** (RTDv2 self-rejected). Single ctor writer.
- **`Lazy<FeedManager>` cross-reload leak** (RTDv2 self-rejected). State cleared on shutdown.
- **`ServerStart never throws`** (RTDv2-026 self-rejected).
- **`RTDv2-018` lambda allocation** (self-rejected). Static lambda is cached.
- **V017 (info leak in trace)** — defense-in-depth, not a real risk.
- **F18 (per-row await allocations)** — micro-opt only.
- **F14 (FileOptions.SequentialScan on write)** — benign noise on Windows.

## Fix plan and dependency order

Apply in this order, building between each:

1. **Marshaling.cs**: V001, V002, V005, V006, V008, V009 (one PR-style commit).
2. **BulkTransfer.cs**: V003, V007, V011, V012.
3. **AddIn.cs**: V004, V013.
4. **DeveloperUtilities.cs**: U-01, U-04, U-05, U-06, U-09, U-10, U-11.
5. **ToolkitLifetime.cs**: RTDv2-007.
6. **RtdServer.cs**: RTDv2-001, RTDv2-002, RTDv2-003, RTDv2-004, RTDv2-005, RTDv2-006, RTDv2-008, RTDv2-011, RTDv2-017.
7. **DirectFileIO.cs**: F1, F2, F3, F5, F7, F8, F9, F12.
8. **VectorizedKernels.cs** + **ParallelUtilities.cs**: F1-F5 (concurrency).

Hashing collision fixes (U-02, U-03) and document/Med/Low items applied alongside their domain commits.

Items DEFERRED (architectural-scope):

- F6 (file I/O OOM cap) — needs a streaming UDF API to truly fix; document instead.
- RTDv2-013 (per-feed dirty tracking) — needs a feed-notification channel; defer.
- F6 (concurrency: `ParallelRowReduce` wasted alloc) — surgical fix is feasible but rewrites the kernel.

Will be added to `audit/DEFERRED.v2.md` if not fixed.
