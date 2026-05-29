# Audit plan

## Repository map

- Language: C# 12, .NET 8 LTS, `net8.0-windows`.
- Single project: `src/ExcelPerfToolkit/ExcelPerfToolkit.csproj`, packed via Excel-DNA into x86/x64 XLLs.
- Source files (9; ~3,824 LOC total):
  - `AddIn.cs` (74) — `IExcelAddIn` lifecycle.
  - `Marshaling.cs` (356) — Excel `object[,]` <-> CLR converters.
  - `BulkTransfer.cs` (222) — `ExcelReference` bulk read/write.
  - `DeveloperUtilities.cs` (914) — utility UDFs.
  - `ParallelUtilities.cs` (490) — `IsThreadSafe` UDFs + `Parallel.For` batch.
  - `VectorizedKernels.cs` (746) — SIMD kernels with scalar fallback.
  - `RtdServer.cs` (484) — multithreaded RTD server + feed manager.
  - `DirectFileIO.cs` (457) — async streaming CSV/TSV.
  - `ToolkitLifetime.cs` (81) — shared shutdown token + TraceSource factory.

## Entry points

- Excel-DNA auto-discovers `[ExcelFunction]` UDFs and `[ProgId]` RTD servers.
- `AddIn.AutoOpen` / `AutoClose` on add-in load/unload.
- Every `[ExcelFunction]` is an entry point from Excel.
- `EPT.Rtd` ProgId is an entry point from Excel's RTD pipeline.

## Hot paths

- Any UDF that takes/returns `object[,]` — invoked once per recalc.
- The RTD throttle `Timer` callback (`RtdServer.FlushTick`) — every 250 ms across all topics.
- The CSV parser inner loop in `DirectFileIO.ReadDelimitedAsync`.
- The SIMD inner loops in `VectorizedKernels`.

## Shared mutable state

| State | File | Synchronization |
| --- | --- | --- |
| `ToolkitLifetime._cts` | `ToolkitLifetime.cs` | `lock (Gate)` |
| `FeedManager._feeds` | `RtdServer.cs` | `ConcurrentDictionary` + `_gate` on shutdown |
| `Feed._subscribers` | `RtdServer.cs` | `ConcurrentDictionary` |
| `Feed._latestValue` | `RtdServer.cs` | `Volatile.Read/Write` |
| `Feed._producer`, `_cts` | `RtdServer.cs` | `lock (_gate)` |
| `RtdServer._topics` | `RtdServer.cs` | `ConcurrentDictionary` |
| `RtdServer._flushTimer` | `RtdServer.cs` | `Interlocked.Exchange` on terminate |
| `TopicRegistration.LastPushed` | `RtdServer.cs` | read/written from flush thread only (by design) |
| `RandomFeed._rng` | `RtdServer.cs` | `lock (_gate)` |
| Static `TraceSource` instances | every file | TraceSource is thread-safe per BCL |

## External I/O

- `FileStream` reads/writes in `DirectFileIO.cs` (async).
- Excel COM boundary via `ExcelReference.GetValue/SetValue` and `XlCall.Excel`.
- No DB, no network, no message queues.

## Concurrency model

- Excel main thread: `BulkTransfer` API, all non-`IsThreadSafe` UDFs, `AddIn.AutoOpen/AutoClose`.
- Excel MTR worker pool: every `IsThreadSafe = true` UDF (`EPT.MT.*`, `EPT.SIMD.*`, `EPT.RTD`, all `DeveloperUtilities` are NOT thread-safe though some are pure).
- Excel RTD thread: `RtdServer.ConnectData`, `DisconnectData`, `ServerStart`, `ServerTerminate`, and any `Topic.UpdateValue` calls (Excel-DNA marshals to this thread).
- Background `Task.Run` per feed: `Feed.RunAsync` producers.
- `Timer` callback thread (`ThreadPool`): `RtdServer.FlushTick`.
- Async file I/O continuations: `DirectFileIO` async pipelines.

## Partition into audit domains

Domains chosen by subsystem coherence; each fits comfortably in one sub-agent's context.

### Domain 1 — Boundary & Conversion (~650 LOC)
**Files:** `AddIn.cs`, `Marshaling.cs`, `BulkTransfer.cs`
**Surface:** Excel boundary, type coercion semantics, lifecycle.

### Domain 2 — Developer Utilities (~914 LOC)
**Files:** `DeveloperUtilities.cs`
**Surface:** Pure transforms, hashing, lookup, split/join, sort.

### Domain 3 — Concurrency: Parallel + SIMD (~1,236 LOC)
**Files:** `ParallelUtilities.cs`, `VectorizedKernels.cs`
**Surface:** Thread-safe UDFs, `Parallel.For`, SIMD path selection, scalar fallbacks.

### Domain 4 — RTD Server + Lifetime (~565 LOC)
**Files:** `RtdServer.cs`, `ToolkitLifetime.cs`
**Surface:** Multithreaded RTD server, FeedManager, throttle timer, background feeds, shared shutdown CTS.

### Domain 5 — Direct File I/O (~457 LOC)
**Files:** `DirectFileIO.cs`
**Surface:** Async streaming CSV/TSV parser + writer, sync-over-async bridge from UDFs.

## Cross-domain seams (call out so they aren't missed)

- **Marshaling coercion semantics** — `Marshaling.TryToDouble` / `ToStringSafe` are consumed by every other file. Any change in semantics ripples everywhere. Domain 1 owns the contract; Domains 2, 3, 5 are consumers and must flag misuse.
- **ToolkitLifetime.ShutdownToken** — produced by Domain 4, consumed by Domain 4 (RTD feeds) and Domain 5 (file I/O). Cancellation semantics across the seam.
- **AutoClose ordering** — `AddIn.AutoClose` calls `ToolkitLifetime.Shutdown()` then `FeedManager.Instance.Shutdown()`. Domain 1 + Domain 4 share this invariant.
- **`XlCall.RTD` from `EPT.RTD` UDF** — UDF is registered `IsThreadSafe = true` (Domain 4), invoked by Excel MTR pool. Excel-DNA documents `XlCall.RTD` as safe; trust boundary is here.
- **Sync-over-async** — `DirectFileIO` UDFs call `.ConfigureAwait(false).GetAwaiter().GetResult()`. Deadlock potential and threadpool exhaustion (Domain 5 audits primarily; Domain 1 flags if the UDF model is misused).
- **`IsThreadSafe` UDFs may not touch the Excel object model** — Domain 3 audits this for `EPT.MT.*` and `EPT.SIMD.*`. Domain 4 audits this for `EPT.RTD`.
- **Static state across UDFs** — all UDFs that are `IsThreadSafe` must avoid shared mutable static state. Static `TraceSource` instances are read-only; static `HardwareAccelerated`/`Avx2Supported`/`FmaSupported` are immutable. Domain 3 + Domain 4 audit.

## Sub-agent dispatch

Phase 2 launches five `Explore` sub-agents in parallel, each restricted to its assigned files plus the cited cross-domain seams. Each writes `audit/reports/<domain>.md`.
