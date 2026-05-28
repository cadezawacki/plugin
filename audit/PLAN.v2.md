# Audit plan — round 2

Round 1 produced 5 sub-agent reports but never reached synthesis or fix
application; the codebase is unchanged from the post-feature commit. Round 2
re-audits with fresh eyes, sharper prompts, lower-level inspection, and a
single-batch parallel dispatch. All outputs are suffixed `.v2`.

## Improvements over round 1

1. **Use `general-purpose` sub-agents (not `Explore`)** — round 1 used Explore
   which is read-only; sub-agents struggled to write reports to disk. v2 sub
   agents have `Write` access and are instructed to write directly.
2. **Tighter scope** — the round-1 boundary agent drifted into RTD territory.
   v2 prompts explicitly forbid out-of-scope findings and tell each agent to
   skip and flag, not investigate.
3. **Lower level** — v2 prompts require sub-agents to *trace the exact CIL
   semantics or .NET BCL contract* for each finding, not summarize. Examples:
   reading the doc for `Timer.Dispose()`, `CancellationTokenSource.Dispose()`,
   `StreamReader.ReadAsync` chunk boundaries.
4. **Independent re-verification by orchestrator before any fix.** Round 1
   already showed at least three false positives (DotProductAvx2 OOB,
   `Array.Sort` stability, MatrixMultiply bCol allocation). v2 will re-walk
   every finding before applying.
5. **Mandatory rejection appendix** in each sub-agent report: every plausible
   finding the agent considered and rejected, with the reasoning. Forces
   self-skepticism.

## Repository map (unchanged since round 1)

- Language: C# 12, .NET 8 LTS, `net8.0-windows`. Single project,
  `src/ExcelPerfToolkit/ExcelPerfToolkit.csproj`, packed as XLL via Excel-DNA.
- 9 source files, ~3,824 LOC.
- Entry points: `IExcelAddIn` lifecycle, `[ExcelFunction]` UDFs, RTD ProgId
  `EPT.Rtd`.
- External I/O: local file streams (CSV/TSV). Excel COM boundary via
  `ExcelReference` + `XlCall`. No DB, no network, no queues.
- Concurrency model: Excel main thread, MTR worker pool (`IsThreadSafe = true`
  UDFs), Excel RTD thread (`Topic.UpdateValue`), per-feed background tasks
  (Task.Run), throttle `Timer` callback on ThreadPool, async file I/O on
  ThreadPool continuations.

## Partition into audit domains (same boundaries as round 1)

### Domain 1 — Boundary & Conversion (~650 LOC)
**Files:** `AddIn.cs`, `Marshaling.cs`, `BulkTransfer.cs`
**Forbidden territory:** RTD server, feed manager, async file I/O, SIMD,
parallel batch. If you find an issue there, flag and stop; do NOT trace into
it.

### Domain 2 — Developer Utilities (~914 LOC)
**Files:** `DeveloperUtilities.cs`
**Forbidden territory:** everything else.

### Domain 3 — Concurrency: Parallel + SIMD (~1,236 LOC)
**Files:** `ParallelUtilities.cs`, `VectorizedKernels.cs`
**Forbidden territory:** RTD server, file I/O, Marshaling internals.

### Domain 4 — RTD Server + Lifetime (~565 LOC)
**Files:** `RtdServer.cs`, `ToolkitLifetime.cs`
**Forbidden territory:** SIMD kernels, file I/O internals.

### Domain 5 — Direct File I/O (~457 LOC)
**Files:** `DirectFileIO.cs`
**Forbidden territory:** RTD server, SIMD, lifetime token plumbing.

## Cross-domain seams (annotate but do not chase)

- `Marshaling.TryToDouble` / `ToStringSafe` — consumed by every other domain;
  Domain 1 owns the contract.
- `ToolkitLifetime.ShutdownToken` — produced by Domain 4, consumed by Domain
  4 and Domain 5.
- `AddIn.AutoOpen / AutoClose` ordering — Domain 1 owns.
- `XlCall.RTD` from `EPT.RTD` UDF — Domain 4 owns.
- Sync-over-async in DirectFileIO UDFs — Domain 5 owns.
- `IsThreadSafe = true` contract — Domain 3 audits SIMD/parallel; Domain 4
  audits RTD UDF.

## Sub-agent outputs

Each agent writes `audit/reports/domain-<n>-<topic>.v2.md` and the report
includes a rejection appendix.

## Round-2 acceptance criteria

- Every finding must include the exact failure trace and a citable file:line.
- Findings the orchestrator cannot re-verify by reading the source are
  dropped.
- Synthesis goes to `audit/SYNTHESIS.v2.md`.
- Fixes are surgical, one finding at a time, build re-checked after each.
- Changelog at `audit/CHANGELOG.v2.md`; deferrals at `audit/DEFERRED.v2.md`.
