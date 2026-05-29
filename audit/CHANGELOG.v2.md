# Round-2 changelog

Mapping of confirmed findings → files changed → one-line fix description.
All fixes are surgical; release builds for both x86 and x64 are clean with
zero warnings after every commit.

## Critical

| Finding | Files | Fix |
| --- | --- | --- |
| **V001** Marshaling.TryToDouble culture-confused decimal | `Marshaling.cs` | Dropped `NumberStyles.AllowThousands` from the invariant `double.TryParse` pass; kept it for the current-culture fallback. `"1,5"` no longer parses as `15`. |
| **U-01** ReDoS in EPT.FINDREPLACE | `DeveloperUtilities.cs` | Added `TimeSpan.FromSeconds(1)` `MatchTimeout` to the `Regex` constructor; per-cell `RegexMatchTimeoutException` caught and surfaced as `#VALUE!` with a logged reason. |
| **RTDv2-001** `Timer.Dispose()` without WaitHandle | `RtdServer.cs` | `ServerTerminate` now uses `timer.Dispose(WaitHandle)` and awaits the wait handle before clearing `_topics`. |
| **RTDv2-002** FlushTick outer try/catch eats whole tick | `RtdServer.cs` | Per-topic try/catch inside the foreach so one bad topic can't starve the rest. |
| **F1 (file I/O)** lastWasCR not reset on opening quote | `DirectFileIO.cs` | Opening-quote and delimiter branches now reset `lastWasCR = false`; `a\r"b"\nc\n` now parses correctly. |

## High

| Finding | Files | Fix |
| --- | --- | --- |
| **V002** `DateTime.ToOADate` throws for years 1–99 | `Marshaling.cs` | Wrapped the OAdate call in try/catch (OverflowException). |
| **V003** `XlCallException` escapes `ResolveRange` | `BulkTransfer.cs` | Caught `XlCallException` and translated to `ArgumentException` with inner. |
| **V004** `SafeGetExcelVersion` mis-renders ExcelError | `AddIn.cs` | Type-check for `ExcelError` before stringifying the return value. |
| **V006** Unbounded string parse in TryToDouble | `Marshaling.cs` | Length cap of 64 chars before parsing. |
| **V007** Multi-cell anchor silently overwrites | `BulkTransfer.cs` | Anchor size must equal block shape (or be single-cell); otherwise `ArgumentException`. |
| **V011** Sheet-name injection via `[`/`]` | `BulkTransfer.cs` | Reject those characters (and `\r`/`\n`) in `sheetName` up front. |
| **V013** OnUnhandledException drops non-Exception payload | `AddIn.cs` | Log the payload verbatim when it isn't an Exception. |
| **F1, F2 (concurrency)** rows*cols int overflow in `FlattenToDoubles` / `BoxFlatDoubles` | `VectorizedKernels.cs` | long-cast guard before allocation and in length-check. |
| **F3 (concurrency)** `MatrixMultiply` shape products overflow | `VectorizedKernels.cs` | long-cast every shape product in the entry guards. |
| **F4 (concurrency)** `L2Normalize` divides by infinity | `VectorizedKernels.cs` | Non-finite norm short-circuits to zeros + 0 return. |
| **F5 (concurrency)** `ParallelUtilities.Dot` int overflow | `ParallelUtilities.cs` | long-cast `na`/`nb`; reject ranges > Int32.MaxValue cells. |
| **U-04, U-05** StackColumns/Unpivot int overflow | `DeveloperUtilities.cs` | long-cast `rows*cols` / `dataRows*valueColumns` with explicit reject. |
| **U-06** SortBlock silently drops scalar descendingFlags | `DeveloperUtilities.cs` | Accept `bool`/`double`/`int`/`long`/`string` scalar; broadcast to all keys via new `CoerceBool` helper. |
| **F2 (file I/O)** mid-field `"` enters quote mode | `DirectFileIO.cs` | Quote inside an unquoted field that already has content is now appended as a literal char instead of flipping state. |
| **F3 (file I/O)** delimiter branch doesn't reset lastWasCR | `DirectFileIO.cs` | `lastWasCR = false` at the top of the delimiter branch. |
| **F5 (file I/O)** ResolveEncoding swallows only ArgumentException | `DirectFileIO.cs` | Also catches `NotSupportedException` (missing CodePagesEncodingProvider). |
| **F7 (file I/O)** UTF-8 BOM asymmetry | `DirectFileIO.cs` | Single `DefaultUtf8 = new UTF8Encoding(false)` used by both default read and default write. |
| **F8 (file I/O)** EOF inside quoted field silently flushed | `DirectFileIO.cs` | Throws `InvalidDataException` (surfaced as `#VALUE!` via the UDF catch) when EOF arrives while `inQuotes==true`. |
| **F9 (file I/O)** `ResolveDelimiter` accepts `"`/CR/LF | `DirectFileIO.cs` | Reject them with a clear `ArgumentException`. |
| **F12 (file I/O)** `FileShare.ReadWrite | Delete` allows torn reads | `DirectFileIO.cs` | Switched to `FileShare.Read`. |
| **RTDv2-003** Subscribe/Unsubscribe race on `_subscribers` | `RtdServer.cs` | Dict mutations + `IsEmpty` check moved inside `_gate`. Introduced `StopLocked()` to call from inside the lock. |
| **RTDv2-004** Stop disposes CTS during Task.Delay | `RtdServer.cs` | Stop now only cancels and nulls; lets GC dispose after the producer observes cancellation. |
| **RTDv2-005** ConnectData overwrites prior reg | `RtdServer.cs` | If the topic id is already in `_topics`, unsubscribe the old reg first. |
| **RTDv2-006** `FeedManager.Shutdown` exception loses Clear | `RtdServer.cs` | foreach in try; per-feed Stop catches Exception; `_feeds.Clear()` in `finally`. |
| **RTDv2-007** ToolkitLifetime.Shutdown leaks CTS | `ToolkitLifetime.cs` | Dispose `_cts` after cancel; cache the cancelled token in a separate field so the getter doesn't touch a disposed source. |
| **RTDv2-008** Linked CTS born cancelled if shutdown already fired | `RtdServer.cs` | Subscribe early-returns if `ToolkitLifetime.ShutdownToken.IsCancellationRequested`. |

## Medium / Low — applied as part of the above commits

| Finding | Files | Fix |
| --- | --- | --- |
| **V005** Convert.ToString ignores provider for non-IConvertible | `Marshaling.cs` | Switch to `IConvertible.ToString(provider)` explicitly; fall back to `value.ToString()`. |
| **V008** All `ExcelError` values hash to one bucket | `Marshaling.cs` | `GetHashCode` includes the enum value via `HashCode.Combine`. |
| **F7 (concurrency)** UDF catch swallows OOM and friends | `VectorizedKernels.cs` | Added `IsCritical` helper; `catch (Exception ex) when (!IsCritical(ex))` rethrows OOM/SOH/AccessViolation/ThreadAbort. |
| **U-09** Per-cell `Span.ToArray()` allocs in Sha256Block | `DeveloperUtilities.cs` | Separator byte arrays allocated once outside the hot loop. |
| **U-10** FlattenIntIndexes silent double→int wrap | `DeveloperUtilities.cs` | Range + integer-truncation check before cast; clearer error message. |
| **RTDv2-011** FlushTick reentrancy | `RtdServer.cs` | `Interlocked.CompareExchange` gate at the top of FlushTick. |
| **RTDv2-017** Unbounded `_feeds` growth | `RtdServer.cs` | Cap at `MaxDistinctFeeds = 1024` distinct specs; reject overflow with `InvalidOperationException`. |

## Build verification

After every fix, a release build of both x86 and x64 was re-run. Final state:

```
dotnet build src/ExcelPerfToolkit/ExcelPerfToolkit.csproj -c Release -p:Platform=x64 -> 0 Warning(s), 0 Error(s)
dotnet build src/ExcelPerfToolkit/ExcelPerfToolkit.csproj -c Release -p:Platform=x86 -> 0 Warning(s), 0 Error(s)
```

## Concurrency invariants documented in source

The following invariants are now stated as comments next to the code that
enforces them; future changes that violate them should be rejected at review:

- `RtdServer.ServerTerminate` returns only after the last `FlushTick`
  callback has completed (Timer.Dispose(WaitHandle)).
- `Feed._subscribers` mutations and the start/stop decision are made under
  `_gate` as a single atomic step.
- `Feed.Stop`/`StopLocked` does not dispose `_cts` while the producer Task
  may be inside `await Task.Delay(token)`.
- `ToolkitLifetime.ShutdownToken` getter never touches a disposed CTS — the
  cancelled token is captured into `_token` before disposal.
- `Subscribe` after `ToolkitLifetime.Shutdown` is a no-op (avoids the
  born-cancelled linked CTS pattern).
- `BulkTransfer.WriteBlock(sheet, anchor, block)` requires the anchor to be
  single-cell or shape-matching with the block.

## Rejected at synthesis (re-traced and discarded)

- DotProductAvx2 out-of-bounds (Round 1 P003).
- Array.Sort claimed unstable (CompareRows tiebreaker preserves order).
- `(double)decimal.MaxValue → Infinity` (decimal range fits within double).
- CellEquality NaN inconsistency (TryToDouble filters NaN).
- IsBlankOrError "contract ambiguity" (intentional design).
- ToolkitLifetime.ShutdownToken race with Reset (lock serializes them).
- TopicRegistration.LastPushed torn write (single-writer + ref assignment).
- Volatile.Read/Write on LatestValue (correct for ref assignment).
- SineFeed._epoch race (single ctor writer).
- Lazy<FeedManager> cross-reload leak (state cleared on shutdown).

See `audit/SYNTHESIS.v2.md` for the full rejected-findings appendix with
reasoning.
