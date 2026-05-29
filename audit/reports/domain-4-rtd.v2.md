# Domain 4 — RTD Server + Lifetime (Round 2, fresh eyes)

**Scope (files in audit):**
- `/home/user/plugin/src/ExcelPerfToolkit/RtdServer.cs`
- `/home/user/plugin/src/ExcelPerfToolkit/ToolkitLifetime.cs`

**Forbidden territory:** AddIn.cs, vector kernels, file I/O, utilities. Any cross-domain observation is noted only in the rejected appendix.

Round 1 confirmed findings have been re-traced under load, restart, and concurrency assumptions, and several were under-rated. Round 2 added 9 new findings driven by lower-level interactions (CTS finalizer, linked-token race, disposed-CTS reads, captured-token registrations, lazy singleton across reloads, COM marshaling queue, etc.).

## Findings Table

| ID | Severity | Location | Category | Failure Scenario (one-line) | Confidence |
|---|---|---|---|---|---|
| RTDv2-001 | Critical | RtdServer.cs:73-79 | Lifecycle / in-flight callback | `Timer.Dispose()` w/o WaitHandle; FlushTick still running while `_topics.Clear()` + `FeedManager.Shutdown()` execute → `UpdateValue` on disposed Topic. | 0.97 |
| RTDv2-002 | Critical | RtdServer.cs:122-138 | Exception scope | One topic's `UpdateValue` throws → outer try/catch eats it, the remaining N-1 topics get no update this tick (and every following tick if the throw is deterministic). | 0.92 |
| RTDv2-003 | High | RtdServer.cs:326-350 | Race / lost subscription | `_subscribers` dict mutation outside `_gate` lets Unsubscribe see `IsEmpty` and call Stop while Subscribe is mid-flight; new subscriber sees no producer until Excel triggers re-subscribe. | 0.88 |
| RTDv2-004 | High | RtdServer.cs:355-371 | TOCTOU on CTS dispose | `Stop()` disposes CTS while producer awaits `Task.Delay(token)`; the Delay's internal registration fires `Cancel()` on the now-being-disposed source → possible `ObjectDisposedException` thrown out of `await Task.Delay`. | 0.78 |
| RTDv2-005 | High | RtdServer.cs:93-94 | Logic / orphan registration | Excel re-issues `ConnectData` for same `TopicId`; new reg overwrites in `_topics`, but the *old* reg is still in `feed._subscribers` → orphan that prevents feed from ever stopping. | 0.85 |
| RTDv2-006 | High | RtdServer.cs:232-243 | Exception safety | `f.Stop()` throw aborts the foreach; `_feeds.Clear()` never runs; stale dead feeds remain in registry across re-opens. | 0.82 |
| RTDv2-007 | High | ToolkitLifetime.cs:54-67 | Resource leak / token reuse | `Shutdown()` cancels but never disposes `_cts`; subsequent `ShutdownToken` getter still returns the cancelled token (functionally OK) but native `SafeWaitHandle` is held until next `Reset()` or finalizer. Worse: if process never calls Reset (Excel exit path), finalizer thread eventually disposes — racing any consumer that captured the token. | 0.80 |
| RTDv2-008 | High | RtdServer.cs:333-335 | Synchronous-fire of linked CTS | If `ToolkitLifetime.ShutdownToken` is *already cancelled* when `Subscribe` runs (e.g. Subscribe arrives during/after AutoClose), `CreateLinkedTokenSource` produces an already-cancelled source. The new `Task.Run(..., ct)` is born Canceled; `_producer.IsCompleted` is true on next Subscribe → infinite "restart that never runs" loop, no producer ever runs, `LatestValue` is stuck at `null`/initial. | 0.85 |
| RTDv2-009 | Medium | RtdServer.cs:73-74 | Timer callback after null-exchange | `Interlocked.Exchange` to null does NOT cancel an in-flight FlushTick. `_topics` is also mutated and `FeedManager.Instance.Shutdown()` cancels Feed CTSes — the running FlushTick can dereference `reg.Topic` after Excel-DNA has already torn down the Topic on `ServerTerminate`. AccessViolation possible at the COM boundary. | 0.78 |
| RTDv2-010 | Medium | RtdServer.cs:124-133 | Iteration over mutating dict | `foreach (var kv in _topics)` runs concurrently with `DisconnectData` (line 112: `_topics.TryRemove`). `ConcurrentDictionary` enumeration is snapshot-ish but the *value* (`TopicRegistration`) and its `reg.Topic` can be torn down by Excel between yield and use. `reg.Topic.UpdateValue` may run against a topic already disconnected; Excel-DNA's marshaler discards but logs noisily. | 0.72 |
| RTDv2-011 | Medium | RtdServer.cs:118-139 | Reentrancy / overlapping FlushTick | The throttle Timer fires every 250ms but the callback is not re-entrancy-guarded. If a tick takes > 250ms (large topic set + slow `Equals(current, reg.LastPushed)` for boxed doubles → boxing alloc per compare), the next tick fires on another ThreadPool thread and both iterate the same `_topics`. Two threads call `reg.Topic.UpdateValue` for the same topic; `reg.LastPushed` writes race (single field, last-writer-wins) → spurious duplicate pushes and `LastPushed` torn semantics. | 0.74 |
| RTDv2-012 | Medium | RtdServer.cs:118-139 | Unbounded marshaling queue / no backpressure | Each `UpdateValue` enqueues into Excel-DNA's per-topic dispatch. If Excel is blocked (modal dialog, heavy recalc, debugger pause), the FlushTick keeps producing change-deltas and the queue grows unbounded. No drop policy, no cap; memory bloat under partial Excel stalls. | 0.65 |
| RTDv2-013 | Medium | RtdServer.cs:118-139 | O(n) iteration on each tick over all topics regardless of dirtiness | Every 250ms the timer walks *all* topics and probes `LatestValue` even for topics that haven't changed in hours (sine: changing constantly; clock: every second; counter:60000 maybe every minute). At 50k topics this is 200k allocs/sec for boxed `Equals` and 50k volatile reads — measurable CPU on idle. | 0.55 |
| RTDv2-014 | Medium | ToolkitLifetime.cs:25-34 | Captured-token then disposed source | A consumer calls `Task.Delay(ToolkitLifetime.ShutdownToken)` which registers a callback on the underlying CTS. If `Reset()` then disposes that CTS, the Delay's registration is left attached to a disposed source. On timeout, Task.Delay's CTR cleanup calls `Dispose()` on its registration — disposing a registration of an already-disposed CTS is documented safe, but if `Cancel()` fires between dispose and the registered callback completing, behavior is implementation-defined. | 0.55 |
| RTDv2-015 | Medium | RtdServer.cs:355-371 | Stop() awaits nothing; no join | `Stop()` cancels and nulls but does not await `_producer`. The producer Task may still be on the ThreadPool, running iteration N, when `Shutdown()` returns and `FeedManager._feeds.Clear()` drops the only managed reference. The Task GC-roots itself via Task scheduler. A subsequent `Subscribe()` on a recreated `Feed` instance starts a *second* producer; the orphan from the previous Feed keeps writing `LatestValue` on its own (different) object until it observes cancellation. Memory pressure unchanged but log noise + delayed shutdown. | 0.70 |
| RTDv2-016 | Medium | ToolkitLifetime.cs:73-80 | Per-call TraceSource allocation | `CreateTraceSource` allocates a new `TraceSource` on every call. Callers (e.g. `FeedManager`, `RtdServer`, `Feed`) capture it in `static readonly` so it's once-per-type. BUT `Feed` base type has one static `TraceSource`; `RandomFeed`, `ClockFeed`, etc. inherit it — fine. Acceptable. **Not a finding** but watch if pattern proliferates. | 0.20 (rejected) |
| RTDv2-017 | Medium | RtdServer.cs:91 | Unbounded `_feeds` growth from malformed specs | `GetOrCreate` calls `CreateFeed` for *every distinct* spec string. Whitespace-trimmed but `clock`, `clock ` (post-trim same), but `CLOCK`/`Clock` — keys are OrdinalIgnoreCase so dedup works. However `counter:500`, `counter:501`, `counter:502` are all distinct feeds with distinct timers — an attacker (or buggy formula generator) could produce thousands of distinct counter rates → thousands of feed Tasks. No cap. | 0.60 |
| RTDv2-018 | Low | RtdServer.cs:225 | Lambda allocates on every GetOrCreate miss | `static key => CreateFeed(key)` — `static` lambda, cached. Single closure-free delegate. Fine. **Not a finding.** | 0.10 (rejected) |
| RTDv2-019 | Low | RtdServer.cs:330-336 | `_producer is null \|\| _producer.IsCompleted` — race window with self-restart | If a producer naturally completes (cannot happen — loops on cancellation, but if `RunAsync` throws then catch in `RunSafelyAsync` swallows and Task completes), the next Subscribe sees `IsCompleted` and restarts. But a buggy custom feed could trip an infinite restart hot loop if the producer always throws on first iteration. No backoff. | 0.50 |
| RTDv2-020 | Low | RtdServer.cs:127-131 | Boxing in `Equals(object?, object?)` | For value-type feeds (`double`), `current` and `reg.LastPushed` are boxed. Each tick: `Volatile.Read` returns the already-boxed ref. `ReferenceEquals` first — boxes are distinct refs, so always false. Then `Equals(object?, object?)` unboxes and compares. Two unboxings per topic per tick. Not catastrophic but wasted CPU at 50k topics. | 0.45 |
| RTDv2-021 | Low | RtdServer.cs:271-285 | `ParseSineArgs` swallows malformed input | `sine:abc:def` silently returns `(1, 1)` — user gets a 1Hz amplitude-1 sine and never knows the spec was malformed. Logic bug for usability, not safety. | 0.40 |
| RTDv2-022 | Low | RtdServer.cs:266-269 | `ParseInt` treats 0 and negatives as defaultValue silently | `counter:0` → defaults to 1000ms; user expected an error or "as fast as possible". | 0.35 |
| RTDv2-023 | Low | ClockFeed/SineFeed/RandomFeed | No initial value seeded for ClockFeed | `ClockFeed` does NOT call `LatestValue = ...` in ctor; first read returns null. ConnectData returns `feed.LatestValue ?? string.Empty` → empty cell flash on first subscribe until first tick (up to ~250ms). Acceptable but inconsistent with siblings. | 0.30 |
| RTDv2-024 | Low | RtdServer.cs:434-455 | `SineFeed` wall-clock skew (DateTime.UtcNow) | NTP step backward → negative `t`, Math.Sin handles it but the curve discontinuity is visible. Use `Stopwatch.Elapsed`. | 0.40 |
| RTDv2-025 | Low | RtdServer.cs:458-484 | `RandomFeed._gate` unnecessary | Producer is single-threaded (only one Task at a time per Feed). The lock around `_rng.NextDouble()` is dead code; remove or justify with a comment. | 0.55 |
| RTDv2-026 | Critical | RtdServer.cs:62-67 | ServerStart never throws — no init failure path | If `new Timer(...)` were to fail (resource exhaustion under load), `ServerStart` would let the exception escape; Excel-DNA reports a fail-to-start, but `_topics` may have been written by a concurrent `ConnectData`. (Excel-DNA contracts ServerStart before ConnectData, so probably fine — re-traced as safe.) **Move to rejected.** | 0.10 (rejected) |

## Per-finding sections

### RTDv2-001 — Critical — Timer.Dispose() without WaitHandle (re-confirmed, severity escalated)

- **Location:** `RtdServer.cs:73-79`.
- **Category:** Lifecycle / in-flight callback race.
- **Scenario:** Excel calls `ServerTerminate`. Line 73: `var timer = Interlocked.Exchange(ref _flushTimer, null); timer?.Dispose();`. The `System.Threading.Timer.Dispose()` overload that takes no `WaitHandle` is documented to return immediately even if a callback is currently executing. Meanwhile, on a ThreadPool thread, `FlushTick` is mid-`foreach`, calling `reg.Topic.UpdateValue(current)`. Line 79 calls `_topics.Clear()` and line 80 calls `FeedManager.Instance.Shutdown()` — Feed CTSes are cancelled and disposed. The in-flight FlushTick continues to walk `_topics` (cleared mid-iteration: `ConcurrentDictionary.Clear` is safe to enumerate against but yields zero entries from the point of clear) and may dereference `reg.Feed.LatestValue` after the feed has been Stopped (LatestValue is still readable — Volatile.Read works on a field that survives Stop). Worse: `reg.Topic` is an `ExcelRtdServer.Topic` provided by Excel-DNA; once `ServerTerminate` completes, Excel-DNA may have released the Topic's underlying COM reference. Calling `UpdateValue` on a torn-down Topic is undefined: best case it's a no-op + log; worst case AccessViolation at the COM boundary.
- **Evidence trace:** Lines 70-81.
- **Fix:** Use `using var waitHandle = new ManualResetEvent(false); timer.Dispose(waitHandle); waitHandle.WaitOne();` (or `timer.DisposeAsync().AsTask().Wait(2000)`). Synchronously await pending callbacks before mutating `_topics` or shutting down feeds.
- **Confidence:** 0.97.

### RTDv2-002 — Critical — Outer try/catch in FlushTick halts every remaining topic on one throw

- **Location:** `RtdServer.cs:122-138`.
- **Category:** Exception scope.
- **Scenario:** `foreach` enumerates 50000 topics. Topic #7's `UpdateValue` throws (Excel busy state, COMException, marshaling failure for a non-marshalable value). The outer `try/catch` at line 122-138 catches and logs once; the remaining 49,993 topics get no value pushed this tick. If the throwing condition is deterministic (e.g. a specific spec produces a value Excel rejects), the same topic throws every 250ms forever, starving everyone else.
- **Evidence trace:** Lines 122-138.
- **Fix:** Move the `try/catch` inside the foreach body around the `UpdateValue` + `LastPushed` write only. Per-topic isolation.
- **Confidence:** 0.92.

### RTDv2-003 — High — `_subscribers` mutation outside `_gate`

- **Location:** `RtdServer.cs:326-350`.
- **Category:** Race.
- **Scenario:** Thread A (Subscribe): writes `_subscribers[id_A] = regA` (line 328, OUTSIDE the lock). Thread B (Unsubscribe): TryRemoves its own id_B, then checks `_subscribers.IsEmpty`. If A's `_subscribers[id_A] = regA` write has not yet committed to memory observable by B (or if B's TryRemove and IsEmpty happen between A's call entry and A's dict write), B sees `IsEmpty == true`, calls `Stop()`. A then enters the lock, sees `_producer is null`, creates a fresh linked CTS and Task. Worst case: A has subscribed but the producer Task hasn't started yet, B's cancellation races into A's CTS and the new Task is born cancelled (see RTDv2-008). The subscriber A is registered with no live producer.
- **Evidence trace:** Lines 326-328, 343-350.
- **Fix:** Move the `_subscribers[reg.Topic.TopicId] = reg;` (line 328) and the `_subscribers.TryRemove + IsEmpty` check (lines 345-346) inside the `lock (_gate)` block. Or use `Subscribe`/`Unsubscribe` lock-acquired before any dict mutation.
- **Confidence:** 0.88.

### RTDv2-004 — High — `Stop()` disposes CTS while producer is in `Task.Delay(token)`

- **Location:** `RtdServer.cs:355-371`.
- **Category:** TOCTOU on CTS dispose.
- **Scenario:** Producer is at `await Task.Delay(_intervalMs, cancellationToken)` (line 429, 453, 481, etc.). `Task.Delay(int, CancellationToken)` internally calls `token.Register(...)` to wire cancellation. `Stop()` acquires `_gate`, calls `_cts.Cancel()` (line 361), then `_cts.Dispose()` (line 367). Per .NET docs, `Cancel()` runs registered callbacks *synchronously on the canceling thread*. The Delay's registration completes the delay's TaskCompletionSource and detaches. Then `Dispose()` is called. If the callback was still executing when Dispose was called (it ran synchronously, so no — but only if there's no async tail), Dispose blocks until callbacks finish (documented). So the *common* path is safe. HOWEVER: the producer's `await` resumes on the ThreadPool; if the resume races with `Dispose` and the producer immediately calls `cancellationToken.ThrowIfCancellationRequested()` or examines the token, reading `IsCancellationRequested` on a disposed CTS's token is documented safe (CT is a value type with a back-pointer that is null-checked). So far OK. But `cancellationToken.Register` post-dispose throws `ObjectDisposedException`. If a custom Feed implementation registers post-Delay, kaboom.
- **Evidence trace:** Lines 357-370 + `Task.Delay(_intervalMs, ct)` at lines 407, 429, 453, 481.
- **Fix:** Cancel only inside the lock; capture the `Task _producer` reference, exit lock, `await _producer.ConfigureAwait(false)` (with a timeout), THEN dispose the CTS. Or skip Dispose and let GC handle it (CTS finalizer is well-behaved).
- **Confidence:** 0.78.

### RTDv2-005 — High — `ConnectData` re-subscribe to same TopicId orphans prior registration

- **Location:** `RtdServer.cs:93-94`.
- **Category:** Logic bug / orphan.
- **Scenario:** Excel re-issues `ConnectData(topic_X, spec_A, ...)` followed by `ConnectData(topic_X, spec_B, ...)` without an intervening DisconnectData (Excel-DNA does this in some edge cases — workbook re-open, cell formula change in place). Line 93: `_topics[topic.TopicId] = reg` *overwrites* the old reg. Then line 94: `feed.Subscribe(reg)` adds the new reg to feed_B's `_subscribers`. The OLD reg (for feed_A) is still in feed_A's `_subscribers[topic_X.TopicId]`. The new reg overwrites that ONLY IF feed_A == feed_B (same spec); but if specs differ, the old reg is orphaned in feed_A's `_subscribers`. Now: feed_A's subscriber count never drops to zero (because nothing will ever Unsubscribe the orphan), so feed_A's producer runs forever. Excel never reads from feed_A again (the topic now points to feed_B). Memory + CPU leak per orphan; in pathological churn workloads, accumulates one orphaned Feed producer per spec-change.
- **Evidence trace:** Lines 90-96 (`_topics[topic.TopicId] = reg` is unconditional).
- **Fix:** Before line 93, `if (_topics.TryGetValue(topic.TopicId, out var prior)) { prior.Feed.Unsubscribe(prior); }`. Then proceed with the new registration.
- **Confidence:** 0.85.

### RTDv2-006 — High — `FeedManager.Shutdown` exception in `f.Stop()` skips Clear

- **Location:** `RtdServer.cs:232-243`.
- **Category:** Exception safety.
- **Scenario:** Stop() inside the lock at line 357 catches `ObjectDisposedException` from `_cts.Cancel()`. But `_cts?.Dispose()` at line 367 is NOT inside the catch, and Dispose can throw if a registered callback re-throws (rare but documented). If any `Stop()` throws, the foreach unwinds, line 240 `_feeds.Clear()` is skipped, and dead feeds remain in `_feeds`. Next `GetOrCreate(same_spec)` returns the dead Feed (its CTS disposed, no producer). Subsequent `Subscribe` enters lock, sees `_producer is null` (Stop set it null before throwing), creates new CTS and starts new producer. So actually self-healing if Stop completes the cancel + null assignment before throwing. But if Stop throws BEFORE setting `_producer = null` (line 369), the dead Feed has `_producer.IsCompleted == false` and no Subscribe will ever restart it.
- **Evidence trace:** Lines 232-243 + lines 355-371.
- **Fix:** Wrap the foreach in try/finally: `finally { _feeds.Clear(); }`. Separately: in `Feed.Stop()`, ensure `_cts = null; _producer = null;` happens unconditionally (try/finally inside Stop too).
- **Confidence:** 0.82.

### RTDv2-007 — High — `ToolkitLifetime.Shutdown()` never disposes `_cts`; finalizer races consumers

- **Location:** `ToolkitLifetime.cs:54-67`.
- **Category:** Resource leak + finalizer race.
- **Scenario:** `Shutdown()` cancels `_cts` but does not dispose it. `_cts` holds a kernel `SafeWaitHandle` (the cancellation event). The handle is released only when (a) the next `Reset()` runs (line 44 disposes the old) or (b) GC runs and the CTS finalizer disposes. On Excel exit path, `Reset()` is never called; the finalizer runs on a non-deterministic thread *after* the AppDomain has potentially been shutting down. If any captured `CancellationToken` is still being used by a consumer (e.g. an orphaned Feed producer Task that never observed cancellation because of the gap in RTDv2-015), the finalizer disposes the underlying CTS while the producer's `Task.Delay(token)` registration is still attached. Documented behavior: Dispose with live registrations races; in practice the runtime tolerates it but may log `ObjectDisposedException` from the deferred Task.Delay continuation.
- **Evidence trace:** Lines 54-67 (no Dispose), versus Reset at line 44 which DOES dispose.
- **Fix:** Add `_cts.Dispose();` after the Cancel inside the lock, with an inner try to swallow ObjectDisposedException on a double-shutdown. OR move disposal contract: Shutdown disposes, Reset constructs new (current Reset assumes the old is not yet disposed → that becomes wrong; adjust Reset accordingly).
- **Confidence:** 0.80.

### RTDv2-008 — High — Linked CTS born cancelled when Subscribe runs after ShutdownToken is fired

- **Location:** `RtdServer.cs:333-335`.
- **Category:** Synchronous-cancel of linked token.
- **Scenario:** Suppose `ToolkitLifetime.ShutdownToken` was cancelled (AutoClose ran). For some reason the RtdServer instance is still alive and a stray `ConnectData` arrives (Excel-DNA edge case during teardown, or unit test wiring). `Subscribe` runs. `CreateLinkedTokenSource(ShutdownToken)` *synchronously fires cancellation on the new source during construction* because the parent is already cancelled. `_cts = (cancelled source)`. `var ct = _cts.Token;` — ct is already cancelled. `Task.Run(() => RunSafelyAsync(ct), ct)` — the second `ct` argument tells the scheduler "if ct is cancelled before the action runs, mark the Task Canceled and never invoke the action". The Task is born in Canceled state; `RunSafelyAsync` never executes. `_producer` is non-null but `IsCompleted == true`. On the next `Subscribe`, line 331 sees `_producer.IsCompleted`, repeats the same cycle, creates another stillborn Task. Subscribers accumulate in `_subscribers`, no producer ever runs, `LatestValue` is stuck at the initial value (null for ClockFeed, 0 for others). Excel cells show stale 0 forever.
- **Evidence trace:** Line 333: `CreateLinkedTokenSource(ToolkitLifetime.ShutdownToken)`. The MSRC reference here is: `CancellationTokenSource.CreateLinkedTokenSource` — if any input token is cancelled, the result is cancelled immediately. Line 335: `Task.Run(action, ct)` with already-cancelled ct → canceled task without invoking the delegate.
- **Fix:** Before creating the linked CTS, check `if (ToolkitLifetime.ShutdownToken.IsCancellationRequested) return;` — reject subscribes during shutdown. Or: check the token before Task.Run and skip the Task scheduling. Better: design so Subscribe is impossible during/after Shutdown (state machine; ServerTerminate refuses ConnectData first).
- **Confidence:** 0.85.

### RTDv2-009 — Medium — In-flight FlushTick uses `reg.Topic` after Excel-DNA may have torn it down

- **Location:** `RtdServer.cs:73-74`, escalation of RTDv2-001.
- **Category:** Use-after-release of Excel-DNA Topic.
- **Scenario:** Beyond just "callback ran", the specific failure mode is `reg.Topic.UpdateValue(current ?? string.Empty)` (line 130) after Excel-DNA has invoked `Topic.Dispose()` (Excel-DNA calls this in its internal teardown of `ExcelRtdServer`). Calling `UpdateValue` on a disposed Topic — depending on Excel-DNA version — either throws `ObjectDisposedException` (caught by the outer try at line 135, which now poisons the WHOLE remaining loop per RTDv2-002), or no-ops with a trace warning, or in the worst case (older Excel-DNA) AccessViolation if the underlying COM object was released.
- **Evidence trace:** Line 73-80 plus line 130.
- **Fix:** Same as RTDv2-001 — wait for callback before clearing. Additionally, gate FlushTick with a volatile `_terminated` flag set at the top of ServerTerminate; FlushTick exits early when set.
- **Confidence:** 0.78.

### RTDv2-010 — Medium — Snapshot enumeration vs. concurrent DisconnectData

- **Location:** `RtdServer.cs:124-133`.
- **Category:** Iteration over mutating collection.
- **Scenario:** `foreach (var kv in _topics)` — ConcurrentDictionary's enumerator is moment-in-time consistent (safe to iterate during mutation; does not throw). HOWEVER, the *value* `TopicRegistration` reference yielded by the enumerator may correspond to a topic that has just been DisconnectData'd. `_topics.TryRemove` (line 112) doesn't affect the already-yielded ref. So FlushTick will call `reg.Topic.UpdateValue` on a Topic that DisconnectData just told Excel-DNA "we're done with". Excel-DNA may treat the late UpdateValue as a no-op or as an error logged.
- **Evidence trace:** Line 124 vs. line 112.
- **Fix:** Track a `Volatile.Read(ref reg.Disconnected)` flag set by DisconnectData; FlushTick skips disconnected regs. Or use a separate snapshot list maintained on Subscribe/Unsubscribe.
- **Confidence:** 0.72.

### RTDv2-011 — Medium — FlushTick is not re-entrancy guarded

- **Location:** `RtdServer.cs:118-139`.
- **Category:** Reentrancy / overlap.
- **Scenario:** `new Timer(FlushTick, null, 250, 250)` schedules a callback every 250ms. The Timer does NOT serialize callbacks — if a previous callback hasn't finished when the period elapses, a SECOND callback is dispatched on a different ThreadPool thread. With 50k topics and per-tick allocation/boxing/Equals, a tick can exceed 250ms under load. Now two threads enumerate `_topics`, both compare `reg.LastPushed` (read), both compute "current != LastPushed" (both observe true), both call `UpdateValue`, both write `reg.LastPushed = current` (last-writer-wins). Excel receives duplicate updates; LastPushed is in a torn but eventually-consistent state.
- **Evidence trace:** Line 65 (Timer ctor) + lines 118-138 (no overlap guard).
- **Fix:** Add `if (Interlocked.Exchange(ref _flushBusy, 1) == 1) return;` at the top, `Interlocked.Exchange(ref _flushBusy, 0);` in a finally. Or use `Timer` with `dueTime = Timeout.Infinite` after each callback and re-schedule from inside the callback (self-pacing).
- **Confidence:** 0.74.

### RTDv2-012 — Medium — No backpressure on Excel-DNA UpdateValue queue

- **Location:** `RtdServer.cs:118-139`.
- **Category:** Unbounded queue / no backpressure.
- **Scenario:** `reg.Topic.UpdateValue(current ?? string.Empty)` enqueues into Excel-DNA's per-server dispatch. If Excel is in a modal dialog / manual-recalc-only / debugger paused / printing, the queue fills. Each tick adds (50k × changed) entries. Over minutes this is millions of entries pending. Memory bloat → eventual OOM. No drop policy, no coalescing (the flush only sends the latest *changed* value, so coalescing of multiple-changes-since-last-flush is already implicit — but if Excel never drains, even one update per topic per tick builds up).
- **Evidence trace:** Line 130 fire-and-forget.
- **Fix:** Track per-topic last-known-acked time; if a topic hasn't been acked for N flushes, skip (Excel will re-pull on its own re-eval). Or query Excel-DNA's pending count and stop pushing when high-water-mark exceeded. Excel-DNA does not expose this metric directly today, so the practical fix is to gate the timer on `XlCall.RTD` health checks or use a fixed cap of in-flight updates.
- **Confidence:** 0.65.

### RTDv2-013 — Medium — O(n) per-tick scan over all topics regardless of dirty-set

- **Location:** `RtdServer.cs:118-139`.
- **Category:** Hidden complexity / wasted CPU.
- **Scenario:** Every 250ms, walk all N topics, read `LatestValue` (volatile read = ordinary mov on x86 but still a load), box-compare with `LastPushed`. At N=50k this is ~200k volatile reads/sec and N boxed Equals calls. The CHANGED set is typically a small fraction. Better: feeds publish to a dirty-flag queue; flush drains the queue.
- **Evidence trace:** Lines 124-133.
- **Fix:** Maintain a `ConcurrentQueue<TopicRegistration>` of "topics with new value since last flush"; producer marks itself dirty when it writes LatestValue (with a CAS guard to avoid duplicate enqueues). FlushTick drains the queue. O(changed) not O(total).
- **Confidence:** 0.55.

### RTDv2-014 — Medium — Captured token outlives source disposal in Reset

- **Location:** `ToolkitLifetime.cs:25-34` + Reset at 40-47.
- **Category:** Token capture / source disposal race.
- **Scenario:** A consumer calls `var ct = ToolkitLifetime.ShutdownToken;` and stashes it. They later call `Task.Delay(ms, ct)` — internally `ct.Register(callback)`. If `Reset()` runs in between, the underlying CTS is `Dispose()`'d. The stashed `ct` value still has a back-pointer to the disposed CTS. `Task.Delay`'s Register on a disposed CTS throws `ObjectDisposedException`, which propagates as a TaskCanceled/faulted to the caller. The bug surfaces as "my long-running operation suddenly threw ObjectDisposedException after a shutdown/reopen cycle".
- **Evidence trace:** Line 31 returns the token. Line 44 disposes the source it came from.
- **Fix:** `Reset()` should cancel the old source first (line 43.5: `_cts.Cancel();`), then dispose. The cancellation makes captured tokens behave as cancelled rather than throwing on Register. Already the case if `Shutdown()` was called before `Reset()`. If `Reset()` is called without prior Shutdown, the disposal-without-cancel makes the disposed source observable as such.
- **Confidence:** 0.55.

### RTDv2-015 — Medium — Stop() does not await producer; orphan Task continues briefly

- **Location:** `RtdServer.cs:355-371`.
- **Category:** Cleanup / partial shutdown.
- **Scenario:** `Stop()` cancels the CTS and immediately nulls `_producer`. The producer Task is still scheduled (or running) on the ThreadPool. The Task holds the (now-disposed) CTS via its `cancellationToken` parameter. The Task continues running until it next checks the token (next iteration of the while loop or next `await Task.Delay` resumption). Between cancel and observation, additional `LatestValue` writes happen — visible to any concurrent FlushTick. After Stop, FeedManager.Shutdown calls `_feeds.Clear()` (line 240), dropping the only reference to the Feed object. The Feed is GC-eligible *while its producer Task is still running and writing to its LatestValue field*. The Task itself keeps the Feed alive via captured `this` in the lambda at line 335, so no use-after-free, but the GC pressure picture is non-obvious.
- **Evidence trace:** Lines 357-370 — no `await _producer` or `_producer.Wait`.
- **Fix:** Capture `var t = _producer;` inside the lock, exit lock, `t?.Wait(TimeSpan.FromSeconds(2));` outside the lock (with a finite timeout to bound shutdown).
- **Confidence:** 0.70.

### RTDv2-016 — Low (rejected) — TraceSource allocation pattern

See table. Not a finding.

### RTDv2-017 — Medium — Unbounded `_feeds` growth on adversarial specs

- **Location:** `RtdServer.cs:91`, `RtdServer.cs:219-226`.
- **Category:** Resource exhaustion via spec churn.
- **Scenario:** A workbook with cells like `=EPT.RTD("counter:" & ROW())` produces a distinct feed per row. 1M cells → 1M live Feed instances, each with a Task running on the ThreadPool. ThreadPool starvation, memory exhaustion. No cap, no validation, no LRU eviction.
- **Evidence trace:** Line 225 `_feeds.GetOrAdd(spec, ...)` with no cap.
- **Fix:** Validate and cap. Reject `counter:N` with N outside `[10ms, 1h]`. Cap total distinct specs at e.g. 10000; reject (return `ExcelError.ExcelErrorValue`) when full. Add a metric.
- **Confidence:** 0.60.

### RTDv2-018 — Low (rejected) — Static lambda already cached

Not a finding.

### RTDv2-019 — Low — Restart loop with no backoff if producer throws on first iteration

- **Location:** `RtdServer.cs:330-336`.
- **Category:** Hot-restart loop.
- **Scenario:** Subscribe sees `_producer.IsCompleted`. If a custom Feed implementation's `RunAsync` throws synchronously before its first `await`, `RunSafelyAsync` catches and Task completes. Next Subscribe (or any code path that triggers Subscribe — there's just the one) starts another, which throws again. Today no caller invokes Subscribe in a tight loop, so the practical impact is one extra restart per ConnectData, not a hot loop. But if Excel-DNA re-issues ConnectData rapidly during recalc, this becomes a CPU pit.
- **Evidence trace:** Lines 330-336.
- **Fix:** Track `_consecutiveFailures`; back off exponentially. Or after N failures, leave `LatestValue = ExcelError.ExcelErrorNA` and refuse further restarts until manual reset.
- **Confidence:** 0.50.

### RTDv2-020 — Low — Boxing in per-tick `Equals` for double-valued feeds

- **Location:** `RtdServer.cs:127-131`.
- **Category:** Allocation / CPU waste.
- **Scenario:** `LatestValue` is `object?`. CounterFeed writes `(double)v` → box. SineFeed/RandomFeed same. Each tick: `current` and `reg.LastPushed` are both already-boxed doubles (references). `ReferenceEquals` is false (distinct boxes per producer write). `Equals(object?, object?)` unboxes both, returns equality of underlying double values. Two unboxings per topic per tick. At N=50k topics × 4 ticks/sec = 200k unboxings/sec. Trivial CPU cost; no allocation here (unboxing doesn't allocate). The allocation is upstream in the producer (one box per `LatestValue =` assignment).
- **Evidence trace:** Lines 127-131 + lines 428, 452, 480.
- **Fix:** Specialize per-feed value type. Or use a typed `Feed<T>` and avoid object. Out of scope for a small fix.
- **Confidence:** 0.45.

### RTDv2-021 — Low — `ParseSineArgs` silently swallows malformed input

- **Location:** `RtdServer.cs:271-285`.
- **Category:** Usability / silent default.
- **Scenario:** `sine:notanumber:alsobad` → `(1.0, 1.0)`. User believes they set freq=notanumber, gets default behavior, debugs for hours.
- **Fix:** Throw `ArgumentException` on parse failure of any non-empty arg; let the caller see `#VALUE!`.
- **Confidence:** 0.40.

### RTDv2-022 — Low — `ParseInt` silently defaults on non-positive values

- **Location:** `RtdServer.cs:266-269`.
- **Category:** Usability / silent default.
- **Scenario:** `counter:0` or `counter:-5` → 1000ms. User intent unclear; failing loud (`ArgumentException`) would surface formula errors.
- **Fix:** Throw or return `ExcelError.ExcelErrorValue` from `GetOrCreate` on invalid args.
- **Confidence:** 0.35.

### RTDv2-023 — Low — ClockFeed has no initial value before first iteration

- **Location:** `RtdServer.cs:398-410`.
- **Category:** Cold-start UX.
- **Scenario:** `ClockFeed` does not seed `LatestValue` in ctor. First Subscribe → `feed.LatestValue` is `null` → `ConnectData` returns `string.Empty`. Cell shows blank for up to ~250ms (Task scheduling + first iteration) until first push. CounterFeed/SineFeed/RandomFeed all seed `(double)0` in ctor (lines 420, 444, 468).
- **Fix:** In ClockFeed ctor: `LatestValue = DateTime.UtcNow.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture);`.
- **Confidence:** 0.30.

### RTDv2-024 — Low — SineFeed uses wall-clock `DateTime.UtcNow`

- **Location:** `RtdServer.cs:438, 451`.
- **Category:** Time handling.
- **Scenario:** NTP step backward → `t = (UtcNow - _epoch).TotalSeconds` becomes smaller than the previous tick or even negative; sine waveform has a visible glitch.
- **Fix:** Use `Stopwatch _sw = Stopwatch.StartNew();` in ctor, read `_sw.Elapsed.TotalSeconds` per tick. Monotonic.
- **Confidence:** 0.40.

### RTDv2-025 — Low — RandomFeed `_gate` is dead code

- **Location:** `RtdServer.cs:460, 476-479`.
- **Category:** Dead synchronization.
- **Scenario:** Only one producer Task per Feed; only the producer touches `_rng`. The `lock (_gate)` around `NextDouble()` serializes a single thread — no-op overhead. Misleading for readers (implies multi-thread access).
- **Fix:** Remove the lock; add a one-line comment explaining the single-producer invariant.
- **Confidence:** 0.55.

### RTDv2-026 — rejected — ServerStart cannot leave inconsistent state

Excel-DNA contracts ServerStart to complete before any ConnectData. The Timer construction is the only failable step and any throw aborts server registration cleanly. Not a finding.

---

## Round-2 inspection notes (the 16 prompts) — explicit dispositions

1. **CTS finalizer interactions:** confirmed real. See RTDv2-007. Finalizer disposes a leaked CTS on a non-deterministic thread; if any captured token's registrations are live, the disposal races.
2. **CreateLinkedTokenSource synchronous fire when parent already cancelled:** confirmed real. See RTDv2-008.
3. **ClockFeed cold start:** confirmed minor. See RTDv2-023.
4. **`_latestValue` initial null:** safe (Volatile.Read returns null cleanly). Not a finding.
5. **Subscribe rerun semantics after both unsubscribe:** safe. `_producer.IsCompleted` correctly triggers restart with fresh CTS. Not a finding.
6. **`IsCancellationRequested` on token of disposed CTS:** per .NET docs (`CancellationToken.IsCancellationRequested`), reading a cancellation state on a token whose source has been disposed is safe — the token caches the cancelled-or-not state via the CTS internal `_state` field, and disposal does not clear it. `Register` does throw, however. RTDv2-014 covers the Register path; RTDv2-004 covers the Task.Delay path. The bare IsCancellationRequested read is not a finding.
7. **RandomFeed lock dead code:** RTDv2-025. Single producer confirmed.
8. **SineFeed `_epoch`:** safe. Single writer (ctor), immutable thereafter. Not a finding.
9. **CounterFeed Interlocked on single-thread field:** safe but redundant. Not a finding (cosmetic).
10. **`_cts` field non-volatile but lock-protected:** safe. Lock entry has full memory fence. Not a finding.
11. **`ShutdownToken` returned and used after Reset:** real, see RTDv2-014.
12. **Lazy<FeedManager> reused across reloads:** the singleton survives AddIn unload/reload. After Shutdown clears `_feeds`, the instance is in a clean state — but Lazy<T> itself never recreates the instance. Acceptable. No retained per-reload state I can find. Not a finding.
13. **Double-shutdown across reloads:** idempotent today. Cross-domain (AddIn.cs) — see rejected appendix.
14. **Excel-DNA UpdateValue queue:** likely bounded by Excel-DNA internally but not documented. RTDv2-012 captures the risk.
15. **CounterFeed overflow:** RTDv2 column documents only. Not actionable.
16. **SineFeed clock skew:** RTDv2-024.

---

## Rejected findings appendix

- **RTDv2-016 (TraceSource allocation pattern):** callers cache in `static readonly`; pattern is fine.
- **RTDv2-018 (static lambda allocation):** confirmed cached; not a finding.
- **RTDv2-026 (ServerStart inconsistent state):** Excel-DNA serializes ServerStart before any ConnectData; not reachable.
- **`_cts` field non-volatile:** the lock provides the memory fence. Not a finding.
- **`Interlocked.Increment` on CounterFeed._value:** redundant given single writer, but not incorrect. Cosmetic.
- **SineFeed `_epoch` race:** single ctor writer, multi-reader. .NET memory model guarantees ctor stores are visible after publication of `this`. Safe.
- **`ShutdownToken.IsCancellationRequested` read after CTS dispose:** safe per .NET docs.
- **Lazy<FeedManager> retained state across reloads:** none. After Shutdown clears `_feeds`, no stale state.

### Cross-domain observations (NOT findings — outside scope per instructions)

- `AddIn.cs` (forbidden) calls `FeedManager.Instance.Shutdown()` on AutoClose; `RtdServer.ServerTerminate` also calls it. Currently idempotent but undocumented. Noted only.
- `AddIn.cs` interaction with `ToolkitLifetime.Shutdown`/`Reset` is the trigger condition for RTDv2-008. The trigger comes from `AddIn.cs`; the bug is in `RtdServer.cs` Subscribe.
- The "infinite stillborn Task restart" of RTDv2-008 only manifests if some other component (e.g. a test harness or a re-open path) calls Subscribe after Shutdown without intervening Reset. AddIn.cs typically calls Reset on AutoOpen, mitigating the trigger — but if AutoOpen is skipped or fails, the failure mode is reachable.

---

## Severity summary

- **Critical:** 2 (RTDv2-001, RTDv2-002)
- **High:** 6 (RTDv2-003, -004, -005, -006, -007, -008)
- **Medium:** 7 (RTDv2-009, -010, -011, -012, -013, -014, -015, -017)
- **Low:** 6 (RTDv2-019, -020, -021, -022, -023, -024, -025)
- **Rejected:** 4

The two Critical findings are both about lifecycle: a Timer that doesn't await its callback before tearing down state, and an exception scope that lets one bad topic poison every sister topic for the entire tick. Both should be fixed together — they compound (a single throw plus a single in-flight-after-Dispose creates exactly the use-after-release scenario at the COM boundary).
