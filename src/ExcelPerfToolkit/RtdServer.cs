using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;

namespace ExcelPerfToolkit;

/// <summary>
/// Multithreaded RTD server that addresses the "real-time feeds throttled through
/// single-threaded marshaling" bottleneck. Excel's built-in RTD model is friendly to
/// background threads, but most homegrown RTD code pumps every tick straight through
/// to Excel's UI thread - which serializes and floods. This server fixes that with
/// three coordinated pieces:
///
///   * A concurrency-safe <see cref="FeedManager"/> that multiplexes many topics over
///     a single producer per distinct feed specification. Two thousand cells
///     subscribing to <c>clock</c> share one timer, not two thousand.
///   * Background producers that push raw values into volatile per-feed state. They
///     never call Excel.
///   * A throttle <see cref="Timer"/> owned by the server that walks subscribed
///     topics every <see cref="DefaultFlushIntervalMs"/> milliseconds and calls
///     <see cref="ExcelRtdServer.Topic.UpdateValue(object)"/> only when the value has
///     actually changed since the last push. Excel-DNA marshals the UpdateValue calls
///     onto Excel's RTD thread - that boundary is the only place the server crosses
///     into Excel.
///
/// The RTD callback into Excel happens on Excel's thread; data acquisition happens on
/// background threads. The boundary is kept safe because:
///   (a) producers write to volatile fields only,
///   (b) the throttle timer copies those fields once per tick into local variables
///       before calling UpdateValue,
///   (c) the throttle timer is the sole writer of <c>UpdateValue</c>; there is no
///       fan-out from producer threads directly into Excel.
/// </summary>
[ComVisible(true)]
[ProgId(ProgIdValue)]
public sealed class RtdServer : ExcelRtdServer
{
    /// <summary>
    /// The ProgId Excel will use to address this server via <c>=RTD("EPT.Rtd", ...)</c>.
    /// </summary>
    public const string ProgIdValue = "EPT.Rtd";

    /// <summary>
    /// Default outbound throttle interval. Excel can keep up with this without UI
    /// stutter even when many topics are subscribed; producers may run faster.
    /// </summary>
    public const int DefaultFlushIntervalMs = 250;

    private static readonly TraceSource TraceSource = ToolkitLifetime.CreateTraceSource("RtdServer");

    private readonly ConcurrentDictionary<int, TopicRegistration> _topics = new();
    private Timer? _flushTimer;
    // Reentrancy gate for FlushTick. If a tick takes longer than the period (e.g. with
    // tens of thousands of topics or a slow UpdateValue queue), Timer fires the
    // callback again on a different ThreadPool thread. Without this gate, two threads
    // would both iterate _topics and race writes to TopicRegistration.LastPushed.
    private int _flushInFlight;

    /// <inheritdoc/>
    protected override bool ServerStart()
    {
        TraceSource.TraceInformation("RTD server starting (flush interval = {0} ms).", DefaultFlushIntervalMs);
        _flushTimer = new Timer(FlushTick, state: null, dueTime: DefaultFlushIntervalMs, period: DefaultFlushIntervalMs);
        return true;
    }

    /// <inheritdoc/>
    protected override void ServerTerminate()
    {
        TraceSource.TraceInformation("RTD server terminating, releasing {0} topics.", _topics.Count);
        var timer = Interlocked.Exchange(ref _flushTimer, null);
        // Invariant: by the time we return from this method, no FlushTick callback
        // can still be running. Timer.Dispose() alone returns immediately without
        // waiting for a pending callback, so we use the WaitHandle overload to join.
        if (timer is not null)
        {
            using var done = new ManualResetEvent(false);
            if (timer.Dispose(done))
            {
                done.WaitOne();
            }
        }
        foreach (var reg in _topics.Values)
        {
            reg.Feed.Unsubscribe(reg);
        }
        _topics.Clear();
        FeedManager.Instance.Shutdown();
    }

    /// <inheritdoc/>
    protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
    {
        ArgumentNullException.ThrowIfNull(topic);
        ArgumentNullException.ThrowIfNull(topicInfo);
        try
        {
            var spec = topicInfo.Count > 0 ? topicInfo[0] : string.Empty;
            var feed = FeedManager.Instance.GetOrCreate(spec);
            var reg = new TopicRegistration(topic, feed);
            // If Excel re-issues ConnectData for the same TopicId without an
            // intervening DisconnectData (rare but observed), the prior registration
            // would otherwise be silently orphaned in its Feed's _subscribers - keeping
            // that Feed's producer running forever for a topic we no longer track.
            // Unsubscribe the old reg before overwriting.
            if (_topics.TryGetValue(topic.TopicId, out var existing))
            {
                existing.Feed.Unsubscribe(existing);
            }
            _topics[topic.TopicId] = reg;
            feed.Subscribe(reg);
            newValues = true;
            return feed.LatestValue ?? string.Empty;
        }
        catch (Exception ex)
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 1, "RTD ConnectData failed for topic {0}: {1}", topic.TopicId, ex.Message);
            return ExcelError.ExcelErrorValue;
        }
    }

    /// <inheritdoc/>
    protected override void DisconnectData(Topic topic)
    {
        if (topic is null)
        {
            return;
        }
        if (_topics.TryRemove(topic.TopicId, out var reg))
        {
            reg.Feed.Unsubscribe(reg);
        }
    }

    private void FlushTick(object? state)
    {
        // Runs on a ThreadPool thread, NOT on Excel's RTD thread. UpdateValue itself
        // is the COM-marshaled boundary; Excel-DNA hands it across safely.
        //
        // Reentrancy gate: if a previous tick is still running (e.g. large topic set
        // backing up Excel-DNA's marshaling queue), skip this tick rather than racing
        // a sibling iteration over the same _topics and LastPushed writes.
        if (Interlocked.CompareExchange(ref _flushInFlight, 1, 0) != 0)
        {
            return;
        }
        try
        {
            foreach (var kv in _topics)
            {
                // Per-topic try/catch: an exception inside one Topic.UpdateValue must
                // NOT abort the rest of the tick. A deterministically-throwing topic
                // would otherwise starve every sibling topic of updates forever.
                var reg = kv.Value;
                try
                {
                    var current = reg.Feed.LatestValue;
                    if (!ReferenceEquals(current, reg.LastPushed) && !Equals(current, reg.LastPushed))
                    {
                        reg.Topic.UpdateValue(current ?? string.Empty);
                        reg.LastPushed = current;
                    }
                }
                catch (Exception ex)
                {
                    TraceSource.TraceEvent(TraceEventType.Warning, 2, "RTD flush failed for topic {0}: {1}", kv.Key, ex.Message);
                }
            }
        }
        finally
        {
            Volatile.Write(ref _flushInFlight, 0);
        }
    }

    // ====================================================================
    // Worksheet UDF wrapper
    // ====================================================================

    /// <summary>
    /// Thread-safe worksheet UDF that subscribes to an RTD topic and returns the
    /// current value. Excel re-evaluates RTD cells on its own schedule; the throttled
    /// server pushes updates via <see cref="ExcelRtdServer.Topic.UpdateValue(object)"/>.
    /// Feed specifications recognized out of the box:
    ///   * <c>clock</c>                  - current wall-clock as text, ticks once per second.
    ///   * <c>counter:N</c>              - increments every N milliseconds.
    ///   * <c>sine:freqHz:amplitude</c>  - sine of wall-clock time, useful for demos.
    ///   * <c>random</c>                 - uniform random doubles in [0, 1).
    /// Marshaling cost: 1 crossing on subscribe (Excel issues an array-formula call),
    /// then exactly one further crossing per throttled push regardless of underlying
    /// producer rate.
    /// Thread-safety: SAFE - registered with <c>IsThreadSafe = true</c>; the body just
    /// hands the topic spec to <see cref="XlCall.RTD(string, string, string[])"/>.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.RTD",
        Description = "Subscribe to an RTD topic. Updates flow on background threads, throttled to Excel.",
        Category = "EPT.Rtd",
        IsThreadSafe = true,
        IsVolatile = false)]
    public static object EptRtd(
        [ExcelArgument(Name = "topic", Description = "Feed spec: 'clock', 'counter:500', 'sine:1:5', 'random'.")] string topic)
    {
        if (string.IsNullOrWhiteSpace(topic))
        {
            return ExcelError.ExcelErrorValue;
        }
        return XlCall.RTD(ProgIdValue, null, topic) ?? ExcelError.ExcelErrorNA;
    }
}

/// <summary>
/// Records a single subscription: which <see cref="ExcelRtdServer.Topic"/> subscribed,
/// which <see cref="Feed"/> it points at, and what value was last pushed to Excel.
/// </summary>
internal sealed class TopicRegistration
{
    public TopicRegistration(ExcelRtdServer.Topic topic, Feed feed)
    {
        Topic = topic;
        Feed = feed;
    }

    public ExcelRtdServer.Topic Topic { get; }

    public Feed Feed { get; }

    public object? LastPushed { get; set; }
}

/// <summary>
/// Process-wide registry of feeds. Multiple topics subscribing to the same feed spec
/// share the same underlying <see cref="Feed"/> instance - so one timer drives any
/// number of cells.
/// </summary>
internal sealed class FeedManager
{
    private static readonly TraceSource TraceSource = ToolkitLifetime.CreateTraceSource("FeedManager");
    private static readonly Lazy<FeedManager> Lazy = new(() => new FeedManager(), LazyThreadSafetyMode.ExecutionAndPublication);

    public static FeedManager Instance => Lazy.Value;

    private readonly ConcurrentDictionary<string, Feed> _feeds = new(StringComparer.OrdinalIgnoreCase);
    private readonly object _gate = new();

    private FeedManager()
    {
    }

    /// <summary>
    /// Gets or creates the feed for <paramref name="spec"/>. Concurrency-safe: only one
    /// feed instance is created per distinct spec even under racing connect calls.
    /// </summary>
    /// <summary>
    /// Hard cap on the number of distinct feed specs we'll create. Each feed owns a
    /// background Task; an attacker (or buggy formula like <c>="counter:"&amp;ROW()</c>)
    /// could otherwise create unbounded feeds.
    /// </summary>
    public const int MaxDistinctFeeds = 1024;

    public Feed GetOrCreate(string spec)
    {
        if (string.IsNullOrWhiteSpace(spec))
        {
            throw new ArgumentException("Spec is required.", nameof(spec));
        }
        // Cap is checked optimistically; if a parallel call races us across the cap
        // we tolerate one or two extra registrations - the dictionary's own concurrency
        // ensures consistency.
        if (_feeds.Count >= MaxDistinctFeeds && !_feeds.ContainsKey(spec))
        {
            throw new InvalidOperationException($"Feed registry full ({MaxDistinctFeeds} distinct specs). Reduce spec variation.");
        }
        return _feeds.GetOrAdd(spec, static key => CreateFeed(key));
    }

    /// <summary>
    /// Stop every feed and clear the registry. Called from
    /// <see cref="RtdServer.ServerTerminate"/> and from <see cref="AddIn.AutoClose"/>.
    /// </summary>
    public void Shutdown()
    {
        lock (_gate)
        {
            // try/finally so that any single Feed.Stop() exception still leaves the
            // registry empty - otherwise a stale dead feed lingers across re-opens.
            try
            {
                foreach (var f in _feeds.Values)
                {
                    try
                    {
                        f.Stop();
                    }
                    catch (Exception ex)
                    {
                        TraceSource.TraceEvent(TraceEventType.Warning, 4, "Feed.Stop threw during shutdown for '{0}': {1}", f.Spec, ex.Message);
                    }
                }
            }
            finally
            {
                _feeds.Clear();
            }
        }
        TraceSource.TraceInformation("FeedManager shut down.");
    }

    private static Feed CreateFeed(string spec)
    {
        var trimmed = spec.Trim();
        var head = trimmed;
        var args = string.Empty;
        var sep = trimmed.IndexOf(':', StringComparison.Ordinal);
        if (sep > 0)
        {
            head = trimmed.Substring(0, sep);
            args = trimmed.Substring(sep + 1);
        }
        return head.ToLowerInvariant() switch
        {
            "clock" => new ClockFeed(trimmed),
            "counter" => new CounterFeed(trimmed, ParseInt(args, defaultValue: 1000)),
            "sine" => new SineFeed(trimmed, ParseSineArgs(args)),
            "random" => new RandomFeed(trimmed),
            "watchfile" => new WatchFileFeed(trimmed, args),
            "watchfolder" => new WatchFolderFeed(trimmed, args),
            _ => throw new ArgumentException($"Unknown feed type '{head}'. Expected one of: clock, counter, sine, random, watchfile, watchfolder."),
        };
    }

    private static int ParseInt(string s, int defaultValue)
    {
        return int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) && v > 0 ? v : defaultValue;
    }

    private static (double FreqHz, double Amplitude) ParseSineArgs(string args)
    {
        var freq = 1d;
        var amp = 1d;
        var parts = args.Split(':');
        if (parts.Length > 0 && double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out var f))
        {
            freq = f;
        }
        if (parts.Length > 1 && double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var a))
        {
            amp = a;
        }
        return (freq, amp);
    }
}

/// <summary>
/// Base class for background feed producers. A feed owns its own background loop and
/// publishes the latest value to a volatile field; the RTD server's throttle timer is
/// the only thing that propagates that value across the Excel boundary.
/// </summary>
internal abstract class Feed
{
    private static readonly TraceSource TraceSource = ToolkitLifetime.CreateTraceSource("Feed");

    private readonly object _gate = new();
    private readonly ConcurrentDictionary<int, TopicRegistration> _subscribers = new();
    private CancellationTokenSource? _cts;
    private Task? _producer;

    protected Feed(string spec)
    {
        Spec = spec;
    }

    public string Spec { get; }

    /// <summary>
    /// Latest produced value. Volatile assignment is provided by the
    /// <see cref="Interlocked"/>/<c>volatile</c> semantics of object reference writes
    /// on .NET; consumers read this on the throttle thread.
    /// </summary>
    public object? LatestValue
    {
        get => Volatile.Read(ref _latestValue);
        protected set => Volatile.Write(ref _latestValue, value);
    }

    private object? _latestValue;

    /// <summary>
    /// Adds <paramref name="reg"/> to the subscriber set and starts the producer if it
    /// is not already running. Concurrency-safe.
    /// </summary>
    public void Subscribe(TopicRegistration reg)
    {
        // Invariant: _subscribers mutation and the start/stop decision are made under
        // the same lock so a concurrent Unsubscribe cannot read IsEmpty and call Stop
        // between the add and the producer-start.
        lock (_gate)
        {
            _subscribers[reg.Topic.TopicId] = reg;
            // Subscribe-after-shutdown is a no-op: a linked CTS over an already-cancelled
            // parent would fire synchronously and Task.Run with a cancelled token never
            // invokes the action, leaving _producer.IsCompleted == true and looping us
            // into a stillborn-restart pattern on every subsequent Subscribe.
            if (ToolkitLifetime.ShutdownToken.IsCancellationRequested)
            {
                return;
            }
            if (_producer is null || _producer.IsCompleted)
            {
                _cts = CancellationTokenSource.CreateLinkedTokenSource(ToolkitLifetime.ShutdownToken);
                var ct = _cts.Token;
                _producer = Task.Run(() => RunSafelyAsync(ct), ct);
            }
        }
    }

    /// <summary>
    /// Removes <paramref name="reg"/>; stops the producer if no subscribers remain.
    /// </summary>
    public void Unsubscribe(TopicRegistration reg)
    {
        // Invariant: the Remove + IsEmpty check + Stop happen under the lock as a
        // single atomic decision, so a concurrent Subscribe cannot see "no producer"
        // immediately after we decided to stop one.
        lock (_gate)
        {
            _subscribers.TryRemove(reg.Topic.TopicId, out _);
            if (_subscribers.IsEmpty)
            {
                StopLocked();
            }
        }
    }

    /// <summary>
    /// Stops the background producer. Safe to call multiple times.
    /// </summary>
    public void Stop()
    {
        lock (_gate)
        {
            StopLocked();
        }
    }

    /// <summary>
    /// Caller must hold <see cref="_gate"/>. Cancels the producer's CTS and clears the
    /// producer reference, but does NOT dispose the CTS while the producer may still
    /// be inside <c>await Task.Delay(token)</c> - Delay's internal registration cleanup
    /// races a synchronous Dispose. Letting the GC finalize the CTS after the Task
    /// completes is the safe path; tradeoff is a brief delay before the unmanaged
    /// wait handle is released.
    /// </summary>
    private void StopLocked()
    {
        try
        {
            _cts?.Cancel();
        }
        catch (ObjectDisposedException)
        {
            // Already disposed; benign.
        }
        _cts = null;
        _producer = null;
    }

    private async Task RunSafelyAsync(CancellationToken cancellationToken)
    {
        try
        {
            await RunAsync(cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            // Expected on shutdown / last unsubscribe.
        }
        catch (Exception ex)
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 3, "Feed '{0}' producer crashed: {1}", Spec, ex.Message);
            LatestValue = ExcelError.ExcelErrorNA;
        }
    }

    /// <summary>
    /// Producer loop. Implementations must observe <paramref name="cancellationToken"/>
    /// and use only non-blocking, asynchronous primitives (e.g.
    /// <see cref="Task.Delay(int, CancellationToken)"/>).
    /// </summary>
    protected abstract Task RunAsync(CancellationToken cancellationToken);
}

internal sealed class ClockFeed : Feed
{
    public ClockFeed(string spec) : base(spec) { }

    protected override async Task RunAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            LatestValue = DateTime.UtcNow.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture);
            await Task.Delay(250, cancellationToken).ConfigureAwait(false);
        }
    }
}

internal sealed class CounterFeed : Feed
{
    private readonly int _intervalMs;
    private long _value;

    public CounterFeed(string spec, int intervalMs) : base(spec)
    {
        _intervalMs = intervalMs;
        LatestValue = (double)0;
    }

    protected override async Task RunAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            var v = Interlocked.Increment(ref _value);
            LatestValue = (double)v;
            await Task.Delay(_intervalMs, cancellationToken).ConfigureAwait(false);
        }
    }
}

internal sealed class SineFeed : Feed
{
    private readonly double _freqHz;
    private readonly double _amplitude;
    private readonly DateTime _epoch = DateTime.UtcNow;

    public SineFeed(string spec, (double FreqHz, double Amplitude) args) : base(spec)
    {
        _freqHz = args.FreqHz;
        _amplitude = args.Amplitude;
        LatestValue = (double)0;
    }

    protected override async Task RunAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            var t = (DateTime.UtcNow - _epoch).TotalSeconds;
            LatestValue = _amplitude * Math.Sin(2 * Math.PI * _freqHz * t);
            await Task.Delay(50, cancellationToken).ConfigureAwait(false);
        }
    }
}

internal sealed class RandomFeed : Feed
{
    private readonly object _gate = new();
    private readonly Random _rng;

    public RandomFeed(string spec) : base(spec)
    {
        // Per-feed Random; not shared - so we don't violate "no shared mutable static
        // state" and don't need lock-free atomic doubles either.
        _rng = new Random();
        LatestValue = (double)0;
    }

    protected override async Task RunAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            double v;
            lock (_gate)
            {
                v = _rng.NextDouble();
            }
            LatestValue = v;
            await Task.Delay(100, cancellationToken).ConfigureAwait(false);
        }
    }
}
