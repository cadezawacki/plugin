using System;
using System.Diagnostics;
using System.Threading;

namespace ExcelPerfToolkit;

/// <summary>
/// Process-wide lifetime and diagnostics support shared by the second-wave features
/// (vectorized kernels, RTD server, direct file I/O). Holds the toolkit-wide
/// <see cref="CancellationTokenSource"/> so background schedulers and async file I/O
/// can be cancelled cleanly when the add-in unloads.
/// Marshaling cost: 0 boundary crossings.
/// Thread-safety: all members are safe to call from any thread.
/// </summary>
internal static class ToolkitLifetime
{
    private static CancellationTokenSource _cts = new();
    // Cache the token alongside the source. CancellationTokenSource.Token throws
    // ObjectDisposedException after Dispose, so we read it once at construction and
    // hand callers the cached struct. Token is a struct of refs, safe to copy.
    private static CancellationToken _token = _cts.Token;
    private static readonly object Gate = new();

    /// <summary>
    /// Cancellation token signaled when the add-in is shutting down. Any background
    /// thread, RTD feed, or async file operation that wants to honor add-in unload
    /// should observe this token. Safe to read after Shutdown disposes the source
    /// because we return the token captured at the last Reset/init.
    /// </summary>
    public static CancellationToken ShutdownToken
    {
        get
        {
            lock (Gate)
            {
                return _token;
            }
        }
    }

    /// <summary>
    /// Resets the lifetime to a fresh, non-cancelled state. Called from
    /// <see cref="AddIn.AutoOpen"/> so that a re-open after unload starts clean.
    /// </summary>
    public static void Reset()
    {
        lock (Gate)
        {
            // Tolerate a prior Shutdown() (which already disposed the CTS) as well as
            // the steady-state case where _cts is live. Either way, replace with a fresh
            // source and refresh the captured token.
            try
            {
                _cts.Dispose();
            }
            catch (ObjectDisposedException)
            {
                // Already disposed by Shutdown(); benign.
            }
            _cts = new CancellationTokenSource();
            _token = _cts.Token;
        }
    }

    /// <summary>
    /// Signals shutdown to every observer and disposes the source. Subsequent
    /// access to <see cref="ShutdownToken"/> yields the already-cancelled token
    /// captured before disposal, never a disposed CTS.
    /// </summary>
    public static void Shutdown()
    {
        lock (Gate)
        {
            try
            {
                _cts.Cancel();
            }
            catch (ObjectDisposedException)
            {
                // Already shut down; benign.
            }
            // Snapshot the cancelled token before disposing the source so the
            // ShutdownToken getter keeps returning a usable (cancelled) token
            // after dispose without touching the disposed CTS.
            _token = _cts.Token;
            // Dispose the CTS after cancelling so the native SafeWaitHandle is released
            // promptly rather than waiting for the finalizer (which races consumers that
            // captured the token).
            try
            {
                _cts.Dispose();
            }
            catch (ObjectDisposedException)
            {
                // Already disposed; benign.
            }
        }
    }

    /// <summary>
    /// Creates a <see cref="TraceSource"/> named consistently with the rest of the
    /// toolkit. Centralized so new files don't drift in naming.
    /// </summary>
    public static TraceSource CreateTraceSource(string componentName)
    {
        if (string.IsNullOrWhiteSpace(componentName))
        {
            throw new ArgumentException("Component name is required.", nameof(componentName));
        }
        return new TraceSource("ExcelPerfToolkit." + componentName, SourceLevels.Information);
    }
}
