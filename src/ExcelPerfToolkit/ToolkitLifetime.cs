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
    private static readonly object Gate = new();

    /// <summary>
    /// Cancellation token signaled when the add-in is shutting down. Any background
    /// thread, RTD feed, or async file operation that wants to honor add-in unload
    /// should observe this token.
    /// </summary>
    public static CancellationToken ShutdownToken
    {
        get
        {
            lock (Gate)
            {
                return _cts.Token;
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
            _cts.Dispose();
            _cts = new CancellationTokenSource();
        }
    }

    /// <summary>
    /// Signals shutdown to every observer and disposes the source. Subsequent
    /// access to <see cref="ShutdownToken"/> yields an already-cancelled token
    /// until <see cref="Reset"/> is called.
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
