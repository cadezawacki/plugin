using System;
using System.Diagnostics;
using ExcelDna.Integration;

namespace ExcelPerfToolkit;

/// <summary>
/// Add-in entry point. Registered automatically by Excel-DNA when the XLL loads.
/// Bottlenecks addressed: registration metadata advertises IsThreadSafe UDFs to Excel's
/// multithreaded recalc engine (bottleneck #3). The bulk transfer and developer utility
/// surfaces hosted by this add-in address bottlenecks #1 and #2.
/// </summary>
public sealed class AddIn : IExcelAddIn
{
    private static readonly TraceSource TraceSource = new("ExcelPerfToolkit", SourceLevels.Information);

    /// <summary>
    /// Called by Excel-DNA after the XLL has loaded. The function registry has already
    /// been populated from the public static methods in this assembly.
    /// Marshaling cost: 0 boundary crossings (executed once at load).
    /// Thread-safety: invoked on Excel's main thread only.
    /// </summary>
    public void AutoOpen()
    {
        var bitness = Environment.Is64BitProcess ? "x64" : "x86";
        TraceSource.TraceInformation(
            "ExcelPerfToolkit loaded. Process={0}, ExcelVersion={1}",
            bitness,
            SafeGetExcelVersion());
        ExcelIntegration.RegisterUnhandledExceptionHandler(OnUnhandledException);
        // Reset the toolkit-wide cancellation source so any background work spun up
        // by VectorizedKernels / RtdServer / DirectFileIO honors a fresh shutdown
        // token. The RTD server itself is registered automatically by Excel-DNA from
        // its [ProgId] attribute on first use; we only need to manage the lifetime
        // of background producers, which observe ToolkitLifetime.ShutdownToken.
        ToolkitLifetime.Reset();
    }

    /// <summary>
    /// Called by Excel-DNA when the workbook is closing or the add-in is unloaded.
    /// Marshaling cost: 0 boundary crossings.
    /// Thread-safety: invoked on Excel's main thread only.
    /// </summary>
    public void AutoClose()
    {
        // Cancel any background producers (RTD feeds, in-flight async file I/O) so
        // the unload returns promptly without leaving worker tasks running against
        // a torn-down host. Each subsystem observes ToolkitLifetime.ShutdownToken
        // and shuts down its own resources.
        ToolkitLifetime.Shutdown();
        FeedManager.Instance.Shutdown();
        TraceSource.TraceInformation("ExcelPerfToolkit unloaded.");
    }

    private static object OnUnhandledException(object exceptionObject)
    {
        var ex = exceptionObject as Exception;
        TraceSource.TraceEvent(TraceEventType.Error, 0, "Unhandled UDF exception: {0}", ex);
        return ExcelError.ExcelErrorValue;
    }

    private static string SafeGetExcelVersion()
    {
        try
        {
            var v = XlCall.Excel(XlCall.xlfGetWorkspace, 2);
            return v?.ToString() ?? "unknown";
        }
        catch (XlCallException)
        {
            return "unknown";
        }
    }
}
