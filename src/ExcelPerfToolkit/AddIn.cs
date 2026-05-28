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
    }

    /// <summary>
    /// Called by Excel-DNA when the workbook is closing or the add-in is unloaded.
    /// Marshaling cost: 0 boundary crossings.
    /// Thread-safety: invoked on Excel's main thread only.
    /// </summary>
    public void AutoClose()
    {
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
