using System;
using System.Diagnostics;
using ExcelDna.Integration;

namespace ExcelPerfToolkit;

/// <summary>
/// Bulk read / write between Excel and managed memory. This file is the direct fix for
/// bottlenecks #1 (cell-by-cell read/write) and #2 (per-cell COM marshaling): every
/// operation here is engineered to cross the Excel boundary exactly once on read, once
/// on write, regardless of how many cells are involved.
///
/// All functions in this file must run on Excel's main thread - they touch
/// <see cref="ExcelReference"/> and <see cref="XlCall"/> in ways that are illegal from
/// thread-safe UDFs. Use the <see cref="ParallelUtilities"/> file when you need
/// computation off the main thread, then come back here to write.
/// </summary>
public static class BulkTransfer
{
    private static readonly TraceSource TraceSource = new("ExcelPerfToolkit.BulkTransfer", SourceLevels.Information);

    /// <summary>
    /// Resolves <paramref name="a1Address"/> on <paramref name="sheetName"/> to an
    /// <see cref="ExcelReference"/>. The sheet may be qualified with a workbook
    /// (<c>[Book1.xlsx]Sheet1</c>) or be a bare sheet name.
    /// Marshaling cost: 2 boundary crossings (one to evaluate the textual reference,
    /// one is implicit in constructing the COM token). Both are constant in N.
    /// Thread-safety: NOT safe for thread-safe UDFs - resolves COM identity.
    /// Accepts: A1-style addresses. Returns: <see cref="ExcelReference"/> or throws.
    /// </summary>
    public static ExcelReference ResolveRange(string sheetName, string a1Address)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
        {
            throw new ArgumentException("Sheet name is required.", nameof(sheetName));
        }
        if (string.IsNullOrWhiteSpace(a1Address))
        {
            throw new ArgumentException("A1 address is required.", nameof(a1Address));
        }

        var qualified = a1Address.Contains('!', StringComparison.Ordinal)
            ? a1Address
            : string.Concat("'", sheetName.Replace("'", "''", StringComparison.Ordinal), "'!", a1Address);

        var token = XlCall.Excel(XlCall.xlfEvaluate, qualified);
        if (token is ExcelReference r)
        {
            return r;
        }
        if (token is ExcelError err)
        {
            throw new ArgumentException($"Range '{qualified}' could not be resolved: {Marshaling.ErrorToText(err)}.");
        }
        throw new ArgumentException($"Range '{qualified}' did not resolve to a cell reference.");
    }

    /// <summary>
    /// Reads the entire range as a dense <c>object[,]</c> in one boundary crossing,
    /// normalizing the single-cell scalar case so callers never have to special-case it.
    /// Marshaling cost: 1 boundary crossing.
    /// Thread-safety: NOT safe for thread-safe UDFs.
    /// Accepts: any valid <see cref="ExcelReference"/>. Returns: <c>object[,]</c>.
    /// Cells may be <see cref="double"/>, <see cref="string"/>, <see cref="bool"/>,
    /// <see cref="ExcelEmpty"/>, <see cref="ExcelError"/>, or a date serial as <see cref="double"/>.
    /// </summary>
    public static object[,] ReadBlock(ExcelReference reference)
    {
        ArgumentNullException.ThrowIfNull(reference);
        return Marshaling.AsArray2D(reference.GetValue());
    }

    /// <summary>
    /// Reads an entire A1 range on a named sheet in exactly one boundary crossing for
    /// the value transfer (the address resolution itself is a separate, constant-cost
    /// crossing - see <see cref="ResolveRange"/>).
    /// Marshaling cost: 1 read crossing + 1 resolve crossing = 2 total, both O(1) in N.
    /// Thread-safety: NOT safe for thread-safe UDFs.
    /// </summary>
    public static object[,] ReadBlock(string sheetName, string a1Address)
    {
        return ReadBlock(ResolveRange(sheetName, a1Address));
    }

    /// <summary>
    /// Writes <paramref name="block"/> to <paramref name="target"/> in exactly one
    /// boundary crossing. The shape of <paramref name="block"/> must match
    /// <paramref name="target"/>'s dimensions.
    /// Marshaling cost: 1 boundary crossing.
    /// Thread-safety: NOT safe for thread-safe UDFs - this writes to the workbook.
    /// Accepts: <c>object[,]</c> sized to the reference. Returns: nothing.
    /// </summary>
    public static void WriteBlock(ExcelReference target, object[,] block)
    {
        ArgumentNullException.ThrowIfNull(target);
        ArgumentNullException.ThrowIfNull(block);

        var refRows = target.RowLast - target.RowFirst + 1;
        var refCols = target.ColumnLast - target.ColumnFirst + 1;
        if (block.GetLength(0) != refRows || block.GetLength(1) != refCols)
        {
            throw new ArgumentException(
                $"Block shape {block.GetLength(0)}x{block.GetLength(1)} does not match target {refRows}x{refCols}.",
                nameof(block));
        }

        var ok = target.SetValue(block);
        if (!ok)
        {
            throw new InvalidOperationException("Excel rejected the bulk write. The range may be protected or in an invalid state.");
        }
    }

    /// <summary>
    /// Writes <paramref name="block"/> into the rectangle anchored at the top-left of
    /// the A1 range. Caller is responsible for ensuring there is enough room.
    /// Marshaling cost: 1 write crossing + 1 resolve crossing = 2 total.
    /// Thread-safety: NOT safe for thread-safe UDFs.
    /// </summary>
    public static void WriteBlock(string sheetName, string anchorA1, object[,] block)
    {
        ArgumentNullException.ThrowIfNull(block);
        var anchor = ResolveRange(sheetName, anchorA1);
        var target = new ExcelReference(
            anchor.RowFirst,
            anchor.RowFirst + block.GetLength(0) - 1,
            anchor.ColumnFirst,
            anchor.ColumnFirst + block.GetLength(1) - 1,
            anchor.SheetId);
        WriteBlock(target, block);
    }

    /// <summary>
    /// Read once, transform in pure managed memory, write once. Round trip is exactly
    /// two boundary crossings (plus one resolve), no matter how many cells are involved.
    /// This is the canonical way to honor the one-crossing rule for in-place edits.
    /// Marshaling cost: 2 transfer crossings.
    /// Thread-safety: NOT safe for thread-safe UDFs.
    /// Accepts: a transform that maps an <c>object[,]</c> to an <c>object[,]</c> of the
    /// same shape. Returns: nothing.
    /// </summary>
    public static void RoundTripTransform(ExcelReference range, Func<object[,], object[,]> transform)
    {
        ArgumentNullException.ThrowIfNull(range);
        ArgumentNullException.ThrowIfNull(transform);
        var input = ReadBlock(range);
        var output = transform(input) ?? throw new InvalidOperationException("Transform returned null.");
        WriteBlock(range, output);
    }

    /// <summary>
    /// Convenience overload of <see cref="RoundTripTransform(ExcelReference, Func{object[,], object[,]})"/>
    /// that resolves the range from sheet and A1.
    /// Marshaling cost: 2 transfer + 1 resolve = 3 crossings, all O(1) in N.
    /// </summary>
    public static void RoundTripTransform(string sheetName, string a1Address, Func<object[,], object[,]> transform)
    {
        RoundTripTransform(ResolveRange(sheetName, a1Address), transform);
    }

    /// <summary>
    /// Reads a range and returns a <c>double[,]</c>. Non-numeric cells become
    /// <paramref name="defaultValue"/>. Logs the count of coerced cells at Verbose level.
    /// Marshaling cost: 1 read crossing.
    /// Thread-safety: NOT safe for thread-safe UDFs (reads via ExcelReference).
    /// </summary>
    public static double[,] ReadDoubleBlock(ExcelReference reference, double defaultValue = 0d)
    {
        var raw = ReadBlock(reference);
        return Marshaling.ToDoubleMatrix(raw, defaultValue);
    }

    /// <summary>
    /// Reads a range and returns a <c>string[,]</c>.
    /// Marshaling cost: 1 read crossing.
    /// Thread-safety: NOT safe for thread-safe UDFs.
    /// </summary>
    public static string[,] ReadStringBlock(ExcelReference reference)
    {
        var raw = ReadBlock(reference);
        return Marshaling.ToStringMatrix(raw);
    }

    /// <summary>
    /// Writes a <c>double[,]</c> back to Excel, boxing once on the way out.
    /// Marshaling cost: 1 write crossing.
    /// Thread-safety: NOT safe for thread-safe UDFs.
    /// </summary>
    public static void WriteDoubleBlock(ExcelReference target, double[,] block)
    {
        WriteBlock(target, Marshaling.BoxDoubleMatrix(block));
    }

    /// <summary>
    /// Excel UDF wrapper around <see cref="ReadBlock(string, string)"/>. Returns the
    /// rectangular range as an array formula result. Designed to demonstrate that even
    /// a worksheet formula returning a whole block costs one boundary crossing.
    /// Marshaling cost: 1 read crossing (resolve crossing is folded into Excel's own
    /// argument evaluation when invoked as a UDF).
    /// Thread-safety: NOT safe for MTR (touches the object model via ExcelReference).
    /// </summary>
    [ExcelFunction(
        Name = "EPT.READRANGE",
        Description = "Reads an A1 range from a named sheet as a single bulk array. One boundary crossing regardless of cell count.",
        Category = "EPT.BulkTransfer",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static object ReadRangeUdf(
        [ExcelArgument(Name = "sheet", Description = "Sheet name, unquoted.")] string sheetName,
        [ExcelArgument(Name = "a1", Description = "A1-style range, e.g. A1:D1000.")] string a1Address)
    {
        try
        {
            return ReadBlock(sheetName, a1Address);
        }
        catch (Exception ex)
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 1, "EPT.READRANGE failed for {0}!{1}: {2}", sheetName, a1Address, ex.Message);
            return ExcelError.ExcelErrorRef;
        }
    }
}
