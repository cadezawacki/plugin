using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;
using ExcelDna.Integration;

namespace ExcelPerfToolkit;

/// <summary>
/// Round 2 regular-expression utilities: match, count, extract, extract-all, and split,
/// applied across a whole block in one pass. Every function compiles the pattern once per
/// call (never a shared static cache, which would violate MTR safety) with a one-second
/// <see cref="Regex"/> match timeout as a surgical ReDoS defense - the same discipline as
/// <see cref="DeveloperUtilities.FindReplace"/>.
///
/// <para>All functions are pure CPU over <c>object[,]</c> and registered
/// <c>IsThreadSafe = true</c>. A pathological pattern that trips the per-cell match timeout
/// surfaces as <c>#VALUE!</c> in that cell rather than pinning the recalc thread. An invalid
/// pattern throws and is surfaced as <c>#VALUE!</c> by the registered unhandled handler.</para>
/// </summary>
public static class RegexUtilities
{
    private static readonly TraceSource TraceSource = new("ExcelPerfToolkit.RegexUtilities", SourceLevels.Information);

    private static readonly TimeSpan MatchTimeout = TimeSpan.FromSeconds(1);

    /// <summary>
    /// Returns a block of <c>TRUE</c>/<c>FALSE</c> indicating whether each cell's text
    /// contains a match for <paramref name="pattern"/>. Blank cells are <c>FALSE</c>;
    /// input error cells pass through unchanged.
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.REGEXMATCH", Description = "TRUE/FALSE per cell: does the text match the pattern?", Category = "EPT.Regex", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] RegexMatch(
        object[,] block,
        [ExcelArgument(Name = "pattern")] string pattern,
        [ExcelArgument(Name = "ignore_case")] bool ignoreCase = false)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rx = Build(pattern, ignoreCase, block.Length);
        return MapText(block, false, (s, r, c) =>
        {
            try
            {
                return rx.IsMatch(s);
            }
            catch (RegexMatchTimeoutException)
            {
                TraceSource.TraceEvent(TraceEventType.Warning, 1, "EPT.REGEXMATCH timeout at ({0},{1}).", r, c);
                return ExcelError.ExcelErrorValue;
            }
        });
    }

    /// <summary>
    /// Returns the number of non-overlapping matches of <paramref name="pattern"/> in each
    /// cell's text. Blank cells are 0; input error cells pass through unchanged.
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.REGEXCOUNT", Description = "Count of non-overlapping matches per cell.", Category = "EPT.Regex", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] RegexCount(
        object[,] block,
        [ExcelArgument(Name = "pattern")] string pattern,
        [ExcelArgument(Name = "ignore_case")] bool ignoreCase = false)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rx = Build(pattern, ignoreCase, block.Length);
        return MapText(block, 0d, (s, r, c) =>
        {
            try
            {
                // Regex.Count scans without materializing Match objects (~350 B/match
                // saved); it honors the instance MatchTimeout.
                return (double)rx.Count(s);
            }
            catch (RegexMatchTimeoutException)
            {
                TraceSource.TraceEvent(TraceEventType.Warning, 2, "EPT.REGEXCOUNT timeout at ({0},{1}).", r, c);
                return ExcelError.ExcelErrorValue;
            }
        });
    }

    /// <summary>
    /// Returns the first match of <paramref name="pattern"/> in each cell, or the capture
    /// group given by <paramref name="groupIndex"/> (0 = the whole match). Cells with no
    /// match - and missing groups - return an empty string. Input error cells pass through.
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.REGEXEXTRACT", Description = "First match (or capture group) per cell.", Category = "EPT.Regex", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] RegexExtract(
        object[,] block,
        [ExcelArgument(Name = "pattern")] string pattern,
        [ExcelArgument(Name = "group_index", Description = "Capture group; 0 (default) is the whole match.")] int groupIndex = 0,
        [ExcelArgument(Name = "ignore_case")] bool ignoreCase = false)
    {
        ArgumentNullException.ThrowIfNull(block);
        if (groupIndex < 0)
        {
            throw new ArgumentException("group_index must be non-negative.", nameof(groupIndex));
        }
        var rx = Build(pattern, ignoreCase, block.Length);
        return MapText(block, string.Empty, (s, r, c) =>
        {
            try
            {
                var m = rx.Match(s);
                if (!m.Success || groupIndex >= m.Groups.Count || !m.Groups[groupIndex].Success)
                {
                    return string.Empty;
                }
                return m.Groups[groupIndex].Value;
            }
            catch (RegexMatchTimeoutException)
            {
                TraceSource.TraceEvent(TraceEventType.Warning, 3, "EPT.REGEXEXTRACT timeout at ({0},{1}).", r, c);
                return ExcelError.ExcelErrorValue;
            }
        });
    }

    /// <summary>
    /// Spills every match of <paramref name="pattern"/> for each row of a single-column input
    /// across columns, padded to the widest row. Optionally returns a capture group instead
    /// of the whole match. Rows whose match trips the timeout get a single <c>#VALUE!</c>.
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.REGEXEXTRACTALL", Description = "All matches per row of a single column, spilled across columns.", Category = "EPT.Regex", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] RegexExtractAll(
        [ExcelArgument(Name = "column", Description = "Single-column input.")] object[,] column,
        [ExcelArgument(Name = "pattern")] string pattern,
        [ExcelArgument(Name = "group_index", Description = "Capture group; 0 (default) is the whole match.")] int groupIndex = 0,
        [ExcelArgument(Name = "ignore_case")] bool ignoreCase = false)
    {
        ArgumentNullException.ThrowIfNull(column);
        if (column.GetLength(1) != 1)
        {
            throw new ArgumentException("EPT.REGEXEXTRACTALL requires a single-column input.", nameof(column));
        }
        if (groupIndex < 0)
        {
            throw new ArgumentException("group_index must be non-negative.", nameof(groupIndex));
        }
        var rx = Build(pattern, ignoreCase, column.GetLength(0));
        var rows = column.GetLength(0);
        var perRow = new object[rows][];
        for (var r = 0; r < rows; r++)
        {
            var cell = column[r, 0];
            if (cell is ExcelError errCell)
            {
                perRow[r] = new object[] { errCell };
                continue;
            }
            if (Marshaling.IsBlankOrError(cell))
            {
                perRow[r] = Array.Empty<object>();
                continue;
            }
            try
            {
                var matches = rx.Matches(Marshaling.ToStringSafe(cell));
                var values = new List<object>(matches.Count);
                foreach (Match m in matches)
                {
                    if (groupIndex < m.Groups.Count && m.Groups[groupIndex].Success)
                    {
                        values.Add(m.Groups[groupIndex].Value);
                    }
                }
                perRow[r] = values.ToArray();
            }
            catch (RegexMatchTimeoutException)
            {
                TraceSource.TraceEvent(TraceEventType.Warning, 4, "EPT.REGEXEXTRACTALL timeout at row {0}.", r);
                perRow[r] = new object[] { ExcelError.ExcelErrorValue };
            }
        }
        return Spill(perRow, rows);
    }

    /// <summary>
    /// Splits each cell of a single-column input by <paramref name="pattern"/>, spilling the
    /// parts across columns padded to the widest row. Rows whose split trips the timeout get
    /// a single <c>#VALUE!</c>.
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.REGEXSPLIT", Description = "Split each cell of a single column by a regex, spilled across columns.", Category = "EPT.Regex", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] RegexSplit(
        [ExcelArgument(Name = "column", Description = "Single-column input.")] object[,] column,
        [ExcelArgument(Name = "pattern")] string pattern,
        [ExcelArgument(Name = "ignore_case")] bool ignoreCase = false)
    {
        ArgumentNullException.ThrowIfNull(column);
        if (column.GetLength(1) != 1)
        {
            throw new ArgumentException("EPT.REGEXSPLIT requires a single-column input.", nameof(column));
        }
        var rx = Build(pattern, ignoreCase, column.GetLength(0));
        var rows = column.GetLength(0);
        var perRow = new object[rows][];
        for (var r = 0; r < rows; r++)
        {
            var cell = column[r, 0];
            if (cell is ExcelError errCell)
            {
                perRow[r] = new object[] { errCell };
                continue;
            }
            if (Marshaling.IsBlankOrError(cell))
            {
                perRow[r] = Array.Empty<object>();
                continue;
            }
            try
            {
                var parts = rx.Split(Marshaling.ToStringSafe(cell));
                var values = new object[parts.Length];
                for (var i = 0; i < parts.Length; i++)
                {
                    values[i] = parts[i];
                }
                perRow[r] = values;
            }
            catch (RegexMatchTimeoutException)
            {
                TraceSource.TraceEvent(TraceEventType.Warning, 5, "EPT.REGEXSPLIT timeout at row {0}.", r);
                perRow[r] = new object[] { ExcelError.ExcelErrorValue };
            }
        }
        return Spill(perRow, rows);
    }

    // ---------- Helpers ----------

    /// <summary>
    /// Compilation costs ~15 ms once and repays roughly 5x per match; only worth paying
    /// when the block is large enough to amortize it within the single call (there is
    /// deliberately no cross-call cache to govern).
    /// </summary>
    private const int CompileThreshold = 16_384;

    private static Regex Build(string pattern, bool ignoreCase, int cellCount)
    {
        if (pattern is null)
        {
            throw new ArgumentNullException(nameof(pattern));
        }
        var options = RegexOptions.CultureInvariant;
        if (ignoreCase)
        {
            options |= RegexOptions.IgnoreCase;
        }
        if (cellCount >= CompileThreshold)
        {
            options |= RegexOptions.Compiled;
        }
        try
        {
            return new Regex(pattern, options, MatchTimeout);
        }
        catch (ArgumentException ex)
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 6, "Invalid regex pattern: {0}", ex.Message);
            throw new ArgumentException("Invalid regex pattern.", nameof(pattern), ex);
        }
    }

    private static object[,] MapText(object[,] block, object blankDefault, Func<string, int, int, object> cellFn)
    {
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var result = new object[rows, cols];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                var cell = block[r, c];
                result[r, c] = cell is ExcelError err
                    ? err
                    : Marshaling.IsBlankOrError(cell)
                        ? blankDefault
                        : cellFn(Marshaling.ToStringSafe(cell), r, c);
            }
        }
        return result;
    }

    /// <summary>Excel's grid is 16,384 columns wide; nothing wider can ever spill.</summary>
    private const int MaxSpillColumns = 16_384;

    private static object[,] Spill(object[][] perRow, int rows)
    {
        var maxCols = 1;
        foreach (var row in perRow)
        {
            if (row.Length > maxCols)
            {
                maxCols = row.Length;
            }
        }
        // A zero-width or pathological pattern (e.g. "" splits a 32,767-char cell into
        // 32,769 parts) can demand a multi-GB or >Int32.MaxValue-element block; fail
        // loudly before allocating inside the Excel process.
        if (maxCols > MaxSpillColumns)
        {
            throw new ArgumentException($"Result would be {maxCols} columns wide; Excel allows at most {MaxSpillColumns}.");
        }
        if ((long)rows * maxCols > int.MaxValue)
        {
            throw new ArgumentException($"Result {rows}x{maxCols} exceeds Int32.MaxValue cells.");
        }
        var result = new object[rows, maxCols];
        for (var r = 0; r < rows; r++)
        {
            var row = perRow[r];
            for (var c = 0; c < maxCols; c++)
            {
                result[r, c] = c < row.Length ? row[c] : string.Empty;
            }
        }
        return result;
    }
}
