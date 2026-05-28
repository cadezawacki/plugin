using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO.Hashing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using ExcelDna.Integration;

namespace ExcelPerfToolkit;

/// <summary>
/// Day-to-day developer utilities. Each function honors the one-crossing rule: it
/// receives a single bulk <c>object[,]</c> from Excel (or an explicit pre-read block),
/// computes entirely in managed memory, and returns a single bulk <c>object[,]</c>.
///
/// Bottlenecks addressed: #1 (cell-by-cell read/write) by always accepting and
/// returning whole blocks, and #2 (per-cell COM marshaling) by never re-entering
/// COM for intermediate values. Functions here are not registered as IsThreadSafe
/// because they are usually invoked at the boundary with array arguments; see
/// <see cref="ParallelUtilities"/> for the MTR-safe surface.
/// </summary>
public static class DeveloperUtilities
{
    private static readonly TraceSource TraceSource = new("ExcelPerfToolkit.DeveloperUtilities", SourceLevels.Information);

    // ---------- Trim and whitespace ----------

    /// <summary>
    /// Trims leading/trailing whitespace and collapses internal runs of whitespace to a
    /// single space across every string cell. Non-string cells are returned unchanged.
    /// Marshaling cost: 0 boundary crossings if the caller passes a block (the array
    /// argument crosses once on entry, the result crosses once on return - 2 total
    /// for the whole UDF invocation). O(1) in N.
    /// Thread-safety: pure; safe for MTR if invoked through the registration in
    /// <see cref="ParallelUtilities"/>.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.TRIMBLOCK",
        Description = "Trim and collapse whitespace across an entire range in one bulk pass.",
        Category = "EPT.DeveloperUtilities",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static object[,] TrimBlock(object[,] block)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var result = new object[rows, cols];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                var cell = block[r, c];
                if (cell is string s)
                {
                    result[r, c] = NormalizeWhitespace(s);
                }
                else
                {
                    result[r, c] = cell;
                }
            }
        }
        return result;
    }

    private static string NormalizeWhitespace(string input)
    {
        if (string.IsNullOrEmpty(input))
        {
            return input;
        }
        var sb = new StringBuilder(input.Length);
        var inWs = false;
        var started = false;
        foreach (var ch in input)
        {
            if (char.IsWhiteSpace(ch))
            {
                if (started)
                {
                    inWs = true;
                }
                continue;
            }
            if (inWs)
            {
                sb.Append(' ');
                inWs = false;
            }
            sb.Append(ch);
            started = true;
        }
        return sb.ToString();
    }

    // ---------- Coerce numeric-looking text ----------

    /// <summary>
    /// Coerces any cell whose string value parses as a number into a real
    /// <see cref="double"/>. Non-string and non-parseable cells pass through.
    /// Marshaling cost: O(1) - one bulk argument in, one bulk result out.
    /// Thread-safety: pure.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.COERCENUMERIC",
        Description = "Convert text that looks numeric into real numbers across an entire range.",
        Category = "EPT.DeveloperUtilities",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static object[,] CoerceNumeric(object[,] block)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var result = new object[rows, cols];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                var cell = block[r, c];
                if (cell is string s && Marshaling.TryToDouble(s, out var d))
                {
                    result[r, c] = d;
                }
                else
                {
                    result[r, c] = cell;
                }
            }
        }
        return result;
    }

    // ---------- Remove duplicate rows ----------

    /// <summary>
    /// Returns a block containing only the first occurrence of each unique row, in
    /// original order. Equality is determined by cell-by-cell comparison using
    /// <see cref="Marshaling.CellEquality"/>.
    /// Marshaling cost: O(1) - one bulk in, one bulk out.
    /// Thread-safety: pure.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.REMOVEDUPLICATEROWS",
        Description = "Remove duplicate rows preserving first occurrence order.",
        Category = "EPT.DeveloperUtilities",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static object[,] RemoveDuplicateRows(object[,] block)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var keep = new List<int>(rows);
        for (var r = 0; r < rows; r++)
        {
            var key = BuildRowKey(block, r, cols);
            if (seen.Add(key))
            {
                keep.Add(r);
            }
        }
        var result = new object[keep.Count, cols];
        for (var i = 0; i < keep.Count; i++)
        {
            var src = keep[i];
            for (var c = 0; c < cols; c++)
            {
                result[i, c] = block[src, c];
            }
        }
        return result;
    }

    private static string BuildRowKey(object[,] block, int row, int cols)
    {
        var sb = new StringBuilder(cols * 8);
        for (var c = 0; c < cols; c++)
        {
            sb.Append(Marshaling.ToStringSafe(block[row, c])).Append('');
        }
        return sb.ToString();
    }

    // ---------- Transpose ----------

    /// <summary>
    /// Returns the transpose of <paramref name="block"/>.
    /// Marshaling cost: O(1).
    /// Thread-safety: pure.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.TRANSPOSE",
        Description = "Transpose a rectangular block in pure managed memory.",
        Category = "EPT.DeveloperUtilities",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static object[,] Transpose(object[,] block)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var result = new object[cols, rows];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                result[c, r] = block[r, c];
            }
        }
        return result;
    }

    // ---------- Fill blanks ----------

    /// <summary>
    /// Replaces blank cells (null / <see cref="ExcelEmpty"/> / <see cref="ExcelMissing"/> /
    /// empty string) with <paramref name="fillValue"/>.
    /// Marshaling cost: O(1).
    /// Thread-safety: pure.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.FILLBLANKS",
        Description = "Replace blank cells across a block with a supplied value.",
        Category = "EPT.DeveloperUtilities",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static object[,] FillBlanks(object[,] block, object fillValue)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var result = new object[rows, cols];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                var v = block[r, c];
                result[r, c] = IsCellBlank(v) ? fillValue : v;
            }
        }
        return result;
    }

    /// <summary>
    /// Replaces blanks in each column with the most recent non-blank value above. Useful
    /// for cleaning hierarchical data that uses merged-looking blanks.
    /// Marshaling cost: O(1).
    /// Thread-safety: pure.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.FILLDOWN",
        Description = "Fill blanks with the value above, per column.",
        Category = "EPT.DeveloperUtilities",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static object[,] FillDown(object[,] block)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var result = new object[rows, cols];
        for (var c = 0; c < cols; c++)
        {
            object? last = ExcelEmpty.Value;
            for (var r = 0; r < rows; r++)
            {
                var v = block[r, c];
                if (IsCellBlank(v))
                {
                    result[r, c] = last ?? ExcelEmpty.Value;
                }
                else
                {
                    result[r, c] = v;
                    last = v;
                }
            }
        }
        return result;
    }

    private static bool IsCellBlank(object? v)
    {
        return v is null
            || v is ExcelEmpty
            || v is ExcelMissing
            || (v is string s && s.Length == 0);
    }

    // ---------- Split / Join ----------

    /// <summary>
    /// Splits each value in a single-column block by <paramref name="delimiter"/> into
    /// multiple columns. The result is padded to the width of the longest split.
    /// Marshaling cost: O(1).
    /// Thread-safety: pure.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.SPLITCOLUMN",
        Description = "Split a single column into many columns by delimiter.",
        Category = "EPT.DeveloperUtilities",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static object[,] SplitColumn(object[,] column, string delimiter)
    {
        ArgumentNullException.ThrowIfNull(column);
        if (string.IsNullOrEmpty(delimiter))
        {
            throw new ArgumentException("Delimiter must be non-empty.", nameof(delimiter));
        }
        if (column.GetLength(1) != 1)
        {
            throw new ArgumentException("SplitColumn requires a single-column block.", nameof(column));
        }
        var rows = column.GetLength(0);
        var parts = new string[rows][];
        var maxCols = 0;
        for (var r = 0; r < rows; r++)
        {
            var s = Marshaling.ToStringSafe(column[r, 0]);
            parts[r] = s.Split(delimiter, StringSplitOptions.None);
            if (parts[r].Length > maxCols)
            {
                maxCols = parts[r].Length;
            }
        }
        if (maxCols == 0)
        {
            maxCols = 1;
        }
        var result = new object[rows, maxCols];
        for (var r = 0; r < rows; r++)
        {
            var row = parts[r];
            for (var c = 0; c < maxCols; c++)
            {
                result[r, c] = c < row.Length ? row[c] : string.Empty;
            }
        }
        return result;
    }

    /// <summary>
    /// Joins every column of each row into a single string with <paramref name="delimiter"/>.
    /// Marshaling cost: O(1).
    /// Thread-safety: pure.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.JOINCOLUMNS",
        Description = "Join all columns of each row into one column using a delimiter.",
        Category = "EPT.DeveloperUtilities",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static object[,] JoinColumns(object[,] block, string delimiter)
    {
        ArgumentNullException.ThrowIfNull(block);
        delimiter ??= string.Empty;
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var result = new object[rows, 1];
        var sb = new StringBuilder();
        for (var r = 0; r < rows; r++)
        {
            sb.Clear();
            for (var c = 0; c < cols; c++)
            {
                if (c > 0)
                {
                    sb.Append(delimiter);
                }
                sb.Append(Marshaling.ToStringSafe(block[r, c]));
            }
            result[r, 0] = sb.ToString();
        }
        return result;
    }

    // ---------- Find and replace ----------

    /// <summary>
    /// Finds and replaces text across every string cell. When <paramref name="useRegex"/>
    /// is true, <paramref name="find"/> is treated as a .NET regex and replacements honor
    /// regex substitutions.
    /// Marshaling cost: O(1).
    /// Thread-safety: pure. Regex compilation is cached locally to the call to avoid a
    /// shared static cache, which would violate MTR safety guidance.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.FINDREPLACE",
        Description = "Find and replace text across a block, with optional regex.",
        Category = "EPT.DeveloperUtilities",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static object[,] FindReplace(
        object[,] block,
        string find,
        string replace,
        [ExcelArgument(Name = "useRegex", Description = "TRUE to treat 'find' as a .NET regex.")] bool useRegex = false)
    {
        ArgumentNullException.ThrowIfNull(block);
        find ??= string.Empty;
        replace ??= string.Empty;
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var result = new object[rows, cols];

        Regex? rx = null;
        if (useRegex)
        {
            try
            {
                // MatchTimeout is the surgical ReDoS defense: any pathological pattern
                // (e.g. (a+)+$) will throw RegexMatchTimeoutException instead of pinning
                // the calc thread forever. The 1s budget is generous for any real Excel
                // workflow and short enough to keep the workbook responsive.
                rx = new Regex(find, RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
            }
            catch (ArgumentException ex)
            {
                TraceSource.TraceEvent(TraceEventType.Warning, 2, "EPT.FINDREPLACE invalid regex: {0}", ex.Message);
                throw new ArgumentException("Invalid regex pattern.", nameof(find), ex);
            }
        }

        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                var cell = block[r, c];
                if (cell is string s)
                {
                    try
                    {
                        result[r, c] = rx is null
                            ? s.Replace(find, replace, StringComparison.Ordinal)
                            : rx.Replace(s, replace);
                    }
                    catch (RegexMatchTimeoutException)
                    {
                        // Pathological pattern tripped the per-match budget. Surface the
                        // failure as a cell error rather than pinning the calc thread.
                        TraceSource.TraceEvent(TraceEventType.Warning, 3, "EPT.FINDREPLACE regex timeout on cell ({0},{1}).", r, c);
                        result[r, c] = ExcelError.ExcelErrorValue;
                    }
                }
                else
                {
                    result[r, c] = cell;
                }
            }
        }
        return result;
    }

    // ---------- Unpivot / stack ----------

    /// <summary>
    /// Stacks all columns of <paramref name="block"/> into a single column, preserving
    /// row-major order.
    /// Marshaling cost: O(1).
    /// Thread-safety: pure.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.STACKCOLUMNS",
        Description = "Stack every column of a block into a single column.",
        Category = "EPT.DeveloperUtilities",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static object[,] StackColumns(object[,] block)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        // Excel's max range is ~16B cells which overflows int multiplication.
        // Compute the product as long and reject before allocation rather than wrapping.
        var total = (long)rows * cols;
        if (total > int.MaxValue)
        {
            throw new ArgumentException($"Block too large: {rows}x{cols} = {total} cells exceeds Int32.MaxValue.", nameof(block));
        }
        var result = new object[(int)total, 1];
        var idx = 0;
        for (var c = 0; c < cols; c++)
        {
            for (var r = 0; r < rows; r++)
            {
                result[idx++, 0] = block[r, c];
            }
        }
        return result;
    }

    /// <summary>
    /// Unpivots a wide block into a long block (Key, Attribute, Value). The first
    /// <paramref name="keyColumns"/> columns are repeated for every value column.
    /// Marshaling cost: O(1).
    /// Thread-safety: pure.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.UNPIVOT",
        Description = "Reshape a wide block to long format with key columns and Attribute/Value columns.",
        Category = "EPT.DeveloperUtilities",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static object[,] Unpivot(object[,] block, int keyColumns)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        if (rows < 1 || keyColumns < 0 || keyColumns >= cols)
        {
            throw new ArgumentException("Unpivot requires at least one header row and 0 <= keyColumns < cols.");
        }

        var dataRows = rows - 1;
        var valueColumns = cols - keyColumns;
        // Guard against int overflow: dataRows * valueColumns can exceed Int32.MaxValue
        // for worksheet-sized blocks. Reject up front rather than wrapping.
        var outRowsLong = (long)dataRows * valueColumns;
        if (outRowsLong > int.MaxValue)
        {
            throw new ArgumentException($"Unpivoted shape {outRowsLong} rows exceeds Int32.MaxValue.", nameof(block));
        }
        var outRows = (int)outRowsLong;
        var outCols = keyColumns + 2; // + Attribute + Value
        var result = new object[outRows, outCols];

        var outIdx = 0;
        for (var r = 1; r < rows; r++)
        {
            for (var c = keyColumns; c < cols; c++)
            {
                for (var k = 0; k < keyColumns; k++)
                {
                    result[outIdx, k] = block[r, k];
                }
                result[outIdx, keyColumns] = block[0, c];
                result[outIdx, keyColumns + 1] = block[r, c];
                outIdx++;
            }
        }
        return result;
    }

    // ---------- Unique values ----------

    /// <summary>
    /// Returns a single-column block of distinct non-blank values found in
    /// <paramref name="block"/>, in first-seen order.
    /// Marshaling cost: O(1).
    /// Thread-safety: pure.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.UNIQUE",
        Description = "List unique non-blank values in first-seen order.",
        Category = "EPT.DeveloperUtilities",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static object[,] Unique(object[,] block)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var values = new List<object>(Math.Min(rows * cols, 1024));
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                var v = block[r, c];
                if (IsCellBlank(v))
                {
                    continue;
                }
                var key = Marshaling.ToStringSafe(v);
                if (seen.Add(key))
                {
                    values.Add(v);
                }
            }
        }
        if (values.Count == 0)
        {
            return new object[1, 1] { { ExcelEmpty.Value } };
        }
        var result = new object[values.Count, 1];
        for (var i = 0; i < values.Count; i++)
        {
            result[i, 0] = values[i];
        }
        return result;
    }

    /// <summary>
    /// Returns a single-column block of (value, count) pairs for every distinct non-blank
    /// value, in first-seen order. Result has two columns.
    /// Marshaling cost: O(1).
    /// Thread-safety: pure.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.UNIQUECOUNT",
        Description = "Distinct values with their occurrence counts.",
        Category = "EPT.DeveloperUtilities",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static object[,] UniqueCount(object[,] block)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var order = new List<string>(64);
        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        var samples = new Dictionary<string, object>(StringComparer.Ordinal);
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                var v = block[r, c];
                if (IsCellBlank(v))
                {
                    continue;
                }
                var key = Marshaling.ToStringSafe(v);
                if (counts.TryGetValue(key, out var existing))
                {
                    counts[key] = existing + 1;
                }
                else
                {
                    counts[key] = 1;
                    order.Add(key);
                    samples[key] = v;
                }
            }
        }
        if (order.Count == 0)
        {
            return new object[1, 2] { { ExcelEmpty.Value, 0d } };
        }
        var result = new object[order.Count, 2];
        for (var i = 0; i < order.Count; i++)
        {
            result[i, 0] = samples[order[i]];
            result[i, 1] = (double)counts[order[i]];
        }
        return result;
    }

    // ---------- Block lookup (left join in managed memory) ----------

    /// <summary>
    /// Joins a left-hand block to a lookup table in managed memory. Every value in
    /// <paramref name="leftKeyColumnIndex"/> of <paramref name="left"/> is looked up in
    /// column 0 of <paramref name="lookup"/>; matching rows of
    /// <paramref name="lookupReturnColumns"/> are returned alongside the original left
    /// row. Misses fill with <see cref="ExcelError.ExcelErrorNA"/>.
    /// Marshaling cost: O(1) - one bulk in, one bulk out. Replaces N classic VLOOKUP
    /// crossings with a single managed-memory hash join.
    /// Thread-safety: pure.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.BLOCKLOOKUP",
        Description = "Hash-join a key column against a lookup table entirely in managed memory.",
        Category = "EPT.DeveloperUtilities",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static object[,] BlockLookup(
        object[,] left,
        int leftKeyColumnIndex,
        object[,] lookup,
        object[,] lookupReturnColumns)
    {
        ArgumentNullException.ThrowIfNull(left);
        ArgumentNullException.ThrowIfNull(lookup);
        ArgumentNullException.ThrowIfNull(lookupReturnColumns);

        var leftRows = left.GetLength(0);
        var leftCols = left.GetLength(1);
        if ((uint)leftKeyColumnIndex >= (uint)leftCols)
        {
            throw new ArgumentOutOfRangeException(nameof(leftKeyColumnIndex));
        }

        var lookupRows = lookup.GetLength(0);
        var lookupCols = lookup.GetLength(1);

        var returnColIndexes = FlattenIntIndexes(lookupReturnColumns);
        foreach (var idx in returnColIndexes)
        {
            if ((uint)idx >= (uint)lookupCols)
            {
                throw new ArgumentOutOfRangeException(nameof(lookupReturnColumns), $"Return column index {idx} out of range.");
            }
        }

        var index = new Dictionary<string, int>(lookupRows, StringComparer.Ordinal);
        for (var r = 0; r < lookupRows; r++)
        {
            var key = Marshaling.ToStringSafe(lookup[r, 0]);
            if (!index.ContainsKey(key))
            {
                index[key] = r;
            }
        }

        var totalCols = leftCols + returnColIndexes.Length;
        var result = new object[leftRows, totalCols];
        for (var r = 0; r < leftRows; r++)
        {
            for (var c = 0; c < leftCols; c++)
            {
                result[r, c] = left[r, c];
            }
            var lkey = Marshaling.ToStringSafe(left[r, leftKeyColumnIndex]);
            if (index.TryGetValue(lkey, out var matchRow))
            {
                for (var i = 0; i < returnColIndexes.Length; i++)
                {
                    result[r, leftCols + i] = lookup[matchRow, returnColIndexes[i]];
                }
            }
            else
            {
                for (var i = 0; i < returnColIndexes.Length; i++)
                {
                    result[r, leftCols + i] = ExcelError.ExcelErrorNA;
                }
            }
        }
        return result;
    }

    private static int[] FlattenIntIndexes(object[,] block)
    {
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var list = new List<int>(rows * cols);
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                var v = block[r, c];
                if (Marshaling.IsBlankOrError(v))
                {
                    continue;
                }
                if (!Marshaling.TryToDouble(v, out var d))
                {
                    throw new ArgumentException("Return column indexes must be numbers.");
                }
                // Bounds-check before the cast so an out-of-range value is reported
                // verbatim instead of silently wrapping to a negative int.
                if (d < 0d || d > int.MaxValue || d != Math.Truncate(d))
                {
                    throw new ArgumentException($"Column index {d} is not a non-negative integer fitting Int32.");
                }
                list.Add((int)d);
            }
        }
        if (list.Count == 0)
        {
            throw new ArgumentException("At least one return column index is required.");
        }
        return list.ToArray();
    }

    // ---------- Sort ----------

    /// <summary>
    /// Returns <paramref name="block"/> sorted by one or more key columns. Sort is stable.
    /// Numeric comparisons are used when both cells parse as numbers, otherwise ordinal
    /// string comparison.
    /// Marshaling cost: O(1).
    /// Thread-safety: pure.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.SORTBLOCK",
        Description = "Sort rows by one or more key columns with optional descending flags.",
        Category = "EPT.DeveloperUtilities",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static object[,] SortBlock(
        object[,] block,
        object[,] keyColumns,
        [ExcelArgument(Name = "descendingFlags", Description = "Optional same-length column of TRUE/FALSE per key.")] object descendingFlags)
    {
        ArgumentNullException.ThrowIfNull(block);
        ArgumentNullException.ThrowIfNull(keyColumns);

        var keys = FlattenIntIndexes(keyColumns);
        // Excel marshals an array argument as object[,]; a single-cell argument like
        // TRUE comes through as a bare bool/double/string. Accept both forms - the
        // scalar case applies the same flag to every key column.
        bool[] desc;
        if (descendingFlags is object[,] flagsBlock)
        {
            var fRows = flagsBlock.GetLength(0);
            var fCols = flagsBlock.GetLength(1);
            var flat = new List<bool>(fRows * fCols);
            for (var r = 0; r < fRows; r++)
            {
                for (var c = 0; c < fCols; c++)
                {
                    var v = flagsBlock[r, c];
                    if (Marshaling.IsBlankOrError(v))
                    {
                        continue;
                    }
                    flat.Add(CoerceBool(v));
                }
            }
            desc = flat.Count >= keys.Length ? flat.Take(keys.Length).ToArray() : flat.Concat(Enumerable.Repeat(false, keys.Length - flat.Count)).ToArray();
        }
        else if (Marshaling.IsBlankOrError(descendingFlags))
        {
            desc = new bool[keys.Length];
        }
        else
        {
            // Scalar TRUE/FALSE/1/0/"TRUE" applies to every key.
            var b = CoerceBool(descendingFlags);
            desc = Enumerable.Repeat(b, keys.Length).ToArray();
        }

        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var order = Enumerable.Range(0, rows).ToArray();
        Array.Sort(order, (a, b) => CompareRows(block, a, b, keys, desc));

        var result = new object[rows, cols];
        for (var i = 0; i < rows; i++)
        {
            var src = order[i];
            for (var c = 0; c < cols; c++)
            {
                result[i, c] = block[src, c];
            }
        }
        return result;
    }

    private static bool CoerceBool(object? value) => value switch
    {
        bool b => b,
        double d => d != 0d,
        int i => i != 0,
        long l => l != 0L,
        string s => s.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || s == "1",
        _ => false,
    };

    private static int CompareRows(object[,] block, int a, int b, int[] keys, bool[] desc)
    {
        for (var k = 0; k < keys.Length; k++)
        {
            var col = keys[k];
            var va = block[a, col];
            var vb = block[b, col];
            int cmp;
            if (Marshaling.TryToDouble(va, out var da) && Marshaling.TryToDouble(vb, out var db))
            {
                cmp = da.CompareTo(db);
            }
            else
            {
                cmp = string.CompareOrdinal(Marshaling.ToStringSafe(va), Marshaling.ToStringSafe(vb));
            }
            if (cmp != 0)
            {
                return desc[k] ? -cmp : cmp;
            }
        }
        return a.CompareTo(b);
    }

    // ---------- Checksum / hash ----------

    /// <summary>
    /// Returns a stable hash (XXH3 64-bit) over the canonical text representation of
    /// <paramref name="block"/>. Useful for cheap change detection over a range. The
    /// algorithm is deterministic across processes and bitnesses.
    /// Marshaling cost: O(1).
    /// Thread-safety: pure.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.HASHBLOCK",
        Description = "Stable XXH3 hash over a block, suitable for change detection.",
        Category = "EPT.DeveloperUtilities",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static string HashBlock(object[,] block)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var hasher = new XxHash3();
        Span<byte> separator = stackalloc byte[1];
        separator[0] = 0x1f;
        Span<byte> rowSeparator = stackalloc byte[1];
        rowSeparator[0] = 0x1e;
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                var text = Marshaling.ToStringSafe(block[r, c]);
                var bytes = Encoding.UTF8.GetBytes(text);
                hasher.Append(bytes);
                hasher.Append(separator);
            }
            hasher.Append(rowSeparator);
        }
        var hashBytes = hasher.GetHashAndReset();
        return Convert.ToHexString(hashBytes);
    }

    /// <summary>
    /// Returns a SHA-256 hash over the canonical text representation of
    /// <paramref name="block"/> for cryptographically strong change detection or
    /// auditing. Slower than <see cref="HashBlock"/> but collision-resistant.
    /// Marshaling cost: O(1).
    /// Thread-safety: pure.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.SHA256BLOCK",
        Description = "SHA-256 hex digest over a block.",
        Category = "EPT.DeveloperUtilities",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static string Sha256Block(object[,] block)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        using var sha = SHA256.Create();
        // Allocate the separator byte arrays ONCE outside the hot loop. The previous
        // implementation called Span.ToArray() per cell, producing up to N*cells throwaway
        // arrays for a worksheet-size block.
        var separator = new byte[] { 0x1f };
        var rowSeparator = new byte[] { 0x1e };
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                var text = Marshaling.ToStringSafe(block[r, c]);
                var bytes = Encoding.UTF8.GetBytes(text);
                sha.TransformBlock(bytes, 0, bytes.Length, null, 0);
                sha.TransformBlock(separator, 0, separator.Length, null, 0);
            }
            sha.TransformBlock(rowSeparator, 0, rowSeparator.Length, null, 0);
        }
        sha.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        return Convert.ToHexString(sha.Hash ?? Array.Empty<byte>());
    }
}
