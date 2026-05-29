using System;
using System.Collections.Generic;
using System.Diagnostics;
using ExcelDna.Integration;

namespace ExcelPerfToolkit;

/// <summary>
/// Boosted lookups. This is the toolkit's answer to "can we make VLOOKUP/XLOOKUP faster?".
///
/// <para><b>The problem.</b> A column of <i>R</i> <c>VLOOKUP</c>/<c>XLOOKUP</c> formulas
/// against a table of <i>M</i> rows re-scans the table for every formula: O(R*M). On a
/// 100k x 100k join that is ten billion comparisons, recomputed on every edit.</para>
///
/// <para><b>The fix.</b> <see cref="XLookupB"/> resolves an entire column of lookup values
/// in a single call. It builds the lookup table's index <b>once</b> - a hash map for exact
/// match, a sorted array for approximate match - then probes it for every lookup value:
/// O(M + R) total, with the table read crossing the Excel boundary exactly once (the
/// one-crossing rule of bottlenecks #1/#2). The function is registered
/// <c>IsThreadSafe = true</c>, so Excel's multithreaded recalc engine (bottleneck #3) can
/// schedule it freely.</para>
///
/// <para>Use <see cref="XLookupB"/> when you want results aligned one-to-one with a column
/// of lookup keys; use <see cref="DeveloperUtilities.BlockLookup"/> when you want to staple
/// several lookup columns onto an existing block (a left join).</para>
/// </summary>
public static class LookupBoost
{
    private static readonly TraceSource TraceSource = new("ExcelPerfToolkit.LookupBoost", SourceLevels.Information);

    /// <summary>
    /// Resolves every cell of <paramref name="lookupValues"/> against
    /// <paramref name="lookupArray"/> in one pass and returns the aligned value from
    /// <paramref name="returnArray"/>. The result has the same shape as
    /// <paramref name="lookupValues"/>.
    ///
    /// <para><paramref name="lookupArray"/> and <paramref name="returnArray"/> must hold the
    /// same number of cells (row-major aligned). Exact match (the default) is
    /// case-insensitive and uses a hash index built once. Approximate match
    /// (<paramref name="matchMode"/> = <c>FALSE</c>) operates over the numeric keys, assumes
    /// they are sorted ascending, and returns the value for the largest key
    /// &lt;= the lookup value (Excel's "next smaller" semantics).</para>
    ///
    /// <para>Misses return <paramref name="ifNotFound"/> when supplied, otherwise
    /// <c>#N/A</c>.</para>
    ///
    /// Marshaling cost: one read per range argument; one bulk write. O(M + R) compute.
    /// Thread-safety: SAFE for MTR (pure CPU; no object-model access).
    /// </summary>
    [ExcelFunction(
        Name = "EPT.XLOOKUPB",
        Description = "Batched exact/approximate lookup: resolves a whole column of keys in one O(M+R) pass.",
        Category = "EPT.Lookup",
        IsThreadSafe = true,
        IsVolatile = false)]
    public static object XLookupB(
        [ExcelArgument(Name = "lookup_values", Description = "One or more keys to resolve. Result mirrors this shape.")] object lookupValues,
        [ExcelArgument(Name = "lookup_array", Description = "The column of keys to search.")] object lookupArray,
        [ExcelArgument(Name = "return_array", Description = "The aligned column of values to return.")] object returnArray,
        [ExcelArgument(Name = "if_not_found", Description = "Optional value returned for misses. Defaults to #N/A.")] object ifNotFound,
        [ExcelArgument(Name = "match_mode", Description = "Optional. TRUE/omitted = exact; FALSE = approximate (sorted-ascending, next-smaller).")] object matchMode)
    {
        try
        {
            var lookupBlock = Marshaling.AsArray2D(lookupValues);
            var keys = Flatten(Marshaling.AsArray2D(lookupArray));
            var returns = Flatten(Marshaling.AsArray2D(returnArray));
            if (keys.Length != returns.Length)
            {
                return ExcelError.ExcelErrorValue;
            }

            var notFound = IsSupplied(ifNotFound) ? ifNotFound : ExcelError.ExcelErrorNA;
            var approximate = matchMode is bool mode && !mode;

            return approximate
                ? ApproximateLookup(lookupBlock, keys, returns, notFound)
                : ExactLookup(lookupBlock, keys, returns, notFound);
        }
        catch (Exception ex) when (!IsCritical(ex))
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 1, "EPT.XLOOKUPB failed: {0}", ex.Message);
            return ExcelError.ExcelErrorValue;
        }
    }

    private static object[,] ExactLookup(object[,] lookupBlock, object?[] keys, object?[] returns, object notFound)
    {
        // Build the key -> first-occurrence index once. Case-insensitive to match XLOOKUP's
        // text comparison; numeric keys canonicalize through ToStringSafe so 5 and 5.0 agree.
        var index = new Dictionary<string, int>(keys.Length, StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < keys.Length; i++)
        {
            var key = Marshaling.ToStringSafe(keys[i]);
            if (!index.ContainsKey(key))
            {
                index[key] = i;
            }
        }

        var rows = lookupBlock.GetLength(0);
        var cols = lookupBlock.GetLength(1);
        var result = new object[rows, cols];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                var key = Marshaling.ToStringSafe(lookupBlock[r, c]);
                result[r, c] = index.TryGetValue(key, out var idx)
                    ? returns[idx] ?? ExcelEmpty.Value
                    : notFound;
            }
        }
        return result;
    }

    private static object[,] ApproximateLookup(object[,] lookupBlock, object?[] keys, object?[] returns, object notFound)
    {
        // Collect the numeric keys with their original return index, then sort ascending so
        // a binary search can find the largest key <= the lookup value in O(log M).
        var points = new List<(double Key, int Index)>(keys.Length);
        for (var i = 0; i < keys.Length; i++)
        {
            if (IsNumeric(keys[i], out var d))
            {
                points.Add((d, i));
            }
        }
        points.Sort(static (a, b) => a.Key.CompareTo(b.Key));

        var rows = lookupBlock.GetLength(0);
        var cols = lookupBlock.GetLength(1);
        var result = new object[rows, cols];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                if (!IsNumeric(lookupBlock[r, c], out var target))
                {
                    result[r, c] = notFound;
                    continue;
                }
                var found = FloorIndex(points, target);
                result[r, c] = found < 0
                    ? notFound
                    : returns[points[found].Index] ?? ExcelEmpty.Value;
            }
        }
        return result;
    }

    /// <summary>Index of the largest entry whose key is &lt;= <paramref name="target"/>, or -1.</summary>
    private static int FloorIndex(List<(double Key, int Index)> points, double target)
    {
        int lo = 0, hi = points.Count - 1, result = -1;
        while (lo <= hi)
        {
            var mid = lo + ((hi - lo) / 2);
            if (points[mid].Key <= target)
            {
                result = mid;
                lo = mid + 1;
            }
            else
            {
                hi = mid - 1;
            }
        }
        return result;
    }

    private static object?[] Flatten(object[,] block)
    {
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var total = (long)rows * cols;
        if (total > int.MaxValue)
        {
            throw new ArgumentException($"Block too large: {rows}x{cols} = {total} cells exceeds Int32.MaxValue.");
        }
        var result = new object?[(int)total];
        var idx = 0;
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                result[idx++] = block[r, c];
            }
        }
        return result;
    }

    private static bool IsSupplied(object? value)
        => value is not null and not ExcelMissing;

    private static bool IsNumeric(object? v, out double d)
    {
        if (v is string)
        {
            d = 0d;
            return false;
        }
        return Marshaling.TryToDouble(v, out d);
    }

    private static bool IsCritical(Exception ex)
        => ex is OutOfMemoryException
            or StackOverflowException
            or AccessViolationException
            or System.Threading.ThreadAbortException;
}
