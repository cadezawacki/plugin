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
    /// Below this many probes the O(M) index build / O(M log M) sort costs more than it
    /// saves; a direct scan is cheaper and allocation-free.
    /// </summary>
    private const int LinearScanProbeLimit = 4;

    /// <summary>
    /// Resolves every cell of <paramref name="lookupValues"/> against
    /// <paramref name="lookupArray"/> in one pass and returns the aligned value from
    /// <paramref name="returnArray"/>. The result has the same shape as
    /// <paramref name="lookupValues"/>.
    ///
    /// <para><paramref name="lookupArray"/> and <paramref name="returnArray"/> must hold the
    /// same number of cells (row-major aligned). Exact match (the default) compares
    /// type-aware: text matches text case-insensitively, numbers match numbers (5 and 5.0
    /// agree), and a number never matches its text form (numeric 5 does not match "5"),
    /// exactly as native XLOOKUP. An error in <paramref name="lookupValues"/> propagates to
    /// the corresponding result cell. Approximate match operates over the numeric keys; if
    /// they are not already sorted ascending they are sorted internally (stably, so the
    /// last original occurrence of a duplicate key wins) and the value for the largest key
    /// &lt;= the lookup value is returned (Excel's "next smaller" semantics).</para>
    ///
    /// <para><paramref name="matchMode"/>: omitted or <c>TRUE</c> or <c>0</c> = exact;
    /// <c>FALSE</c> or <c>-1</c> (XLOOKUP's next-smaller mode) = approximate. Other numeric
    /// modes are rejected with <c>#VALUE!</c>. Misses return <paramref name="ifNotFound"/>
    /// when supplied, otherwise <c>#N/A</c>.</para>
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
        [ExcelArgument(Name = "match_mode", Description = "Optional. TRUE/0/omitted = exact; FALSE/-1 = approximate (next-smaller).")] object matchMode)
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
            var approximate = ResolveMatchMode(matchMode);

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

    private static bool ResolveMatchMode(object matchMode)
    {
        // Documented legacy form: TRUE/omitted = exact, FALSE = approximate.
        if (matchMode is bool mode)
        {
            return !mode;
        }
        if (matchMode is null or ExcelMissing or ExcelEmpty)
        {
            return false;
        }
        if (matchMode is string s && bool.TryParse(s, out var parsed))
        {
            return !parsed;
        }
        if (Marshaling.TryToDouble(matchMode, out var mm))
        {
            // XLOOKUP-style numeric modes: 0 = exact, -1 = next-smaller (approximate).
            // 1 (next-larger) and 2 (wildcard) are not implemented; reject them loudly
            // rather than silently running exact match.
            if (mm == 0d)
            {
                return false;
            }
            if (mm == -1d)
            {
                return true;
            }
        }
        throw new ArgumentException("match_mode must be TRUE/0 (exact) or FALSE/-1 (approximate next-smaller).");
    }

    private static object[,] ExactLookup(object[,] lookupBlock, object?[] keys, object?[] returns, object notFound)
    {
        var rows = lookupBlock.GetLength(0);
        var cols = lookupBlock.GetLength(1);
        var result = new object[rows, cols];

        // A handful of probes (the formula-copied-down shape): scan instead of paying
        // the full index build.
        if ((long)rows * cols <= LinearScanProbeLimit)
        {
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    var cell = lookupBlock[r, c];
                    if (cell is ExcelError probeErr)
                    {
                        result[r, c] = probeErr;
                        continue;
                    }
                    var probe = CellKey.From(cell);
                    var found = -1;
                    for (var i = 0; i < keys.Length; i++)
                    {
                        if (probe.Equals(CellKey.From(keys[i])))
                        {
                            found = i;
                            break;
                        }
                    }
                    result[r, c] = found >= 0 ? returns[found] ?? ExcelEmpty.Value : notFound;
                }
            }
            return result;
        }

        // Build the typed key -> first-occurrence index once. Type-aware comparison
        // matches XLOOKUP: text is case-insensitive, numbers compare as doubles (5 and
        // 5.0 agree because both are the double 5.0), and types never cross.
        var index = new Dictionary<CellKey, int>(keys.Length);
        for (var i = 0; i < keys.Length; i++)
        {
            index.TryAdd(CellKey.From(keys[i]), i);
        }

        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                var cell = lookupBlock[r, c];
                if (cell is ExcelError probeErr)
                {
                    result[r, c] = probeErr;
                    continue;
                }
                result[r, c] = index.TryGetValue(CellKey.From(cell), out var idx)
                    ? returns[idx] ?? ExcelEmpty.Value
                    : notFound;
            }
        }
        return result;
    }

    private static object[,] ApproximateLookup(object[,] lookupBlock, object?[] keys, object?[] returns, object notFound)
    {
        var rows = lookupBlock.GetLength(0);
        var cols = lookupBlock.GetLength(1);
        var result = new object[rows, cols];

        // Collect the numeric keys with their original return index.
        var points = new List<(double Key, int Index)>(keys.Length);
        for (var i = 0; i < keys.Length; i++)
        {
            if (IsNumeric(keys[i], out var d))
            {
                points.Add((d, i));
            }
        }

        if ((long)rows * cols <= LinearScanProbeLimit)
        {
            // O(M) floor scan per probe - cheaper than sorting for a handful of probes.
            // `>=` on equal keys makes the last original occurrence win, matching the
            // sorted path's tiebreak below.
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    var cell = lookupBlock[r, c];
                    if (cell is ExcelError probeErr)
                    {
                        result[r, c] = probeErr;
                        continue;
                    }
                    if (!IsNumeric(cell, out var target))
                    {
                        result[r, c] = notFound;
                        continue;
                    }
                    var best = -1;
                    var bestKey = double.NegativeInfinity;
                    for (var p = 0; p < points.Count; p++)
                    {
                        var (key, _) = points[p];
                        if (key <= target && (best < 0 || key >= bestKey))
                        {
                            best = p;
                            bestKey = key;
                        }
                    }
                    result[r, c] = best < 0
                        ? notFound
                        : returns[points[best].Index] ?? ExcelEmpty.Value;
                }
            }
            return result;
        }

        // The dominant real-world input is already sorted ascending (the documented
        // contract): an O(M) check avoids an O(M log M) sort. When sorting is needed,
        // the index tiebreak makes duplicate keys deterministic (last occurrence wins).
        var sorted = true;
        for (var i = 1; i < points.Count; i++)
        {
            if (points[i - 1].Key > points[i].Key)
            {
                sorted = false;
                break;
            }
        }
        if (!sorted)
        {
            points.Sort(static (a, b) =>
            {
                var byKey = a.Key.CompareTo(b.Key);
                return byKey != 0 ? byKey : a.Index.CompareTo(b.Index);
            });
        }

        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                var cell = lookupBlock[r, c];
                if (cell is ExcelError probeErr)
                {
                    result[r, c] = probeErr;
                    continue;
                }
                if (!IsNumeric(cell, out var target))
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

    /// <summary>
    /// A typed cell identity for exact-match lookups. Numbers, text, bools, errors, and
    /// blanks live in disjoint keyspaces: numeric 5 never matches text "5". Text compares
    /// case-insensitively (XLOOKUP's text semantics); numbers compare by value, so 5 and
    /// 5.0 agree. No string is materialized for numeric keys (the previous string
    /// canonicalization allocated ~63 B per numeric cell on both the build and probe
    /// sides).
    /// </summary>
    private readonly struct CellKey : IEquatable<CellKey>
    {
        private enum Kind : byte
        {
            Blank,
            Number,
            Text,
            Bool,
            Error,
        }

        private readonly Kind _kind;
        private readonly double _number;
        private readonly string? _text;

        private CellKey(Kind kind, double number, string? text)
        {
            _kind = kind;
            _number = number;
            _text = text;
        }

        public static CellKey From(object? cell) => cell switch
        {
            null or ExcelEmpty or ExcelMissing => new CellKey(Kind.Blank, 0d, null),
            bool b => new CellKey(Kind.Bool, b ? 1d : 0d, null),
            ExcelError err => new CellKey(Kind.Error, (int)err, null),
            string s => new CellKey(Kind.Text, 0d, s),
            _ => Marshaling.TryToDouble(cell, out var d)
                ? new CellKey(Kind.Number, d, null)
                : new CellKey(Kind.Text, 0d, Marshaling.ToStringSafe(cell)),
        };

        public bool Equals(CellKey other)
        {
            if (_kind != other._kind)
            {
                return false;
            }
            return _kind == Kind.Text
                ? string.Equals(_text, other._text, StringComparison.OrdinalIgnoreCase)
                : _number.Equals(other._number);
        }

        public override bool Equals(object? obj) => obj is CellKey other && Equals(other);

        public override int GetHashCode()
            => _kind == Kind.Text
                ? HashCode.Combine(_kind, _text!.GetHashCode(StringComparison.OrdinalIgnoreCase))
                : HashCode.Combine(_kind, _number);
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
