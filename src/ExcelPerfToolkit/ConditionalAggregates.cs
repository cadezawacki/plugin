using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelDna.Integration;

namespace ExcelPerfToolkit;

/// <summary>
/// Lightning-fast conditional aggregation: a family of <c>*IFS</c> functions, weighted
/// statistics, conditional <c>SUMPRODUCT</c>, and a single-pass <c>GROUP BY</c>. This is
/// the toolkit's answer to "what is Excel missing natively, and how do we make the things
/// it does have faster?".
///
/// <para><b>Why this is fast.</b> Excel's native <c>SUMIFS</c>/<c>COUNTIFS</c> re-scan
/// every criteria range on every formula instance. A column of <i>R</i> SUMIFS formulas
/// over a criteria range of <i>M</i> cells is O(R*M). Each function here takes the value
/// range and criteria ranges as one bulk <c>object[,]</c> each (the one-crossing rule of
/// bottlenecks #1/#2), builds the match mask in a single managed-memory pass, and is
/// registered <c>IsThreadSafe = true</c> so Excel's multithreaded recalc engine
/// (bottleneck #3) can fan the whole column out across cores.</para>
///
/// <para><b>What Excel lacks that lives here.</b> <c>MEDIANIFS</c>, <c>PERCENTILEIFS</c>,
/// <c>MODEIFS</c>, <c>STDEVIFS</c>/<c>VARIFS</c> (sample and population),
/// <c>GEOMEANIFS</c>/<c>HARMEANIFS</c>, <c>PRODUCTIFS</c>, <c>DISTINCTCOUNTIFS</c>,
/// <c>FIRSTIFS</c>/<c>LASTIFS</c>, weighted average / weighted median / weighted stdev
/// (with and without conditions), a conditional <c>SUMPRODUCT</c>, and a one-pass
/// <c>GROUP BY</c> that returns distinct keys with an aggregate.</para>
///
/// <para><b>Criteria semantics.</b> Criteria strings follow Excel's own grammar:
/// comparison prefixes (<c>"&gt;5"</c>, <c>"&lt;=10"</c>, <c>"&lt;&gt;x"</c>, <c>"=v"</c>),
/// case-insensitive text equality with <c>*</c>/<c>?</c> wildcards (<c>~</c> escapes a
/// literal wildcard), the empty-operand blank/non-blank forms (<c>""</c> and <c>"&lt;&gt;"</c>),
/// and bare numeric criteria. Numeric comparisons only match genuinely numeric cells, not
/// numeric-looking text - exactly as Excel does.</para>
///
/// <para><b>Thread-safety.</b> Every UDF here is pure CPU over its <c>object[,]</c>
/// arguments and registered <c>IsThreadSafe = true</c>: no shared mutable state, no access
/// to the Excel object model, MTR-eligible.</para>
/// </summary>
public static class ConditionalAggregates
{
    private static readonly TraceSource TraceSource = new("ExcelPerfToolkit.ConditionalAggregates", SourceLevels.Information);

    // ====================================================================
    // Conditional aggregates over a value range (SUMIFS family)
    // ====================================================================

    /// <summary>
    /// Counts the cells across the criteria ranges that satisfy every criterion. Faster,
    /// MTR-eligible equivalent of Excel's native <c>COUNTIFS</c>.
    /// Marshaling cost: one read per range argument; one scalar write. O(1) in cell count.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.COUNTIFS",
        Description = "Fast, MTR-safe COUNTIFS: count cells matching every (range, criteria) pair.",
        Category = "EPT.Conditional",
        IsThreadSafe = true,
        IsVolatile = false)]
    public static object CountIfs(
        [ExcelArgument(Name = "criteria_range1")] object criteriaRange1,
        [ExcelArgument(Name = "criteria1")] object criteria1,
        [ExcelArgument(Name = "more", Description = "Additional criteria_range, criteria pairs.")] params object[] morePairs)
        => Guard("EPT.COUNTIFS", () =>
        {
            var pairs = Combine(criteriaRange1, criteria1, morePairs);
            var block = Marshaling.AsArray2D(criteriaRange1);
            var mask = BuildMask(block.GetLength(0), block.GetLength(1), pairs);
            var count = 0;
            foreach (var m in mask)
            {
                if (m)
                {
                    count++;
                }
            }
            return (double)count;
        });

    /// <summary>
    /// Sums the numeric cells of <paramref name="sumRange"/> whose row matches every
    /// criterion. Faster, MTR-eligible equivalent of Excel's native <c>SUMIFS</c>.
    /// Non-numeric cells in the sum range are ignored, as in Excel. Returns 0 when nothing
    /// matches.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.SUMIFS",
        Description = "Fast, MTR-safe SUMIFS over a numeric value range.",
        Category = "EPT.Conditional",
        IsThreadSafe = true,
        IsVolatile = false)]
    public static object SumIfs(
        [ExcelArgument(Name = "sum_range")] object sumRange,
        [ExcelArgument(Name = "criteria_range1")] object criteriaRange1,
        [ExcelArgument(Name = "criteria1")] object criteria1,
        [ExcelArgument(Name = "more")] params object[] morePairs)
        => Guard("EPT.SUMIFS", () => AggregateIfs("sum", sumRange, Combine(criteriaRange1, criteria1, morePairs)));

    /// <summary>Conditional arithmetic mean. <c>#DIV/0!</c> when nothing matches. Like Excel <c>AVERAGEIFS</c>.</summary>
    [ExcelFunction(Name = "EPT.AVERAGEIFS", Description = "Fast, MTR-safe AVERAGEIFS.", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object AverageIfs(object averageRange, object criteriaRange1, object criteria1, params object[] morePairs)
        => Guard("EPT.AVERAGEIFS", () => AggregateIfs("average", averageRange, Combine(criteriaRange1, criteria1, morePairs)));

    /// <summary>Conditional minimum. Returns 0 when nothing matches, like Excel <c>MINIFS</c>.</summary>
    [ExcelFunction(Name = "EPT.MINIFS", Description = "Fast, MTR-safe MINIFS.", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object MinIfs(object minRange, object criteriaRange1, object criteria1, params object[] morePairs)
        => Guard("EPT.MINIFS", () => AggregateIfs("min", minRange, Combine(criteriaRange1, criteria1, morePairs)));

    /// <summary>Conditional maximum. Returns 0 when nothing matches, like Excel <c>MAXIFS</c>.</summary>
    [ExcelFunction(Name = "EPT.MAXIFS", Description = "Fast, MTR-safe MAXIFS.", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object MaxIfs(object maxRange, object criteriaRange1, object criteria1, params object[] morePairs)
        => Guard("EPT.MAXIFS", () => AggregateIfs("max", maxRange, Combine(criteriaRange1, criteria1, morePairs)));

    /// <summary>Conditional median - <b>not available natively</b> in Excel. <c>#NUM!</c> when nothing matches.</summary>
    [ExcelFunction(Name = "EPT.MEDIANIFS", Description = "Conditional median (no native Excel equivalent).", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object MedianIfs(object medianRange, object criteriaRange1, object criteria1, params object[] morePairs)
        => Guard("EPT.MEDIANIFS", () => AggregateIfs("median", medianRange, Combine(criteriaRange1, criteria1, morePairs)));

    /// <summary>
    /// Conditional percentile (inclusive, matching <c>PERCENTILE.INC</c>) - <b>no native
    /// Excel equivalent</b>. <paramref name="k"/> is in [0, 1]. <c>#NUM!</c> when nothing
    /// matches or <paramref name="k"/> is out of range.
    /// </summary>
    [ExcelFunction(Name = "EPT.PERCENTILEIFS", Description = "Conditional inclusive percentile (no native Excel equivalent).", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object PercentileIfs(
        [ExcelArgument(Name = "value_range")] object valueRange,
        [ExcelArgument(Name = "k", Description = "Percentile in [0,1].")] double k,
        object criteriaRange1,
        object criteria1,
        params object[] morePairs)
        => Guard("EPT.PERCENTILEIFS", () =>
        {
            var values = Marshaling.AsArray2D(valueRange);
            var rows = values.GetLength(0);
            var cols = values.GetLength(1);
            var mask = BuildMask(rows, cols, Combine(criteriaRange1, criteria1, morePairs));
            var nums = new List<double>();
            var i = 0;
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++, i++)
                {
                    if (!mask[i])
                    {
                        continue;
                    }
                    var cell = values[r, c];
                    if (cell is ExcelError err)
                    {
                        // Native Excel propagates an error from a matched value cell.
                        return err;
                    }
                    if (IsNumericCell(cell, out var d))
                    {
                        nums.Add(d);
                    }
                }
            }
            return PercentileInc(nums, k);
        });

    /// <summary>Conditional sample standard deviation (n-1) - <b>no native Excel equivalent</b>. <c>#DIV/0!</c> with fewer than 2 matches.</summary>
    [ExcelFunction(Name = "EPT.STDEVIFS", Description = "Conditional sample standard deviation (no native Excel equivalent).", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object StdevIfs(object valueRange, object criteriaRange1, object criteria1, params object[] morePairs)
        => Guard("EPT.STDEVIFS", () => AggregateIfs("stdev", valueRange, Combine(criteriaRange1, criteria1, morePairs)));

    /// <summary>Conditional population standard deviation (n) - <b>no native Excel equivalent</b>.</summary>
    [ExcelFunction(Name = "EPT.STDEVPIFS", Description = "Conditional population standard deviation (no native Excel equivalent).", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object StdevPIfs(object valueRange, object criteriaRange1, object criteria1, params object[] morePairs)
        => Guard("EPT.STDEVPIFS", () => AggregateIfs("stdevp", valueRange, Combine(criteriaRange1, criteria1, morePairs)));

    /// <summary>Conditional sample variance (n-1) - <b>no native Excel equivalent</b>.</summary>
    [ExcelFunction(Name = "EPT.VARIFS", Description = "Conditional sample variance (no native Excel equivalent).", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object VarIfs(object valueRange, object criteriaRange1, object criteria1, params object[] morePairs)
        => Guard("EPT.VARIFS", () => AggregateIfs("var", valueRange, Combine(criteriaRange1, criteria1, morePairs)));

    /// <summary>Conditional population variance (n) - <b>no native Excel equivalent</b>.</summary>
    [ExcelFunction(Name = "EPT.VARPIFS", Description = "Conditional population variance (no native Excel equivalent).", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object VarPIfs(object valueRange, object criteriaRange1, object criteria1, params object[] morePairs)
        => Guard("EPT.VARPIFS", () => AggregateIfs("varp", valueRange, Combine(criteriaRange1, criteria1, morePairs)));

    /// <summary>Conditional mode (most frequent value) - <b>no native Excel equivalent</b>. <c>#N/A</c> when no value repeats.</summary>
    [ExcelFunction(Name = "EPT.MODEIFS", Description = "Conditional mode (no native Excel equivalent).", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object ModeIfs(object valueRange, object criteriaRange1, object criteria1, params object[] morePairs)
        => Guard("EPT.MODEIFS", () => AggregateIfs("mode", valueRange, Combine(criteriaRange1, criteria1, morePairs)));

    /// <summary>Conditional geometric mean - <b>no native Excel equivalent</b>. <c>#NUM!</c> if any matched value is non-positive.</summary>
    [ExcelFunction(Name = "EPT.GEOMEANIFS", Description = "Conditional geometric mean (no native Excel equivalent).", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object GeoMeanIfs(object valueRange, object criteriaRange1, object criteria1, params object[] morePairs)
        => Guard("EPT.GEOMEANIFS", () => AggregateIfs("geomean", valueRange, Combine(criteriaRange1, criteria1, morePairs)));

    /// <summary>Conditional harmonic mean - <b>no native Excel equivalent</b>. <c>#NUM!</c> if any matched value is non-positive.</summary>
    [ExcelFunction(Name = "EPT.HARMEANIFS", Description = "Conditional harmonic mean (no native Excel equivalent).", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object HarMeanIfs(object valueRange, object criteriaRange1, object criteria1, params object[] morePairs)
        => Guard("EPT.HARMEANIFS", () => AggregateIfs("harmean", valueRange, Combine(criteriaRange1, criteria1, morePairs)));

    /// <summary>Conditional product of matched numeric cells - <b>no native Excel equivalent</b>. Returns 0 when nothing matches.</summary>
    [ExcelFunction(Name = "EPT.PRODUCTIFS", Description = "Conditional product (no native Excel equivalent).", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object ProductIfs(object valueRange, object criteriaRange1, object criteria1, params object[] morePairs)
        => Guard("EPT.PRODUCTIFS", () => AggregateIfs("product", valueRange, Combine(criteriaRange1, criteria1, morePairs)));

    /// <summary>Conditional count of distinct values among matched cells - <b>no native Excel equivalent</b>.</summary>
    [ExcelFunction(Name = "EPT.DISTINCTCOUNTIFS", Description = "Conditional distinct count (no native Excel equivalent).", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object DistinctCountIfs(object valueRange, object criteriaRange1, object criteria1, params object[] morePairs)
        => Guard("EPT.DISTINCTCOUNTIFS", () => AggregateIfs("distinct", valueRange, Combine(criteriaRange1, criteria1, morePairs)));

    /// <summary>First matched value in row-major order - <b>no native Excel equivalent</b>. <c>#N/A</c> when nothing matches.</summary>
    [ExcelFunction(Name = "EPT.FIRSTIFS", Description = "First matched value (no native Excel equivalent).", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object FirstIfs(object valueRange, object criteriaRange1, object criteria1, params object[] morePairs)
        => Guard("EPT.FIRSTIFS", () => AggregateIfs("first", valueRange, Combine(criteriaRange1, criteria1, morePairs)));

    /// <summary>Last matched value in row-major order - <b>no native Excel equivalent</b>. <c>#N/A</c> when nothing matches.</summary>
    [ExcelFunction(Name = "EPT.LASTIFS", Description = "Last matched value (no native Excel equivalent).", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object LastIfs(object valueRange, object criteriaRange1, object criteria1, params object[] morePairs)
        => Guard("EPT.LASTIFS", () => AggregateIfs("last", valueRange, Combine(criteriaRange1, criteria1, morePairs)));

    // ====================================================================
    // Weighted statistics (no native Excel equivalents)
    // ====================================================================

    /// <summary>
    /// Weighted arithmetic mean <c>sum(value*weight) / sum(weight)</c> - <b>no native Excel
    /// equivalent</b> (the usual workaround is <c>SUMPRODUCT(v,w)/SUM(w)</c>). Rows where
    /// either the value or the weight is non-numeric are skipped. <c>#DIV/0!</c> when the
    /// total weight is zero. <c>#VALUE!</c> when the ranges differ in size.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.WAVG", Description = "Weighted average (no native Excel equivalent).", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object WAvg(
        [ExcelArgument(Name = "values")] object valueRange,
        [ExcelArgument(Name = "weights")] object weightRange)
        => Guard("EPT.WAVG", () => WeightedAverage(valueRange, weightRange, null));

    /// <summary>
    /// Weighted average restricted to rows matching every criterion - <b>no native Excel
    /// equivalent</b>. Same skip/zero-weight rules as <see cref="WAvg"/>.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.WAVGIFS", Description = "Conditional weighted average (no native Excel equivalent).", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object WAvgIfs(
        [ExcelArgument(Name = "values")] object valueRange,
        [ExcelArgument(Name = "weights")] object weightRange,
        object criteriaRange1,
        object criteria1,
        params object[] morePairs)
        => Guard("EPT.WAVGIFS", () => WeightedAverage(valueRange, weightRange, Combine(criteriaRange1, criteria1, morePairs)));

    /// <summary>
    /// Weighted median - <b>no native Excel equivalent</b>. Pairs are sorted by value and the
    /// returned value is the one at which the cumulative weight first reaches half of the
    /// total weight (the lower weighted median). Only positive weights contribute.
    /// <c>#DIV/0!</c> when the total weight is zero; <c>#VALUE!</c> on a size mismatch.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.WMEDIAN", Description = "Weighted median (no native Excel equivalent).", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object WMedian(object valueRange, object weightRange)
        => Guard("EPT.WMEDIAN", () =>
        {
            var v = FlattenRowMajor(Marshaling.AsArray2D(valueRange));
            var w = FlattenRowMajor(Marshaling.AsArray2D(weightRange));
            if (v.Length != w.Length)
            {
                return ExcelError.ExcelErrorValue;
            }
            var pairs = new List<(double Value, double Weight)>(v.Length);
            var total = 0d;
            for (var i = 0; i < v.Length; i++)
            {
                if (v[i] is ExcelError ve)
                {
                    return ve;
                }
                if (w[i] is ExcelError we)
                {
                    return we;
                }
                if (IsNumericCell(v[i], out var x) && IsNumericCell(w[i], out var wt) && wt > 0d)
                {
                    pairs.Add((x, wt));
                    total += wt;
                }
            }
            if (pairs.Count == 0 || total <= 0d)
            {
                return ExcelError.ExcelErrorDiv0;
            }
            pairs.Sort(static (a, b) => a.Value.CompareTo(b.Value));
            var half = total / 2d;
            var cumulative = 0d;
            foreach (var (value, weight) in pairs)
            {
                cumulative += weight;
                if (cumulative >= half)
                {
                    return value;
                }
            }
            return pairs[pairs.Count - 1].Value;
        });

    /// <summary>
    /// Weighted (population) standard deviation - <b>no native Excel equivalent</b>:
    /// <c>sqrt(sum(w*(x-wmean)^2) / sum(w))</c>. Only positive weights contribute.
    /// <c>#DIV/0!</c> when the total weight is zero; <c>#VALUE!</c> on a size mismatch.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.WSTDEV", Description = "Weighted population standard deviation (no native Excel equivalent).", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object WStdev(object valueRange, object weightRange)
        => Guard("EPT.WSTDEV", () =>
        {
            var v = FlattenRowMajor(Marshaling.AsArray2D(valueRange));
            var w = FlattenRowMajor(Marshaling.AsArray2D(weightRange));
            if (v.Length != w.Length)
            {
                return ExcelError.ExcelErrorValue;
            }
            // Collect the typed pairs once: the second pass must not re-run the 10-branch
            // TryToDouble dispatch (4 conversions per row where 2 suffice).
            var kept = new List<(double X, double W)>();
            double sumW = 0d, sumWx = 0d;
            for (var i = 0; i < v.Length; i++)
            {
                if (v[i] is ExcelError ve)
                {
                    return ve;
                }
                if (w[i] is ExcelError we)
                {
                    return we;
                }
                if (IsNumericCell(v[i], out var x) && IsNumericCell(w[i], out var wt) && wt > 0d)
                {
                    kept.Add((x, wt));
                    sumW += wt;
                    sumWx += wt * x;
                }
            }
            if (sumW <= 0d)
            {
                return ExcelError.ExcelErrorDiv0;
            }
            var mean = sumWx / sumW;
            var sumWss = 0d;
            foreach (var (x, wt) in kept)
            {
                var d = x - mean;
                sumWss += wt * d * d;
            }
            return Math.Sqrt(sumWss / sumW);
        });

    // ====================================================================
    // Conditional SUMPRODUCT
    // ====================================================================

    /// <summary>
    /// Conditional <c>SUMPRODUCT</c>: <c>sum(a*b)</c> over rows matching every criterion -
    /// the clean, fast form of the awkward <c>SUMPRODUCT(a*b*(cond1)*(cond2))</c> idiom.
    /// Rows where either factor is non-numeric are skipped. Returns 0 when nothing matches;
    /// <c>#VALUE!</c> when the two value ranges differ in size.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.SUMPRODUCTIFS", Description = "Conditional SUMPRODUCT of two ranges (no clean native equivalent).", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object SumProductIfs(
        [ExcelArgument(Name = "range_a")] object rangeA,
        [ExcelArgument(Name = "range_b")] object rangeB,
        object criteriaRange1,
        object criteria1,
        params object[] morePairs)
        => Guard("EPT.SUMPRODUCTIFS", () =>
        {
            var aBlock = Marshaling.AsArray2D(rangeA);
            var a = FlattenRowMajor(aBlock);
            var b = FlattenRowMajor(Marshaling.AsArray2D(rangeB));
            if (a.Length != b.Length)
            {
                return ExcelError.ExcelErrorValue;
            }
            var mask = BuildMask(aBlock.GetLength(0), aBlock.GetLength(1), Combine(criteriaRange1, criteria1, morePairs));
            var sum = 0d;
            for (var i = 0; i < a.Length; i++)
            {
                if (!mask[i])
                {
                    continue;
                }
                if (a[i] is ExcelError ea)
                {
                    return ea;
                }
                if (b[i] is ExcelError eb)
                {
                    return eb;
                }
                if (IsNumericCell(a[i], out var x) && IsNumericCell(b[i], out var y))
                {
                    sum += x * y;
                }
            }
            return sum;
        });

    // ====================================================================
    // GROUP BY (single pass, returns a spilled key/aggregate table)
    // ====================================================================

    /// <summary>
    /// One-pass <c>GROUP BY</c>: returns the distinct keys of <paramref name="keyRange"/>
    /// (in first-seen order) with one aggregate of <paramref name="valueRange"/> per group.
    /// <paramref name="keyRange"/> may span several columns to form a composite key; the
    /// first column of <paramref name="valueRange"/> supplies the values. The result is a
    /// spilled block of <c>keyColumns + 1</c> columns.
    ///
    /// <para><paramref name="operation"/> is one of: <c>count</c>, <c>sum</c>,
    /// <c>average</c>, <c>min</c>, <c>max</c>, <c>median</c>, <c>stdev</c>, <c>stdevp</c>,
    /// <c>var</c>, <c>varp</c>, <c>product</c>, <c>mode</c>, <c>geomean</c>, <c>harmean</c>,
    /// <c>distinct</c>, <c>first</c>, <c>last</c>.</para>
    ///
    /// Marshaling cost: one read per range argument; one bulk write. O(1) in cell count.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.GROUPBY", Description = "Single-pass GROUP BY returning distinct keys with one aggregate per group.", Category = "EPT.Conditional", IsThreadSafe = true, IsVolatile = false)]
    public static object GroupBy(
        [ExcelArgument(Name = "key_range", Description = "One or more key columns; composite key when more than one.")] object keyRange,
        [ExcelArgument(Name = "value_range", Description = "Values to aggregate (first column is used).")] object valueRange,
        [ExcelArgument(Name = "operation", Description = "sum|count|average|min|max|median|stdev|stdevp|var|varp|product|mode|geomean|harmean|distinct|first|last")] string operation)
        => Guard("EPT.GROUPBY", () =>
        {
            var keys = Marshaling.AsArray2D(keyRange);
            var values = Marshaling.AsArray2D(valueRange);
            var keyRows = keys.GetLength(0);
            var keyCols = keys.GetLength(1);
            if (keyRows == 0 || keyCols == 0 || values.GetLength(0) != keyRows)
            {
                return ExcelError.ExcelErrorValue;
            }

            // Normalize the operation once, not once per group in the result loop.
            var op = (operation ?? string.Empty).Trim().ToLowerInvariant();

            // Text keys group case-insensitively (Excel's text equality), matching the
            // "distinct" aggregate's policy; the displayed key keeps first-seen casing.
            var order = new List<string>();
            var groups = new Dictionary<string, List<object?>>(StringComparer.OrdinalIgnoreCase);
            var keyCells = new Dictionary<string, object?[]>(StringComparer.OrdinalIgnoreCase);
            var sb = new StringBuilder();
            for (var r = 0; r < keyRows; r++)
            {
                sb.Clear();
                for (var c = 0; c < keyCols; c++)
                {
                    AppendCellKey(sb, keys[r, c]);
                }
                var key = sb.ToString();
                if (!groups.TryGetValue(key, out var list))
                {
                    list = new List<object?>();
                    groups[key] = list;
                    // The display row is only retained for first-seen keys; allocating it
                    // per input row would be rows-many wasted arrays.
                    var keyRow = new object?[keyCols];
                    for (var c = 0; c < keyCols; c++)
                    {
                        keyRow[c] = keys[r, c];
                    }
                    keyCells[key] = keyRow;
                    order.Add(key);
                }
                list.Add(values[r, 0]);
            }

            var result = new object[order.Count, keyCols + 1];
            for (var i = 0; i < order.Count; i++)
            {
                var key = order[i];
                var cells = keyCells[key];
                for (var c = 0; c < keyCols; c++)
                {
                    result[i, c] = cells[c] ?? ExcelEmpty.Value;
                }
                result[i, keyCols] = Aggregate(op, groups[key]);
            }
            return result;
        });

    // ====================================================================
    // Vectorized CASE WHEN + conditional text join
    // ====================================================================

    /// <summary>Excel's hard limit on the number of characters a single cell can hold.</summary>
    private const int MaxCellChars = 32_767;

    /// <summary>
    /// Vectorized <c>CASE WHEN</c> over a whole block - the one-call replacement for nested
    /// <c>IF</c>/<c>IFS</c>/<c>SWITCH</c> chains. Each cell of <paramref name="testRange"/>
    /// is evaluated against the (criteria, result) pairs in order and receives the result of
    /// the FIRST matching criterion. Criteria use the same Excel grammar as the *IFS family
    /// (<c>"&gt;=5"</c>, <c>"&lt;&gt;x"</c>, wildcards, blank forms, bare values) and each is
    /// parsed ONCE per call - not once per cell, as a column of nested IFs effectively does.
    /// Each result may be a scalar (broadcast) or a block shaped like the test range
    /// (per-cell). A trailing unpaired argument is the default for unmatched cells; without
    /// one they return <c>#N/A</c>, as native <c>SWITCH</c> does.
    /// Marshaling cost: one read per range argument; one bulk write.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.CASEWHEN",
        Description = "Vectorized CASE WHEN: first matching (criteria, result) pair per cell; trailing unpaired value is the default.",
        Category = "EPT.Conditional",
        IsThreadSafe = true,
        IsVolatile = false)]
    public static object CaseWhen(
        [ExcelArgument(Name = "test_range", Description = "Values to test. Result mirrors this shape.")] object testRange,
        [ExcelArgument(Name = "criteria1", Description = "First criterion (Excel criteria grammar).")] object criteria1,
        [ExcelArgument(Name = "result1", Description = "Result when criteria1 matches; scalar or same-shaped block.")] object result1,
        [ExcelArgument(Name = "more", Description = "Additional criteria, result pairs; a trailing unpaired value is the default.")] params object[] more)
        => Guard("EPT.CASEWHEN", () =>
        {
            var block = Marshaling.AsArray2D(testRange);
            var rows = block.GetLength(0);
            var cols = block.GetLength(1);

            var pairCount = 1 + (more.Length / 2);
            var criteria = new Criterion[pairCount];
            var results = new Func<int, int, object>[pairCount];
            criteria[0] = Criterion.Parse(criteria1);
            results[0] = ResultProvider(result1, rows, cols);
            for (var p = 1; p < pairCount; p++)
            {
                criteria[p] = Criterion.Parse(more[2 * (p - 1)]);
                results[p] = ResultProvider(more[(2 * (p - 1)) + 1], rows, cols);
            }
            Func<int, int, object> fallback = more.Length % 2 == 1
                ? ResultProvider(more[^1], rows, cols)
                : static (_, _) => ExcelError.ExcelErrorNA;

            var result = new object[rows, cols];
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    var cell = block[r, c];
                    var matched = false;
                    for (var p = 0; p < pairCount; p++)
                    {
                        if (criteria[p].Matches(cell))
                        {
                            result[r, c] = results[p](r, c);
                            matched = true;
                            break;
                        }
                    }
                    if (!matched)
                    {
                        result[r, c] = fallback(r, c);
                    }
                }
            }
            return result;
        });

    /// <summary>
    /// Joins the text of the cells matching every criterion - the conditional
    /// <c>TEXTJOIN</c> Excel lacks natively (the workaround is a <c>TEXTJOIN(IF(...))</c>
    /// array formula that re-scans per cell). Blank matched cells are skipped; an error in
    /// a matched cell propagates, as native <c>TEXTJOIN</c> does; a result beyond the
    /// 32,767-character cell limit returns <c>#VALUE!</c>.
    /// Marshaling cost: one read per range argument; one scalar write.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.TEXTJOINIFS",
        Description = "Join the text of cells matching every (range, criteria) pair (no native equivalent).",
        Category = "EPT.Conditional",
        IsThreadSafe = true,
        IsVolatile = false)]
    public static object TextJoinIfs(
        [ExcelArgument(Name = "text_range", Description = "Cells whose text is joined.")] object textRange,
        [ExcelArgument(Name = "delimiter")] string delimiter,
        [ExcelArgument(Name = "criteria_range1")] object criteriaRange1,
        [ExcelArgument(Name = "criteria1")] object criteria1,
        [ExcelArgument(Name = "more")] params object[] morePairs)
        => Guard("EPT.TEXTJOINIFS", () =>
        {
            var values = Marshaling.AsArray2D(textRange);
            var rows = values.GetLength(0);
            var cols = values.GetLength(1);
            var mask = BuildMask(rows, cols, Combine(criteriaRange1, criteria1, morePairs));
            delimiter ??= string.Empty;
            var sb = new StringBuilder();
            var first = true;
            var i = 0;
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++, i++)
                {
                    if (!mask[i])
                    {
                        continue;
                    }
                    var cell = values[r, c];
                    if (cell is ExcelError err)
                    {
                        return err;
                    }
                    if (IsBlank(cell))
                    {
                        continue;
                    }
                    if (!first)
                    {
                        sb.Append(delimiter);
                    }
                    sb.Append(Marshaling.ToStringSafe(cell));
                    first = false;
                    if (sb.Length > MaxCellChars)
                    {
                        return ExcelError.ExcelErrorValue;
                    }
                }
            }
            return sb.ToString();
        });

    /// <summary>
    /// Adapts a CASE WHEN result argument into a per-cell accessor: a 1x1 block or scalar
    /// broadcasts; a block matching the test range's shape supplies per-cell values; any
    /// other shape is rejected.
    /// </summary>
    private static Func<int, int, object> ResultProvider(object arg, int rows, int cols)
    {
        if (arg is object[,] b)
        {
            if (b.Length == 1)
            {
                var single = b[0, 0] ?? ExcelEmpty.Value;
                return (_, _) => single;
            }
            if (b.GetLength(0) == rows && b.GetLength(1) == cols)
            {
                return (r, c) => b[r, c] ?? ExcelEmpty.Value;
            }
            throw new ArgumentException($"A result block must be a scalar or match the test range's {rows}x{cols} shape.");
        }
        var value = arg is null or ExcelMissing ? ExcelEmpty.Value : arg;
        return (_, _) => value;
    }

    // ====================================================================
    // Shared aggregation engine
    // ====================================================================

    /// <summary>
    /// Applies a named aggregation to a list of matched cell values. The operation name
    /// must already be trimmed and lower-cased (callers normalize once per call). Numeric
    /// aggregations extract the genuinely-numeric cells (numeric-looking text and logicals
    /// are ignored, as in Excel) and propagate an <see cref="ExcelError"/> found in a
    /// matched cell, as native SUMIFS/AVERAGEIFS do. Returns a boxed <see cref="double"/>,
    /// an original cell value (<c>first</c>/<c>last</c>), or an <see cref="ExcelError"/>
    /// sentinel for empty / undefined cases.
    /// </summary>
    private static object Aggregate(string op, IReadOnlyList<object?> cells)
    {
        switch (op)
        {
            case "count":
                return (double)cells.Count;
            case "distinct":
            case "distinctcount":
            {
                var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var keyBuilder = new StringBuilder();
                foreach (var c in cells)
                {
                    if (IsBlank(c) || c is ExcelError)
                    {
                        continue;
                    }
                    keyBuilder.Clear();
                    AppendCellKey(keyBuilder, c);
                    set.Add(keyBuilder.ToString());
                }
                return (double)set.Count;
            }
            case "first":
                foreach (var c in cells)
                {
                    if (!IsBlank(c) && c is not ExcelError)
                    {
                        return c!;
                    }
                }
                return ExcelError.ExcelErrorNA;
            case "last":
                for (var i = cells.Count - 1; i >= 0; i--)
                {
                    var c = cells[i];
                    if (!IsBlank(c) && c is not ExcelError)
                    {
                        return c!;
                    }
                }
                return ExcelError.ExcelErrorNA;
        }

        var nums = new List<double>(cells.Count);
        foreach (var c in cells)
        {
            // Native Excel propagates an error from a matched value cell for the numeric
            // aggregates (SUMIFS over a matched #DIV/0! returns #DIV/0!); count/first/
            // last/distinct above define their own skip semantics.
            if (c is ExcelError err)
            {
                return err;
            }
            if (IsNumericCell(c, out var d))
            {
                nums.Add(d);
            }
        }

        switch (op)
        {
            case "sum":
                return Sum(nums);
            case "product":
            {
                if (nums.Count == 0)
                {
                    return 0d;
                }
                var p = 1d;
                foreach (var v in nums)
                {
                    p *= v;
                }
                return p;
            }
            case "average":
            case "mean":
                return nums.Count == 0 ? ExcelError.ExcelErrorDiv0 : Sum(nums) / nums.Count;
            case "min":
            {
                if (nums.Count == 0)
                {
                    return 0d;
                }
                var m = nums[0];
                foreach (var v in nums)
                {
                    if (v < m)
                    {
                        m = v;
                    }
                }
                return m;
            }
            case "max":
            {
                if (nums.Count == 0)
                {
                    return 0d;
                }
                var m = nums[0];
                foreach (var v in nums)
                {
                    if (v > m)
                    {
                        m = v;
                    }
                }
                return m;
            }
            case "median":
                return nums.Count == 0 ? ExcelError.ExcelErrorNum : Median(nums);
            case "stdev":
                return nums.Count < 2 ? ExcelError.ExcelErrorDiv0 : Math.Sqrt(Variance(nums, sample: true));
            case "stdevp":
                return nums.Count < 1 ? ExcelError.ExcelErrorDiv0 : Math.Sqrt(Variance(nums, sample: false));
            case "var":
                return nums.Count < 2 ? ExcelError.ExcelErrorDiv0 : Variance(nums, sample: true);
            case "varp":
                return nums.Count < 1 ? ExcelError.ExcelErrorDiv0 : Variance(nums, sample: false);
            case "mode":
                return Mode(nums);
            case "geomean":
                return GeoMean(nums);
            case "harmean":
                return HarMean(nums);
            default:
                throw new ArgumentException($"Unknown aggregation '{op}'. See EPT.GROUPBY help for the supported list.");
        }
    }

    /// <summary>Collects the matched value cells, then defers to <see cref="Aggregate"/>.</summary>
    private static object AggregateIfs(string op, object valueRange, object[] pairs)
    {
        var values = Marshaling.AsArray2D(valueRange);
        var rows = values.GetLength(0);
        var cols = values.GetLength(1);
        var mask = BuildMask(rows, cols, pairs);
        var matched = new List<object?>();
        var i = 0;
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++, i++)
            {
                if (mask[i])
                {
                    matched.Add(values[r, c]);
                }
            }
        }
        return Aggregate(op, matched);
    }

    private static object WeightedAverage(object valueRange, object weightRange, object[]? pairs)
    {
        var vBlock = Marshaling.AsArray2D(valueRange);
        var v = FlattenRowMajor(vBlock);
        var w = FlattenRowMajor(Marshaling.AsArray2D(weightRange));
        if (v.Length != w.Length)
        {
            return ExcelError.ExcelErrorValue;
        }
        var mask = pairs is null ? null : BuildMask(vBlock.GetLength(0), vBlock.GetLength(1), pairs);
        double sumW = 0d, sumWx = 0d;
        for (var i = 0; i < v.Length; i++)
        {
            if (mask is not null && !mask[i])
            {
                continue;
            }
            if (v[i] is ExcelError ve)
            {
                return ve;
            }
            if (w[i] is ExcelError we)
            {
                return we;
            }
            if (IsNumericCell(v[i], out var x) && IsNumericCell(w[i], out var wt))
            {
                sumW += wt;
                sumWx += wt * x;
            }
        }
        return sumW == 0d ? ExcelError.ExcelErrorDiv0 : sumWx / sumW;
    }

    // ====================================================================
    // Match mask + criteria parsing
    // ====================================================================

    /// <summary>
    /// Builds the per-element AND of every criterion, iterating the criteria blocks in
    /// place (no flattened copies). Each pair is a (criteria range, criteria scalar); the
    /// range must have exactly the value range's shape, as native COUNTIFS/SUMIFS require.
    /// </summary>
    private static bool[] BuildMask(int rows, int cols, object[] pairs)
    {
        if (pairs.Length % 2 != 0)
        {
            throw new ArgumentException("Criteria must be supplied as (range, criteria) pairs.");
        }
        var total = (long)rows * cols;
        if (total > int.MaxValue)
        {
            throw new ArgumentException($"Block too large: {rows}x{cols} = {total} cells exceeds Int32.MaxValue.");
        }
        var n = (int)total;
        var mask = new bool[n];
        Array.Fill(mask, true);
        for (var p = 0; p < pairs.Length; p += 2)
        {
            var range = Marshaling.AsArray2D(pairs[p]);
            if (range.GetLength(0) != rows || range.GetLength(1) != cols)
            {
                throw new ArgumentException(
                    $"Criteria range is {range.GetLength(0)}x{range.GetLength(1)} but the value range is {rows}x{cols}; shapes must match.");
            }
            var criterion = Criterion.Parse(pairs[p + 1]);
            var i = 0;
            if (criterion.IsPlainNumeric)
            {
                // Fast path for the dominant shape: a plain numeric criterion over a
                // numeric column. `is double` + an inlined compare skips the layered
                // blank/error/type dispatch; anything that isn't a raw double falls back
                // to the general matcher so semantics stay identical.
                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < cols; c++, i++)
                    {
                        if (!mask[i])
                        {
                            continue;
                        }
                        var cell = range[r, c];
                        mask[i] = cell is double d ? criterion.MatchesNumber(d) : criterion.Matches(cell);
                    }
                }
            }
            else
            {
                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < cols; c++, i++)
                    {
                        if (mask[i] && !criterion.Matches(range[r, c]))
                        {
                            mask[i] = false;
                        }
                    }
                }
            }
        }
        return mask;
    }

    private static object[] Combine(object range1, object criteria1, object[] more)
    {
        var pairs = new object[2 + more.Length];
        pairs[0] = range1;
        pairs[1] = criteria1;
        Array.Copy(more, 0, pairs, 2, more.Length);
        return pairs;
    }

    private static object?[] FlattenRowMajor(object[,] block)
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

    private enum Op
    {
        Equal,
        NotEqual,
        Greater,
        Less,
        GreaterEqual,
        LessEqual,
    }

    private enum TokKind
    {
        Literal,
        Any,
        Star,
    }

    private readonly struct Tok
    {
        public Tok(TokKind kind, char ch)
        {
            Kind = kind;
            Ch = ch;
        }

        public TokKind Kind { get; }

        public char Ch { get; }
    }

    /// <summary>
    /// A parsed Excel-style criterion. Numeric criteria match only genuinely numeric cells;
    /// text criteria match case-insensitively with <c>*</c>/<c>?</c> wildcards (and <c>~</c>
    /// to escape a literal wildcard). The empty-operand forms match blank / non-blank cells.
    /// </summary>
    private sealed class Criterion
    {
        private readonly Op _op;
        private readonly bool _emptyOperand;
        private readonly bool _isNumeric;
        private readonly bool _isBool;
        private readonly bool _isError;
        private readonly bool _boolValue;
        private readonly ExcelError _errorValue;
        private readonly double _number;
        private readonly string _text;
        private readonly Tok[] _tokens;

        private Criterion(
            Op op,
            bool emptyOperand,
            bool isNumeric,
            double number,
            string text,
            Tok[] tokens,
            bool isBool = false,
            bool boolValue = false,
            bool isError = false,
            ExcelError errorValue = default)
        {
            _op = op;
            _emptyOperand = emptyOperand;
            _isNumeric = isNumeric;
            _number = number;
            _text = text;
            _tokens = tokens;
            _isBool = isBool;
            _boolValue = boolValue;
            _isError = isError;
            _errorValue = errorValue;
        }

        /// <summary>True for a plain numeric comparison; enables the fast mask loop.</summary>
        public bool IsPlainNumeric => _isNumeric;

        public static Criterion Parse(object? criteria)
        {
            switch (criteria)
            {
                case null:
                case ExcelEmpty:
                case ExcelMissing:
                    return new Criterion(Op.Equal, emptyOperand: true, false, 0d, string.Empty, Array.Empty<Tok>());
                case bool b:
                    // Excel keeps logicals a distinct type: criterion TRUE matches TRUE
                    // cells, never numeric 1 (and vice versa).
                    return new Criterion(Op.Equal, false, false, 0d, string.Empty, Array.Empty<Tok>(), isBool: true, boolValue: b);
                case ExcelError ce:
                    return new Criterion(Op.Equal, false, false, 0d, string.Empty, Array.Empty<Tok>(), isError: true, errorValue: ce);
                case double or int or long or float or decimal or DateTime:
                    Marshaling.TryToDouble(criteria, out var dn);
                    return new Criterion(Op.Equal, false, true, dn, string.Empty, Array.Empty<Tok>());
                case object[,] block:
                    if (block.Length == 1)
                    {
                        return Parse(block[0, 0]);
                    }
                    // Without this, ToStringSafe would render the array as the literal
                    // text "System.Object[,]" and the criterion would silently match
                    // nothing.
                    throw new ArgumentException("criteria must be a single value, not a multi-cell range.");
            }

            var s = Marshaling.ToStringSafe(criteria);
            var op = Op.Equal;
            var start = 0;
            if (s.StartsWith(">=", StringComparison.Ordinal))
            {
                op = Op.GreaterEqual;
                start = 2;
            }
            else if (s.StartsWith("<=", StringComparison.Ordinal))
            {
                op = Op.LessEqual;
                start = 2;
            }
            else if (s.StartsWith("<>", StringComparison.Ordinal))
            {
                op = Op.NotEqual;
                start = 2;
            }
            else if (s.StartsWith(">", StringComparison.Ordinal))
            {
                op = Op.Greater;
                start = 1;
            }
            else if (s.StartsWith("<", StringComparison.Ordinal))
            {
                op = Op.Less;
                start = 1;
            }
            else if (s.StartsWith("=", StringComparison.Ordinal))
            {
                op = Op.Equal;
                start = 1;
            }

            var operand = s.Substring(start);
            if (operand.Length == 0)
            {
                return new Criterion(op, emptyOperand: true, false, 0d, string.Empty, Array.Empty<Tok>());
            }
            if (Marshaling.TryToDouble(operand, out var num))
            {
                return new Criterion(op, false, true, num, operand, Array.Empty<Tok>());
            }
            if ((op == Op.Equal || op == Op.NotEqual) && TryParseErrorText(operand, out var errVal))
            {
                // Native COUNTIF matches error cells by their textual form
                // (=COUNTIF(rng,"#DIV/0!") counts #DIV/0! cells).
                return new Criterion(op, false, false, 0d, operand, Array.Empty<Tok>(), isError: true, errorValue: errVal);
            }
            if (TryParseDateLiteral(operand, out var serial))
            {
                // Native Excel parses date/time-literal operands (">1/1/2025") into
                // serials and compares numerically; without this the criterion degrades
                // to an ordinal STRING comparison against the serial's text form, which
                // matches essentially every date.
                return new Criterion(op, false, true, serial, operand, Array.Empty<Tok>());
            }
            return new Criterion(op, false, false, 0d, operand, Tokenize(operand));
        }

        private static bool TryParseErrorText(string s, out ExcelError err)
        {
            switch (s.ToUpperInvariant())
            {
                case "#DIV/0!": err = ExcelError.ExcelErrorDiv0; return true;
                case "#N/A": err = ExcelError.ExcelErrorNA; return true;
                case "#NAME?": err = ExcelError.ExcelErrorName; return true;
                case "#NULL!": err = ExcelError.ExcelErrorNull; return true;
                case "#NUM!": err = ExcelError.ExcelErrorNum; return true;
                case "#REF!": err = ExcelError.ExcelErrorRef; return true;
                case "#VALUE!": err = ExcelError.ExcelErrorValue; return true;
                case "#GETTING_DATA": err = ExcelError.ExcelErrorGettingData; return true;
                default: err = default; return false;
            }
        }

        private static bool TryParseDateLiteral(string operand, out double serial)
        {
            serial = 0d;
            // Require at least one digit so month names alone ("March") stay text
            // criteria; locale-dependent parsing mirrors how Excel itself interprets
            // date-literal criteria per locale.
            var hasDigit = false;
            foreach (var ch in operand)
            {
                if (char.IsAsciiDigit(ch))
                {
                    hasDigit = true;
                    break;
                }
            }
            if (!hasDigit)
            {
                return false;
            }
            if (!DateTime.TryParse(operand, CultureInfo.CurrentCulture, DateTimeStyles.NoCurrentDateDefault, out var dt))
            {
                return false;
            }
            if (dt.Date == DateTime.MinValue.Date)
            {
                // Time-only literal ("13:30"): Excel time criteria are day fractions.
                serial = dt.TimeOfDay.TotalDays;
                return true;
            }
            try
            {
                serial = dt.ToOADate();
                return true;
            }
            catch (OverflowException)
            {
                return false;
            }
        }

        public bool Matches(object? cell)
        {
            if (cell is ExcelError cellErr)
            {
                if (_isError)
                {
                    return _op == Op.NotEqual ? cellErr != _errorValue : cellErr == _errorValue;
                }
                return false;
            }
            var blank = IsBlank(cell);
            if (_emptyOperand)
            {
                return _op == Op.NotEqual ? !blank : blank;
            }
            if (blank)
            {
                // Native Excel counts blanks for "<>x": a blank cell is "not x" for any
                // non-empty operand, numeric or textual.
                return _op == Op.NotEqual;
            }
            if (_isError)
            {
                // An error criterion matches only error cells; "<>#N/A" is satisfied by
                // any non-error cell.
                return _op == Op.NotEqual;
            }
            if (_isBool)
            {
                if (cell is bool cb)
                {
                    return _op == Op.NotEqual ? cb != _boolValue : cb == _boolValue;
                }
                return _op == Op.NotEqual;
            }

            if (_isNumeric)
            {
                if (!IsNumericCell(cell, out var d))
                {
                    // A numeric criterion only matches numeric cells; a "<>" criterion is
                    // satisfied by everything that is not that number, including text.
                    return _op == Op.NotEqual;
                }
                return MatchesNumber(d);
            }

            // Text/wildcard criteria match text cells only, as native Excel does:
            // COUNTIF(range,"*") never counts numbers or logicals. "<>text" is satisfied
            // by any non-text cell.
            if (cell is not string text)
            {
                return _op == Op.NotEqual;
            }
            if (_op == Op.Equal)
            {
                return WildcardMatch(_tokens, text);
            }
            if (_op == Op.NotEqual)
            {
                return !WildcardMatch(_tokens, text);
            }
            var cmp = string.Compare(text, _text, StringComparison.OrdinalIgnoreCase);
            return _op switch
            {
                Op.Greater => cmp > 0,
                Op.Less => cmp < 0,
                Op.GreaterEqual => cmp >= 0,
                Op.LessEqual => cmp <= 0,
                _ => false,
            };
        }

        private static Tok[] Tokenize(string pattern)
        {
            var tokens = new List<Tok>(pattern.Length);
            for (var i = 0; i < pattern.Length; i++)
            {
                var c = pattern[i];
                if (c == '~' && i + 1 < pattern.Length)
                {
                    tokens.Add(new Tok(TokKind.Literal, pattern[i + 1]));
                    i++;
                }
                else if (c == '*')
                {
                    tokens.Add(new Tok(TokKind.Star, c));
                }
                else if (c == '?')
                {
                    tokens.Add(new Tok(TokKind.Any, c));
                }
                else
                {
                    tokens.Add(new Tok(TokKind.Literal, c));
                }
            }
            return tokens.ToArray();
        }

        /// <summary>
        /// The numeric comparison core, callable directly by the fast mask loop for raw
        /// <see cref="double"/> cells. Semantics identical to the numeric branch of
        /// <see cref="Matches"/>.
        /// </summary>
        public bool MatchesNumber(double d)
        {
            if (!double.IsFinite(d))
            {
                return _op == Op.NotEqual;
            }
            return _op switch
            {
                Op.Equal => d == _number,
                Op.NotEqual => d != _number,
                Op.Greater => d > _number,
                Op.Less => d < _number,
                Op.GreaterEqual => d >= _number,
                Op.LessEqual => d <= _number,
                _ => false,
            };
        }

        private static bool WildcardMatch(Tok[] tokens, string text)
        {
            int ti = 0, xi = 0, star = -1, starX = 0;
            while (xi < text.Length)
            {
                if (ti < tokens.Length
                    && (tokens[ti].Kind == TokKind.Any
                        || (tokens[ti].Kind == TokKind.Literal && CharEqual(tokens[ti].Ch, text[xi]))))
                {
                    ti++;
                    xi++;
                }
                else if (ti < tokens.Length && tokens[ti].Kind == TokKind.Star)
                {
                    star = ti;
                    starX = xi;
                    ti++;
                }
                else if (star != -1)
                {
                    ti = star + 1;
                    starX++;
                    xi = starX;
                }
                else
                {
                    return false;
                }
            }
            while (ti < tokens.Length && tokens[ti].Kind == TokKind.Star)
            {
                ti++;
            }
            return ti == tokens.Length;
        }

        private static bool CharEqual(char a, char b)
            => a == b || char.ToUpperInvariant(a) == char.ToUpperInvariant(b);
    }

    // ====================================================================
    // Numeric helpers
    // ====================================================================

    private static double Sum(List<double> nums)
    {
        var s = 0d;
        foreach (var v in nums)
        {
            s += v;
        }
        return s;
    }

    private static double Median(List<double> nums)
    {
        // The list is always a freshly built scratch collection: sort it in place
        // instead of paying a ToArray copy.
        var span = CollectionsMarshal.AsSpan(nums);
        span.Sort();
        var n = span.Length;
        var mid = n / 2;
        return (n & 1) == 1 ? span[mid] : (span[mid - 1] + span[mid]) / 2d;
    }

    private static object PercentileInc(List<double> nums, double k)
    {
        if (nums.Count == 0 || double.IsNaN(k) || k < 0d || k > 1d)
        {
            return ExcelError.ExcelErrorNum;
        }
        var span = CollectionsMarshal.AsSpan(nums);
        span.Sort();
        if (span.Length == 1)
        {
            return span[0];
        }
        var rank = k * (span.Length - 1);
        var lo = (int)Math.Floor(rank);
        var hi = (int)Math.Ceiling(rank);
        if (lo == hi)
        {
            return span[lo];
        }
        var frac = rank - lo;
        return span[lo] + ((span[hi] - span[lo]) * frac);
    }

    private static double Variance(List<double> nums, bool sample)
    {
        var n = nums.Count;
        var mean = Sum(nums) / n;
        var ss = 0d;
        foreach (var v in nums)
        {
            var d = v - mean;
            ss += d * d;
        }
        return sample ? ss / (n - 1) : ss / n;
    }

    private static object Mode(List<double> nums)
    {
        if (nums.Count == 0)
        {
            return ExcelError.ExcelErrorNA;
        }
        var counts = new Dictionary<double, int>(nums.Count);
        var order = new List<double>(nums.Count);
        foreach (var v in nums)
        {
            if (counts.TryGetValue(v, out var c))
            {
                counts[v] = c + 1;
            }
            else
            {
                counts[v] = 1;
                order.Add(v);
            }
        }
        var best = 0d;
        var bestCount = 0;
        foreach (var v in order)
        {
            if (counts[v] > bestCount)
            {
                bestCount = counts[v];
                best = v;
            }
        }
        return bestCount < 2 ? ExcelError.ExcelErrorNA : best;
    }

    private static object GeoMean(List<double> nums)
    {
        if (nums.Count == 0)
        {
            return ExcelError.ExcelErrorNum;
        }
        var sumLn = 0d;
        foreach (var v in nums)
        {
            if (v <= 0d)
            {
                return ExcelError.ExcelErrorNum;
            }
            sumLn += Math.Log(v);
        }
        return Math.Exp(sumLn / nums.Count);
    }

    private static object HarMean(List<double> nums)
    {
        if (nums.Count == 0)
        {
            return ExcelError.ExcelErrorNum;
        }
        var sumReciprocal = 0d;
        foreach (var v in nums)
        {
            if (v <= 0d)
            {
                return ExcelError.ExcelErrorNum;
            }
            sumReciprocal += 1d / v;
        }
        return nums.Count / sumReciprocal;
    }

    // ====================================================================
    // Cell classification + boundary guard
    // ====================================================================

    private static bool IsBlank(object? v)
        => v is null or ExcelEmpty or ExcelMissing || (v is string s && s.Length == 0);

    /// <summary>
    /// True only for genuinely numeric cells (number, date serial). Numeric-looking
    /// <see cref="string"/> cells and logicals are deliberately excluded so criteria and
    /// value handling match Excel, which treats <c>"5"</c> as text and keeps <c>TRUE</c>
    /// a distinct type (native SUMIFS ignores logicals in the sum range; <c>COUNTIF(rng,
    /// 1)</c> does not count TRUE).
    /// </summary>
    private static bool IsNumericCell(object? v, out double d)
    {
        if (v is string or bool)
        {
            d = 0d;
            return false;
        }
        return Marshaling.TryToDouble(v, out d);
    }

    /// <summary>
    /// Appends an injective, type-tagged encoding of a cell to a composite-key builder:
    /// <c>{tag}{length}:{text}</c>. The tag keeps numbers, text, bools, errors, and blanks
    /// in disjoint keyspaces (numeric 5 never merges with text "5"); the length prefix
    /// makes concatenation injective, closing the separator-collision class fixed in
    /// round 2 as U-02/U-03.
    /// </summary>
    private static void AppendCellKey(StringBuilder sb, object? cell)
    {
        char tag;
        string text;
        switch (cell)
        {
            case null:
            case ExcelEmpty:
            case ExcelMissing:
            case string { Length: 0 }:
                // Blank and zero-length text intentionally share a key, preserving the
                // pre-existing grouping behavior for empty cells.
                tag = '_';
                text = string.Empty;
                break;
            case bool b:
                tag = 'B';
                text = b ? "1" : "0";
                break;
            case ExcelError err:
                tag = 'E';
                text = ((int)err).ToString(CultureInfo.InvariantCulture);
                break;
            case string s:
                tag = 'S';
                text = s;
                break;
            default:
                if (Marshaling.TryToDouble(cell, out var d))
                {
                    tag = 'N';
                    text = d.ToString("R", CultureInfo.InvariantCulture);
                }
                else
                {
                    tag = 'S';
                    text = Marshaling.ToStringSafe(cell);
                }
                break;
        }
        sb.Append(tag).Append(text.Length).Append(':').Append(text);
    }

    /// <summary>
    /// Runs a UDF body and converts any non-critical exception into <c>#VALUE!</c> with a
    /// logged reason, mirroring the boundary discipline of the rest of the toolkit.
    /// </summary>
    private static object Guard(string name, Func<object> body)
    {
        try
        {
            return body();
        }
        catch (Exception ex) when (!IsCritical(ex))
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 1, "{0} failed: {1}", name, ex.Message);
            return ExcelError.ExcelErrorValue;
        }
    }

    private static bool IsCritical(Exception ex)
        => ex is OutOfMemoryException
            or StackOverflowException
            or AccessViolationException
            or System.Threading.ThreadAbortException;
}
