using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelDna.Integration;

namespace ExcelPerfToolkit;

/// <summary>
/// Round 3 series cleaning and robust statistics: directional blank filling, outlier
/// flagging, and quantiles. All operate on a whole block in one managed-memory pass and are
/// pure CPU, so each is registered <c>IsThreadSafe = true</c> and eligible for Excel's
/// multithreaded recalc.
///
/// <para>Bottlenecks addressed: #1 and #2 (one bulk in, one bulk out, no COM re-entry).</para>
/// </summary>
public static class SeriesUtilities
{
    // ---------- Fill forward (generalized fill-down) ----------

    /// <summary>
    /// Fills blank cells with the most recent non-blank value in the chosen direction:
    /// <c>down</c> (default), <c>up</c>, <c>right</c>, or <c>left</c>. <c>down</c>/<c>up</c>
    /// carry values along each column; <c>right</c>/<c>left</c> along each row. Leading
    /// blanks with no prior value stay blank. Generalizes <see cref="DeveloperUtilities.FillDown"/>.
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.FILLFORWARD", Description = "Fill blanks from the previous value in a direction (down|up|right|left).", Category = "EPT.Series", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] FillForward(
        object[,] block,
        [ExcelArgument(Name = "direction", Description = "down (default) | up | right | left")] string direction = "down")
    {
        ArgumentNullException.ThrowIfNull(block);
        var dir = string.IsNullOrWhiteSpace(direction) ? "down" : direction.Trim().ToLowerInvariant();
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var result = new object[rows, cols];

        switch (dir)
        {
            case "down":
                for (var c = 0; c < cols; c++)
                {
                    object? last = null;
                    for (var r = 0; r < rows; r++)
                    {
                        last = FillStep(block[r, c], last, result, r, c);
                    }
                }
                break;
            case "up":
                for (var c = 0; c < cols; c++)
                {
                    object? last = null;
                    for (var r = rows - 1; r >= 0; r--)
                    {
                        last = FillStep(block[r, c], last, result, r, c);
                    }
                }
                break;
            case "right":
                for (var r = 0; r < rows; r++)
                {
                    object? last = null;
                    for (var c = 0; c < cols; c++)
                    {
                        last = FillStep(block[r, c], last, result, r, c);
                    }
                }
                break;
            case "left":
                for (var r = 0; r < rows; r++)
                {
                    object? last = null;
                    for (var c = cols - 1; c >= 0; c--)
                    {
                        last = FillStep(block[r, c], last, result, r, c);
                    }
                }
                break;
            default:
                throw new ArgumentException("direction must be one of: down, up, right, left.", nameof(direction));
        }
        return result;
    }

    private static object? FillStep(object? value, object? last, object[,] result, int r, int c)
    {
        if (IsCellBlank(value))
        {
            result[r, c] = last ?? ExcelEmpty.Value;
            return last;
        }
        result[r, c] = value!;
        return value;
    }

    // ---------- Outliers ----------

    /// <summary>
    /// Flags numeric outliers across all numeric cells of the block, returning a same-shaped
    /// block of <c>TRUE</c>/<c>FALSE</c>. Methods: <c>iqr</c> (default, Tukey fences at
    /// <c>threshold</c>·IQR, default 1.5), <c>zscore</c> (|z| &gt; threshold, default 3), and
    /// <c>mad</c> (modified z-score from the median absolute deviation, default 3). A
    /// <paramref name="threshold"/> of 0 selects the per-method default. Non-numeric cells are
    /// <c>FALSE</c>; fewer than two points (or zero spread) yields all <c>FALSE</c>.
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.OUTLIERS", Description = "Flag numeric outliers (iqr|zscore|mad) as a TRUE/FALSE block.", Category = "EPT.Series", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] Outliers(
        object[,] block,
        [ExcelArgument(Name = "method", Description = "iqr (default) | zscore | mad")] string method = "iqr",
        [ExcelArgument(Name = "threshold", Description = "0 (default) uses the per-method default.")] double threshold = 0)
    {
        ArgumentNullException.ThrowIfNull(block);
        var m = string.IsNullOrWhiteSpace(method) ? "iqr" : method.Trim().ToLowerInvariant();
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);

        // Parse every cell exactly once: the flagging pass below reads the cached
        // values instead of re-running TryToDouble (a second full string re-parse on
        // text-numeric blocks).
        var n = rows * cols;
        var vals = new double[n];
        var ok = new bool[n];
        var count = 0;
        var idx = 0;
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++, idx++)
            {
                if (Marshaling.TryToDouble(block[r, c], out var d))
                {
                    vals[idx] = d;
                    ok[idx] = true;
                    count++;
                }
            }
        }

        Func<double, bool> isOutlier;
        if (count < 2)
        {
            isOutlier = static _ => false;
        }
        else
        {
            var nums = new double[count];
            var k2 = 0;
            for (var i = 0; i < n; i++)
            {
                if (ok[i])
                {
                    nums[k2++] = vals[i];
                }
            }
            switch (m)
            {
                case "iqr":
                {
                    Array.Sort(nums);
                    var q1 = PercentileSorted(nums, 0.25);
                    var q3 = PercentileSorted(nums, 0.75);
                    var iqr = q3 - q1;
                    var k = threshold > 0d ? threshold : 1.5;
                    var lower = q1 - (k * iqr);
                    var upper = q3 + (k * iqr);
                    isOutlier = d => d < lower || d > upper;
                    break;
                }
                case "zscore":
                {
                    var (mean, std) = MeanStd(nums, sample: true);
                    var k = threshold > 0d ? threshold : 3d;
                    if (std <= 0d)
                    {
                        isOutlier = static _ => false;
                    }
                    else
                    {
                        isOutlier = d => Math.Abs(d - mean) / std > k;
                    }
                    break;
                }
                case "mad":
                {
                    Array.Sort(nums);
                    var median = PercentileSorted(nums, 0.5);
                    // The deviation multiset doesn't depend on order, so overwrite the
                    // sorted buffer in place instead of allocating a second array.
                    for (var i = 0; i < nums.Length; i++)
                    {
                        nums[i] = Math.Abs(nums[i] - median);
                    }
                    Array.Sort(nums);
                    var mad = PercentileSorted(nums, 0.5);
                    var k = threshold > 0d ? threshold : 3d;
                    if (mad <= 0d)
                    {
                        isOutlier = static _ => false;
                    }
                    else
                    {
                        isOutlier = d => Math.Abs(0.6745 * (d - median) / mad) > k;
                    }
                    break;
                }
                default:
                    throw new ArgumentException("method must be one of: iqr, zscore, mad.", nameof(method));
            }
        }

        var result = new object[rows, cols];
        idx = 0;
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++, idx++)
            {
                result[r, c] = ok[idx] && isOutlier(vals[idx]);
            }
        }
        return result;
    }

    // ---------- Quantiles ----------

    /// <summary>
    /// Returns the inclusive quantile (matching <c>PERCENTILE.INC</c>) of the block's numeric
    /// values for each probability in <paramref name="probabilities"/>. The result mirrors the
    /// shape of <paramref name="probabilities"/>. Probabilities outside [0, 1], non-numeric
    /// probabilities, or an empty data set yield <c>#NUM!</c> in that cell.
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.QUANTILES", Description = "Inclusive quantiles of a block's numeric values for each requested probability.", Category = "EPT.Series", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] Quantiles(
        object[,] block,
        [ExcelArgument(Name = "probabilities", Description = "One or more probabilities in [0,1]; result mirrors this shape.")] object[,] probabilities)
    {
        ArgumentNullException.ThrowIfNull(block);
        ArgumentNullException.ThrowIfNull(probabilities);

        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var values = new List<double>(rows * cols);
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                if (Marshaling.TryToDouble(block[r, c], out var d))
                {
                    values.Add(d);
                }
            }
        }
        var sorted = CollectionsMarshal.AsSpan(values);
        sorted.Sort();

        var pr = probabilities.GetLength(0);
        var pc = probabilities.GetLength(1);
        var result = new object[pr, pc];
        for (var r = 0; r < pr; r++)
        {
            for (var c = 0; c < pc; c++)
            {
                if (sorted.Length == 0
                    || !Marshaling.TryToDouble(probabilities[r, c], out var p)
                    || p < 0d || p > 1d)
                {
                    result[r, c] = ExcelError.ExcelErrorNum;
                }
                else
                {
                    result[r, c] = PercentileSorted(sorted, p);
                }
            }
        }
        return result;
    }

    // ---------- Running (cumulative) aggregates ----------

    /// <summary>
    /// Cumulative aggregate down each column in one O(N) pass: row <i>r</i> holds the
    /// aggregate of rows 1..<i>r</i>. Replaces the classic expanding-range pattern
    /// (<c>=SUM($A$1:A1)</c> copied down), which is O(N²) and recomputed on every edit.
    /// Operations: <c>sum</c>, <c>count</c>, <c>average</c>, <c>min</c>, <c>max</c>,
    /// <c>product</c>. Non-numeric cells leave the running state unchanged (native
    /// SUM-over-blank semantics); an error cell propagates to itself and every later row of
    /// its column, exactly as the native expanding range would. Before the first numeric
    /// value the output matches the native empty-range result: 0 for
    /// sum/count/min/max/product, <c>#DIV/0!</c> for average.
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.RUNNINGAGG", Description = "Cumulative sum|count|average|min|max|product down each column in one O(N) pass (replaces =SUM($A$1:A1) copied down).", Category = "EPT.Series", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] RunningAgg(
        object[,] block,
        [ExcelArgument(Name = "operation", Description = "sum|count|average|min|max|product")] string operation)
    {
        ArgumentNullException.ThrowIfNull(block);
        var op = (operation ?? string.Empty).Trim().ToLowerInvariant();
        if (op is not ("sum" or "count" or "average" or "min" or "max" or "product"))
        {
            throw new ArgumentException("operation must be one of: sum, count, average, min, max, product.", nameof(operation));
        }
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var result = new object[rows, cols];
        for (var c = 0; c < cols; c++)
        {
            double sum = 0d, product = 1d, min = double.PositiveInfinity, max = double.NegativeInfinity;
            var count = 0;
            ExcelError? poisoned = null;
            for (var r = 0; r < rows; r++)
            {
                var cell = block[r, c];
                if (poisoned is null && cell is ExcelError err)
                {
                    poisoned = err;
                }
                if (poisoned is not null)
                {
                    result[r, c] = poisoned;
                    continue;
                }
                if (IsNumericCell(cell, out var d))
                {
                    count++;
                    sum += d;
                    product *= d;
                    if (d < min)
                    {
                        min = d;
                    }
                    if (d > max)
                    {
                        max = d;
                    }
                }
                result[r, c] = op switch
                {
                    "sum" => (object)sum,
                    "count" => (double)count,
                    "average" => count == 0 ? (object)ExcelError.ExcelErrorDiv0 : sum / count,
                    "min" => count == 0 ? 0d : min,
                    "max" => count == 0 ? 0d : max,
                    _ => count == 0 ? 0d : product,
                };
            }
        }
        return result;
    }

    // ---------- Moving (trailing window) aggregates ----------

    /// <summary>
    /// Trailing-window aggregate down each column in a single pass - the fast, non-volatile
    /// replacement for the <c>=AVERAGE(OFFSET(...))</c> moving-average idiom, which is both
    /// volatile and O(N·W). The window covers the current row and the <c>window - 1</c> rows
    /// above it. Sums and counts update incrementally (O(N) per column); min/max use a
    /// monotonic deque (O(N)); median maintains a sorted window (O(N·W)). Operations:
    /// <c>sum</c>, <c>average</c>, <c>count</c>, <c>min</c>, <c>max</c>, <c>median</c>,
    /// <c>stdev</c> (sample).
    ///
    /// <para>Non-numeric cells contribute nothing; a cell emits a value only when the
    /// window holds at least <paramref name="minPeriods"/> numeric cells (default: the
    /// window size), otherwise it stays blank - pandas' <c>min_periods</c> semantics. An
    /// error cell poisons every window containing it, as a native windowed formula
    /// would.</para>
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.MOVINGAGG", Description = "Trailing-window sum|average|count|min|max|median|stdev down each column in one pass (replaces volatile OFFSET windows).", Category = "EPT.Series", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] MovingAgg(
        object[,] block,
        [ExcelArgument(Name = "window", Description = "Trailing window size in rows (>= 1), including the current row.")] double window,
        [ExcelArgument(Name = "operation", Description = "sum|average|count|min|max|median|stdev")] string operation,
        [ExcelArgument(Name = "min_periods", Description = "Optional minimum numeric cells in the window to emit a value; defaults to the window size.")] object minPeriods)
    {
        ArgumentNullException.ThrowIfNull(block);
        if (double.IsNaN(window) || window < 1d || window != Math.Truncate(window) || window > 1_048_576d)
        {
            throw new ArgumentException("window must be an integer between 1 and 1,048,576 rows.", nameof(window));
        }
        var w = (int)window;
        var op = (operation ?? string.Empty).Trim().ToLowerInvariant();
        var needSums = op is "sum" or "average" or "count" or "stdev";
        var needMinMax = op is "min" or "max";
        var needMedian = op == "median";
        if (!needSums && !needMinMax && !needMedian)
        {
            throw new ArgumentException("operation must be one of: sum, average, count, min, max, median, stdev.", nameof(operation));
        }
        var minP = ResolveMinPeriods(minPeriods, w);

        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var result = new object[rows, cols];
        var vals = new double[rows];
        var valid = new bool[rows];
        var errs = new ExcelError?[rows];
        // Monotonic deque of row indexes; sign folds min into the max recurrence.
        var deque = needMinMax ? new int[rows] : Array.Empty<int>();
        var sign = op == "min" ? -1d : 1d;

        for (var c = 0; c < cols; c++)
        {
            // Parse the column once; the sliding pass below never re-runs TryToDouble.
            for (var r = 0; r < rows; r++)
            {
                var cell = block[r, c];
                errs[r] = cell as ExcelError?;
                valid[r] = errs[r] is null && IsNumericCell(cell, out vals[r]);
            }

            double sum = 0d, sumSq = 0d;
            int count = 0, errorCount = 0;
            int dHead = 0, dTail = 0;
            var sortedWindow = needMedian ? new List<double>(Math.Min(w, rows)) : null;

            for (var r = 0; r < rows; r++)
            {
                // Enter row r.
                if (errs[r] is not null)
                {
                    errorCount++;
                }
                else if (valid[r])
                {
                    var d = vals[r];
                    count++;
                    if (needSums)
                    {
                        sum += d;
                        sumSq += d * d;
                    }
                    if (needMinMax)
                    {
                        var v = d * sign;
                        while (dTail > dHead && vals[deque[dTail - 1]] * sign <= v)
                        {
                            dTail--;
                        }
                        deque[dTail++] = r;
                    }
                    if (needMedian)
                    {
                        InsertSorted(sortedWindow!, d);
                    }
                }

                // Leave row r - window.
                var leave = r - w;
                if (leave >= 0)
                {
                    if (errs[leave] is not null)
                    {
                        errorCount--;
                    }
                    else if (valid[leave])
                    {
                        count--;
                        if (needSums)
                        {
                            sum -= vals[leave];
                            sumSq -= vals[leave] * vals[leave];
                        }
                        if (needMedian)
                        {
                            RemoveSorted(sortedWindow!, vals[leave]);
                        }
                    }
                }
                if (needMinMax)
                {
                    var floor = r - w + 1;
                    while (dTail > dHead && deque[dHead] < floor)
                    {
                        dHead++;
                    }
                }

                // Emit.
                if (errorCount > 0)
                {
                    var start = Math.Max(0, r - w + 1);
                    for (var k = start; k <= r; k++)
                    {
                        if (errs[k] is { } windowErr)
                        {
                            result[r, c] = windowErr;
                            break;
                        }
                    }
                }
                else if (count < minP)
                {
                    result[r, c] = ExcelEmpty.Value;
                }
                else
                {
                    result[r, c] = op switch
                    {
                        "sum" => (object)sum,
                        "count" => (double)count,
                        "average" => sum / count,
                        "min" or "max" => vals[deque[dHead]],
                        "median" => MedianSorted(sortedWindow!),
                        // Sample stdev from the incremental sums; clamp tiny negative
                        // round-off before the sqrt.
                        _ => count < 2
                            ? (object)ExcelError.ExcelErrorDiv0
                            : Math.Sqrt(Math.Max(0d, (sumSq - (sum * sum / count)) / (count - 1))),
                    };
                }
            }
        }
        return result;
    }

    // ---------- Block-wide ranking ----------

    /// <summary>
    /// Ranks every numeric cell against all numeric cells of the block in one
    /// O(N log N) pass - the batch replacement for a column of native <c>RANK</c> formulas,
    /// which is O(N²) because each formula re-scans the range. <paramref name="order"/>
    /// follows native <c>RANK</c>: 0 or omitted ranks descending (largest = 1); any other
    /// number ranks ascending. <paramref name="ties"/>: <c>eq</c> (default, like
    /// <c>RANK.EQ</c>), <c>avg</c> (<c>RANK.AVG</c>), <c>dense</c> (no gaps after ties -
    /// <b>no native equivalent</b>), or <c>ordinal</c> (unique 1..N, first-seen order
    /// breaks ties). Non-numeric cells return blank; error cells propagate.
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.RANKB", Description = "Rank every numeric cell against the whole block in one O(N log N) pass (eq|avg|dense|ordinal ties).", Category = "EPT.Series", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] RankB(
        object[,] block,
        [ExcelArgument(Name = "order", Description = "0/omitted = descending (largest = 1, the native RANK default); any other number = ascending.")] object order,
        [ExcelArgument(Name = "ties", Description = "eq (default, RANK.EQ) | avg (RANK.AVG) | dense | ordinal")] object ties)
    {
        ArgumentNullException.ThrowIfNull(block);
        var ascending = ResolveOrder(order);
        var mode = ResolveTies(ties);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var result = new object[rows, cols];
        var points = new List<(double Value, int Flat)>(rows * cols);
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                var cell = block[r, c];
                if (cell is ExcelError err)
                {
                    result[r, c] = err;
                }
                else if (IsNumericCell(cell, out var d))
                {
                    points.Add((d, (r * cols) + c));
                }
                else
                {
                    result[r, c] = ExcelEmpty.Value;
                }
            }
        }

        var dir = ascending ? 1 : -1;
        points.Sort((a, b) =>
        {
            var byValue = dir * a.Value.CompareTo(b.Value);
            return byValue != 0 ? byValue : a.Flat.CompareTo(b.Flat);
        });

        var i = 0;
        var dense = 0;
        while (i < points.Count)
        {
            var j = i;
            while (j < points.Count && points[j].Value == points[i].Value)
            {
                j++;
            }
            dense++;
            for (var k = i; k < j; k++)
            {
                var rank = mode switch
                {
                    "eq" => (double)(i + 1),
                    "avg" => (i + j + 1) / 2d,
                    "dense" => (double)dense,
                    _ => (double)(k + 1), // ordinal
                };
                var flat = points[k].Flat;
                result[flat / cols, flat % cols] = rank;
            }
            i = j;
        }
        return result;
    }

    /// <summary>
    /// <c>PERCENTRANK.INC</c> of every numeric cell against all numeric cells of the block
    /// in one O(N log N) pass - the batch replacement for a column of native
    /// <c>PERCENTRANK</c> formulas (O(N²)). For each value the rank fraction is (count of
    /// strictly smaller values) / (N - 1), truncated to <paramref name="significance"/>
    /// significant decimal digits (default 3, like native). Fewer than two numeric cells
    /// yields <c>#N/A</c> for the numeric cells. Non-numeric cells return blank; error
    /// cells propagate.
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.PERCENTRANKB", Description = "PERCENTRANK.INC of every numeric cell against the whole block in one O(N log N) pass.", Category = "EPT.Series", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] PercentRankB(
        object[,] block,
        [ExcelArgument(Name = "significance", Description = "Optional decimal digits (1-15) kept by truncation; defaults to 3, like native PERCENTRANK.")] object significance)
    {
        ArgumentNullException.ThrowIfNull(block);
        var pow = Math.Pow(10d, ResolveSignificance(significance));
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var result = new object[rows, cols];
        var nums = new List<double>(rows * cols);
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                if (IsNumericCell(block[r, c], out var d))
                {
                    nums.Add(d);
                }
            }
        }
        var sorted = nums.ToArray();
        Array.Sort(sorted);
        var n = sorted.Length;
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                var cell = block[r, c];
                if (cell is ExcelError err)
                {
                    result[r, c] = err;
                }
                else if (!IsNumericCell(cell, out var d))
                {
                    result[r, c] = ExcelEmpty.Value;
                }
                else if (n < 2)
                {
                    result[r, c] = ExcelError.ExcelErrorNA;
                }
                else
                {
                    var frac = LowerBound(sorted, d) / (double)(n - 1);
                    result[r, c] = Math.Floor(frac * pow) / pow;
                }
            }
        }
        return result;
    }

    // ---------- Helpers ----------

    /// <summary>
    /// True only for genuinely numeric cells: numeric-looking text and logicals are
    /// excluded so the running/moving/rank aggregates match native SUM/RANK semantics.
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

    private static int ResolveMinPeriods(object minPeriods, int window)
    {
        if (minPeriods is null or ExcelMissing or ExcelEmpty)
        {
            return window;
        }
        if (!Marshaling.TryToDouble(minPeriods, out var d) || d < 1d || d != Math.Truncate(d) || d > int.MaxValue)
        {
            throw new ArgumentException("min_periods must be an integer >= 1.", nameof(minPeriods));
        }
        return (int)d;
    }

    private static bool ResolveOrder(object order)
    {
        if (order is null or ExcelMissing or ExcelEmpty)
        {
            return false;
        }
        if (order is bool b)
        {
            return b;
        }
        if (Marshaling.TryToDouble(order, out var d))
        {
            return d != 0d;
        }
        throw new ArgumentException("order must be 0 (descending, the native default) or a nonzero number (ascending).", nameof(order));
    }

    private static string ResolveTies(object ties)
    {
        if (ties is null or ExcelMissing or ExcelEmpty)
        {
            return "eq";
        }
        var s = Marshaling.ToStringSafe(ties).Trim().ToLowerInvariant();
        if (s.Length == 0)
        {
            return "eq";
        }
        if (s is not ("eq" or "avg" or "dense" or "ordinal"))
        {
            throw new ArgumentException("ties must be one of: eq, avg, dense, ordinal.", nameof(ties));
        }
        return s;
    }

    private static int ResolveSignificance(object significance)
    {
        if (significance is null or ExcelMissing or ExcelEmpty)
        {
            return 3;
        }
        if (!Marshaling.TryToDouble(significance, out var d) || d < 1d || d != Math.Truncate(d) || d > 15d)
        {
            throw new ArgumentException("significance must be an integer between 1 and 15.", nameof(significance));
        }
        return (int)d;
    }

    private static void InsertSorted(List<double> sorted, double value)
    {
        var at = sorted.BinarySearch(value);
        sorted.Insert(at < 0 ? ~at : at, value);
    }

    private static void RemoveSorted(List<double> sorted, double value)
    {
        var at = sorted.BinarySearch(value);
        if (at >= 0)
        {
            sorted.RemoveAt(at);
        }
    }

    private static double MedianSorted(List<double> sorted)
    {
        var mid = sorted.Count / 2;
        return (sorted.Count & 1) == 1 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2d;
    }

    /// <summary>First index in <paramref name="sortedAsc"/> whose value is &gt;= <paramref name="value"/>.</summary>
    private static int LowerBound(double[] sortedAsc, double value)
    {
        int lo = 0, hi = sortedAsc.Length;
        while (lo < hi)
        {
            var mid = lo + ((hi - lo) / 2);
            if (sortedAsc[mid] < value)
            {
                lo = mid + 1;
            }
            else
            {
                hi = mid;
            }
        }
        return lo;
    }

    private static bool IsCellBlank(object? v)
        => v is null or ExcelEmpty or ExcelMissing || (v is string s && s.Length == 0);

    private static (double Mean, double Std) MeanStd(ReadOnlySpan<double> values, bool sample)
    {
        var n = values.Length;
        var sum = 0d;
        foreach (var v in values)
        {
            sum += v;
        }
        var mean = sum / n;
        var ss = 0d;
        foreach (var v in values)
        {
            var d = v - mean;
            ss += d * d;
        }
        var denom = sample ? n - 1 : n;
        var std = denom > 0 ? Math.Sqrt(ss / denom) : 0d;
        return (mean, std);
    }

    /// <summary>Inclusive percentile (PERCENTILE.INC) of an ascending-sorted sequence.</summary>
    private static double PercentileSorted(ReadOnlySpan<double> sortedAsc, double p)
    {
        if (sortedAsc.Length == 1)
        {
            return sortedAsc[0];
        }
        var rank = p * (sortedAsc.Length - 1);
        var lo = (int)Math.Floor(rank);
        var hi = (int)Math.Ceiling(rank);
        if (lo == hi)
        {
            return sortedAsc[lo];
        }
        var frac = rank - lo;
        return sortedAsc[lo] + ((sortedAsc[hi] - sortedAsc[lo]) * frac);
    }
}
