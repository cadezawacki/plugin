using System;
using System.Collections.Generic;
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

        Func<double, bool> isOutlier;
        if (values.Count < 2)
        {
            isOutlier = static _ => false;
        }
        else
        {
            switch (m)
            {
                case "iqr":
                {
                    var sorted = values.ToArray();
                    Array.Sort(sorted);
                    var q1 = PercentileSorted(sorted, 0.25);
                    var q3 = PercentileSorted(sorted, 0.75);
                    var iqr = q3 - q1;
                    var k = threshold > 0d ? threshold : 1.5;
                    var lower = q1 - (k * iqr);
                    var upper = q3 + (k * iqr);
                    isOutlier = d => d < lower || d > upper;
                    break;
                }
                case "zscore":
                {
                    var (mean, std) = MeanStd(values, sample: true);
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
                    var sorted = values.ToArray();
                    Array.Sort(sorted);
                    var median = PercentileSorted(sorted, 0.5);
                    var deviations = new double[values.Count];
                    for (var i = 0; i < values.Count; i++)
                    {
                        deviations[i] = Math.Abs(values[i] - median);
                    }
                    Array.Sort(deviations);
                    var mad = PercentileSorted(deviations, 0.5);
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
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                result[r, c] = Marshaling.TryToDouble(block[r, c], out var d) && isOutlier(d);
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
        var sorted = values.ToArray();
        Array.Sort(sorted);

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

    // ---------- Helpers ----------

    private static bool IsCellBlank(object? v)
        => v is null or ExcelEmpty or ExcelMissing || (v is string s && s.Length == 0);

    private static (double Mean, double Std) MeanStd(List<double> values, bool sample)
    {
        var n = values.Count;
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

    /// <summary>Inclusive percentile (PERCENTILE.INC) of an ascending-sorted array.</summary>
    private static double PercentileSorted(double[] sortedAsc, double p)
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
