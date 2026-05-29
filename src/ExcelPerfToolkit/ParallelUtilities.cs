using System;
using System.Diagnostics;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelPerfToolkit;

/// <summary>
/// Thread-safe UDFs and explicit parallel batch transforms. This file is the direct fix
/// for bottleneck #3 (VBA single-threaded recalc): every UDF in here is registered with
/// <c>IsThreadSafe = true</c> and <c>IsVolatile = false</c>, so Excel's MTR engine is
/// free to schedule it across cores during a recalc that would otherwise serialize on
/// VBA's single-thread limit.
///
/// **Rules followed by every function here**:
///   * No shared mutable static state. The only static fields below are immutable
///     configuration constants and a <see cref="TraceSource"/>; TraceSource itself is
///     thread-safe and only invoked outside the hot loop.
///   * No access to the Excel object model. We do not call <c>ExcelReference.GetValue</c>,
///     <c>ExcelReference.SetValue</c>, <c>ExcelDnaUtil.Application</c>, or any non-pure
///     <see cref="XlCall"/> entry. Thread-safe UDFs run on worker threads and the
///     object model is owned by the main thread.
///   * Inputs and outputs are <c>object[,]</c> only - the marshaler handles the single
///     boundary crossing per invocation.
/// </summary>
public static class ParallelUtilities
{
    private static readonly TraceSource TraceSource = new("ExcelPerfToolkit.ParallelUtilities", SourceLevels.Information);

    /// <summary>
    /// Default row-count threshold below which <see cref="ParallelBatchTransform"/> runs
    /// serially. Parallelization overhead (task scheduling, partitioning, false sharing
    /// of cache lines on small blocks) only pays off above this size in our measurements.
    /// </summary>
    public const int DefaultParallelThreshold = 4_096;

    /// <summary>
    /// Delegate used by <see cref="ParallelBatchTransform"/>. Receives the row index,
    /// a read-only span of the input row, and a writable span of the same length for
    /// the output row. Implementations must be pure (no shared state).
    /// </summary>
    public delegate void RowTransform(int rowIndex, ReadOnlySpan<double> input, Span<double> output);

    // ------------------------------------------------------------------
    // Thread-safe UDF #1: row sums
    // ------------------------------------------------------------------

    /// <summary>
    /// Returns a single-column block of row sums over a numeric block. All non-numeric
    /// cells are treated as 0. Pure CPU; safe for Excel's multithreaded recalc because:
    /// (a) inputs arrive as an immutable <c>object[,]</c> on the worker thread,
    /// (b) computation reads only locals, (c) the return value is freshly allocated.
    /// Marshaling cost: 1 read + 1 write = 2 crossings, both O(1) in N.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.MT.ROWSUMS",
        Description = "Thread-safe per-row sums over a numeric block (MTR-eligible).",
        Category = "EPT.Parallel",
        IsThreadSafe = true,
        IsVolatile = false)]
    public static object[,] RowSums(object[,] block)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var result = new object[rows, 1];
        for (var r = 0; r < rows; r++)
        {
            var sum = 0d;
            for (var c = 0; c < cols; c++)
            {
                if (Marshaling.TryToDouble(block[r, c], out var d))
                {
                    sum += d;
                }
            }
            result[r, 0] = sum;
        }
        return result;
    }

    // ------------------------------------------------------------------
    // Thread-safe UDF #2: per-row z-score normalization
    // ------------------------------------------------------------------

    /// <summary>
    /// Returns a block where each row has been z-score normalized: <c>(x - mean) / stdev</c>.
    /// Rows with zero variance are returned as zeros. Pure CPU; safe for MTR for the
    /// same reasons as <see cref="RowSums"/>.
    /// Marshaling cost: 1 read + 1 write = 2 crossings.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.MT.ROWZSCORES",
        Description = "Thread-safe per-row z-score normalization (MTR-eligible).",
        Category = "EPT.Parallel",
        IsThreadSafe = true,
        IsVolatile = false)]
    public static object[,] RowZScores(object[,] block)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var result = new object[rows, cols];
        for (var r = 0; r < rows; r++)
        {
            var sum = 0d;
            var count = 0;
            for (var c = 0; c < cols; c++)
            {
                if (Marshaling.TryToDouble(block[r, c], out var d))
                {
                    sum += d;
                    count++;
                }
            }
            if (count == 0)
            {
                for (var c = 0; c < cols; c++)
                {
                    result[r, c] = 0d;
                }
                continue;
            }
            var mean = sum / count;
            var sumSq = 0d;
            for (var c = 0; c < cols; c++)
            {
                if (Marshaling.TryToDouble(block[r, c], out var d))
                {
                    var diff = d - mean;
                    sumSq += diff * diff;
                }
            }
            var stdev = count > 1 ? Math.Sqrt(sumSq / (count - 1)) : 0d;
            if (stdev == 0d)
            {
                for (var c = 0; c < cols; c++)
                {
                    result[r, c] = 0d;
                }
                continue;
            }
            for (var c = 0; c < cols; c++)
            {
                if (Marshaling.TryToDouble(block[r, c], out var d))
                {
                    result[r, c] = (d - mean) / stdev;
                }
                else
                {
                    result[r, c] = 0d;
                }
            }
        }
        return result;
    }

    // ------------------------------------------------------------------
    // Thread-safe UDF #3: element-wise polynomial evaluation
    // ------------------------------------------------------------------

    /// <summary>
    /// Evaluates a polynomial <c>c0 + c1*x + c2*x^2 + ... + ck*x^k</c> element-wise over
    /// <paramref name="block"/>. Coefficients are taken as the row-major flattening of
    /// <paramref name="coefficients"/>, lowest order first. Horner's method is used.
    /// Marshaling cost: 1 read + 1 write per invocation.
    /// Thread-safety: SAFE for MTR. No shared state; reads only locals.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.MT.POLYEVAL",
        Description = "Thread-safe element-wise polynomial evaluation with Horner's method (MTR-eligible).",
        Category = "EPT.Parallel",
        IsThreadSafe = true,
        IsVolatile = false)]
    public static object[,] PolyEval(object[,] block, object[,] coefficients)
    {
        ArgumentNullException.ThrowIfNull(block);
        ArgumentNullException.ThrowIfNull(coefficients);

        var coeffsList = new System.Collections.Generic.List<double>(16);
        var cRows = coefficients.GetLength(0);
        var cCols = coefficients.GetLength(1);
        for (var r = 0; r < cRows; r++)
        {
            for (var c = 0; c < cCols; c++)
            {
                var v = coefficients[r, c];
                if (Marshaling.IsBlankOrError(v))
                {
                    continue;
                }
                if (Marshaling.TryToDouble(v, out var d))
                {
                    coeffsList.Add(d);
                }
            }
        }
        if (coeffsList.Count == 0)
        {
            return Marshaling.ErrorBlock(ExcelError.ExcelErrorValue);
        }
        var coeffs = coeffsList.ToArray();
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var result = new object[rows, cols];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                if (!Marshaling.TryToDouble(block[r, c], out var x))
                {
                    result[r, c] = 0d;
                    continue;
                }
                // Horner: start with the highest-order coefficient.
                var acc = coeffs[coeffs.Length - 1];
                for (var i = coeffs.Length - 2; i >= 0; i--)
                {
                    acc = acc * x + coeffs[i];
                }
                result[r, c] = acc;
            }
        }
        return result;
    }

    // ------------------------------------------------------------------
    // Explicit parallel batch transform
    // ------------------------------------------------------------------

    /// <summary>
    /// Applies <paramref name="rowTransform"/> to each row of <paramref name="input"/>
    /// in parallel, partitioning by row. Falls back to serial execution when the row
    /// count is below <paramref name="parallelThresholdRows"/> (default
    /// <see cref="DefaultParallelThreshold"/>) because the overhead of starting tasks,
    /// partitioning the iteration space, and contending the result array's cache lines
    /// dominates on small inputs.
    ///
    /// Thread-safety: <paramref name="rowTransform"/> must be pure: it must not touch
    /// shared state, Excel's object model, or anything that mutates outside its row.
    /// The transform is called once per row and may write to any column of that row
    /// in <paramref name="output"/> - rows do not share memory and writes do not race.
    ///
    /// Marshaling cost: 0 boundary crossings; this operates on managed arrays only.
    /// </summary>
    public static void ParallelBatchTransform(
        double[,] input,
        double[,] output,
        RowTransform rowTransform,
        int parallelThresholdRows = DefaultParallelThreshold,
        int? maxDegreeOfParallelism = null)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentNullException.ThrowIfNull(output);
        ArgumentNullException.ThrowIfNull(rowTransform);
        var rows = input.GetLength(0);
        var cols = input.GetLength(1);
        if (output.GetLength(0) != rows || output.GetLength(1) != cols)
        {
            throw new ArgumentException("Output dimensions must match input.", nameof(output));
        }
        if (rows == 0 || cols == 0)
        {
            return;
        }

        if (rows < parallelThresholdRows)
        {
            // Serial path: avoid the parallel scheduler entirely on small inputs.
            var inputBuf = new double[cols];
            var outputBuf = new double[cols];
            for (var r = 0; r < rows; r++)
            {
                CopyRow(input, r, inputBuf);
                rowTransform(r, inputBuf, outputBuf);
                WriteRow(output, r, outputBuf);
            }
            return;
        }

        var options = new ParallelOptions
        {
            MaxDegreeOfParallelism = maxDegreeOfParallelism ?? Environment.ProcessorCount,
        };

        // Per-thread row buffers avoid allocating inside the hot loop.
        Parallel.For(
            fromInclusive: 0,
            toExclusive: rows,
            parallelOptions: options,
            localInit: () => (new double[cols], new double[cols]),
            body: (r, _, locals) =>
            {
                var (inBuf, outBuf) = locals;
                CopyRow(input, r, inBuf);
                rowTransform(r, inBuf, outBuf);
                WriteRow(output, r, outBuf);
                return locals;
            },
            localFinally: _ => { });
    }

    private static void CopyRow(double[,] source, int row, double[] destination)
    {
        var cols = source.GetLength(1);
        for (var c = 0; c < cols; c++)
        {
            destination[c] = source[row, c];
        }
    }

    private static void WriteRow(double[,] destination, int row, double[] source)
    {
        var cols = destination.GetLength(1);
        for (var c = 0; c < cols; c++)
        {
            destination[row, c] = source[c];
        }
    }

    /// <summary>
    /// UDF wrapper that exposes the parallel batch transform as a worksheet function for
    /// a useful default workload: element-wise <c>tanh</c>, which is CPU-bound enough to
    /// benefit from MTR. Note that this UDF is NOT registered as <c>IsThreadSafe</c>:
    /// it already drives its own parallelism with <see cref="Parallel.For"/>, so we
    /// don't want Excel to additionally schedule it across MTR worker threads (that
    /// would create N*M threads). Use it once per recalc on a large block.
    /// Marshaling cost: 1 read + 1 write = 2 crossings.
    /// Thread-safety: NOT registered as MTR-safe by design (it forks internally).
    /// </summary>
    [ExcelFunction(
        Name = "EPT.PARALLELTANH",
        Description = "Element-wise tanh using internal Parallel.For; demonstrates explicit parallel batch path.",
        Category = "EPT.Parallel",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static object[,] ParallelTanh(object[,] block)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var input = Marshaling.ToDoubleMatrix(block);
        var output = new double[rows, cols];
        ParallelBatchTransform(input, output, static (_, src, dst) =>
        {
            for (var i = 0; i < src.Length; i++)
            {
                dst[i] = Math.Tanh(src[i]);
            }
        });
        return Marshaling.BoxDoubleMatrix(output);
    }

    /// <summary>
    /// UDF wrapper around the parallel batch path that takes a row-reduction selector
    /// expressed as one of a small set of named operations. This avoids exposing a raw
    /// delegate over the wire while still demonstrating the parallel path on a useful
    /// workload (e.g. per-row standard deviation across a million-row block).
    /// Operations: <c>sum</c>, <c>mean</c>, <c>min</c>, <c>max</c>, <c>stdev</c>.
    /// Marshaling cost: 1 read + 1 write = 2 crossings.
    /// Thread-safety: NOT registered as MTR-safe (drives its own Parallel.For).
    /// </summary>
    [ExcelFunction(
        Name = "EPT.PARALLELROWREDUCE",
        Description = "Per-row reduction (sum|mean|min|max|stdev) using internal Parallel.For.",
        Category = "EPT.Parallel",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static object[,] ParallelRowReduce(object[,] block, string operation)
    {
        ArgumentNullException.ThrowIfNull(block);
        if (string.IsNullOrWhiteSpace(operation))
        {
            return Marshaling.ErrorBlock(ExcelError.ExcelErrorValue);
        }
        var op = operation.Trim().ToLowerInvariant();
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var input = Marshaling.ToDoubleMatrix(block);
        var output = new double[rows, 1];

        RowTransform body = op switch
        {
            "sum" => static (_, src, dst) =>
            {
                var s = 0d;
                foreach (var v in src) { s += v; }
                dst[0] = s;
            },
            "mean" => static (_, src, dst) =>
            {
                var s = 0d;
                foreach (var v in src) { s += v; }
                dst[0] = src.Length == 0 ? 0d : s / src.Length;
            },
            "min" => static (_, src, dst) =>
            {
                if (src.Length == 0) { dst[0] = 0d; return; }
                var m = src[0];
                for (var i = 1; i < src.Length; i++) { if (src[i] < m) { m = src[i]; } }
                dst[0] = m;
            },
            "max" => static (_, src, dst) =>
            {
                if (src.Length == 0) { dst[0] = 0d; return; }
                var m = src[0];
                for (var i = 1; i < src.Length; i++) { if (src[i] > m) { m = src[i]; } }
                dst[0] = m;
            },
            "stdev" => static (_, src, dst) =>
            {
                if (src.Length < 2) { dst[0] = 0d; return; }
                var sum = 0d;
                foreach (var v in src) { sum += v; }
                var mean = sum / src.Length;
                var ss = 0d;
                foreach (var v in src)
                {
                    var d = v - mean;
                    ss += d * d;
                }
                dst[0] = Math.Sqrt(ss / (src.Length - 1));
            },
            _ => throw new ArgumentException($"Unknown operation '{operation}'. Expected one of: sum, mean, min, max, stdev.", nameof(operation)),
        };

        // The reduction output is one column wide; the row buffer in the helper is sized
        // to the input cols, but the writer only touches dst[0..1]. Allocate matching
        // shapes by using a 2-D output of width 1 and an input projection trick:
        // since the helper writes one full row to output, we expand and then take col 0.
        var expanded = new double[rows, cols];
        ParallelBatchTransform(input, expanded, (_, src, dst) =>
        {
            body(0, src, dst);
        });
        var result = new object[rows, 1];
        for (var r = 0; r < rows; r++)
        {
            result[r, 0] = expanded[r, 0];
        }
        return result;
    }

    /// <summary>
    /// Thread-safe UDF that returns the dot product of two equal-length vectors. Inputs
    /// may be column-shaped or row-shaped; both are flattened in row-major order.
    /// Marshaling cost: 2 reads (one per arg) + 1 write = 3 crossings.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.MT.DOT",
        Description = "Thread-safe dot product of two equal-length flattened vectors (MTR-eligible).",
        Category = "EPT.Parallel",
        IsThreadSafe = true,
        IsVolatile = false)]
    public static object Dot(object[,] a, object[,] b)
    {
        ArgumentNullException.ThrowIfNull(a);
        ArgumentNullException.ThrowIfNull(b);
        var ra = a.GetLength(0);
        var ca = a.GetLength(1);
        var rb = b.GetLength(0);
        var cb = b.GetLength(1);
        // long-cast so an int-multiplication overflow can't make a mismatched pair of
        // shapes appear equal under matched wrap.
        var na = (long)ra * ca;
        var nb = (long)rb * cb;
        if (na != nb || na > int.MaxValue)
        {
            return ExcelError.ExcelErrorValue;
        }
        var sum = 0d;
        var ia = 0;
        for (var r = 0; r < ra; r++)
        {
            for (var c = 0; c < ca; c++)
            {
                var br = ia / cb;
                var bc = ia % cb;
                if (Marshaling.TryToDouble(a[r, c], out var va)
                    && Marshaling.TryToDouble(b[br, bc], out var vb))
                {
                    sum += va * vb;
                }
                ia++;
            }
        }
        return sum;
    }
}
