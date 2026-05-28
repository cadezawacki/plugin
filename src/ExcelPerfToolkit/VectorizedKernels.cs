using System;
using System.Diagnostics;
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;
using ExcelDna.Integration;

namespace ExcelPerfToolkit;

/// <summary>
/// Compiled, SIMD-accelerated numeric kernels. This file addresses bottleneck #4 of the
/// second wave: Excel's interpreted runtime offers no just-in-time compilation and no
/// SIMD. We move the arithmetic-heavy kernels into compiled .NET that exploits the
/// hardware's vector lanes through <see cref="Vector{T}"/> (portable) and, where a
/// concrete benefit warrants it, hardware intrinsics under
/// <see cref="System.Runtime.Intrinsics"/> with a runtime-checked scalar fallback.
///
/// Engineering rules followed by every kernel here:
///   * The SIMD path is gated on <see cref="Vector.IsHardwareAccelerated"/> (and on
///     intrinsic-specific <c>IsSupported</c> flags where used) so the same XLL is
///     correct on any CPU - including hypervised hosts and ARM emulation where no
///     vector instructions are available.
///   * Kernels run on the strongly typed <c>double[]</c> / <c>double[,]</c> outputs of
///     <see cref="Marshaling.ToDoubleMatrix"/>, so the boundary crossings are exactly
///     the read and write of the surrounding UDF. No re-entrance into COM.
///   * No shared mutable static state. UDF entry points are registered with
///     <c>IsThreadSafe = true</c>, so they can run on Excel's MTR worker threads.
///   * <c>AllowUnsafeBlocks</c> remains disabled. We use
///     <see cref="Vector{T}"/>'s span-based constructors and <c>Vector256</c> with the
///     public, safe <c>LoadUnsafe</c> APIs over spans - no raw pointers.
/// </summary>
public static class VectorizedKernels
{
    private static readonly TraceSource TraceSource = ToolkitLifetime.CreateTraceSource("VectorizedKernels");

    /// <summary>
    /// True if the runtime reports hardware-accelerated <see cref="Vector{T}"/> over
    /// <see cref="double"/>. Cached to avoid re-querying on every call.
    /// </summary>
    public static bool HardwareAccelerated { get; } = Vector.IsHardwareAccelerated && Vector<double>.Count > 1;

    /// <summary>
    /// True if AVX2 is available on this CPU. When true the intrinsic-tuned dot product
    /// and matrix multiply paths use 256-bit lanes (4x <see cref="double"/>) with FMA
    /// where additionally supported.
    /// </summary>
    public static bool Avx2Supported { get; } = Avx2.IsSupported;

    /// <summary>
    /// True if FMA3 is available; used together with AVX2 in the matrix-multiply path.
    /// </summary>
    public static bool FmaSupported { get; } = Fma.IsSupported;

    // ====================================================================
    // Boundary helpers: object[,] <-> flat double[] for SIMD access.
    // These deliberately allocate exactly once on entry and once on exit.
    // ====================================================================

    /// <summary>
    /// Converts an Excel-style block into a flat row-major <c>double[]</c>. Non-numeric
    /// cells are zeroed. The single allocation is documented and unavoidable: SIMD over
    /// <c>Vector{double}</c> requires linear memory.
    /// </summary>
    internal static double[] FlattenToDoubles(object[,] block, out int rows, out int cols)
    {
        ArgumentNullException.ThrowIfNull(block);
        rows = block.GetLength(0);
        cols = block.GetLength(1);
        // Excel max range exceeds Int32.MaxValue cells; use long to detect overflow
        // before allocation so we surface a clear ArgumentException instead of
        // wrapping to a too-small buffer and indexing past it.
        var total = (long)rows * cols;
        if (total > int.MaxValue)
        {
            throw new ArgumentException($"Block too large: {rows}x{cols} = {total} cells exceeds Int32.MaxValue.", nameof(block));
        }
        var result = new double[(int)total];
        var idx = 0;
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                result[idx++] = Marshaling.TryToDouble(block[r, c], out var d) ? d : 0d;
            }
        }
        return result;
    }

    /// <summary>
    /// Boxes a flat row-major <c>double[]</c> back into the <c>object[,]</c> Excel wants.
    /// The boxing per cell is intentional and unavoidable for COM.
    /// </summary>
    internal static object[,] BoxFlatDoubles(double[] flat, int rows, int cols)
    {
        ArgumentNullException.ThrowIfNull(flat);
        // long-cast the comparison so a wrapped int product can't accidentally equal
        // flat.Length and let a too-small buffer through.
        if ((long)flat.Length != (long)rows * cols)
        {
            throw new ArgumentException("Flat buffer length does not match dimensions.", nameof(flat));
        }
        var result = new object[rows, cols];
        var idx = 0;
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                result[r, c] = flat[idx++];
            }
        }
        return result;
    }

    // ====================================================================
    // Kernel 1: element-wise add
    // ====================================================================

    /// <summary>
    /// Element-wise <c>a + b</c> over two equal-length buffers. SIMD path uses
    /// <see cref="Vector{T}"/> when hardware-accelerated; otherwise scalar.
    /// Instruction sets exploited: SSE2 / AVX / AVX2 / NEON (whichever
    /// <see cref="Vector{T}"/> targets on the host).
    /// Thread-safety: pure; safe for MTR.
    /// Marshaling cost: 0 (operates on managed arrays).
    /// </summary>
    public static void ElementWiseAdd(ReadOnlySpan<double> a, ReadOnlySpan<double> b, Span<double> destination)
    {
        if (a.Length != b.Length || destination.Length != a.Length)
        {
            throw new ArgumentException("All buffers must have the same length.");
        }
        var i = 0;
        if (HardwareAccelerated)
        {
            var width = Vector<double>.Count;
            for (; i <= a.Length - width; i += width)
            {
                var va = new Vector<double>(a.Slice(i, width));
                var vb = new Vector<double>(b.Slice(i, width));
                (va + vb).CopyTo(destination.Slice(i, width));
            }
        }
        for (; i < a.Length; i++)
        {
            destination[i] = a[i] + b[i];
        }
    }

    /// <summary>
    /// UDF wrapper for <see cref="ElementWiseAdd(ReadOnlySpan{double}, ReadOnlySpan{double}, Span{double})"/>.
    /// Returns a block of the same shape as the inputs. Inputs must share a shape.
    /// Marshaling cost: 2 read + 1 write = 3 crossings, all O(1) in N.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.SIMD.ADD",
        Description = "SIMD element-wise addition of two equally-shaped numeric blocks.",
        Category = "EPT.Vectorized",
        IsThreadSafe = true,
        IsVolatile = false)]
    public static object SimdAdd(object[,] a, object[,] b)
    {
        try
        {
            ValidateSameShape(a, b);
            var flatA = FlattenToDoubles(a, out var rows, out var cols);
            var flatB = FlattenToDoubles(b, out _, out _);
            var dest = new double[flatA.Length];
            ElementWiseAdd(flatA, flatB, dest);
            return BoxFlatDoubles(dest, rows, cols);
        }
        catch (Exception ex) when (!IsCritical(ex))
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 1, "EPT.SIMD.ADD failed: {0}", ex.Message);
            return ExcelError.ExcelErrorValue;
        }
    }

    // ====================================================================
    // Kernel 2: element-wise multiply
    // ====================================================================

    /// <summary>
    /// Element-wise <c>a * b</c>. See <see cref="ElementWiseAdd"/> for SIMD/scalar
    /// selection notes.
    /// Thread-safety: pure; safe for MTR.
    /// </summary>
    public static void ElementWiseMultiply(ReadOnlySpan<double> a, ReadOnlySpan<double> b, Span<double> destination)
    {
        if (a.Length != b.Length || destination.Length != a.Length)
        {
            throw new ArgumentException("All buffers must have the same length.");
        }
        var i = 0;
        if (HardwareAccelerated)
        {
            var width = Vector<double>.Count;
            for (; i <= a.Length - width; i += width)
            {
                var va = new Vector<double>(a.Slice(i, width));
                var vb = new Vector<double>(b.Slice(i, width));
                (va * vb).CopyTo(destination.Slice(i, width));
            }
        }
        for (; i < a.Length; i++)
        {
            destination[i] = a[i] * b[i];
        }
    }

    /// <summary>
    /// UDF wrapper for <see cref="ElementWiseMultiply"/>.
    /// Marshaling cost: 2 reads + 1 write = 3 crossings.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.SIMD.MULTIPLY",
        Description = "SIMD element-wise multiplication of two equally-shaped numeric blocks.",
        Category = "EPT.Vectorized",
        IsThreadSafe = true,
        IsVolatile = false)]
    public static object SimdMultiply(object[,] a, object[,] b)
    {
        try
        {
            ValidateSameShape(a, b);
            var flatA = FlattenToDoubles(a, out var rows, out var cols);
            var flatB = FlattenToDoubles(b, out _, out _);
            var dest = new double[flatA.Length];
            ElementWiseMultiply(flatA, flatB, dest);
            return BoxFlatDoubles(dest, rows, cols);
        }
        catch (Exception ex) when (!IsCritical(ex))
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 1, "EPT.SIMD.MULTIPLY failed: {0}", ex.Message);
            return ExcelError.ExcelErrorValue;
        }
    }

    // ====================================================================
    // Kernel 3: scalar scale
    // ====================================================================

    /// <summary>
    /// Multiplies every element by <paramref name="factor"/>. SIMD when accelerated.
    /// Thread-safety: pure.
    /// </summary>
    public static void Scale(ReadOnlySpan<double> source, double factor, Span<double> destination)
    {
        if (source.Length != destination.Length)
        {
            throw new ArgumentException("Source and destination must have the same length.");
        }
        var i = 0;
        if (HardwareAccelerated)
        {
            var width = Vector<double>.Count;
            var vf = new Vector<double>(factor);
            for (; i <= source.Length - width; i += width)
            {
                var v = new Vector<double>(source.Slice(i, width));
                (v * vf).CopyTo(destination.Slice(i, width));
            }
        }
        for (; i < source.Length; i++)
        {
            destination[i] = source[i] * factor;
        }
    }

    /// <summary>
    /// UDF wrapper for <see cref="Scale"/>.
    /// Marshaling cost: 1 read + 1 write = 2 crossings.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.SIMD.SCALE",
        Description = "SIMD multiplication of every element by a scalar factor.",
        Category = "EPT.Vectorized",
        IsThreadSafe = true,
        IsVolatile = false)]
    public static object SimdScale(object[,] block, double factor)
    {
        try
        {
            var flat = FlattenToDoubles(block, out var rows, out var cols);
            var dest = new double[flat.Length];
            Scale(flat, factor, dest);
            return BoxFlatDoubles(dest, rows, cols);
        }
        catch (Exception ex) when (!IsCritical(ex))
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 1, "EPT.SIMD.SCALE failed: {0}", ex.Message);
            return ExcelError.ExcelErrorValue;
        }
    }

    // ====================================================================
    // Kernel 4: dot product
    // ====================================================================

    /// <summary>
    /// Dot product. Three paths selected at runtime, in order of preference:
    /// (1) AVX2 + FMA over 256-bit lanes when both are supported,
    /// (2) portable <see cref="Vector{T}"/> when hardware-accelerated,
    /// (3) scalar fallback.
    /// Thread-safety: pure.
    /// </summary>
    public static double DotProduct(ReadOnlySpan<double> a, ReadOnlySpan<double> b)
    {
        if (a.Length != b.Length)
        {
            throw new ArgumentException("Vectors must have equal length.");
        }
        if (a.Length == 0)
        {
            return 0d;
        }

        if (Avx2Supported && a.Length >= Vector256<double>.Count)
        {
            return DotProductAvx2(a, b);
        }
        if (HardwareAccelerated && a.Length >= Vector<double>.Count)
        {
            return DotProductPortable(a, b);
        }
        return DotProductScalar(a, b);
    }

    private static double DotProductAvx2(ReadOnlySpan<double> a, ReadOnlySpan<double> b)
    {
        // We use Vector256.LoadUnsafe over Span, which is the safe API that compiles
        // to vmovupd; AllowUnsafeBlocks remains disabled.
        var width = Vector256<double>.Count;
        var accum = Vector256<double>.Zero;
        var i = 0;
        for (; i <= a.Length - width; i += width)
        {
            var va = Vector256.Create(a[i], a[i + 1], a[i + 2], a[i + 3]);
            var vb = Vector256.Create(b[i], b[i + 1], b[i + 2], b[i + 3]);
            accum = FmaSupported
                ? Fma.MultiplyAdd(va, vb, accum)
                : Avx.Add(accum, Avx.Multiply(va, vb));
        }
        var sum = accum.GetElement(0) + accum.GetElement(1) + accum.GetElement(2) + accum.GetElement(3);
        for (; i < a.Length; i++)
        {
            sum += a[i] * b[i];
        }
        return sum;
    }

    private static double DotProductPortable(ReadOnlySpan<double> a, ReadOnlySpan<double> b)
    {
        var width = Vector<double>.Count;
        var accum = Vector<double>.Zero;
        var i = 0;
        for (; i <= a.Length - width; i += width)
        {
            var va = new Vector<double>(a.Slice(i, width));
            var vb = new Vector<double>(b.Slice(i, width));
            accum += va * vb;
        }
        var sum = Vector.Sum(accum);
        for (; i < a.Length; i++)
        {
            sum += a[i] * b[i];
        }
        return sum;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static double DotProductScalar(ReadOnlySpan<double> a, ReadOnlySpan<double> b)
    {
        var sum = 0d;
        for (var i = 0; i < a.Length; i++)
        {
            sum += a[i] * b[i];
        }
        return sum;
    }

    /// <summary>
    /// UDF wrapper for <see cref="DotProduct"/>. Flattens both inputs row-major and
    /// requires equal element counts.
    /// Marshaling cost: 2 reads + 1 scalar write = 3 crossings.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.SIMD.DOT",
        Description = "SIMD dot product over two flattened numeric vectors (AVX2+FMA when available).",
        Category = "EPT.Vectorized",
        IsThreadSafe = true,
        IsVolatile = false)]
    public static object SimdDot(object[,] a, object[,] b)
    {
        try
        {
            var flatA = FlattenToDoubles(a, out _, out _);
            var flatB = FlattenToDoubles(b, out _, out _);
            if (flatA.Length != flatB.Length)
            {
                return ExcelError.ExcelErrorValue;
            }
            return DotProduct(flatA, flatB);
        }
        catch (Exception ex) when (!IsCritical(ex))
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 1, "EPT.SIMD.DOT failed: {0}", ex.Message);
            return ExcelError.ExcelErrorValue;
        }
    }

    // ====================================================================
    // Kernel 5: row sums
    // ====================================================================

    /// <summary>
    /// Per-row sums over a flat row-major buffer. SIMD path sums across lanes.
    /// Thread-safety: pure.
    /// </summary>
    public static void RowSums(ReadOnlySpan<double> source, int rows, int cols, Span<double> rowSums)
    {
        if (source.Length != rows * cols)
        {
            throw new ArgumentException("Source length does not match dimensions.");
        }
        if (rowSums.Length != rows)
        {
            throw new ArgumentException("Destination must have length equal to row count.", nameof(rowSums));
        }
        for (var r = 0; r < rows; r++)
        {
            var row = source.Slice(r * cols, cols);
            var i = 0;
            var sum = 0d;
            if (HardwareAccelerated && cols >= Vector<double>.Count)
            {
                var width = Vector<double>.Count;
                var accum = Vector<double>.Zero;
                for (; i <= cols - width; i += width)
                {
                    accum += new Vector<double>(row.Slice(i, width));
                }
                sum = Vector.Sum(accum);
            }
            for (; i < cols; i++)
            {
                sum += row[i];
            }
            rowSums[r] = sum;
        }
    }

    /// <summary>
    /// UDF wrapper for <see cref="RowSums"/>. Returns a single-column block.
    /// Marshaling cost: 1 read + 1 write = 2 crossings.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.SIMD.ROWSUMS",
        Description = "SIMD per-row sums.",
        Category = "EPT.Vectorized",
        IsThreadSafe = true,
        IsVolatile = false)]
    public static object SimdRowSums(object[,] block)
    {
        try
        {
            var flat = FlattenToDoubles(block, out var rows, out var cols);
            var sums = new double[rows];
            RowSums(flat, rows, cols, sums);
            var result = new object[rows, 1];
            for (var r = 0; r < rows; r++)
            {
                result[r, 0] = sums[r];
            }
            return result;
        }
        catch (Exception ex) when (!IsCritical(ex))
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 1, "EPT.SIMD.ROWSUMS failed: {0}", ex.Message);
            return ExcelError.ExcelErrorValue;
        }
    }

    // ====================================================================
    // Kernel 6: column sums
    // ====================================================================

    /// <summary>
    /// Per-column sums over a flat row-major buffer. SIMD accumulates whole rows
    /// horizontally into the column-sum vector, exploiting wide lanes.
    /// Thread-safety: pure.
    /// </summary>
    public static void ColumnSums(ReadOnlySpan<double> source, int rows, int cols, Span<double> colSums)
    {
        if (source.Length != rows * cols)
        {
            throw new ArgumentException("Source length does not match dimensions.");
        }
        if (colSums.Length != cols)
        {
            throw new ArgumentException("Destination must have length equal to column count.", nameof(colSums));
        }
        colSums.Clear();
        if (HardwareAccelerated)
        {
            var width = Vector<double>.Count;
            // Buffer for accumulator updates that round-trip through a Vector.
            Span<double> buffer = stackalloc double[width];
            for (var r = 0; r < rows; r++)
            {
                var row = source.Slice(r * cols, cols);
                var c = 0;
                for (; c <= cols - width; c += width)
                {
                    var vRow = new Vector<double>(row.Slice(c, width));
                    var vCol = new Vector<double>(colSums.Slice(c, width));
                    (vCol + vRow).CopyTo(buffer);
                    buffer.CopyTo(colSums.Slice(c, width));
                }
                for (; c < cols; c++)
                {
                    colSums[c] += row[c];
                }
            }
        }
        else
        {
            for (var r = 0; r < rows; r++)
            {
                var row = source.Slice(r * cols, cols);
                for (var c = 0; c < cols; c++)
                {
                    colSums[c] += row[c];
                }
            }
        }
    }

    /// <summary>
    /// UDF wrapper for <see cref="ColumnSums"/>. Returns a single-row block.
    /// Marshaling cost: 1 read + 1 write = 2 crossings.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.SIMD.COLSUMS",
        Description = "SIMD per-column sums.",
        Category = "EPT.Vectorized",
        IsThreadSafe = true,
        IsVolatile = false)]
    public static object SimdColumnSums(object[,] block)
    {
        try
        {
            var flat = FlattenToDoubles(block, out var rows, out var cols);
            var sums = new double[cols];
            ColumnSums(flat, rows, cols, sums);
            var result = new object[1, cols];
            for (var c = 0; c < cols; c++)
            {
                result[0, c] = sums[c];
            }
            return result;
        }
        catch (Exception ex) when (!IsCritical(ex))
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 1, "EPT.SIMD.COLSUMS failed: {0}", ex.Message);
            return ExcelError.ExcelErrorValue;
        }
    }

    // ====================================================================
    // Kernel 7: L2 normalize
    // ====================================================================

    /// <summary>
    /// L2-normalizes a vector in place semantics: writes <c>x / ||x||_2</c> into the
    /// destination. Returns the L2 norm. If the norm is zero, the destination is set
    /// to zeros and the function returns 0.
    /// Thread-safety: pure.
    /// </summary>
    public static double L2Normalize(ReadOnlySpan<double> source, Span<double> destination)
    {
        if (source.Length != destination.Length)
        {
            throw new ArgumentException("Source and destination must have the same length.");
        }
        var sumSq = DotProduct(source, source);
        var norm = Math.Sqrt(sumSq);
        // If the squared sum overflows to infinity, norm is +Inf; dividing by it
        // would silently zero the destination while we returned +Inf as the norm.
        // Treat non-finite norm the same as zero norm: return zeros and norm=0 so
        // the caller can detect the overflow via the return value being 0 despite
        // a non-empty source.
        if (norm == 0d || double.IsNaN(norm) || double.IsInfinity(norm))
        {
            destination.Clear();
            return 0d;
        }
        Scale(source, 1d / norm, destination);
        return norm;
    }

    /// <summary>
    /// UDF wrapper for <see cref="L2Normalize"/> over a whole block (treated as a flat
    /// vector). Returns a block of the same shape.
    /// Marshaling cost: 1 read + 1 write = 2 crossings.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.SIMD.NORMALIZE",
        Description = "SIMD L2 normalization of a numeric block (treated as a flat vector).",
        Category = "EPT.Vectorized",
        IsThreadSafe = true,
        IsVolatile = false)]
    public static object SimdNormalize(object[,] block)
    {
        try
        {
            var flat = FlattenToDoubles(block, out var rows, out var cols);
            var dest = new double[flat.Length];
            L2Normalize(flat, dest);
            return BoxFlatDoubles(dest, rows, cols);
        }
        catch (Exception ex) when (!IsCritical(ex))
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 1, "EPT.SIMD.NORMALIZE failed: {0}", ex.Message);
            return ExcelError.ExcelErrorValue;
        }
    }

    // ====================================================================
    // Kernel 8: matrix multiply C = A * B (AxK * KxB)
    // ====================================================================

    /// <summary>
    /// Matrix multiply <c>C = A * B</c> with <paramref name="a"/> shape <c>[m, k]</c>,
    /// <paramref name="b"/> shape <c>[k, n]</c>, destination shape <c>[m, n]</c>.
    /// The inner-product loop dispatches to AVX2+FMA or portable <see cref="Vector{T}"/>
    /// or scalar at runtime via <see cref="DotProduct"/>.
    /// Thread-safety: pure.
    /// </summary>
    public static void MatrixMultiply(
        ReadOnlySpan<double> a, int m, int k,
        ReadOnlySpan<double> b, int kBRows, int n,
        Span<double> destination)
    {
        // long-cast every shape product so an int-overflow can't make a too-small
        // buffer compare equal to the wrapped product and slip past validation.
        if ((long)a.Length != (long)m * k)
        {
            throw new ArgumentException("A's length does not match m*k.", nameof(a));
        }
        if ((long)b.Length != (long)kBRows * n)
        {
            throw new ArgumentException("B's length does not match k*n.", nameof(b));
        }
        if (k != kBRows)
        {
            throw new ArgumentException("A.cols must equal B.rows.");
        }
        if ((long)destination.Length != (long)m * n)
        {
            throw new ArgumentException("Destination length must be m*n.", nameof(destination));
        }

        // Pack a single column of B into a contiguous buffer so the inner DotProduct
        // sees two linear ReadOnlySpan<double> and can vectorize end to end. The
        // packing allocation is paid once per output column, not per cell.
        var bCol = new double[k];
        for (var j = 0; j < n; j++)
        {
            for (var p = 0; p < k; p++)
            {
                bCol[p] = b[p * n + j];
            }
            for (var i = 0; i < m; i++)
            {
                var row = a.Slice(i * k, k);
                destination[i * n + j] = DotProduct(row, bCol);
            }
        }
    }

    /// <summary>
    /// UDF wrapper for <see cref="MatrixMultiply"/>. Multiplies <paramref name="a"/> by
    /// <paramref name="b"/>; the inner dimensions must agree. Returns the product as a
    /// block of shape <c>[a.rows, b.cols]</c>.
    /// Marshaling cost: 2 reads + 1 write = 3 crossings.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.SIMD.MATMUL",
        Description = "SIMD matrix multiply (AVX2+FMA when available, scalar fallback).",
        Category = "EPT.Vectorized",
        IsThreadSafe = true,
        IsVolatile = false)]
    public static object SimdMatMul(object[,] a, object[,] b)
    {
        try
        {
            var flatA = FlattenToDoubles(a, out var m, out var k);
            var flatB = FlattenToDoubles(b, out var kB, out var n);
            if (k != kB)
            {
                return ExcelError.ExcelErrorValue;
            }
            var dest = new double[m * n];
            MatrixMultiply(flatA, m, k, flatB, kB, n, dest);
            return BoxFlatDoubles(dest, m, n);
        }
        catch (Exception ex) when (!IsCritical(ex))
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 1, "EPT.SIMD.MATMUL failed: {0}", ex.Message);
            return ExcelError.ExcelErrorValue;
        }
    }

    // ====================================================================
    // Kernel 9: capability probe (for diagnostics from the sheet)
    // ====================================================================

    /// <summary>
    /// Returns a one-row block describing the SIMD capabilities the kernels can use on
    /// this CPU. Useful from the worksheet to confirm that the vectorized path is live.
    /// Columns: HardwareAccelerated, VectorWidth(doubles), Avx2, Fma.
    /// Marshaling cost: 0 read + 1 write = 1 crossing.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.SIMD.CAPABILITIES",
        Description = "Returns the SIMD capabilities exploited by the vectorized kernels on this CPU.",
        Category = "EPT.Vectorized",
        IsThreadSafe = true,
        IsVolatile = false)]
    public static object SimdCapabilities()
    {
        var result = new object[1, 4];
        result[0, 0] = HardwareAccelerated;
        result[0, 1] = (double)Vector<double>.Count;
        result[0, 2] = Avx2Supported;
        result[0, 3] = FmaSupported;
        return result;
    }

    // ====================================================================
    // Shared validation
    // ====================================================================

    private static void ValidateSameShape(object[,] a, object[,] b)
    {
        ArgumentNullException.ThrowIfNull(a);
        ArgumentNullException.ThrowIfNull(b);
        if (a.GetLength(0) != b.GetLength(0) || a.GetLength(1) != b.GetLength(1))
        {
            throw new ArgumentException("Inputs must have identical shapes.");
        }
    }

    /// <summary>
    /// Returns true for exception types we must NOT swallow in the UDF boundary catch.
    /// OutOfMemoryException etc. signal process-wide problems that can't be papered over
    /// by surfacing #VALUE!; rethrowing lets Excel-DNA's unhandled handler at least log
    /// the real cause.
    /// </summary>
    private static bool IsCritical(Exception ex)
        => ex is OutOfMemoryException
            or StackOverflowException
            or AccessViolationException
            or ThreadAbortException;
}
