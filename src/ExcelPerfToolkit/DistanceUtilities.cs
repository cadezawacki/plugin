using System;
using ExcelDna.Integration;

namespace ExcelPerfToolkit;

/// <summary>
/// Round 4 pairwise distance / similarity. <see cref="Distance"/> treats each row of its
/// input as an observation vector and returns the full matrix of distances between every row
/// of <c>matrix_a</c> and every row of <c>matrix_b</c> (or of <c>matrix_a</c> against itself).
///
/// <para>The Euclidean and cosine paths are expressed through inner products and routed
/// through <see cref="VectorizedKernels.DotProduct"/>, so they ride the AVX2/FMA or portable
/// <c>Vector&lt;double&gt;</c> kernels (bottleneck #4) - Euclidean uses the identity
/// <c>||a-b||^2 = a.a + b.b - 2 a.b</c> with the self-dot products precomputed once per row,
/// falling back to a direct difference pass for any pair where the identity would cancel
/// catastrophically. The function is pure CPU and registered <c>IsThreadSafe = true</c>
/// (bottleneck #3).</para>
/// </summary>
public static class DistanceUtilities
{
    /// <summary>
    /// Returns the <c>rows(A) x rows(B)</c> matrix of distances between observation vectors.
    /// When <paramref name="matrixB"/> is omitted, <paramref name="matrixA"/> is compared
    /// against itself (computing one triangle and mirroring it; the diagonal is exactly 0).
    /// The feature dimension (column count) of the two inputs must match.
    /// Non-numeric cells are treated as 0. An error scalar passed as
    /// <paramref name="matrixB"/> propagates instead of being treated as omitted.
    ///
    /// <para><paramref name="metric"/> is one of <c>euclidean</c> (default), <c>cosine</c>
    /// (returns 1 - cosine similarity), <c>manhattan</c>, or <c>chebyshev</c>.</para>
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.DISTANCE", Description = "Pairwise distance matrix between row vectors (euclidean|cosine|manhattan|chebyshev).", Category = "EPT.Distance", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] Distance(
        [ExcelArgument(Name = "matrix_a", Description = "Rows are observation vectors.")] object[,] matrixA,
        [ExcelArgument(Name = "matrix_b", Description = "Optional second set of vectors; defaults to matrix_a.")] object matrixB,
        [ExcelArgument(Name = "metric", Description = "euclidean (default) | cosine | manhattan | chebyshev")] string metric = "euclidean")
    {
        ArgumentNullException.ThrowIfNull(matrixA);
        if (matrixB is ExcelError bErr)
        {
            // An error scalar here is a broken upstream reference, not "omitted";
            // silently computing a self-distance matrix would mask it.
            return Marshaling.ErrorBlock(bErr);
        }
        var bBlock = Marshaling.IsBlankOrError(matrixB) ? matrixA : Marshaling.AsArray2D(matrixB);
        var met = string.IsNullOrWhiteSpace(metric) ? "euclidean" : metric.Trim().ToLowerInvariant();
        var self = ReferenceEquals(bBlock, matrixA);

        var flatA = Flatten(matrixA, out var rowsA, out var dim);
        double[] flatB;
        int rowsB, dimB;
        if (self)
        {
            // Same matrix: reuse the flat buffer instead of duplicating the copy and
            // re-parsing every cell.
            flatB = flatA;
            rowsB = rowsA;
            dimB = dim;
        }
        else
        {
            flatB = Flatten(bBlock, out rowsB, out dimB);
        }
        if (dim != dimB)
        {
            throw new ArgumentException($"matrix_a has {dim} columns but matrix_b has {dimB}; feature dimensions must match.");
        }
        if ((long)rowsA * rowsB > int.MaxValue)
        {
            throw new ArgumentException($"Result {rowsA}x{rowsB} exceeds Int32.MaxValue cells.");
        }

        var result = new object[rowsA, rowsB];
        if (dim == 0)
        {
            for (var i = 0; i < rowsA; i++)
            {
                for (var j = 0; j < rowsB; j++)
                {
                    result[i, j] = 0d;
                }
            }
            return result;
        }

        switch (met)
        {
            case "euclidean":
            {
                var selfA = SelfDots(flatA, rowsA, dim);
                var selfB = self ? selfA : SelfDots(flatB, rowsB, dim);
                for (var i = 0; i < rowsA; i++)
                {
                    var rowA = flatA.AsSpan(i * dim, dim);
                    for (var j = self ? i : 0; j < rowsB; j++)
                    {
                        if (self && j == i)
                        {
                            result[i, i] = 0d;
                            continue;
                        }
                        var dot = VectorizedKernels.DotProduct(rowA, flatB.AsSpan(j * dim, dim));
                        var sq = selfA[i] + selfB[j] - (2d * dot);
                        // The identity cancels catastrophically when the distance is
                        // tiny relative to the magnitudes (verified: rows of (1e8) and
                        // (1e8+1) come back as 0 instead of 1). Below ~1e-8 of the
                        // combined self-dots fewer than half the significant bits
                        // survive; recompute that pair directly. The slow path fires
                        // only for near-duplicate pairs.
                        if (sq <= 1e-8 * (selfA[i] + selfB[j]))
                        {
                            var rowB = flatB.AsSpan(j * dim, dim);
                            sq = 0d;
                            for (var k = 0; k < dim; k++)
                            {
                                var diff = rowA[k] - rowB[k];
                                sq += diff * diff;
                            }
                        }
                        var value = Math.Sqrt(sq > 0d ? sq : 0d);
                        if (self)
                        {
                            var boxed = (object)value;
                            result[i, j] = boxed;
                            result[j, i] = boxed;
                        }
                        else
                        {
                            result[i, j] = value;
                        }
                    }
                }
                break;
            }
            case "cosine":
            {
                var selfA = SelfDots(flatA, rowsA, dim);
                var selfB = self ? selfA : SelfDots(flatB, rowsB, dim);
                for (var i = 0; i < rowsA; i++)
                {
                    var rowA = flatA.AsSpan(i * dim, dim);
                    for (var j = self ? i : 0; j < rowsB; j++)
                    {
                        if (self && j == i)
                        {
                            // sqrt(s)*sqrt(s) != s for about half of all magnitudes, so
                            // computing the diagonal leaves +/-1e-16 residue; it is 0 by
                            // definition.
                            result[i, i] = 0d;
                            continue;
                        }
                        var dot = VectorizedKernels.DotProduct(rowA, flatB.AsSpan(j * dim, dim));
                        var denom = Math.Sqrt(selfA[i]) * Math.Sqrt(selfB[j]);
                        object boxed;
                        if (denom == 0d)
                        {
                            boxed = selfA[i] == 0d && selfB[j] == 0d ? 0d : 1d;
                        }
                        else
                        {
                            boxed = 1d - (dot / denom);
                        }
                        result[i, j] = boxed;
                        if (self)
                        {
                            result[j, i] = boxed;
                        }
                    }
                }
                break;
            }
            case "manhattan":
                ElementwiseMetric(flatA, rowsA, flatB, rowsB, dim, result, manhattan: true, self);
                break;
            case "chebyshev":
                ElementwiseMetric(flatA, rowsA, flatB, rowsB, dim, result, manhattan: false, self);
                break;
            default:
                throw new ArgumentException("metric must be one of: euclidean, cosine, manhattan, chebyshev.", nameof(metric));
        }
        return result;
    }

    private static void ElementwiseMetric(double[] flatA, int rowsA, double[] flatB, int rowsB, int dim, object[,] result, bool manhattan, bool self)
    {
        for (var i = 0; i < rowsA; i++)
        {
            var baseA = i * dim;
            for (var j = self ? i : 0; j < rowsB; j++)
            {
                if (self && j == i)
                {
                    result[i, i] = 0d;
                    continue;
                }
                var baseB = j * dim;
                var acc = 0d;
                // Two specialized loops keep the metric branch out of the inner loop.
                if (manhattan)
                {
                    for (var k = 0; k < dim; k++)
                    {
                        acc += Math.Abs(flatA[baseA + k] - flatB[baseB + k]);
                    }
                }
                else
                {
                    for (var k = 0; k < dim; k++)
                    {
                        var diff = Math.Abs(flatA[baseA + k] - flatB[baseB + k]);
                        if (diff > acc)
                        {
                            acc = diff;
                        }
                    }
                }
                if (self)
                {
                    var boxed = (object)acc;
                    result[i, j] = boxed;
                    result[j, i] = boxed;
                }
                else
                {
                    result[i, j] = acc;
                }
            }
        }
    }

    private static double[] SelfDots(double[] flat, int rows, int dim)
    {
        var self = new double[rows];
        for (var i = 0; i < rows; i++)
        {
            var row = flat.AsSpan(i * dim, dim);
            self[i] = VectorizedKernels.DotProduct(row, row);
        }
        return self;
    }

    private static double[] Flatten(object[,] block, out int rows, out int cols)
    {
        rows = block.GetLength(0);
        cols = block.GetLength(1);
        var total = (long)rows * cols;
        if (total > int.MaxValue)
        {
            throw new ArgumentException($"Block too large: {rows}x{cols} = {total} cells exceeds Int32.MaxValue.");
        }
        var flat = new double[(int)total];
        var idx = 0;
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                flat[idx++] = Marshaling.TryToDouble(block[r, c], out var d) ? d : 0d;
            }
        }
        return flat;
    }
}
