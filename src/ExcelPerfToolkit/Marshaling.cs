using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;
using ExcelDna.Integration;

namespace ExcelPerfToolkit;

/// <summary>
/// Centralized, allocation-conscious marshaling helpers used by the entire toolkit.
/// Everything that converts between Excel's <c>object[,]</c> wire format and managed
/// representations lives here so callers benefit from a single fast path.
///
/// Bottlenecks addressed: #2 (COM per-cell marshaling) by amortizing classification and
/// conversion across whole blocks, and #1 (cell-by-cell read/write) indirectly by making
/// it cheap to operate on the bulk arrays produced by <see cref="BulkTransfer"/>.
/// </summary>
public static class Marshaling
{
    /// <summary>
    /// Excel returns a single-cell read as the bare scalar, not as a 1x1 array. This
    /// helper normalizes any value returned from <see cref="ExcelReference.GetValue"/> to
    /// a dense <c>object[,]</c> for uniform downstream handling.
    /// Marshaling cost: 0 boundary crossings.
    /// Thread-safety: pure; safe in any context.
    /// Accepts: anything Excel-DNA can hand back. Returns: <c>object[,]</c>.
    /// </summary>
    public static object[,] AsArray2D(object value)
    {
        if (value is object[,] block)
        {
            return block;
        }

        var single = new object[1, 1];
        single[0, 0] = value;
        return single;
    }

    /// <summary>
    /// True if the value is one of Excel's sentinel cell states: <see cref="ExcelEmpty"/>,
    /// <see cref="ExcelMissing"/>, an <see cref="ExcelError"/>, or a CLR null. Lets
    /// callers skip cells they cannot meaningfully transform.
    /// Marshaling cost: 0 boundary crossings.
    /// Thread-safety: pure.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool IsBlankOrError(object? value)
    {
        return value is null
            || value is ExcelEmpty
            || value is ExcelMissing
            || value is ExcelError;
    }

    /// <summary>
    /// True if the value carries an <see cref="ExcelError"/> sentinel produced by Excel
    /// (e.g. <c>#N/A</c>, <c>#DIV/0!</c>, <c>#VALUE!</c>).
    /// Marshaling cost: 0 boundary crossings.
    /// Thread-safety: pure.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool IsExcelError(object? value) => value is ExcelError;

    /// <summary>
    /// Best-effort conversion of an Excel cell value to a finite <see cref="double"/>.
    /// Recognizes numeric scalars, booleans, numeric-looking strings, and dates as serials.
    /// <see cref="ExcelEmpty"/>, <see cref="ExcelMissing"/>, errors, and unparseable strings
    /// produce <see langword="false"/>.
    /// Marshaling cost: 0 boundary crossings.
    /// Thread-safety: pure.
    /// </summary>
    public static bool TryToDouble(object? value, out double result)
    {
        switch (value)
        {
            case null:
            case ExcelEmpty:
            case ExcelMissing:
            case ExcelError:
                result = 0d;
                return false;
            case double d:
                result = d;
                return !double.IsNaN(d) && !double.IsInfinity(d);
            case int i:
                result = i;
                return true;
            case long l:
                result = l;
                return true;
            case float f:
                result = f;
                return !float.IsNaN(f) && !float.IsInfinity(f);
            case decimal m:
                result = (double)m;
                return true;
            case bool b:
                result = b ? 1d : 0d;
                return true;
            case DateTime dt:
                // DateTime.ToOADate throws OverflowException for years 1-99 (excluding MinValue).
                // TryToDouble must not let that escape: contract is "best-effort, no throws".
                try
                {
                    result = dt.ToOADate();
                    return true;
                }
                catch (OverflowException)
                {
                    result = 0d;
                    return false;
                }
            case string s:
                // Invariant culture uses ',' as its NumberGroupSeparator. Combined with
                // NumberStyles.AllowThousands this would parse "1,5" as 15 (treating ',' as
                // a thousands separator), silently 10x-ing any comma-decimal cell from
                // de-DE / fr-FR / es-ES workbooks. We must NOT allow thousands separators
                // in the invariant pass; the current-culture fallback may still use them.
                // We also cap the length: TryParse is O(n) and would otherwise scan
                // multi-MB cell strings twice for no benefit.
                if (s.Length > 64)
                {
                    result = 0d;
                    return false;
                }
                // On .NET Core 3.0+, double.TryParse succeeds for "NaN", "Infinity", and
                // out-of-range literals like "1e999" (returning ±Infinity). The contract
                // here is "finite double", so re-check finiteness after either parse —
                // a single non-finite text cell must not poison sort-based consumers
                // (QUANTILES, OUTLIERS, DISTANCE).
                return (double.TryParse(
                    s,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out result)
                    || double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out result))
                    && double.IsFinite(result);
            default:
                result = 0d;
                return false;
        }
    }

    /// <summary>
    /// Best-effort conversion of an Excel cell value to a string. Numbers use the invariant
    /// round-trip form. <see cref="ExcelEmpty"/>/<see cref="ExcelMissing"/> become the empty
    /// string. Errors become Excel's textual form (e.g. <c>#N/A</c>).
    /// Marshaling cost: 0 boundary crossings.
    /// Thread-safety: pure.
    /// </summary>
    public static string ToStringSafe(object? value)
    {
        return value switch
        {
            null => string.Empty,
            ExcelEmpty => string.Empty,
            ExcelMissing => string.Empty,
            ExcelError err => ErrorToText(err),
            string s => s,
            double d => d.ToString("R", CultureInfo.InvariantCulture),
            int i => i.ToString(CultureInfo.InvariantCulture),
            long l => l.ToString(CultureInfo.InvariantCulture),
            float f => f.ToString("R", CultureInfo.InvariantCulture),
            decimal m => m.ToString(CultureInfo.InvariantCulture),
            bool b => b ? "TRUE" : "FALSE",
            DateTime dt => dt.ToString("o", CultureInfo.InvariantCulture),
            // Convert.ToString(object, IFormatProvider) ignores the provider for types
            // that don't implement IConvertible (it falls through to value.ToString()).
            // Call IConvertible.ToString(provider) explicitly when we can; otherwise
            // accept the type's own ToString().
            IConvertible c => c.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            _ => value.ToString() ?? string.Empty,
        };
    }

    /// <summary>
    /// Renders an <see cref="ExcelError"/> as the textual symbol Excel uses on the grid.
    /// </summary>
    public static string ErrorToText(ExcelError err) => err switch
    {
        ExcelError.ExcelErrorDiv0 => "#DIV/0!",
        ExcelError.ExcelErrorNA => "#N/A",
        ExcelError.ExcelErrorName => "#NAME?",
        ExcelError.ExcelErrorNull => "#NULL!",
        ExcelError.ExcelErrorNum => "#NUM!",
        ExcelError.ExcelErrorRef => "#REF!",
        ExcelError.ExcelErrorValue => "#VALUE!",
        ExcelError.ExcelErrorGettingData => "#GETTING_DATA",
        _ => "#ERR",
    };

    /// <summary>
    /// True if every cell in <paramref name="block"/> is a <see cref="double"/>. Lets bulk
    /// code take a non-boxing fast path when Excel hands back a homogeneous numeric range.
    /// Marshaling cost: 0 boundary crossings.
    /// Thread-safety: pure.
    /// </summary>
    public static bool IsAllDouble(object[,] block)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                if (block[r, c] is not double)
                {
                    return false;
                }
            }
        }
        return true;
    }

    /// <summary>
    /// Returns a new <c>double[,]</c> copy of <paramref name="block"/>. Cells that cannot
    /// be interpreted as numbers become <paramref name="defaultValue"/>.
    /// Marshaling cost: 0 boundary crossings.
    /// Thread-safety: pure.
    /// </summary>
    public static double[,] ToDoubleMatrix(object[,] block, double defaultValue = 0d)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var result = new double[rows, cols];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                result[r, c] = TryToDouble(block[r, c], out var d) ? d : defaultValue;
            }
        }
        return result;
    }

    /// <summary>
    /// Returns a new <c>string[,]</c> copy of <paramref name="block"/> using
    /// <see cref="ToStringSafe"/> on every cell.
    /// Marshaling cost: 0 boundary crossings.
    /// Thread-safety: pure.
    /// </summary>
    public static string[,] ToStringMatrix(object[,] block)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var result = new string[rows, cols];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                result[r, c] = ToStringSafe(block[r, c]);
            }
        }
        return result;
    }

    /// <summary>
    /// Boxes a <c>double[,]</c> back into an <c>object[,]</c>. The single deliberate
    /// allocation per cell is unavoidable: Excel-DNA requires <c>object</c> on the wire.
    /// Marshaling cost: 0 boundary crossings.
    /// Thread-safety: pure.
    /// </summary>
    public static object[,] BoxDoubleMatrix(double[,] block)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var result = new object[rows, cols];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                // Boxing here is intentional: Excel's COM interop only accepts object[,].
                result[r, c] = block[r, c];
            }
        }
        return result;
    }

    /// <summary>
    /// Boxes a <c>string[,]</c> back into an <c>object[,]</c>. Null entries become the
    /// empty string so Excel writes a blank cell rather than <see cref="ExcelError.ExcelErrorNull"/>.
    /// </summary>
    public static object[,] BoxStringMatrix(string?[,] block)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var result = new object[rows, cols];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                result[r, c] = block[r, c] ?? string.Empty;
            }
        }
        return result;
    }

    /// <summary>
    /// Returns a flat row-major <see cref="Span{T}"/> over a one-dimensional column of
    /// values pulled from a single column of a two-dimensional block. Avoids per-cell
    /// boxing when the caller only needs one column.
    /// Marshaling cost: 0 boundary crossings.
    /// Thread-safety: pure.
    /// </summary>
    public static double[] ExtractColumnAsDoubles(object[,] block, int columnIndex, double defaultValue = 0d)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        if ((uint)columnIndex >= (uint)cols)
        {
            throw new ArgumentOutOfRangeException(nameof(columnIndex));
        }
        var result = new double[rows];
        for (var r = 0; r < rows; r++)
        {
            result[r] = TryToDouble(block[r, columnIndex], out var d) ? d : defaultValue;
        }
        return result;
    }

    /// <summary>
    /// Returns a one-cell <c>object[,]</c> that wraps a single error sentinel. Useful at
    /// boundary entry points that must return an array to satisfy the registration but
    /// also want to surface a validation failure.
    /// Marshaling cost: 0 boundary crossings.
    /// Thread-safety: pure.
    /// </summary>
    public static object[,] ErrorBlock(ExcelError err)
    {
        var result = new object[1, 1];
        result[0, 0] = err;
        return result;
    }

    /// <summary>
    /// Allocates a destination block of the same shape as <paramref name="source"/>.
    /// Marshaling cost: 0 boundary crossings.
    /// </summary>
    public static object[,] AllocateLike(object[,] source)
    {
        ArgumentNullException.ThrowIfNull(source);
        return new object[source.GetLength(0), source.GetLength(1)];
    }

    /// <summary>
    /// Equality comparer for object cell values that treats Excel sentinels consistently.
    /// </summary>
    public static IEqualityComparer<object?> CellEquality { get; } = new CellEqualityComparer();

    private sealed class CellEqualityComparer : IEqualityComparer<object?>
    {
        public new bool Equals(object? x, object? y)
        {
            if (IsBlankOrError(x) && IsBlankOrError(y))
            {
                if (x is ExcelError ex && y is ExcelError ey)
                {
                    return ex == ey;
                }
                return x?.GetType() == y?.GetType();
            }
            if (TryToDouble(x, out var dx) && TryToDouble(y, out var dy))
            {
                return dx.Equals(dy);
            }
            return string.Equals(ToStringSafe(x), ToStringSafe(y), StringComparison.Ordinal);
        }

        public int GetHashCode(object? obj)
        {
            if (IsBlankOrError(obj))
            {
                // Include the specific ExcelError enum value so that NA / DIV/0 / REF
                // don't all collide into a single bucket - keeps hashed-collection probes
                // O(1) when error-heavy data flows through CellEquality.
                if (obj is ExcelError err)
                {
                    return HashCode.Combine(typeof(ExcelError), (int)err);
                }
                return obj?.GetType().GetHashCode() ?? 0;
            }
            if (TryToDouble(obj, out var d))
            {
                return d.GetHashCode();
            }
            return ToStringSafe(obj).GetHashCode(StringComparison.Ordinal);
        }
    }
}
