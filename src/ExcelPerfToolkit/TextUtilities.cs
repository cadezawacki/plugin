using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using ExcelDna.Integration;

namespace ExcelPerfToolkit;

/// <summary>
/// Round 1 string micro-utilities: case conversion, padding, repetition, reversal, and
/// templated fill. Every function honors the one-crossing rule (one bulk <c>object[,]</c>
/// in, one bulk <c>object[,]</c> out) and is pure CPU over its arguments, so all are
/// registered <c>IsThreadSafe = true</c> and eligible for Excel's multithreaded recalc.
///
/// <para>Bottlenecks addressed: #1 (cell-by-cell read/write) and #2 (per-cell COM
/// marshaling) - whole blocks are transformed in managed memory without re-entering COM.
/// Validation failures throw and are surfaced as <c>#VALUE!</c> by the add-in's registered
/// unhandled-exception handler; per-cell limits embed an <see cref="ExcelError"/> in the
/// offending cell.</para>
/// </summary>
public static class TextUtilities
{
    /// <summary>Excel's hard limit on the number of characters a single cell can hold.</summary>
    private const int MaxCellChars = 32_767;

    // ---------- Case conversion ----------

    /// <summary>
    /// Capitalizes the first letter of each word and lower-cases the rest, like Excel's
    /// native <c>PROPER</c> but over a whole block in one pass. Word boundaries are any
    /// run of non-letters. Non-string cells pass through unchanged.
    /// Marshaling cost: O(1) - one bulk in, one bulk out. Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.PROPER", Description = "Proper-case every string cell in a block (first letter of each word upper, rest lower).", Category = "EPT.Text", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] Proper(object[,] block)
        => MapCells(block, c => c is string s ? ToProper(s) : (c ?? ExcelEmpty.Value));

    /// <summary>
    /// Title-cases every string cell using the invariant culture's title-casing, which - unlike
    /// <see cref="Proper"/> - leaves all-upper acronyms intact. Non-string cells pass through.
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.TITLECASE", Description = "Title-case every string cell (preserves acronyms).", Category = "EPT.Text", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] TitleCase(object[,] block)
        => MapCells(block, c => c is string s ? CultureInfo.InvariantCulture.TextInfo.ToTitleCase(s) : (c ?? ExcelEmpty.Value));

    /// <summary>
    /// Converts every string cell to <c>camelCase</c>: words (runs of letters/digits) are
    /// Pascal-cased and concatenated, then the very first character is lower-cased. Separators
    /// are dropped. Non-string cells pass through unchanged.
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.CAMELCASE", Description = "Convert every string cell to camelCase (drops separators).", Category = "EPT.Text", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] CamelCase(object[,] block)
        => MapCells(block, c => c is string s ? ToCamel(s) : (c ?? ExcelEmpty.Value));

    // ---------- Padding ----------

    /// <summary>
    /// Pads the text form of every non-blank cell on the left to <paramref name="totalWidth"/>
    /// using <paramref name="padChar"/> (default space). Cells already at or above the width
    /// are returned unchanged. Blank and error cells pass through.
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.PADLEFT", Description = "Left-pad each cell's text to a fixed width.", Category = "EPT.Text", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] PadLeft(
        object[,] block,
        [ExcelArgument(Name = "total_width")] double totalWidth,
        [ExcelArgument(Name = "pad_char", Description = "Optional pad character; defaults to a space.")] object padChar)
    {
        var width = WidthOf(totalWidth);
        var pad = PadCharOf(padChar, ' ');
        return MapCells(block, c => IsBlankCell(c) ? (c ?? ExcelEmpty.Value) : Marshaling.ToStringSafe(c).PadLeft(width, pad));
    }

    /// <summary>
    /// Pads the text form of every non-blank cell on the right to <paramref name="totalWidth"/>
    /// using <paramref name="padChar"/> (default space). Blank and error cells pass through.
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.PADRIGHT", Description = "Right-pad each cell's text to a fixed width.", Category = "EPT.Text", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] PadRight(
        object[,] block,
        [ExcelArgument(Name = "total_width")] double totalWidth,
        [ExcelArgument(Name = "pad_char", Description = "Optional pad character; defaults to a space.")] object padChar)
    {
        var width = WidthOf(totalWidth);
        var pad = PadCharOf(padChar, ' ');
        return MapCells(block, c => IsBlankCell(c) ? (c ?? ExcelEmpty.Value) : Marshaling.ToStringSafe(c).PadRight(width, pad));
    }

    /// <summary>
    /// Left-pads the text form of every non-blank cell with <c>'0'</c> to
    /// <paramref name="totalWidth"/>. Convenient for zero-padding IDs and codes. The pad is
    /// purely textual: a value like <c>-5</c> becomes <c>"00-5"</c>, so apply it to codes,
    /// not to signed numbers you intend to keep numeric. Blank and error cells pass through.
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.ZEROPAD", Description = "Left-pad each cell's text with zeros to a fixed width.", Category = "EPT.Text", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] ZeroPad(object[,] block, [ExcelArgument(Name = "total_width")] double totalWidth)
    {
        var width = WidthOf(totalWidth);
        return MapCells(block, c => IsBlankCell(c) ? (c ?? ExcelEmpty.Value) : Marshaling.ToStringSafe(c).PadLeft(width, '0'));
    }

    // ---------- Repeat / reverse ----------

    /// <summary>
    /// Repeats the text form of every non-blank cell <paramref name="count"/> times. A result
    /// that would exceed the 32,767-character cell limit throws (surfaced as <c>#VALUE!</c>).
    /// Blank and error cells pass through.
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.REPEAT", Description = "Repeat each cell's text N times.", Category = "EPT.Text", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] Repeat(object[,] block, [ExcelArgument(Name = "count")] double count)
    {
        if (double.IsNaN(count) || count < 0d || count != Math.Truncate(count))
        {
            throw new ArgumentException("count must be a non-negative integer.", nameof(count));
        }
        var n = (int)count;
        return MapCells(block, c =>
        {
            if (IsBlankCell(c))
            {
                return c ?? ExcelEmpty.Value;
            }
            var s = Marshaling.ToStringSafe(c);
            if ((long)n * s.Length > MaxCellChars)
            {
                throw new ArgumentException($"Repeated string exceeds the {MaxCellChars}-character cell limit.");
            }
            var sb = new StringBuilder(n * s.Length);
            for (var i = 0; i < n; i++)
            {
                sb.Append(s);
            }
            return sb.ToString();
        });
    }

    /// <summary>
    /// Reverses the characters of every non-blank cell's text. Reversal is by UTF-16 code
    /// unit; surrogate-pair characters (some emoji) are not preserved. Blank and error cells
    /// pass through.
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.REVERSE", Description = "Reverse the characters of each cell's text.", Category = "EPT.Text", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] Reverse(object[,] block)
        => MapCells(block, c =>
        {
            if (IsBlankCell(c))
            {
                return c ?? ExcelEmpty.Value;
            }
            var chars = Marshaling.ToStringSafe(c).ToCharArray();
            Array.Reverse(chars);
            return new string(chars);
        });

    // ---------- Template fill ----------

    /// <summary>
    /// Renders <paramref name="template"/> once per data row, substituting <c>{placeholder}</c>
    /// tokens with cell values - a mail-merge in a single formula. Placeholders may be a
    /// 0-based column index (<c>{0}</c>) or, when <paramref name="hasHeaderRow"/> is true
    /// (the default), a header name from row 0 (<c>{Region}</c>, case-insensitive). Unknown
    /// placeholders are left literal; use <c>{{</c> and <c>}}</c> for literal braces. Returns
    /// a single column of rendered strings, one per data row.
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.TEMPLATEFILL", Description = "Render a {placeholder} template once per data row (mail-merge).", Category = "EPT.Text", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] TemplateFill(
        [ExcelArgument(Name = "template", Description = "Template text with {name} or {index} placeholders.")] string template,
        [ExcelArgument(Name = "data", Description = "Data block; row 0 holds header names when has_header_row is TRUE.")] object[,] data,
        [ExcelArgument(Name = "has_header_row", Description = "TRUE (default) treats row 0 as header names.")] bool hasHeaderRow = true)
    {
        ArgumentNullException.ThrowIfNull(data);
        template ??= string.Empty;
        var rows = data.GetLength(0);
        var cols = data.GetLength(1);

        Dictionary<string, int>? names = null;
        int dataStart, outRows;
        if (hasHeaderRow)
        {
            names = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (var c = 0; c < cols; c++)
            {
                var header = Marshaling.ToStringSafe(data[0, c]);
                if (header.Length > 0 && !names.ContainsKey(header))
                {
                    names[header] = c;
                }
            }
            dataStart = 1;
            outRows = rows - 1;
        }
        else
        {
            dataStart = 0;
            outRows = rows;
        }

        if (outRows <= 0)
        {
            return new object[1, 1] { { ExcelEmpty.Value } };
        }
        var result = new object[outRows, 1];
        for (var r = 0; r < outRows; r++)
        {
            result[r, 0] = RenderTemplate(template, data, dataStart + r, names, cols);
        }
        return result;
    }

    // ---------- Helpers ----------

    private static object[,] MapCells(object[,] block, Func<object?, object> cellFn)
    {
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var result = new object[rows, cols];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                result[r, c] = cellFn(block[r, c]);
            }
        }
        return result;
    }

    private static bool IsBlankCell(object? v)
        => Marshaling.IsBlankOrError(v) || (v is string s && s.Length == 0);

    private static int WidthOf(double totalWidth)
    {
        if (double.IsNaN(totalWidth) || totalWidth < 0d || totalWidth > MaxCellChars)
        {
            throw new ArgumentException($"total_width must be between 0 and {MaxCellChars}.", nameof(totalWidth));
        }
        return (int)totalWidth;
    }

    private static char PadCharOf(object padChar, char fallback)
    {
        if (Marshaling.IsBlankOrError(padChar))
        {
            return fallback;
        }
        var s = Marshaling.ToStringSafe(padChar);
        return s.Length > 0 ? s[0] : fallback;
    }

    private static string ToProper(string s)
    {
        var sb = new StringBuilder(s.Length);
        var newWord = true;
        foreach (var ch in s)
        {
            if (char.IsLetter(ch))
            {
                sb.Append(newWord ? char.ToUpperInvariant(ch) : char.ToLowerInvariant(ch));
                newWord = false;
            }
            else
            {
                sb.Append(ch);
                newWord = true;
            }
        }
        return sb.ToString();
    }

    private static string ToCamel(string s)
    {
        var sb = new StringBuilder(s.Length);
        var i = 0;
        while (i < s.Length)
        {
            // Skip separators between words.
            while (i < s.Length && !char.IsLetterOrDigit(s[i]))
            {
                i++;
            }
            if (i >= s.Length)
            {
                break;
            }
            // Consume one word and Pascal-case it.
            var first = true;
            while (i < s.Length && char.IsLetterOrDigit(s[i]))
            {
                var ch = s[i];
                sb.Append(first ? char.ToUpperInvariant(ch) : char.ToLowerInvariant(ch));
                first = false;
                i++;
            }
        }
        if (sb.Length > 0)
        {
            sb[0] = char.ToLowerInvariant(sb[0]);
        }
        return sb.ToString();
    }

    private static string RenderTemplate(string template, object[,] data, int row, Dictionary<string, int>? names, int cols)
    {
        var sb = new StringBuilder(template.Length);
        var i = 0;
        while (i < template.Length)
        {
            var ch = template[i];
            if (ch == '{')
            {
                if (i + 1 < template.Length && template[i + 1] == '{')
                {
                    sb.Append('{');
                    i += 2;
                    continue;
                }
                var end = template.IndexOf('}', i + 1);
                if (end < 0)
                {
                    sb.Append(template, i, template.Length - i);
                    break;
                }
                var token = template.Substring(i + 1, end - i - 1);
                if (TryResolveToken(token, names, cols, out var col))
                {
                    sb.Append(Marshaling.ToStringSafe(data[row, col]));
                }
                else
                {
                    sb.Append('{').Append(token).Append('}');
                }
                i = end + 1;
            }
            else if (ch == '}' && i + 1 < template.Length && template[i + 1] == '}')
            {
                sb.Append('}');
                i += 2;
            }
            else
            {
                sb.Append(ch);
                i++;
            }
        }
        return sb.ToString();
    }

    private static bool TryResolveToken(string token, Dictionary<string, int>? names, int cols, out int column)
    {
        if (int.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out var idx) && idx >= 0 && idx < cols)
        {
            column = idx;
            return true;
        }
        if (names is not null && names.TryGetValue(token, out column) && column < cols)
        {
            return true;
        }
        column = -1;
        return false;
    }
}
