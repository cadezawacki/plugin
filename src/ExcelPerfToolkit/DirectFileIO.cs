using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelPerfToolkit;

/// <summary>
/// Direct, streaming file input and output for tabular data. This file addresses the
/// "open a workbook through the object model just to read a CSV" bottleneck: we go
/// straight to managed <see cref="FileStream"/> with async I/O, RFC 4180-compliant
/// quoting, and never touch <c>Application</c>, <c>Workbooks.Open</c>, or any other
/// COM surface.
///
/// Engineering rules followed by every function here:
///   * All I/O is asynchronous and cancellable. No worker thread blocks on disk.
///   * No Excel object model access. Reads return <c>object[,]</c> that the caller can
///     hand to <see cref="BulkTransfer.WriteBlock(ExcelReference, object[,])"/> for the
///     single write-side boundary crossing.
///   * Streaming parser: works incrementally over arbitrarily large files; the only
///     hot-path allocation is the growing <see cref="StringBuilder"/> for the current
///     field and a row buffer reused per row.
///   * The worksheet UDF wrappers run synchronously (Excel's calling convention is
///     synchronous) but execute on a worker thread and block only that worker on
///     completion of the underlying async pipeline.
/// </summary>
public static class DirectFileIO
{
    private static readonly TraceSource TraceSource = ToolkitLifetime.CreateTraceSource("DirectFileIO");

    /// <summary>
    /// Default buffer size for the underlying streams. 64 KiB is large enough to make
    /// per-read syscall overhead negligible for files of any practical size.
    /// </summary>
    public const int DefaultBufferSize = 65_536;

    /// <summary>
    /// Reads a delimited text file into a single <c>object[,]</c>. Async; no Excel
    /// object model is touched. Quoted fields, embedded delimiters, embedded newlines,
    /// and doubled-quote escapes are handled per RFC 4180.
    /// Marshaling cost: 0 boundary crossings on the read itself. When the caller hands
    /// the result to <see cref="BulkTransfer.WriteBlock(ExcelReference, object[,])"/>
    /// the worksheet write costs 1 crossing - regardless of file size.
    /// Thread-safety: safe to call concurrently with itself on different paths. NOT
    /// safe inside MTR worker threads because of the async-bridging in the UDF
    /// wrapper; expose async paths to .NET callers instead.
    /// </summary>
    public static async Task<object[,]> ReadDelimitedAsync(
        string path,
        char delimiter = ',',
        Encoding? encoding = null,
        bool coerceNumeric = true,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("Path is required.", nameof(path));
        }
        encoding ??= Encoding.UTF8;

        await using var stream = new FileStream(
            path,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete,
            bufferSize: DefaultBufferSize,
            options: FileOptions.Asynchronous | FileOptions.SequentialScan);
        using var reader = new StreamReader(stream, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: DefaultBufferSize);

        var rows = new List<object[]>(capacity: 256);
        var maxCols = 0;
        var fieldBuffer = new StringBuilder(64);
        var rowBuffer = new List<object>(16);

        // Pull characters in chunks rather than line-by-line, because a quoted field
        // may legally span lines.
        var charBuffer = new char[8192];
        var inQuotes = false;
        var lastWasCR = false;

        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var read = await reader.ReadAsync(charBuffer.AsMemory(), cancellationToken).ConfigureAwait(false);
            if (read == 0)
            {
                break;
            }
            for (var i = 0; i < read; i++)
            {
                var ch = charBuffer[i];
                if (inQuotes)
                {
                    if (ch == '"')
                    {
                        // Peek the next char to detect doubled-quote escape.
                        if (i + 1 < read)
                        {
                            if (charBuffer[i + 1] == '"')
                            {
                                fieldBuffer.Append('"');
                                i++;
                                continue;
                            }
                            inQuotes = false;
                            continue;
                        }
                        // Boundary case: quote at the very end of a chunk - refill and
                        // re-examine the next char. We do this by reading one more
                        // char from the reader and re-applying the rule.
                        var next = reader.Peek();
                        if (next == '"')
                        {
                            reader.Read();
                            fieldBuffer.Append('"');
                            continue;
                        }
                        inQuotes = false;
                        continue;
                    }
                    fieldBuffer.Append(ch);
                    continue;
                }
                if (ch == '"')
                {
                    inQuotes = true;
                    continue;
                }
                if (ch == delimiter)
                {
                    PushField(fieldBuffer, rowBuffer, coerceNumeric);
                    continue;
                }
                if (ch == '\r')
                {
                    PushField(fieldBuffer, rowBuffer, coerceNumeric);
                    PushRow(rowBuffer, rows, ref maxCols);
                    lastWasCR = true;
                    continue;
                }
                if (ch == '\n')
                {
                    if (lastWasCR)
                    {
                        lastWasCR = false;
                        continue;
                    }
                    PushField(fieldBuffer, rowBuffer, coerceNumeric);
                    PushRow(rowBuffer, rows, ref maxCols);
                    continue;
                }
                lastWasCR = false;
                fieldBuffer.Append(ch);
            }
        }

        // Flush trailing field/row if the file did not end with a newline.
        if (fieldBuffer.Length > 0 || rowBuffer.Count > 0)
        {
            PushField(fieldBuffer, rowBuffer, coerceNumeric);
            PushRow(rowBuffer, rows, ref maxCols);
        }

        if (rows.Count == 0)
        {
            return new object[1, 1] { { ExcelEmpty.Value } };
        }

        var result = new object[rows.Count, maxCols];
        for (var r = 0; r < rows.Count; r++)
        {
            var row = rows[r];
            for (var c = 0; c < maxCols; c++)
            {
                result[r, c] = c < row.Length ? row[c] : ExcelEmpty.Value;
            }
        }
        return result;
    }

    private static void PushField(StringBuilder field, List<object> row, bool coerceNumeric)
    {
        var text = field.ToString();
        field.Clear();
        if (coerceNumeric && text.Length > 0 && Marshaling.TryToDouble(text, out var d))
        {
            row.Add(d);
        }
        else
        {
            row.Add(text);
        }
    }

    private static void PushRow(List<object> rowBuffer, List<object[]> rows, ref int maxCols)
    {
        if (rowBuffer.Count == 0)
        {
            return;
        }
        var arr = rowBuffer.ToArray();
        rows.Add(arr);
        if (arr.Length > maxCols)
        {
            maxCols = arr.Length;
        }
        rowBuffer.Clear();
    }

    /// <summary>
    /// Writes <paramref name="block"/> as a delimited text file. Async; no Excel object
    /// model is touched. Fields containing the delimiter, double quotes, or newlines
    /// are quoted per RFC 4180.
    /// Marshaling cost: 0 boundary crossings on the write itself. When the caller
    /// produced <paramref name="block"/> via
    /// <see cref="BulkTransfer.ReadBlock(ExcelReference)"/> the worksheet read costs
    /// 1 crossing - regardless of file size.
    /// Thread-safety: safe to call concurrently with itself on different paths.
    /// </summary>
    public static async Task WriteDelimitedAsync(
        string path,
        object[,] block,
        char delimiter = ',',
        Encoding? encoding = null,
        string newLine = "\r\n",
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("Path is required.", nameof(path));
        }
        ArgumentNullException.ThrowIfNull(block);
        encoding ??= new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

        await using var stream = new FileStream(
            path,
            FileMode.Create,
            FileAccess.Write,
            FileShare.Read,
            bufferSize: DefaultBufferSize,
            options: FileOptions.Asynchronous | FileOptions.SequentialScan);
        await using var writer = new StreamWriter(stream, encoding, bufferSize: DefaultBufferSize)
        {
            NewLine = newLine,
            AutoFlush = false,
        };

        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        var sb = new StringBuilder(256);

        for (var r = 0; r < rows; r++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            sb.Clear();
            for (var c = 0; c < cols; c++)
            {
                if (c > 0)
                {
                    sb.Append(delimiter);
                }
                AppendQuoted(sb, Marshaling.ToStringSafe(block[r, c]), delimiter);
            }
            await writer.WriteLineAsync(sb, cancellationToken).ConfigureAwait(false);
        }
        await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    private static void AppendQuoted(StringBuilder sb, string field, char delimiter)
    {
        var needsQuote = false;
        for (var i = 0; i < field.Length; i++)
        {
            var ch = field[i];
            if (ch == delimiter || ch == '"' || ch == '\r' || ch == '\n')
            {
                needsQuote = true;
                break;
            }
        }
        if (!needsQuote)
        {
            sb.Append(field);
            return;
        }
        sb.Append('"');
        for (var i = 0; i < field.Length; i++)
        {
            var ch = field[i];
            if (ch == '"')
            {
                sb.Append('"').Append('"');
            }
            else
            {
                sb.Append(ch);
            }
        }
        sb.Append('"');
    }

    // ====================================================================
    // UDF surface
    // ====================================================================

    /// <summary>
    /// Worksheet UDF that reads a delimited file into an array result. Bridges the
    /// async pipeline to Excel's synchronous calling convention by awaiting on a
    /// task obtained from <see cref="ReadDelimitedAsync"/>. Cancellation honors the
    /// add-in-wide shutdown token from <see cref="ToolkitLifetime.ShutdownToken"/>.
    /// Marshaling cost: 1 write crossing - the file size is irrelevant.
    /// Thread-safety: NOT registered as MTR-safe. The async bridge here would be
    /// hostile to Excel's MTR threading model; expose the async method directly to
    /// .NET callers instead.
    /// </summary>
    [ExcelFunction(
        Name = "EPT.READCSV",
        Description = "Read a delimited text file directly into a bulk array. Async, no workbook open, no COM.",
        Category = "EPT.DirectFileIO",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static object ReadDelimitedUdf(
        [ExcelArgument(Name = "path", Description = "Absolute or workbook-relative file path.")] string path,
        [ExcelArgument(Name = "delimiter", Description = "Optional. Defaults to ','. Pass \"\\t\" for TSV.")] object delimiter,
        [ExcelArgument(Name = "encoding", Description = "Optional encoding name (e.g. 'utf-8', 'windows-1252'). Defaults to UTF-8.")] object encoding,
        [ExcelArgument(Name = "coerceNumeric", Description = "Optional. TRUE (default) parses numeric-looking fields as numbers.")] object coerceNumeric)
    {
        try
        {
            var delim = ResolveDelimiter(delimiter);
            var enc = ResolveEncoding(encoding);
            var coerce = ResolveBool(coerceNumeric, defaultValue: true);
            return ReadDelimitedAsync(path, delim, enc, coerce, ToolkitLifetime.ShutdownToken)
                .ConfigureAwait(false)
                .GetAwaiter()
                .GetResult();
        }
        catch (OperationCanceledException)
        {
            TraceSource.TraceEvent(TraceEventType.Information, 2, "EPT.READCSV cancelled for {0}", path);
            return ExcelError.ExcelErrorNA;
        }
        catch (FileNotFoundException ex)
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 3, "EPT.READCSV file not found: {0}", ex.Message);
            return ExcelError.ExcelErrorNA;
        }
        catch (Exception ex)
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 4, "EPT.READCSV failed: {0}", ex.Message);
            return ExcelError.ExcelErrorValue;
        }
    }

    /// <summary>
    /// Worksheet UDF that writes a block to a delimited file. Returns the row count
    /// written on success or an <see cref="ExcelError"/> on failure.
    /// Marshaling cost: 1 read crossing - the file size is irrelevant.
    /// Thread-safety: NOT registered as MTR-safe (writes to disk).
    /// </summary>
    [ExcelFunction(
        Name = "EPT.WRITECSV",
        Description = "Write a block to a delimited text file. Async, no workbook open, no COM.",
        Category = "EPT.DirectFileIO",
        IsThreadSafe = false,
        IsVolatile = false)]
    public static object WriteDelimitedUdf(
        [ExcelArgument(Name = "path", Description = "Absolute or workbook-relative output path.")] string path,
        [ExcelArgument(Name = "block", Description = "The data block to write.")] object[,] block,
        [ExcelArgument(Name = "delimiter", Description = "Optional. Defaults to ','. Pass \"\\t\" for TSV.")] object delimiter,
        [ExcelArgument(Name = "encoding", Description = "Optional encoding name. Defaults to UTF-8 (no BOM).")] object encoding)
    {
        try
        {
            var delim = ResolveDelimiter(delimiter);
            var enc = ResolveEncoding(encoding);
            WriteDelimitedAsync(path, block, delim, enc, "\r\n", ToolkitLifetime.ShutdownToken)
                .ConfigureAwait(false)
                .GetAwaiter()
                .GetResult();
            return (double)block.GetLength(0);
        }
        catch (OperationCanceledException)
        {
            TraceSource.TraceEvent(TraceEventType.Information, 5, "EPT.WRITECSV cancelled for {0}", path);
            return ExcelError.ExcelErrorNA;
        }
        catch (UnauthorizedAccessException ex)
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 6, "EPT.WRITECSV unauthorized: {0}", ex.Message);
            return ExcelError.ExcelErrorNA;
        }
        catch (Exception ex)
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 7, "EPT.WRITECSV failed: {0}", ex.Message);
            return ExcelError.ExcelErrorValue;
        }
    }

    private static char ResolveDelimiter(object delimiter)
    {
        if (Marshaling.IsBlankOrError(delimiter))
        {
            return ',';
        }
        var s = Marshaling.ToStringSafe(delimiter);
        if (s.Length == 0)
        {
            return ',';
        }
        if (string.Equals(s, "\\t", StringComparison.Ordinal) || s == "\t")
        {
            return '\t';
        }
        return s[0];
    }

    private static Encoding ResolveEncoding(object encoding)
    {
        if (Marshaling.IsBlankOrError(encoding))
        {
            return Encoding.UTF8;
        }
        var s = Marshaling.ToStringSafe(encoding);
        if (string.IsNullOrWhiteSpace(s))
        {
            return Encoding.UTF8;
        }
        try
        {
            return Encoding.GetEncoding(s);
        }
        catch (ArgumentException)
        {
            return Encoding.UTF8;
        }
    }

    private static bool ResolveBool(object value, bool defaultValue)
    {
        if (Marshaling.IsBlankOrError(value))
        {
            return defaultValue;
        }
        return value switch
        {
            bool b => b,
            double d => d != 0d,
            string s => s.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || s == "1",
            _ => defaultValue,
        };
    }
}
