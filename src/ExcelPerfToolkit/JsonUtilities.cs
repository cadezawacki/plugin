using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelPerfToolkit;

/// <summary>
/// Round 7 JSON utilities, built entirely on the in-box <see cref="System.Text.Json"/> -
/// no third-party dependency. Two pure-CPU, MTR-safe cell functions extract from JSON text
/// already in the grid (<c>EPT.JSONPATH</c>, <c>EPT.PARSEJSON</c>), and three async,
/// object-model-free file functions read and write JSON on disk (<c>EPT.READJSON</c>,
/// <c>EPT.READNDJSON</c>, <c>EPT.WRITEJSON</c>) following the streaming bridge established
/// by <see cref="DirectFileIO"/>.
///
/// <para><b>Path syntax.</b> The <c>path</c>/<c>pointer</c> arguments accept a documented
/// subset: dotted keys and <c>[index]</c> array steps, with an optional leading <c>$</c> or
/// <c>$.</c> (e.g. <c>data.items[0].name</c>). Wildcards and filter expressions are not
/// supported.</para>
///
/// <para><b>Value mapping.</b> JSON numbers become <see cref="double"/>, strings stay
/// strings, booleans stay booleans, null becomes a blank cell, and nested objects/arrays are
/// returned as their compact JSON text so they still fit a single cell.</para>
/// </summary>
public static class JsonUtilities
{
    private static readonly TraceSource TraceSource = ToolkitLifetime.CreateTraceSource("JsonUtilities");

    // ====================================================================
    // Pure-CPU cell functions (MTR-safe)
    // ====================================================================

    /// <summary>
    /// Extracts the value at <paramref name="path"/> from the JSON text in each cell,
    /// returning a block of the same shape. Scalars come back typed; nested objects/arrays
    /// come back as compact JSON text. A cell whose JSON fails to parse yields <c>#VALUE!</c>;
    /// a path that does not resolve yields <c>#N/A</c>; a blank cell stays blank.
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.JSONPATH", Description = "Extract a value at a dotted/indexed path from JSON text in each cell.", Category = "EPT.Json", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] JsonPath(
        [ExcelArgument(Name = "json", Description = "Cell(s) containing JSON text.")] object[,] json,
        [ExcelArgument(Name = "path", Description = "Dotted/indexed path, e.g. data.items[0].name.")] string path)
    {
        ArgumentNullException.ThrowIfNull(json);
        var rows = json.GetLength(0);
        var cols = json.GetLength(1);
        var result = new object[rows, cols];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                var cell = json[r, c];
                if (Marshaling.IsBlankOrError(cell))
                {
                    result[r, c] = ExcelEmpty.Value;
                    continue;
                }
                try
                {
                    using var doc = JsonDocument.Parse(Marshaling.ToStringSafe(cell));
                    result[r, c] = TryNavigate(doc.RootElement, path, out var node)
                        ? JsonToCell(node)
                        : ExcelError.ExcelErrorNA;
                }
                catch (JsonException)
                {
                    result[r, c] = ExcelError.ExcelErrorValue;
                }
            }
        }
        return result;
    }

    /// <summary>
    /// Parses a single JSON document and spills it into a table. An array of objects becomes a
    /// table with a header row (the union of keys, in first-seen order) when
    /// <paramref name="hasHeaderRow"/> is true (the default); an array of scalars becomes one
    /// column; an object becomes a two-column key/value table; a scalar becomes a single cell.
    /// An optional <paramref name="path"/> navigates to a sub-node first. Invalid JSON returns
    /// <c>#VALUE!</c>; an unresolved path returns <c>#N/A</c>.
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.PARSEJSON", Description = "Parse a JSON document into a spilled table.", Category = "EPT.Json", IsThreadSafe = true, IsVolatile = false)]
    public static object[,] ParseJson(
        [ExcelArgument(Name = "json", Description = "A single JSON document.")] string json,
        [ExcelArgument(Name = "path", Description = "Optional sub-node path to expand.")] string? path = null,
        [ExcelArgument(Name = "has_header_row", Description = "TRUE (default) emits a header row for arrays of objects.")] bool hasHeaderRow = true)
    {
        if (string.IsNullOrEmpty(json))
        {
            return new object[1, 1] { { ExcelEmpty.Value } };
        }
        try
        {
            using var doc = JsonDocument.Parse(json);
            if (!TryNavigate(doc.RootElement, path, out var node))
            {
                return new object[1, 1] { { ExcelError.ExcelErrorNA } };
            }
            return ExpandToTable(node, hasHeaderRow);
        }
        catch (JsonException)
        {
            return new object[1, 1] { { ExcelError.ExcelErrorValue } };
        }
    }

    // ====================================================================
    // Async file cores (no Excel object model)
    // ====================================================================

    /// <summary>
    /// Reads a JSON file and expands it into a table, optionally navigating to
    /// <paramref name="pointer"/> first. Async; no object model touched.
    /// </summary>
    public static async Task<object[,]> ReadJsonAsync(string path, string? pointer, bool hasHeaderRow, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("Path is required.", nameof(path));
        }
        var bytes = await File.ReadAllBytesAsync(path, cancellationToken).ConfigureAwait(false);
        using var doc = JsonDocument.Parse(bytes);
        if (!TryNavigate(doc.RootElement, pointer, out var node))
        {
            return new object[1, 1] { { ExcelError.ExcelErrorNA } };
        }
        return ExpandToTable(node, hasHeaderRow);
    }

    /// <summary>
    /// Reads newline-delimited JSON (one JSON value per line), streaming line by line, and
    /// builds a table from the union of object keys (in first-seen order). Non-object lines
    /// are placed under a single <c>value</c> column. Async; no object model touched.
    /// </summary>
    public static async Task<object[,]> ReadNdjsonAsync(string path, bool hasHeaderRow, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("Path is required.", nameof(path));
        }

        await using var stream = new FileStream(
            path,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            bufferSize: DirectFileIO.DefaultBufferSize,
            options: FileOptions.Asynchronous | FileOptions.SequentialScan);
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: DirectFileIO.DefaultBufferSize);

        var order = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var rows = new List<Dictionary<string, object>>(256);

        while (await reader.ReadLineAsync(cancellationToken).ConfigureAwait(false) is { } line)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }
            using var doc = JsonDocument.Parse(line);
            var element = doc.RootElement;
            var dict = new Dictionary<string, object>(StringComparer.Ordinal);
            if (element.ValueKind == JsonValueKind.Object)
            {
                foreach (var prop in element.EnumerateObject())
                {
                    if (seen.Add(prop.Name))
                    {
                        order.Add(prop.Name);
                    }
                    dict[prop.Name] = JsonToCell(prop.Value);
                }
            }
            else
            {
                if (seen.Add("value"))
                {
                    order.Add("value");
                }
                dict["value"] = JsonToCell(element);
            }
            rows.Add(dict);
        }

        if (rows.Count == 0)
        {
            return new object[1, 1] { { ExcelEmpty.Value } };
        }
        return BuildObjectTable(order, rows, hasHeaderRow);
    }

    /// <summary>
    /// Writes <paramref name="block"/> as a JSON array. With <paramref name="hasHeaderRow"/>
    /// true (the default), row 0 supplies object keys and each later row becomes an object;
    /// otherwise each row becomes an array. Returns the number of records written. Async; no
    /// object model touched.
    /// </summary>
    public static async Task<int> WriteJsonAsync(string path, object[,] block, bool hasHeaderRow, bool indent, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("Path is required.", nameof(path));
        }
        ArgumentNullException.ThrowIfNull(block);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);

        await using var stream = new FileStream(
            path,
            FileMode.Create,
            FileAccess.Write,
            FileShare.Read,
            bufferSize: DirectFileIO.DefaultBufferSize,
            options: FileOptions.Asynchronous | FileOptions.SequentialScan);
        await using var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = indent });

        string[]? keys = null;
        var dataStart = 0;
        if (hasHeaderRow && rows > 0)
        {
            keys = new string[cols];
            for (var c = 0; c < cols; c++)
            {
                var name = Marshaling.ToStringSafe(block[0, c]);
                keys[c] = name.Length > 0 ? name : "col" + c.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            dataStart = 1;
        }

        writer.WriteStartArray();
        var written = 0;
        for (var r = dataStart; r < rows; r++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (keys is not null)
            {
                writer.WriteStartObject();
                for (var c = 0; c < cols; c++)
                {
                    writer.WritePropertyName(keys[c]);
                    WriteCellJson(writer, block[r, c]);
                }
                writer.WriteEndObject();
            }
            else
            {
                writer.WriteStartArray();
                for (var c = 0; c < cols; c++)
                {
                    WriteCellJson(writer, block[r, c]);
                }
                writer.WriteEndArray();
            }
            written++;
            if (writer.BytesPending > DirectFileIO.DefaultBufferSize)
            {
                await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
            }
        }
        writer.WriteEndArray();
        await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
        return written;
    }

    // ====================================================================
    // UDF surface for the file functions (async bridge, not MTR-safe)
    // ====================================================================

    /// <summary>
    /// Worksheet UDF: read a JSON file into a spilled table. Marshaling cost: 1 write
    /// crossing. Thread-safety: NOT MTR-safe (async bridge + disk).
    /// </summary>
    [ExcelFunction(Name = "EPT.READJSON", Description = "Read a JSON file into a bulk array. Async, no workbook open, no COM.", Category = "EPT.Json", IsThreadSafe = false, IsVolatile = false)]
    public static object ReadJsonUdf(
        [ExcelArgument(Name = "path")] string path,
        [ExcelArgument(Name = "pointer", Description = "Optional sub-node path to expand.")] object pointer,
        [ExcelArgument(Name = "has_header_row", Description = "TRUE (default) emits a header row for arrays of objects.")] object hasHeaderRow)
        => RunFileRead("EPT.READJSON", path, () => ReadJsonAsync(path, PathOf(pointer), ResolveBool(hasHeaderRow, true), ToolkitLifetime.ShutdownToken));

    /// <summary>
    /// Worksheet UDF: read a newline-delimited JSON file into a spilled table. Marshaling
    /// cost: 1 write crossing. Thread-safety: NOT MTR-safe (async bridge + disk).
    /// </summary>
    [ExcelFunction(Name = "EPT.READNDJSON", Description = "Read newline-delimited JSON into a bulk array. Async, no workbook open, no COM.", Category = "EPT.Json", IsThreadSafe = false, IsVolatile = false)]
    public static object ReadNdjsonUdf(
        [ExcelArgument(Name = "path")] string path,
        [ExcelArgument(Name = "has_header_row", Description = "TRUE (default) emits a header row.")] object hasHeaderRow)
        => RunFileRead("EPT.READNDJSON", path, () => ReadNdjsonAsync(path, ResolveBool(hasHeaderRow, true), ToolkitLifetime.ShutdownToken));

    /// <summary>
    /// Worksheet UDF: write a block to a JSON file. Returns the record count. Marshaling
    /// cost: 1 read crossing. Thread-safety: NOT MTR-safe (async bridge + disk).
    /// </summary>
    [ExcelFunction(Name = "EPT.WRITEJSON", Description = "Write a block to a JSON file. Async, no workbook open, no COM.", Category = "EPT.Json", IsThreadSafe = false, IsVolatile = false)]
    public static object WriteJsonUdf(
        [ExcelArgument(Name = "path")] string path,
        [ExcelArgument(Name = "block")] object[,] block,
        [ExcelArgument(Name = "has_header_row", Description = "TRUE (default) treats row 0 as object keys.")] object hasHeaderRow,
        [ExcelArgument(Name = "indent", Description = "TRUE pretty-prints the output. Default FALSE.")] object indent)
    {
        try
        {
            var written = WriteJsonAsync(path, block, ResolveBool(hasHeaderRow, true), ResolveBool(indent, false), ToolkitLifetime.ShutdownToken)
                .ConfigureAwait(false)
                .GetAwaiter()
                .GetResult();
            return (double)written;
        }
        catch (OperationCanceledException)
        {
            TraceSource.TraceEvent(TraceEventType.Information, 1, "EPT.WRITEJSON cancelled for {0}", path);
            return ExcelError.ExcelErrorNA;
        }
        catch (UnauthorizedAccessException ex)
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 2, "EPT.WRITEJSON unauthorized: {0}", ex.Message);
            return ExcelError.ExcelErrorNA;
        }
        catch (Exception ex)
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 3, "EPT.WRITEJSON failed: {0}", ex.Message);
            return ExcelError.ExcelErrorValue;
        }
    }

    private static object RunFileRead(string name, string path, Func<Task<object[,]>> read)
    {
        try
        {
            return read().ConfigureAwait(false).GetAwaiter().GetResult();
        }
        catch (OperationCanceledException)
        {
            TraceSource.TraceEvent(TraceEventType.Information, 4, "{0} cancelled for {1}", name, path);
            return ExcelError.ExcelErrorNA;
        }
        catch (FileNotFoundException ex)
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 5, "{0} file not found: {1}", name, ex.Message);
            return ExcelError.ExcelErrorNA;
        }
        catch (DirectoryNotFoundException ex)
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 6, "{0} directory not found: {1}", name, ex.Message);
            return ExcelError.ExcelErrorNA;
        }
        catch (JsonException ex)
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 7, "{0} invalid JSON: {1}", name, ex.Message);
            return ExcelError.ExcelErrorValue;
        }
        catch (Exception ex)
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 8, "{0} failed: {1}", name, ex.Message);
            return ExcelError.ExcelErrorValue;
        }
    }

    // ====================================================================
    // JSON <-> cell helpers
    // ====================================================================

    private static object JsonToCell(JsonElement element) => element.ValueKind switch
    {
        JsonValueKind.Number => element.TryGetDouble(out var d) ? d : (object)element.GetRawText(),
        JsonValueKind.String => element.GetString() ?? string.Empty,
        JsonValueKind.True => true,
        JsonValueKind.False => false,
        JsonValueKind.Null => ExcelEmpty.Value,
        JsonValueKind.Undefined => ExcelEmpty.Value,
        _ => element.GetRawText(),
    };

    private static void WriteCellJson(Utf8JsonWriter writer, object? cell)
    {
        if (Marshaling.IsBlankOrError(cell))
        {
            if (cell is ExcelError err)
            {
                writer.WriteStringValue(Marshaling.ErrorToText(err));
            }
            else
            {
                writer.WriteNullValue();
            }
            return;
        }
        switch (cell)
        {
            case bool b:
                writer.WriteBooleanValue(b);
                return;
            case string s:
                writer.WriteStringValue(s);
                return;
        }
        if (Marshaling.TryToDouble(cell, out var number))
        {
            writer.WriteNumberValue(number);
            return;
        }
        writer.WriteStringValue(Marshaling.ToStringSafe(cell));
    }

    private static object[,] ExpandToTable(JsonElement node, bool hasHeaderRow)
    {
        switch (node.ValueKind)
        {
            case JsonValueKind.Array:
                return ExpandArray(node, hasHeaderRow);
            case JsonValueKind.Object:
                return ExpandObject(node);
            default:
                return new object[1, 1] { { JsonToCell(node) } };
        }
    }

    private static object[,] ExpandArray(JsonElement array, bool hasHeaderRow)
    {
        var count = 0;
        var allObjects = true;
        foreach (var el in array.EnumerateArray())
        {
            count++;
            if (el.ValueKind != JsonValueKind.Object)
            {
                allObjects = false;
            }
        }
        if (count == 0)
        {
            return new object[1, 1] { { ExcelEmpty.Value } };
        }
        if (!allObjects)
        {
            var single = new object[count, 1];
            var i = 0;
            foreach (var el in array.EnumerateArray())
            {
                single[i++, 0] = JsonToCell(el);
            }
            return single;
        }

        var order = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var rows = new List<Dictionary<string, object>>(count);
        foreach (var el in array.EnumerateArray())
        {
            var dict = new Dictionary<string, object>(StringComparer.Ordinal);
            foreach (var prop in el.EnumerateObject())
            {
                if (seen.Add(prop.Name))
                {
                    order.Add(prop.Name);
                }
                dict[prop.Name] = JsonToCell(prop.Value);
            }
            rows.Add(dict);
        }
        return BuildObjectTable(order, rows, hasHeaderRow);
    }

    private static object[,] ExpandObject(JsonElement obj)
    {
        var count = 0;
        foreach (var _ in obj.EnumerateObject())
        {
            count++;
        }
        var result = new object[Math.Max(count, 1), 2];
        if (count == 0)
        {
            result[0, 0] = ExcelEmpty.Value;
            result[0, 1] = ExcelEmpty.Value;
            return result;
        }
        var i = 0;
        foreach (var prop in obj.EnumerateObject())
        {
            result[i, 0] = prop.Name;
            result[i, 1] = JsonToCell(prop.Value);
            i++;
        }
        return result;
    }

    private static object[,] BuildObjectTable(List<string> order, List<Dictionary<string, object>> rows, bool emitHeader)
    {
        var cols = Math.Max(order.Count, 1);
        var outRows = rows.Count + (emitHeader ? 1 : 0);
        var result = new object[outRows, cols];
        var rowOffset = 0;
        if (emitHeader)
        {
            for (var c = 0; c < cols; c++)
            {
                result[0, c] = c < order.Count ? order[c] : string.Empty;
            }
            rowOffset = 1;
        }
        for (var i = 0; i < rows.Count; i++)
        {
            var dict = rows[i];
            for (var c = 0; c < cols; c++)
            {
                result[rowOffset + i, c] = c < order.Count && dict.TryGetValue(order[c], out var v)
                    ? v
                    : ExcelEmpty.Value;
            }
        }
        return result;
    }

    private static bool TryNavigate(JsonElement root, string? path, out JsonElement result)
    {
        result = root;
        if (string.IsNullOrEmpty(path))
        {
            return true;
        }
        var p = path;
        if (p.StartsWith("$", StringComparison.Ordinal))
        {
            p = p.Substring(1);
        }
        var current = root;
        var i = 0;
        while (i < p.Length)
        {
            if (p[i] == '.')
            {
                i++;
                continue;
            }
            if (p[i] == '[')
            {
                var end = p.IndexOf(']', i + 1);
                if (end < 0
                    || !int.TryParse(p.AsSpan(i + 1, end - i - 1), System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var idx)
                    || idx < 0
                    || current.ValueKind != JsonValueKind.Array)
                {
                    result = default;
                    return false;
                }
                if (!TryGetArrayElement(current, idx, out current))
                {
                    result = default;
                    return false;
                }
                i = end + 1;
            }
            else
            {
                var start = i;
                while (i < p.Length && p[i] != '.' && p[i] != '[')
                {
                    i++;
                }
                var key = p.Substring(start, i - start);
                if (current.ValueKind != JsonValueKind.Object || !current.TryGetProperty(key, out current))
                {
                    result = default;
                    return false;
                }
            }
        }
        result = current;
        return true;
    }

    private static bool TryGetArrayElement(JsonElement array, int index, out JsonElement element)
    {
        var k = 0;
        foreach (var el in array.EnumerateArray())
        {
            if (k == index)
            {
                element = el;
                return true;
            }
            k++;
        }
        element = default;
        return false;
    }

    private static string? PathOf(object value)
    {
        if (Marshaling.IsBlankOrError(value))
        {
            return null;
        }
        var s = Marshaling.ToStringSafe(value);
        return s.Length == 0 ? null : s;
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
