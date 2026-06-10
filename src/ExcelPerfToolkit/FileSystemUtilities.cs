using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelPerfToolkit;

/// <summary>
/// Round 9 filesystem utilities: file metadata (<c>EPT.FILEINFO</c>) and bulk folder ingest
/// (<c>EPT.READFOLDER</c>). Both work directly against <see cref="System.IO"/> - no workbook
/// is ever opened and the Excel object model is never touched. <c>EPT.READFOLDER</c> fans out
/// to the existing readers (<see cref="DirectFileIO.ReadDelimitedAsync"/>,
/// <see cref="JsonUtilities.ReadJsonAsync"/>, <see cref="JsonUtilities.ReadNdjsonAsync"/>) by
/// file extension and concatenates the results, aligning columns by header name.
///
/// <para>Both are registered <c>IsThreadSafe = false</c> because they touch the disk; the
/// async folder reader honors <see cref="ToolkitLifetime.ShutdownToken"/>.</para>
/// </summary>
public static class FileSystemUtilities
{
    private static readonly TraceSource TraceSource = ToolkitLifetime.CreateTraceSource("FileSystemUtilities");

    private static readonly string[] InfoHeader =
    {
        "Path", "Exists", "IsDirectory", "SizeBytes", "Modified", "Created", "Extension", "Name",
    };

    /// <summary>
    /// Returns a metadata table for one or more paths: a header row followed by one row per
    /// path with columns <c>Path, Exists, IsDirectory, SizeBytes, Modified, Created,
    /// Extension, Name</c>. <c>Modified</c>/<c>Created</c> are local date serials. Missing
    /// paths report <c>Exists = FALSE</c> with blank metadata; a per-path error never fails
    /// the whole call.
    /// Marshaling cost: O(1). Thread-safety: NOT MTR-safe (touches the disk).
    /// </summary>
    [ExcelFunction(Name = "EPT.FILEINFO", Description = "Metadata table for one or more file/folder paths.", Category = "EPT.FileSystem", IsThreadSafe = false, IsVolatile = false)]
    public static object[,] FileInfoFn(
        [ExcelArgument(Name = "paths", Description = "A path or a block of paths.")] object paths)
    {
        var block = Marshaling.AsArray2D(paths);
        var list = new List<string>();
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                if (!Marshaling.IsBlankOrError(block[r, c]))
                {
                    list.Add(Marshaling.ToStringSafe(block[r, c]));
                }
            }
        }

        var result = new object[list.Count + 1, InfoHeader.Length];
        for (var c = 0; c < InfoHeader.Length; c++)
        {
            result[0, c] = InfoHeader[c];
        }
        for (var i = 0; i < list.Count; i++)
        {
            WriteInfoRow(result, i + 1, list[i]);
        }
        return result;
    }

    private static void WriteInfoRow(object[,] result, int row, string path)
    {
        // Columns: Path, Exists, IsDirectory, SizeBytes, Modified, Created, Extension, Name.
        result[row, 0] = path;
        result[row, 1] = false;
        result[row, 2] = false;
        result[row, 3] = ExcelEmpty.Value;
        result[row, 4] = ExcelEmpty.Value;
        result[row, 5] = ExcelEmpty.Value;
        result[row, 6] = ExcelEmpty.Value;
        result[row, 7] = ExcelEmpty.Value;
        try
        {
            var isDir = Directory.Exists(path);
            var isFile = File.Exists(path);
            result[row, 1] = isDir || isFile;
            result[row, 2] = isDir;
            if (isFile)
            {
                var fi = new FileInfo(path);
                result[row, 3] = (double)fi.Length;
                result[row, 4] = fi.LastWriteTime.ToOADate();
                result[row, 5] = fi.CreationTime.ToOADate();
                result[row, 6] = fi.Extension;
                result[row, 7] = fi.Name;
            }
            else if (isDir)
            {
                var di = new DirectoryInfo(path);
                result[row, 4] = di.LastWriteTime.ToOADate();
                result[row, 5] = di.CreationTime.ToOADate();
                result[row, 6] = string.Empty;
                result[row, 7] = di.Name;
            }
            else
            {
                result[row, 6] = Path.GetExtension(path);
                result[row, 7] = Path.GetFileName(path);
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException or PathTooLongException)
        {
            // Leave the row as "does not exist / unknown"; one bad path must not fail the rest.
            TraceSource.TraceEvent(TraceEventType.Information, 1, "EPT.FILEINFO could not stat '{0}': {1}", path, ex.Message);
        }
    }

    /// <summary>
    /// Reads every file in <paramref name="folder"/> matching <paramref name="pattern"/> and
    /// concatenates them into one table. The reader is chosen per file by extension
    /// (<c>.csv</c>/<c>.tsv</c> delimited, <c>.json</c> document, <c>.ndjson</c>/<c>.jsonl</c>
    /// newline-delimited; anything else is treated as CSV). With <paramref name="hasHeaderRow"/>
    /// true (the default), each file's first row is treated as a header and rows are aligned to
    /// the union of all headers (in first-seen order) with a single header row emitted;
    /// otherwise rows are stacked and padded to the widest file. Async; no object model touched.
    /// </summary>
    public static async Task<object[,]> ReadFolderAsync(string folder, string pattern, bool recursive, bool hasHeaderRow, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(folder))
        {
            throw new ArgumentException("Folder is required.", nameof(folder));
        }
        if (!Directory.Exists(folder))
        {
            throw new DirectoryNotFoundException($"Folder not found: {folder}");
        }
        var searchPattern = string.IsNullOrEmpty(pattern) ? "*" : pattern;
        var option = recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
        var files = Directory.GetFiles(folder, searchPattern, option);
        Array.Sort(files, StringComparer.OrdinalIgnoreCase);
        if (files.Length == 0)
        {
            return new object[1, 1] { { ExcelEmpty.Value } };
        }

        var blocks = new List<object[,]>(files.Length);
        foreach (var file in files)
        {
            cancellationToken.ThrowIfCancellationRequested();
            blocks.Add(await ReadOneAsync(file, hasHeaderRow, cancellationToken).ConfigureAwait(false));
        }

        return hasHeaderRow ? ConcatByHeader(blocks) : ConcatStacked(blocks);
    }

    private static Task<object[,]> ReadOneAsync(string path, bool hasHeaderRow, CancellationToken cancellationToken)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        return ext switch
        {
            ".tsv" => DirectFileIO.ReadDelimitedAsync(path, '\t', null, true, cancellationToken),
            ".json" => JsonUtilities.ReadJsonAsync(path, null, hasHeaderRow, cancellationToken),
            ".ndjson" or ".jsonl" => JsonUtilities.ReadNdjsonAsync(path, hasHeaderRow, cancellationToken),
            _ => DirectFileIO.ReadDelimitedAsync(path, ',', null, true, cancellationToken),
        };
    }

    private static object[,] ConcatByHeader(List<object[,]> blocks)
    {
        var order = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var dataRows = 0;
        foreach (var b in blocks)
        {
            var br = b.GetLength(0);
            var bc = b.GetLength(1);
            if (br >= 1)
            {
                for (var c = 0; c < bc; c++)
                {
                    var name = Marshaling.ToStringSafe(b[0, c]);
                    if (seen.Add(name))
                    {
                        order.Add(name);
                    }
                }
                dataRows += br - 1;
            }
        }

        var cols = Math.Max(order.Count, 1);
        var result = new object[dataRows + 1, cols];
        for (var r = 0; r <= dataRows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                result[r, c] = ExcelEmpty.Value;
            }
        }
        for (var c = 0; c < order.Count; c++)
        {
            result[0, c] = order[c];
        }

        var outRow = 1;
        foreach (var b in blocks)
        {
            var br = b.GetLength(0);
            var bc = b.GetLength(1);
            if (br < 1)
            {
                continue;
            }
            var map = new int[bc];
            for (var c = 0; c < bc; c++)
            {
                map[c] = order.IndexOf(Marshaling.ToStringSafe(b[0, c]));
            }
            for (var r = 1; r < br; r++)
            {
                for (var c = 0; c < bc; c++)
                {
                    var oc = map[c];
                    if (oc >= 0)
                    {
                        result[outRow, oc] = b[r, c];
                    }
                }
                outRow++;
            }
        }
        return result;
    }

    private static object[,] ConcatStacked(List<object[,]> blocks)
    {
        var total = 0;
        var maxCols = 1;
        foreach (var b in blocks)
        {
            total += b.GetLength(0);
            if (b.GetLength(1) > maxCols)
            {
                maxCols = b.GetLength(1);
            }
        }
        var result = new object[total, maxCols];
        var outRow = 0;
        foreach (var b in blocks)
        {
            var br = b.GetLength(0);
            var bc = b.GetLength(1);
            for (var r = 0; r < br; r++)
            {
                for (var c = 0; c < maxCols; c++)
                {
                    result[outRow, c] = c < bc ? b[r, c] : ExcelEmpty.Value;
                }
                outRow++;
            }
        }
        return result;
    }

    /// <summary>
    /// Worksheet UDF: read and concatenate a folder of files. Marshaling cost: 1 write
    /// crossing. Thread-safety: NOT MTR-safe (async bridge + disk).
    /// </summary>
    [ExcelFunction(Name = "EPT.READFOLDER", Description = "Read and concatenate matching files in a folder (CSV/TSV/JSON/NDJSON). Async, no COM.", Category = "EPT.FileSystem", IsThreadSafe = false, IsVolatile = false)]
    public static object ReadFolderUdf(
        [ExcelArgument(Name = "folder")] string folder,
        [ExcelArgument(Name = "pattern", Description = "Glob like *.csv. Default *.")] object pattern,
        [ExcelArgument(Name = "recursive", Description = "TRUE to recurse subfolders. Default FALSE.")] object recursive,
        [ExcelArgument(Name = "has_header_row", Description = "TRUE (default) aligns files by header name.")] object hasHeaderRow)
    {
        try
        {
            return ReadFolderAsync(
                    folder,
                    PatternOf(pattern),
                    ResolveBool(recursive, false),
                    ResolveBool(hasHeaderRow, true),
                    ToolkitLifetime.ShutdownToken)
                .ConfigureAwait(false)
                .GetAwaiter()
                .GetResult();
        }
        catch (OperationCanceledException)
        {
            TraceSource.TraceEvent(TraceEventType.Information, 2, "EPT.READFOLDER cancelled for {0}", folder);
            return ExcelError.ExcelErrorNA;
        }
        catch (DirectoryNotFoundException ex)
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 3, "EPT.READFOLDER folder not found: {0}", ex.Message);
            return ExcelError.ExcelErrorNA;
        }
        catch (Exception ex)
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 4, "EPT.READFOLDER failed: {0}", ex.Message);
            return ExcelError.ExcelErrorValue;
        }
    }

    /// <summary>
    /// Live RTD cell that increments whenever the watched file changes (create/change/delete/
    /// rename). Use it as a recalculation trigger - e.g. wrap an <c>EPT.READCSV</c> in a
    /// formula that also references this cell so the import re-runs when the file is rewritten.
    /// Backed by a shared <see cref="FileSystemWatcher"/> feed; bursts are coalesced by the RTD
    /// server's throttle. Returns <c>#N/A</c> if the containing folder does not exist.
    /// Thread-safety: SAFE for MTR (hands the spec to the RTD server).
    /// </summary>
    [ExcelFunction(Name = "EPT.WATCHFILE", Description = "Live change counter for a file (RTD trigger).", Category = "EPT.FileSystem", IsThreadSafe = true, IsVolatile = false)]
    public static object WatchFile([ExcelArgument(Name = "path")] string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return ExcelError.ExcelErrorValue;
        }
        return XlCall.RTD(RtdServer.ProgIdValue, null, "watchfile:" + path) ?? ExcelError.ExcelErrorNA;
    }

    /// <summary>
    /// Live RTD cell that increments whenever any file in the watched folder changes
    /// (non-recursive). See <see cref="WatchFile"/> for usage and semantics.
    /// Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.WATCHFOLDER", Description = "Live change counter for a folder (RTD trigger).", Category = "EPT.FileSystem", IsThreadSafe = true, IsVolatile = false)]
    public static object WatchFolder([ExcelArgument(Name = "path")] string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return ExcelError.ExcelErrorValue;
        }
        return XlCall.RTD(RtdServer.ProgIdValue, null, "watchfolder:" + path) ?? ExcelError.ExcelErrorNA;
    }

    private static string PatternOf(object pattern)
    {
        if (Marshaling.IsBlankOrError(pattern))
        {
            return "*";
        }
        var s = Marshaling.ToStringSafe(pattern);
        return s.Length == 0 ? "*" : s;
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
