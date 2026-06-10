using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelPerfToolkit;

/// <summary>
/// Result caching: an in-memory session cache (<c>EPT.MEMOIZE</c>, <c>EPT.CACHE.GET</c>,
/// <c>EPT.CACHE.CLEAR</c>) and a disk-backed cache that survives workbook reopen
/// (<c>EPT.DISKCACHE.WRITE/READ/CLEAR</c>). Both store whole blocks keyed by a string; build
/// content-addressed keys with the existing <c>EPT.HASHBLOCK</c> over the inputs.
///
/// <para><b>Honesty about Excel semantics.</b> Excel evaluates a UDF's arguments before
/// calling it, so <c>EPT.MEMOIZE(key, heavyFormula)</c> does not stop <c>heavyFormula</c> from
/// being computed - it stores the result so later <c>EPT.CACHE.GET</c> calls (and other cells)
/// can reuse it without recomputing, and so a disk cache can carry it across sessions. These
/// functions are stateful, so they are registered <c>IsThreadSafe = false</c> (they run on
/// Excel's calc thread) and are <b>not</b> volatile: a reader only re-reads when its key
/// argument changes.</para>
/// </summary>
public static class CacheUtilities
{
    private static readonly TraceSource TraceSource = ToolkitLifetime.CreateTraceSource("CacheUtilities");

    /// <summary>Soft cap on in-memory entries; the oldest entry is evicted past this.</summary>
    private const int MaxMemoryEntries = 4096;

    private static readonly ConcurrentDictionary<string, CacheEntry> Memory = new(StringComparer.Ordinal);
    private static long _stampSequence;

    private sealed class CacheEntry
    {
        public CacheEntry(object[,] value, long stamp)
        {
            Value = value;
            Stamp = stamp;
        }

        public object[,] Value { get; }

        public long Stamp { get; }
    }

    // ====================================================================
    // In-memory cache
    // ====================================================================

    /// <summary>
    /// Stores <paramref name="value"/> in the session cache under <paramref name="key"/> and
    /// returns it unchanged (a pass-through). Subsequent <see cref="CacheGet"/> calls with the
    /// same key return the stored block without recomputation.
    /// Thread-safety: NOT MTR-safe (stateful).
    /// </summary>
    [ExcelFunction(Name = "EPT.MEMOIZE", Description = "Store a value in the session cache under a key and return it (pass-through).", Category = "EPT.Cache", IsThreadSafe = false, IsVolatile = false)]
    public static object Memoize(
        [ExcelArgument(Name = "key")] string key,
        [ExcelArgument(Name = "value")] object value)
    {
        if (string.IsNullOrEmpty(key))
        {
            return ExcelError.ExcelErrorValue;
        }
        var stamp = Interlocked.Increment(ref _stampSequence);
        Memory[key] = new CacheEntry(Marshaling.AsArray2D(value), stamp);
        if (Memory.Count > MaxMemoryEntries)
        {
            EvictOldest();
        }
        return value;
    }

    /// <summary>
    /// Returns the cached block for <paramref name="key"/>, or <paramref name="ifMissing"/>
    /// when supplied, otherwise <c>#N/A</c>.
    /// Thread-safety: NOT MTR-safe (stateful).
    /// </summary>
    [ExcelFunction(Name = "EPT.CACHE.GET", Description = "Read a value from the session cache.", Category = "EPT.Cache", IsThreadSafe = false, IsVolatile = false)]
    public static object CacheGet(
        [ExcelArgument(Name = "key")] string key,
        [ExcelArgument(Name = "if_missing", Description = "Optional value returned on a miss. Defaults to #N/A.")] object ifMissing)
    {
        if (string.IsNullOrEmpty(key))
        {
            return ExcelError.ExcelErrorValue;
        }
        if (Memory.TryGetValue(key, out var entry))
        {
            return entry.Value;
        }
        return Marshaling.IsBlankOrError(ifMissing) ? ExcelError.ExcelErrorNA : ifMissing;
    }

    /// <summary>
    /// Clears one key (when supplied) or the entire session cache. Returns the number of
    /// entries removed.
    /// Thread-safety: NOT MTR-safe (stateful).
    /// </summary>
    [ExcelFunction(Name = "EPT.CACHE.CLEAR", Description = "Clear one key or the whole session cache; returns the count removed.", Category = "EPT.Cache", IsThreadSafe = false, IsVolatile = false)]
    public static object CacheClear(
        [ExcelArgument(Name = "key", Description = "Optional. Omit to clear everything.")] object key)
    {
        if (Marshaling.IsBlankOrError(key))
        {
            var n = Memory.Count;
            Memory.Clear();
            return (double)n;
        }
        return Memory.TryRemove(Marshaling.ToStringSafe(key), out _) ? 1d : 0d;
    }

    /// <summary>Entries evicted per scan once the cap is exceeded.</summary>
    private const int EvictBatch = 64;

    private static void EvictOldest()
    {
        // Evict a batch per scan: a single full enumeration amortized over 64 inserts
        // instead of an O(N) walk on every Memoize past the cap.
        var entries = new List<(long Stamp, string Key)>(Memory.Count);
        foreach (var kv in Memory)
        {
            entries.Add((kv.Value.Stamp, kv.Key));
        }
        entries.Sort(static (a, b) => a.Stamp.CompareTo(b.Stamp));
        var toRemove = Math.Min(EvictBatch, entries.Count);
        for (var i = 0; i < toRemove; i++)
        {
            Memory.TryRemove(entries[i].Key, out _);
        }
    }

    // ====================================================================
    // Disk cache
    // ====================================================================

    /// <summary>
    /// Persists <paramref name="block"/> to the on-disk cache under <paramref name="key"/> and
    /// returns the row count. The value survives workbook reopen.
    /// Thread-safety: NOT MTR-safe (disk + async bridge).
    /// </summary>
    [ExcelFunction(Name = "EPT.DISKCACHE.WRITE", Description = "Persist a block to the disk cache under a key. Returns row count.", Category = "EPT.Cache", IsThreadSafe = false, IsVolatile = false)]
    public static object DiskWrite(
        [ExcelArgument(Name = "key")] string key,
        [ExcelArgument(Name = "block")] object[,] block)
    {
        if (string.IsNullOrEmpty(key))
        {
            return ExcelError.ExcelErrorValue;
        }
        ArgumentNullException.ThrowIfNull(block);
        try
        {
            WriteDiskAsync(DiskPath(key, ensureDirectory: true), block, ToolkitLifetime.ShutdownToken)
                .ConfigureAwait(false)
                .GetAwaiter()
                .GetResult();
            return (double)block.GetLength(0);
        }
        catch (OperationCanceledException)
        {
            return ExcelError.ExcelErrorNA;
        }
        catch (Exception ex)
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 1, "EPT.DISKCACHE.WRITE failed: {0}", ex.Message);
            return ExcelError.ExcelErrorValue;
        }
    }

    /// <summary>
    /// Reads the block stored under <paramref name="key"/> from the disk cache, or returns
    /// <paramref name="ifMissing"/> when supplied (otherwise <c>#N/A</c>).
    /// Thread-safety: NOT MTR-safe (disk + async bridge).
    /// </summary>
    [ExcelFunction(Name = "EPT.DISKCACHE.READ", Description = "Read a block from the disk cache.", Category = "EPT.Cache", IsThreadSafe = false, IsVolatile = false)]
    public static object DiskRead(
        [ExcelArgument(Name = "key")] string key,
        [ExcelArgument(Name = "if_missing", Description = "Optional value returned on a miss. Defaults to #N/A.")] object ifMissing)
    {
        if (string.IsNullOrEmpty(key))
        {
            return ExcelError.ExcelErrorValue;
        }
        try
        {
            var block = ReadDiskAsync(DiskPath(key), ToolkitLifetime.ShutdownToken)
                .ConfigureAwait(false)
                .GetAwaiter()
                .GetResult();
            if (block is not null)
            {
                return block;
            }
            return Marshaling.IsBlankOrError(ifMissing) ? ExcelError.ExcelErrorNA : ifMissing;
        }
        catch (OperationCanceledException)
        {
            return ExcelError.ExcelErrorNA;
        }
        catch (Exception ex)
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 2, "EPT.DISKCACHE.READ failed: {0}", ex.Message);
            return ExcelError.ExcelErrorValue;
        }
    }

    /// <summary>
    /// Deletes one key (when supplied) or every entry in the disk cache. Returns the number of
    /// files removed.
    /// Thread-safety: NOT MTR-safe (disk).
    /// </summary>
    [ExcelFunction(Name = "EPT.DISKCACHE.CLEAR", Description = "Delete one key or the whole disk cache; returns the count removed.", Category = "EPT.Cache", IsThreadSafe = false, IsVolatile = false)]
    public static object DiskClear(
        [ExcelArgument(Name = "key", Description = "Optional. Omit to clear everything.")] object key)
    {
        try
        {
            if (Marshaling.IsBlankOrError(key))
            {
                var dir = CacheDirectory();
                if (!Directory.Exists(dir))
                {
                    return 0d;
                }
                var files = Directory.GetFiles(dir, "*.json");
                foreach (var f in files)
                {
                    File.Delete(f);
                }
                return (double)files.Length;
            }
            var path = DiskPath(Marshaling.ToStringSafe(key));
            if (File.Exists(path))
            {
                File.Delete(path);
                return 1d;
            }
            return 0d;
        }
        catch (Exception ex)
        {
            TraceSource.TraceEvent(TraceEventType.Warning, 3, "EPT.DISKCACHE.CLEAR failed: {0}", ex.Message);
            return ExcelError.ExcelErrorValue;
        }
    }

    // ====================================================================
    // Disk serialization (typed JSON round-trip)
    // ====================================================================

    private static async Task WriteDiskAsync(string path, object[,] block, CancellationToken cancellationToken)
    {
        // Write to a sibling temp file and atomically swap it in. Truncate-in-place
        // destroyed the previous good entry the moment the stream opened, so any
        // cancel/crash mid-write left a torn file that poisoned every later read of
        // that key. File.Move(overwrite) is an atomic same-volume replace: readers
        // (including other Excel instances) see the old or the new entry, never a mix.
        var tmp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            await using (var stream = new FileStream(
                tmp,
                FileMode.CreateNew,
                FileAccess.Write,
                FileShare.None,
                bufferSize: DirectFileIO.DefaultBufferSize,
                options: FileOptions.Asynchronous | FileOptions.SequentialScan))
            await using (var writer = new Utf8JsonWriter(stream))
            {
                var rows = block.GetLength(0);
                var cols = block.GetLength(1);
                writer.WriteStartObject();
                writer.WriteNumber("rows", rows);
                writer.WriteNumber("cols", cols);
                writer.WriteStartArray("cells");
                for (var r = 0; r < rows; r++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    writer.WriteStartArray();
                    for (var c = 0; c < cols; c++)
                    {
                        WriteTypedCell(writer, block[r, c]);
                    }
                    writer.WriteEndArray();
                    if (writer.BytesPending > DirectFileIO.DefaultBufferSize)
                    {
                        await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
                    }
                }
                writer.WriteEndArray();
                writer.WriteEndObject();
                await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
            }
            File.Move(tmp, path, overwrite: true);
        }
        catch
        {
            try
            {
                File.Delete(tmp);
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
            throw;
        }
    }

    private static async Task<object[,]?> ReadDiskAsync(string path, CancellationToken cancellationToken)
    {
        if (!File.Exists(path))
        {
            return null;
        }
        JsonDocument doc;
        try
        {
            await using var stream = new FileStream(
                path,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                bufferSize: DirectFileIO.DefaultBufferSize,
                options: FileOptions.Asynchronous | FileOptions.SequentialScan);
            doc = await JsonDocument.ParseAsync(stream, default, cancellationToken).ConfigureAwait(false);
        }
        catch (FileNotFoundException)
        {
            return null;
        }
        catch (JsonException)
        {
            // A torn or corrupt entry is a MISS, not an error: callers fall back to
            // ifMissing instead of returning #VALUE! forever. Delete it so the cache
            // heals rather than re-parsing the same damage on every read.
            try
            {
                File.Delete(path);
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
            return null;
        }
        using var document = doc;
        var root = document.RootElement;
        var rows = root.GetProperty("rows").GetInt32();
        var cols = root.GetProperty("cols").GetInt32();
        if (rows < 0 || cols < 0 || (long)rows * cols > int.MaxValue)
        {
            throw new InvalidDataException("Corrupt disk-cache entry: invalid dimensions.");
        }
        var result = new object[rows, cols];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                result[r, c] = ExcelEmpty.Value;
            }
        }
        var ri = 0;
        foreach (var rowEl in root.GetProperty("cells").EnumerateArray())
        {
            if (ri >= rows)
            {
                break;
            }
            var ci = 0;
            foreach (var cellEl in rowEl.EnumerateArray())
            {
                if (ci >= cols)
                {
                    break;
                }
                result[ri, ci] = ReadTypedCell(cellEl);
                ci++;
            }
            ri++;
        }
        return result;
    }

    private static void WriteTypedCell(Utf8JsonWriter writer, object? cell)
    {
        if (cell is ExcelError err)
        {
            writer.WriteStartObject();
            writer.WriteNumber("$err", (int)err);
            writer.WriteEndObject();
            return;
        }
        if (cell is null or ExcelEmpty or ExcelMissing)
        {
            writer.WriteNullValue();
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

    private static object ReadTypedCell(JsonElement element) => element.ValueKind switch
    {
        JsonValueKind.Number => element.GetDouble(),
        JsonValueKind.String => element.GetString() ?? string.Empty,
        JsonValueKind.True => true,
        JsonValueKind.False => false,
        JsonValueKind.Null => ExcelEmpty.Value,
        // C# enum casts do not range-check; an undefined ExcelError payload would be
        // handed to Excel's marshaler as an undefined error state.
        JsonValueKind.Object => element.TryGetProperty("$err", out var ep)
                && ep.TryGetInt32(out var ev)
                && Enum.IsDefined(typeof(ExcelError), ev)
            ? (ExcelError)ev
            : (object)element.GetRawText(),
        _ => ExcelEmpty.Value,
    };

    // ====================================================================
    // Cache directory + key hashing
    // ====================================================================

    private static string CacheDirectory()
        => Path.Combine(Path.GetTempPath(), "ExcelPerfToolkit", "diskcache");

    private static string DiskPath(string key, bool ensureDirectory = false)
    {
        var dir = CacheDirectory();
        if (ensureDirectory)
        {
            // Reads and clears don't need the metadata syscall: a missing directory
            // simply means a miss / nothing to delete.
            Directory.CreateDirectory(dir);
        }
        var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(key)));
        return Path.Combine(dir, hash + ".json");
    }
}
