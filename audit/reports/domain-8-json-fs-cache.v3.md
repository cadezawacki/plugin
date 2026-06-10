# Domain 8, Round 3 — JSON / FileSystem / WatchFeeds / Cache audit

Scope: `JsonUtilities.cs`, `FileSystemUtilities.cs`, `WatchFeeds.cs`, `CacheUtilities.cs`.
Contract context consulted (no findings reported in): `RtdServer.cs`, `ToolkitLifetime.cs`, `DirectFileIO.cs`, `Marshaling.cs`, `readme.md`.

All BCL behavior claims marked **[verified]** were confirmed empirically on .NET SDK 8.0.422 (`/tmp/probe8`,
Linux host; Windows-only semantics reasoned from documentation and cited). Probe transcript at the end of
Pass 1.

---

## Summary table

| ID | Severity | One-line |
|---|---|---|
| BUG-HIGH-1 | HIGH | `EPT.READFOLDER`: one locked/malformed file aborts the entire folder ingest with `#VALUE!` |
| BUG-HIGH-2 | HIGH | `EPT.READFOLDER` recursive: one inaccessible subdirectory throws `UnauthorizedAccessException` and kills the whole call |
| BUG-HIGH-3 | HIGH | `EPT.READJSON` (and READFOLDER `.json`): UTF-8 BOM files fail to parse → `#VALUE!` |
| BUG-MEDIUM-4 | MEDIUM | Recursive enumeration follows reparse-point/symlink cycles: files duplicated ~40× (verified) / multi-minute hang on Windows junction loops |
| BUG-MEDIUM-5 | MEDIUM | `JsonToCell`: int64 > 2^53 silently rounded; out-of-range numbers become `Infinity`; `GetRawText` fallback is dead code |
| BUG-MEDIUM-6 | MEDIUM | `EPT.PARSEJSON`: omitted `has_header_row` arrives as FALSE — plain Excel-DNA ignores the C# `= true` default |
| BUG-MEDIUM-7 | MEDIUM | Disk cache: truncate-in-place write; cancel/crash destroys the previous entry and poisons all future reads (`#VALUE!`, not a miss) |
| BUG-MEDIUM-8 | MEDIUM | `ConcatByHeader` treats row 0 of *every* block as a header: single-object/scalar/empty JSON & empty files corrupt the union and silently drop data |
| BUG-MEDIUM-9 | MEDIUM | Watch feeds: fatal `FileSystemWatcher` error (watched dir deleted) leaves a dead watcher parked forever — counter silently freezes |
| BUG-LOW-10 | LOW | `WatchFileFeed` with a bare filename watches `"."` = the Excel process CWD (effectively random) |
| BUG-LOW-11 | LOW | `EPT.READNDJSON`: one bad line fails the whole file and the trace reports `LineNumber: 0` (per-line parse loses file position) |
| BUG-LOW-12 | LOW | `ReadTypedCell`: unvalidated `(ExcelError)ev` cast lets a corrupt cache file box an out-of-range enum into Excel |
| BUG-LOW-13 | LOW | `ConcatByHeader`: duplicate column names within one file collapse onto one output column, last value wins, data silently dropped |
| BUG-LOW-14 | LOW | Disk cache: unbounded growth in `%TEMP%` with no TTL/eviction; memory cache is count-capped but byte-unbounded |
| MEM-1 | — | NDJSON/array-of-objects expansion allocates a `Dictionary<string,object>` per row: ~60 MB transient for 100k×10 → ~10 MB with indexed rows |
| MEM-2 | — | `File.ReadAllBytesAsync` materializes whole JSON files on the LOH (10 MB/call garbage) → `JsonDocument.ParseAsync(Stream)` uses pooled buffers |
| MEM-3 | — | `TryNavigate` re-parses the path per cell: ~100 B substring churn × cells (~10 MB per 100k cells) |
| MEM-4 | — | READFOLDER holds every per-file block until concat: ~2× peak retention of the result's reference arrays |
| PERF-1 | — | JSONPATH: hoist path compilation out of the per-cell loop (navigation phase 3–5×; whole call ~1.1×) |
| PERF-2 | — | READFOLDER reads files strictly serially: bounded parallelism gives 4–8× wall clock on many small files |
| PERF-3 | — | `ConcatByHeader` uses `List.IndexOf` per column per file: O(files × cols × union) → `Dictionary<string,int>` is ~40× on that phase |
| PERF-4 | — | `EPT.FILEINFO` performs 3 stat calls per path where 1–2 suffice (≈3× on network shares) |
| PERF-5 | — | `BuildObjectTable` does rows×cols string-hash lookups: ~1M for 100k×10 (30–60 ms) eliminated by MEM-1's indexed rows |
| PERF-6 | — | `EvictOldest` full O(4096) scan per insert at cap; `DiskPath` calls `Directory.CreateDirectory` on every read |
| ARCH-1 | — | READFOLDER resilience: per-file isolation, safe enumeration options, optional source/status column, bounded parallelism |
| ARCH-2 | — | Shared atomic-write helper (temp + `File.Move(overwrite)`) for disk cache and JSON/CSV writers |
| ARCH-3 | — | Cache governance: byte-budgeted memory cache, disk cache TTL/pruning, corrupt-entry-as-miss semantics |
| ARCH-4 | — | Self-healing watch feeds: re-arm loop on watcher error / late-created directories |

---

## Pass 1 — Deep bug scan

### BUG-HIGH-1
```
File: /home/user/plugin/src/ExcelPerfToolkit/FileSystemUtilities.cs:145-152 (and 155-165)
Category: logic_error
Description: ReadFolderAsync has no per-file fault isolation. Any exception from any single file —
  IOException (sharing violation on a file open in another app), JsonException (one malformed .json),
  InvalidDataException (CSV with an unterminated quote, thrown by DirectFileIO.ReadDelimitedAsync:185) —
  propagates out of the per-file loop and fails the entire EPT.READFOLDER call with #VALUE!. This
  directly contradicts the standard set by EPT.FILEINFO in the same file, whose doc comment (line 35-36)
  promises "a per-path error never fails the whole call".
Trigger condition: =EPT.READFOLDER("C:\data", "*.csv") over 1,000 files while one of them is held open
  by another process with a write lock (e.g. the producer app is mid-write, or the file is an
  Excel-locked .csv). Also: one .json file with a trailing comma in a folder of 500 valid ones.
Trace:
  1. ReadFolderAsync:138 enumerates 1,000 files; loop at :146 awaits ReadOneAsync per file.
  2. File #371 is open elsewhere with FileShare.None. ReadOneAsync:163 → DirectFileIO.ReadDelimitedAsync
     → new FileStream(..., FileShare.Read) (DirectFileIO.cs:65-73) throws IOException
     ("being used by another process") on Windows.
  3. Exception leaves the foreach at FileSystemUtilities.cs:149, unwinds through ReadFolderAsync.
  4. ReadFolderUdf:297 `catch (Exception ex)` → returns ExcelError.ExcelErrorValue.
  5. 999 successfully readable files are discarded; the user sees a bare #VALUE! with no indication
     of which file failed (the trace message has the folder, not the file).
Fix (FileSystemUtilities.cs:146-151), before:
        foreach (var file in files)
        {
            cancellationToken.ThrowIfCancellationRequested();
            blocks.Add(await ReadOneAsync(file, hasHeaderRow, cancellationToken).ConfigureAwait(false));
        }
  after:
        foreach (var file in files)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                blocks.Add(await ReadOneAsync(file, hasHeaderRow, cancellationToken).ConfigureAwait(false));
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException or InvalidDataException)
            {
                TraceSource.TraceEvent(TraceEventType.Warning, 5,
                    "EPT.READFOLDER skipped unreadable file '{0}': {1}", file, ex.Message);
            }
        }
  (Optionally surface skipped count — see ARCH-1.)
```

### BUG-HIGH-2
```
File: /home/user/plugin/src/ExcelPerfToolkit/FileSystemUtilities.cs:137-138
Category: logic_error
Description: Directory.GetFiles(folder, pattern, SearchOption.AllDirectories) uses the legacy-compatible
  enumeration options (EnumerationOptions.Compatible: IgnoreInaccessible = false, AttributesToSkip = 0 —
  documented at learn.microsoft.com/dotnet/api/system.io.enumerationoptions.compatible). The first
  subdirectory the user cannot enter throws UnauthorizedAccessException and aborts the entire
  enumeration before a single file is read.
  [verified] Probe G2 (run as unprivileged user, .NET 8.0.422): GetFiles over a tree containing one
  mode-000 subdir → "THROWS UnauthorizedAccessException: Access to the path '/tmp/probe8-root/locked'
  is denied." With EnumerationOptions { IgnoreInaccessible = true } the same tree enumerates fine
  (probe M2: "OK, 1 files").
Trigger condition: =EPT.READFOLDER("D:\", "*.csv", TRUE) or any recursive scan that crosses
  "System Volume Information", "$RECYCLE.BIN", a DPAPI-protected profile subfolder, or any ACL-denied
  directory — extremely common the moment `recursive` is TRUE on anything but a private folder.
Trace:
  1. ReadFolderAsync:138 → Directory.GetFiles(folder, "*", SearchOption.AllDirectories).
  2. The enumerator descends into D:\System Volume Information → CreateFileW/FindFirstFileW fails
     with ERROR_ACCESS_DENIED → UnauthorizedAccessException (IgnoreInaccessible=false).
  3. GetFiles throws before returning any array; no file has been read.
  4. ReadFolderUdf:297 generic catch → #VALUE! for the whole call.
Fix (FileSystemUtilities.cs:137-138), before:
        var option = recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
        var files = Directory.GetFiles(folder, searchPattern, option);
  after:
        var options = new EnumerationOptions
        {
            RecurseSubdirectories = recursive,
            IgnoreInaccessible = true,
            AttributesToSkip = FileAttributes.ReparsePoint, // also fixes BUG-MEDIUM-4
            MatchType = MatchType.Win32,                    // keep DOS-style pattern semantics
        };
        var files = Directory.GetFiles(folder, searchPattern, options);
```

### BUG-HIGH-3
```
File: /home/user/plugin/src/ExcelPerfToolkit/JsonUtilities.cs:128-129
Category: logic_error
Description: ReadJsonAsync loads the file with File.ReadAllBytesAsync and parses with
  JsonDocument.Parse(ReadOnlyMemory<byte>). That overload does NOT tolerate a UTF-8 BOM.
  [verified] Probe A (.NET 8.0.422): JsonDocument.Parse(bytes-with-EF-BB-BF) → JsonException
  "'0xEF' is an invalid start of a value". Probe B: JsonDocument.Parse(Stream) on the identical
  bytes succeeds (the stream-based path skips the BOM). So every UTF-8-with-BOM JSON file —
  Windows Notepad "UTF-8 with BOM", Visual Studio default for some templates, PowerShell 5.1
  `Out-File -Encoding utf8`, many export tools — makes EPT.READJSON return #VALUE!
  ("invalid JSON") even though the file is perfectly valid. EPT.READFOLDER inherits the failure for
  every `.json` file (and, combined with BUG-HIGH-1, one BOM'd file kills the whole folder ingest).
  Note the inconsistency: ReadNdjsonAsync:156 uses StreamReader with detectEncodingFromByteOrderMarks
  and handles BOMs fine — the two readers disagree on the same byte sequence.
Trigger condition: =EPT.READJSON("C:\data\export.json") where the file was produced by
  `ConvertTo-Json ... | Out-File -Encoding utf8` under Windows PowerShell 5.1.
Trace:
  1. ReadJsonAsync:128 reads N bytes; bytes[0..2] = EF BB BF.
  2. :129 JsonDocument.Parse(bytes) → Utf8JsonReader sees 0xEF as the first token byte →
     JsonException at LineNumber 0, BytePositionInLine 0.
  3. RunFileRead:355 catches JsonException → returns ExcelError.ExcelErrorValue.
  4. User sees #VALUE! and a trace line claiming "invalid JSON" for a valid document.
Fix (JsonUtilities.cs:128-129), before:
        var bytes = await File.ReadAllBytesAsync(path, cancellationToken).ConfigureAwait(false);
        using var doc = JsonDocument.Parse(bytes);
  after:
        await using var stream = new FileStream(
            path, FileMode.Open, FileAccess.Read, FileShare.Read,
            bufferSize: DirectFileIO.DefaultBufferSize,
            options: FileOptions.Asynchronous | FileOptions.SequentialScan);
        using var doc = await JsonDocument.ParseAsync(stream, default, cancellationToken).ConfigureAwait(false);
  (Also removes the LOH allocation — see MEM-2. The stream path additionally accepts UTF-16/32 BOMs? No —
  it only skips the UTF-8 BOM; UTF-16 files still fail, acceptable to document.)
```

### BUG-MEDIUM-4
```
File: /home/user/plugin/src/ExcelPerfToolkit/FileSystemUtilities.cs:137-138
Category: logic_error
Description: The same Compatible enumeration options recurse INTO reparse points (junctions,
  directory symlinks) with no cycle detection (AttributesToSkip = 0). A self-referencing junction —
  trivially created by `mklink /J`, and present in every Windows profile as the legacy
  "Application Data" → "AppData\Roaming" junction chain — causes the enumerator to revisit the same
  tree at ever-deeper paths.
  [verified] Probe F (.NET 8.0.422, Linux): a folder with 1 real file and a self-symlink returned
  42 files — the single f.csv was enumerated 41 times (once per nesting level until the OS's
  40-deep symlink-resolution limit, ELOOP, cut it off). On Windows there is no ELOOP equivalent for
  junctions and .NET uses \\?\-prefixed paths (~32k char limit), so a 4-char-name junction loop
  recurses thousands of levels: every file in the folder is returned thousands of times, the calc
  thread stalls for minutes inside GetFiles, and the result block (if it ever materializes)
  contains massively duplicated rows. Known BCL behavior: dotnet/runtime#18243 — recursive
  enumeration follows junctions/symlinks, callers must skip FileAttributes.ReparsePoint themselves.
Trigger condition: =EPT.READFOLDER("C:\Users\me", "*.csv", TRUE) — the profile contains junction
  chains; or any folder where a tool created a looping junction.
Trace:
  1. GetFiles descends loopdir → loopdir\self → loopdir\self\self → ... appending one level per pass.
  2. Each level re-yields every matching file with a longer path; files array grows linearly with depth.
  3. On Linux, depth 41 fails path resolution (ELOOP) and enumeration stops: 42 entries for 1 file.
     On Windows the loop continues until the extended path limit; wall clock is minutes, memory is
     O(depth × files), and ConcatByHeader then concatenates the duplicates as real data rows.
Fix: same change as BUG-HIGH-2 — AttributesToSkip = FileAttributes.ReparsePoint in EnumerationOptions.
  [verified] Probe F2: with ReparsePoint skipped the looped tree returns exactly 1 file in 0 ms.
  (Cost: legitimately symlinked files are skipped too; document it. If symlink-following matters,
  cycle-detect on (volume serial, file index) — not worth it here.)
```

### BUG-MEDIUM-5
```
File: /home/user/plugin/src/ExcelPerfToolkit/JsonUtilities.cs:371-380 (line 373)
Category: data_corruption
Description: JsonToCell maps JSON numbers with
      element.TryGetDouble(out var d) ? d : (object)element.GetRawText()
  On .NET Core, JsonElement.TryGetDouble NEVER returns false for a syntactically valid JSON number:
  [verified] Probe C (.NET 8.0.422):
      1e400                          → TryGetDouble=True, value=Infinity
      9007199254740993  (2^53 + 1)   → TryGetDouble=True, value=9007199254740992   (silently rounded)
      18446744073709551615 (UInt64.Max) → True, 1.8446744073709552E+19             (silently rounded)
  So (a) the GetRawText() fallback is dead code, (b) int64-valued IDs above 2^53 — snowflake IDs,
  database bigints, ticks — are silently corrupted in their low digits, and (c) out-of-range numbers
  surface as double.Infinity, which Excel cannot represent (Excel-DNA marshals non-finite doubles to
  #NUM!), instead of the intended raw-text passthrough. Affects EPT.JSONPATH, EPT.PARSEJSON,
  EPT.READJSON, EPT.READNDJSON (all funnel through JsonToCell).
Trigger condition: =EPT.JSONPATH(A1, "id") where A1 = {"id": 1234567890123456789}
  → returns 1234567890123456770 (a different number), no error, no warning.
Trace:
  1. TryNavigate resolves the "id" element, ValueKind = Number, raw text "1234567890123456789".
  2. JsonToCell:373 → TryGetDouble → true, d = 1.2345678901234568E18.
  3. Cell shows 1234567890123456770. A user joining on that ID downstream gets wrong matches.
Fix (JsonUtilities.cs:373), before:
        JsonValueKind.Number => element.TryGetDouble(out var d) ? d : (object)element.GetRawText(),
  after:
        JsonValueKind.Number => NumberToCell(element),
  with:
        private const long MaxExactDoubleInteger = 9007199254740992; // 2^53
        private static object NumberToCell(JsonElement element)
        {
            if (element.TryGetInt64(out var l))
            {
                return l >= -MaxExactDoubleInteger && l <= MaxExactDoubleInteger
                    ? (double)l
                    : (object)element.GetRawText();   // preserve digits as text
            }
            return element.TryGetDouble(out var d) && double.IsFinite(d)
                ? d
                : (object)element.GetRawText();
        }
```

### BUG-MEDIUM-6
```
File: /home/user/plugin/src/ExcelPerfToolkit/JsonUtilities.cs:90-93
Category: logic_error
Description: ParseJson is the only UDF in these four files that declares native optional parameters:
      string? path = null, bool hasHeaderRow = true
  The project references plain ExcelDna.AddIn/Integration 1.8.0 (ExcelPerfToolkit.csproj) with no
  ExcelDna.Registration parameter-conversion pipeline, so C# default parameter values are ignored at
  registration: an omitted argument arrives as the marshaled default of the native type — FALSE for
  bool. Therefore =EPT.PARSEJSON(A1) over an array of objects emits NO header row, directly
  contradicting the attribute text ("TRUE (default) emits a header row", line 93), the XML doc
  (line 82-84), and the readme. Every sibling UDF (ReadJsonUdf:285, ReadNdjsonUdf:295,
  ReadFolderUdf:273) correctly takes `object` + ResolveBool(value, true) precisely to get
  default-TRUE semantics; ParseJson deviates. (`path = null` happens to be harmless: an omitted
  string marshals as empty, and TryNavigate:526 treats "" as root.)
Trigger condition: =EPT.PARSEJSON("[{""a"":1},{""a"":2}]") — expected 3 rows (header "a", 1, 2);
  actual 2 rows (1, 2). Column identity is silently lost.
Trace:
  1. Excel calls PARSEJSON with 1 argument; the registered native bool param receives FALSE.
  2. ParseJson:106 → ExpandToTable(node, hasHeaderRow: false) → ExpandArray:469 →
     BuildObjectTable(order, rows, emitHeader: false) → no header row.
Fix (JsonUtilities.cs:90-93), before:
        public static object[,] ParseJson(
            [ExcelArgument(...)] string json,
            [ExcelArgument(...)] string? path = null,
            [ExcelArgument(...)] bool hasHeaderRow = true)
  after:
        public static object[,] ParseJson(
            [ExcelArgument(...)] string json,
            [ExcelArgument(...)] object path,
            [ExcelArgument(...)] object hasHeaderRow)
        {
            ...
            if (!TryNavigate(doc.RootElement, PathOf(path), out var node)) ...
            return ExpandToTable(node, ResolveBool(hasHeaderRow, true));
```

### BUG-MEDIUM-7
```
File: /home/user/plugin/src/ExcelPerfToolkit/CacheUtilities.cs:255-263 (write), 188-208 + 291-298 (read)
Category: data_corruption / race_condition
Description: WriteDiskAsync opens the cache file with FileMode.Create — the previous good entry is
  truncated to zero at open, then rewritten in place with periodic flushes (:281-284). Two failure
  modes follow:
  (a) Durability: a cancellation (DiskWrite:157 passes ToolkitLifetime.ShutdownToken — i.e. every
      add-in unload during a write), process kill, or power loss mid-write leaves a zero-length or
      truncated-JSON file. The PREVIOUS value is already destroyed. On the next session,
      ReadDiskAsync:298 → JsonDocument.Parse(truncated bytes) → JsonException → DiskRead:204 catches
      Exception → returns #VALUE! — not the ifMissing fallback. The poisoned entry returns #VALUE!
      forever until the user manually runs EPT.DISKCACHE.CLEAR; the cache cannot heal itself.
  (b) Cross-instance torn access: the cache directory is shared by every Excel instance of the user
      (Path.GetTempPath()\ExcelPerfToolkit\diskcache, :384-385). Instance A writes with
      FileShare.Read (holds Write access); instance B's File.ReadAllBytesAsync requests Read with
      FileShare.Read — B's share mode does not admit A's Write access, so B's open fails with a
      sharing violation IOException (Windows sharing semantics:
      learn.microsoft.com/windows/win32/fileio/creating-and-opening-files). DiskRead's generic catch
      maps that transient condition to #VALUE! instead of treating it as a miss/retry.
Trigger condition: (a) Close Excel while =EPT.DISKCACHE.WRITE("k", bigBlock) is mid-flight (a 100k-row
  block takes hundreds of ms); reopen; =EPT.DISKCACHE.READ("k", fallback) → #VALUE!, fallback ignored.
  (b) Two open workbooks in separate Excel instances, one writing key K while the other reads K.
Trace (a):
  1. DiskWrite:157 → WriteDiskAsync; FileStream(FileMode.Create) truncates <hash>.json at :257-263.
  2. Utf8JsonWriter flushes the first 64 KiB at :281-284; ShutdownToken fires; :274
     ThrowIfCancellationRequested throws; stream disposes — file is a 64 KiB JSON prefix.
  3. Next session: ReadDiskAsync:297 reads the prefix; :298 Parse throws JsonException
     ("expected end of data"); DiskRead:204 → #VALUE!. ifMissing (:198) is unreachable because the
     exception bypasses the null-return path.
Fix: write to a sibling temp file and atomically move into place; treat unparseable entries as a miss:
  WriteDiskAsync — before:
        await using var stream = new FileStream(path, FileMode.Create, ...);
  after:
        var tmp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        await using (var stream = new FileStream(tmp, FileMode.CreateNew, FileAccess.Write,
            FileShare.None, DirectFileIO.DefaultBufferSize, FileOptions.Asynchronous))
        { ... write + final flush ... }
        File.Move(tmp, path, overwrite: true);   // MoveFileEx(MOVEFILE_REPLACE_EXISTING): atomic
                                                 // replace on the same NTFS volume; readers see
                                                 // either the old or the new complete file.
  ReadDiskAsync — wrap the parse:
        try { using var doc = JsonDocument.Parse(bytes); ... }
        catch (JsonException) { return null; }   // corrupt entry == miss; optionally File.Delete(path)
  (Keep an IOException retry-or-miss for the cross-instance sharing case as well.)
```

### BUG-MEDIUM-8
```
File: /home/user/plugin/src/ExcelPerfToolkit/FileSystemUtilities.cs:167-231 (esp. 174-187, 204-229)
Category: data_corruption
Description: ConcatByHeader unconditionally treats b[0, c] of every block as a header name. But three
  block shapes produced by the per-extension readers have no header row:
  (1) A .json file containing a single OBJECT: JsonUtilities.ExpandObject (JsonUtilities.cs:472-494)
      returns an N×2 key/value table regardless of hasHeaderRow. Row 0 is the FIRST key/value PAIR —
      real data. ConcatByHeader promotes it to two column names and drops it from the data.
  (2) A .json file containing a scalar: 1×1 block; the value itself ("42") becomes a column header.
  (3) An empty .csv: DirectFileIO returns 1×1 [[ExcelEmpty]] (DirectFileIO.cs:195-198);
      ToStringSafe(ExcelEmpty) = "" so a phantom empty-named column joins the union (:180-184).
Trigger condition: =EPT.READFOLDER("C:\configs", "*.json") over files like
  {"name":"Alice","age":30} — extremely common config/export shape.
Trace (file = {"name":"Alice","age":30}):
  1. ReadOneAsync:161 → ReadJsonAsync → ExpandObject → block = [["name","Alice"],["age",30.0]].
  2. ConcatByHeader:180 reads row 0 → union gains columns "name" and "Alice" (a VALUE as a header).
  3. dataRows += 1 (:186); the only emitted data row is ["age", 30] (:218-228), placed under the
     columns "name" and "Alice".
  4. Result: header row [name, Alice]; data row [age, 30]. The pair name=Alice is gone; the column
     identity is garbage. No error anywhere.
Fix: make ReadOneAsync mark blocks as headered or not, and have ConcatByHeader only consume a header
  from blocks that actually have one — e.g. return (object[,] block, bool hasHeader) from
  ReadOneAsync: false for JSON object/scalar expansions and for 1×1 empty blocks; for header-less
  blocks, map columns positionally into reserved "col{i}" union entries (or skip empty blocks
  entirely: if block is 1×1 && IsBlankOrError(block[0,0]) → continue, which alone fixes case (3)):
        if (br == 1 && bc == 1 && Marshaling.IsBlankOrError(b[0, 0])) { continue; }  // empty file
  plus a HasHeader flag consulted at :178 and :214 for cases (1)/(2).
```

### BUG-MEDIUM-9
```
File: /home/user/plugin/src/ExcelPerfToolkit/WatchFeeds.cs:79-101 (esp. 81, 91)
Category: resource_leak / logic_error (silently dead watcher)
Description: WatchFeedRunner arms the watcher once and parks forever on
  Task.Delay(Timeout.Infinite, ct). The Error handler (:81) merely calls onChange() — it bumps the
  counter once and does nothing else. For RECOVERABLE errors (InternalBufferOverflowException after
  an event burst > the default 8 KB buffer) that is actually adequate: the watcher keeps running and
  a bump is exactly the right trigger semantics. But for FATAL errors the Win32 ReadDirectoryChangesW
  loop terminates and never restarts: watched directory deleted or renamed, network share dropped,
  handle invalidated (Win32Exception ERROR_ACCESS_DENIED / ERROR_NETNAME_DELETED surfaced via the
  Error event — learn.microsoft.com/dotnet/api/system.io.filesystemwatcher.error). After that single
  bump the feed is a zombie: RunAsync still parked at :91, Feed._producer not completed, LatestValue
  frozen — the cell silently never updates again even after the directory is recreated. The user has
  no signal that the trigger is dead; downstream "re-import on change" formulas silently stop
  refreshing — the worst failure mode for a recalculation trigger.
  Related: a directory that does not exist YET makes the FileSystemWatcher ctor throw
  ([verified] probe H: ArgumentException) → Feed.RunSafelyAsync (RtdServer.cs:485-489) sets #N/A and
  the producer COMPLETES; the cell stays #N/A forever even after the folder is created (a producer
  restart only happens on a new Subscribe).
Trigger condition: =EPT.WATCHFOLDER("C:\drop") in a workbook; an upstream job deletes and recreates
  C:\drop nightly. After the first delete the counter bumps once and then freezes permanently.
Trace:
  1. Directory deleted → FSW raises Deleted (maybe) then Error(Win32Exception ACCESS_DENIED).
  2. OnError:81 → onChange() → counter++ → one last RTD push.
  3. The native watch loop is dead; EnableRaisingEvents is still true but no callbacks ever fire.
  4. RunAsync remains awaiting Task.Delay(Infinite) at :91; nothing observes the watcher's health.
  5. Directory recreated; events occur; cell never changes. No #N/A, no trace entry.
Fix: convert RunAsync into a re-arm loop; complete a TCS from the Error handler:
        public static async Task RunAsync(string directory, string filter, bool includeSubdirectories,
            Action onChange, CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                TaskCompletionSource fatal = new(TaskCreationOptions.RunContinuationsAsynchronously);
                try
                {
                    using var watcher = new FileSystemWatcher(directory, filter) { ... };
                    void OnError(object? s, ErrorEventArgs e)
                    {
                        onChange();
                        if (e.GetException() is not InternalBufferOverflowException) { fatal.TrySetResult(); }
                    }
                    ... attach handlers ...; watcher.EnableRaisingEvents = true;
                    await Task.WhenAny(fatal.Task, Task.Delay(Timeout.Infinite, cancellationToken)).ConfigureAwait(false);
                    cancellationToken.ThrowIfCancellationRequested();
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception) { /* directory missing/access denied: fall through to retry */ }
                onChange(); // signal consumers across the gap
                await Task.Delay(2000, cancellationToken).ConfigureAwait(false); // re-arm backoff
            }
        }
  This also heals the not-yet-existing-directory case (ctor throw → retry every 2 s) and the
  recreated-directory case. Optionally raise InternalBufferSize to 64 KB for watchfolder feeds.
```

### BUG-LOW-10
```
File: /home/user/plugin/src/ExcelPerfToolkit/WatchFeeds.cs:29-32
Category: logic_error
Description: WatchFileFeed maps a bare or relative path to a watch on "." :
      var dir = Path.GetDirectoryName(path);          // "data.csv" → ""
      _directory = string.IsNullOrEmpty(dir) ? "." : dir;
  "." is the Excel PROCESS current directory — typically C:\Windows\System32 or wherever the last
  File-Open dialog landed, and it changes at runtime. The feed silently watches an unrelated, drifting
  directory; the counter either never moves or bumps on unrelated churn (System32 is chatty).
Trigger condition: =EPT.WATCHFILE("data.csv") — any user who assumes workbook-relative resolution.
Trace:
  1. Spec "watchfile:data.csv" → ctor: dir = "" → _directory = ".", _filter = "data.csv".
  2. WatchFeedRunner arms FileSystemWatcher(".") → resolves against Environment.CurrentDirectory.
  3. The intended file (next to the workbook) is never watched; no error is surfaced.
Fix: reject relative paths at the feed boundary instead of guessing:
        if (!Path.IsPathFullyQualified(path))
        {
            throw new ArgumentException($"WATCHFILE requires an absolute path; got '{path}'.");
        }
  (Feed.RunSafelyAsync turns the throw into #N/A in the cell — wait, the ctor runs in
  FeedManager.CreateFeed on the ConnectData path, so the throw surfaces as #VALUE! from
  RtdServer.ConnectData:122-126. Either way it is visible instead of silent.)
```

### BUG-LOW-11
```
File: /home/user/plugin/src/ExcelPerfToolkit/JsonUtilities.cs:162-168 + 355-359
Category: logic_error (diagnosability / granularity)
Description: ReadNdjsonAsync parses each line as its own JsonDocument (:168). One malformed line
  anywhere in the file throws JsonException, which RunFileRead:355 maps to a whole-call #VALUE!.
  Two problems: (a) granularity — a 1M-line NDJSON file with one bad line yields nothing rather than
  999,999 rows; (b) the logged JsonException reports positions relative to the LINE
  ("LineNumber: 0"), so the trace cannot tell the user where the file is broken.
Trigger condition: =EPT.READNDJSON("C:\logs\events.ndjson") where line 412,377 was torn by a log
  rotation.
Trace:
  1. Loop at :162 reads line 412,377; :168 Parse throws JsonException{LineNumber=0, BytePositionInLine=58}.
  2. Exception unwinds; RunFileRead logs "EPT.READNDJSON invalid JSON: ... LineNumber: 0 ..." — wrong/
     useless position — and returns #VALUE! discarding 412,376 parsed rows.
Fix: count lines and rethrow with file position (keeps strictness but fixes diagnosability), or
  skip-and-trace per line (matches READFOLDER's intended resilience). Minimal version:
        var lineNo = 0;
        while (await reader.ReadLineAsync(cancellationToken).ConfigureAwait(false) is { } line)
        {
            lineNo++;
            if (string.IsNullOrWhiteSpace(line)) { continue; }
            try { using var doc = JsonDocument.Parse(line); ... }
            catch (JsonException ex)
            {
                throw new JsonException($"Invalid JSON on line {lineNo} of '{path}': {ex.Message}", ex);
            }
            ...
        }
```

### BUG-LOW-12
```
File: /home/user/plugin/src/ExcelPerfToolkit/CacheUtilities.cs:374-376
Category: data_corruption
Description: ReadTypedCell rehydrates errors via an unvalidated enum cast:
      element.TryGetProperty("$err", out var ep) && ep.TryGetInt32(out var ev)
          ? (ExcelError)ev
  A corrupt or hand-edited cache file with {"$err": 999} produces a boxed ExcelError with an
  undefined value (C# enum casts do not range-check). That box is handed to Excel-DNA's marshaler as
  the cell value; the xltypeErr payload 999 is not one of Excel's defined error codes, so the cell
  renders an undefined error state rather than the defensive ExcelEmpty the rest of ReadTypedCell
  falls back to.
Trigger condition: =EPT.DISKCACHE.READ("k") after the entry file was damaged or written by a
  different/older serializer version.
Trace:
  1. ReadDiskAsync:328 → ReadTypedCell; ValueKind=Object; "$err"=999 parses as int.
  2. (ExcelError)999 boxes successfully; result[r,c] holds it; the block crosses into Excel.
Fix:
        JsonValueKind.Object => element.TryGetProperty("$err", out var ep)
                && ep.TryGetInt32(out var ev) && Enum.IsDefined(typeof(ExcelError), ev)
            ? (ExcelError)ev
            : (object)element.GetRawText(),
```

### BUG-LOW-13
```
File: /home/user/plugin/src/ExcelPerfToolkit/FileSystemUtilities.cs:213-217 + 218-229
Category: data_corruption
Description: Within one file, duplicate header names map to the SAME union column:
  map[c] = order.IndexOf(name) returns the first occurrence's index for both duplicates, and the
  inner write loop (:222-226) writes them sequentially to the same result cell — the right-most
  duplicate silently wins, the left one's data is dropped. CSVs with repeated column names ("Value",
  "Value") are common in instrument exports.
Trigger condition: =EPT.READFOLDER over a CSV whose header is "id,Value,Value" — every row loses its
  first Value.
Trace:
  1. Header pass (:178-185): seen.Add("Value") true once → union has one "Value" at index 1.
  2. map = [0, 1, 1]. Row write: result[outRow,1] = b[r,1]; then result[outRow,1] = b[r,2].
  3. Column 1 holds only b[r,2]; b[r,1] is unrecoverable. No warning.
Fix: disambiguate duplicates during the union pass per block, e.g. track a per-block
  Dictionary<string,int> occurrence counter and rename repeats "Value#2" before seen.Add — or at
  minimum trace a warning when seen.Add(name) is false for a duplicate within the same block.
```

### BUG-LOW-14
```
File: /home/user/plugin/src/ExcelPerfToolkit/CacheUtilities.cs:32-35, 384-393
Category: memory_leak (unbounded growth) / contract drift
Description: (a) The disk cache has NO eviction, TTL, or size budget: every distinct key ever written
  leaves a file in %TEMP%\ExcelPerfToolkit\diskcache forever. Content-addressed keys (the documented
  pattern: EPT.HASHBLOCK over inputs) change whenever inputs change, so churned workbooks accrete
  dead entries indefinitely — 100k-row blocks serialize to multi-MB JSON each. (b) The memory cache
  caps ENTRIES at 4096 (:33) but not bytes: 50 memoized 100k×10 blocks ≈ 50 × ~30 MB (8 MB object[,]
  refs + ~22 MB boxed doubles) ≈ 1.5 GB pinned for the session with zero pressure-based relief.
  (c) Placement in Path.GetTempPath() means Windows Storage Sense / "Disk Cleanup → Temporary files"
  deletes the cache at will, quietly weakening the documented "survives workbook reopen" contract
  (CacheUtilities.cs:15-17, readme) — reads degrade to a miss, which is safe, but the persistence
  promise is unreliable.
Trigger condition: A daily-refresh workbook writing content-addressed keys for a month → thousands of
  orphaned multi-MB files; a session memoizing several large import blocks → GBs of working set.
Trace: DiskPath:387-393 only ever creates; nothing enumerates by age; DiskClear is manual-only.
  Memoize:72-76 evicts strictly by count.
Fix: see ARCH-3 (byte budget + LRU for memory; startup/periodic age-based prune for disk; move to
  %LOCALAPPDATA% if persistence is a real contract). Minimal patch: in DiskWrite, after a successful
  write, prune files older than N days when the directory exceeds M entries.
```

### Probe transcript (BCL verification, .NET SDK 8.0.422)

```
A: Parse(bytes+BOM) THROWS JsonException: '0xEF' is an invalid start of a value
B: Parse(Stream+BOM) OK, a=1
C[0]: raw=1e400 TryGetDouble=True value=Infinity
C[1]: raw=9007199254740993 TryGetDouble=True value=9007199254740992
C[2]: raw=18446744073709551615 TryGetDouble=True value=1.8446744073709552E+19
C[3]: raw=1.23456789012345678901234567890 TryGetDouble=True value=1.2345678901234567
D: GetDouble(1e400) = Infinity
E: after-dispose THROWS ObjectDisposedException
H: FSW ctor THROWS ArgumentException (non-existent directory)
G2 (as nobody): GetFiles(SearchOption.AllDirectories) THROWS UnauthorizedAccessException
M2 (as nobody): GetFiles(EnumerationOptions{IgnoreInaccessible=true}) OK, 1 files
F: symlink loop, GetFiles(AllDirectories) returned 42 files (1 real file, duplicated 41x)
F2: with AttributesToSkip=ReparsePoint: 1 file in 0 ms
I: int.TryParse NumberStyles.Integer accepts "+1" and " 1 " (path-index lenience, harmless)
J: Utf8JsonWriter double round-trip exact for 0.30000000000000004 and double.MaxValue
K: File.Delete(missing) is a no-op
```

---

## Pass 2 — Memory optimization

### MEM-1
```
File: /home/user/plugin/src/ExcelPerfToolkit/JsonUtilities.cs:160,170-190 (ReadNdjsonAsync) and 455-468 (ExpandArray)
Current memory cost: one Dictionary<string,object> per data row. On x64, a 10-entry
  Dictionary ≈ 80 B (object) + 92 B (int[17] buckets) + 432 B (Entry[17] @ 24 B) ≈ ~600 B/row,
  excluding the boxed values themselves. For a 100k-row × 10-key NDJSON file: ~60 MB of transient
  dictionary garbage (plus 1M hash inserts), all dead the moment BuildObjectTable finishes.
Optimization: maintain the union as Dictionary<string,int> name→column (append-only, so earlier rows
  simply have shorter arrays), store each row as object[] indexed by column:
  before (per row):
        var dict = new Dictionary<string, object>(StringComparer.Ordinal);
        ... dict[prop.Name] = JsonToCell(prop.Value); ... rows.Add(dict);
  after:
        // shared, built once per call:
        var index = new Dictionary<string, int>(StringComparer.Ordinal);
        // per row:
        var row = new object[Math.Min(index.Count + 8, ...)]; // or List<object> grown by column index
        foreach (var prop in element.EnumerateObject())
        {
            if (!index.TryGetValue(prop.Name, out var col)) { col = index.Count; index.Add(prop.Name, col); order.Add(prop.Name); }
            if (col >= row.Length) Array.Resize(ref row, index.Count);
            row[col] = JsonToCell(prop.Value);
        }
        rows.Add(row);
  and in BuildObjectTable read result[r+off, c] = c < rowArr.Length && rowArr[c] is not null ? rowArr[c] : ExcelEmpty.Value.
Expected reduction: ~600 B/row → ~104 B/row (object[10] = 24 B header + 80 B slots): ~6× less
  transient allocation (~60 MB → ~10 MB for 100k×10), removes 1M dictionary inserts on build and
  1M lookups on read-out (see PERF-5). Same change applies verbatim to ExpandArray (EPT.READJSON /
  EPT.PARSEJSON over arrays of objects).
```

### MEM-2
```
File: /home/user/plugin/src/ExcelPerfToolkit/JsonUtilities.cs:128 and /home/user/plugin/src/ExcelPerfToolkit/CacheUtilities.cs:297
Current memory cost: File.ReadAllBytesAsync allocates one exact-size byte[] of the whole file. A
  10 MB JSON file → a 10 MB Large Object Heap allocation per call (LOH threshold 85 KB), pure garbage
  after parse. Recalc-driven re-reads (e.g. EPT.READJSON keyed off EPT.WATCHFILE) produce repeated
  LOH garbage → gen2/LOH fragmentation and full blocking GCs inside the calc thread.
Optimization: JsonDocument.ParseAsync(Stream) — the stream path rents its working buffer from
  ArrayPool<byte>.Shared (pooled up to 2^30 elements), so steady-state per-call allocation for the
  raw bytes drops to ~0. Code change shown in BUG-HIGH-3 (same fix resolves the BOM bug); apply the
  identical transformation to CacheUtilities.ReadDiskAsync:297-298.
Expected reduction: 10 MB LOH garbage per call → ~0 pooled (file-byte buffer); for a 100k-row disk
  cache entry (~5-8 MB JSON) the same per-read saving. JsonDocument metadata remains pooled in both
  variants (no regression).
```

### MEM-3
```
File: /home/user/plugin/src/ExcelPerfToolkit/JsonUtilities.cs:523-579 (TryNavigate: 531-534, 548, 569)
Current memory cost: the path string is re-tokenized for every cell of EPT.JSONPATH:
  - :533 p.Substring(1) when the path starts with '$' — one string per cell;
  - :569 p.Substring(start, i - start) — one string per object step per cell.
  For 100k cells with path "data.items[0].name": 100k × (3 key substrings + 1 '$' strip) ≈ 100k ×
  ~100 B ≈ 10 MB of gen0 churn, plus the re-tokenization CPU (PERF-1).
Optimization: parse the path ONCE per UDF call into a step array before the r/c loops in
  JsonPath:54, then navigate with the precompiled steps:
        readonly struct PathStep { public readonly string? Key; public readonly int Index; ... }
        private static PathStep[] CompilePath(string? path) { ... existing tokenizer, run once ... }
        private static bool TryNavigateSteps(JsonElement root, PathStep[] steps, out JsonElement result) { ... }
  TryGetProperty(string) accepts the cached key strings; zero per-cell path allocations remain.
Expected reduction: per-cell path allocations 100 B → 0 (≈10 MB per 100k-cell call); CPU effect in PERF-1.
```

### MEM-4
```
File: /home/user/plugin/src/ExcelPerfToolkit/FileSystemUtilities.cs:145-152 + 167-231
Current memory cost: every per-file object[,] is retained in `blocks` until ConcatByHeader finishes
  building the final array. For 1,000 files × 1,000 rows × 10 cols, the final block holds 10M
  references (80 MB) and the per-file blocks hold another 10M (80 MB) — peak ≈ 2× the result's
  reference storage (the boxed cell values themselves are shared by reference, so they are not doubled).
Optimization: the union/header pass already iterates blocks once (:172-188); after computing `map`
  for a block and copying its rows (:205-230), null out the list slot (blocks[i] = null!) so each
  source block becomes collectible as soon as it is consumed, instead of all of them surviving until
  the method returns:
        for (var i = 0; i < blocks.Count; i++) { var b = blocks[i]; ...copy...; blocks[i] = null!; }
Expected reduction: peak extra retention drops from ~80 MB (all shells live at once) toward the
  single largest block (~80 KB–8 MB), i.e. near-1× peak instead of 2× for the reference arrays.
  Honest caveat: the dominant memory (boxed values + final array) is inherent to the API shape;
  this trims the avoidable half only.
```

---

## Pass 3 — CPU / throughput

### PERF-1
```
File: /home/user/plugin/src/ExcelPerfToolkit/JsonUtilities.cs:54-76 + 523-579
Hot path: YES — per cell of every EPT.JSONPATH array call.
Current cost: per cell = JsonDocument.Parse (~1-2 µs per KB of JSON; this is the irreducible work)
  + path re-tokenization (~150-400 ns + 3-4 allocations for a 3-step path) + char-walk branching.
  For 100k cells × 1 KB JSON: parse ≈ 150 ms, path re-parse ≈ 25-40 ms + GC pressure from MEM-3.
Optimization: CompilePath once per call (code shape in MEM-3); per-cell work becomes Parse +
  TryGetProperty probes only.
Expected speedup: navigation phase 3-5× (tokenize+alloc eliminated, property probe remains); whole
  call ≈ 1.1-1.2× (document parse dominates and is untouched). Also removes ~10 MB/100k-cell churn.
Before → after: see MEM-3.
```

### PERF-2
```
File: /home/user/plugin/src/ExcelPerfToolkit/FileSystemUtilities.cs:145-152
Hot path: YES — entire EPT.READFOLDER wall clock.
Current cost: strictly serial `await ReadOneAsync` per file. 1,000 × 64 KB CSVs on SSD ≈ 1,000 ×
  (open 0.1-0.3 ms + read 0.2 ms + parse 0.5-2 ms) ≈ 1-2.5 s, single core. On an SMB share each open
  adds 1-5 ms RTT: 3-10 s dominated by latency, not bandwidth.
Optimization: bounded parallel reads preserving order:
  before:
        foreach (var file in files) { blocks.Add(await ReadOneAsync(file, ...)); }
  after:
        var results = new object[files.Length][,];
        using var gate = new SemaphoreSlim(Math.Min(8, Environment.ProcessorCount));
        var tasks = files.Select(async (file, i) =>
        {
            await gate.WaitAsync(cancellationToken).ConfigureAwait(false);
            try { results[i] = await ReadOneAsync(file, hasHeaderRow, cancellationToken).ConfigureAwait(false); }
            finally { gate.Release(); }
        });
        await Task.WhenAll(tasks).ConfigureAwait(false);
        var blocks = results.ToList();   // original sorted order preserved by index
Expected speedup: 4-6× on SSD (parse parallelizes across cores), up to ~8× on latency-bound network
  shares with gate=8. Determinism preserved (results indexed by the pre-sorted files array). The UDF
  is IsThreadSafe=false so the calc thread blocks either way — this shrinks how long.
```

### PERF-3
```
File: /home/user/plugin/src/ExcelPerfToolkit/FileSystemUtilities.cs:213-217
Hot path: yes within READFOLDER concat for wide unions.
Current cost: map[c] = order.IndexOf(name) — List<string>.IndexOf is a linear ordinal scan. For
  1,000 files × 20 cols against a 200-column union: 1,000 × 20 × avg 100 string compares ≈ 2M
  comparisons ≈ 30-50 ms. Degenerates quadratically as the union grows.
Optimization: build Dictionary<string,int> headerIndex alongside `order` in the union pass (:178-185),
  then map[c] = headerIndex.TryGetValue(name, out var idx) ? idx : -1.
Expected speedup: mapping phase ~40× (O(1) probes); whole-call effect <1% for typical folders (I/O
  dominates) but removes the quadratic cliff for 1,000+-column unions. Cost: one small dictionary.
```

### PERF-4
```
File: /home/user/plugin/src/ExcelPerfToolkit/FileSystemUtilities.cs:81-103
Hot path: yes for EPT.FILEINFO over large path blocks, especially on UNC paths.
Current cost: 3 metadata syscalls per existing file path — Directory.Exists (:83), File.Exists (:84),
  FileInfo first-property access (:90) — each a full stat / network round trip. Local NTFS ≈ 2-10 µs
  each (10k paths: 60-300 ms); SMB ≈ 0.5-2 ms each (10k paths: 15-60 s).
Optimization: stat once via FileInfo and branch on it; only fall back to DirectoryInfo when not a file:
  before:
        var isDir = Directory.Exists(path);
        var isFile = File.Exists(path);
        ... new FileInfo(path) ...
  after:
        var fi = new FileInfo(path);
        if (fi.Exists)                                  // 1 stat, fields cached for Length/times/Name
        { ...file row from fi...; result[row,1] = true; }
        else
        {
            var di = new DirectoryInfo(path);           // 2nd stat only for dirs/missing
            if (di.Exists) { ...dir row...; } else { ...missing row...; }
        }
Expected speedup: 3 stats → 1 for files (the common case), 2 for dirs/missing: ~3× on network shares,
  ~2-3× locally for the stat phase. Identical output table.
```

### PERF-5
```
File: /home/user/plugin/src/ExcelPerfToolkit/JsonUtilities.cs:510-519 (BuildObjectTable)
Hot path: yes for large READNDJSON / READJSON results.
Current cost: result fill does dict.TryGetValue(order[c]) per output cell — rows × cols string-hash
  lookups. 100k rows × 10 cols = 1M hashes+probes ≈ 30-60 ms, plus cache misses across 100k
  dictionaries.
Optimization: falls out of MEM-1 — with column-indexed object[] rows the fill becomes a bounds check
  + array read per cell (~1-2 ns):
        result[rowOffset + i, c] = c < rowArr.Length && rowArr[c] is not null ? rowArr[c] : ExcelEmpty.Value;
Expected speedup: table-build phase ~10-20×; whole READNDJSON call ~1.2-1.5× (parse dominates).
```

### PERF-6
```
File: /home/user/plugin/src/ExcelPerfToolkit/CacheUtilities.cs:118-134, 387-393
Hot path: no (cache management), included for completeness.
Current cost: (a) EvictOldest enumerates the full ConcurrentDictionary (acquiring its bucket locks)
  on EVERY Memoize once Count > 4096: ~4096-entry scan ≈ 50-100 µs per call → a 10k-Memoize recalc
  burns 0.5-1 s purely on eviction scans. (b) DiskPath calls Directory.CreateDirectory on every
  read AND write (:390) — an extra metadata syscall per DISKCACHE.READ even on pure hits/misses.
Optimization: (a) evict in batches — when over cap, take the 64 oldest in one scan
  (one O(N) pass amortized over 64 inserts → ~64× fewer scans); (b) ensure the directory once:
        private static readonly Lazy<string> EnsuredDir = new(() =>
        { var d = CacheDirectory(); Directory.CreateDirectory(d); return d; });
Expected speedup: eviction overhead ÷64 at cap; one syscall removed per disk-cache op (≈10-50 µs
  local, ~1 ms UNC).
```

---

## Pass 4 — Architectural wins

### ARCH-1 — Resilient, observable folder ingest
```
Scope: FileSystemUtilities.ReadFolderAsync / ReadOneAsync / ConcatByHeader (subsumes BUG-HIGH-1,
  BUG-HIGH-2, BUG-MEDIUM-4, BUG-MEDIUM-8, BUG-LOW-13; incorporates PERF-2/PERF-3).
Current pattern: enumerate (legacy options) → serial read, fail-fast on any file → header union that
  assumes every block is headered tabular data → silent drops on mismatch.
Proposed pattern: enumerate with EnumerationOptions{IgnoreInaccessible, AttributesToSkip=ReparsePoint};
  per-file try/catch producing a (block, hasHeader, error) record; bounded-parallel reads (gate 8);
  optional trailing "__source"/"__status" columns (opt-in flag) so users can see which file each row
  came from and which files were skipped; positional "col{i}" mapping for header-less blocks.
Impact estimate: converts the flagship ingest UDF from "any one of 1,000 files can zero the result"
  to per-file degradation + 4-8× faster wall clock; eliminates the two HIGH bugs at their root.
Effort: ~1-1.5 days incl. tests. Risk: low — output shape unchanged unless the opt-in columns are
  requested; skip behavior is strictly additive relative to today's total failure.
```

### ARCH-2 — Shared atomic-write protocol
```
Scope: CacheUtilities.WriteDiskAsync (BUG-MEDIUM-7), JsonUtilities.WriteJsonAsync, and (seam)
  DirectFileIO.WriteDelimitedAsync.
Current pattern: FileMode.Create truncates the destination in place; any cancel/crash/power event
  mid-write destroys the previous content and leaves a torn file; concurrent readers in other Excel
  instances race the writer's handle.
Proposed pattern: one helper — write to "{path}.{guid}.tmp" in the same directory (same volume →
  atomic rename), final flush, then File.Move(tmp, path, overwrite: true) (MoveFileEx
  MOVEFILE_REPLACE_EXISTING; readers observe either the complete old or complete new file), with a
  try/finally File.Delete(tmp) on failure. Disk-cache read side treats parse failures as a miss.
Impact estimate: closes the only data-destroying failure mode in the four files; cache becomes
  crash-safe and multi-instance-safe; WRITEJSON/WRITECSV stop being able to destroy a good output
  file on a cancelled save.
Effort: ~0.5 day. Risk: low; rename adds one metadata op per write (~0.1 ms) — negligible against
  the write itself. Caveat to document: File.Move(overwrite) across volumes is copy+delete, but the
  tmp file is always created beside the destination so this cannot occur.
```

### ARCH-3 — Cache governance (size, age, health)
```
Scope: CacheUtilities (BUG-LOW-14, PERF-6).
Current pattern: memory cache = count-capped (4096) ConcurrentDictionary with O(N) single-evict, no
  byte accounting; disk cache = append-only directory in %TEMP% with no TTL, no size budget, no
  corruption recovery, manual CLEAR only.
Proposed pattern: memory — track approximate bytes per entry (rows × cols × 32 B heuristic), evict
  LRU batches past a configurable byte budget (default e.g. 256 MB); disk — a tiny manifest-free
  policy: on first write per session, prune entries with LastWriteTimeUtc older than N days (default
  30) and trim to a byte budget by oldest-first; treat unparseable entries as a miss and delete them
  (heals BUG-MEDIUM-7 leftovers); optionally relocate to %LOCALAPPDATA%\ExcelPerfToolkit\diskcache so
  Windows temp cleanup stops silently voiding the documented persistence.
Impact estimate: bounds worst-case working set from "GBs, silent" to a configured budget; disk cache
  stops accreting dead content-addressed entries forever; failure modes become self-healing.
Effort: ~1 day. Risk: low-medium — eviction can drop entries users expected to live; document the
  budgets and keep CLEAR semantics unchanged.
```

### ARCH-4 — Self-healing watch feeds
```
Scope: WatchFeedRunner (BUG-MEDIUM-9, BUG-LOW-10).
Current pattern: arm once, park forever; Error → counter bump only; missing directory → permanent
  #N/A; deleted directory → permanently frozen counter.
Proposed pattern: re-arm loop with backoff (code shape in BUG-MEDIUM-9): fatal Error or
  arm-failure → dispose, bump (so dependents recalc across the gap), retry every ~2 s; surfaces
  ExcelErrorNA via LatestValue while unarmed and resumes counting when the directory reappears.
  Raise InternalBufferSize to 64 KB (the documented max worth using) for watchfolder feeds to shrink
  the overflow window on bursty directories.
Impact estimate: the recalculation-trigger contract ("this cell moves when the target changes")
  becomes true across directory recreation, share drops, and pre-creation subscription — the three
  realistic lifecycles of a watched drop-folder.
Effort: ~0.5 day. Risk: low — the Feed base contract (RunAsync observes the token, publishes via
  LatestValue) is unchanged; retry loop is contained in WatchFeedRunner.
```

---

## Seam notes (out-of-scope observations, one line each)

- `DirectFileIO.cs:117-121` — `reader.Peek()`/`reader.Read()` are synchronous calls on a `FileOptions.Asynchronous` stream inside the async read loop (chunk-boundary quote handling): a sync-over-async micro-block on the worker thread.
- `DirectFileIO.cs:158-163` — a blank line mid-file becomes a one-cell `[""]` row (PushField runs unconditionally on `\r`/`\n`), which then feeds READFOLDER's concat as a data row.
- `DirectFileIO.WriteDelimitedAsync` (`DirectFileIO.cs:266-272`) — same truncate-in-place write as BUG-MEDIUM-7; if ARCH-2 lands, apply the helper there too (accepted round-2 surface, so noted only).
- `RtdServer.cs:250` — `FeedManager._feeds` is `OrdinalIgnoreCase`, so `watchfile:` specs dedupe case-insensitively: correct for Windows paths, would alias distinct paths on a case-sensitive filesystem (academic for `net8.0-windows`).

---

## Rejection appendix

| # | Candidate finding | Why rejected |
|---|---|---|
| R1 | JsonElement use-after-Dispose (the classic JsonDocument lifetime bug) | Audited every `using var doc` scope: `JsonPath:66-69`, `ParseJson:101-106`, `ReadJsonAsync:129-134`, `ReadNdjsonAsync:168-189`, `CacheUtilities.ReadDiskAsync:298-333`. Every JsonElement is fully materialized to object (JsonToCell/ExpandToTable/ReadTypedCell) before the document disposes; no element escapes a scope. Verified the failure mode is real if it occurred (probe E: ObjectDisposedException) — it does not occur here. |
| R2 | WatchFeeds `Bump` race: two events can publish `LatestValue` out of order (count regresses 3→2) | `Interlocked.Increment` makes counts unique and fresh; an out-of-order `Volatile.Write` can publish a lower count transiently, but no newly produced value can equal the previously pushed `LastPushed` (all batch values are > it), so the flush timer (RtdServer.cs:164-169) always detects ≥1 change per batch. Trigger semantics (the only contract) cannot lose a recalc. Benign. |
| R3 | FILEINFO TOCTOU: file deleted between `File.Exists` (:84) and `fi.Length` (:90) | `FileInfo.Length` then throws FileNotFoundException, an IOException subclass caught by the filter at FileSystemUtilities.cs:110; the row degrades to Exists=true with blank metadata and the call continues. Explicitly designed-for; no corruption. |
| R4 | Session-cache races (prompt: "IsThreadSafe=false UDFs still race with anything else touching the static") | All five touchpoints of `Memory` are IsThreadSafe=false UDFs serialized on Excel's calc thread; no RTD/background code references `CacheUtilities`. Residual VBA `Application.Run` reentrancy is covered anyway: ConcurrentDictionary + Interlocked `_stampSequence`; the worst concurrent-EvictOldest outcome is evicting two entries. No invariant breaks. |
| R5 | Disk-cache key→filename injection (path traversal via `..`, invalid chars, post-sanitization collision) | `DiskPath` (CacheUtilities.cs:391) hex-encodes SHA-256 of the key: output alphabet is [0-9A-F], no traversal or invalid chars possible; collision probability ≈ 2⁻¹²⁸. Not exploitable. |
| R6 | `WriteJsonAsync` truncate-in-place destroys prior file on cancel | Real, but byte-for-byte the same pattern as the round-2-accepted `DirectFileIO.WriteDelimitedAsync` (FileMode.Create, same flush cadence). Not a NEW deviation; routed to ARCH-2 instead of a bug entry (the disk cache variant IS reported — BUG-MEDIUM-7 — because its read-back/poisoning semantics are new). |
| R7 | Sync-over-async deadlock in the six `.GetAwaiter().GetResult()` bridges | Audited all six (JsonUtilities:286/296/311-314/338, FileSystemUtilities:277-285, CacheUtilities:157-160/190-193): every async core is `ConfigureAwait(false)` throughout, no SynchronizationContext capture, all wired to `ToolkitLifetime.ShutdownToken`, no fire-and-forget tasks. Matches the accepted round-2 pattern exactly; no new deviation. |
| R8 | `ReadDiskAsync` dimension attack (huge rows×cols → OOM) | `(long)rows * cols > int.MaxValue` guarded at CacheUtilities.cs:302; short/jagged `cells` arrays are bounds-checked (:317-331) and padded with ExcelEmpty (:307-313). A merely-large allocation OOMs into the generic catch → #VALUE!. Adequate. |
| R9 | `TryGetArrayElement` linear scan vs `JsonElement` indexer | The built-in `element[index]` also walks the metadata rows O(index); no asymptotic or measured win available. Rejected as PERF. |
| R10 | NotifyFilter missing `Attributes`/`Security` | The documented contract (WatchFeeds.cs:17-19, readme) is create/change/delete/rename — fully covered by LastWrite|FileName|DirectoryName|Size|CreationTime. Attribute-only touches not triggering is consistent with the docs. |
| R11 | 8 KB `InternalBufferSize` overflow loses events | For a change COUNTER, the Error handler's `onChange()` (WatchFeeds.cs:81) converts buffer overflow into exactly the promised signal (a bump); individual lost events are irrelevant to the semantics. Buffer-size increase kept only as optional hardening under ARCH-4. The FATAL error case is reported separately (BUG-MEDIUM-9). |
| R12 | `FeedManager.CreateFeed` splitting `watchfile:C:\x.csv` on the drive colon | RtdServer.cs:321-325 splits at the FIRST colon (index 9, after "watchfile"); `args` = "C:\x.csv" intact. Verified by inspection. |
| R13 | DiskClear TOCTOU: file deleted between `File.Exists` (:237) and `File.Delete` (:239) | File.Delete on a missing path is a no-op ([verified] probe K); worst case the count is off by one. Locked-file failure mid-bulk-clear returns #VALUE! after partial deletion — accepted as a manual-maintenance path. |
| R14 | `int.TryParse(NumberStyles.Integer)` in TryNavigate accepts "[+1]"/"[ 1 ]" | [verified] probe I: both parse. Pure lenience; resolves to the same element; cannot mis-resolve. |
| R15 | `ParseJson("")` returns blank instead of an error | Consistent with the documented JSONPATH behavior ("a blank cell stays blank", JsonUtilities.cs:42) and harmless. |
| R16 | `WriteCellJson` serializes ExcelError as its display string ("#N/A") — lossy round-trip via WRITEJSON→READJSON | Intentional interchange design: JSON has no error type, and the comment-free but deliberate split exists — the CACHE path preserves errors typed via `$err` (CacheUtilities.cs:336-344) where fidelity matters; the JSON-export path favors human-readable output. |
| R17 | Search-pattern injection (`pattern` = "..\\*") | .NET throws ArgumentException for path-walking search patterns; caught at ReadFolderUdf:297 → #VALUE!. The folder argument itself is user-trusted by design (local add-in, not a privilege boundary). |
| R18 | EPT.JSONPATH `IsThreadSafe = true` with static state | The method touches only locals and a readonly TraceSource; JsonDocument instances are per-call; Marshaling helpers are pure. MTR-safe as registered. |
| R19 | `ResolveBool` returns FALSE (not the default) for unrecognized strings like "yes" (JsonUtilities.cs:617, FileSystemUtilities.cs:357) | Micro-inconsistency with the `_ => defaultValue` arm, but the same helper is copied identically in DirectFileIO (accepted round 2); fixing one copy would diverge behavior across the family. Belongs to a cross-cutting helper consolidation, not this domain's bug list. |
| R20 | `ConcatByHeader` pre-fills the whole result with ExcelEmpty then overwrites mapped cells (double store) | The pre-fill is required for unmapped gap cells; the extra reference write per cell is ~1 ns and removing it requires per-cell branch logic that costs as much. Not worth it. |
| R21 | FILEINFO `ToOADate()` on pre-1900 filesystem timestamps (e.g. 1601-01-01) | Produces a negative OA serial, which Excel renders as a number (not an error) and DateTime.ToOADate only throws for years 1-99 — not reachable from real filesystem timestamps (FILETIME epoch is 1601). The Marshaling.TryToDouble guard pattern confirms the BCL edge is years 1-99 only. |
| R22 | `Memoize` stores the caller's `object[,]` by reference (no defensive copy) — aliasing if a caller mutates | Excel-DNA materializes a fresh `object[,]` per UDF invocation and nothing in the toolkit mutates argument blocks (audited: all utilities allocate result arrays). CacheGet returning the shared instance to multiple cells is read-only marshaling. No observable aliasing path. |
| R23 | `WriteJsonAsync` duplicate header names → duplicate JSON property names | Legal JSON (RFC 8259 does not forbid), mirrors what the sheet actually shows, and consumers (including this toolkit's own readers, last-wins) handle it. Cosmetic at most. |
