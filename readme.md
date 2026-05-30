# ExcelPerfToolkit

A production-grade .NET 8 library packaged as an Excel-DNA XLL add-in. It is a
developer-utility toolkit that eliminates the **three highest-impact performance
bottlenecks** in Excel/VBA workflows.

## The three bottlenecks this toolkit fixes

1. **Cell-by-cell worksheet read and write.** Every `Range.Value` access through
   the COM object model crosses the managed/COM boundary. A naive VBA loop that
   touches *N* cells pays *N* round trips: catastrophic past a few thousand
   cells. The fix is bulk transfer: read an entire range into a managed array
   in one marshaled call, operate on it in pure managed memory, and write it
   back in one call.
2. **COM per-cell marshaling more broadly.** Same root cause, wider surface.
   Every public entry point in this toolkit is designed so that operating on
   *N* cells costs roughly *one* crossing of the managed-to-COM boundary,
   never *N*.
3. **VBA single-threaded recalc.** VBA cannot use Excel's multithreaded
   recalculation engine, so CPU-bound work serializes on a single core. The
   fix is to expose CPU-bound utilities as thread-safe UDFs registered with
   `IsThreadSafe = true`, letting Excel schedule them across cores during a
   normal recalc, plus an explicit `Parallel.For` batch path for heavy
   transforms.

## Second-wave bottlenecks (the second three)

After the first three fixes landed, three more surfaced on production workloads:

4. **Interpreted runtime with no JIT and no SIMD.** Excel's formula engine is
   interpreted, so arithmetic-heavy work pays per-cell dispatch costs and
   never sees a vector instruction. The fix is to move heavy kernels into
   compiled, vectorized .NET that targets `Vector<T>` (portable SIMD) and,
   where measurably worthwhile, `System.Runtime.Intrinsics` (AVX2/FMA) with a
   runtime-checked scalar fallback so the same XLL stays correct on any CPU.
5. **Real-time feeds throttled through single-threaded marshaling.** Each
   tick of a real-time feed pumped through `Application` is a separate COM
   crossing on Excel's UI thread; many cells subscribed to the same feed
   multiply that cost. The fix is a properly multithreaded RTD server backed
   by a concurrency-safe feed manager that multiplexes many topics over
   shared producers and throttles outbound updates to a configurable
   interval.
6. **File I/O performed by opening workbooks through the object model.**
   Loading a CSV via `Workbooks.Open` spins up the entire object model, fires
   events, and serializes through Excel's UI thread. The fix is direct
   streaming I/O through managed `FileStream`s with async + cancellation,
   never touching the object model.

## Third-wave bottlenecks (conditional aggregation and lookups)

Two more bottlenecks surface specifically around conditional aggregation and
lookups - the work most real spreadsheets spend their recalc budget on:

7. **Conditional functions re-scan their ranges per formula.** A column of
   *R* native `SUMIFS`/`COUNTIFS`/`AVERAGEIFS` formulas, each scanning a
   criteria range of *M* cells, is O(*R*·*M*) and is recomputed on every
   edit. The fix is to take the value and criteria ranges as one bulk
   `object[,]` each (one crossing apiece), build the match mask in a single
   managed-memory pass, and register every function `IsThreadSafe = true` so
   Excel's MTR engine fans the column out across cores. The same engine also
   adds the conditional aggregates Excel **lacks natively** - `MEDIANIFS`,
   `PERCENTILEIFS`, `MODEIFS`, `STDEVIFS`/`VARIFS`, `GEOMEANIFS`/`HARMEANIFS`,
   `PRODUCTIFS`, `DISTINCTCOUNTIFS`, `FIRSTIFS`/`LASTIFS` - plus weighted
   statistics (`WAVG`, `WAVGIFS`, `WMEDIAN`, `WSTDEV`), a conditional
   `SUMPRODUCT`, and a one-pass `GROUP BY`.
8. **`VLOOKUP`/`XLOOKUP` re-scan the table per lookup.** A column of *R*
   lookups against a table of *M* rows is O(*R*·*M*). The fix is to build the
   table's index **once** - a hash map for exact match, a sorted array for
   approximate match - and probe it for every lookup value: O(*M* + *R*)
   total, table read in a single crossing. That is `EPT.XLOOKUPB`.

## Architecture

```
ExcelPerfToolkit (single net8.0-windows class library, packed as XLL)
├── AddIn.cs                 - IExcelAddIn entry point + lifetime wiring (boots #1..#6)
├── Marshaling.cs            - Allocation-conscious helpers (#2)
├── BulkTransfer.cs          - Read/write/round-trip a whole range (#1, #2)
├── DeveloperUtilities.cs    - The day-to-day utility surface  (#1, #2)
├── ParallelUtilities.cs     - Thread-safe UDFs + Parallel.For (#3)
├── VectorizedKernels.cs     - SIMD kernels with scalar fallback (#4)
├── RtdServer.cs             - Multithreaded RTD server + feed manager (#5)
├── DirectFileIO.cs          - Async CSV/TSV streaming, no object model (#6)
├── ConditionalAggregates.cs - *IFS family, weighted stats, SUMPRODUCTIFS, GROUPBY (#7)
├── LookupBoost.cs           - Batched exact/approximate lookup, O(M+R) (#8)
├── TextUtilities.cs         - Case/pad/repeat/reverse/templatefill over blocks (#1, #2)
├── RegexUtilities.cs        - Match/count/extract/extractall/split over blocks (#1, #2)
├── SeriesUtilities.cs       - Fill-forward, outliers, quantiles (#1, #2)
├── DateUtilities.cs         - WORKDAYADD (vectorized WORKDAY.INTL) (#1, #2)
├── DistanceUtilities.cs     - Pairwise distance matrices via SIMD dot products (#4)
├── JsonUtilities.cs         - JSONPATH/PARSEJSON + async READ/WRITE JSON & NDJSON (#1, #2, #6)
└── ToolkitLifetime.cs       - Shared shutdown CancellationTokenSource + TraceSource factory
```

Everything is built on three primitives:

- `BulkTransfer.ReadBlock(...)` - one COM crossing pulls a whole `object[,]`.
- `BulkTransfer.WriteBlock(...)` - one COM crossing pushes a whole `object[,]`.
- `BulkTransfer.RoundTripTransform(...)` - read, transform in pure managed
  memory, write back. Two crossings total, regardless of cell count.

`ExcelDnaUtil.Application` (late-bound COM) is **not** used on any hot path.
We rely exclusively on `ExcelReference` + `XlCall` and Excel-DNA's native
`object[,]` marshaling, which is dramatically faster than
`Microsoft.Office.Interop.Excel`.

## Build

Prerequisites: **.NET 8 SDK** and Visual Studio Build Tools or the `dotnet`
CLI on Windows.

```bash
# Restore once.
dotnet restore ExcelPerfToolkit.sln

# x64 (default; the recommended bitness for modern Excel).
dotnet build src/ExcelPerfToolkit/ExcelPerfToolkit.csproj -c Release -p:Platform=x64

# x86 (only if you specifically have 32-bit Office).
dotnet build src/ExcelPerfToolkit/ExcelPerfToolkit.csproj -c Release -p:Platform=x86
```

The packed XLLs are produced under:

```
src/ExcelPerfToolkit/bin/x64/Release/net8.0-windows/publish/
  ExcelPerfToolkit-AddIn64-packed.xll   # load this from 64-bit Excel
src/ExcelPerfToolkit/bin/x86/Release/net8.0-windows/publish/
  ExcelPerfToolkit-AddIn-packed.xll     # load this from 32-bit Excel
```

To install: in Excel, **File -> Options -> Add-Ins -> Manage: Excel Add-ins ->
Go -> Browse**, point at the packed XLL, and tick it.

## Function catalogue

All functions are registered under category prefixes (`EPT.BulkTransfer`,
`EPT.DeveloperUtilities`, `EPT.Parallel`). Every signature is documented with
its marshaling cost (number of boundary crossings) and its thread-safety with
respect to Excel's multithreaded recalc.

### Bulk transfer (`BulkTransfer.cs`)

| Function | Signature | Description |
| --- | --- | --- |
| `EPT.READRANGE` | `(sheet, a1) -> object[,]` | Read an A1 range as a single bulk array. **1 boundary crossing.** |
| `BulkTransfer.WriteBlock` | `(ref, object[,])` | Push an entire array back in one crossing. |
| `BulkTransfer.RoundTripTransform` | `(ref, Func<object[,], object[,]>)` | Read - transform in managed - write. Two crossings total. |
| `BulkTransfer.ReadDoubleBlock` / `WriteDoubleBlock` | typed convenience | Returns `double[,]`; coerces non-numerics. |
| `BulkTransfer.ReadStringBlock` | typed convenience | Returns `string[,]`. |

### Developer utilities (`DeveloperUtilities.cs`)

| Function | Signature | Description |
| --- | --- | --- |
| `EPT.TRIMBLOCK` | `(object[,]) -> object[,]` | Trim + collapse whitespace in every string cell. |
| `EPT.COERCENUMERIC` | `(object[,]) -> object[,]` | Convert numeric-looking text to real numbers. |
| `EPT.REMOVEDUPLICATEROWS` | `(object[,]) -> object[,]` | Drop duplicate rows, keep first occurrence. |
| `EPT.TRANSPOSE` | `(object[,]) -> object[,]` | Transpose a block. |
| `EPT.FILLBLANKS` | `(object[,], object) -> object[,]` | Replace blanks with a supplied value. |
| `EPT.FILLDOWN` | `(object[,]) -> object[,]` | Fill blanks with value above, per column. |
| `EPT.SPLITCOLUMN` | `(object[,], string) -> object[,]` | Split single column into many by delimiter. |
| `EPT.JOINCOLUMNS` | `(object[,], string) -> object[,]` | Join all columns of each row with delimiter. |
| `EPT.FINDREPLACE` | `(object[,], string, string, bool) -> object[,]` | Find/replace with optional regex. |
| `EPT.STACKCOLUMNS` | `(object[,]) -> object[,]` | Stack every column into one. |
| `EPT.UNPIVOT` | `(object[,], int) -> object[,]` | Reshape wide to long (keys + attribute + value). |
| `EPT.UNIQUE` | `(object[,]) -> object[,]` | Distinct non-blank values, first-seen order. |
| `EPT.UNIQUECOUNT` | `(object[,]) -> object[,]` | Distinct values with counts. |
| `EPT.BLOCKLOOKUP` | `(object[,], int, object[,], object[,]) -> object[,]` | Hash-join a key column against a lookup table. |
| `EPT.SORTBLOCK` | `(object[,], object[,], object) -> object[,]` | Stable sort by one or more key columns. |
| `EPT.HASHBLOCK` | `(object[,]) -> string` | XXH3 hex digest over a block. |
| `EPT.SHA256BLOCK` | `(object[,]) -> string` | SHA-256 hex digest over a block. |

### Parallel & thread-safe (`ParallelUtilities.cs`)

| Function | Thread-safe for MTR? | Description |
| --- | --- | --- |
| `EPT.MT.ROWSUMS` | yes | Per-row sums over a numeric block. |
| `EPT.MT.ROWZSCORES` | yes | Per-row z-score normalization. |
| `EPT.MT.POLYEVAL` | yes | Element-wise polynomial via Horner's method. |
| `EPT.MT.DOT` | yes | Dot product of two equal-length flattened vectors. |
| `EPT.PARALLELTANH` | no (forks internally) | Element-wise `tanh` via `Parallel.For`. |
| `EPT.PARALLELROWREDUCE` | no (forks internally) | Per-row `sum`/`mean`/`min`/`max`/`stdev`. |
| `ParallelUtilities.ParallelBatchTransform` | API only | Partitioned `Parallel.For` with configurable degree and serial fallback below `DefaultParallelThreshold = 4096` rows. |

### Vectorized (SIMD) kernels (`VectorizedKernels.cs`)

All kernels select between an AVX2+FMA path (where supported), a portable
`Vector<T>` path, and a scalar fallback at runtime. Every UDF is
`IsThreadSafe = true` and safe inside Excel's multithreaded recalc.

| Function | Signature | Description |
| --- | --- | --- |
| `EPT.SIMD.ADD` | `(object[,], object[,]) -> object[,]` | Element-wise `a + b` with SIMD. |
| `EPT.SIMD.MULTIPLY` | `(object[,], object[,]) -> object[,]` | Element-wise `a * b` with SIMD. |
| `EPT.SIMD.SCALE` | `(object[,], double) -> object[,]` | Multiply every element by a scalar. |
| `EPT.SIMD.DOT` | `(object[,], object[,]) -> double` | Dot product (AVX2+FMA, portable, or scalar). |
| `EPT.SIMD.ROWSUMS` | `(object[,]) -> object[,]` | SIMD per-row sums. |
| `EPT.SIMD.COLSUMS` | `(object[,]) -> object[,]` | SIMD per-column sums. |
| `EPT.SIMD.NORMALIZE` | `(object[,]) -> object[,]` | L2 normalize a block as a flat vector. |
| `EPT.SIMD.MATMUL` | `(object[,], object[,]) -> object[,]` | Matrix multiply, packed inner loop. |
| `EPT.SIMD.CAPABILITIES` | `() -> object[,]` | Returns `(HardwareAccelerated, VectorWidth, Avx2, Fma)` for diagnostics. |

Internal .NET-only entry points (callable from other add-ins) take spans of
`double` directly: `ElementWiseAdd`, `ElementWiseMultiply`, `Scale`,
`DotProduct`, `RowSums`, `ColumnSums`, `L2Normalize`, `MatrixMultiply`.

### Real-time data (`RtdServer.cs`)

A multithreaded RTD server with ProgId `EPT.Rtd`. Topics multiplex over a
single producer per distinct feed spec; outbound updates are throttled by a
250 ms flush timer.

| Function | Thread-safe for MTR? | Description |
| --- | --- | --- |
| `EPT.RTD` | yes | Subscribe to a topic. Feed specs: `clock`, `counter:N`, `sine:freq:amp`, `random`. |
| `=RTD("EPT.Rtd",,topic)` | n/a | Direct Excel-native form, identical semantics. |

### Direct file I/O (`DirectFileIO.cs`)

Streaming async CSV/TSV read and write. Never touches the Excel object
model. Honors `ToolkitLifetime.ShutdownToken` for clean cancellation.

| Function | Signature | Description |
| --- | --- | --- |
| `EPT.READCSV` | `(path, delim?, encoding?, coerceNumeric?) -> object[,]` | Stream a delimited file into a bulk array. |
| `EPT.WRITECSV` | `(path, block, delim?, encoding?) -> double` | Stream a block to a delimited file; returns row count. |
| `DirectFileIO.ReadDelimitedAsync` | async API | Same as `EPT.READCSV` with a `CancellationToken`. |
| `DirectFileIO.WriteDelimitedAsync` | async API | Same as `EPT.WRITECSV` with a `CancellationToken`. |

### Conditional aggregation (`ConditionalAggregates.cs`)

Every function is `IsThreadSafe = true` (MTR-eligible) and pure CPU over its
`object[,]` arguments. Criteria follow Excel's grammar: comparison prefixes
(`">5"`, `"<=10"`, `"<>x"`), case-insensitive text with `*`/`?` wildcards
(`~` escapes a literal wildcard), the empty-operand blank/non-blank forms
(`""`, `"<>"`), and bare numeric criteria. Numeric comparisons match only
genuinely numeric cells - never numeric-looking text - exactly as Excel does.

| Function | Signature | Native? | Description |
| --- | --- | --- | --- |
| `EPT.COUNTIFS` | `(critRange1, crit1, ...) -> double` | faster | Count cells matching every pair. |
| `EPT.SUMIFS` | `(sumRange, critRange1, crit1, ...) -> double` | faster | Sum of matched numeric cells. |
| `EPT.AVERAGEIFS` | `(avgRange, ...) -> double` | faster | Mean of matched cells. |
| `EPT.MINIFS` / `EPT.MAXIFS` | `(range, ...) -> double` | faster | Min / max of matched cells. |
| `EPT.MEDIANIFS` | `(range, ...) -> double` | **new** | Conditional median. |
| `EPT.PERCENTILEIFS` | `(range, k, ...) -> double` | **new** | Conditional inclusive percentile, `k` in [0,1]. |
| `EPT.STDEVIFS` / `EPT.STDEVPIFS` | `(range, ...) -> double` | **new** | Conditional sample / population stdev. |
| `EPT.VARIFS` / `EPT.VARPIFS` | `(range, ...) -> double` | **new** | Conditional sample / population variance. |
| `EPT.MODEIFS` | `(range, ...) -> double` | **new** | Conditional mode (`#N/A` if none repeats). |
| `EPT.GEOMEANIFS` / `EPT.HARMEANIFS` | `(range, ...) -> double` | **new** | Conditional geometric / harmonic mean. |
| `EPT.PRODUCTIFS` | `(range, ...) -> double` | **new** | Conditional product. |
| `EPT.DISTINCTCOUNTIFS` | `(range, ...) -> double` | **new** | Count of distinct matched values. |
| `EPT.FIRSTIFS` / `EPT.LASTIFS` | `(range, ...) -> value` | **new** | First / last matched value. |
| `EPT.WAVG` | `(values, weights) -> double` | **new** | Weighted average. |
| `EPT.WAVGIFS` | `(values, weights, critRange1, crit1, ...) -> double` | **new** | Conditional weighted average. |
| `EPT.WMEDIAN` | `(values, weights) -> double` | **new** | Weighted median (lower). |
| `EPT.WSTDEV` | `(values, weights) -> double` | **new** | Weighted population stdev. |
| `EPT.SUMPRODUCTIFS` | `(rangeA, rangeB, critRange1, crit1, ...) -> double` | **new** | Conditional `sum(a*b)`. |
| `EPT.GROUPBY` | `(keyRange, valueRange, operation) -> object[,]` | **new** | One-pass GROUP BY; spills distinct keys + one aggregate. |

`EPT.GROUPBY` operations: `sum`, `count`, `average`, `min`, `max`, `median`,
`stdev`, `stdevp`, `var`, `varp`, `product`, `mode`, `geomean`, `harmean`,
`distinct`, `first`, `last`. `keyRange` may span several columns to form a
composite key.

### Boosted lookups (`LookupBoost.cs`)

| Function | Signature | Description |
| --- | --- | --- |
| `EPT.XLOOKUPB` | `(lookupValues, lookupArray, returnArray, [ifNotFound], [matchMode]) -> object[,]` | Resolve a whole column of keys in one O(M+R) pass. Exact (hash) by default; `matchMode = FALSE` does sorted-ascending approximate (next-smaller) match. Result mirrors `lookupValues`' shape; misses return `ifNotFound` or `#N/A`. **`IsThreadSafe = true`.** |

### Text utilities (`TextUtilities.cs`)

All MTR-safe, whole-block, one-crossing. Case functions pass non-string cells through;
pad/repeat/reverse pass blanks and errors through.

| Function | Signature | Description |
| --- | --- | --- |
| `EPT.PROPER` | `(block) -> object[,]` | Proper-case each word (first upper, rest lower). |
| `EPT.TITLECASE` | `(block) -> object[,]` | Title-case, preserving all-caps acronyms. |
| `EPT.CAMELCASE` | `(block) -> object[,]` | `camelCase`, dropping separators. |
| `EPT.PADLEFT` / `EPT.PADRIGHT` | `(block, totalWidth, [padChar]) -> object[,]` | Pad each cell's text to a width. |
| `EPT.ZEROPAD` | `(block, totalWidth) -> object[,]` | Left-pad with `0` to a width. |
| `EPT.REPEAT` | `(block, count) -> object[,]` | Repeat each cell's text N times. |
| `EPT.REVERSE` | `(block) -> object[,]` | Reverse each cell's characters. |
| `EPT.TEMPLATEFILL` | `(template, data, [hasHeaderRow]) -> object[,]` | Mail-merge: render `{name}`/`{index}` placeholders once per data row. |

### Regex utilities (`RegexUtilities.cs`)

Pattern compiled once per call with a 1s match timeout (ReDoS-safe). Per-cell timeouts
surface as `#VALUE!` in that cell; invalid patterns return `#VALUE!`.

| Function | Signature | Description |
| --- | --- | --- |
| `EPT.REGEXMATCH` | `(block, pattern, [ignoreCase]) -> object[,]` | TRUE/FALSE per cell. |
| `EPT.REGEXCOUNT` | `(block, pattern, [ignoreCase]) -> object[,]` | Count of matches per cell. |
| `EPT.REGEXEXTRACT` | `(block, pattern, [groupIndex], [ignoreCase]) -> object[,]` | First match (or capture group) per cell. |
| `EPT.REGEXEXTRACTALL` | `(column, pattern, [groupIndex], [ignoreCase]) -> object[,]` | All matches per row of a single column, spilled across columns. |
| `EPT.REGEXSPLIT` | `(column, pattern, [ignoreCase]) -> object[,]` | Split each cell of a single column by a regex, spilled across columns. |

### Series & robust stats (`SeriesUtilities.cs`)

| Function | Signature | Description |
| --- | --- | --- |
| `EPT.FILLFORWARD` | `(block, [direction]) -> object[,]` | Fill blanks from the previous value `down`/`up`/`right`/`left`. |
| `EPT.OUTLIERS` | `(block, [method], [threshold]) -> object[,]` | TRUE/FALSE outlier flags via `iqr`/`zscore`/`mad`. |
| `EPT.QUANTILES` | `(block, probabilities) -> object[,]` | Inclusive quantiles for each requested probability. |

### Dates (`DateUtilities.cs`)

| Function | Signature | Description |
| --- | --- | --- |
| `EPT.WORKDAYADD` | `(startDates, days, [weekendMask], [holidays]) -> object` | Add working days with a 7-char `Mon..Sun` weekend mask and holiday list; scalar or block, broadcast. |

### Distance (`DistanceUtilities.cs`)

| Function | Signature | Description |
| --- | --- | --- |
| `EPT.DISTANCE` | `(matrixA, [matrixB], [metric]) -> object[,]` | Pairwise distance matrix between row vectors. `euclidean` (default) and `cosine` route through the SIMD dot-product kernel; also `manhattan`, `chebyshev`. |

### JSON (`JsonUtilities.cs`)

Built on the in-box `System.Text.Json` (no third-party dependency). Path syntax is a
documented subset: dotted keys and `[index]` steps with an optional leading `$`/`$.`
(e.g. `data.items[0].name`). The two cell functions are MTR-safe; the three file functions
are async and never open a workbook.

| Function | Signature | MTR? | Description |
| --- | --- | --- | --- |
| `EPT.JSONPATH` | `(json, path) -> object[,]` | yes | Extract a value at `path` from the JSON in each cell. |
| `EPT.PARSEJSON` | `(json, [path], [hasHeaderRow]) -> object[,]` | yes | Expand one JSON document into a spilled table. |
| `EPT.READJSON` | `(path, [pointer], [hasHeaderRow]) -> object[,]` | no | Read a JSON file into a bulk array. |
| `EPT.READNDJSON` | `(path, [hasHeaderRow]) -> object[,]` | no | Read newline-delimited JSON, streamed line by line. |
| `EPT.WRITEJSON` | `(path, block, [hasHeaderRow], [indent]) -> double` | no | Write a block as a JSON array of objects/arrays; returns the record count. |

Internal .NET-only async entry points: `JsonUtilities.ReadJsonAsync`,
`ReadNdjsonAsync`, `WriteJsonAsync` (each takes a `CancellationToken`).

## Before and after: the one-crossing rule, demonstrated

### Before - the classic slow VBA loop (N boundary crossings)

```vba
' Trim every cell in A1:A100000. With N = 100,000 cells this performs
' 100,000 COM crossings to read AND 100,000 to write = 200,000 crossings.
Sub TrimSlow()
    Dim r As Range, c As Range
    Set r = Worksheets("Data").Range("A1:A100000")
    For Each c In r.Cells
        c.Value = Trim(c.Value)
    Next c
End Sub
```

Empirically: tens of seconds on a modern machine, locking the UI throughout.

### After - single bulk toolkit call (2 boundary crossings, total)

As a worksheet array formula that returns the trimmed block in place of the
loop:

```
=EPT.TRIMBLOCK(Data!A1:A100000)
```

Excel hands the UDF the range as one `object[,]` (one crossing in) and the UDF
returns one `object[,]` (one crossing out). Two crossings total, regardless of
*N*.

Or, equivalently, from VBA by reading and writing through the toolkit:

```vba
Sub TrimFast()
    Dim ws As Worksheet
    Set ws = Worksheets("Data")
    Dim arr As Variant
    ' Excel's own Range.Value already does a single bulk crossing when you
    ' assign a whole Range to a Variant - that's the trick the toolkit
    ' generalizes.
    arr = ws.Range("A1:A100000").Value          ' 1 crossing in
    ' Hand off to the toolkit for the managed-memory transform.
    arr = Application.Run("EPT.TRIMBLOCK", arr)
    ws.Range("A1:A100000").Value = arr          ' 1 crossing out
End Sub
```

The reduction is from *N* crossings (in the For-Each loop) to **2 crossings**
total. Both numbers are documented in the XML comments on `EPT.TRIMBLOCK` and
`BulkTransfer.RoundTripTransform`.

## Before and after: interpreted scalar loop -> vectorized kernel

### Before - scalar VBA arithmetic (no SIMD, interpreted)

```vba
' Element-wise multiply A1:A1000000 by B1:B1000000, write into C.
' Each iteration is interpreted; no vector instructions are ever issued.
Sub ScalarMultiply()
    Dim a As Variant, b As Variant, c() As Double
    a = Worksheets("Data").Range("A1:A1000000").Value
    b = Worksheets("Data").Range("B1:B1000000").Value
    ReDim c(1 To 1000000, 1 To 1)
    Dim i As Long
    For i = 1 To 1000000
        c(i, 1) = a(i, 1) * b(i, 1)            ' 1 interpreted op per element
    Next i
    Worksheets("Data").Range("C1:C1000000").Value = c
End Sub
```

### After - vectorized kernel (`Vector<double>` lanes, AVX2 where present)

```
=EPT.SIMD.MULTIPLY(Data!A1:A1000000, Data!B1:B1000000)
```

```vba
Sub VectorMultiply()
    Dim a As Variant, b As Variant, result As Variant
    a = Worksheets("Data").Range("A1:A1000000").Value
    b = Worksheets("Data").Range("B1:B1000000").Value
    result = Application.Run("EPT.SIMD.MULTIPLY", a, b)
    Worksheets("Data").Range("C1:C1000000").Value = result
End Sub
```

The kernel runs in compiled .NET, processes `Vector<double>.Count` elements
per iteration (4 on AVX2 hosts, 2 on SSE2-only hosts, 1 on hosts with no
hardware acceleration), and falls back to scalar on the tail. The same XLL
works correctly on any CPU - check `=EPT.SIMD.CAPABILITIES()` to see what
the host can use.

## Before and after: single-threaded feed -> multithreaded RTD server

### Before - polling on Excel's UI thread

```vba
' One Timer per cell, each pumping through the object model on the UI thread.
' At 50 subscribed cells and 10 Hz this monopolizes the UI.
Public Sub StartPolling()
    Application.OnTime Now + TimeValue("00:00:01"), "PollAll"
End Sub

Public Sub PollAll()
    Dim i As Long
    For i = 1 To 50
        Worksheets("Feeds").Cells(i, 1).Value = GetFeedValue(i)   ' COM crossing
    Next i
    Application.OnTime Now + TimeValue("00:00:01"), "PollAll"
End Sub
```

### After - throttled, multiplexed RTD server

```
=EPT.RTD("clock")
=EPT.RTD("counter:500")
=EPT.RTD("sine:2:5")
```

Or equivalently using Excel's native RTD function:

```
=RTD("EPT.Rtd",,"sine:2:5")
```

Subscriptions are de-duplicated by feed spec: 2,000 cells subscribing to
`clock` share a single background producer. Producers run on `ThreadPool`
threads; a single 250 ms throttle timer is the only thing that pushes new
values across the COM boundary into Excel - and only when the value has
actually changed.

## Before and after: object-model file read -> direct streaming I/O

### Before - opening a workbook just to read a CSV

```vba
' Spins up the full object model, fires Workbook_Open events, marshals every
' cell through the UI thread, and shows the file briefly in the workbook tabs.
Sub ImportCsvSlow()
    Dim wb As Workbook
    Set wb = Workbooks.Open("C:\data\prices.csv")
    Dim src As Range
    Set src = wb.Worksheets(1).UsedRange
    src.Copy Worksheets("Data").Range("A1")
    wb.Close SaveChanges:=False
End Sub
```

### After - direct streaming read, no object model touched

```
=EPT.READCSV("C:\data\prices.csv")
```

```vba
Sub ImportCsvFast()
    Dim arr As Variant
    arr = Application.Run("EPT.READCSV", "C:\data\prices.csv", ",")
    Worksheets("Data").Range("A1").Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
End Sub
```

The read is fully async under the hood, RFC 4180-compliant, and never
touches `Workbooks`, `Application`, or any other COM surface. The result
crosses the Excel boundary exactly once - on the bulk write into `Data!A1`.

## Before and after: per-formula re-scan -> one-pass conditional aggregate

### Before - a column of native SUMIFS (O(R*M))

```
' Filling C2:C100000, each SUMIFS re-scans A:A and B:B (M cells).
' R formulas x M criteria cells = O(R*M), recomputed on every edit.
C2: =SUMIFS($B$2:$B$100000, $A$2:$A$100000, ">"&D2)
... copied down 100,000 rows ...
```

### After - one MTR-safe call per cell, fanned across cores

```
=EPT.SUMIFS($B$2:$B$100000, $A$2:$A$100000, ">"&D2)
```

Each call builds the match mask in one managed-memory pass and is
`IsThreadSafe = true`, so Excel's multithreaded recalc engine schedules the
whole column across cores. The same engine answers questions Excel cannot
natively: a weighted average of matched rows in a single formula -

```
=EPT.WAVGIFS(Sales!Price, Sales!Qty, Sales!Region, "West", Sales!Year, 2025)
```

or a one-pass pivot that spills distinct keys with their aggregate -

```
=EPT.GROUPBY(Sales!Region, Sales!Revenue, "sum")
```

## Before and after: per-lookup re-scan -> one indexed pass

### Before - a column of VLOOKUP (O(R*M))

```
' Each VLOOKUP re-scans the M-row table; R lookups = O(R*M).
B2: =VLOOKUP(A2, Table!$A$2:$E$100000, 5, FALSE)
... copied down R rows ...
```

### After - build the index once, probe it for the whole column

```
=EPT.XLOOKUPB(A2:A100000, Table!A2:A100000, Table!E2:E100000)
```

`EPT.XLOOKUPB` hashes the lookup column once, then resolves every key in
O(M + R) total and spills the aligned results - exact match by default, or
sorted-ascending approximate match with a trailing `FALSE`.

## Engineering standards

- C# 12 on .NET 8 LTS, targeting `net8.0-windows`.
- Nullable reference types: **enabled**.
- Warnings as errors: **enabled** in `Directory.Build.props`.
- Unsafe code: **disabled** project-wide (no concrete justification was
  needed).
- Excel-DNA via `ExcelDna.AddIn` (packed XLL build).
- Late-bound COM via `ExcelDnaUtil.Application` is **not used** anywhere in
  this codebase; every hot path uses `ExcelReference` + `XlCall` and
  Excel-DNA's native `object[,]` marshaling.
- Every public function carries an XML doc comment stating its marshaling
  cost, its thread-safety, and the Excel types it accepts and returns.
- Input validation at the boundary: malformed ranges return an
  `ExcelError.ExcelErrorRef` with a logged reason rather than throwing into
  Excel.

## Which file addresses which bottleneck

| File | Bottleneck(s) addressed |
| --- | --- |
| `AddIn.cs` | Boot path for all six; logs registration metadata; resets `ToolkitLifetime` on open, cancels background work on close. |
| `Marshaling.cs` | **#2** (per-cell COM marshaling) - single fast classification layer used by every other file; also lays the groundwork for **#1** by making bulk blocks cheap to traverse. |
| `BulkTransfer.cs` | **#1** (cell-by-cell read/write) and **#2** (per-cell COM marshaling) - direct fix: one read crossing, one write crossing, regardless of cell count. |
| `DeveloperUtilities.cs` | **#1** and **#2** - every utility consumes and returns whole blocks, never re-entering COM mid-computation. |
| `ParallelUtilities.cs` | **#3** (VBA single-threaded recalc) - thread-safe UDFs let Excel's MTR engine schedule work across cores; the explicit `Parallel.For` path covers heavier batch transforms. |
| `VectorizedKernels.cs` | **#4** (interpreted runtime, no JIT or SIMD) - compiled .NET kernels with `Vector<T>` and AVX2+FMA paths, runtime-checked scalar fallback. |
| `RtdServer.cs` | **#5** (real-time feeds throttled through single-threaded marshaling) - multithreaded RTD server, concurrency-safe feed manager, throttled outbound updates. |
| `DirectFileIO.cs` | **#6** (file I/O via the object model) - async streaming CSV/TSV through managed file streams, no workbook ever opened. |
| `ConditionalAggregates.cs` | **#7** (conditional functions re-scan per formula) - one-pass match mask over bulk blocks, MTR-safe, plus the conditional aggregates and weighted statistics Excel lacks natively. |
| `LookupBoost.cs` | **#8** (lookups re-scan the table per formula) - build the table index once, probe it for a whole column of lookup values in O(M+R). |
| `TextUtilities.cs` | **#1**/**#2** - whole-block case, padding, repeat, reverse, and templated fill in pure managed memory. |
| `RegexUtilities.cs` | **#1**/**#2** - block-wide regex match/count/extract/split with a per-cell ReDoS timeout. |
| `SeriesUtilities.cs` | **#1**/**#2** - directional fill, outlier flagging, and quantiles over a whole block. |
| `DateUtilities.cs` | **#1**/**#2** - vectorized working-day arithmetic with custom weekends and holidays. |
| `DistanceUtilities.cs` | **#4** - pairwise distance matrices expressed through the SIMD dot-product kernel. |
| `JsonUtilities.cs` | **#1**/**#2** (in-grid JSON extraction) and **#6** (async file read/write that never opens a workbook), on the in-box `System.Text.Json`. |
| `ToolkitLifetime.cs` | Shared shutdown `CancellationTokenSource` and `TraceSource` factory used by the second-wave files. |

## See also

- `docs/usage.md` - copy-paste examples for every public function from both
  the formula bar and VBA, with the mandatory MTR safety warning.
