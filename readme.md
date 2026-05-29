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
| `ToolkitLifetime.cs` | Shared shutdown `CancellationTokenSource` and `TraceSource` factory used by the second-wave files. |

## See also

- `docs/usage.md` - copy-paste examples for every public function from both
  the formula bar and VBA, with the mandatory MTR safety warning.
