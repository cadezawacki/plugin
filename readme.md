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

## Architecture

```
ExcelPerfToolkit (single net8.0-windows class library, packed as XLL)
├── AddIn.cs                 - IExcelAddIn entry point (boots #1, #2, #3)
├── Marshaling.cs            - Allocation-conscious helpers (#2)
├── BulkTransfer.cs          - Read/write/round-trip a whole range (#1, #2)
├── DeveloperUtilities.cs    - The day-to-day utility surface  (#1, #2)
└── ParallelUtilities.cs     - Thread-safe UDFs + Parallel.For (#3)
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
| `AddIn.cs` | Boot path for all three; logs registration metadata. |
| `Marshaling.cs` | **#2** (per-cell COM marshaling) - single fast classification layer used by every other file; also lays the groundwork for **#1** by making bulk blocks cheap to traverse. |
| `BulkTransfer.cs` | **#1** (cell-by-cell read/write) and **#2** (per-cell COM marshaling) - direct fix: one read crossing, one write crossing, regardless of cell count. |
| `DeveloperUtilities.cs` | **#1** and **#2** - every utility consumes and returns whole blocks, never re-entering COM mid-computation. |
| `ParallelUtilities.cs` | **#3** (VBA single-threaded recalc) - thread-safe UDFs let Excel's MTR engine schedule work across cores; the explicit `Parallel.For` path covers heavier batch transforms. |

## See also

- `docs/usage.md` - copy-paste examples for every public function from both
  the formula bar and VBA, with the mandatory MTR safety warning.
