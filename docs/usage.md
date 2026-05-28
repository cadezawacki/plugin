# ExcelPerfToolkit usage

This document gives copy-paste examples for **every public function** in the
toolkit, used two ways:

- **Formula** - typed into a cell, usually as a dynamic-array formula.
- **VBA** - invoked via `Application.Run` so a macro can drive the toolkit
  without taking a separate reference.

> **Critical safety warning for thread-safe UDFs.** Functions registered with
> `IsThreadSafe = true` (every `EPT.MT.*` function below) execute on Excel's
> multithreaded recalc worker threads. They **must never** touch the Excel
> object model. That means: no `ExcelDnaUtil.Application`, no
> `ExcelReference.GetValue` / `SetValue`, no `XlCall` entries beyond the
> handful Microsoft documents as thread-safe (`xlfXxx` evaluation functions
> over already-marshaled arguments). Doing so will crash Excel, corrupt
> recalc state, or hang on the COM apartment. Pure CPU work over the
> `object[,]` arguments only.

## Conventions

- `Data!A1:Z1000` is the example source range throughout.
- Where a formula returns an array, the examples show it as a single
  dynamic-array formula. Older Excel versions need it entered as an array
  formula with `Ctrl+Shift+Enter` and a pre-selected output range.
- All function names use the `EPT.` prefix and are case-insensitive in Excel.

## Bulk transfer

### `EPT.READRANGE(sheet, a1)`

Read a whole range as a bulk array.

```
=EPT.READRANGE("Data","A1:D1000")
```

```vba
Sub LoadBlockExample()
    Dim block As Variant
    block = Application.Run("EPT.READRANGE", "Data", "A1:D1000")
    ' block is now Variant(1 To 1000, 1 To 4) and arrived in ONE crossing.
End Sub
```

## Developer utilities

### `EPT.TRIMBLOCK(block)`

Trim and collapse whitespace across every string cell.

```
=EPT.TRIMBLOCK(Data!A1:D1000)
```

```vba
Sub TrimExample()
    Dim arr As Variant
    arr = Worksheets("Data").Range("A1:D1000").Value
    arr = Application.Run("EPT.TRIMBLOCK", arr)
    Worksheets("Data").Range("A1:D1000").Value = arr
End Sub
```

### `EPT.COERCENUMERIC(block)`

Convert numeric-looking text to real numbers.

```
=EPT.COERCENUMERIC(Data!A1:A500)
```

```vba
arr = Application.Run("EPT.COERCENUMERIC", arr)
```

### `EPT.REMOVEDUPLICATEROWS(block)`

Drop duplicate rows, keeping the first occurrence.

```
=EPT.REMOVEDUPLICATEROWS(Data!A1:D5000)
```

```vba
arr = Application.Run("EPT.REMOVEDUPLICATEROWS", arr)
```

### `EPT.TRANSPOSE(block)`

Transpose a block (toolkit version - not Excel's built-in TRANSPOSE).

```
=EPT.TRANSPOSE(Data!A1:Z100)
```

```vba
arr = Application.Run("EPT.TRANSPOSE", arr)
```

### `EPT.FILLBLANKS(block, fillValue)`

Replace blank cells with a supplied value.

```
=EPT.FILLBLANKS(Data!A1:D1000, 0)
=EPT.FILLBLANKS(Data!A1:D1000, "N/A")
```

```vba
arr = Application.Run("EPT.FILLBLANKS", arr, 0)
```

### `EPT.FILLDOWN(block)`

Replace blanks with the most recent non-blank value above, per column.
Useful for cleaning hierarchical exports.

```
=EPT.FILLDOWN(Data!A1:D1000)
```

```vba
arr = Application.Run("EPT.FILLDOWN", arr)
```

### `EPT.SPLITCOLUMN(column, delimiter)`

Split one column into many by a delimiter.

```
=EPT.SPLITCOLUMN(Data!A1:A1000, "|")
```

```vba
arr = Application.Run("EPT.SPLITCOLUMN", Worksheets("Data").Range("A1:A1000").Value, "|")
```

### `EPT.JOINCOLUMNS(block, delimiter)`

Join every column of each row into a single column.

```
=EPT.JOINCOLUMNS(Data!A1:D1000, " | ")
```

```vba
arr = Application.Run("EPT.JOINCOLUMNS", arr, " | ")
```

### `EPT.FINDREPLACE(block, find, replace, useRegex)`

Find-and-replace across every string cell. Pass `TRUE` for regex mode.

```
=EPT.FINDREPLACE(Data!A1:D1000, "  ", " ", FALSE)
=EPT.FINDREPLACE(Data!A1:D1000, "[0-9]+", "#", TRUE)
```

```vba
arr = Application.Run("EPT.FINDREPLACE", arr, "abc", "xyz", False)
arr = Application.Run("EPT.FINDREPLACE", arr, "\s+", " ", True)
```

### `EPT.STACKCOLUMNS(block)`

Stack every column into a single column, in column-major order.

```
=EPT.STACKCOLUMNS(Data!A1:D250)
```

```vba
arr = Application.Run("EPT.STACKCOLUMNS", arr)
```

### `EPT.UNPIVOT(block, keyColumns)`

Reshape a wide block into a long block. The first `keyColumns` columns are
preserved as keys; the remaining columns become Attribute/Value pairs. Row 0
is treated as the header row.

```
=EPT.UNPIVOT(Data!A1:F500, 2)
```

```vba
arr = Application.Run("EPT.UNPIVOT", arr, 2)
```

### `EPT.UNIQUE(block)`

Distinct non-blank values, in first-seen order, as a single column.

```
=EPT.UNIQUE(Data!A1:D1000)
```

```vba
arr = Application.Run("EPT.UNIQUE", arr)
```

### `EPT.UNIQUECOUNT(block)`

Distinct values with occurrence counts (two columns: value, count).

```
=EPT.UNIQUECOUNT(Data!A1:A1000)
```

```vba
arr = Application.Run("EPT.UNIQUECOUNT", arr)
```

### `EPT.BLOCKLOOKUP(left, leftKeyColumnIndex, lookup, returnColumnIndexes)`

Hash-join the left block to the lookup block. The lookup's column 0 is the
key. `returnColumnIndexes` is a small range of 0-based column indexes into
the lookup table.

```
=EPT.BLOCKLOOKUP(Data!A1:C1000, 0, Lookup!A1:E50, {1;2;4})
```

```vba
Dim leftBlk As Variant, lkpBlk As Variant, cols(0 To 2) As Variant
leftBlk = Worksheets("Data").Range("A1:C1000").Value
lkpBlk = Worksheets("Lookup").Range("A1:E50").Value
cols(0) = 1: cols(1) = 2: cols(2) = 4
arr = Application.Run("EPT.BLOCKLOOKUP", leftBlk, 0, lkpBlk, cols)
```

### `EPT.SORTBLOCK(block, keyColumns, descendingFlags)`

Stable sort by one or more key columns. `keyColumns` is a small range of
0-based column indexes; `descendingFlags` is an optional matching range of
TRUE/FALSE.

```
=EPT.SORTBLOCK(Data!A1:D1000, {2;0}, {TRUE;FALSE})
```

```vba
Dim keys(0 To 1) As Variant, dirs(0 To 1) As Variant
keys(0) = 2: keys(1) = 0
dirs(0) = True: dirs(1) = False
arr = Application.Run("EPT.SORTBLOCK", arr, keys, dirs)
```

### `EPT.HASHBLOCK(block)`

Fast, stable XXH3 hex digest. Good for change detection.

```
=EPT.HASHBLOCK(Data!A1:Z1000)
```

```vba
Dim h As String
h = Application.Run("EPT.HASHBLOCK", arr)
```

### `EPT.SHA256BLOCK(block)`

Cryptographic SHA-256 hex digest. Slower; collision-resistant.

```
=EPT.SHA256BLOCK(Data!A1:Z1000)
```

```vba
Dim h As String
h = Application.Run("EPT.SHA256BLOCK", arr)
```

## Thread-safe (MTR-eligible) UDFs

All of the following are registered with `IsThreadSafe = true`. They are
**pure CPU**; they do not and must not touch the Excel object model. Excel
schedules them across cores during a normal recalc.

### `EPT.MT.ROWSUMS(block)`

```
=EPT.MT.ROWSUMS(Data!A1:Z1000)
```

```vba
arr = Application.Run("EPT.MT.ROWSUMS", arr)
```

### `EPT.MT.ROWZSCORES(block)`

```
=EPT.MT.ROWZSCORES(Data!A1:Z1000)
```

```vba
arr = Application.Run("EPT.MT.ROWZSCORES", arr)
```

### `EPT.MT.POLYEVAL(block, coefficients)`

Coefficients are taken low-order-first from the flattened row-major form of
`coefficients`. For example `{1, 2, 3}` evaluates `1 + 2x + 3x^2`.

```
=EPT.MT.POLYEVAL(Data!A1:Z1000, {1, -1, 0.5})
```

```vba
Dim coeffs(0 To 2) As Variant
coeffs(0) = 1: coeffs(1) = -1: coeffs(2) = 0.5
arr = Application.Run("EPT.MT.POLYEVAL", arr, coeffs)
```

### `EPT.MT.DOT(a, b)`

```
=EPT.MT.DOT(A1:A1000, B1:B1000)
```

```vba
Dim a As Variant, b As Variant, d As Double
a = Worksheets("Data").Range("A1:A1000").Value
b = Worksheets("Data").Range("B1:B1000").Value
d = Application.Run("EPT.MT.DOT", a, b)
```

## Explicit parallel batch transforms

These functions drive their own `Parallel.For` internally. They are
**not** registered as `IsThreadSafe` because we don't want Excel's MTR
engine to additionally fan them out across MTR worker threads (which would
multiply the thread count). Call them once per recalc on a large block.

### `EPT.PARALLELTANH(block)`

```
=EPT.PARALLELTANH(Data!A1:Z10000)
```

```vba
arr = Application.Run("EPT.PARALLELTANH", arr)
```

### `EPT.PARALLELROWREDUCE(block, operation)`

`operation` is one of `"sum"`, `"mean"`, `"min"`, `"max"`, `"stdev"`.

```
=EPT.PARALLELROWREDUCE(Data!A1:Z10000, "stdev")
```

```vba
arr = Application.Run("EPT.PARALLELROWREDUCE", arr, "mean")
```

## Vectorized (SIMD) kernels

All `EPT.SIMD.*` UDFs are registered with `IsThreadSafe = true`. They
operate entirely on numeric content; non-numeric cells are coerced to 0.
The SIMD path is gated at runtime: hosts without `Vector<T>` hardware
acceleration silently use the scalar fallback and produce identical
results.

### `EPT.SIMD.CAPABILITIES()`

Returns a one-row block of `(HardwareAccelerated, VectorWidth, Avx2, Fma)`.

```
=EPT.SIMD.CAPABILITIES()
```

```vba
Dim caps As Variant
caps = Application.Run("EPT.SIMD.CAPABILITIES")
```

### `EPT.SIMD.ADD(a, b)` / `EPT.SIMD.MULTIPLY(a, b)`

Element-wise add or multiply over two equally-shaped numeric blocks.

```
=EPT.SIMD.ADD(Data!A1:Z1000, Data!AA1:AZ1000)
=EPT.SIMD.MULTIPLY(Data!A1:Z1000, Data!AA1:AZ1000)
```

```vba
arr = Application.Run("EPT.SIMD.ADD", a, b)
arr = Application.Run("EPT.SIMD.MULTIPLY", a, b)
```

### `EPT.SIMD.SCALE(block, factor)`

Multiply every element by a scalar.

```
=EPT.SIMD.SCALE(Data!A1:Z1000, 2.5)
```

```vba
arr = Application.Run("EPT.SIMD.SCALE", a, 2.5)
```

### `EPT.SIMD.DOT(a, b)`

Dot product over two flattened equal-length vectors. AVX2+FMA path is used
when supported.

```
=EPT.SIMD.DOT(Data!A1:A100000, Data!B1:B100000)
```

```vba
Dim d As Double
d = Application.Run("EPT.SIMD.DOT", a, b)
```

### `EPT.SIMD.ROWSUMS(block)` / `EPT.SIMD.COLSUMS(block)`

```
=EPT.SIMD.ROWSUMS(Data!A1:Z10000)
=EPT.SIMD.COLSUMS(Data!A1:Z10000)
```

```vba
arr = Application.Run("EPT.SIMD.ROWSUMS", a)
arr = Application.Run("EPT.SIMD.COLSUMS", a)
```

### `EPT.SIMD.NORMALIZE(block)`

L2 normalize the block treated as a flat vector.

```
=EPT.SIMD.NORMALIZE(Data!A1:Z1000)
```

```vba
arr = Application.Run("EPT.SIMD.NORMALIZE", a)
```

### `EPT.SIMD.MATMUL(a, b)`

Matrix multiply; inner dimensions must agree.

```
=EPT.SIMD.MATMUL(Data!A1:C100, Data!E1:G3)
```

```vba
arr = Application.Run("EPT.SIMD.MATMUL", a, b)
```

## Real-time data: `EPT.RTD`

Subscribe to a topic on the multithreaded `EPT.Rtd` server. Updates push
from background threads through a 250 ms throttle into Excel.

```
=EPT.RTD("clock")               ' wall-clock UTC, ticks ~4x/sec on the producer, throttled to 4 Hz outbound
=EPT.RTD("counter:500")         ' increments every 500 ms
=EPT.RTD("sine:2:5")            ' sine wave, freq=2 Hz, amplitude=5
=EPT.RTD("random")              ' uniform random doubles in [0, 1)
```

The native Excel form is equivalent:

```
=RTD("EPT.Rtd",,"sine:2:5")
```

From VBA, drive RTD subscriptions by writing the formula into a cell. RTD
subscriptions cannot be opened by `Application.Run`; that is by design - the
server is owned by Excel's RTD lifecycle:

```vba
Worksheets("Live").Range("A1").Formula = "=EPT.RTD(""clock"")"
```

> **Critical safety warning - restated.** The RTD producer threads (the
> `Feed.RunAsync` loops in `RtdServer.cs`) and the throttle flush timer
> **must never** touch the Excel object model. `Topic.UpdateValue` is the
> single safe boundary into Excel; Excel-DNA marshals it onto Excel's
> dedicated RTD thread. Any thread-safe UDF you write that participates in
> RTD acquisition must obey the same rule: no `ExcelDnaUtil.Application`,
> no `ExcelReference.GetValue` / `SetValue`, no non-pure `XlCall` entries.
> Doing so will crash Excel or hang the COM apartment.

## Direct file I/O

### `EPT.READCSV(path, [delimiter], [encoding], [coerceNumeric])`

Stream a delimited text file directly into a bulk array.

```
=EPT.READCSV("C:\data\prices.csv")
=EPT.READCSV("C:\data\prices.tsv","\t")
=EPT.READCSV("C:\data\prices.csv",",","windows-1252",FALSE)
```

```vba
Dim arr As Variant
arr = Application.Run("EPT.READCSV", "C:\data\prices.csv")
Worksheets("Data").Range("A1").Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
```

### `EPT.WRITECSV(path, block, [delimiter], [encoding])`

Stream a worksheet block out to a delimited file. Returns the row count.

```
=EPT.WRITECSV("C:\out\result.csv", Data!A1:Z1000)
=EPT.WRITECSV("C:\out\result.tsv", Data!A1:Z1000,"\t")
```

```vba
Dim n As Double
n = Application.Run("EPT.WRITECSV", "C:\out\result.csv", _
                    Worksheets("Data").Range("A1:Z1000").Value)
```

## Public .NET-only entry points

These are not registered as UDFs because they take delegates,
`ExcelReference`, spans of `double`, or `CancellationToken` directly, but
they are public on the assembly so any other add-in can consume them.

- `BulkTransfer.ResolveRange(sheet, a1) -> ExcelReference`
- `BulkTransfer.ReadBlock(ExcelReference) -> object[,]`
- `BulkTransfer.WriteBlock(ExcelReference, object[,])`
- `BulkTransfer.RoundTripTransform(ExcelReference, Func<object[,], object[,]>)`
- `BulkTransfer.ReadDoubleBlock` / `WriteDoubleBlock` / `ReadStringBlock`
- `ParallelUtilities.ParallelBatchTransform(double[,], double[,], RowTransform, int, int?)`
- `VectorizedKernels.ElementWiseAdd` / `ElementWiseMultiply` / `Scale` / `DotProduct` / `RowSums` / `ColumnSums` / `L2Normalize` / `MatrixMultiply` - all on `Span<double>` / `ReadOnlySpan<double>`.
- `DirectFileIO.ReadDelimitedAsync(path, delim, encoding, coerceNumeric, ct) -> Task<object[,]>`
- `DirectFileIO.WriteDelimitedAsync(path, block, delim, encoding, newLine, ct) -> Task`
- `Marshaling.*` - the full conversion surface.

## VBA pattern: read-bulk, transform-managed, write-bulk

This is the canonical pattern that replaces every cell-by-cell loop:

```vba
Sub BulkPattern(sheetName As String, addr As String)
    Dim ws As Worksheet
    Set ws = Worksheets(sheetName)
    Dim arr As Variant
    arr = ws.Range(addr).Value                                ' 1 crossing in
    arr = Application.Run("EPT.TRIMBLOCK", arr)               ' pure managed
    arr = Application.Run("EPT.COERCENUMERIC", arr)           ' pure managed
    arr = Application.Run("EPT.FILLBLANKS", arr, 0)           ' pure managed
    ws.Range(addr).Value = arr                                ' 1 crossing out
End Sub
```

Two crossings total, regardless of how many cells are involved.
