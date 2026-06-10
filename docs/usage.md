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

## Conditional aggregation (MTR-eligible)

All `EPT.*IFS`, weighted, `SUMPRODUCTIFS`, and `GROUPBY` functions are
registered with `IsThreadSafe = true`. They are pure CPU over their
`object[,]` arguments and never touch the Excel object model. Criteria use
Excel's grammar: `">5"`, `"<=10"`, `"<>x"`, `"=v"`, the wildcards `*`/`?`
(with `~` escaping a literal wildcard), and the empty-operand blank/non-blank
forms `""` and `"<>"`. Numeric comparisons match only genuinely numeric
cells, not numeric-looking text - exactly like Excel.

### Faster `*IFS` (drop-in for the native versions)

```
=EPT.COUNTIFS(Data!A1:A100000, ">0", Data!B1:B100000, "West")
=EPT.SUMIFS(Data!C1:C100000, Data!A1:A100000, ">0", Data!B1:B100000, "West")
=EPT.AVERAGEIFS(Data!C1:C100000, Data!B1:B100000, "W*")
=EPT.MINIFS(Data!C1:C100000, Data!B1:B100000, "West")
=EPT.MAXIFS(Data!C1:C100000, Data!B1:B100000, "West")
```

```vba
Dim total As Double
total = Application.Run("EPT.SUMIFS", _
        Worksheets("Data").Range("C1:C100000").Value, _
        Worksheets("Data").Range("A1:A100000").Value, ">0", _
        Worksheets("Data").Range("B1:B100000").Value, "West")
```

### Conditional aggregates Excel lacks natively

```
=EPT.MEDIANIFS(Data!C1:C100000, Data!B1:B100000, "West")
=EPT.PERCENTILEIFS(Data!C1:C100000, 0.95, Data!B1:B100000, "West")
=EPT.STDEVIFS(Data!C1:C100000, Data!B1:B100000, "West")
=EPT.VARPIFS(Data!C1:C100000, Data!B1:B100000, "West")
=EPT.MODEIFS(Data!C1:C100000, Data!B1:B100000, "West")
=EPT.GEOMEANIFS(Data!C1:C100000, Data!B1:B100000, "West")
=EPT.HARMEANIFS(Data!C1:C100000, Data!B1:B100000, "West")
=EPT.PRODUCTIFS(Data!C1:C100000, Data!B1:B100000, "West")
=EPT.DISTINCTCOUNTIFS(Data!C1:C100000, Data!B1:B100000, "West")
=EPT.FIRSTIFS(Data!C1:C100000, Data!B1:B100000, "West")
=EPT.LASTIFS(Data!C1:C100000, Data!B1:B100000, "West")
```

### Weighted statistics

```
=EPT.WAVG(Data!Price, Data!Qty)
=EPT.WAVGIFS(Data!Price, Data!Qty, Data!Region, "West", Data!Year, 2025)
=EPT.WMEDIAN(Data!Price, Data!Qty)
=EPT.WSTDEV(Data!Price, Data!Qty)
```

```vba
Dim w As Double
w = Application.Run("EPT.WAVG", _
        Worksheets("Data").Range("A1:A100000").Value, _
        Worksheets("Data").Range("B1:B100000").Value)
```

### Conditional SUMPRODUCT

```
=EPT.SUMPRODUCTIFS(Data!Price, Data!Qty, Data!Region, "West")
```

### `EPT.GROUPBY(key_range, value_range, operation)`

Returns a spilled table of distinct keys with one aggregate per group, in
first-seen order. `key_range` may span several columns to form a composite
key; the first column of `value_range` supplies the values. `operation` is
one of `sum`, `count`, `average`, `min`, `max`, `median`, `stdev`, `stdevp`,
`var`, `varp`, `product`, `mode`, `geomean`, `harmean`, `distinct`, `first`,
`last`.

```
=EPT.GROUPBY(Sales!A2:A100000, Sales!D2:D100000, "sum")
=EPT.GROUPBY(Sales!A2:B100000, Sales!D2:D100000, "average")
```

```vba
Dim g As Variant
g = Application.Run("EPT.GROUPBY", _
        Worksheets("Sales").Range("A2:A100000").Value, _
        Worksheets("Sales").Range("D2:D100000").Value, "sum")
```

## Boosted lookups

### `EPT.XLOOKUPB(lookup_values, lookup_array, return_array, [if_not_found], [match_mode])`

Resolves an entire column of lookup keys in one O(M+R) pass: the lookup
column is indexed once, then probed for every key. The result mirrors the
shape of `lookup_values`. Exact match (default) is case-insensitive; pass
`FALSE` as `match_mode` for sorted-ascending approximate (next-smaller)
match. Misses return `if_not_found` when supplied, otherwise `#N/A`.

```
=EPT.XLOOKUPB(A2:A100000, Table!A2:A100000, Table!E2:E100000)
=EPT.XLOOKUPB(A2:A100000, Table!A2:A100000, Table!E2:E100000, "n/a")
=EPT.XLOOKUPB(A2:A100000, Brackets!A2:A20, Brackets!B2:B20, "", FALSE)
```

```vba
Dim res As Variant
res = Application.Run("EPT.XLOOKUPB", _
        Worksheets("Keys").Range("A2:A100000").Value, _
        Worksheets("Table").Range("A2:A100000").Value, _
        Worksheets("Table").Range("E2:E100000").Value)
```

## Text utilities (MTR-eligible)

All `EPT.*` text functions are `IsThreadSafe = true`, whole-block, one-crossing. Case
functions leave non-string cells untouched; pad/repeat/reverse pass blanks and errors
through.

```
=EPT.PROPER(Data!A1:A1000)
=EPT.TITLECASE(Data!A1:A1000)
=EPT.CAMELCASE(Data!A1:A1000)
=EPT.PADLEFT(Data!A1:A1000, 10)
=EPT.PADRIGHT(Data!A1:A1000, 10, ".")
=EPT.ZEROPAD(Data!A1:A1000, 6)
=EPT.REPEAT(Data!A1:A1000, 3)
=EPT.REVERSE(Data!A1:A1000)
```

`EPT.TEMPLATEFILL(template, data, [has_header_row])` renders a template once per data
row. Placeholders are `{HeaderName}` (when a header row is present) or `{0}`-style column
indexes; use `{{`/`}}` for literal braces.

```
=EPT.TEMPLATEFILL("Dear {Name}, your balance is {Amount}.", Data!A1:B100)
=EPT.TEMPLATEFILL("{0}-{1}", Data!A2:B100, FALSE)
```

```vba
arr = Application.Run("EPT.TEMPLATEFILL", "Hi {Name}", Worksheets("Data").Range("A1:A100").Value)
```

## Regex utilities (MTR-eligible)

Pattern is compiled once per call with a 1-second match timeout. `ignore_case` defaults to
FALSE. `EPT.REGEXEXTRACTALL` and `EPT.REGEXSPLIT` require a single-column input and spill
across columns.

```
=EPT.REGEXMATCH(Data!A1:A1000, "^[A-Z]{2}\d+$")
=EPT.REGEXCOUNT(Data!A1:A1000, "\d")
=EPT.REGEXEXTRACT(Data!A1:A1000, "(\d{4})-(\d{2})", 1)
=EPT.REGEXEXTRACTALL(Data!A1:A1000, "\d+")
=EPT.REGEXSPLIT(Data!A1:A1000, "\s*,\s*")
```

```vba
arr = Application.Run("EPT.REGEXEXTRACT", arr, "[0-9]+", 0, False)
```

## Series cleaning & robust stats (MTR-eligible)

```
=EPT.FILLFORWARD(Data!A1:D1000)            ' down (default)
=EPT.FILLFORWARD(Data!A1:D1000, "right")
=EPT.OUTLIERS(Data!A1:A1000)               ' iqr, 1.5x fences
=EPT.OUTLIERS(Data!A1:A1000, "zscore", 2.5)
=EPT.QUANTILES(Data!A1:A1000, {0;0.25;0.5;0.75;1})
```

```vba
flags = Application.Run("EPT.OUTLIERS", arr, "mad", 3)
```

## Working-day arithmetic

`EPT.WORKDAYADD(start_dates, days, [weekend_mask], [holidays])` adds working days using a
7-character `Mon..Sun` weekend mask (`'1'` = non-working, default `"0000011"`). `start_dates`
and `days` may each be a scalar or a block (a scalar is broadcast).

```
=EPT.WORKDAYADD(A2, 10)
=EPT.WORKDAYADD(A2:A100, 5, "0000011", Holidays!A1:A20)
=EPT.WORKDAYADD(A2, -3, "1000001")        ' Monday & Sunday as weekend
```

```vba
serial = Application.Run("EPT.WORKDAYADD", DateSerial(2026, 1, 2), 10)
```

## Pairwise distance

`EPT.DISTANCE(matrix_a, [matrix_b], [metric])` treats each row as an observation vector and
returns the full distance matrix. Omit `matrix_b` to compare `matrix_a` to itself.

```
=EPT.DISTANCE(Data!A2:E50)                 ' 49x49 euclidean self-distances
=EPT.DISTANCE(Data!A2:E50, Centroids!A2:E5, "cosine")
=EPT.DISTANCE(Data!A2:E50, Data!A2:E50, "manhattan")
```

```vba
m = Application.Run("EPT.DISTANCE", a, b, "euclidean")
```

## JSON

Built on the in-box `System.Text.Json`. `EPT.JSONPATH` and `EPT.PARSEJSON` are MTR-safe
and operate on JSON text already in the grid; `EPT.READJSON`/`EPT.READNDJSON`/`EPT.WRITEJSON`
are async file functions that never open a workbook. Path syntax is dotted keys plus
`[index]` with an optional leading `$`/`$.`.

### `EPT.JSONPATH(json, path)` and `EPT.PARSEJSON(json, [path], [hasHeaderRow])`

```
=EPT.JSONPATH(A1:A1000, "address.city")
=EPT.JSONPATH(A1:A1000, "items[0].sku")
=EPT.PARSEJSON(A1)                          ' expand a document to a table
=EPT.PARSEJSON(A1, "data.rows")             ' expand a sub-node
=EPT.PARSEJSON(A1, "", FALSE)               ' no header row
```

```vba
city = Application.Run("EPT.JSONPATH", arr, "address.city")
tbl = Application.Run("EPT.PARSEJSON", jsonText, "data.rows")
```

### `EPT.READJSON` / `EPT.READNDJSON` / `EPT.WRITEJSON`

```
=EPT.READJSON("C:\data\orders.json")
=EPT.READJSON("C:\data\payload.json", "data.orders")
=EPT.READNDJSON("C:\data\events.ndjson")
=EPT.WRITEJSON("C:\out\rows.json", Data!A1:D1000)        ' array of objects (row 1 = keys)
=EPT.WRITEJSON("C:\out\rows.json", Data!A1:D1000, FALSE, TRUE)  ' arrays, pretty-printed
```

```vba
arr = Application.Run("EPT.READJSON", "C:\data\orders.json")
Worksheets("Data").Range("A1").Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
n = Application.Run("EPT.WRITEJSON", "C:\out\rows.json", _
                    Worksheets("Data").Range("A1:D1000").Value)
```

## Filesystem

`EPT.FILEINFO` and `EPT.READFOLDER` work directly against the filesystem (no workbook
open). `EPT.WATCHFILE`/`EPT.WATCHFOLDER` are live RTD triggers.

```
=EPT.FILEINFO("C:\data\prices.csv")
=EPT.FILEINFO(A1:A20)                          ' metadata for a column of paths
=EPT.READFOLDER("C:\data", "*.csv")            ' concatenate, align by header
=EPT.READFOLDER("C:\data", "*.json", TRUE)     ' recurse subfolders
=EPT.WATCHFILE("C:\data\prices.csv")           ' increments when the file changes
=EPT.WATCHFOLDER("C:\data")                     ' increments on any change in the folder
```

A common pattern: make an import depend on the watch trigger so it re-runs on change.

```
=IF(EPT.WATCHFILE("C:\data\prices.csv")>=0, EPT.READCSV("C:\data\prices.csv"), "")
```

## Result caching

`EPT.MEMOIZE`/`EPT.CACHE.*` are an in-memory session cache; `EPT.DISKCACHE.*` persists
across reopen. Build keys from input content with `EPT.HASHBLOCK`. (Excel computes a UDF's
arguments first, so these store and reuse results - they do not stop the inner formula from
computing once.)

```
=EPT.MEMOIZE("prices_v1", EPT.READCSV("C:\data\prices.csv"))
=EPT.CACHE.GET("prices_v1")
=EPT.CACHE.GET("missing", "n/a")
=EPT.CACHE.CLEAR()                              ' clear everything; returns count
=EPT.DISKCACHE.WRITE("daily_"&TODAY(), Data!A1:Z1000)
=EPT.DISKCACHE.READ("daily_"&TODAY())
=EPT.DISKCACHE.CLEAR("daily_"&TODAY())
```

```vba
Application.Run "EPT.DISKCACHE.WRITE", "snapshot", Worksheets("Data").Range("A1:Z1000").Value
arr = Application.Run("EPT.DISKCACHE.READ", "snapshot")
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
