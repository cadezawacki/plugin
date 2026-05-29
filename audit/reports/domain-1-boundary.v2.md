# Domain 1 — Boundary & Conversion (round 2)

**Files in scope:** `AddIn.cs`, `Marshaling.cs`, `BulkTransfer.cs`
**Auditor stance:** Round 2, fresh eyes. Round 1 produced shallow "might fail" findings; this pass traces every concrete path to a specific input or interleaving with BCL-contract citations.

## Findings table

| ID | Severity | Location | Category | One-line scenario | Confidence |
| --- | --- | --- | --- | --- | --- |
| V001 | Critical | Marshaling.cs:104-110 (`TryToDouble` string branch) | Data corruption | `"1,5"` parses to `15.0` via `InvariantCulture + AllowThousands`; user-entered decimals silently wrong by 10x | 0.95 |
| V002 | High | Marshaling.cs:101-103 (`TryToDouble` DateTime branch) | Unhandled exception | `DateTime` with year 1–99 (other than `MinValue`) throws `OverflowException` from `ToOADate()`; `TryToDouble` lets it escape | 0.90 |
| V003 | High | BulkTransfer.cs:46-55 (`ResolveRange`) | Unhandled exception | `xlfEvaluate` raises `XlCallException` for many failure modes; caller `ReadRangeUdf` only catches generic `Exception` (works), but other callers (`ReadBlock(string,string)`, `WriteBlock(string,string,object[,])`, `RoundTripTransform(string,…)`) propagate it as `XlCallException`, not the documented `ArgumentException` | 0.85 |
| V004 | High | AddIn.cs:62-73 (`SafeGetExcelVersion`) | Error swallowing / wrong catch scope | The catch only handles `XlCallException`. `XlCall.Excel` returns an `ExcelError` boxed in `object` for "not ready" rather than throwing in some host states; `ExcelError.ExcelErrorNA.ToString()` yields `"-2146826246"` or similar (the boxed enum's `ToString`), surfacing garbage as "Excel version" in the startup trace | 0.75 |
| V005 | High | Marshaling.cs:124-142 (`ToStringSafe` default branch) | NullReferenceException risk / contract violation | `Convert.ToString(value, IFormatProvider)` returns `null` if `value.ToString()` returns null; the `?? string.Empty` guards that. But for types that don't implement `IFormatProvider`-aware `IConvertible`, `Convert.ToString` ignores the provider and calls `value.ToString()` directly — culture-sensitive types render in current culture, not invariant | 0.85 |
| V006 | High | Marshaling.cs:104-110 (`TryToDouble` string branch) | DoS / unbounded work | `double.TryParse` on extremely long numeric strings is `O(n)` with no length cap. A malicious 1 MB cell containing `"1" * 1_000_000` is parsed twice (once invariant, once current culture) before returning `false` | 0.70 |
| V007 | Medium | BulkTransfer.cs:120-131 (`WriteBlock(string, string, object[,])`) | Logic bug | The "anchor" range may itself be a multi-cell range; the code uses only `RowFirst`/`ColumnFirst`, silently overwriting cells outside the user-specified anchor and ignoring `RowLast`/`ColumnLast` | 0.85 |
| V008 | Medium | Marshaling.cs:343-354 (`CellEqualityComparer.GetHashCode`) | Hash quality / collision | All blank/error sentinels hash to `GetType().GetHashCode()` — every `ExcelError` value (NA, DIV/0, REF, etc.) collides into one bucket. Hot-loop lookups on error-heavy data degrade to O(n) per probe in `HashSet`/`Dictionary` keyed by `CellEquality` | 0.80 |
| V009 | Medium | Marshaling.cs:73-115 (`TryToDouble`) | Missing type coverage | `uint`, `short`, `ushort`, `byte`, `sbyte`, `ulong` all fall through to `default` and return `false` even though they trivially convert. Excel-DNA never produces these, but `AsArray2D` accepts any `object` from user UDF wrappers and other internal callers | 0.55 |
| V010 | Medium | Marshaling.cs:255-269 (`BoxStringMatrix`) | Shape mismatch crash | No check that `block` is non-null after the null-guard or that nested null cells aren't `string` nulls (they are by signature `string?[,]`). But the boxed result is `object[,]` and the comment says "Null entries become the empty string" — that branch is fine. However, the `BoxDoubleMatrix` sibling at line 234-249 also has no NaN/Infinity guard: a `double.NaN` boxed and sent to `ExcelReference.SetValue` may result in Excel writing the cell as `#NUM!` or `#VALUE!` silently | 0.60 |
| V011 | Medium | BulkTransfer.cs:31-56 (`ResolveRange`) | Injection via sheet name | The replace `'` → `''` only escapes single quotes. A sheet name containing `]` (e.g. workbook qualifier injection like `]Sheet1`) or a leading `[` reaches `xlfEvaluate` unsanitized, potentially resolving to a different workbook than intended | 0.65 |
| V012 | Medium | BulkTransfer.cs:107-111 (`WriteBlock`) | Error message hides root cause | `target.SetValue(block)` returns `bool` with no further information. A failed write (protected sheet, calc-engine state, sheet hidden) all surface as the same generic `InvalidOperationException`. Lossy diagnostic — debuggability cost only | 0.50 |
| V013 | Medium | AddIn.cs:55-60 (`OnUnhandledException`) | Error swallowing | If `exceptionObject` is not an `Exception` (Excel-DNA invokes the handler with the *original return value* when a UDF returns a non-`ExcelError`-recognized object), `ex` is null and the trace logs `"Unhandled UDF exception: "` with no payload, then unconditionally returns `#VALUE!`. The non-exception payload is dropped on the floor | 0.70 |
| V014 | Low | AddIn.cs:44-53 (`AutoClose`) | Ordering invariant | `ToolkitLifetime.Shutdown()` cancels and *disposes nothing* (per ToolkitLifetime.cs:54-67 — only `Cancel` is called, not `Dispose`); the CTS object lives until `Reset()` is called on the next `AutoOpen`. `Reset()` then disposes the old CTS. But if Excel never calls `AutoOpen` again, the cancelled CTS is held until process exit — a one-time leak, not a runtime hazard | 0.55 |
| V015 | Low | Marshaling.cs:131 (`ToStringSafe` ExcelError branch) | Lossy formatting | `ErrorToText` returns a 6–13 char string for known errors and `"#ERR"` for unknown values. A future Excel error sentinel added to the enum (e.g. `ExcelErrorSpill`) would render as `"#ERR"` and be indistinguishable from any other unknown — version-pinning risk | 0.45 |
| V016 | Low | Marshaling.cs:133-136 (`ToStringSafe` double formatting) | Round-trip invariant violation | `double.ToString("R")` round-trips, but `float.ToString("R")` returns a string with no decimal point for whole numbers (e.g. `5f` → `"5"`). `TryToDouble("5")` returns `5.0`, which is a `double`. The round-trip is float → string → double, losing the original type. `CellEquality.Equals(5f, "5")` returns true (both go via `TryToDouble`), so practically consistent | 0.40 |
| V017 | Low | BulkTransfer.cs:202-221 (`ReadRangeUdf`) | Information leak in trace | The exception message is logged verbatim. If a sheet name was injected with credential-looking characters, the trace contains it. Defense-in-depth, not a real risk in normal workflows | 0.35 |

---

## Detailed entries

### V001 | Critical | Marshaling.cs:104-110 | Data corruption — culture-confused decimal parsing

**Concrete failure scenario.** A worksheet built in Germany contains the cell value `"1,5"` (one and a half, German decimal). Code path:

1. `TryToDouble("1,5", out result)` enters the `string s` case at Marshaling.cs:104.
2. Line 105: `double.TryParse("1,5", NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out result)`.
3. **InvariantCulture's `NumberFormat.NumberGroupSeparator` is `","`** (verified — per .NET docs `NumberFormatInfo.InvariantInfo.NumberGroupSeparator == ","`).
4. With `NumberStyles.AllowThousands` set, `","` is parsed as a *group separator*, not as a decimal point. `"1,5"` is interpreted as `1 group-sep 5` which `double.TryParse` accepts as `15.0`.
5. The method returns `true` with `result = 15.0`. The second `TryParse` (current culture) is never reached.

**Evidence (code trace).**
```
Marshaling.cs:104    case string s:
Marshaling.cs:105        return double.TryParse(
Marshaling.cs:106            s,
Marshaling.cs:107            NumberStyles.Float | NumberStyles.AllowThousands,
Marshaling.cs:108            CultureInfo.InvariantCulture,
Marshaling.cs:109            out result)
Marshaling.cs:110            || double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out result);
```

This silently 10x's any German/French/Spanish decimal input that uses comma as the decimal separator and happens to have a single comma. `"1,5"` → `15.0`. `"1,75"` → `175.0`. `"123,4"` → `1234.0`.

**Proposed surgical fix.** Drop `AllowThousands` for the invariant pass (numeric strings from Excel cells are not thousands-formatted in invariant form), or use `NumberStyles.Float` (no `AllowThousands`) for both passes:
```csharp
return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out result)
    || double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out result);
```

**Confidence: 0.95.** This is reproducible with a single `double.TryParse` call. Cells reaching `TryToDouble` as strings (rather than as `double`) are the exact case where users typed text — i.e., locale-formatted numbers.

---

### V002 | High | Marshaling.cs:101-103 | Unhandled `OverflowException` from `DateTime.ToOADate`

**Concrete failure scenario.** A caller (e.g. `ToDoubleMatrix` at Marshaling.cs:190-204) passes a block containing `new DateTime(50, 1, 1)` (year 50). Code path:

1. `TryToDouble` hits `case DateTime dt` at line 101.
2. `dt.ToOADate()` throws `OverflowException` because OLE Automation dates discontinue before midnight, December 30, 1899; per the .NET reference source `DateTime.TicksToOADate` throws `OverflowException` when `value < OADateMinAsTicks` (~year 100) and `value != 0`.
3. The exception propagates out of `TryToDouble`, bubbling through `ToDoubleMatrix`'s row/col loop and aborting the entire bulk conversion mid-block.

`DateTime.MinValue` is the explicit special case (ticks == 0 → returns 0.0), but any `DateTime(year, m, d)` with `1 <= year <= 99` throws.

**Evidence.**
```
Marshaling.cs:101    case DateTime dt:
Marshaling.cs:102        result = dt.ToOADate();
Marshaling.cs:103        return true;
```

`TryToDouble`'s entire contract is "best-effort … unparseable values produce `false`". An exception escape is a contract violation.

**Proposed surgical fix.**
```csharp
case DateTime dt:
    try { result = dt.ToOADate(); return true; }
    catch (OverflowException) { result = 0d; return false; }
```

**Confidence: 0.90.** BCL contract: `DateTime.ToOADate` throws `OverflowException` for dates before year 100 (excluding MinValue) and never for dates after.

---

### V003 | High | BulkTransfer.cs:46-55 | `XlCallException` escapes `ResolveRange`

**Concrete failure scenario.** A UDF calls `BulkTransfer.ReadBlock("Sheet1", "A1:Z100")`. The workbook is closed or Excel is mid-calculation in a state where `xlfEvaluate` is uncallable. Code path:

1. Line 46: `XlCall.Excel(XlCall.xlfEvaluate, qualified)`.
2. Excel-DNA's `XlCall.Excel` raises `XlCallException` (e.g. `XlCallException.XlReturn.XlReturnInvXlfn` or `XlReturnFailed`). The exception is *not* caught by the `is ExcelReference` / `is ExcelError` branches because the call never returned a value.
3. `XlCallException` escapes `ResolveRange`. Callers of `ResolveRange` (`ReadBlock(string,string)`, the multi-arg `WriteBlock` overload at BulkTransfer.cs:120-131, and `RoundTripTransform(string,string,…)` at BulkTransfer.cs:156-159) all advertise that they throw `ArgumentException` on bad input; they actually throw `XlCallException`.

`ReadRangeUdf` at line 212-220 catches `Exception` broadly so this is masked from the worksheet path, but every C# caller observes a different exception type from what the docstring implies.

**Evidence.**
```
BulkTransfer.cs:46        var token = XlCall.Excel(XlCall.xlfEvaluate, qualified);
BulkTransfer.cs:47        if (token is ExcelReference r)
BulkTransfer.cs:48        {
BulkTransfer.cs:49            return r;
BulkTransfer.cs:50        }
BulkTransfer.cs:51        if (token is ExcelError err)
```

**Proposed surgical fix.** Wrap the `XlCall.Excel` in try/catch and translate:
```csharp
object token;
try { token = XlCall.Excel(XlCall.xlfEvaluate, qualified); }
catch (XlCallException ex)
{
    throw new ArgumentException($"Range '{qualified}' could not be evaluated: {ex.Message}.", nameof(a1Address), ex);
}
```

**Confidence: 0.85.** Excel-DNA's `XlCall.Excel` is documented to throw `XlCallException` for any non-`XlReturnSuccess` return value, regardless of input validity.

---

### V004 | High | AddIn.cs:62-73 | `SafeGetExcelVersion` mis-renders `ExcelError`

**Concrete failure scenario.** `AutoOpen` runs early in the boot sequence. `xlfGetWorkspace` returns `ExcelError.ExcelErrorNA` (boxed as `object`) instead of throwing — Excel-DNA boxes the error result rather than raising `XlCallException` when the call succeeds at the C-API level but returns an error sentinel.

1. Line 66: `v = XlCall.Excel(XlCall.xlfGetWorkspace, 2)` returns `(object)ExcelError.ExcelErrorNA`.
2. Line 67: `v?.ToString()` calls `ExcelError.ExcelErrorNA.ToString()` which is the enum's default `ToString` — returns `"ExcelErrorNA"` (the enum name) or the underlying int representation depending on the boxed type.
3. The trace line logs `"ExcelVersion=ExcelErrorNA"` — confusing operational signal; admins reading logs see "version: ExcelErrorNA" and think Excel is broken.
4. The `catch (XlCallException)` is never entered.

**Evidence.**
```
AddIn.cs:64        try
AddIn.cs:65        {
AddIn.cs:66            var v = XlCall.Excel(XlCall.xlfGetWorkspace, 2);
AddIn.cs:67            return v?.ToString() ?? "unknown";
AddIn.cs:68        }
AddIn.cs:69        catch (XlCallException)
```

**Proposed surgical fix.**
```csharp
var v = XlCall.Excel(XlCall.xlfGetWorkspace, 2);
if (v is ExcelError) return "unknown";
return v?.ToString() ?? "unknown";
```

**Confidence: 0.75.** Excel-DNA's behavior: a C-API call returning an `xltypeErr` value bubbles up as an `ExcelError` boxed in `object`, not as a thrown `XlCallException`. The `XlCallException` path is reserved for `xlretFailed`, `xlretInvXlfn`, etc.

---

### V005 | High | Marshaling.cs:124-142 | `Convert.ToString(value, provider)` ignores provider for non-`IConvertible` types

**Concrete failure scenario.** A caller invokes `ToStringSafe(new SomeStruct { ... })` where `SomeStruct` is a custom type that overrides `ToString()` with a culture-sensitive format and does *not* implement `IConvertible`. Code path:

1. The switch falls through to `_ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty` at line 140.
2. **BCL contract:** `Convert.ToString(object, IFormatProvider)` checks `value is IConvertible`. If yes, calls `IConvertible.ToString(provider)`. If no, calls `value.ToString()` — *ignoring the provider entirely*. (Per .NET 8 reference source `Convert.ToString(object, IFormatProvider)`.)
3. For a `Vector3` (no IConvertible), the result is `value.ToString()` in current culture, not invariant. A `Vector3(1.5f, 2.5f, 3.5f)` on a `de-DE` machine renders `"<1,5 2,5 3,5>"` even though invariant was requested.

**Evidence.**
```
Marshaling.cs:140    _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty,
```

**Proposed surgical fix.**
```csharp
_ => value is IConvertible c
       ? c.ToString(CultureInfo.InvariantCulture)
       : value.ToString() ?? string.Empty,
```
…or accept the limitation and document it. The bigger systemic bug is that `Convert.ToString`'s signature is misleading — many engineers (the original author included) assume it always respects the provider.

**Confidence: 0.85.** This is a literal `Convert.ToString` BCL contract gotcha verifiable from the reference source.

---

### V006 | High | Marshaling.cs:104-110 | Unbounded-length string parse

**Concrete failure scenario.** A workbook cell contains a 4 MB string. Code path:

1. `TryToDouble(s)` hits the string branch.
2. `double.TryParse(s, …, InvariantCulture, out _)` runs in O(n) over the string length. Internally it allocates and walks character by character.
3. If the invariant parse fails (it will for a 4 MB blob), `double.TryParse` runs *again* on the same 4 MB string against current culture.
4. Two full O(n) scans, ~8 MB of work per cell. Multiply by a 100k-cell block in `ToDoubleMatrix` — gigabytes of parser work and pinned scan allocations during a single `ReadDoubleBlock`.

This isn't a remote attack vector in normal Excel use, but a workbook with one large cell becomes a DoS for any code path that hits `TryToDouble`.

**Evidence.** Lines 104-110, no length cap.

**Proposed surgical fix.** Reject strings over a sane threshold up front:
```csharp
case string s:
    if (s.Length > 64) return (result = 0d) != 0d; // unreachable; just bail
    // …existing parses
```
or `s.Length <= 64` guard.

**Confidence: 0.70.** Real exposure depends on whether large strings reach this path; in a CSV-importing workbook they easily can.

---

### V007 | Medium | BulkTransfer.cs:120-131 | Multi-cell anchor silently overwrites

**Concrete failure scenario.** Caller invokes `WriteBlock("Sheet1", "A1:B2", a3x3Block)`:

1. Line 123: `anchor = ResolveRange("Sheet1", "A1:B2")` — succeeds, anchor covers a 2×2 range.
2. Line 124-129: a new `ExcelReference` is built using only `anchor.RowFirst`/`ColumnFirst` (top-left of the 2×2) plus `block.GetLength(0)-1`/`block.GetLength(1)-1`. The target is now a 3×3 range starting at A1, ignoring that the user said "A1:B2".
3. `WriteBlock(target, block)` at line 130 happily writes a 3×3 block. Cells A1:C3 are overwritten — including C3 which was *outside* the user's anchor.

Anywhere the caller meant the A1 string as a bounding range, they get silent expansion. This is a "happy path" overrun: no error, just wrong cells stomped.

**Evidence.**
```
BulkTransfer.cs:120    public static void WriteBlock(string sheetName, string anchorA1, object[,] block)
BulkTransfer.cs:121    {
…
BulkTransfer.cs:123        var anchor = ResolveRange(sheetName, anchorA1);
BulkTransfer.cs:124        var target = new ExcelReference(
BulkTransfer.cs:125            anchor.RowFirst,
BulkTransfer.cs:126            anchor.RowFirst + block.GetLength(0) - 1,
BulkTransfer.cs:127            anchor.ColumnFirst,
BulkTransfer.cs:128            anchor.ColumnFirst + block.GetLength(1) - 1,
BulkTransfer.cs:129            anchor.SheetId);
```

**Proposed surgical fix.** If the anchor is multi-cell, validate dimensions match the block, or document and require an explicit single-cell anchor:
```csharp
if ((anchor.RowLast != anchor.RowFirst || anchor.ColumnLast != anchor.ColumnFirst)
    && (anchor.RowLast - anchor.RowFirst + 1 != block.GetLength(0)
        || anchor.ColumnLast - anchor.ColumnFirst + 1 != block.GetLength(1)))
{
    throw new ArgumentException("Anchor range size does not match block; pass a single cell anchor or matching range.");
}
```

**Confidence: 0.85.** The code reads `RowFirst`/`ColumnFirst` only; if the user passes a multi-cell A1, the additional cells are silently lost.

---

### V008 | Medium | Marshaling.cs:343-354 | All errors hash to the same bucket

**Concrete failure scenario.** A caller builds `new HashSet<object?>(Marshaling.CellEquality)` and inserts a column with 100k cells, of which 90% are `ExcelError.ExcelErrorNA` and 10% are `ExcelError.ExcelErrorDiv0`. Code path:

1. `GetHashCode` at line 345-348: `if (IsBlankOrError(obj)) return obj?.GetType().GetHashCode() ?? 0;`.
2. **Both** `ExcelErrorNA` and `ExcelErrorDiv0` have `obj.GetType() == typeof(ExcelError)` — same hash.
3. `Equals` at line 326-335 distinguishes them (line 330-333: `if (x is ExcelError ex && y is ExcelError ey) return ex == ey;`), so they aren't deduped — they all land in the same bucket and probe linearly.

Result: a hash set keyed by `CellEquality` collapses all error variants into one bucket — Set/Dict lookups on error-heavy data degrade to O(n) probes per access, defeating the data structure.

**Evidence.**
```
Marshaling.cs:343    public int GetHashCode(object? obj)
Marshaling.cs:344    {
Marshaling.cs:345        if (IsBlankOrError(obj))
Marshaling.cs:346        {
Marshaling.cs:347            return obj?.GetType().GetHashCode() ?? 0;
Marshaling.cs:348        }
```

**Proposed surgical fix.**
```csharp
if (IsBlankOrError(obj))
{
    if (obj is ExcelError err) return HashCode.Combine(typeof(ExcelError), (int)err);
    return obj?.GetType().GetHashCode() ?? 0;
}
```

**Confidence: 0.80.** Confirmed by tracing the comparer methods. Practical impact depends on whether errors dominate the keyed data.

---

### V009 | Medium | Marshaling.cs:73-115 | Missing integer type coverage in `TryToDouble`

**Concrete failure scenario.** A future internal caller assembles an `object[,]` from a typed source (e.g. a `byte[]` blob read from `DirectFileIO`) and passes it through `ToDoubleMatrix`. Cells of type `byte`, `short`, `ushort`, `uint`, `ulong`, `sbyte` all fall through to the `default` branch at Marshaling.cs:111-113 and return `false`, producing `defaultValue` (typically `0d`) — *silently wrong* numeric output.

Excel-DNA's wire types are `double`/`string`/`bool`/`ExcelError`/`ExcelEmpty`/`ExcelMissing`/`object[,]`, so a *direct* read from Excel won't hit this. But `BulkTransfer.RoundTripTransform` accepts any `Func<object[,], object[,]>` and the transform's output may contain other numeric types.

**Evidence.** Lines 73-115 enumerate `double, int, long, float, decimal, bool, DateTime, string` and no other integer widths.

**Proposed surgical fix.** Add explicit cases for `byte`/`sbyte`/`short`/`ushort`/`uint`/`ulong`, or use `IConvertible.ToDouble(InvariantCulture)` as a final fallback before `default`.

**Confidence: 0.55.** The exposure is real but requires a caller to construct non-Excel-native numeric objects.

---

### V010 | Medium | Marshaling.cs:234-249 | `BoxDoubleMatrix` doesn't sanitize NaN/Infinity

**Concrete failure scenario.** `VectorizedKernels` returns a `double[,]` containing `double.NaN` (e.g. a divide-by-zero in a kernel). The caller does `WriteDoubleBlock(target, result)` at BulkTransfer.cs:189-192, which calls `Marshaling.BoxDoubleMatrix` then `WriteBlock`. Code path:

1. `BoxDoubleMatrix` at line 234-249 boxes each cell unchanged. NaN/±Infinity are preserved.
2. `target.SetValue(boxed)` sends the boxed NaN to Excel via COM. Excel's behavior for a NaN double in `xltypeNum` is implementation-defined: in current Excel builds the cell becomes `#NUM!`, in others the IEEE bit pattern is stored verbatim (`-1.#IND`) and silently displays as a tiny value.

Either way: undefined surface behavior at the Excel boundary, with no diagnostic on the C# side.

**Evidence.** Lines 234-249, no IEEE checks.

**Proposed surgical fix.** Map NaN/Infinity to `ExcelError.ExcelErrorNum` on the boundary:
```csharp
var v = block[r, c];
result[r, c] = (double.IsNaN(v) || double.IsInfinity(v))
    ? (object)ExcelError.ExcelErrorNum
    : v;
```

**Confidence: 0.60.** The Excel side is not deterministic across versions; on modern Excel, NaN does render as `#NUM!` cleanly, so this is more of a hardening finding than an outright bug.

---

### V011 | Medium | BulkTransfer.cs:42-44 | Sheet name escaping doesn't handle bracket injection

**Concrete failure scenario.** Caller invokes `ResolveRange("[Other.xlsx]Sheet1", "A1")`:

1. Line 42-44 builds `qualified = "'[Other.xlsx]Sheet1'!A1"` (the brackets are not escaped — the `Replace` only handles single quotes).
2. Line 46: `XlCall.Excel(XlCall.xlfEvaluate, "'[Other.xlsx]Sheet1'!A1")`. Excel's expression evaluator resolves this to *another workbook* if `Other.xlsx` is open, reading from a workbook other than the caller intended.

Real Excel sheet names disallow `[`, `]`, `?`, `*`, `:`, `/`, `\`, but the API takes a `string` and the function trusts it. A user-supplied `sheetName` (e.g. from a UDF argument) can therefore exfiltrate or pollute data from other open workbooks.

**Evidence.**
```
BulkTransfer.cs:42        var qualified = a1Address.Contains('!', StringComparison.Ordinal)
BulkTransfer.cs:43            ? a1Address
BulkTransfer.cs:44            : string.Concat("'", sheetName.Replace("'", "''", StringComparison.Ordinal), "'!", a1Address);
```

**Proposed surgical fix.** Validate `sheetName` against Excel's allowed set before quoting:
```csharp
if (sheetName.IndexOfAny(new[] { '[', ']', '*', '?', ':', '/', '\\' }) >= 0)
{
    throw new ArgumentException("Sheet name contains illegal characters.", nameof(sheetName));
}
```

**Confidence: 0.65.** The path is reachable from any caller, including the worksheet-exposed `ReadRangeUdf`.

---

### V012 | Medium | BulkTransfer.cs:107-111 | Lossy `SetValue` failure reporting

**Concrete failure scenario.** Sheet is protected. `target.SetValue(block)` returns `false`. Code path:

1. Line 108-111 throws `InvalidOperationException("Excel rejected the bulk write. The range may be protected or in an invalid state.")`.
2. No information about *why*, *which sheet*, or what cells. Diagnostic value at the bottom of the funnel is poor.

Not a correctness bug — a diagnostic / debuggability ding. Including it because production triage needs more context.

**Evidence.** Lines 107-111.

**Proposed surgical fix.** Include the target dimensions and sheet ID in the exception:
```csharp
throw new InvalidOperationException(
    $"Excel rejected bulk write to sheet {target.SheetId} range {target.RowFirst}-{target.RowLast}, {target.ColumnFirst}-{target.ColumnLast} ({refRows}x{refCols}).");
```

**Confidence: 0.50.** Quality-of-error issue only.

---

### V013 | High | AddIn.cs:55-60 | `OnUnhandledException` drops non-`Exception` payloads

**Concrete failure scenario.** Excel-DNA's `RegisterUnhandledExceptionHandler` documentation states the handler is invoked with "the Exception or the return value that caused the failure". A UDF that returns an object of a type Excel-DNA cannot marshal (e.g. an arbitrary CLR object that isn't a recognized wire type) calls the handler with *that object*, not with an `Exception`. Code path:

1. Line 57: `var ex = exceptionObject as Exception;` — `ex` is `null`.
2. Line 58: `TraceSource.TraceEvent(TraceEventType.Error, 0, "Unhandled UDF exception: {0}", ex);` — formats with null. Output: `"Unhandled UDF exception: "` (empty).
3. The actual offending payload (`exceptionObject`) is never logged.
4. Returns `#VALUE!` regardless.

Operationally this means UDFs returning bad types produce a silent `#VALUE!` with no signal in the trace about what the offending return value was.

**Evidence.**
```
AddIn.cs:55    private static object OnUnhandledException(object exceptionObject)
AddIn.cs:56    {
AddIn.cs:57        var ex = exceptionObject as Exception;
AddIn.cs:58        TraceSource.TraceEvent(TraceEventType.Error, 0, "Unhandled UDF exception: {0}", ex);
AddIn.cs:59        return ExcelError.ExcelErrorValue;
AddIn.cs:60    }
```

**Proposed surgical fix.**
```csharp
if (exceptionObject is Exception ex)
    TraceSource.TraceEvent(TraceEventType.Error, 0, "Unhandled UDF exception: {0}", ex);
else
    TraceSource.TraceEvent(TraceEventType.Error, 0, "Unhandled UDF payload: {0} (type={1})",
        exceptionObject, exceptionObject?.GetType().FullName ?? "null");
```

**Confidence: 0.70.** Excel-DNA's handler contract is documented as `Func<object,object>` and the input is not constrained to `Exception`.

---

### V014 | Low | AddIn.cs:44-53 | `AutoClose` ordering and CTS-disposal interaction

**Concrete failure scenario.** Walking the explicit shutdown order with the helper class:

1. Line 50: `ToolkitLifetime.Shutdown()` → ToolkitLifetime.cs:54-67 — calls `_cts.Cancel()`. The CTS is *not* disposed here (only `Cancel`).
2. Line 51: `FeedManager.Instance.Shutdown()` — iterates feeds and calls `f.Stop()`. The Feed code at RtdServer.cs:333 reads `ToolkitLifetime.ShutdownToken` while creating linked sources; after step 1 that token is already cancelled (consistent — the linked source fires immediately).

So in this ordering there is no use-after-dispose: `Shutdown()` doesn't dispose, only `Reset()` does (ToolkitLifetime.cs:40-47). The window the round-1 hint was after — "what if Shutdown disposed in some path" — doesn't exist; the code is consistent.

What *does* happen is that if Excel never calls `AutoOpen` again, the cancelled-but-undisposed CTS leaks until process exit. One CTS, ~hundreds of bytes — not a runtime hazard.

**Evidence.** AddIn.cs:50-51 + ToolkitLifetime.cs:40-67.

**Proposed surgical fix.** None required; one-line note in `Shutdown()` that it does not dispose, or `Dispose` the source at end of `AutoClose` after `FeedManager` has stopped:
```csharp
ToolkitLifetime.Shutdown();
FeedManager.Instance.Shutdown();
// optional: ToolkitLifetime.DisposeIfShutdown();
```

**Confidence: 0.55.** Not a bug, a tidiness note.

---

### V015 | Low | Marshaling.cs:131, 147-158 | Future-enum collapse in `ErrorToText`

**Concrete failure scenario.** A newer Excel-DNA release adds `ExcelError.ExcelErrorSpill` (the `#SPILL!` value that exists in dynamic-array Excel). The current switch at lines 147-158 maps every unknown enum value to `"#ERR"`. After upgrade, all `#SPILL!` errors render as `"#ERR"` in any code path that goes through `ToStringSafe` (e.g., string-block writes back to Excel for diagnostics or CSV export via `DirectFileIO`).

**Evidence.**
```
Marshaling.cs:147    public static string ErrorToText(ExcelError err) => err switch
...
Marshaling.cs:157        _ => "#ERR",
```

**Proposed surgical fix.** Fall back to `err.ToString()` (the enum name) instead of `"#ERR"` for unknown values, so triage retains some signal:
```csharp
_ => err.ToString(),
```

**Confidence: 0.45.** Forward-compat only.

---

### V016 | Low | Marshaling.cs:133-136 | Round-trip type narrowing

**Concrete failure scenario.** A `float 5.5f` flows through `ToStringSafe`:

1. Line 136: `f.ToString("R", InvariantCulture)` → `"5.5"`.
2. Round-trip: `TryToDouble("5.5")` succeeds (invariant), returns `5.5d`.
3. The result is `double`, not `float`.

`CellEquality` masks this because both sides go through `TryToDouble`. But any caller that round-trips `ToStringSafe` → user input → re-parsed will discover the type changed silently. Practically benign; pure compatibility note.

**Evidence.** Lines 133-136.

**Proposed surgical fix.** None practical — Excel's wire format uses double anyway. Document that round-trip is value-preserving but not type-preserving.

**Confidence: 0.40.**

---

### V017 | Low | BulkTransfer.cs:202-221 | Trace log includes raw input on failure

**Concrete failure scenario.** A worksheet UDF passes a string-typed argument that gets logged on failure:

```
BulkTransfer.cs:218        TraceSource.TraceEvent(TraceEventType.Warning, 1, "EPT.READRANGE failed for {0}!{1}: {2}", sheetName, a1Address, ex.Message);
```

If `sheetName` was attacker-controlled and contained log-injection sequences (newlines, CRLF) the trace listener could be tricked into splitting an entry. Excel-DNA UDF inputs are scrubbed by Excel and unlikely to contain CRLF, but defense-in-depth would sanitize.

**Confidence: 0.35.** Not a real exploit; low-impact polish.

---

## Rejected findings

Each is a concern I considered, tried to construct a concrete failure for, and could not. Listed so synthesis doesn't re-walk them.

- **`AsArray2D` for `string` value (Marshaling.cs:28-38).** A `string s` parameter creates a 1×1 `object[,]` with `[0,0] = s`. No issue — `object[,]` is correctly typed; downstream code accepts string cells. Same for `ExcelError.ExcelErrorRef` — wraps cleanly. Rejected.

- **`AsArray2D` returns the original array reference, not a copy.** If a caller mutates the returned array, they mutate Excel-DNA's internal block. This is by design (zero-copy hot path) and documented as "0 boundary crossings". Rejected as not a bug.

- **`IsBlankOrError` conflates blank and error (round 1 B006).** This is a public, documented API with name "blank *or* error". Callers distinguish via the sibling `IsExcelError`. Round 1's "contract ambiguity" was a documentation concern, not a defect. Rejected.

- **`decimal.MaxValue → double` overflow (round 1 B004).** `(decimal)7.92e28` cast to `double` produces `7.92e28d`, not infinity. `double.MaxValue ≈ 1.79e308` >> `decimal.MaxValue ≈ 7.92e28`. The cast loses precision but does *not* produce Infinity. Round 1 was wrong; rejected.

- **`CellEquality.Equals(int 5, double 5.0)` consistency.** Both go through `TryToDouble`, both produce `5.0d`. `GetHashCode` on both also produces `(5.0d).GetHashCode()`. Consistent. Rejected.

- **`CellEquality.Equals(ExcelEmpty.Value, null)` symmetric/transitive.** Both are `IsBlankOrError`. The branch at line 330-333 only fires when *both* are `ExcelError`; otherwise falls through to `x?.GetType() == y?.GetType()`. `typeof(ExcelEmpty) != null` for `null`, so `Equals(null, ExcelEmpty.Value)` returns `null == typeof(ExcelEmpty)` → false. `Equals(ExcelEmpty.Value, ExcelEmpty.Value)` returns true. Hash: same type, same hash. Consistent. Rejected.

- **`CellEqualityComparer` NaN inconsistency (round 1 B008).** `TryToDouble` rejects NaN at line 85 (`return !double.IsNaN(d) && !double.IsInfinity(d)`), so the comparer's double-path never sees NaN — it falls through to `ToStringSafe` and compares strings. Round 1 was self-rejecting; confirmed rejected here.

- **`ReadBlock` for a 3-D / multi-sheet `ExcelReference`.** Excel-DNA's `ExcelReference.GetValue()` for a multi-sheet reference returns the value of the first inner sheet only (per Excel-DNA source — `GetValue` calls `xlCoerce` which is single-sheet). No silent multi-sheet flattening, no corruption — caller gets first sheet, which is the C-API semantics. Not a defect *in this domain*; document at most. Rejected.

- **`WriteBlock` with `block.GetLength(0) == 0`.** `ExcelReference` cannot validly refer to a zero-row range — `RowLast >= RowFirst` is enforced by Excel before `ExcelReference` is constructable through `xlfEvaluate`. The shape mismatch check at line 100 fires correctly: `0 != (RowLast - RowFirst + 1) ≥ 1`. The thrown `ArgumentException` is the right behavior. Rejected.

- **`BoxDoubleMatrix` cache behavior (round 2 prompt item 15).** Row-major iteration on a row-major-stored `double[,]` is cache-optimal. Confirmed no issue. Rejected.

- **`ReadRangeUdf` catching `Exception`.** The catch is broad but explicitly logs; returning `ExcelErrorRef` is the correct UDF contract. Not error swallowing in the bad sense — the trace at line 218 captures the message. Rejected.

- **`ResolveRange` `'O'Brien` sheet name escaping.** `sheetName.Replace("'", "''")` turns `O'Brien` into `O''Brien`; wrapped in `'…'` it becomes `'O''Brien'` — that *is* the correct Excel escape for a sheet name with a single quote. Verified against Excel's reference syntax. Rejected.

- **`SafeGetExcelVersion` exception scope (round 2 prompt item 14).** The catch is `XlCallException` only. Any other exception type (`COMException` from a torn-down host?) escapes `AutoOpen` and aborts add-in load. Excel-DNA's `XlCall.Excel` is documented to throw `XlCallException` and not other exception types under normal operation. The `ExcelError` boxed-return case is filed as V004; non-`XlCallException` exceptions from `XlCall.Excel` are not a documented failure mode. Rejected as not concrete enough.

- **Cross-domain: `FeedManager.Shutdown` called twice (round 1 B005).** One-line note: yes, `RtdServer.ServerTerminate` and `AddIn.AutoClose` both call it; the inner `lock(_gate)` + `_feeds.Clear()` makes it idempotent. Outside this domain. Noted, not filed.

- **Cross-domain: `Feed.Subscribe`/`Unsubscribe` race (round 1 B001/B002).** Outside this domain. Noted, not filed.

- **Cross-domain: `ToolkitLifetime.Reset` dispose-then-reassign (round 1 B003).** Outside this domain. Noted, not filed.

- **Cross-domain: `RtdServer.FlushTick` exception scope (round 1 B007).** Outside this domain. Noted, not filed.
