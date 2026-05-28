# Production-Grade Audit: DeveloperUtilities.cs

**Domain:** Developer Utilities  
**File:** `/home/user/plugin/src/ExcelPerfToolkit/DeveloperUtilities.cs` (914 lines, 17 UDFs)  
**Audit Date:** 2026-05-28  
**Scope:** Full file under extreme load, partial failures, high concurrency, edge cases, and shutdown safety.

---

## Findings Summary

| Count | Severity |
|-------|----------|
| 1     | Critical |
| 3     | High     |
| 4     | Medium   |
| 2     | Low      |

---

## Findings by ID

### 1 | Critical | FlattenIntIndexes:729 | Integer Overflow (Unchecked Cast) | cast of double to int without bounds validation
**Location:** `/home/user/plugin/src/ExcelPerfToolkit/DeveloperUtilities.cs:729`

**Category:** Integer Overflow / Type Safety Violation

**Concrete Failure Scenario:**
A user passes a lookup return-column index block containing the double value `2.147483648e9` (or any value > `Int32.MaxValue` = 2,147,483,647). The cast `(int)d` silently wraps or truncates to a negative number or small positive integer, e.g., `2147483648 → -2147483648` in unchecked context. This causes:
1. `SortBlock` line 762 to call `FlattenIntIndexes` with a block containing `2.5e9`
2. Index cast truncates to unexpected value
3. An array index out-of-bounds exception is raised in `CompareRows:815` when accessing `block[a, col]`, causing a runtime crash instead of a validation error
4. OR the cast wraps to a negative index, which also crashes in array access

**Evidence:**
```csharp
// Line 725-729: TryToDouble succeeds for any finite double
if (!Marshaling.TryToDouble(v, out var d))  // d could be 2e9+
{
    throw new ArgumentException("...");
}
list.Add((int)d);  // UNCHECKED: silently truncates/wraps if d > 2147483647
```

The contract in Marshaling.TryToDouble (line 85 in Marshaling.cs) returns `true` for any finite double, including very large values. No upper-bound validation before cast.

**Proposed Surgical Fix:**
Replace line 729 with explicit bounds checking:
```csharp
if (d < 0d || d > int.MaxValue)
{
    throw new ArgumentException($"Column index {d} out of range [0, {int.MaxValue}].", nameof(d));
}
list.Add((int)d);
```

Alternatively (stricter):
```csharp
list.Add(checked { (int)d });  // Throws OverflowException at runtime
```

**Confidence:** 0.95

---

### 2 | High | Unpivot:502 | Integer Overflow (Unchecked Multiplication) | dataRows * valueColumns could exceed Int32.MaxValue
**Location:** `/home/user/plugin/src/ExcelPerfToolkit/DeveloperUtilities.cs:502`

**Category:** Integer Overflow / Memory Allocation Failure

**Concrete Failure Scenario:**
A user passes a 50,000-row block with 50,000 columns and `keyColumns=0`. Then:
- `dataRows = 50000 - 1 = 49999`
- `valueColumns = 50000 - 0 = 50000`
- `outRows = 49999 * 50000 = 2,499,950,000` (exceeds `Int32.MaxValue`)

In unchecked context, the multiplication wraps to a negative number or small positive (modulo 2^32), e.g., `2499950000 % 2^32 ≈ 405482704`. The subsequent array allocation at line 504:
```csharp
var result = new object[outRows, outCols];
```
silently allocates a 405M-row array instead of the 2.5B-row array intended. The loop at line 507-519 then **writes beyond array bounds**, corrupting memory or causing an IndexOutOfRangeException on the first out-of-bounds write.

**Evidence:**
```csharp
// Line 500-504: No overflow check
var dataRows = rows - 1;           // = 49999
var valueColumns = cols - keyColumns; // = 50000
var outRows = dataRows * valueColumns; // UNCHECKED: 2499950000 wraps to ~405M
var outCols = keyColumns + 2;
var result = new object[outRows, outCols]; // Allocates wrong size
```

Array bounds are checked at allocation but the **multiply result itself** is not validated.

**Proposed Surgical Fix:**
Replace lines 502-504 with overflow checking:
```csharp
// Check for overflow before allocation
checked
{
    outRows = dataRows * valueColumns;
}
var outCols = keyColumns + 2;
var result = new object[outRows, outCols];
```

Or (defensive, with graceful error):
```csharp
if (dataRows > 0 && valueColumns > int.MaxValue / dataRows)
{
    throw new ArgumentException("Unpivot output exceeds maximum array size.", nameof(block));
}
var outRows = dataRows * valueColumns;
```

**Confidence:** 0.98

---

### 3 | High | StackColumns:466 | Integer Overflow (Unchecked Multiplication) | rows * cols could exceed Int32.MaxValue
**Location:** `/home/user/plugin/src/ExcelPerfToolkit/DeveloperUtilities.cs:466`

**Category:** Integer Overflow / Memory Allocation Failure

**Concrete Failure Scenario:**
A user passes a 100,000 row × 100,000 column block. Then:
- `rows * cols = 10,000,000,000` (exceeds `Int32.MaxValue`)
- The multiplication wraps in unchecked context to ~1.7B (modulo 2^32)
- Array allocation creates a ~1.7B-row array instead of the intended 10B-row array
- The loop at lines 468-473 writes beyond bounds, causing memory corruption or IndexOutOfRangeException

**Evidence:**
```csharp
// Line 464-466: No overflow check on rows * cols
var rows = block.GetLength(0);
var cols = block.GetLength(1);
var result = new object[rows * cols, 1]; // UNCHECKED multiply
```

Excel-DNA array dimensions are `int[,]`, so the first dimension of `object[rows * cols, 1]` is a single unchecked multiply.

**Proposed Surgical Fix:**
```csharp
checked
{
    var resultRows = rows * cols;  // Throws OverflowException if overflow
}
var result = new object[resultRows, 1];
```

Or (defensive):
```csharp
if (rows > 0 && cols > int.MaxValue / rows)
{
    throw new ArgumentException("Stack result exceeds maximum array size.", nameof(block));
}
var result = new object[rows * cols, 1];
```

**Confidence:** 0.98

---

### 4 | High | Sha256Block:905-906, 908-909 | Excessive Stack Allocation (EXTREME) | separator.ToArray() and rowSeparator.ToArray() called per cell in inner loops
**Location:** `/home/user/plugin/src/ExcelPerfToolkit/DeveloperUtilities.cs:905-906, 908-909`

**Category:** Performance Pathology / Allocation Hot Path

**Concrete Failure Scenario:**
A user passes a 10,000 × 10,000 block and calls `EPT.SHA256BLOCK`. The code:
1. Allocates a 1-byte `separator` Span on the stack (line 894)
2. In the inner loop (line 862), calls `separator.ToArray()` **100,000,000 times** (10K × 10K), creating 100M temporary byte[1] arrays on the managed heap
3. Similarly, calls `rowSeparator.ToArray()` 10,000 times
4. Total: ~100M GC allocations

This causes:
- Sustained Gen0 → Gen1 → Gen2 promotion pressure
- Pauses in the 100ms+ range during GC (STOP-THE-WORLD on Excel's main thread)
- Excel UI frozen, user perceives hang or crash
- Potential OutOfMemoryException if heap is already stressed

The `stackalloc` was **negated** by immediately converting to heap arrays.

**Evidence:**
```csharp
// Line 856-859: Span allocated once
Span<byte> separator = stackalloc byte[1];
separator[0] = 0x1f;
Span<byte> rowSeparator = stackalloc byte[1];
rowSeparator[0] = 0x1e;

for (var r = 0; r < rows; r++)  // 10K iterations
{
    for (var c = 0; c < cols; c++)  // 10K iterations
    {
        // ...
        var sep = separator.ToArray();  // LINE 905: Heap alloc inside nested loop
        sha.TransformBlock(sep, 0, sep.Length, null, 0);
    }
    var rowSep = rowSeparator.ToArray();  // LINE 908: Heap alloc inside loop
    sha.TransformBlock(rowSep, 0, rowSep.Length, null, 0);
}
```

**Proposed Surgical Fix:**
Replace lines 905-906 and 908-909 to use the Span directly:
```csharp
sha.TransformBlock(separator.AsReadOnlySpan().ToArray(), 0, 1, null, 0);
// OR better:
Span<byte> sepBuffer = stackalloc byte[1] { 0x1f };
sha.TransformBlock(new byte[] { 0x1f }, 0, 1, null, 0);
```

But **the real fix** is to rewrite to avoid repeated allocations:
```csharp
byte[] sepBytes = new byte[] { 0x1f };
byte[] rowSepBytes = new byte[] { 0x1e };

for (var r = 0; r < rows; r++)
{
    for (var c = 0; c < cols; c++)
    {
        // ...
        sha.TransformBlock(sepBytes, 0, 1, null, 0);
    }
    sha.TransformBlock(rowSepBytes, 0, 1, null, 0);
}
```

This allocates exactly 2 byte arrays (once) instead of ~rows*cols + rows times.

**Confidence:** 0.99

---

### 5 | High | HashBlock:865 | Per-Cell String Allocation (Repeated UTF8 Encoding) | Encoding.UTF8.GetBytes called per cell without caching or reuse
**Location:** `/home/user/plugin/src/ExcelPerfToolkit/DeveloperUtilities.cs:865`

**Category:** Performance Pathology / Allocation Hot Path

**Concrete Failure Scenario:**
A user has a 5,000 × 1,000 block (5M cells) of mixed strings and numbers. Calling `EPT.HASHBLOCK`:
1. Per cell, `Marshaling.ToStringSafe(block[r, c])` is called → allocates a string (if not already a string in the cell)
2. `Encoding.UTF8.GetBytes(text)` is called → allocates a byte[] array
3. Total: 5M string conversions + 5M byte[] allocations = **10M GC allocations** (just for hashing)

**Evidence:**
```csharp
for (var r = 0; r < rows; r++)
{
    for (var c = 0; c < cols; c++)
    {
        var text = Marshaling.ToStringSafe(block[r, c]);  // Allocation (conversion)
        var bytes = Encoding.UTF8.GetBytes(text);        // Allocation (UTF8 encode)
        hasher.Append(bytes);
        hasher.Append(separator);
    }
    hasher.Append(rowSeparator);
}
```

Each `GetBytes` call allocates a new byte[] on the heap. With 5M cells, this is a sustained allocation rate that triggers GC overhead.

**Proposed Surgical Fix:**
Use `Encoding.UTF8.GetBytes(text, 0, text.Length, buffer, 0)` overload with a pre-allocated reusable buffer:
```csharp
byte[] buffer = new byte[1024];  // Or estimate max cell size
for (var r = 0; r < rows; r++)
{
    for (var c = 0; c < cols; c++)
    {
        var text = Marshaling.ToStringSafe(block[r, c]);
        int byteCount = Encoding.UTF8.GetByteCount(text);
        if (byteCount > buffer.Length)
        {
            buffer = new byte[byteCount];
        }
        Encoding.UTF8.GetBytes(text, 0, text.Length, buffer, 0);
        hasher.Append(buffer.AsSpan(0, byteCount));
        hasher.Append(separator);
    }
    hasher.Append(rowSeparator);
}
```

This reuses a single buffer across all cells.

**Confidence:** 0.92

---

### 6 | Medium | Sha256Block:904, 905-906, 909 | CRITICAL DESIGN ISSUE: SHA256 API Misuse | TransformBlock with null output buffer allocated/deallocated per cell
**Location:** `/home/user/plugin/src/ExcelPerfToolkit/DeveloperUtilities.cs:904-906, 909`

**Category:** API Misuse / Inefficient Hashing Pattern

**Concrete Failure Scenario:**
The `TransformBlock` API (line 904, 906, 909) is being called with null output buffer:
```csharp
sha.TransformBlock(bytes, 0, bytes.Length, null, 0);
```

While this is technically valid, each call incurs:
1. Parameter validation
2. Internal state copy (for null buffer case)
3. Hash state update

For a 10K × 10K block, this is 100M+10K calls to `TransformBlock`. **Better pattern:** accumulate data in a buffer and call `TransformBlock` less frequently, or use `Stream` abstraction.

Also, on line 905-906, `separator.ToArray()` creates a heap allocation per cell (see Finding #4).

**Evidence:**
```csharp
// INEFFICIENT: Many small TransformBlock calls
sha.TransformBlock(bytes, 0, bytes.Length, null, 0);  // Per cell
sha.TransformBlock(sep, 0, sep.Length, null, 0);      // Per cell
// ...
sha.TransformBlock(rowSep, 0, rowSep.Length, null, 0); // Per row
```

**Proposed Surgical Fix:**
Combine input into a single buffer and call `TransformBlock` once per row:
```csharp
var rowBuffer = new StringBuilder(cols * 64);
for (var r = 0; r < rows; r++)
{
    rowBuffer.Clear();
    for (var c = 0; c < cols; c++)
    {
        var text = Marshaling.ToStringSafe(block[r, c]);
        rowBuffer.Append(text).Append('');  // Separator
    }
    rowBuffer.Append('');  // Row separator
    var rowBytes = Encoding.UTF8.GetBytes(rowBuffer.ToString());
    sha.TransformBlock(rowBytes, 0, rowBytes.Length, null, 0);
}
sha.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
```

(This still has allocation overhead from StringBuilder, but batches API calls.)

**Confidence:** 0.85

---

### 7 | Medium | BuildRowKey:182-187 | Per-Row Key String Allocation (Hash Collision Risk) | Separator character '\0' may not be unique enough
**Location:** `/home/user/plugin/src/ExcelPerfToolkit/DeveloperUtilities.cs:185`

**Category:** Hashing / Collision Risk

**Concrete Failure Scenario:**
The row key is built by concatenating cell strings with a separator character (line 185):
```csharp
sb.Append(Marshaling.ToStringSafe(block[row, c])).Append('');  // Appends literal null char \0
```

Two different rows can produce the same hash key:
- **Row A:** `["hello", "world"]` → `"hello\0world"`
- **Row B:** `["hello\0world", ""]` → `"hello\0world\0"`

Wait—control characters (\0x00–\0x1F) are illegal in Excel cells per ECMA-376, so this **should not occur in practice**. However:
1. If a user somehow injects a null byte through COM interop or malformed XLL data, the separator is defeated
2. The code assumes Marshaling.ToStringSafe() never produces control characters—this is not documented
3. The choice of '\0' (the weakest separator, the C string terminator) is suspicious

**Evidence:**
```csharp
// Line 185: Separator is a literal null char
sb.Append(Marshaling.ToStringSafe(block[row, c])).Append('');
// Produces "cell1\0cell2\0cell3"
// But also "cell1\0cell2\0cell3" if cell1 = "cell1\0cell2" and cell2 is empty
```

The contract in line 16 says "honors the one-crossing rule," implying inputs are Excel cells. Excel forbids control chars, so this is **theory only**. But it's fragile.

**Proposed Surgical Fix:**
Use a stronger, multi-byte separator (e.g., `` which Excel-DNA recognizes):
```csharp
sb.Append(Marshaling.ToStringSafe(block[row, c])).Append('');  // Unit separator
```

Or use explicit byte encoding (like HashBlock does):
```csharp
// Convert entire row to bytes with explicit separators
var rowKey = BuildRowKeyBytes(block, row, cols);  // Returns byte[] hash
// Store in HashSet<string> with hex encoding
```

**Confidence:** 0.65 (low practical risk, but poor defensive design)

---

### 8 | Medium | RemoveDuplicateRows:162 | Per-Row Key Building is Linear but Inefficient | StringBuilder reallocation per row could be optimized
**Location:** `/home/user/plugin/src/ExcelPerfToolkit/DeveloperUtilities.cs:180-188`

**Category:** Performance / Allocation Efficiency

**Concrete Failure Scenario:**
For a 100,000-row × 1,000-column block, `BuildRowKey` is called 100K times (line 162):
```csharp
for (var r = 0; r < rows; r++)
{
    var key = BuildRowKey(block, r, cols);  // Called 100K times
    // ...
}
```

Each call allocates a **new StringBuilder** (line 182) with an estimated capacity of `cols * 8`. For 1K columns, this is 8KB per row. With 100K rows, that's **800MB of StringBuilder buffers allocated and GC'd** (since StringBuilder is short-lived). While this is O(n) not O(n²), the constant factor is high.

**Evidence:**
```csharp
private static string BuildRowKey(object[,] block, int row, int cols)
{
    var sb = new StringBuilder(cols * 8);  // NEW allocation per call
    for (var c = 0; c < cols; c++)
    {
        sb.Append(Marshaling.ToStringSafe(block[row, c])).Append('');
    }
    return sb.ToString();  // String copy allocation
}
```

Called 100K times, this is a hot path.

**Proposed Surgical Fix:**
Reuse a single StringBuilder across all rows (thread-safe since non-IsThreadSafe function):
```csharp
var seen = new HashSet<string>(StringComparer.Ordinal);
var keep = new List<int>(rows);
var keyBuilder = new StringBuilder(cols * 8);  // Single, reused
for (var r = 0; r < rows; r++)
{
    keyBuilder.Clear();
    for (var c = 0; c < cols; c++)
    {
        keyBuilder.Append(Marshaling.ToStringSafe(block[r, c])).Append('');
    }
    var key = keyBuilder.ToString();
    if (seen.Add(key))
    {
        keep.Add(r);
    }
}
```

This reduces allocations by 100K→1 (excluding the strings themselves).

**Confidence:** 0.90

---

### 9 | Medium | SortBlock:797 | Array.Sort Stability Claim vs. Reality | Documentation claims "stable sort" but Array.Sort uses Quicksort (unstable)
**Location:** `/home/user/plugin/src/ExcelPerfToolkit/DeveloperUtilities.cs:742, 797`

**Category:** Documentation Bug / Correctness (Observable Impact)

**Concrete Failure Scenario:**
A user relies on the documented behavior "Sort is stable" (line 742) to preserve original row order for equal keys. They sort a 10-row block by a column with many duplicates:
- Row 0: `[Key=5, Data="A"]`
- Row 1: `[Key=5, Data="B"]`
- Row 2: `[Key=5, Data="C"]`

After sort, they expect:
- Rows with Key=5 in original order: `A, B, C`

But `Array.Sort` (used on line 797) uses **Introsort** (a hybrid of Quicksort and Heapsort) which is **NOT stable**. The order of rows with equal keys is undefined:
- Possible result: `A, C, B` or `B, A, C` etc.

**Evidence:**
```csharp
// Line 742: Documentation claim
/// Returns <paramref name="block"/> sorted by one or more key columns. Sort is stable.

// Line 797: Implementation
var order = Enumerable.Range(0, rows).ToArray();
Array.Sort(order, (a, b) => CompareRows(...));  // Array.Sort is NOT stable
```

From .NET documentation, `Array.Sort<T>(T[], IComparer<T>)` uses Introsort, which is unstable.

**Proposed Surgical Fix:**
Option 1: Use a stable sort implementation (e.g., LINQ OrderBy, which uses merge sort):
```csharp
var order = Enumerable.Range(0, rows)
    .OrderBy(i => i, new RowComparer(block, keys, desc))
    .ToArray();

private sealed class RowComparer : IComparer<int>
{
    // Implement IComparer<int> wrapping CompareRows
}
```

Option 2: Update documentation to say "sort order for equal keys is undefined":
```csharp
/// Sort order for rows with equal keys is undefined (Array.Sort is not stable).
```

**Confidence:** 0.88

---

### 10 | Medium | BlockLookup:678-681 | Silent Duplicate-Key Handling (First-Win Semantics) | Documentation does not clarify behavior with duplicate lookup keys
**Location:** `/home/user/plugin/src/ExcelPerfToolkit/DeveloperUtilities.cs:678-681`

**Category:** Correctness / Undocumented Behavior

**Concrete Failure Scenario:**
A lookup table has duplicate keys in column 0:
```
Row 0: [Key="Alice", Value=100]
Row 1: [Key="Alice", Value=200]  // Duplicate
```

When a left row is matched to "Alice", the code silently returns the **first** occurrence (line 680):
```csharp
if (!index.ContainsKey(key))
{
    index[key] = r;  // Only stores first row for each key
}
```

A user who expects to match to Row 1 (e.g., the most recent or highest-value entry) gets unexpected results silently. The documentation (line 629) does not mention this behavior.

**Evidence:**
```csharp
// Line 674-681: Build lookup index (first-win)
var index = new Dictionary<string, int>(lookupRows, StringComparer.Ordinal);
for (var r = 0; r < lookupRows; r++)
{
    var key = Marshaling.ToStringSafe(lookup[r, 0]);
    if (!index.ContainsKey(key))  // Only first match stored
    {
        index[key] = r;
    }
}
```

The docstring (line 629-634) does not say "first matching row" or "behavior with duplicate keys."

**Proposed Surgical Fix:**
Update documentation to clarify:
```csharp
/// When the lookup table contains duplicate keys (multiple rows with the same value in column 0),
/// the first occurrence is matched. Subsequent duplicates are ignored.
```

Or, if "last" or "all" semantics are desired, modify the index to store a List<int> instead of a single int.

**Confidence:** 0.75

---

### 11 | Low | SplitColumn:327 | String.Split with StringSplitOptions.None (Empty Parts Retained) | May not match user expectation
**Location:** `/home/user/plugin/src/ExcelPerfToolkit/DeveloperUtilities.cs:327`

**Category:** Semantic Mismatch / Edge Case

**Concrete Failure Scenario:**
A user splits `"a||b"` by delimiter `"|"` expecting to drop empty parts. The code calls:
```csharp
parts[r] = s.Split(delimiter, StringSplitOptions.None);
```

Result: `["a", "", "b"]` (with empty string in the middle). User sees empty cells in the output and is confused, expecting `["a", "b"]`.

The documentation (line 299) does not specify behavior of empty parts.

**Evidence:**
```csharp
// Line 327: Split with None (retains empty)
parts[r] = s.Split(delimiter, StringSplitOptions.None);
// "a||b".Split("|", None) → ["a", "", "b"]
```

**Proposed Surgical Fix:**
Either:
1. Update documentation: "Empty parts resulting from consecutive delimiters are retained."
2. Or change to `StringSplitOptions.RemoveEmptyEntries` if that's the intended behavior.

**Confidence:** 0.55 (purely semantic; depends on user intent)

---

### 12 | Low | FillDown:270, 276 | Potential Null Reference (ExcelEmpty.Value Assumption) | Code assumes ExcelEmpty.Value is never null
**Location:** `/home/user/plugin/src/ExcelPerfToolkit/DeveloperUtilities.cs:270, 276`

**Category:** Defensive Programming / Null Safety

**Concrete Failure Scenario:**
The code initializes:
```csharp
object? last = ExcelEmpty.Value;  // Assume non-null
```

Then later (line 276):
```csharp
result[r, c] = last ?? ExcelEmpty.Value;  // Null-coalescing suggests last could be null
```

The null-coalescing operator on line 276 contradicts the assumption on line 270 that `last` is initially non-null. If `ExcelEmpty.Value` is unexpectedly null (e.g., due to a breaking change in ExcelDna), or if a cell contains a true `null`, the `last ?? ExcelEmpty.Value` could produce unexpected results.

**Evidence:**
```csharp
// Line 270: Initialized to ExcelEmpty.Value
object? last = ExcelEmpty.Value;

for (var r = 0; r < rows; r++)
{
    var v = block[r, c];
    if (IsCellBlank(v))
    {
        result[r, c] = last ?? ExcelEmpty.Value;  // LINE 276: null-coalesce contradicts initialization
    }
    else
    {
        result[r, c] = v;
        last = v;  // last could now be anything (or null if v is null)
    }
}
```

**Proposed Surgical Fix:**
Clarify intent:
```csharp
object? last = null;  // Explicitly allow null

for (var r = 0; r < rows; r++)
{
    var v = block[r, c];
    if (IsCellBlank(v))
    {
        result[r, c] = last ?? ExcelEmpty.Value;  // If no prior value, use ExcelEmpty
    }
    else
    {
        result[r, c] = v;
        last = v;
    }
}
```

**Confidence:** 0.60 (low practical risk; more about code clarity)

---

## Summary of Actionable Items

| Priority | Issue | Fix Complexity | Risk If Unfixed |
|----------|-------|-----------------|------------------|
| **CRITICAL** | Integer overflow in cast at FlattenIntIndexes:729 | Low (add bounds check) | Silent memory corruption, crash |
| **CRITICAL** | Integer overflow in Unpivot:502 multiply | Low (add checked block) | Out-of-bounds write, memory corruption |
| **CRITICAL** | Integer overflow in StackColumns:466 multiply | Low (add checked block) | Out-of-bounds write, memory corruption |
| **HIGH** | Excessive allocations in Sha256Block (separator.ToArray per cell) | Medium (refactor loop) | 100M+ GC allocations, user-visible hangs on large blocks |
| **HIGH** | Excessive allocations in HashBlock (UTF8.GetBytes per cell) | Medium (reuse buffer) | 5M+ GC allocations, sustained heap pressure |
| **HIGH** | SHA256 API misuse (many small TransformBlock calls) | Medium (batch input) | Inefficiency, potential timeout on large blocks |
| **MEDIUM** | BuildRowKey reallocation per row in RemoveDuplicateRows | Low (reuse StringBuilder) | 800MB+ allocations, GC pauses on large blocks |
| **MEDIUM** | Array.Sort stability claim contradicts implementation | Low (fix docs or use OrderBy) | Silent incorrect sort for equal keys |
| **MEDIUM** | BlockLookup duplicate-key behavior undocumented | Low (add doc comment) | User confusion, unexpected results |
| **LOW** | SplitColumn empty-parts behavior undocumented | Low (add doc comment) | User confusion |
| **LOW** | FillDown null-safety inconsistency | Low (clarify initialization) | Code clarity, potential edge-case issue |

---

## Notes on Excluded Items

- **Regex in FindReplace (line 418):** Regex compilation is done once per call (not per cell), and the regex object is reused across all cells in the block. This is correct—no per-cell recompilation. ✓
- **Transpose with non-rectangular input:** `object[,]` is always rectangular in .NET; jagged input is impossible. ✓
- **Unique with empty result (line 560-562):** Correctly returns a 1×1 block with ExcelEmpty.Value for empty input. ✓
- **Unpivot with dataRows=0 (header only):** Returns a 0×(keyColumns+2) array. Excel-DNA handles 0-row arrays cleanly (returns empty range). ✓
- **Thread safety:** All functions are marked `IsThreadSafe = false` and are safe for single-threaded (main-thread) Excel use. No concurrent mutations observed. ✓

---

## End of Audit
