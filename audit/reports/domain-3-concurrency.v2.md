# Domain 3 — Concurrency: Parallel + SIMD (Round 2)

Scope:
- `/home/user/plugin/src/ExcelPerfToolkit/ParallelUtilities.cs`
- `/home/user/plugin/src/ExcelPerfToolkit/VectorizedKernels.cs`

Methodology: line-by-line trace of every kernel and UDF; per-instruction trace
of every SIMD block-loop boundary for length classes (`0`, `1`, `width-1`,
`width`, `width+1`, large); explicit integer-arithmetic overflow analysis for
every shape multiplication; MTR contract verification for every
`IsThreadSafe=true` UDF; aliasing analysis for every parallel write path.

## Findings table

| ID  | Severity | Location                                | Category                | Confidence |
|-----|----------|-----------------------------------------|-------------------------|-----------:|
| F1  | High     | VectorizedKernels.cs:70                 | Integer overflow        | 0.95       |
| F2  | High     | VectorizedKernels.cs:89                 | Integer overflow        | 0.90       |
| F3  | High     | VectorizedKernels.cs:637,641,649,695    | Integer overflow        | 0.90       |
| F4  | High     | VectorizedKernels.cs:582-589 (L2Normalize) | Numerical correctness | 0.95       |
| F5  | Med      | ParallelUtilities.cs:466-467            | Integer overflow        | 0.85       |
| F6  | Med      | ParallelUtilities.cs:433-444 (ParallelRowReduce) | Memory / wasted alloc | 0.95   |
| F7  | Med      | VectorizedKernels.cs:154-167 (SimdAdd) etc. | Error swallowing    | 0.70       |
| F8  | Low      | ParallelUtilities.cs:289-302            | Allocation discipline   | 0.80       |
| F9  | Low      | VectorizedKernels.cs:498-519 (ColumnSums) | Excess copy          | 0.85       |
| F10 | Low      | VectorizedKernels.cs:331-332 (DotProductAvx2) | Codegen quality   | 0.65       |
| F11 | Low      | ParallelUtilities.cs:434-437 (rowIndex dropped) | Latent invariant | 0.75      |

---

## F1 — Critical-path integer overflow in `FlattenToDoubles` allocation

- ID: F1
- Severity: High
- Location: `VectorizedKernels.cs:70`
- Category: Integer overflow / unchecked arithmetic
- Concrete failure scenario: A user invokes any SIMD UDF (`EPT.SIMD.ADD`,
  `EPT.SIMD.MULTIPLY`, `EPT.SIMD.SCALE`, `EPT.SIMD.DOT`, `EPT.SIMD.ROWSUMS`,
  `EPT.SIMD.COLSUMS`, `EPT.SIMD.NORMALIZE`, `EPT.SIMD.MATMUL`) on a sufficiently
  large block — e.g. a whole-sheet reference on a modern Excel sheet with
  `rows=1_048_576, cols=16_384` (the worksheet maximum), or even more modest
  shapes such as `rows=200_000, cols=12_000 = 2_400_000_000`. The expression
  `rows * cols` is computed as `int` and wraps: `200_000 * 12_000 = 2.4e9` >
  `int.MaxValue (2_147_483_647)`. With overflow checking off (the C# default),
  the wrap silently produces a smaller positive (or negative) int. `new
  double[rows*cols]` then either throws `OverflowException`/`ArgumentOutOfRangeException`
  (negative wrap), or — worse — allocates a *too-small* buffer and the
  subsequent `result[idx++] = ...` loop overruns it, throwing
  `IndexOutOfRangeException` mid-population. Either way the kernel terminates
  with an unhandled exception that is rethrown by the UDF wrapper's outer
  catch into `#VALUE!`. The exception path is reached but the kernel
  reports a misleading reason and the workbook silently produces the wrong
  result for the user's query.
- Evidence (code trace):
  ```csharp
  // VectorizedKernels.cs:65-80
  internal static double[] FlattenToDoubles(object[,] block, out int rows, out int cols)
  {
      ArgumentNullException.ThrowIfNull(block);
      rows = block.GetLength(0);
      cols = block.GetLength(1);
      var result = new double[rows * cols];  // <-- int multiply, wraps
      var idx = 0;
      for (var r = 0; r < rows; r++)
      {
          for (var c = 0; c < cols; c++)
          {
              result[idx++] = Marshaling.TryToDouble(block[r, c], out var d) ? d : 0d;
          }
      }
      return result;
  }
  ```
  `block.GetLength(int)` returns `int`. `rows*cols` is computed with the
  multiplicative `int*int` operator and is not in a `checked` context. For
  `rows=200_000, cols=12_000` the result is `(int)(2_400_000_000L) = -1_894_967_296`,
  which `new double[]` rejects with `OverflowException`/`ArgumentOutOfRangeException`.
  For `rows=65_536, cols=65_536 = 4_294_967_296 ≡ 0 mod 2^32`, `new double[0]`
  succeeds, and the inner fill at `result[idx++]` immediately overruns on the
  first iteration.
- Proposed surgical fix: compute the product as `long` and validate:
  ```csharp
  long total = (long)rows * cols;
  if (total > int.MaxValue) throw new ArgumentException("Block too large to flatten (rows*cols exceeds Int32.MaxValue).");
  var result = new double[(int)total];
  ```
- Confidence: 0.95

---

## F2 — Same overflow in `BoxFlatDoubles` length check

- ID: F2
- Severity: High
- Location: `VectorizedKernels.cs:89`
- Category: Integer overflow
- Concrete failure scenario: All callers reach `BoxFlatDoubles` *after*
  `FlattenToDoubles` has already crashed for the same overflowing dimensions,
  so this is unreachable on its own through current call sites. However, the
  helper is `internal` and any future caller (or a refactor that pre-allocates
  `flat`) hits the same defect: the equality `flat.Length != rows * cols`
  silently lies because the right-hand side wraps. Worst case the function
  proceeds to write into a wrong-shape `object[,]` and corrupts data.
- Evidence:
  ```csharp
  // VectorizedKernels.cs:86-103
  internal static object[,] BoxFlatDoubles(double[] flat, int rows, int cols)
  {
      ArgumentNullException.ThrowIfNull(flat);
      if (flat.Length != rows * cols)  // <-- int multiply, wraps
      {
          throw new ArgumentException("Flat buffer length does not match dimensions.", nameof(flat));
      }
      ...
  }
  ```
- Proposed surgical fix:
  ```csharp
  long expected = (long)rows * cols;
  if (flat.Length != expected) throw new ArgumentException(...);
  ```
- Confidence: 0.90

---

## F3 — `MatrixMultiply` shape arithmetic overflow

- ID: F3
- Severity: High
- Location: `VectorizedKernels.cs:637, 641, 649, 695, 696`
- Category: Integer overflow
- Concrete failure scenario: `EPT.SIMD.MATMUL` with `a` shape `[m=50_000,
  k=50_000]` and `b` shape `[k=50_000, n=50_000]`. `m*k = 2_500_000_000 >
  int.MaxValue`; the validity checks `a.Length != m*k` etc. compare against a
  wrapped value, so an A buffer that *actually* has the correct length compares
  unequal to the wrapped product and throws an erroneous "A's length does not
  match m*k" — or, for a shape where both `m*k` and the actual `a.Length` wrap
  to the same int, the check passes but `destination[i * n + j]` at line 667
  uses `i*n+j` (also `int*int`); for `i*n` near the wrap point, this either
  throws `IndexOutOfRangeException` or writes the wrong cell, *silently
  corrupting the output matrix*. Same for `b[p * n + j]` at line 662.
- Evidence:
  ```csharp
  // VectorizedKernels.cs:632-670
  public static void MatrixMultiply(
      ReadOnlySpan<double> a, int m, int k,
      ReadOnlySpan<double> b, int kBRows, int n,
      Span<double> destination)
  {
      if (a.Length != m * k) ...                  // line 637
      if (b.Length != kBRows * n) ...             // line 641
      if (destination.Length != m * n) ...        // line 649
      ...
      for (var j = 0; j < n; j++) {
          for (var p = 0; p < k; p++) {
              bCol[p] = b[p * n + j];             // line 662, p*n int
          }
          for (var i = 0; i < m; i++) {
              ...
              destination[i * n + j] = ...;       // line 667, i*n int
          }
      }
  }
  ```
  And `SimdMatMul`:
  ```csharp
  // VectorizedKernels.cs:695-696
  var dest = new double[m * n];                   // line 695, int wrap
  MatrixMultiply(flatA, m, k, flatB, kB, n, dest);
  ```
- Proposed surgical fix: cast indices and allocations to long, validate up
  front, and use `checked` arithmetic for the inner index expressions (or
  precompute `i * (long)n + j` and cast to `int` only after a guard).
- Confidence: 0.90

---

## F4 — `L2Normalize` produces all-zeros for finite-but-large vectors (overflow in sumSq)

- ID: F4
- Severity: High
- Location: `VectorizedKernels.cs:576-591`
- Category: Numerical correctness / silent data corruption
- Concrete failure scenario: User normalizes a vector with magnitude near
  `double.MaxValue`, e.g. `[1e200, 1e200]`. The dot product accumulates `1e400
  + 1e400 = +Infinity` (each square overflows to `+Inf` individually). Then
  `Math.Sqrt(+Inf) = +Inf`, which is not equal to `0d`, so the
  zero-norm short-circuit at line 584 does not fire. `Scale(source, 1d /
  +Inf, dest)` evaluates `1d / +Inf = +0d`, and every element of `dest` becomes
  `source[i] * 0 = 0` (assuming finite source elements). The function returns
  `+Infinity` as the norm but writes a zero vector to `destination` — the
  correct answer for `[1e200, 1e200]` is `[1/√2, 1/√2]`. The caller is told
  the operation succeeded (no exception, no error code) but receives garbage.
- Evidence:
  ```csharp
  // VectorizedKernels.cs:576-591
  public static double L2Normalize(ReadOnlySpan<double> source, Span<double> destination)
  {
      ...
      var sumSq = DotProduct(source, source);     // can overflow to +Inf
      var norm = Math.Sqrt(sumSq);                // sqrt(+Inf) = +Inf
      if (norm == 0d)
      {
          destination.Clear();
          return 0d;
      }
      Scale(source, 1d / norm, destination);      // 1/+Inf = +0, dest=0
      return norm;                                // returns +Inf
  }
  ```
- Proposed surgical fix: either (a) two-pass normalize by max element to avoid
  overflow:
  ```csharp
  double maxAbs = 0;
  for (int i = 0; i < source.Length; i++) maxAbs = Math.Max(maxAbs, Math.Abs(source[i]));
  if (maxAbs == 0) { destination.Clear(); return 0; }
  // Scale down, then accumulate, then scale up the norm
  double inv = 1.0 / maxAbs;
  double sumSq = 0;
  for (int i = 0; i < source.Length; i++) { var v = source[i] * inv; sumSq += v * v; }
  double norm = maxAbs * Math.Sqrt(sumSq);
  ```
  or (b) explicitly handle the `double.IsInfinity(norm)` case by returning a
  defined error (UDF returns `#VALUE!`) instead of silently emitting zeros.
- Confidence: 0.95

---

## F5 — `ParallelUtilities.Dot` integer overflow on `na`/`nb` length comparison

- ID: F5
- Severity: Med
- Location: `ParallelUtilities.cs:466-467`
- Category: Integer overflow
- Concrete failure scenario: User calls `EPT.MT.DOT(a, b)` with two large
  blocks of mismatched-but-overflowing shapes. Example: `a` is `[80_000,
  30_000]` (total 2.4e9), `b` is `[30_000, 80_000]` (also 2.4e9 — but
  conceptually the same number of elements). Both overflow to the same wrapped
  int. The function proceeds — fine in this case. Worse example: `a` is
  `[65_536, 65_536]` total `2^32 ≡ 0`; `b` is `[1, 1]` total `1`. `na=0, nb=1`,
  guard fires correctly. But consider `a` `[2_147_483_648 / cb, cb]` where the
  product overflows to `int.MinValue`: comparison with another negative gives
  wrong equality. Inside the loop, `ia++` is incremented up to `ra*ca`
  iterations; at iteration 2^31 it overflows to `int.MinValue`, then `ia/cb`
  is negative, then `b[br, bc]` throws `IndexOutOfRangeException` — propagates
  out of the thread-safe UDF as an unhandled exception and produces `#VALUE!`
  (Excel-DNA default for unhandled exceptions from UDFs).
- Evidence:
  ```csharp
  // ParallelUtilities.cs:462-489
  var ra = a.GetLength(0);
  var ca = a.GetLength(1);
  var rb = b.GetLength(0);
  var cb = b.GetLength(1);
  var na = ra * ca;          // line 466, int wrap
  var nb = rb * cb;          // line 467, int wrap
  if (na != nb) return ExcelError.ExcelErrorValue;
  ...
  var ia = 0;
  for (var r = 0; r < ra; r++)
      for (var c = 0; c < ca; c++) {
          var br = ia / cb;
          var bc = ia % cb;
          ...
          ia++;              // overflows to int.MinValue after 2^31 increments
      }
  ```
- Proposed surgical fix: validate `(long)ra*ca == (long)rb*cb` and bail to
  `#VALUE!` if either side exceeds `int.MaxValue` (Excel sheets can technically
  reach `1_048_576 * 16_384 = 1.7e10`, so the cap is hit by whole-sheet args).
- Confidence: 0.85

---

## F6 — `ParallelRowReduce` allocates `expanded[rows, cols]` then discards all but column 0

- ID: F6
- Severity: Med
- Location: `ParallelUtilities.cs:429-444`
- Category: Memory waste / latent correctness if op writes more than `dst[0]`
- Concrete failure scenario: For a `[1_000_000, 1_000]` input block the
  function intends to return a `[1_000_000, 1]` reduction. It allocates
  `expanded = new double[1_000_000, 1_000]` = 8 GB of doubles to satisfy
  `ParallelBatchTransform`'s precondition `output.GetLength(1) == input.GetLength(1)`,
  then reads only column 0 into a `[1_000_000, 1]` result. The peak working
  set increases by ~8 GB for no semantic reason. On a 16 GB workstation this
  triggers `OutOfMemoryException` (caught nowhere here — propagates as
  `#VALUE!` from Excel-DNA). On a smaller box it may cause paging and bring
  the whole Excel session to a crawl. The comment at lines 429-432 acknowledges
  this is a hack ("the helper writes one full row to output, we expand and
  then take col 0"). It is also wrong in spirit: each row's `dst[c]` for
  `c > 0` is whatever stale data the thread's `outBuf` holds from a previous
  row — currently harmless because we drop those columns, but a footgun for
  future maintainers and for any reduction op that decides to write `dst[1]`.
- Evidence:
  ```csharp
  // ParallelUtilities.cs:429-444
  // The reduction output is one column wide; the row buffer in the helper is sized
  // to the input cols, but the writer only touches dst[0..1]. Allocate matching
  // shapes by using a 2-D output of width 1 and an input projection trick:
  // since the helper writes one full row to output, we expand and then take col 0.
  var expanded = new double[rows, cols];          // <-- O(rows*cols) waste
  ParallelBatchTransform(input, expanded, (_, src, dst) =>
  {
      body(0, src, dst);
  });
  var result = new object[rows, 1];
  for (var r = 0; r < rows; r++)
  {
      result[r, 0] = expanded[r, 0];
  }
  return result;
  ```
- Proposed surgical fix: add a `ParallelBatchTransform` overload that takes a
  scalar-output delegate `(int row, ReadOnlySpan<double> src, out double dst)`
  and writes directly into a `double[rows]` accumulator, or relax the helper's
  precondition to allow `output.GetLength(1) < input.GetLength(1)` provided
  the body declares its written width.
- Confidence: 0.95

---

## F7 — UDFs catch and swallow `OutOfMemoryException` and other fatal exceptions

- ID: F7
- Severity: Med
- Location: `VectorizedKernels.cs:154-167, 213-228, 273-287, 387-403, 458-476, 545-563, 605-618, 685-703`
- Category: Error swallowing / process invariant violation
- Concrete failure scenario: Each SIMD UDF wraps its body in `try { ... }
  catch (Exception ex) { TraceSource.TraceEvent(Warning, ...); return
  ExcelError.ExcelErrorValue; }`. A bare `catch (Exception)` traps
  `OutOfMemoryException`, `AccessViolationException` (if marshaled), and
  `StackOverflowException`'s subordinate frames (only the top SOE bypasses
  catch). Under MTR, when one worker thread hits OOM during the
  multi-GB allocation in `FlattenToDoubles` for a max-shape input, the OOM
  is converted to a warning trace + `#VALUE!`. The runtime is in a degraded
  state (managed heap is exhausted; subsequent allocations on other threads
  may also fail) but Excel continues calling more UDFs, each of which then
  also OOMs. The user sees `#VALUE!` everywhere with no log breadcrumb that a
  fatal allocator failure was the root cause. A bare `catch (Exception)` here
  also conceals genuine bugs (e.g. a null-deref in a future refactor) from
  the developer because the only signal is a TraceSource warning that nobody
  watches under normal operation.
- Evidence: representative case at `VectorizedKernels.cs:154-167`
  ```csharp
  public static object SimdAdd(object[,] a, object[,] b)
  {
      try
      {
          ValidateSameShape(a, b);
          var flatA = FlattenToDoubles(a, out var rows, out var cols);
          var flatB = FlattenToDoubles(b, out _, out _);  // may OOM
          var dest = new double[flatA.Length];            // may OOM
          ElementWiseAdd(flatA, flatB, dest);
          return BoxFlatDoubles(dest, rows, cols);
      }
      catch (Exception ex)                                // catches OOM
      {
          TraceSource.TraceEvent(TraceEventType.Warning, 1, "EPT.SIMD.ADD failed: {0}", ex.Message);
          return ExcelError.ExcelErrorValue;
      }
  }
  ```
- Proposed surgical fix: narrow the catch to expected exception types
  (`ArgumentException`, `InvalidOperationException`,
  `IndexOutOfRangeException`, `OverflowException`), and either let
  `OutOfMemoryException` propagate (the safer choice — Excel-DNA's host
  fence will translate it) or rethrow it after tracing:
  ```csharp
  catch (Exception ex) when (ex is not OutOfMemoryException and not StackOverflowException)
  { ... return ExcelError.ExcelErrorValue; }
  ```
- Confidence: 0.70

---

## F8 — `ParallelBatchTransform` per-thread buffers re-allocated each call; no pooling

- ID: F8
- Severity: Low
- Location: `ParallelUtilities.cs:289-302`
- Category: Allocation discipline / GC pressure
- Concrete failure scenario: Each `Parallel.For` invocation creates fresh
  `(double[cols], double[cols])` buffers via `localInit` — one tuple per
  thread per call. For `cols=16_384` that is `2 * 16_384 * 8 = 256 KB` per
  thread, each allocation landing on LOH (>= 85 KB). On an 8-core box this
  is 2 MB of LOH allocations per `ParallelBatchTransform` call. Under a recalc
  that fans out e.g. 1000 cells calling `EPT.PARALLELTANH` over a wide block,
  that is 2 GB of LOH allocation pressure per recalc, defeating the kernel's
  intent (avoiding allocation in the hot loop). The `localFinally: _ => { }`
  is a no-op — the buffers are simply released to the GC. No pooling.
- Evidence:
  ```csharp
  // ParallelUtilities.cs:289-302
  Parallel.For(
      fromInclusive: 0,
      toExclusive: rows,
      parallelOptions: options,
      localInit: () => (new double[cols], new double[cols]),  // 2x LOH
      body: ...,
      localFinally: _ => { });                                // dropped
  ```
- Proposed surgical fix: pool the buffers via `ArrayPool<double>.Shared.Rent(cols)`
  in `localInit` and `Return` in `localFinally`.
- Confidence: 0.80

---

## F9 — `ColumnSums` SIMD path round-trips every store through a `stackalloc` buffer

- ID: F9
- Severity: Low
- Location: `VectorizedKernels.cs:498-519`
- Category: Excess copy / suboptimal codegen
- Concrete failure scenario: For each `width`-aligned column block, the code
  does `(vCol + vRow).CopyTo(buffer)` then `buffer.CopyTo(colSums.Slice(c,
  width))`. The intermediate `stackalloc double[width]` buffer adds two extra
  span copies per inner iteration. The `Vector<double>.CopyTo(Span<double>)`
  API accepts the destination directly. For a `[10_000, 16_384]` block, that
  is `10_000 * (16_384 / 4) * 2 extra copies = ~80M unnecessary copy ops`.
  Not a bug, but a clear performance miss for a kernel whose only purpose is
  speed.
- Evidence:
  ```csharp
  // VectorizedKernels.cs:498-519
  Span<double> buffer = stackalloc double[width];
  for (var r = 0; r < rows; r++)
  {
      var row = source.Slice(r * cols, cols);
      var c = 0;
      for (; c <= cols - width; c += width)
      {
          var vRow = new Vector<double>(row.Slice(c, width));
          var vCol = new Vector<double>(colSums.Slice(c, width));
          (vCol + vRow).CopyTo(buffer);                  // extra copy 1
          buffer.CopyTo(colSums.Slice(c, width));        // extra copy 2
      }
      ...
  }
  ```
- Proposed surgical fix:
  ```csharp
  (vCol + vRow).CopyTo(colSums.Slice(c, width));
  ```
- Confidence: 0.85

---

## F10 — `DotProductAvx2` uses `Vector256.Create(a[i], a[i+1], a[i+2], a[i+3])` instead of `LoadUnsafe`

- ID: F10
- Severity: Low
- Location: `VectorizedKernels.cs:331-332`
- Category: Codegen quality / SIMD inefficiency
- Concrete failure scenario: The XML comment at line 324-325 says
  "Vector256.LoadUnsafe over Span, which is the safe API that compiles to
  vmovupd". The actual code uses `Vector256.Create(a[i], a[i+1], a[i+2],
  a[i+3])` — a 4-arg constructor — which JIT-compiles to four scalar loads
  followed by lane inserts (`vmovsd` + `vunpcklpd` chain) rather than a single
  `vmovupd`. The benchmark numbers for "AVX2 path" therefore understate the
  hardware's true throughput by ~2-4x on the load side. Not a correctness
  bug. The advertised optimisation does not match the code.
- Evidence:
  ```csharp
  // VectorizedKernels.cs:322-336
  private static double DotProductAvx2(ReadOnlySpan<double> a, ReadOnlySpan<double> b)
  {
      // We use Vector256.LoadUnsafe over Span, which is the safe API that compiles
      // to vmovupd; AllowUnsafeBlocks remains disabled.
      var width = Vector256<double>.Count;
      var accum = Vector256<double>.Zero;
      var i = 0;
      for (; i <= a.Length - width; i += width)
      {
          var va = Vector256.Create(a[i], a[i + 1], a[i + 2], a[i + 3]);  // 4 loads
          var vb = Vector256.Create(b[i], b[i + 1], b[i + 2], b[i + 3]);  // 4 loads
          accum = FmaSupported
              ? Fma.MultiplyAdd(va, vb, accum)
              : Avx.Add(accum, Avx.Multiply(va, vb));
      }
      ...
  }
  ```
- Proposed surgical fix: use the documented API:
  ```csharp
  var va = Vector256.LoadUnsafe(ref System.Runtime.InteropServices.MemoryMarshal.GetReference(a), (nuint)i);
  var vb = Vector256.LoadUnsafe(ref System.Runtime.InteropServices.MemoryMarshal.GetReference(b), (nuint)i);
  ```
  (still no `unsafe` block required).
- Confidence: 0.65

---

## F11 — `ParallelRowReduce` body invocation hard-codes rowIndex=0

- ID: F11
- Severity: Low
- Location: `ParallelUtilities.cs:434-437`
- Category: Latent invariant / dead parameter
- Concrete failure scenario: The lambda passed to `ParallelBatchTransform`
  invokes `body(0, src, dst)` discarding the actual row index. None of the
  current reductions (`sum`, `mean`, `min`, `max`, `stdev`) consult their
  row-index parameter — they bind it to `_`. But the invariant is silently
  broken: the `RowTransform` contract says `rowIndex` is the row being
  processed, and any future operation that needs it (e.g. row-weighted sum)
  would silently receive `0` for every row, producing wrong answers with no
  exception. This is a tripwire for future maintainers.
- Evidence:
  ```csharp
  // ParallelUtilities.cs:434-437
  ParallelBatchTransform(input, expanded, (_, src, dst) =>
  {
      body(0, src, dst);   // <-- should be (row, src, dst)
  });
  ```
  Note the outer lambda's first parameter (the real row index) is discarded
  with `_`.
- Proposed surgical fix: forward the row index:
  ```csharp
  ParallelBatchTransform(input, expanded, (row, src, dst) =>
  {
      body(row, src, dst);
  });
  ```
- Confidence: 0.75

---

## Rejected findings (round 2)

The following candidate issues were investigated and explicitly rejected with
a concrete trace.

- **OOB read in `DotProductAvx2` for `Length=5, width=4`** (round-1 finding).
  Trace: `for (; i <= a.Length - width; i += width)` → `i <= 1`. i=0 reads
  a[0..3] (OK), i=4 (4 <= 1 false) exits. Tail: i=4 < 5 → reads a[4] (OK).
  No OOB. **Rejected.**

- **Per-output-column allocation in `MatrixMultiply`.** `var bCol = new
  double[k]` at line 657 is outside the `j` loop (which begins line 658).
  Allocated once, reused per j. **Rejected.**

- **Type-init failure on non-x86 CPUs from `Avx2.IsSupported` / `Fma.IsSupported`.**
  Per .NET docs, these `IsSupported` static properties return `false` on
  unsupported platforms; they do not throw `PlatformNotSupportedException`.
  Static initializer (lines 41, 48, 53) is safe. **Rejected.**

- **`PolyEval` crash for `coeffs.Length == 1`.** Trace: `acc =
  coeffs[coeffs.Length - 1] = coeffs[0]`. Then `for (var i = -1; i >= 0; ...)`
  — the loop condition `-1 >= 0` is false on first evaluation, never enters.
  Returns `acc = coeffs[0]`. Correct for a constant polynomial. **Rejected.**

- **`PolyEval` crash for `coeffs.Length == 0`.** The guard at line 200
  `if (coeffsList.Count == 0)` returns `ErrorBlock` before the kernel runs.
  **Rejected.**

- **`ElementWiseAdd`/`Multiply`/`Scale` SIMD block-loop OOB for `Length=0`,
  `Length=1`, `Length=width-1`.** Trace: SIMD loop condition `i <= length -
  width` with `length=0, width=4` is `i <= -4`, never enters; tail `i < 0`
  never enters; no work, no error. For `length=1, width=4`: `i <= -3` never
  enters; tail i=0 runs; OK. **Rejected.**

- **`Dot` (ParallelUtilities) computes a "wrong" inner product.** Trace
  reproduced in the task brief: indexing `b[ia/cb, ia%cb]` with `ia`
  incremented row-major across `a` reproduces a row-major flat-vector dot
  product regardless of which side is row- or column-shaped, exactly matching
  the docstring "Inputs may be column-shaped or row-shaped; both are flattened
  in row-major order." **Rejected.**

- **`RowSums` (SIMD) tail handling when `cols < Vector<double>.Count`.**
  Gate at line 429 (`cols >= Vector<double>.Count`) skips SIMD; tail loop
  runs for `i in [0, cols)`. For `cols=0`: tail also doesn't enter; `sum=0`
  written. For `cols=1`: tail runs once, sum=row[0]. **Rejected.**

- **`stackalloc` blowing the stack in `ColumnSums`.** `stackalloc
  double[width]` for `width ∈ {2,4,8}` on current CPUs = 16-64 bytes. Safe
  even under MTR worker thread stacks. **Rejected.**

- **`MatrixMultiply` shape check post-`SimdMatMul`'s `k!=kB` short-circuit.**
  `SimdMatMul` returns `#VALUE!` before calling `MatrixMultiply` on shape
  mismatch (line 691-694). MatrixMultiply's own checks are belt-and-braces.
  **Rejected.**

- **Race in `ParallelBatchTransform` writes to `output`.** `Parallel.For`
  partitions the iteration space such that each `r` is assigned to exactly
  one worker. Each iteration writes only `output[r, *]`. No two threads write
  the same cell. The .NET memory model guarantees that the joining `Parallel.For`
  acts as a release/acquire barrier, so the caller sees all writes after
  return. **Rejected.**

- **Static `TraceSource` field thread-safety in `ParallelUtilities`.**
  `TraceSource.TraceEvent` is documented thread-safe; the field is initialized
  once at type init and never mutated. **Rejected.**

- **`ParallelTanh` re-entrance with MTR creating N*M threads.** UDF is
  registered `IsThreadSafe=false` (line 337) precisely to prevent this. The
  comment at lines 327-329 documents the rationale. **Rejected.**

- **`SimdCapabilities` MTR-unsafe due to reading static properties.** The
  properties are read-only auto-properties initialized at type init; reading
  them from multiple threads is safe (a single torn read of a `bool` field
  on a x64 or ARM64 .NET runtime is atomic; even if it weren't, the value
  never changes). **Rejected.**

## Out-of-scope notes

(No issues to report outside the two files in scope.)
