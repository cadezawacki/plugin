# Deferred — round 3

Confirmed findings whose fix would require an architectural change beyond
surgical scope, plus optimization proposals intentionally not taken this
round. Each references the full write-up in the domain reports.

## Architectural proposals (domain reports, Pass 4)

- **D6 ARCH-1 — Batched *IFS variants** (`EPT.SUMIFS.BATCH(sum, crit_range,
  criteria_vector)` → spilled column): turns a column of R formulas from
  O(R·M) into O(M+R) (~5000× at R=10k, M=100k). New UDF surface; defer.
- **D6 ARCH-2 — Content-hash-keyed cross-call index cache for XLOOKUPB**:
  5–30× per recalc wave for multi-formula workbooks. Needs a bounded
  concurrent cache and an explicit staleness story (content hashing, since
  Excel never notifies UDFs of edits); defer.
- **D6 ARCH-3 — Shared typed CellKey across GROUPBY/distinct**: the
  correctness half landed (typed string tags in `AppendCellKey`, typed
  struct in LookupBoost); unifying both on one struct comparer is cleanup
  for a follow-up.
- **D6 ARCH-4 / PERF-5 — Streaming accumulators + quickselect**: the
  remaining win after this round's in-place iteration is replacing the
  matched-cells list with per-op accumulator state and Array.Sort with
  quickselect for median/percentile (~10× on the selection phase). Defer:
  rewrites the aggregation engine's shape.
- **D7 ARCH-1 — Cross-call compiled-Regex cache**: the per-call
  `RegexOptions.Compiled` threshold landed; a process-wide bounded cache
  (which is MTR-safe — Regex matching is documented thread-safe) would also
  amortize construction across recalcs. Defer: cache governance.
- **D7 ARCH-2 — Closed-form workday engine**: prefix-sum weekday table +
  sorted holiday array gives O(1) per cell (~50–150× at year-scale day
  counts) on top of this round's 3.1× serial walk. Defer: WORKDAY.INTL edge
  semantics need a property-test grid first.
- **D7 ARCH-3 — Central spill-shape governance**: `RegexUtilities.Spill`
  got explicit caps; hoisting a shared `Marshaling.AllocateSpill` for every
  spilling UDF is follow-up cleanup.
- **D7 ARCH-4 — Span/Rune text-transform layer + SIMD `DiffSquaredDot` /
  manhattan/chebyshev kernels**: PROPER measured 2.7× less garbage with
  `string.Create`; manhattan/chebyshev ~3–4× with `Vector.Abs/Max`. This
  round took the correctness (Rune) and branch-split halves; the span
  kernel layer is deferred.
- **D8 ARCH-1 — Full READFOLDER resilience surface** (opt-in
  `__source`/`__status` columns; positional `col{i}` mapping for
  header-less JSON object/scalar blocks — BUG-MEDIUM-8 cases 1–2): the
  fault isolation, safe enumeration, parallelism, and empty-block skip
  landed; the schema-affecting parts need a user-facing design decision.
- **D8 ARCH-2 — Shared atomic-write helper**: the disk cache got the
  temp+rename protocol; applying it to `EPT.WRITEJSON` and (round-2
  surface) `EPT.WRITECSV` should ride one shared helper.
- **D8 ARCH-3 / BUG-LOW-14 — Cache governance**: byte budgets, disk-cache
  TTL/pruning, and relocation out of %TEMP% (Windows Storage Sense can
  silently void the "survives reopen" promise). Defer with documentation.
- **D8 ARCH-4 — Pre-creation watches**: watching a directory that does not
  exist yet (feed arms when it appears). This round deliberately kept the
  documented first-arm `#N/A` and only added post-arm re-arming.

## Other deferred items

- **D6 — locale-dependent criteria parsing** (`">1,5"` numeric on fr-FR,
  text on en-US; date literals per CurrentCulture): mirrors Excel's own
  locale behavior; documented in `Marshaling.TryToDouble` seam notes.
  Workbook results remain machine-locale-dependent by design.
- **D7 — TitleCase/Outliers/Quantiles/Distance accepting numeric-looking
  text and bools** via `TryToDouble`: documented toolkit-wide block-function
  semantics (deliberately broader than the *IFS family's genuine-numeric
  rule). No change.
- **D7 seam — `VectorizedKernels.DotProductAvx2` loads lanes via
  `Vector256.Create(a[i],…)` (4 scalar inserts) instead of
  `Vector256.LoadUnsafe`**: measurable dot-product throughput left on the
  table for `EPT.DISTANCE`; touching the audited SIMD kernel is out of this
  round's charter.
- **D8 — `ResolveBool` copies** in DirectFileIO / JsonUtilities /
  FileSystemUtilities / TextUtilities return FALSE (not the default) for
  unrecognized strings: micro-inconsistency, fix belongs to a cross-cutting
  helper consolidation.
- **D8 MEM-4 residual**: peak retention during READFOLDER concat is now
  bounded by the largest source block, but the final `object[,]` plus boxed
  values remain inherent to the UDF return shape.
- **READFOLDER/READNDJSON unbounded result size**: same class as round-2's
  deferred F6 (file I/O) — Excel's one-bulk-array UDF model; documented
  limitation.
