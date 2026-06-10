# Domain 6 — Conditional aggregation & lookups (round 3)

Files audited (every line):
- `/home/user/plugin/src/ExcelPerfToolkit/ConditionalAggregates.cs` (1093 LOC)
- `/home/user/plugin/src/ExcelPerfToolkit/LookupBoost.cs` (211 LOC)

Contract context read (no findings reported in): `Marshaling.cs`, `readme.md`.

All "Empirically verified" items were tested on the .NET 8.0.422 SDK at `/tmp/dotnet/dotnet`
(`/tmp/probe6`, Release). Representative sizing throughout: **100k-row single-criterion
`EPT.SUMIFS` call** and **100k-key / 100k-probe `EPT.XLOOKUPB` call**.

---

## Summary table

| ID | Severity | One-line |
| --- | --- | --- |
| BUG-HIGH-001 | HIGH | Date-literal criteria (`">1/1/2025"`) silently degrade to ordinal string comparison — matches essentially every date |
| BUG-MEDIUM-002 | MEDIUM | `<>` criteria never match blank cells; native COUNTIFS/SUMIFS count blanks for `"<>x"` |
| BUG-MEDIUM-003 | MEDIUM | Booleans conflated with numbers in both criteria matching and value-range aggregation |
| BUG-MEDIUM-004 | MEDIUM | Text/wildcard criteria match numeric/bool cells via their string forms (`"*"` counts numbers) |
| BUG-MEDIUM-005 | MEDIUM | XLOOKUPB exact match collides keys across types (5 vs `"5"`, TRUE vs `"TRUE"`, errors as `"#N/A"` text) |
| BUG-MEDIUM-006 | MEDIUM | Errors in the value range of matched rows are silently dropped instead of propagating |
| BUG-MEDIUM-007 | MEDIUM | XLOOKUPB `match_mode` convention is the inverse of VLOOKUP's and ignores XLOOKUP's numeric modes |
| BUG-LOW-008 | LOW | GROUPBY composite-key `\u001f` separator can merge distinct groups |
| BUG-LOW-009 | LOW | GROUPBY (Ordinal, case-sensitive) vs DISTINCTCOUNTIFS (OrdinalIgnoreCase) key policy inconsistency + type conflation |
| BUG-LOW-010 | LOW | Criteria ranges validated by cell count only, not shape (2×3 accepted against 3×2) |
| BUG-LOW-011 | LOW | Approximate lookup returns an arbitrary duplicate's value (unstable `List<T>.Sort`) |
| BUG-LOW-012 | LOW | Error criteria (`"#N/A"`) can never match error cells; native COUNTIF matches them |
| BUG-LOW-013 | LOW | Multi-cell range passed as criteria becomes the literal text `"System.Object[,]"` and silently matches nothing |
| MEM-1 | — | `FlattenRowMajor` copies every range: 1.6 MB/call avoidable |
| MEM-2 | — | `AggregateIfs` full-capacity `matched` list + second `nums` list: up to 1.6 MB/call |
| MEM-3 | — | `CountIfs` flattens a whole range just to read its length: 800 KB/call |
| MEM-4 | — | GROUPBY allocates a `keyRow` array for every row, not every group: 3.2 MB/100k rows |
| MEM-5 | — | XLOOKUPB stringifies every key and every probe: 12.6 MB/call measured |
| MEM-6 | — | Text-criterion mask loop stringifies every non-string cell: 6.3 MB/100k numeric cells |
| PERF-1 | — | Layered per-cell dispatch in mask build (2–3× on the mask phase) |
| PERF-2 | — | Four passes + three materializations where one streaming pass suffices (1.5–2×) |
| PERF-3 | — | ApproximateLookup always sorts: 47.6 ms vs 0.21 ms sortedness check (≈5× whole call when presorted) |
| PERF-4 | — | ExactLookup string keys + double hashing: 3.2× index build measured |
| PERF-5 | — | Median/percentile full `Array.Sort` (29.3 ms/100k) where quickselect is ~10× cheaper |
| PERF-6 | — | WSTDEV converts every cell four times across its two passes (~1.7×) |
| PERF-7 | — | XLOOKUPB builds the full index even for a single probe (~60× for R=1 calls) |
| PERF-8 | — | GROUPBY re-parses the operation string (Trim+ToLowerInvariant) once per group |
| ARCH-1 | — | Batched *IFS variant (criteria vector → spilled column): O(R·M) → O(M+R) |
| ARCH-2 | — | Content-hash-keyed cross-call index cache for XLOOKUPB |
| ARCH-3 | — | Typed `CellKey` comparer shared by lookup/groupby/distinct (fixes the type-conflation bug family) |
| ARCH-4 | — | Criterion-compiled, streaming-accumulator aggregation engine |

---

## Pass 1 — Deep bug scan

```
BUG-HIGH-001
File: /home/user/plugin/src/ExcelPerfToolkit/ConditionalAggregates.cs:786-795, 834-851
Category: logic_error
Description: A criterion whose operand is a date literal — ">1/1/2025", "<=12/31/2024",
  ">=2025-01-01" — is not recognized as a date. Criterion.Parse tries only
  Marshaling.TryToDouble on the operand (line 791); date strings fail numeric parsing
  (empirically verified: double.TryParse("1/1/2025", Float, Invariant) = false), so the
  criterion falls through to the TEXT path (line 795). Matches() then stringifies every
  numeric date cell via ToStringSafe ("R" format → "45658") and performs an ORDINAL
  string comparison against "1/1/2025" (line 843). Native Excel COUNTIF/SUMIFS parse
  date-literal operands into serial numbers and compare numerically.
Trigger condition: =EPT.COUNTIFS(A1:A100000, ">1/1/2025") where column A holds dates.
Trace (cell = DATE(2020,6,15), serial 43997.0 — a date that should NOT match ">1/1/2025"):
  1. Parse(">1/1/2025"): s.StartsWith(">") → op=Greater, start=1 (lines 770-774).
  2. operand = "1/1/2025" (line 786). TryToDouble fails → line 795: text criterion,
     _text="1/1/2025", _isNumeric=false.
  3. Matches(43997.0): not error, not blank, _isNumeric false → line 834:
     text = ToStringSafe(43997.0) = "43997" (verified: 45658d.ToString("R") = "45658").
  4. Line 843: string.Compare("43997", "1/1/2025", OrdinalIgnoreCase) = 3 (> 0,
     empirically verified) because '4' (0x34) > '1' (0x31).
  5. Op.Greater → cmp > 0 → MATCH. Every serial ≥ 10 matches (verified:
     cmp("10","1/1/2025") = 1; only serial "1" compares < 0), so ">1/1/2025" matches
     effectively every date in the sheet → COUNTIFS returns ~the full count, SUMIFS
     ~the full sum. Silently, plausibly wrong.
Fix (Criterion.Parse, after the numeric attempt at line 791):
  before:
      if (Marshaling.TryToDouble(operand, out var num))
      {
          return new Criterion(op, false, true, num, operand, Array.Empty<Tok>());
      }
      return new Criterion(op, false, false, 0d, operand, Tokenize(operand));
  after:
      if (Marshaling.TryToDouble(operand, out var num))
      {
          return new Criterion(op, false, true, num, operand, Array.Empty<Tok>());
      }
      if (DateTime.TryParse(operand, CultureInfo.CurrentCulture, DateTimeStyles.None, out var dt))
      {
          return new Criterion(op, false, true, dt.ToOADate(), operand, Array.Empty<Tok>());
      }
      return new Criterion(op, false, false, 0d, operand, Tokenize(operand));
  (CurrentCulture mirrors how Excel itself interprets date-literal criteria per locale.
  Restrict to ops other than Equal-with-wildcards if desired; comparison ops are the
  dangerous case. Consider TimeSpan literals out of scope.)
```

```
BUG-MEDIUM-002
File: /home/user/plugin/src/ExcelPerfToolkit/ConditionalAggregates.cs:804-812
Category: logic_error
Description: Matches() returns false for every blank cell once the criterion has a
  non-empty operand (line 809-812: `if (blank) { return false; }`), including for the
  NotEqual operator. Native Excel counts blank cells for "<>" criteria: COUNTIF(range,
  "<>5") and COUNTIF(range, "<>apple") both include empty cells (a blank is "not 5" /
  "not apple"). The code's own comment at lines 818-820 claims `a "<>" criterion is
  satisfied by everything that is not that number, including text` — but blanks are
  rejected three lines earlier.
Trigger condition: A1=5, A2=(empty), A3=3. =EPT.COUNTIFS(A1:A3, "<>5") returns 1;
  native COUNTIFS returns 2.
Trace:
  1. BuildMask → criterion "<>5": op=NotEqual, _isNumeric=true, _number=5.
  2. Matches(A2 = ExcelEmpty): line 804 blank=true; _emptyOperand=false so line 805-808
     skipped; line 809-812 → return false.
  3. Mask[1]=false → count = 1 (A3 only). Native = 2.
Fix:
  before:
      if (blank)
      {
          return false;
      }
  after:
      if (blank)
      {
          return _op == Op.NotEqual;
      }
```

```
BUG-MEDIUM-003
File: /home/user/plugin/src/ExcelPerfToolkit/ConditionalAggregates.cs:745-746, 1061-1069
  (plus every numeric-extraction site: 476-483, 246, 289-294, 303, 342, 582, 606)
Category: logic_error
Description: Booleans are conflated with numbers in both directions. (a)
  Criterion.Parse converts a bool criterion to the number 1/0 (lines 745-746), so
  criterion TRUE matches numeric-1 cells and criterion 1 matches TRUE cells — native
  Excel keeps logicals a distinct type (COUNTIF(A1,TRUE) with A1=1 returns 0). (b)
  IsNumericCell (lines 1061-1069) excludes strings but not bools; Marshaling.TryToDouble
  maps bool→1/0, so TRUE cells in a sum/average/min/max/value range are aggregated as
  1 — native SUMIFS/AVERAGEIFS/MINIFS ignore logical values in the value range
  (TRUE contributes 0 to the sum and is excluded from the AVERAGE denominator).
Trigger condition: A1:A3 = {TRUE, 1, 2}; B1:B3 = {10, 20, 30}.
  - =EPT.COUNTIFS(A1:A3, 1) returns 2 (TRUE matched); native COUNTIFS returns 1.
  - =EPT.SUMIFS(A1:A3, B1:B3, ">0") returns 4 (TRUE summed as 1); native returns 3.
  - =EPT.AVERAGEIFS(A1:A3, B1:B3, ">0") returns 4/3; native returns 3/2.
Trace (=EPT.SUMIFS(A1:A3, B1:B3, ">0")):
  1. Mask = [true,true,true].
  2. Aggregate("sum"): nums extraction loop (476-483): IsNumericCell(true) → not a
     string → TryToDouble(bool) → 1d, returns true → nums = [1, 1, 2] → sum = 4.
  3. Native Excel: logicals ignored in sum range → 1 + 2 = 3.
Fix:
  IsNumericCell, before:
      if (v is string)
      {
          d = 0d;
          return false;
      }
      return Marshaling.TryToDouble(v, out d);
  after:
      if (v is string or bool)
      {
          d = 0d;
          return false;
      }
      return Marshaling.TryToDouble(v, out d);
  Criterion: add a bool branch that matches only bool cells of equal value:
      case bool b:
          return new Criterion(Op.Equal, false, isNumeric: false, 0d,
              b ? "TRUE" : "FALSE", Array.Empty<Tok>());   // plus a _isBool flag and a
  dedicated `cell is bool cb && cb == _bool` arm in Matches() — do NOT route bools
  through the wildcard text path.
  Note: WAVG/WSTDEV/WMEDIAN/SUMPRODUCTIFS inherit the same fix via IsNumericCell.
```

```
BUG-MEDIUM-004
File: /home/user/plugin/src/ExcelPerfToolkit/ConditionalAggregates.cs:834-851
Category: logic_error
Description: Once a criterion is textual, Matches() stringifies ANY non-blank,
  non-error cell with ToStringSafe and runs the wildcard/ordinal comparison on the
  string form. Numeric and bool cells therefore match text criteria. Native Excel
  type-segregates: wildcards and text comparisons match text cells only — COUNTIF
  (range,"*") is the canonical "count text entries" idiom and never counts numbers.
Trigger condition: A1:A4 = {5, "x", TRUE, 7}. =EPT.COUNTIFS(A1:A4, "*") returns 4;
  native COUNTIFS returns 1. Likewise "<b" matches the numeric cell 5
  (string.Compare("5","b",OrdinalIgnoreCase) = -13, empirically verified), and "1*"
  matches the number 12 (wildcard port verified: '1*' vs "12" = true, '*' vs "5" =
  true, '*' vs "TRUE" = true).
Trace (criterion "*", cell 5.0):
  1. Parse("*"): no prefix, operand "*" non-numeric → text criterion, tokens=[Star].
  2. Matches(5.0): not error/blank; _isNumeric=false → line 834 text =
     ToStringSafe(5.0) = "5".
  3. Line 835-838: Op.Equal → WildcardMatch([Star], "5") = true (verified) → MATCH.
Fix (Matches, text branch — insert before line 834):
  before:
      var text = Marshaling.ToStringSafe(cell);
  after:
      if (cell is not string textCell)
      {
          // Text criteria match text cells only; "<>text" is satisfied by any
          // non-text cell, mirroring the numeric branch at lines 816-821.
          return _op == Op.NotEqual;
      }
      var text = textCell;
  (Also resolves MEM-6: no more per-cell ToStringSafe allocation in this loop.)
```

```
BUG-MEDIUM-005
File: /home/user/plugin/src/ExcelPerfToolkit/LookupBoost.cs:87-114
Category: data_corruption
Description: ExactLookup canonicalizes every key and probe through
  Marshaling.ToStringSafe into a Dictionary<string,int>. This collapses distinct Excel
  types onto one keyspace: numeric 5.0 → "5" collides with text "5"; bool TRUE → "TRUE"
  collides with text "TRUE"; error cells → "#N/A"/"#DIV/0!" collide with text of the
  same spelling; ExcelEmpty and "" both → "". Native XLOOKUP exact match never matches
  a number to a text cell, never matches a text "#N/A" to an error, and propagates an
  error lookup value as that error.
Trigger condition: lookup table keys = {5 (number), "apple"}; lookup value = "5"
  (text, e.g. imported CSV). EPT.XLOOKUPB returns the row of numeric 5; native XLOOKUP
  returns #N/A. Conversely an #N/A error in lookup_values silently resolves to a row
  if the table also contains an error cell, instead of propagating #N/A.
Trace:
  1. Index build (lines 92-99): keys[0]=5.0 → ToStringSafe → "5" → index["5"]=0.
  2. Probe (line 108): lookupBlock[0,0]="5" (string) → key "5" → TryGetValue hit →
     returns returns[0]. Native: type mismatch → #N/A.
Fix: replace string canonicalization with a type-tagged key (see ARCH-3):
  before:
      var index = new Dictionary<string, int>(keys.Length, StringComparer.OrdinalIgnoreCase);
      ... var key = Marshaling.ToStringSafe(keys[i]); ...
  after:
      readonly struct CellKey : IEquatable<CellKey>  // tag: 0=num,1=text,2=bool,3=err,4=blank
      {
          public readonly byte Tag; public readonly double Num; public readonly string? Text;
          // Equals: tag must match; num by ==; text by OrdinalIgnoreCase.
      }
      var index = new Dictionary<CellKey, int>(keys.Length);
  Numbers stay binary-exact (5 and 5.0 still agree — both are double 5.0), text stays
  case-insensitive, types never cross. Propagate ExcelError lookup values directly:
      if (lookupBlock[r, c] is ExcelError e) { result[r, c] = e; continue; }
```

```
BUG-MEDIUM-006
File: /home/user/plugin/src/ExcelPerfToolkit/ConditionalAggregates.cs:476-483 (with 1061-1069)
Category: logic_error (silent error masking)
Description: When a matched row's VALUE cell holds an ExcelError, the numeric
  extraction loop silently skips it (IsNumericCell → TryToDouble's `case ExcelError:
  return false`), and SUMIFS/AVERAGEIFS/MINIFS/... return a clean number. Native Excel
  propagates an error from a matched cell in the sum/average range (SUMIFS over a
  matched #DIV/0! cell returns #DIV/0!; errors in NON-matching rows are ignored —
  which the mask already handles correctly). The doc at lines 84-85 says "Non-numeric
  cells in the sum range are ignored, as in Excel" — Excel ignores text/logicals, not
  errors. (Confidence: high for SUMIF/SUMIFS error propagation on matched rows;
  Excel-side behavior not re-verifiable in this environment.)
Trigger condition: B2 = #DIV/0! (a failed upstream formula), row 2 matches the
  criteria. =EPT.SUMIFS(B:B, A:A, "x") returns the sum of the other rows; native
  SUMIFS returns #DIV/0!. The broken input is invisible downstream.
Trace:
  1. AggregateIfs: mask[1]=true → matched contains ExcelError.ExcelErrorDiv0 (boxed).
  2. Aggregate("sum") line 479: IsNumericCell(ExcelErrorDiv0) → TryToDouble case
     ExcelError → false → skipped. Sum excludes the errored row. No error surfaces.
Fix (Aggregate, before the numeric extraction at line 476):
      foreach (var c in cells)
      {
          if (c is ExcelError err)
          {
              return err;   // mirror native propagation for value-range aggregates
          }
      }
  Apply to the numeric aggregate path only (count/first/last already define their own
  skip semantics); same check in MatchedNumbers (575-588), WeightedAverage (606),
  SumProductIfs (342), WMedian (246), WStdev (289/303).
```

```
BUG-MEDIUM-007
File: /home/user/plugin/src/ExcelPerfToolkit/LookupBoost.cs:61, 74
Category: logic_error (API semantics trap)
Description: `var approximate = matchMode is bool mode && !mode;` — FALSE selects
  APPROXIMATE match. This is the exact inverse of VLOOKUP's range_lookup (FALSE =
  exact) — the muscle memory of every Excel user — and incompatible with XLOOKUP's
  numeric match_mode (0 = exact, -1 = next-smaller), which this function's name
  evokes. Any numeric mode (0, -1, 1) is silently ignored: `matchMode is bool` fails
  for double 0.0, so the call quietly runs EXACT mode.
Trigger condition: =EPT.XLOOKUPB(A2:A100, T!A:A, T!B:B, , FALSE) written by a VLOOKUP
  user expecting exact match → runs approximate: misses return the next-smaller row's
  value instead of #N/A — plausible-looking wrong joins. =EPT.XLOOKUPB(..., -1)
  expecting XLOOKUP next-smaller → silently exact.
Trace:
  1. matchMode = false (bool) → approximate=true → ApproximateLookup: a probe of 7
     against keys {5, 10} returns the value at key 5 instead of #N/A.
  2. matchMode = -1.0 (double) → `is bool` fails → approximate=false → exact.
Fix (align with XLOOKUP and keep TRUE/FALSE only as documented legacy):
  before:
      var approximate = matchMode is bool mode && !mode;
  after:
      bool approximate;
      if (matchMode is bool mode) { approximate = !mode; }            // documented form
      else if (Marshaling.TryToDouble(matchMode, out var mm)) { approximate = mm == -1d; } // XLOOKUP -1
      else { approximate = false; }
  Strongly consider a breaking change to match VLOOKUP/XLOOKUP conventions outright;
  at minimum the divergence must be called out in the function Description string.
```

```
BUG-LOW-008
File: /home/user/plugin/src/ExcelPerfToolkit/ConditionalAggregates.cs:389-399
Category: data_corruption
Description: GROUPBY builds composite keys by joining cell strings with '\u001f'
  (line 397) with no escaping. Two distinct composite keys whose cells contain the
  separator collide: ("a\u001f", "b") and ("a", "\u001fb") both serialize to
  "a\u001f\u001fb\u001f" and are merged into one group with a combined aggregate.
Trigger condition: key cells containing U+001F (binary imports, control-char-laden
  CSV). Rare but silent when it happens.
Trace: row1 = ["a\u001f","b"] → "a\u001f"+SEP+"b"+SEP; row2 = ["a","\u001fb"] →
  "a"+SEP+"\u001fb"+SEP; both = "a\u001f\u001fb\u001f" → groups.TryGetValue hit on
  row2 → values merged.
Fix: length-prefix each segment so the encoding is injective:
  before:
      sb.Append(Marshaling.ToStringSafe(cell)).Append('\u001f');
  after:
      var s = Marshaling.ToStringSafe(cell);
      sb.Append(s.Length).Append(':').Append(s);
```

```
BUG-LOW-009
File: /home/user/plugin/src/ExcelPerfToolkit/ConditionalAggregates.cs:386-387, 397, 444
Category: logic_error
Description: Two inconsistencies in key identity. (a) GROUPBY uses
  StringComparer.Ordinal (line 386) — case-SENSITIVE — while DISTINCTCOUNTIFS uses
  OrdinalIgnoreCase (line 444): "Apple" and "apple" are two GROUPBY groups but one
  distinct value, and both diverge from Excel's case-insensitive text equality. (b)
  Both stringify cells via ToStringSafe, so numeric 5 and text "5" (and TRUE vs
  "TRUE") merge into one key — the same type conflation as BUG-MEDIUM-005.
Trigger condition: key column = {"Apple", "apple"} → GROUPBY spills 2 rows;
  DISTINCTCOUNTIFS over the same data counts 1. Key column = {5, "5"} → one merged
  group.
Trace: GroupBy line 386 comparer Ordinal → "Apple\u001f" != "apple\u001f" → 2 groups.
  Aggregate "distinct" line 444 OrdinalIgnoreCase → set.Count = 1.
Fix: pick one policy (OrdinalIgnoreCase matches Excel text semantics) and add a type
  tag per segment (or adopt ARCH-3's CellKey):
      sb.Append(TypeTag(cell)).Append(s.Length).Append(':').Append(s.ToUpperInvariant());
```

```
BUG-LOW-010
File: /home/user/plugin/src/ExcelPerfToolkit/ConditionalAggregates.cs:636-640
Category: logic_error
Description: BuildMask validates criteria ranges by total cell count only
  (`range.Length != n`). A 2×3 criteria range against a 3×2 value range (both 6
  cells) is accepted and matched positionally in row-major order. Native Excel
  requires identical shape and returns #VALUE!.
Trigger condition: =EPT.SUMIFS(B1:B6, A1:C2, "x") — shape mismatch silently produces
  a row-major-aligned (i.e., scrambled) match instead of #VALUE!.
Trace: FlattenRowMajor(A1:C2).Length = 6 = n → no throw → mask built with
  transposed pairing.
Fix: thread the (rows, cols) of the value range into BuildMask and compare both:
  before:
      if (range.Length != n)
  after:
      var block = Marshaling.AsArray2D(pairs[p]);
      if (block.GetLength(0) != valueRows || block.GetLength(1) != valueCols)
          throw new ArgumentException("Criteria range shape must match the value range.");
```

```
BUG-LOW-011
File: /home/user/plugin/src/ExcelPerfToolkit/LookupBoost.cs:121-129, 153-170
Category: logic_error (nondeterministic-vs-spec result)
Description: ApproximateLookup sorts (Key, Index) tuples comparing Key only.
  List<T>.Sort is unstable (empirically verified: after sorting 64 duplicate keys,
  original order was NOT preserved — the first element ended up last). FloorIndex
  returns the LAST entry ≤ target in sorted order, so for duplicate keys the winning
  row is an arbitrary artifact of introsort's partitioning, not "first" or "last"
  occurrence. Native VLOOKUP/XLOOKUP approximate match has a defined answer
  (last-in-range for the classic binary-search behavior).
Trigger condition: lookup table keys {10, 10, 10} with returns {a, b, c}; probe 10 →
  returns whichever duplicate introsort left in the highest slot (verified: input
  order scrambled).
Trace: points.Sort(a.Key.CompareTo(b.Key)) → duplicates permuted → FloorIndex picks
  points[last-equal].Index → arbitrary among {0,1,2}.
Fix: make the comparator total so "last occurrence in the original range" wins:
  before:
      points.Sort(static (a, b) => a.Key.CompareTo(b.Key));
  after:
      points.Sort(static (a, b) =>
      {
          var c = a.Key.CompareTo(b.Key);
          return c != 0 ? c : a.Index.CompareTo(b.Index);
      });
```

```
BUG-LOW-012
File: /home/user/plugin/src/ExcelPerfToolkit/ConditionalAggregates.cs:800-803, 737-752
Category: logic_error
Description: Matches() rejects every ExcelError cell unconditionally (lines 800-803),
  and Parse turns an error criterion (or "#N/A" text) into an ordinary text criterion.
  Therefore no criterion can ever match an error cell. Native COUNTIF supports
  counting errors by their text form (=COUNTIF(rng,"#DIV/0!") counts #DIV/0! cells).
Trigger condition: =EPT.COUNTIFS(A1:A10, "#N/A") over a column containing #N/A errors
  returns 0; native returns the error count.
Trace: Parse("#N/A") → text criterion tokens for "#N/A". Matches(#N/A error cell) →
  line 800-803 `cell is ExcelError` → false before the text path is reached.
Fix: in Parse, recognize the eight error spellings (compare against
  Marshaling.ErrorToText forms) and store an _errorValue; in Matches, handle errors
  first:
      if (cell is ExcelError cellErr)
      {
          return _isError && _op == Op.Equal ? cellErr == _errorValue
               : _isError && _op == Op.NotEqual ? cellErr != _errorValue
               : false;
      }
```

```
BUG-LOW-013
File: /home/user/plugin/src/ExcelPerfToolkit/ConditionalAggregates.cs:737-752
Category: logic_error (silent wrong result)
Description: If a multi-cell range arrives as the criteria argument (object[,]),
  Parse falls through every case to Marshaling.ToStringSafe(criteria), whose default
  branch yields "System.Object[,]". The criterion becomes the literal text
  "System.Object[,]" and matches nothing — the function returns 0/empty with no
  error. Native dynamic-array Excel spills one result per criteria element (or
  implicitly intersects in legacy mode).
Trigger condition: =EPT.COUNTIFS(A:A, B1:B3) → always 0, silently.
Trace: criteria = object[2..,1..] → not null/bool/number/string → ToStringSafe →
  `_ => value.ToString()` → "System.Object[,]" → tokenized → matches nothing.
Fix (minimum viable — fail loudly; spilling per-element is a feature decision):
      case object[,] block when block.Length == 1:
          return Parse(block[0, 0]);
      case object[,]:
          throw new ArgumentException("criteria must be a single value, not a multi-cell range.");
```

---

## Pass 2 — Memory optimization

Representative call: `=EPT.SUMIFS(B1:B100000, A1:A100000, ">5")`, ~50% match rate.
Current allocation per call ≈ **2.9 MB** (excluding Excel-DNA's own marshaling);
achievable ≈ **0.1 MB** (the mask) — ~29× reduction.

```
MEM-1
File: ConditionalAggregates.cs:662-681 (FlattenRowMajor), call sites 68, 236-237,
      280-281, 332-333, 561, 577, 592-593, 636
Current cost: one object?[] copy per range argument: 24 + 100000*8 ≈ 800 KB each.
  SUMIFS touches 2 ranges → 1.6 MB/call of pure reference copying; WAVGIFS/
  SUMPRODUCTIFS touch 3 → 2.4 MB.
Optimization: iterate the object[,] in place. The arrays come straight from Excel-DNA
  and are never mutated; row-major order is preserved by a nested loop with a running
  flat index.
Expected reduction: 1.6-2.4 MB/call → 0.
Before:
    var values = FlattenRowMajor(Marshaling.AsArray2D(valueRange));
    var mask = BuildMask(values.Length, pairs);
    for (var i = 0; i < values.Length; i++) { if (mask[i]) matched.Add(values[i]); }
After:
    var block = Marshaling.AsArray2D(valueRange);
    int rows = block.GetLength(0), cols = block.GetLength(1);
    var mask = BuildMask(rows, cols, pairs);          // BuildMask itself iterates 2-D
    var i = 0;
    for (var r = 0; r < rows; r++)
        for (var c = 0; c < cols; c++, i++)
            if (mask[i]) Accumulate(block[r, c]);
```

```
MEM-2
File: ConditionalAggregates.cs:563 (matched list), 476 (nums list)
Current cost: `new List<object?>(values.Length)` pre-sizes to the FULL range — 800 KB
  backing array even when 1 row matches; Aggregate then materializes a second
  List<double> sized cells.Count (≤ 800 KB → 400 KB at 50% match). Up to 1.2 MB of
  intermediates whose only purpose is to be iterated once.
Optimization: stream matched cells directly into per-op accumulators (sum/count/min/
  max/product/geomean/harmean need O(1) state; median/percentile/mode/stdev need only
  the List<double> of numerics). See ARCH-4 for the structure.
Expected reduction: 1.2 MB → 0 for the O(1)-state ops (sum/count/avg/min/max/product/
  geomean/harmean/first/last); 1.2 MB → 0.4 MB for the order-statistics ops.
Before:
    var matched = new List<object?>(values.Length);
    ... return Aggregate(op, matched);
After (sum sketch):
    double sum = 0d;
    ForEachMatched(block, mask, cell => { if (IsNumericCell(cell, out var d)) sum += d; });
    return sum;
```

```
MEM-3
File: ConditionalAggregates.cs:68
Current cost: `FlattenRowMajor(Marshaling.AsArray2D(criteriaRange1)).Length` allocates
  an 800 KB copy of the first criteria range purely to read its length — then
  BuildMask flattens the SAME range again at line 636 (another 800 KB).
Optimization: compute the count from the dimensions.
Expected reduction: 800 KB/call → 0 (plus one fewer 100k-element pass).
Before:
    var n = FlattenRowMajor(Marshaling.AsArray2D(criteriaRange1)).Length;
After:
    var block = Marshaling.AsArray2D(criteriaRange1);
    var n = block.GetLength(0) * block.GetLength(1);   // both ≤ 1,048,576 / 16,384: no overflow
```

```
MEM-4
File: ConditionalAggregates.cs:392-404
Current cost: GROUPBY allocates `new object?[keyCols]` for EVERY row (line 392) but
  retains it only for first-seen keys (line 404). 100k rows, 1 key column, 1k groups:
  ~99k wasted arrays × 32 B = 3.2 MB garbage per call (64 B/row for 5 key columns →
  6.4 MB).
Optimization: build the StringBuilder key from keys[r,c] directly; copy the row only
  on a dictionary miss.
Expected reduction: rows×(24+8·keyCols) → groups×(24+8·keyCols); 3.2 MB → 32 KB at
  1k groups.
Before:
    var keyRow = new object?[keyCols];
    for (var c = 0; c < keyCols; c++) { var cell = keys[r, c]; keyRow[c] = cell; sb.Append(...); }
    ...
    if (!groups.TryGetValue(key, out var list)) { ...; keyCells[key] = keyRow; ... }
After:
    for (var c = 0; c < keyCols; c++) { sb.Append(Marshaling.ToStringSafe(keys[r, c])).Append('\u001f'); }
    var key = sb.ToString();
    if (!groups.TryGetValue(key, out var list))
    {
        var keyRow = new object?[keyCols];
        for (var c = 0; c < keyCols; c++) { keyRow[c] = keys[r, c]; }
        ...
    }
```

```
MEM-5
File: LookupBoost.cs:91-99 (index build), 104-113 (probes), 66-67 (Flatten)
Current cost (measured): stringifying 100k doubles via ToStringSafe's "R" format
  allocates 6.30 MB (measured with GC.GetAllocatedBytesForCurrentThread). Index build
  + probe side for a 100k×100k numeric join ≈ 12.6 MB of transient strings, plus
  1.6 MB for the two Flatten copies, plus ~2.8 MB for the Dictionary itself — ~17 MB
  per call, ~14 MB avoidable.
Optimization: type-tagged CellKey comparer (ARCH-3 / BUG-MEDIUM-005 fix) — numeric
  keys hash the double directly; only genuine text cells keep their (existing) string
  reference. Drop the Flatten copies by indexing the 2-D arrays.
Expected reduction: ~17 MB → ~3 MB per call for numeric tables (5-6×); zero string
  garbage on the probe side.
Before:
    var key = Marshaling.ToStringSafe(keys[i]);
After:
    var key = CellKey.From(keys[i]);   // double → (tag=Num, value); string → (tag=Text, s)
```

```
MEM-6
File: ConditionalAggregates.cs:834
Current cost: with a text criterion over a mixed/numeric column, Matches() calls
  ToStringSafe on every non-string cell: ~6.3 MB per 100k numeric cells (same
  measurement as MEM-5), discarded immediately.
Optimization: the BUG-MEDIUM-004 fix (text criteria only inspect string cells) makes
  this zero by construction.
Expected reduction: 6.3 MB/100k numeric cells → 0.
Before/After: see BUG-MEDIUM-004.
```

---

## Pass 3 — CPU / throughput

```
PERF-1
File: ConditionalAggregates.cs:642-648 (mask loop), 798-852 (Matches), 1061-1069
Hot path: YES (every cell of every criteria range, every call).
Current cost: per cell, a numeric criterion pays: sealed-method call → IsBlank (4
  type tests + string length) → IsNumericCell → TryToDouble (10-branch type switch)
  → _op switch. ~15-25 ns/cell vs ~2-3 ns for `cell is double d` + one compare.
Optimization: hoist the criterion kind out of the loop and run specialized loops; the
  numeric loop pattern-matches double directly (bool excluded per BUG-MEDIUM-003):
Before:
    for (var i = 0; i < n; i++)
        if (mask[i] && !criterion.Matches(range[i])) mask[i] = false;
After:
    if (criterion.IsSimpleNumeric)
    {
        for (var i = 0; i < n; i++)
        {
            if (!mask[i]) continue;
            mask[i] = range[i] is double d && criterion.CompareNumber(d); // inlined op
        }
    }
    else { /* existing general loop */ }
Expected speedup: 2-3× on the mask-build phase (typically ~half the call) → ~1.5×
  end-to-end for COUNTIFS/SUMIFS on numeric data.
```

```
PERF-2
File: ConditionalAggregates.cs:559-572, 434-556
Hot path: YES.
Current cost: a SUMIFS call makes 4 sequential passes over ~100k elements (flatten
  values, flatten+match criteria, copy matched, extract numerics) plus the aggregate
  pass, touching ~3 MB of freshly allocated memory (cache-hostile).
Optimization: fuse into mask build + one accumulation pass over the original block
  (MEM-1/MEM-2 are the allocation half of this change; ARCH-4 is the structure).
Expected speedup: 1.5-2× end-to-end on 100k-row calls (fewer passes + GC pressure:
  2.9 MB/call at recalc fan-out of 8 MTR threads is ~23 MB/s of gen0 churn per
  100ms recalc wave).
Before → After: see MEM-1/MEM-2 snippets.
```

```
PERF-3
File: LookupBoost.cs:121-129
Hot path: YES (index build phase of every approximate call).
Current cost (measured): sorting 100k (double,int) tuples via the Comparison<T>
  delegate = 47.6 ms. A sortedness pre-scan = 0.21 ms (~230× cheaper). The function's
  own doc (lines 40-42) says approximate mode "assumes they are sorted ascending" —
  the dominant real-world input is already sorted.
Optimization: O(M) presorted check; sort only when needed. Also sort an array
  (Array.Sort with struct comparer) instead of List+lambda to dodge delegate dispatch.
Expected speedup: presorted input: build phase 47.6 ms → 0.21 ms; whole 100k-probe
  call ~60 ms → ~12 ms ≈ 5×. Unsorted input: ~1.3× from the struct comparer alone.
Before:
    points.Sort(static (a, b) => a.Key.CompareTo(b.Key));
After:
    var sorted = true;
    for (var i = 1; i < points.Count; i++)
        if (points[i - 1].Key > points[i].Key) { sorted = false; break; }
    if (!sorted) points.Sort(static (a, b) => ...);   // with BUG-LOW-011's tiebreak
```

```
PERF-4
File: LookupBoost.cs:91-99
Hot path: YES (exact-match index build).
Current cost (measured): string-key dictionary build over 100k numeric keys = 32.5 ms
  (ToString("R") 22.1 ms of it); double-key dictionary = 10.3 ms → 3.2×. Within the
  dictionary insert itself, ContainsKey+indexer does two hash probes per key:
  measured 8.94 ms vs TryAdd 3.79 ms (2.4×, warmed, 7-run average).
Optimization: CellKey typed dictionary (ARCH-3) + TryAdd which also preserves
  first-occurrence semantics exactly.
Expected speedup: index build ~32.5 ms → ~5 ms (≈6×) for numeric tables; probe side
  similar (no ToStringSafe per probe).
Before:
    if (!index.ContainsKey(key)) { index[key] = i; }
After:
    index.TryAdd(key, i);
```

```
PERF-5
File: ConditionalAggregates.cs:935-942 (Median), 944-965 (PercentileInc)
Hot path: YES for MEDIANIFS/PERCENTILEIFS over large matches.
Current cost (measured): Array.Sort of 100k doubles = 29.3 ms, O(n log n), plus an
  800 KB nums.ToArray() copy. Median/percentile need only 1-2 order statistics.
Optimization: quickselect (Floyd-Rivest or simple 3-median-of-medians pivot) on a
  span over the list's backing array: O(n) expected, ~2-3 ms at 100k; no ToArray copy
  (CollectionsMarshal.AsSpan(nums)).
Expected speedup: ~10× on the selection phase; eliminates the 800 KB copy.
Before:
    var arr = nums.ToArray();
    Array.Sort(arr);
    ... arr[mid] ...
After:
    var span = System.Runtime.InteropServices.CollectionsMarshal.AsSpan(nums);
    var lo = QuickSelect(span, mid);          // and (mid-1) for the even case / hi rank
```

```
PERF-6
File: ConditionalAggregates.cs:287-308 (WStdev)
Hot path: yes when used at scale.
Current cost: pass 1 calls IsNumericCell on v[i] AND w[i]; pass 2 repeats both — 4
  TryToDouble dispatches per row where 2 suffice (~40 ns/row wasted → ~4 ms/100k).
Optimization: collect (x, wt) into a List<(double,double)> during pass 1; pass 2
  iterates the typed list.
Expected speedup: ~1.7× for WSTDEV (plus it makes the two passes provably see the
  same row set).
Before:
    for (...) { if (IsNumericCell(v[i], out var x) && IsNumericCell(w[i], out var wt) && wt > 0d) { sumW += wt; sumWx += wt * x; } }
    ... second identical filter loop ...
After:
    var rowsKept = new List<(double X, double W)>();
    // pass 1 fills rowsKept and sums; pass 2 iterates rowsKept only.
```

```
PERF-7
File: LookupBoost.cs:87-99
Hot path: YES in the realistic misuse mode (formula copied down → R=1 per call).
Current cost: a 1-probe call against a 100k-row table still builds the full index:
  32.5 ms + ~14 MB (measured, MEM-5/PERF-4) versus a linear scan with early-out:
  ~0.5 ms average, 0 alloc. A column of 100k such formulas → ~54 minutes of CPU and
  ~1.4 TB of allocation traffic across the recalc.
Optimization: probe-count heuristic — below a threshold, scan instead of index.
Expected speedup: ~60× per single-probe call.
Before:
    var index = new Dictionary<string, int>(keys.Length, ...); // always
After:
    var probeCount = lookupBlock.Length;
    if (probeCount * (long)keys.Length <= 4 * keys.Length)   // probeCount <= 4
        return LinearLookup(lookupBlock, keys, returns, notFound);
```

```
PERF-8
File: ConditionalAggregates.cs:419, 434-437
Hot path: yes for high-cardinality GROUPBY.
Current cost: Aggregate() runs `(operation ?? "").Trim().ToLowerInvariant()` plus the
  string switch once PER GROUP (line 419 in the result loop): 100k groups → ~100k
  string allocations (~4 MB) and ~5 ms, and an unknown-op error only surfaces after
  the entire grouping pass.
Optimization: parse the operation to an enum once before the row loop; pass the enum.
Expected speedup: ~1.05-1.1× for 100k-group calls; unknown ops fail fast.
Before:
    result[i, keyCols] = Aggregate(operation, groups[key]);
After:
    var op = ParseOp(operation);            // once, before the r-loop; throws early
    ...
    result[i, keyCols] = Aggregate(op, groups[key]);
```

---

## Pass 4 — Architectural wins

```
ARCH-1
Scope: ConditionalAggregates.cs — entire *IFS family.
Current pattern: each formula instance scans its M-cell ranges once: a column of R
  formulas is O(R·M) compute and R range reads, mitigated only by MTR fan-out.
Proposed pattern: batched variants mirroring XLookupB's shape: EPT.SUMIFS.BATCH(
  sum_range, criteria_range, criteria_vector) → spilled column. For equality criteria,
  one dictionary-aggregation pass groups sums by key (O(M)), then probes per criterion
  (O(R)): O(M+R) total. For comparison criteria, sort once (O(M log M)) + prefix sums
  → O(log M) per probe.
Impact estimate: R=10k formulas over M=100k rows: 10^9 cell-evaluations → ~2×10^5:
  ~5000× compute reduction, R× fewer boundary reads (the readme's own bottleneck #7
  analysis, actually solved rather than parallelized).
Effort: medium (new UDFs; reuses Criterion + Aggregate).
Risk: low — new functions, no behavior change to existing ones.
```

```
ARCH-2
Scope: LookupBoost.cs — ExactLookup/ApproximateLookup index build.
Current pattern: the index is rebuilt on every call; a workbook with 20 XLOOKUPB
  formulas against the same 100k-row table pays 20 index builds per recalc.
Proposed pattern: process-wide ConcurrentDictionary<ulong, CachedIndex> keyed on a
  content hash of the key+return columns (XXH3 already exists per readme's
  EPT.HASHBLOCK), with size-capped LRU eviction. Content hashing makes staleness
  structurally impossible — Excel does NOT notify UDFs on range edits, so any
  identity-keyed (sheet/address) cache would serve stale indexes; hashing the actual
  cell contents (O(M), ~0.05 ms/100k doubles with XXH3) sidesteps that hazard
  entirely while still amortizing the expensive part (dictionary/sort build, string
  materialization).
Impact estimate: per-call cost for cached hits drops from ~33 ms+14 MB to ~1 ms hash
  + probes: 5-30× per recalc wave for multi-formula workbooks.
Effort: medium. Risk: medium — memory retention (cap + LRU mandatory), and the cache
  must be lock-free/concurrent because IsThreadSafe=true means MTR threads race on it.
```

```
ARCH-3
Scope: LookupBoost.cs ExactLookup; ConditionalAggregates.cs GROUPBY (386-407) and
  Aggregate "distinct" (442-454).
Current pattern: all three canonicalize cells to strings via ToStringSafe — the root
  cause of BUG-MEDIUM-005 and BUG-LOW-009, and of MEM-5/MEM-6's 6.3 MB-per-100k-cells
  string churn. (Marshaling.CellEquality is NOT a drop-in replacement: its TryToDouble
  path coerces numeric STRINGS to numbers, reintroducing the same conflation.)
Proposed pattern: one shared readonly struct CellKey { byte Tag; double Num; string?
  Text; } with tag ∈ {Number, Text(OrdinalIgnoreCase), Bool, Error, Blank}, equality
  and hash per tag. Use it as the dictionary key everywhere a cell is an identity.
Impact estimate: fixes the cross-type-collision bug family in one place; removes
  ~12.6 MB/call of strings in XLOOKUPB (measured); 3-6× index/groupby build speedup
  (measured 10.3 ms vs 32.5 ms dictionary builds).
Effort: small-medium (one struct + three call sites). Risk: low; behavior change is
  precisely the bug fix (document that 5 no longer equals "5").
```

```
ARCH-4
Scope: ConditionalAggregates.cs Aggregate/AggregateIfs (434-588) + all callers.
Current pattern: parse-op-string → materialize matched object list → materialize
  numeric list → aggregate: 4 passes, ~2.9 MB transient per 100k-row call, op string
  re-parsed per GROUPBY group.
Proposed pattern: enum AggOp parsed once + a streaming accumulator struct per op
  (Count/Sum/Min/Max/Product/GeoMean/HarMean/First/Last = O(1) state; Stdev/Var via
  Welford or mean+M2; Median/Percentile/Mode collect List<double> only). AggregateIfs
  feeds cells straight from the object[,] through the mask into the accumulator.
Impact estimate: umbrella for MEM-1/2, PERF-1/2/8: ~29× allocation reduction and
  1.5-2× throughput on every *IFS call; under 8-thread MTR recalc the gen0 pressure
  drop is the difference between continuous background GC and none.
Effort: medium (mechanical, well-tested surface). Risk: low — pure refactor with
  identical numeric semantics (sum order preserved by row-major iteration).
```

---

## Seam notes

- `Marshaling.cs:127-132` — `TryToDouble`'s CurrentCulture fallback makes numeric
  criterion parsing locale-dependent (`">1,5"` is numeric >1.5 on fr-FR, text on
  en-US); this mirrors Excel's own locale behavior but means workbook results are
  machine-dependent — worth one line in the function docs.
- `Marshaling.cs:363-367` — `CellEquality` coerces numeric-looking strings to numbers
  (`"5"` == 5.0), so it must not be adopted as the fix for the Domain-6 type-conflation
  bugs without changing that contract.
- `Marshaling.cs:155` — `ToStringSafe` renders large doubles in scientific form
  (`1e17 → "1E+17"`, empirically verified); any consumer using it for display-form
  keys diverges from Excel's General format.

---

## Rejection appendix

1. **MTR race conditions / shared mutable state** — exhaustively scanned both files:
   the only statics are `TraceSource` instances (thread-safe per BCL contract, no
   listeners configured by default) and the `Marshaling.CellEquality` singleton
   (immutable). Every UDF is pure over its arguments. No finding.
2. **`params object[] morePairs` receiving trailing `ExcelMissing` padding** (which
   would make `pairs.Length % 2 != 0` throw on every call) — Excel-DNA's params
   expansion strips trailing missing arguments when building the array; if it did
   not, no 2-argument `EPT.COUNTIFS` call could ever succeed, which would have been
   caught immediately. Rejected.
3. **Numeric-looking text excluded from numeric criteria/value ranges** — native
   COUNTIF famously coerces text-numbers (the 15-digit COUNTIF problem), so this IS
   a native divergence; however the audit spec and the file's documented contract
   both mandate "numeric comparisons match only genuinely numeric cells". Treated as
   intended design, not a bug.
4. **NaN/Infinity poisoning accumulators** — `TryToDouble` rejects NaN/±Inf doubles
   at ingestion (Marshaling.cs:85), and Excel cells cannot hold them anyway; sums/
   products can still overflow to +Inf, which Excel renders as an error — same
   observable class as native `#NUM!` on overflow. Rejected.
5. **`Mode` Dictionary<double,int> +0.0/−0.0 split** — empirically verified: the
   dictionary merges them (Count = 1, hash codes equal). Matches Excel (0 = -0).
   Rejected.
6. **`rows*cols` int overflow in flatten helpers** — both files compute
   `(long)rows * cols` and guard against > int.MaxValue (ConditionalAggregates.cs:
   666-670, LookupBoost.cs:176-180). Rejected.
7. **PercentileInc index math** — traced k=0 (rank 0 → arr[0]), k=1 (rank n-1, lo==hi
   → arr[n-1]), k=0.5/n=4 (rank 1.5 → (arr[1]+arr[2])/2): exactly PERCENTILE.INC's
   `k(n-1)` interpolation. NaN/out-of-range k guarded at line 946. Rejected.
8. **FloorIndex binary-search boundaries** — traced empty list (hi=-1 → -1 →
   notFound), target below min (never ≤ → -1), target above max (result=count-1),
   exact hit (lo=mid+1 continues to find the LAST ≤). Correct floor search. Rejected.
9. **Median off-by-one** — odd n → arr[n/2]; even n → (arr[mid-1]+arr[mid])/2.
   Correct. Rejected.
10. **`Guard` exception swallowing** — converts to #VALUE! with a TraceSource warning;
    pattern accepted in round 2 for the whole toolkit. The lost ArgumentException
    detail (e.g., GROUPBY's helpful unknown-op message) never reaches the user, but
    that is the toolkit-wide boundary discipline, not a Domain-6 defect. Rejected.
11. **WMedian unstable sort** — ties are by VALUE; whichever equal-valued pair crosses
    the half-weight threshold, the returned value is identical. Rejected (contrast
    BUG-LOW-011, where the returned VALUE differs).
12. **Mode tie-breaking (first-seen wins) vs Excel MODE.SNGL** — MODEIFS is documented
    as having no native equivalent; first-seen is a defensible self-contract.
    Rejected.
13. **`~` escaping any following character (`"~a"` → literal `a`)** — verified the
    tokenizer drops the tilde; Excel's documented escapes are only `~*`, `~?`, `~~`,
    and its handling of `~a` is itself ambiguous/undocumented. Impact ≈ nil. Rejected.
14. **Trailing `~` kept as a literal tilde** (Tokenize line 861: `i + 1 <
    pattern.Length` fails → falls to the literal branch) — plausible reading of
    Excel's behavior; unverifiable here; impact ≈ nil. Rejected.
15. **`Criterion.Parse` ignoring TryToDouble's return for double/DateTime criteria
    (line 748)** — a NaN criterion would yield `_number=NaN` (matches nothing) and a
    year-0001 DateTime yields 0, but Excel cannot deliver either through `object`
    UDF arguments (cells hold finite doubles; dates arrive as serials). Unreachable.
    Rejected.
16. **`""` vs `"="` criteria both matching empty-string text** — native Excel
    distinguishes `"="` (truly blank only) from `""` (blank or zero-length text);
    `IsBlank` (line 1053-1054) merges them. Real but third-order: zero-length-string
    cells only arise from formula results pasted as values. Rejected as a finding;
    noting here for completeness.
17. **`"<>"` (non-blank) counting zero-length-string cells** — same `IsBlank` merge as
    #16, same rarity. Rejected.
18. **`Aggregate("count")` counting blank/text/error cells in GROUPBY** — "count" in
    GROUPBY counts group ROWS (COUNTA-like), a defensible documented-by-example
    contract; EPT.COUNTIFS itself counts mask hits, which matches native COUNTIFS.
    Rejected.
19. **`ThreadAbortException` in IsCritical (both files)** — dead code on .NET 8
    (never thrown), harmless. Rejected.
20. **`double k` / typed parameters receiving Excel coercions** — Excel coerces or
    errors before the UDF runs for register-typed double args; omitting k yields 0.0
    → k=0 → minimum, mildly surprising but native PERCENTILE requires k at the
    formula level. Rejected.
21. **BuildMask not short-circuiting** — it already skips `Matches` once `mask[i]` is
    false (line 644 `mask[i] && ...`). Rejected as a PERF finding (already done).
22. **WAVG allowing negative weights** (vs WMEDIAN/WSTDEV requiring wt > 0) —
    explicitly documented difference in the XML docs (lines 198-204 vs 225-231).
    Rejected.
23. **`points.Sort` Comparison-delegate overhead** — real but folded into PERF-3's
    struct-comparer note rather than a standalone finding.
24. **Approximate mode ignoring text keys entirely** (native VLOOKUP-approximate
    handles text ordering) — explicitly documented: "operates over the numeric keys"
    (LookupBoost.cs:40-42). Self-contract. Rejected.
25. **`ExcelEmpty`/blank probe matching the first blank table row in ExactLookup** —
    subsumed by BUG-MEDIUM-005's typed-key fix (Blank gets its own tag; policy for
    blank probes becomes explicit). Not double-reported.
26. **Strings > 32767 chars returned to Excel** — cells cannot contain them on input,
    and first/last return original cell values; GROUPBY returns original key cells.
    Unreachable. Rejected.
27. **`string.Compare(..., OrdinalIgnoreCase)` vs Excel's culture-ish text ordering
    for `>`/`<` text criteria** — real divergence for diacritics/locale collation,
    but text ordering criteria are rare and the case-folding behavior (the common
    case) is correct; also partially subsumed by BUG-MEDIUM-004 (numbers leaking into
    this comparison is the harmful part). Rejected as a separate finding.
28. **GROUPBY `Dictionary` without capacity hint** — groups dictionary growth causes
    ~17 rehashes at 100k distinct keys; sizing to keyRows would over-allocate for
    low-cardinality data. Judgment call, no clear win. Rejected.
```
