# Implementation Plan: Bill of Entry Converter (Milestone 1)

## Overview

This plan implements the BOE-to-CTN-Excel converter in Python using FastAPI, `pdfplumber`, and
`openpyxl`, as specified in the design. Work proceeds bottom-up: project scaffolding and shared data
models first, then the pure logic (Value_Calculator, line-item assembly), then the PDF_Parser and
Excel_Generator, then the Conversion Orchestrator that sequences and cross-checks them, and finally the
Web_Interface. Property-based tests (Properties 1–17 from the design) and example/integration tests are
placed as sub-tasks next to the code they validate so errors surface early. Each step builds on the
prior ones and ends wired into the running system, leaving no orphaned code.

## Tasks

- [x] 1. Set up project structure, dependencies, and shared data models
  - [x] 1.1 Scaffold the project and define shared data models
    - Create the Python package layout (e.g. `boe_converter/` with `models.py`, `parser.py`,
      `calculator.py`, `excel_writer.py`, `orchestrator.py`, `validator.py`, `web/`) and a `tests/` tree
    - Add dependencies (`fastapi`, `uvicorn`, `pdfplumber`, `openpyxl`, `hypothesis`, `pytest`,
      `python-multipart`) via `pyproject.toml`/`requirements.txt`
    - Configure `pytest` and `hypothesis` (default profile: `max_examples=100`)
    - Implement the frozen dataclasses from the design Data Models: `RawValue`, `HeaderBlock`, `LineItem`,
      `ExtractedDocument`, `ComputedLine`, `Totals`, `ComputedDocument`, `ReviewFlag`, `Discrepancy`,
      `ConversionSummary`, plus `ReviewFlagSet` helper
    - _Requirements: 8.1_

- [x] 2. Implement the Value_Calculator (pure per-line and totals computation)
  - [x] 2.1 Implement per-line monetary computations
    - Implement `compute_line(item, usd_rate)` producing `amount_usd`, `purchase_inr`, `sws_amount`,
      `total_customs_duty`, `igst_amount`, `combined_duty`, `land_cost_excl_gst`, `land_cost_incl_gst`,
      and `purchase_rate_per_unit` per the normative formulas; set rate-per-unit to 0 when qty == 0;
      retain full floating-point precision (no rounding)
    - Implement the `pcs` rule: `qty * 12` only when the unit trimmed and upper-cased equals `"DOZ"`,
      otherwise blank
    - Leave any dependent value `None` and emit a `ReviewFlag` when a required input is missing/non-numeric
    - _Requirements: 6.1, 6.2, 6.3, 6.4, 6.5, 6.6, 6.7, 6.8, 6.9, 6.10, 6.12, 6.13, 5.8, 5.9_

  - [x]* 2.2 Write property test for per-line monetary computations
    - **Property 5: Per-line monetary computations satisfy their formulas and algebraic relations**
    - **Validates: Requirements 6.1, 6.2, 6.3, 6.4, 6.5, 6.6, 6.7, 6.8, 6.9, 6.10**

  - [x]* 2.3 Write property test for the pcs rule
    - **Property 6: pcs is QTY×12 exactly when the unit is DOZ, otherwise blank**
    - **Validates: Requirements 5.8, 5.9**

  - [x]* 2.4 Write property test for missing/non-numeric computation inputs
    - **Property 7: Missing or non-numeric inputs leave dependent values blank and flagged**
    - **Validates: Requirements 6.13**

  - [x] 2.5 Implement totals aggregation
    - Implement `compute_totals(lines, pkg_count)` and `compute(doc, usd_rate)` producing column-wise
      sums (USD amount, assessable value, total customs duty, IGST, both land-cost columns); every total
      is 0 when there are no lines; carry the package count through as a `RawValue`
    - _Requirements: 7.1, 7.2, 7.3, 7.5_

  - [x]* 2.6 Write property test for totals aggregation
    - **Property 8: Column totals equal the sums of their per-line values, zero when empty**
    - **Validates: Requirements 7.1, 7.2, 7.3**

- [x] 3. Checkpoint - Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise.

- [x] 4. Implement orientation-aware PDF extraction primitives and header parsing
  - [x] 4.1 Implement upright-word extraction and positional row reconstruction
    - Implement `PdfParser._upright_words(page)` to keep only characters with orientation == 0 (defeating
      rotated margin labels like `SLIATED`/`YTUD`/`SEITUD`) and reconstruct rows by bounding box
    - Provide a verbatim capture helper that records each field as `RawValue(raw_text=...)` with a
      separate non-destructive numeric parse, setting `is_missing`/`is_unparseable` flags
    - _Requirements: 2.11, 3.6_

  - [x]* 4.2 Write property test for verbatim capture round-trip
    - **Property 3: Extraction preserves printed values verbatim (capture round-trip)**
    - **Validates: Requirements 2.11, 3.6**

  - [x]* 4.3 Write orientation-robustness regression test
    - Assert `_upright_words` excludes the BOE's rotated margin labels so rotated text never contaminates
      extracted rows (uses the reference PDF)
    - _Requirements: 3.6_

  - [x] 4.4 Implement header-block extraction
    - Implement `_extract_header(pages)` to extract BE No/Date, Invoice No/Date, invoice amount + currency,
      package count, party name, container details, and B/L No/Date (where present); emit
      `ReviewFlag(header, field, MISSING)` for any required field that cannot be located/resolved; never
      substitute defaults
    - _Requirements: 2.1, 2.2, 2.3, 2.4, 2.5, 2.6, 2.7, 2.8, 2.9, 2.10_

  - [x]* 4.5 Write unit tests for header extraction against the reference PDF
    - Assert each header field equals the reference PDF's known value; assert a missing field yields a flag
    - _Requirements: 2.1, 2.2, 2.3, 2.4, 2.5, 2.6, 2.7, 2.8, 2.9, 2.10_

- [x] 5. Implement line-item extraction and two-pass assembly
  - [x] 5.1 Implement Part II (invoice) and Part III (duty) item extraction
    - Implement `_extract_invoice_items(pages)` → `{item_serial: invoice values}` (CTH, description,
      unit price, qty, UQC, amount) and `_extract_duty_items(pages)` → `{item_serial: assess value, BCD
      rate/amt, SWS amt, IGST rate, total duty}`; treat exemption-driven zero rates/amounts as numeric `0`
    - _Requirements: 3.4, 3.5, 3.6, 3.7, 3.8, 3.9, 3.10, 3.11, 3.12, 3.14_

  - [x] 5.2 Implement merge/assembly keyed by item serial
    - Implement `_merge_items(inv, duty, declared_count)` joining on `item_serial`; stitch multi-page item
      blocks so each field appears exactly once; never drop a serial present in either source; emit
      `ReviewFlag(line_item, serial, field, MISSING/UNPARSEABLE)` for missing/unresolvable fields; build
      `ExtractedDocument` with `declared_item_count`
    - _Requirements: 3.1, 3.13, 3.15, 9.2, 9.3, 9.5_

  - [x]* 5.3 Write property test for line-item completeness and flagging
    - **Property 1: No line item is ever dropped, and missing/failed fields are flagged**
    - **Validates: Requirements 3.1, 3.15, 5.11, 9.5, 2.10**

  - [x]* 5.4 Write property test for multi-page item merge equivalence
    - **Property 4: Multi-page item merge is equivalent to the unsplit item**
    - **Validates: Requirements 3.13**

  - [x]* 5.5 Write property test for unparseable-field retention
    - **Property 15: Unparseable fields retain raw text, are flagged, and are not removed**
    - **Validates: Requirements 9.2, 9.3**

  - [x]* 5.6 Write unit tests for line-item extraction against the reference PDF
    - Assert per-line fields for known items, the exemption-zero case, and a known multi-page item
    - _Requirements: 3.4, 3.5, 3.6, 3.7, 3.8, 3.9, 3.10, 3.11, 3.12, 3.14_

- [x] 6. Checkpoint - Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise.

- [x] 7. Implement the Excel_Generator (exact CTN layout)
  - [x] 7.1 Implement header-block and item-table writing
    - Implement `generate`, `_write_header_block` (D1:G8), and `_write_item_table` (header row 12, data
      from row 13): write Sr. no. as the dense sequence 1..N in ascending serial order with no blank rows;
      write direct + computed values at full precision into their mapped columns A..Y; reproduce each
      header label character-for-character (including quirks like `'Rate Per USDin purcahse '` and
      `'Ratepurchase per unitPcs/KGS/SET'`); leave review-flagged cells and no-source columns blank;
      write raw text for unparseable fields
    - _Requirements: 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 4.7, 4.8, 4.9, 5.1, 5.2, 5.3, 5.4, 5.5, 5.6, 5.7, 5.10, 5.11, 5.12, 6.11, 8.4, 8.5, 8.6_

  - [x] 7.2 Implement totals row and auxiliary template writing
    - Implement `_write_totals_row` (row 61: sums for G, L, M, N, O, P, Q, S, U, W, X plus package count)
      and `_write_aux_templates` (reproduce `DETAILS AS PER CHALLANS`, `DETAILS AS PER TALLY`,
      `CLEARANCE AND FORWARDING INVOICE`, and C&F block labels verbatim at sample positions with empty
      data cells); ensure exactly one sheet named `Sheet1`
    - _Requirements: 7.4, 7.5, 8.1, 8.2, 8.3_

  - [x]* 7.3 Write property test for Sr. no. sequence and dense ascending rows
    - **Property 9: Sr. no. forms the consecutive sequence 1..N and rows are dense and ascending**
    - **Validates: Requirements 5.1, 5.2**

  - [x]* 7.4 Write property test for full-precision value placement
    - **Property 10: Every value is written to its mapped cell unchanged at full precision**
    - **Validates: Requirements 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 4.7, 5.3, 5.4, 5.5, 5.6, 5.7, 5.10, 6.11, 6.12, 7.4, 7.5, 8.4, 8.5**

  - [x]* 7.5 Write property test for empty no-source and sample-empty cells
    - **Property 11: No-source and sample-empty cells are always empty**
    - **Validates: Requirements 4.8, 4.9, 8.6**

  - [x]* 7.6 Write property test for output structure and fixed labels
    - **Property 12: Output structure and fixed labels match the sample exactly**
    - **Validates: Requirements 5.12, 8.1, 8.2, 8.3**

  - [x]* 7.7 Write unit tests for header/auxiliary labels against the sample workbook
    - Compare item-table and auxiliary labels to the exact strings captured from `1357 ctn llp.xlsx`
    - _Requirements: 5.12, 8.3_

- [x] 8. Checkpoint - Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise.

- [x] 9. Implement upload validation and the Conversion Orchestrator
  - [x] 9.1 Implement upload validation
    - Implement `UploadValidator.validate(raw, filename)` with ordered checks: size ≤ 50 MB → readable
      PDF → recognizable BOE (marker tokens `BILL OF ENTRY`, `Port Code`, `BE No`, `PART - II`,
      `PART - III`); return the first failure with its message, else OK with an opened handle
    - _Requirements: 1.2, 1.5, 1.6_

  - [x]* 9.2 Write unit tests for rejection paths
    - Oversized file, corrupt/non-PDF bytes, and a valid non-BOE PDF each yield the correct message and
      no output
    - _Requirements: 1.2, 1.5, 1.6_

  - [x] 9.3 Implement orchestration, cross-checks, and summary
    - Sequence validate → parse → compute → generate; aggregate all `ReviewFlag`s; produce
      `Discrepancy(ITEM_COUNT)` when extracted count ≠ declared (mark output not complete);
      `Discrepancy(INVOICE_TOTAL)` when computed vs declared total differ by > 0.01 USD (retain workbook)
      and a "could not be verified" message when the declared total is absent/unreadable;
      `Discrepancy(RECOMPUTE)` when an extracted value differs from its recompute by > 0.01; build the
      `ConversionSummary`; enforce atomic output (issue a `download_token` only after a fully successful
      build); enforce the 60-second budget treating overruns as post-recognition failures
    - _Requirements: 3.2, 3.3, 7.6, 7.7, 9.1, 9.4, 1.3, 1.7_

  - [x]* 9.4 Write property test for extracted item-count preservation and mismatch reporting
    - **Property 2: Extracted item count is preserved and mismatches are reported**
    - **Validates: Requirements 3.2, 3.3**

  - [x]* 9.5 Write property test for invoice-total verification tolerance
    - **Property 13: Invoice-total verification reports out-of-tolerance differences**
    - **Validates: Requirements 7.6**

  - [x]* 9.6 Write property test for extracted-vs-recomputed mismatch reporting
    - **Property 14: Extracted-vs-recomputed mismatches beyond tolerance are reported**
    - **Validates: Requirements 9.4**

  - [x]* 9.7 Write property test for completion-summary accuracy
    - **Property 16: Completion summary accurately reflects the conversion**
    - **Validates: Requirements 9.1**

  - [x]* 9.8 Write property test for atomic failure (no downloadable output)
    - **Property 17: Failure after BOE recognition yields no downloadable output**
    - **Validates: Requirements 1.7**

- [x] 10. Checkpoint - Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise.

- [x] 11. Implement the Web_Interface and wire the system end to end
  - [x] 11.1 Implement the HTTP API and upload page
    - Implement `POST /api/convert` (multipart: single `file` + `usd_rate`) returning
      `{ download_token, summary }` on success and the mapped 4xx/422 error bodies on rejection/failure;
      implement `GET /api/download/{download_token}` returning the `.xlsx` bytes (404 on unknown/expired);
      serve a single-file upload control with a processing indicator and a results/summary view enabling
      download; flag that the endpoint is unauthenticated (Milestone 1 local tool)
    - _Requirements: 1.1, 1.4, 1.8_

  - [x]* 11.2 Write end-to-end and atomicity integration tests
    - Convert the reference PDF with `usd_rate = 95.3`: assert a downloadable `Sheet1` workbook within
      60s with 45 line items and totals at row 61; golden-output comparison against `1357 ctn llp.xlsx`
      evaluated values within tolerance; inject a post-recognition failure and assert no token is issued
    - _Requirements: 1.3, 1.7, 1.8_

- [x] 12. Final checkpoint - Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise.

## Grand-Total Verification (Requirements 10 & 11)

The tasks below extend the completed Milestone 1 system with Part I duty-summary extraction and
grand-total verification with red-fill highlighting. They build on the existing components without
altering Milestone 1 behavior: the parser gains one extraction stage, the orchestrator gains one
cross-check and reorders `convert()` so verification runs before `generate()`, and the Excel_Generator
gains conditional red fills on the Totals_Row. Property tests map 1:1 to design Properties 18–20 and use
the 1.00 INR tolerance; example tests assert extraction against both reference BOEs.

- [x] 13. Extend shared data models for grand-total verification
  - [x] 13.1 Add declared duty-summary fields and grand-total verification models
    - In `boe_converter/models.py`, add `declared_bcd`, `declared_sws`, `declared_igst`, and
      `declared_total_duty` `RawValue` fields to `HeaderBlock`, each with
      `field(default_factory=RawValue.missing)` so an absent value is never inferred
    - Add `"GRAND_TOTAL"` to the `Discrepancy.kind` `Literal`
    - Add the module constant `GRAND_TOTAL_TOLERANCE_INR = 1.00`
    - Add the frozen dataclasses `GrandTotalCheck` (`column: Literal["P","Q","S"]`, `declared: float|None`,
      `excel_total: float`, `status: Literal["MATCH","MISMATCH","UNVERIFIED"]`) and
      `GrandTotalVerification` (`checks: list[GrandTotalCheck]`) with the `highlight_columns` property
      returning the set of column letters whose check status is `MISMATCH`
    - _Requirements: 10.1, 10.2, 10.3, 10.4, 11.1, 11.2, 11.3, 11.4_

- [x] 14. Implement duty-summary extraction in the PDF_Parser
  - [x] 14.1 Implement `_extract_duty_summary` and wire it into `_extract_header`
    - In `boe_converter/parser.py`, add `_extract_duty_summary(self, pages) -> dict[str, RawValue]`
      returning `declared_bcd`/`declared_sws`/`declared_igst`/`declared_total_duty`
    - Reconstruct rows from `_upright_words(page)` (orientation-aware) so rotated margin labels are
      excluded, locate the Part I `C. DUTY SUMMARY` band and its label tokens (`1.BCD`, `3.SWS`,
      `7.IGST`, `14.TOTAL DUTY`, accepting the bare `TOTAL DUTY` variant), and read each numeric value
      with the existing positional column-window approach via `self._capture(raw_text, numeric=True)`
      (verbatim, non-destructive numeric parse per Req 10.5)
    - Wire the four `RawValue`s onto `HeaderBlock.declared_bcd`/`declared_sws`/`declared_igst`/
      `declared_total_duty`; a label found but unresolvable yields `RawValue.unparseable(...)`, a label
      not found yields `RawValue.missing()`
    - Emit `ReviewFlag(scope="header", field_name="declared_*", reason="MISSING"|"UNPARSEABLE")` for each
      missing/unreadable declared value, reported by field name
    - _Requirements: 10.1, 10.2, 10.3, 10.4, 10.5, 10.6_

  - [ ]* 14.2 Write example extraction tests against both reference BOEs
    - Assert the four declared duty-summary values for `205090022062026INNSA1BE0230620261842.pdf`
      (BCD `184250.2`, SWS `18810.2`, IGST `258391`, TOTAL DUTY `465302`) preserving printed digits
    - Assert extraction against `229411903072026INNSA1BE0040720261315.pdf` to confirm the label
      set/column geometry generalise beyond the single sample
    - _Requirements: 10.1, 10.2, 10.3, 10.4, 10.5_

  - [ ]* 14.3 Write missing/unreadable duty-summary example test
    - Assert a missing/unresolvable duty-summary value is recorded as `RawValue.missing()` and produces a
      reported `declared_*` review flag, with no inferred/default value substituted
    - _Requirements: 10.6_

- [x] 15. Implement grand-total verification in the Conversion Orchestrator
  - [x] 15.1 Implement `_check_grand_totals` / `_verify_grand_totals` and reorder `convert()`
    - In `boe_converter/orchestrator.py`, add `_verify_grand_totals(extracted, computed)` producing a
      `GrandTotalVerification` from the three comparisons: P (`Totals.total_customs_duty` vs
      `declared_bcd + declared_sws`), Q (`Totals.total_igst` vs `declared_igst`), and S
      (`Totals.total_customs_duty + Totals.total_igst` vs `declared_total_duty`), using
      `GRAND_TOTAL_TOLERANCE_INR` (1.00) — `MATCH` when within tolerance, `MISMATCH` when beyond,
      `UNVERIFIED` when a required declared value is missing/unreadable (column P is `UNVERIFIED` when
      either `declared_bcd` or `declared_sws` is missing)
    - Add `_check_grand_totals(...)` alongside `_check_item_count` / `_check_invoice_total` /
      `_check_recompute`, emitting `Discrepancy(kind="GRAND_TOTAL", expected=declared, actual=excel_total,
      message=…names column P/Q/S…)` for each `MISMATCH`, and reporting "the column X grand total could
      not be verified against the BOE" for each `UNVERIFIED`
    - Reorder `convert()` so the grand-total verification runs **before** `generate()`, then thread the
      `GrandTotalVerification` into `generate(...)`
    - Append the `GRAND_TOTAL` discrepancies into the returned cross-check list so they appear in
      `ConversionSummary`; never let them affect `output_complete` — the workbook is always retained and
      downloadable
    - _Requirements: 11.1, 11.2, 11.3, 11.6, 11.7, 11.8, 11.9_

  - [ ]* 15.2 Write property test for grand-total retention and summary inclusion
    - **Property 20: Grand-total mismatches are retained and surfaced in the summary**
    - **Validates: Requirements 11.8, 11.9**

- [x] 16. Implement red-fill highlighting in the Excel_Generator
  - [x] 16.1 Add the optional `grand_total_check` parameter and apply red fills
    - In `boe_converter/excel_writer.py`, add `grand_total_check: GrandTotalVerification | None = None` to
      `generate(...)`; when `None`, preserve exact Milestone 1 behavior (no fills)
    - After the Totals_Row is written, for each column letter in
      `grand_total_check.highlight_columns` apply `PatternFill(fill_type="solid", fgColor="FFFF0000",
      start_color="FFFF0000")` to that Totals_Row cell using column indices P=16, Q=17, S=19 (the existing
      `COL_*` constants)
    - Only `MISMATCH` columns are filled; `MATCH` and `UNVERIFIED` columns keep their existing unstyled
      cells
    - _Requirements: 11.4, 11.5, 11.7_

  - [ ]* 16.2 Write property test for grand-total mismatch highlight and reporting at the tolerance boundary
    - **Property 18: Grand-total mismatch beyond tolerance is highlighted red and reported; within tolerance is not**
    - **Validates: Requirements 11.1, 11.2, 11.3, 11.4, 11.5, 11.6**

  - [ ]* 16.3 Write property test for missing declared value yielding could-not-verify with no fill
    - **Property 19: Missing declared value yields could-not-verify with no fill**
    - **Validates: Requirements 10.6, 11.7**

- [x] 17. Checkpoint - Ensure all grand-total verification tests pass
  - Ensure all tests pass, ask the user if questions arise.

## Notes

- Tasks marked with `*` are optional and can be skipped for a faster MVP; they cover property, unit, and
  integration tests.
- Each task references specific requirements (and properties, where applicable) for traceability.
- Property tests map 1:1 to design Properties 1–20 and use `hypothesis` with `max_examples >= 100`;
  float comparisons use the 0.01 tolerance where specified, the 1.00 INR tolerance for the grand-total
  checks (Properties 18–20), and exact equality where full precision is asserted.
- Grand-total verification (Requirements 10 & 11, tasks 13–17) extends Milestone 1 without altering its
  behavior; `convert()` is reordered so verification runs before `generate()`, and the workbook is always
  retained and downloadable even on a grand-total mismatch.
- Checkpoints ensure incremental validation across components.
- The Value_Calculator and assembly logic are pure; the Orchestrator owns sequencing, cross-checks, and
  atomic output.

## Task Dependency Graph

```json
{
  "waves": [
    { "id": 0, "tasks": ["1.1"] },
    { "id": 1, "tasks": ["2.1", "2.5", "4.1", "4.4", "5.1", "9.1"] },
    { "id": 2, "tasks": ["2.2", "2.3", "2.4", "2.6", "4.2", "4.3", "4.5", "5.2", "9.2"] },
    { "id": 3, "tasks": ["5.3", "5.4", "5.5", "5.6", "7.1"] },
    { "id": 4, "tasks": ["7.2", "9.3"] },
    { "id": 5, "tasks": ["7.3", "7.4", "7.5", "7.6", "7.7", "9.4", "9.5", "9.6", "9.7", "9.8"] },
    { "id": 6, "tasks": ["11.1"] },
    { "id": 7, "tasks": ["11.2"] },
    { "id": 8, "tasks": ["13.1"] },
    { "id": 9, "tasks": ["14.1", "16.1"] },
    { "id": 10, "tasks": ["15.1"] },
    { "id": 11, "tasks": ["14.2", "14.3", "15.2", "16.2", "16.3"] }
  ]
}
```

## Milestone 1.1 — Field-Accuracy Corrections (tasks)

These tasks implement the bug-fix batch (Requirements 12–18). They are localized edits to existing
components; each ends wired into the running system with tests. Reference inputs:
`3. BE - 221981730062026INNSA1BE0040720261600.pdf`, `INV 1054.pdf`, and the client workbooks
`bill_of_entry - with mistake.xlsx` / `bill_of_entry - corrected copy.xlsx`.

- [x] 18. Shared unit-to-pieces factor table (DOZ/GRS/THD)
  - [x] 18.1 Extract the piece-conversion factors into one shared module
    - Create `boe_converter/units.py` with `UNIT_TO_PCS = {"DOZ": 12.0, "GRS": 144.0, "THD": 1000.0}` and
      `pcs_factor(unit) -> float | None` (trim + upper-case lookup)
    - Refactor `tally_exporter._UNIT_TO_PCS` / `_convert_to_pcs` to consume the shared table
    - _Requirements: 14.6_

- [x] 19. Fix the Excel `pcs` column to convert DOZ/GRS/THD
  - [x] 19.1 Compute `pcs` with the unit factor in the calculator and writer
    - In `calculator._pcs`, use `units.pcs_factor(unit)`: `pcs = qty * factor` when a factor exists, else
      `None`; expose the factor on `ComputedLine` (e.g. `pcs_factor`)
    - In `excel_writer._write_item_row`, emit `=H{row}*{factor}` (or the literal `qty*factor`) using the
      line's factor; blank when no factor
    - _Requirements: 14.1, 14.2, 14.3, 14.4, 14.5_
  - [ ]* 19.2 Update/extend the pcs property test for GRS and THD
    - **Property 6 (extended): pcs = QTY×factor for DOZ/GRS/THD, blank otherwise**
    - _Validates: Requirements 14.1, 14.2, 14.3, 14.4_

- [x] 20. Use the BOE-declared SWS rate instead of a fixed 10%
  - [x] 20.1 Extract per-line SWS rate/amount and drive the calculator from it
    - Ensure the parser records per-line `sws_rate` and `sws_amount` (exemption 0 → numeric 0)
    - In `calculator._compute_line`, compute `sws_amount` from the BOE value (amount if present incl. 0,
      else `cust_aidc * sws_rate`), falling back to `cust_aidc * 0.10` with a `ReviewFlag` only when no
      SWS data exists; stop applying the unconditional 10%
    - Write the actual SWS rate into Excel column V
    - _Requirements: 15.1, 15.2, 15.3, 15.4, 15.5_
  - [ ]* 20.2 Property/example test for SWS-from-BOE and the 0% case
    - Assert a 0 SWS rate yields 0 surcharge and reconciling P/Q/S totals; a present rate is used verbatim
    - _Validates: Requirements 15.2, 15.3, 15.5_

- [x] 21. Populate USD Amt / USD Rate and keep dates verbatim in the header
  - [x] 21.1 Header USD amount + rate + verbatim dates
    - Parser: capture the BOE USD exchange rate into `HeaderBlock.usd_rate_boe` (default missing); confirm
      the USD invoice amount lands on `HeaderBlock.invoice_amount`; store header dates as verbatim strings
    - Excel writer: write `G3` from the USD invoice amount; write `G2` from `usd_rate_boe` when present
      else the User-supplied rate; write header dates as their raw strings (no `datetime` coercion)
    - _Requirements: 12.1, 12.2, 12.3, 12.4, 12.5, 13.1, 13.2_
  - [ ]* 21.2 Example test against the reference BOE/invoice and corrected workbook
    - Assert `G3` = the BOE USD amount (`13757.09` for the reference), `G4` = the verbatim invoice-date
      string, and `G2` = the rate source per Req 12.3/12.4
    - _Validates: Requirements 12.2, 12.3, 13.1_

- [ ] 22. Remove round-off from computed and extracted monetary values
  - [ ] 22.1 Audit and fix precision loss
    - Locate the rounding that produced `Q15 = 55.3` (IGST/GST); remove value-level rounding on the
      extraction/calculation path (keep rounding in `number_format` only). Keep Unit Price in USD sourced
      from the BOE Part II UPI field, not from optional invoice attachments.
    - _Requirements: 16.1, 16.2, 16.3_
  - [ ]* 22.2 Regression test asserting full precision
    - Assert a known line's GST value is written unrounded (e.g. `55.29535`) and the Unit Price in USD
      cell matches the BOE-extracted UPI value.
    - _Validates: Requirements 16.1, 16.2_

- [ ] 23. Ensure the grand-total mismatch highlight fires
  - [~] 23.1 Reconcile the verification basis and confirm the fill persists
    - After the SWS fix (task 20), confirm `_verify_grand_totals` detects a genuine mismatch and that
      `generate(...)` receives the `GrandTotalVerification` on every conversion; confirm the red fill on
      P/Q/S survives to the saved workbook
    - _Requirements: 17.1, 17.2, 17.3_
  - [ ]* 23.2 End-to-end test: a mismatching BOE yields a red-filled Totals_Row cell
    - _Validates: Requirements 17.1, 17.2_

- [ ] 24. Keep NOS-unit line items in the JSON/Tally export
  - [~] 24.1 Ensure NOS items are extracted and never dropped from the JSON
    - Confirm `NOS` is accepted by the parser unit allow-list and survives `_merge_items`; ensure
      `TallyExporter.build` emits an inventory allocation for every line including NOS
    - _Requirements: 18.1, 18.2, 18.3, 18.4_
  - [ ]* 24.2 Regression test: a NOS-unit line appears in `allledgerentries` and in the Excel Unit column
    - _Validates: Requirements 18.2, 18.3_

- [~] 25. Checkpoint — full suite + reference-file conversion
  - Convert `3. BE - 221981730062026INNSA1BE0040720261600.pdf` with `INV 1054.pdf`; diff the output against
    `bill_of_entry - corrected copy.xlsx` for the corrected cells (G2/G3/G4, pcs factors, SWS/GST totals,
    grand-total highlight); ensure all tests pass

## Task Dependency Graph (Milestone 1.1)

```json
{
  "waves": [
    { "id": 0, "tasks": ["18.1"] },
    { "id": 1, "tasks": ["19.1", "20.1", "21.1", "22.1", "24.1"] },
    { "id": 2, "tasks": ["23.1"] },
    { "id": 3, "tasks": ["19.2", "20.2", "21.2", "22.2", "23.2", "24.2"] },
    { "id": 4, "tasks": ["25"] }
  ]
}
```
