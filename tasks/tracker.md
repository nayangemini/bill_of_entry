# Task Tracker
<!-- Append-only. Newest at TOP. -->
<!-- Format: ## YYYY-MM-DD HH:MM:SS — <summary> -->

## 2026-10-05 19:30:03 — Reviewed, fixed and merged PR #2 (MTS unit, tape pc override, invoice cartons)
**Type:** task-complete
**Outcome:** PR #2 from nayangemini is merged into `main` (`27c297c`). Review confirmed
its claims but found the tape override was lost in the default Tally path once a line
was mapped in Step 2 (Excel 26880 PCS, Tally 272 KGS on the real bill), and that the
rule fired on any word containing "tape", on lines already declared in pieces, and on
chosen Tally master names. Fixed on the PR branch before merging (`0136a0a`): the
booking quantity is decided once in the calculator and carried on the line; the tape
rule needs the whole word and a weight unit; two guards added to the invoice row
parser. 92 tests added; full suite 439 passed, 9 skipped; real bills differ from the
previous `main` only on the weighed tape line.
**Open:** the digit-join threshold and row-stitching rule were not changed (need the
contributor's live bills to re-verify).
**Files changed:** boe_converter/units.py, boe_converter/models.py, boe_converter/calculator.py, boe_converter/excel_writer.py, boe_converter/tally_exporter.py, boe_converter/invoice_parser.py, streamlit_app.py, tests/test_units.py, tests/test_stock_quantity.py, tests/test_invoice_row_parsing.py

## 2026-08-26 11:06:02 — BUG-002/003/004 reopened and fixed at the real root cause
**Type:** task-complete
**Outcome:** The tester reported the three field bugs were still present. They were:
the earlier change removed the wrong `"India"` default but emitted a blank field, and
Tally fills a blank field with its own company defaults. The actual cause is that the
CTN workbook carries no identity cells, so the Excel round-trip dropped buyer GSTIN /
state and seller country entirely. Identity now travels as OOXML custom document
properties, leaving the golden-tested layout and the one-sheet rule untouched. Export
now warns when any field would be left for Tally to default.
**Open:** the tester must re-download the Excel from Step 1 — workbooks downloaded
before this change carry no identity and will still export blank (now with a warning).
**Files changed:** boe_converter/excel_writer.py, boe_converter/excel_reader.py, streamlit_app.py, tests/test_excel_identity_roundtrip.py

## 2026-08-23 08:28:58 — BOE -> Tally purchase pipeline complete (31/31 fields match Tally)
**Type:** task-complete
**Outcome:** Closed the last gap in the purchase pipeline: the cost centre had no
route in on the PDF path, so vouchers went out allocating to a nameless cost centre.
Operator now supplies it; absent it, no allocations are emitted at all. Party ledger
now allocated, matching Tally's shape exactly. Confirmed the linked GitHub repo is
this same project and that CO4-Workbook1.xlsx is not its output.
**Open:** the sales-voucher side is not built - it needs a customer/price/quantity
source the BOE cannot provide. Question put to the user.
**Files changed:** boe_converter/tally_exporter.py, streamlit_app.py, tests/test_tally_cost_centre.py

## 2026-08-22 22:38:35 — Per-company Tally export + three defects found while verifying
**Type:** task-complete
**Outcome:** Ledger names, GST tax unit, Tally company name and party spelling are
now per-company stored data rather than derived. Verifying against Tally's own
export surfaced three further defects: the voucher was a paisa out of balance, the
`28-MAR-26` invoice-date form was unparseable, and foreign workbook layouts were
read into the wrong fields without error. All fixed. 28/28 header fields now match
Tally exactly. Suite: 326 passed, 9 skipped.
**Files changed:** boe_converter/tally_exporter.py, boe_converter/tally_store.py, boe_converter/excel_reader.py, streamlit_app.py, tests/test_tally_company_generic.py, tests/test_tally_numbering.py, tests/test_tally_export.py

## 2026-08-22 20:27:46 — Fixed the Tally numbering-series import exception (CO-04 data set)
**Type:** task-complete
**Outcome:** Traced the import error to `numberingstyle: "Auto"` combined with a
missing GST tax-unit declaration. Voucher is now numbered manually with the BE
number and declares its `gstregistration` / `vchstatustaxunit`; `referencedate`
corrected to the invoice date. Verified field-by-field against Tally's own export
of the same voucher. Suite: 283 passed, 9 skipped.
**Files changed:** boe_converter/tally_exporter.py, tests/test_tally_numbering.py, audit/changelog.md

## 2026-08-22 18:48:37 — Fixed BUG-001..005 in the Tally purchase-voucher export
**Type:** task-complete
**Outcome:** All five reported bugs traced to two root causes and fixed at source:
(1) the Excel-upload path silently dropped data on the round-trip (`amount_usd`,
`total_amount_usd`, header `G3`) and carried no party identity at all, and
(2) `_voucher_shell` hardcoded `"India"` as the supplier country and omitted both
GST registration-type keys. Buyer state is now decoded from the GSTIN state code, and
stored buyer/seller records fill the gaps the workbook cannot carry. 32 new regression
tests; full suite 270 passed / 9 skipped.
**Files changed:** boe_converter/excel_reader.py, boe_converter/tally_exporter.py, streamlit_app.py, tests/test_tally_voucher_bugs.py, docs/plan.md, tasks/todo.md, audit/changelog.md
