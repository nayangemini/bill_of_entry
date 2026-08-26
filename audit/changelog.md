# Change Log
<!-- Append-only. Newest at TOP. Written AFTER implementation is complete. -->
<!-- Format: YYYY-MM-DD HH:MM:SS - <one line summary> -->

---

2026-08-26 11:06:02 - BUG-002/003/004 reopened: carry buyer/seller identity through the Excel round-trip

### Changes
- The earlier fix removed the hardcoded `"India"` supplier-country fallback but left
  the field **blank**, which is not a fix: Tally substitutes its own company defaults
  for a blank field on import, so the bill still showed India for a Chinese supplier
  and an empty place of supply. The tester was right that nothing had changed.
- Root cause is one level up. The CTN workbook is a costing table with no GSTIN,
  state or country cells, so `PDF -> Excel -> re-upload` discarded the buyer and
  seller identity outright and the exporter had nothing left to emit:

      BEFORE write: gstin=27AAYFG7003K1ZW state=Maharashtra country=China
      AFTER  read : gstin=None            state=None       country=None

- The workbook now carries that identity as OOXML custom document properties
  (`write_identity_props` / `read_identity_props`): invisible in the grid, outside
  every golden-tested cell, unaffected by the one-sheet rule (Req 8.1), and
  preserved across the load-and-resave the Step 2 name mapping performs. Chosen
  over visible cells (the layout is reproduced character-for-character from the
  sample and golden-tested) and over a hidden sheet (`test_e2e_integration.py`
  asserts `wb.sheetnames == ["Sheet1"]`).
- Because a blank field fails *silently* in Tally, the app now warns at export time
  naming each field Tally would substitute a default for.

### Verification
With no stored buyer/seller configured, the reported path yields
`countryofresidence=China`, `placeofsupply=Maharashtra`,
`consigneestatename=Maharashtra`, `cmpgstin=27AAYFG7003K1ZW`.
Pre-change workbooks still read (identity stays missing, never invented).
Full suite: 347 passed, 9 skipped.

### Known follow-up
The GSTIN state-code table now exists twice: `parser._GSTIN_STATE` and
`tally_exporter._GSTIN_STATE_CODES`. Worth consolidating into a shared module
alongside `boe_converter/units.py`.

### Files
- `boe_converter/excel_writer.py` - modified
- `boe_converter/excel_reader.py` - modified
- `streamlit_app.py` - modified
- `tests/test_excel_identity_roundtrip.py` - created

---

2026-08-23 08:28:58 - Complete the BOE -> Tally purchase pipeline: cost centre, party allocation, no nameless allocations

### Changes
- **Cost centre is now supplied by the operator.** `header.details` (the consignment
  code, e.g. `CO-04 CTN 1255`) is hardcoded `RawValue.missing()` on the PDF path
  (`parser.py:522`, "out of scope for task 4.4") and the app had no input for it, so
  the direct BOE -> JSON route emitted `costcentrename: ""` **and** cost-centre
  allocations named `""` - asking Tally to allocate to a nameless cost centre.
  `build(..., cost_centre=...)` now takes it, defaulting to the workbook's `Details`
  cell, and the app has a "Consignment / cost centre" box.
- **No cost centre means no allocations at all.** New `_cost_centre_block()` returns
  an empty dict when there is nothing to allocate to, so the voucher simply carries
  no `categoryallocations`, and `iscostcentre` now reflects reality instead of being
  hardcoded `True`.
- **The party ledger is now allocated to the cost centre**, matching Tally's own
  voucher. The allocation shape now matches exactly: entries carrying stock items
  allocate per item, entries without (party, IGST purchase/payable) allocate at
  entry level, and Custom Duty Payable is not cost-centre tracked.

### Why this matters beyond the purchase voucher
The cost centre is what ties this purchase to the sales vouchers raised against the
same consignment - all ten sales vouchers in the CO-04 data carry
`CO-04 CTN 1255`, and 20 of their 26 stock items are the ones this purchase brought
in. Without it the link is lost.

### Verification
CO-04 Bill of Entry through the full pipeline, diffed against Tally's own export of
the same voucher: **31/31 header fields identical**, ledger + cost-centre allocation
shape identical, same cost centre throughout, voucher balances to 0.00. Full suite:
339 passed, 9 skipped.

### Also confirmed (investigation, no code change)
- `origin` is https://github.com/rushabhgandhi13/boe-converter - the linked repo *is*
  this project, not a separate one.
- The reference template `1357 ctn llp.xlsx` has this project's layout (header row
  12, `BILLING AMOUNT` at column C), and both have been unchanged since the initial
  commit. `CO4-Workbook1.xlsx` is therefore *not* this converter's output, so
  refusing it stands.
- `1. INV 1255-OK.pdf` is Prayan Impex's supplier invoice (202603030, 28-03-2026) -
  an input to the BOE converter, not a sales document.

### Files
- `boe_converter/tally_exporter.py` - modified
- `streamlit_app.py` - modified
- `tests/test_tally_cost_centre.py` - created

---

2026-08-22 22:38:35 - Make the Tally export per-company; fix voucher imbalance, date parsing and silent workbook misreads

### Changes
- **Per-company ledger names.** New `ledgers(company, kind, rate_bp, name)` table and
  a `LedgerBook` the exporter consults for every ledger name, keyed by IGST rate in
  integer basis points. Ledger names cannot be derived: this company's Tally holds
  `Igst Purchase @18.00%` where the convention gives `IGST Purchase @ 18.00 %`, and
  importing the wrong spelling creates a duplicate ledger. The convention remains a
  fallback for companies with nothing stored, and `derived_ledger_names()` reports
  every fallback so the app can warn instead of failing silently.
- **Party ledger** now uses the stored seller name verbatim (companies differ:
  `Prayan Impex Company Limited` vs `DIA IMPEX COMPANY LIMITED`) instead of
  title-casing the BOE's. Likewise the stored buyer name is the canonical spelling.
- **Per-company identity**: `CompanyProfile.tax_unit` (GST registration name) and
  `.tally_name` (the FY-suffixed Tally company name driving `basicbuyername`),
  persisted as new `buyers` columns and editable in the app.
- **Voucher date** is now the caller's choice (`build(..., voucher_date=...)`),
  pre-filled with the BE date in the app - booking date is a bookkeeping decision.
- **The voucher now balances to exactly 0.00.** `land_cost_excl_gst == purchase_inr
  + duty` holds in full precision, but each aggregate was rounded to paise
  independently, leaving a one-paisa residue - invalid double entry (Tally's own
  voucher sums to exactly 0.00). The supplier payable is now derived from the other
  entries, which the identity defines anyway; purchase ledgers stay pinned to the
  sum of their stock items. The balance assertions carried tolerances (0.5, 0.05,
  0.01) that were hiding this; they now assert exact equality in `Decimal`.
- **Date parsing**: `_tally_date` now handles the alphabetic-month form a BOE prints
  its invoice date in (`28-MAR-26`). It previously returned "" and silently fell
  back to the BE date, losing `referencedate`.
- **Foreign workbooks are refused, loudly.** `read_workbook` used fixed cell
  references, so the client's own sheet (no `BILLING AMOUNT` column, header on row
  15) was read column-for-column into the wrong fields - party name came back as
  "USD Rate" - with no error. Anchor labels are now checked first and a mismatch
  raises `ExcelReadError` naming the offending cell.

### Verification
Ran the CO-04 Bill of Entry through the full pipeline and diffed against Tally's own
export of the same voucher: **28/28 header fields identical**, `gstregistration`
identical, ledger-name set identical, no derived ledgers, and both vouchers balance
to 0.00. Full suite: 326 passed, 9 skipped (Neon integration tests).

### Files
- `boe_converter/tally_exporter.py` - modified
- `boe_converter/tally_store.py` - modified
- `boe_converter/excel_reader.py` - modified
- `streamlit_app.py` - modified
- `tests/test_tally_company_generic.py` - created
- `tests/test_tally_numbering.py` - modified
- `tests/test_tally_export.py` - modified (tolerances tightened to exact)

---

2026-08-22 20:27:46 - Fix Tally import exception: "voucher numbering series already used for another GST Registration"

### Changes
- Root cause: the voucher asked Tally to allocate a number from the Purchase
  voucher type's **automatic** numbering series (`numberingstyle: "Auto"`) while
  declaring no GST tax unit. Tally binds each automatic series to one GST
  registration, so it could not match the series to this voucher's registration.
  All eleven Tally-exported vouchers in the CO-04 data set (1 Purchase, 10 Sales)
  instead use `numberingstyle: "Manual"` with an explicit `vouchernumber`,
  `vouchernumberseries: "Default"` and `vchstatustaxunit`.
- `tally_exporter` now numbers the voucher manually with the BE number (falling
  back to the invoice number), emits `vouchernumberseries: "Default"`, and only
  falls back to `Auto` when no number is known at all.
- Emits the GST registration binding: `gstregistration` (value/taxtype/
  taxregistration) plus `vchstatustaxunit`, `vchstatusvouchertype`,
  `vchstatustaxadjustment` and `vchstatusdate`. The tax-unit name defaults to
  Tally's own `"<state> Registration"` and is overridable per company via the new
  `CompanyProfile.tax_unit`. Nothing is emitted when the state is unknown.
- `referencedate` now carries the *supplier invoice* date rather than the BE date,
  pairing correctly with `reference` (the invoice number), as Tally's own
  vouchers do. Falls back to the BE date when the invoice date is absent.

### Verification
Ran the CO-04 Bill of Entry through the full pipeline and diffed the result
against Tally's own export of the same voucher (`Purchase_8668342.json`): all
numbering and GST-registration fields match exactly, along with country, place of
supply, ship-to state, GSTIN, pincode and reference. Full suite: 283 passed,
9 skipped (Neon integration tests).

### Files
- `boe_converter/tally_exporter.py` - modified
- `tests/test_tally_numbering.py` - created

---

2026-08-22 18:48:37 - Fix BUG-001..005: Tally voucher narration, supplier country, place of supply, ship-to state and GST registration type

### Changes
- **BUG-001** `excel_reader._read_line` computed `amount_usd` from column L and then
  discarded it instead of passing it to `ComputedLine`, and `_totals_from_lines`
  never set `total_amount_usd` - so the Tally narration printed `USD 0.00`. Both
  now carry the value, header cell `G3` ("USD Amt") is read into `invoice_amount`
  and stands in when column L has no cached values, and `_narration` formats the
  total with thousands separators to match Tally (`USD 31,453.57`).
- **BUG-002** `_voucher_shell` emitted `"countryofresidence": seller_country or "India"`.
  A Bill of Entry is an import document, so defaulting the *supplier's* country to
  India was always wrong; the fallback is removed.
- **BUG-003 / BUG-004** `placeofsupply`, `cmpgststate` and `consigneestatename` all
  come from `_buyer_state`, which was blank on the Excel path. It now falls back to
  decoding the buyer GSTIN's state code (`27` -> Maharashtra) via the new
  `state_from_gstin`. Nothing is invented: with no state *and* no GSTIN it stays blank.
- **BUG-005** `gstregistrationtype` and `cmpgstregistrationtype` were never emitted.
  The supplier's type is now `OIDAR` unless the supplier is explicitly in India
  (`Regular`); the company's type is a new `CompanyProfile.gst_registration_type`
  field defaulting to `Regular`.
- Added `company_profile_for` / `seller_profile_for`: build a profile from a stored
  buyer/seller record filling **only the gaps** the document itself cannot carry, so
  a re-uploaded CTN workbook (which has no identity cells) still produces a complete
  voucher while the BOE always wins where it states a value.
- Streamlit now auto-matches the stored buyer/seller by name when the user leaves the
  dropdown on "(from BOE)", instead of silently exporting blank identity fields.

### Verification
Full suite: 270 passed, 9 skipped (Neon integration tests, need `BOE_TEST_DATABASE_URL`).
End-to-end against `bill_of_entry - corrected copy.xlsx` now matches the hand-entered
Tally export on all five reported fields.

### Files
- `boe_converter/excel_reader.py` — modified
- `boe_converter/tally_exporter.py` — modified
- `streamlit_app.py` — modified
- `tests/test_tally_voucher_bugs.py` — created
- `docs/plan.md` — created
- `tasks/todo.md` — created
