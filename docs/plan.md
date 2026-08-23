# Plan

## 2026-08-22 — Make the Tally export generic across companies (decisions from the CO-04 data set)

### What the CO-04 data proved
Only one company's data is available (Gemini Unicom LLP), so the design is driven
by the *differences* between it and the earlier company rather than by guesswork:

| Evidence | Implication |
|---|---|
| Tally has `Igst Purchase @18.00%`; our convention yields `IGST Purchase @ 18.00 %` | Ledger names are hand-made per company and cannot be derived |
| Party ledger is `Prayan Impex Company Limited` here, `DIA IMPEX COMPANY LIMITED` for the other company | Party ledger casing is per-company too |
| `CO4-Workbook1.xlsx` lacks the `BILLING AMOUNT` column and puts its header on row 15 | The reader's fixed cell references silently misread foreign layouts |
| `basicbuyername` is `Gemini Unicom LLP (F.Y. 2026-27)` but `consigneemailingname` is `Gemini Unicom LLP` | The Tally company name is its own per-company field |
| `vchstatustaxunit` is `Maharashtra Registration` | The tax-unit name is per-company (defaulted from state) |
| Voucher dated 20 Apr, BE dated 14 Apr | The voucher date is a booking decision, not a document fact |

### Decisions (confirmed with the user)
1. **Ledger names** — store a per-company ledger list, looked up by IGST rate.
   Never silently invent: fall back to the convention only when nothing is stored,
   and surface exactly which names were derived.
2. **`CO4-Workbook1.xlsx`** — reference data only, not an upload format. The reader
   must *reject* unrecognised layouts loudly instead of producing garbage.
3. **Narration** — one fixed template for all companies. Keep the existing
   `<details> USD <amount> @<rate>` form (the format BUG-001 was raised against);
   operator-typed variations are not reproduced.
4. **Voucher date** — user-selectable, pre-filled with the BE date.

### Approach
- `tally_store`: new `ledgers(company, kind, rate_bp, name)` table mirroring the
  existing per-company `stock_items` design. Rate held as integer basis points
  (5% -> 500) so lookups never depend on float equality.
- `tally_exporter`: a `LedgerBook` consulted for every ledger name, with the
  current convention as the documented fallback and
  `derived_ledger_names()` reporting what was not found. The party ledger uses the
  stored seller name verbatim rather than `.title()`.
- `excel_reader`: validate anchor labels in the header block and the item-table
  header row before reading; raise `ExcelReadError` naming the mismatch.
- `CompanyProfile`: add `tally_name` (the FY-suffixed Tally company name driving
  `basicbuyername`) alongside the `tax_unit` added with the numbering fix.
- `build(computed, usd_rate, voucher_date=None)` — voucher and effective dates come
  from the caller, defaulting to the BE date; `referencedate` stays the invoice date.

### Trade-off
The convention fallback is kept deliberately: a company with no stored ledgers must
still produce a voucher. The safeguard is that the fallback is *reported*, so a
derived name is a visible choice rather than a silent duplicate-ledger bug.


## 2026-08-22 — BUG-001..005: Tally voucher identity + narration wrong on the Excel-upload path

### Evidence
Diffed `Tally_Exported.json` (ground truth, hand-entered voucher) against
`Software_Generated.json` for the same consignment (CO-49 CTN-1100 / inv 202607044):

| field | Tally (correct) | software |
|---|---|---|
| `narration` | `CO-49 CTN-1100 USD 31,453.57 @96.05` | `CO-49 CTN-1100 USD 0.00 @96.05` |
| `countryofresidence` | `China` | `India` |
| `placeofsupply` | `Maharashtra` | `""` |
| `consigneestatename` | `Maharashtra` | `""` |
| `cmpgststate` | `Maharashtra` | `""` |
| `gstregistrationtype` | `OIDAR` | *absent* |
| `cmpgstregistrationtype` | `Regular` | *absent* |
| `cmpgstin` | `27AAYFG7003K1ZW` | `""` |

Reproduced deterministically via `read_workbook(...)` + `TallyExporter().build(...)`
on `bill_of_entry - corrected copy.xlsx`.

### Root causes (all 5 bugs reproduce only on the **Excel-upload** path)
1. **BUG-001** `excel_reader._read_line` computes `amount_usd` (line 187) then
   drops it on the floor — it is never passed to `ComputedLine(...)`. And
   `_totals_from_lines` never sets `total_amount_usd`, so it stays `0.0` and the
   narration prints `USD 0.00`. Header cell `G3` ("USD Amt") is likewise ignored.
   Secondary: `_narration` used `f"{x:.2f}"`, so it lacked Tally's thousands
   separators (`31,453.57`).
2. **BUG-002** `excel_reader._read_header` never populates `seller_country`, so
   `_voucher_shell` hit its hardcoded `seller_country or "India"` fallback. On an
   import Bill of Entry the supplier is by definition overseas, so defaulting the
   *supplier's* country to India is always wrong.
3. **BUG-003 / BUG-004** Same root cause: `_read_header` never populates
   `buyer_state`/`buyer_gstin`, so `placeofsupply`, `cmpgststate` and
   `consigneestatename` (all fed by `_buyer_state`) collapse to `""`.
4. **BUG-005** `_voucher_shell` never emitted `gstregistrationtype` or
   `cmpgstregistrationtype` at all — not a value bug, a missing-key bug.

### Approach
Fix at the source of each defect, not at the symptom:
- `excel_reader`: stop losing data on the round-trip — carry `amount_usd` per
  line, read `G3` into `invoice_amount`, and sum `total_amount_usd`.
- `tally_exporter`: derive buyer state from the buyer GSTIN's state code (the
  first two digits of a GSTIN *are* the state — this is decoding, not guessing);
  emit the two GST registration-type keys; drop the wrong `"India"` supplier
  fallback; format the narration with thousands separators.
- `streamlit_app`: auto-resolve the stored buyer/seller record by name when the
  user has not explicitly picked one, so the Excel path recovers the identity the
  workbook cannot carry.

### Trade-offs
- The "never invent a missing value" contract
  (`test_missing_buyer_seller_fields_are_blank_not_guessed`) is preserved: state
  is only *derived* from a GSTIN that is actually present; when everything is
  absent the fields stay blank rather than being fabricated.
- Deliberately out of scope (present in Tally's export, not reported as bugs):
  `gstregistration` tax-unit object, `vchstatustaxunit`, `vouchernumber`.
