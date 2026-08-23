# Todo
<!-- Ephemeral. Rewrite as work evolves. NOT append-only. -->

## In Progress
- (none)

## Up Next
- [ ] Import the fixed JSON into Tally and confirm the numbering-series error is gone
- [ ] Store Gemini Unicom LLP's real ledger names under Manage stored data -> Ledgers
      (the 18% one is `Igst Purchase @18.00%`, not `IGST Purchase @ 18.00 %`)
- [ ] Set `tax_unit` = "Maharashtra Registration" and `tally_name` =
      "Gemini Unicom LLP (F.Y. 2026-27)" on the Gemini Unicom LLP buyer record
- [ ] Open question: our IGST amounts differ slightly from Tally's (16787.92 vs
      16787.80 at 5%; 271084.65 vs 271085.00 at 18%) - per-line vs per-group
      rounding. Needs a decision from the user before changing the calculator.
- [ ] Repeat the ground-truth diff once data for a second company arrives

## Done (this session)
- [x] Root-caused and fixed the "numbering series already used for another GST
      Registration" import error (Auto numbering + no declared tax unit)
- [x] `referencedate` corrected to the supplier invoice date
- [x] Per-company ledger master (store table, LedgerBook, UI, derived-name warning)
- [x] Per-company tax unit + Tally company name; stored party/buyer spelling wins
- [x] User-selectable voucher date, defaulting to the BE date
- [x] Voucher now balances to exactly 0.00 (was a paisa out)
- [x] `28-MAR-26` style dates now parse
- [x] Foreign workbook layouts refused loudly instead of misread
- [x] Full suite green: 326 passed, 9 skipped
