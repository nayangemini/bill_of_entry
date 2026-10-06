# Todo
<!-- Ephemeral. Rewrite as work evolves. NOT append-only. -->

## In Progress
- (none)

## Up Next
- [ ] PR #2 follow-up: the 6pt digit-join threshold and the split-row stitching rule in
      `invoice_parser.py` were left as the contributor tuned them. Re-check both against
      their live bills (BE 2243083 and the `CTN`-unit invoice) when those are available
- [ ] PR #2 follow-up: a tape line declared in DOZ/NOS/ROL keeps its BOE quantity. If a
      live bill needs the piece-count override for another unit, add it to
      `units.WEIGHT_UNITS`
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
- [x] Merged PR #3 (`5562247`): unit merged with its neighbour (`THD$0.12`) now marks
      an invoice line item; conflict with PR #2's parser resolved by hand
- [x] 28 tests added; full suite green: 467 passed, 9 skipped
- [x] Both real invoices parse identically to before
