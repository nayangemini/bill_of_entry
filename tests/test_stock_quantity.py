"""A line's booking quantity is decided once and every output agrees on it.

Two kinds of line are booked under a different quantity than the BOE prints:

- ``MTS`` books as ``KGS`` (x1000), the Tally stock unit;
- a tape line declared by weight whose name carries the piece count in
  brackets - ``PTFE TEFLON TAPE (26880 pc)``, 272 KGS - books 26880 PCS under
  the base name.

Monetary amounts never change; only quantity, unit and the per-unit rates do.

The decision is made by ``ValueCalculator`` and carried on the ``ComputedLine``
(``stock_qty`` / ``stock_unit``). It used to be re-derived from the description
by each output, so mapping a line to a Tally name in Step 2 (which replaces the
description) silently dropped it: the Excel showed 26880 PCS while the Tally
voucher booked 272 KGS. These tests pin the three paths to one answer.

All data is synthetic; the tape line mirrors a real BOE line.
"""

from __future__ import annotations

import pytest

from boe_converter.calculator import ValueCalculator
from boe_converter.excel_reader import read_workbook, write_tally_names
from boe_converter.excel_writer import (
    COL_AMOUNT,
    COL_DESCRIPTION,
    COL_QTY,
    COL_RATE_PER_UNIT,
    COL_UNIT,
    COL_UNIT_PRICE_USD,
    ExcelGenerator,
)
from boe_converter.models import ComputedLine, LineItem, RawValue, ReviewFlagSet
from boe_converter.tally_exporter import TallyExporter, apply_stock_names

from tests.test_tally_export import make_doc, make_line, rv

USD_RATE = 95.3
TAPE = "PTFE TEFLON TAPE (26880 pc)"


def _item(description: str, qty, unit: str, serial: int = 1) -> LineItem:
    quantity = qty if isinstance(qty, RawValue) else rv(qty)
    return LineItem(
        item_serial=serial,
        cth_hsn=rv("39199090"),
        description=rv(description),
        unit_price_usd=rv(0.8),
        quantity=quantity,
        unit=rv(unit),
        assessable_value=rv(20000.0),
        bcd_rate=rv(0.1),
        bcd_amount=rv(2000.0),
        igst_rate=rv(0.18),
        total_duty=rv(6000.0),
    )


def _line(description: str, qty, unit: str, serial: int = 1) -> ComputedLine:
    return ValueCalculator().compute_line(_item(description, qty, unit, serial), USD_RATE)


def _sheet(line: ComputedLine, formulas: bool):
    gen = ExcelGenerator(use_formulas=formulas)
    return gen.build_workbook(make_doc([line]), ReviewFlagSet([]))["Sheet1"]


def _stock_entries(node, out=None) -> list[dict]:
    """Every inventory allocation in a voucher, in order."""
    out = [] if out is None else out
    if isinstance(node, dict):
        if "stockitemname" in node:
            out.append(node)
        for value in node.values():
            _stock_entries(value, out)
    elif isinstance(node, list):
        for value in node:
            _stock_entries(value, out)
    return out


def _booked(doc) -> tuple[str, str, str, str]:
    entry = _stock_entries(TallyExporter().build(doc, USD_RATE))[0]
    return entry["stockitemname"], entry["actualqty"], entry["rate"], entry["amount"]


# ---------------------------------------------------------------------------
# Calculator: the decision
# ---------------------------------------------------------------------------
def test_calculator_books_mts_as_kgs():
    line = _line("SILICONE SEALANT", 1.024, "MTS")
    assert (line.stock_qty, line.stock_unit) == (1024.0, "KGS")
    assert line.source.description.parsed == "SILICONE SEALANT"
    assert line.purchase_rate_per_unit == line.land_cost_excl_gst / 1024.0


def test_calculator_books_a_weighed_tape_line_in_pieces():
    line = _line(TAPE, 272.0, "KGS")
    assert (line.stock_qty, line.stock_unit) == (26880.0, "PCS")
    assert line.source.description.parsed == "PTFE TEFLON TAPE"
    assert line.source.description.raw_text == "PTFE TEFLON TAPE"
    assert line.purchase_rate_per_unit == line.land_cost_excl_gst / 26880.0


def test_calculator_keeps_the_boe_quantity_and_every_amount():
    """Only quantity/unit/rate are re-expressed; nothing monetary moves."""
    tape = _line(TAPE, 272.0, "KGS")
    plain = _line("PTFE TEFLON", 272.0, "KGS")
    assert tape.source.quantity.parsed == 272.0
    assert tape.source.unit.parsed == "KGS"
    for name in (
        "amount_usd", "purchase_inr", "total_customs_duty", "igst_amount",
        "land_cost_excl_gst", "land_cost_incl_gst",
    ):
        assert getattr(tape, name) == getattr(plain, name), name


def test_calculator_leaves_an_ordinary_line_alone():
    item = _item("KEYCHAIN", 100.0, "KGS")
    line = ValueCalculator().compute_line(item, USD_RATE)
    assert line.stock_qty is None and line.stock_unit is None
    assert line.source is item
    assert line.purchase_rate_per_unit == line.land_cost_excl_gst / 100.0


@pytest.mark.parametrize(
    "description,qty,unit",
    [
        ("TAPERED ROLLER BEARING (50PCS)", 200.0, "KGS"),  # not the word "tape"
        ("MASKING TAPE (12PCS)", 200.0, "DOZ"),            # already in pieces
        ("SALO TAPE", 4500.0, "KGS"),                      # no bracket
    ],
)
def test_calculator_does_not_override_lookalikes(description, qty, unit):
    line = _line(description, qty, unit)
    assert line.stock_qty is None
    assert line.source.description.parsed == description


def test_calculator_cannot_convert_an_unreadable_mts_quantity():
    line = _line("SILICONE SEALANT", RawValue.unparseable("1.O24"), "MTS")
    assert line.stock_qty is None and line.stock_unit is None


# ---------------------------------------------------------------------------
# Excel workbook
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("formulas", [False, True])
def test_excel_writes_mts_as_kgs(formulas):
    line = _line("SILICONE SEALANT", 1.024, "MTS")
    ws = _sheet(line, formulas)
    assert ws.cell(13, COL_QTY).value == 1024.0
    assert ws.cell(13, COL_UNIT).value == "KGS"
    # K is re-expressed per KG so that L = H * K still gives the line amount.
    k = ws.cell(13, COL_UNIT_PRICE_USD).value
    assert k == pytest.approx(line.amount_usd / 1024.0)
    assert 1024.0 * k == pytest.approx(line.amount_usd)
    if formulas:
        assert ws.cell(13, COL_AMOUNT).value == "=H13*K13"
        assert ws.cell(13, COL_RATE_PER_UNIT).value == "=O13/H13"
    else:
        assert ws.cell(13, COL_AMOUNT).value == line.amount_usd
        assert ws.cell(13, COL_RATE_PER_UNIT).value == line.land_cost_excl_gst / 1024.0


@pytest.mark.parametrize("formulas", [False, True])
def test_excel_writes_a_weighed_tape_line_in_pieces(formulas):
    line = _line(TAPE, 272.0, "KGS")
    ws = _sheet(line, formulas)
    assert ws.cell(13, COL_DESCRIPTION).value == "PTFE TEFLON TAPE"
    assert ws.cell(13, COL_QTY).value == 26880
    assert ws.cell(13, COL_UNIT).value == "PCS"
    assert ws.cell(13, COL_UNIT_PRICE_USD).value == pytest.approx(line.amount_usd / 26880)
    if formulas:
        assert ws.cell(13, COL_RATE_PER_UNIT).value == "=O13/H13"
    else:
        assert ws.cell(13, COL_RATE_PER_UNIT).value == line.land_cost_excl_gst / 26880


def test_excel_writes_an_ordinary_line_verbatim():
    ws = _sheet(_line("KEYCHAIN", 100.0, "KGS"), formulas=False)
    assert ws.cell(13, COL_DESCRIPTION).value == "KEYCHAIN"
    assert ws.cell(13, COL_QTY).value == 100.0
    assert ws.cell(13, COL_UNIT).value == "KGS"
    assert ws.cell(13, COL_UNIT_PRICE_USD).value == 0.8


def test_excel_keeps_unit_and_quantity_together_when_mts_is_unreadable():
    """A quantity that cannot be multiplied must not sit beside a ``KGS`` label."""
    ws = _sheet(_line("SILICONE SEALANT", RawValue.unparseable("1.O24"), "MTS"), formulas=False)
    assert ws.cell(13, COL_QTY).value == "1.O24"
    assert ws.cell(13, COL_UNIT).value == "MTS"
    assert ws.cell(13, COL_UNIT_PRICE_USD).value == 0.8


# ---------------------------------------------------------------------------
# Tally voucher: every path books the same quantity
# ---------------------------------------------------------------------------
def test_tally_books_an_unmapped_tape_line_in_pieces():
    doc = make_doc([_line(TAPE, 272.0, "KGS")])
    name, qty, rate, _amount = _booked(doc)
    assert name == "PTFE TEFLON TAPE"
    assert qty == " 26880.00 PCS"
    assert rate.endswith("/PCS")


def test_tally_keeps_the_piece_count_when_the_line_is_mapped_in_step_2():
    """The regression: mapping replaced the description and lost the bracket."""
    doc = make_doc([_line(TAPE, 272.0, "KGS")])
    unmapped = _booked(doc)
    mapped = _booked(apply_stock_names(doc, {1: "TEFLON TAPE 12MM"}))
    assert mapped[0] == "TEFLON TAPE 12MM"
    assert mapped[1:] == unmapped[1:] == (" 26880.00 PCS", unmapped[2], unmapped[3])


def test_tally_default_path_and_excel_upload_path_agree():
    doc = make_doc([_line(TAPE, 272.0, "KGS"), _line("SILICONE SEALANT", 1.024, "MTS", 2)])
    mapping = {1: "TEFLON TAPE 12MM", 2: "SILICONE SEALANT 280ML"}
    in_memory = TallyExporter().build(apply_stock_names(doc, mapping), USD_RATE)

    xlsx = ExcelGenerator(use_formulas=False).generate(doc, ReviewFlagSet([]))
    uploaded = TallyExporter().build(read_workbook(write_tally_names(xlsx, mapping)), USD_RATE)

    def rows(voucher):
        return [
            (e["stockitemname"], e["actualqty"], e["rate"], e["amount"])
            for e in _stock_entries(voucher)
        ]

    assert rows(in_memory) == rows(uploaded)
    assert [r[:2] for r in rows(in_memory)] == [
        ("TEFLON TAPE 12MM", " 26880.00 PCS"),
        ("SILICONE SEALANT 280ML", " 1024.00 KGS"),
    ]


def test_tally_never_rewrites_a_chosen_stock_name():
    """A Tally master may itself be called ``... TAPE (12PCS)``; it is a name,
    not an instruction to book 12 pieces."""
    doc = make_doc([_line("BOPP FILM", 100.0, "KGS")])
    name, qty, _rate, _amount = _booked(apply_stock_names(doc, {1: "CELLO TAPE (12PCS)"}))
    assert name == "CELLO TAPE (12PCS)"
    assert qty == " 100.00 KGS"


def test_tally_never_rewrites_a_stock_name_read_from_an_uploaded_workbook():
    doc = make_doc([_line(TAPE, 272.0, "KGS")])
    xlsx = ExcelGenerator(use_formulas=False).generate(doc, ReviewFlagSet([]))
    rebuilt = read_workbook(write_tally_names(xlsx, {1: "CELLO TAPE (12PCS)"}))
    name, qty, _rate, _amount = _booked(rebuilt)
    assert name == "CELLO TAPE (12PCS)"
    assert qty == " 26880.00 PCS"


def test_tally_converts_mts_for_a_line_built_without_the_calculator():
    """Documents assembled by hand (or by older code) still get MTS -> KGS."""
    line = make_line(1, 0.05, 1.024, "MTS", "32141000")
    assert line.stock_qty is None
    assert _booked(make_doc([line]))[1] == " 1024.00 KGS"


def test_tally_still_converts_dozens_to_pieces():
    assert _booked(make_doc([_line("SALO TAPE", 5.0, "DOZ")]))[1] == " 60.00 PCS"
