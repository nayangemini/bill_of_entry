"""Regression tests for the Tally import exception:

    "The voucher numbering series selected is already used for another
     GST Registration."

Cause: the voucher asked Tally to allocate a number from the Purchase voucher
type's *automatic* numbering series while declaring no GST tax unit, so Tally
could not bind the series to this voucher's registration.

Ground truth: all eleven Tally-exported vouchers in the CO-04 data set (one
Purchase, ten Sales) use ``numberingstyle: "Manual"`` with an explicit
``vouchernumber``, ``vouchernumberseries: "Default"`` and
``vchstatustaxunit: "Maharashtra Registration"``. All data here is synthetic.
"""

from __future__ import annotations

import pytest

from boe_converter.models import RawValue
from boe_converter.tally_exporter import CompanyProfile, TallyExporter

from tests.test_tally_export import make_doc, make_line, rv


def _voucher(doc, rate: float = 93.8, **kw) -> dict:
    return TallyExporter(**kw).build(doc, rate)["tallymessage"][0]


def _doc(**over):
    return make_doc([make_line(1, 0.05, 100, "KGS", "9")], **over)


# ---------------------------------------------------------------------------
# Numbering: never ask Tally to allocate from the automatic series
# ---------------------------------------------------------------------------
def test_numbering_is_manual_not_auto():
    """``Auto`` is what triggers the numbering-series exception on import."""
    assert _voucher(_doc())["numberingstyle"] == "Manual"


def test_voucher_number_is_the_bill_of_entry_number():
    """Tally's own export numbers the purchase voucher with the BE number."""
    v = _voucher(_doc(be_no=rv("8668342")))
    assert v["vouchernumber"] == "8668342"


def test_voucher_number_series_is_default():
    assert _voucher(_doc())["vouchernumberseries"] == "Default"


def test_voucher_number_falls_back_to_the_invoice_number():
    v = _voucher(_doc(be_no=RawValue.missing(), invoice_no=rv("202603030")))
    assert v["vouchernumber"] == "202603030"
    assert v["numberingstyle"] == "Manual"


def test_auto_numbering_only_when_no_number_is_known_at_all():
    """With nothing to number it with, Tally must do the numbering itself."""
    v = _voucher(_doc(be_no=RawValue.missing(), invoice_no=RawValue.missing()))
    assert v["numberingstyle"] == "Auto"
    assert "vouchernumber" not in v


# ---------------------------------------------------------------------------
# Tax unit: bind the voucher to a GST registration
# ---------------------------------------------------------------------------
def test_voucher_declares_its_gst_tax_unit():
    v = _voucher(_doc())
    assert v["vchstatustaxunit"] == "Maharashtra Registration"
    assert v["gstregistration"] == {
        "value": "Maharashtra Registration",
        "taxtype": "GST",
        "taxregistration": "27AAYFG7003K1ZW",
    }


def test_tax_unit_name_is_overridable_per_company():
    """Another company may name its registration something else entirely."""
    v = _voucher(_doc(), company=CompanyProfile(tax_unit="Gujarat Registration"))
    assert v["vchstatustaxunit"] == "Gujarat Registration"
    assert v["gstregistration"]["value"] == "Gujarat Registration"


def test_tax_unit_derives_from_the_buyer_state():
    v = _voucher(_doc(buyer_state=rv("Gujarat"), buyer_gstin=rv("24AAAAA0000A1Z5")))
    assert v["vchstatustaxunit"] == "Gujarat Registration"


def test_no_tax_unit_emitted_when_the_state_is_unknown():
    """Never assert a registration we cannot substantiate."""
    v = _voucher(_doc(buyer_state=RawValue.missing(), buyer_gstin=RawValue.missing()))
    assert "vchstatustaxunit" not in v
    assert "gstregistration" not in v


def test_vchstatus_block_accompanies_the_tax_unit():
    v = _voucher(_doc())
    assert v["vchstatusvouchertype"] == "Purchase"
    assert v["vchstatustaxadjustment"] == "Default"
    assert v["vchstatusdate"] == v["date"]


# ---------------------------------------------------------------------------
# The exact shape Tally accepted, end to end
# ---------------------------------------------------------------------------
def test_matches_the_ground_truth_numbering_and_registration_block():
    v = _voucher(_doc(be_no=rv("8668342"), details=rv("CO-04 CTN 1255")))
    assert v["numberingstyle"] == "Manual"
    assert v["vouchernumber"] == "8668342"
    assert v["vouchernumberseries"] == "Default"
    assert v["vchstatustaxunit"] == "Maharashtra Registration"
    assert v["gstregistration"]["taxregistration"] == v["cmpgstin"]


# ---------------------------------------------------------------------------
# referencedate pairs with `reference` (the supplier invoice), not the BE
# ---------------------------------------------------------------------------
def test_reference_date_is_the_invoice_date_not_the_be_date():
    v = _voucher(_doc(invoice_date=rv("28/03/2026"), be_date=rv("14/04/2026")))
    assert v["reference"] == "ZJXY26050983"      # the supplier invoice number
    assert v["referencedate"] == "20260328"      # ...and its date
    assert v["date"] == "20260414"               # the BE date stays the voucher date


def test_reference_date_falls_back_to_the_be_date():
    v = _voucher(_doc(invoice_date=RawValue.missing(), be_date=rv("14/04/2026")))
    assert v["referencedate"] == "20260414"


# ---------------------------------------------------------------------------
# Dates as BOEs actually print them
# ---------------------------------------------------------------------------
@pytest.mark.parametrize(
    "text,expected",
    [
        ("14/04/2026", "20260414"),
        ("14-04-2026", "20260414"),
        ("28-MAR-26", "20260328"),      # the form this BOE prints its invoice date in
        ("28-Mar-2026", "20260328"),
        ("1-JAN-26", "20260101"),
        ("31-DEC-2026", "20261231"),
        ("", ""),
        ("not a date", ""),
        ("28-XYZ-26", ""),              # an unknown month is never guessed
    ],
)
def test_tally_date_parses_the_formats_boes_print(text, expected):
    from boe_converter.tally_exporter import tally_date

    assert tally_date(text) == expected


def test_reference_date_handles_an_alphabetic_month():
    """`28-MAR-26` fell through to the BE date, so the invoice date was lost."""
    v = _voucher(_doc(invoice_date=rv("28-MAR-26"), be_date=rv("14/04/2026")))
    assert v["referencedate"] == "20260328"
