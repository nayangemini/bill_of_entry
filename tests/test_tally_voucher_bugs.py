"""Regression tests for BUG-001..BUG-005 (Tally purchase-voucher fields).

Every one of these reproduces a defect observed by diffing a hand-entered Tally
voucher export against the converter's output for the same consignment, and each
fails against the pre-fix code. All data is synthetic.

Ground truth (Tally's own export, CO-49 CTN-1100 / invoice 202607044)::

    "narration":              "CO-49 CTN-1100 USD 31,453.57 @96.05"
    "countryofresidence":     "China"
    "placeofsupply":          "Maharashtra"
    "consigneestatename":     "Maharashtra"
    "gstregistrationtype":    "OIDAR"
    "cmpgstregistrationtype": "Regular"
"""

from __future__ import annotations

import pytest

from boe_converter.excel_reader import read_workbook
from boe_converter.excel_writer import ExcelGenerator
from boe_converter.models import RawValue, ReviewFlagSet
from boe_converter.tally_exporter import (
    CompanyProfile,
    SellerProfile,
    TallyExporter,
    company_profile_for,
    seller_profile_for,
    state_from_gstin,
)

from tests.test_tally_export import make_doc, make_line, rv


def _voucher(doc, rate: float = 96.05, **kw) -> dict:
    return TallyExporter(**kw).build(doc, rate)["tallymessage"][0]


# ---------------------------------------------------------------------------
# BUG-001 - narration must carry the real USD invoice total, comma-grouped
# ---------------------------------------------------------------------------
def test_bug001_narration_uses_real_usd_total_with_thousands_separator():
    """Tally prints ``USD 31,453.57``; we printed ``USD 0.00``."""
    # qty 3145.357 x $10/unit = $31,453.57, the ground-truth invoice total.
    doc = make_doc(
        [make_line(1, 0.05, 3145.357, "KGS", "9")], details=rv("CO-49 CTN-1100")
    )
    v = _voucher(doc)
    assert v["narration"] == "CO-49 CTN-1100 USD 31,453.57 @96.05"


def test_bug001_excel_roundtrip_preserves_usd_amount():
    """The Excel-upload path dropped column L, so the narration total was 0.00."""
    doc = make_doc(
        [make_line(1, 0.05, 200, "KGS", "9"), make_line(2, 0.05, 300, "KGS", "8")],
        details=rv("CO-49 CTN-1100"),
    )
    xlsx = ExcelGenerator(use_formulas=False).generate(doc, ReviewFlagSet([]))
    rebuilt = read_workbook(xlsx)

    expected = sum(line.amount_usd or 0.0 for line in doc.lines)  # 2000 + 3000
    assert rebuilt.lines[0].amount_usd == doc.lines[0].amount_usd
    assert rebuilt.totals.total_amount_usd == expected
    assert _voucher(rebuilt)["narration"] == "CO-49 CTN-1100 USD 5,000.00 @96.05"


def test_bug001_narration_groups_a_plain_thousand():
    """Guards the separator itself: 1000.0 must print as ``1,000.00``."""
    doc = make_doc([make_line(1, 0.05, 100, "KGS", "9")], details=rv("CO-49 CTN-1100"))
    assert _voucher(doc)["narration"] == "CO-49 CTN-1100 USD 1,000.00 @96.05"


# ---------------------------------------------------------------------------
# BUG-002 - supplier country must never default to India on an import BOE
# ---------------------------------------------------------------------------
def test_bug002_supplier_country_is_china_not_india():
    doc = make_doc([make_line(1, 0.05, 100, "KGS", "9")], seller_country=rv("China"))
    assert _voucher(doc)["countryofresidence"] == "China"


def test_bug002_unknown_supplier_country_is_not_silently_india():
    """The old code emitted ``India`` for *any* supplier whose country was absent."""
    doc = make_doc(
        [make_line(1, 0.05, 100, "KGS", "9")], seller_country=RawValue.missing()
    )
    assert _voucher(doc)["countryofresidence"] != "India"


def test_bug002_seller_profile_country_wins():
    doc = make_doc(
        [make_line(1, 0.05, 100, "KGS", "9")], seller_country=RawValue.missing()
    )
    v = _voucher(doc, seller=SellerProfile(country="China"))
    assert v["countryofresidence"] == "China"


# ---------------------------------------------------------------------------
# BUG-003 / BUG-004 - place of supply and ship-to state must be populated
# ---------------------------------------------------------------------------
def test_bug003_place_of_supply_derived_from_gstin_state_code():
    """GSTIN 27... *is* Maharashtra - decoding it is not inventing a value."""
    doc = make_doc(
        [make_line(1, 0.05, 100, "KGS", "9")],
        buyer_state=RawValue.missing(),
        buyer_gstin=rv("27AAYFG7003K1ZW"),
    )
    v = _voucher(doc)
    assert v["placeofsupply"] == "Maharashtra"
    assert v["cmpgststate"] == "Maharashtra"


def test_bug004_ship_to_state_derived_from_gstin_state_code():
    doc = make_doc(
        [make_line(1, 0.05, 100, "KGS", "9")],
        buyer_state=RawValue.missing(),
        buyer_gstin=rv("27AAYFG7003K1ZW"),
    )
    assert _voucher(doc)["consigneestatename"] == "Maharashtra"


def test_bug003_state_from_profile_gstin_on_excel_path():
    """Excel carries no identity, so a stored buyer profile must supply it."""
    doc = make_doc(
        [make_line(1, 0.05, 100, "KGS", "9")],
        buyer_state=RawValue.missing(),
        buyer_gstin=RawValue.missing(),
    )
    v = _voucher(doc, company=CompanyProfile(gstin="27AAYFG7003K1ZW"))
    assert v["placeofsupply"] == "Maharashtra"
    assert v["consigneestatename"] == "Maharashtra"


def test_bug003_explicit_state_still_wins_over_gstin():
    doc = make_doc(
        [make_line(1, 0.05, 100, "KGS", "9")],
        buyer_state=rv("Gujarat"),
        buyer_gstin=rv("27AAYFG7003K1ZW"),
    )
    assert _voucher(doc)["placeofsupply"] == "Gujarat"


def test_bug003_unknown_state_and_gstin_stays_blank_not_guessed():
    """The 'never invent' contract survives: nothing known -> nothing emitted."""
    doc = make_doc(
        [make_line(1, 0.05, 100, "KGS", "9")],
        buyer_state=RawValue.missing(),
        buyer_gstin=RawValue.missing(),
    )
    v = _voucher(doc)
    assert v["placeofsupply"] == ""
    assert v["consigneestatename"] == ""


# ---------------------------------------------------------------------------
# BUG-005 - GST registration type must be present
# ---------------------------------------------------------------------------
def test_bug005_gst_registration_type_is_oidar_for_overseas_supplier():
    doc = make_doc([make_line(1, 0.05, 100, "KGS", "9")], seller_country=rv("China"))
    v = _voucher(doc)
    assert v["gstregistrationtype"] == "OIDAR"
    assert v["cmpgstregistrationtype"] == "Regular"


def test_bug005_registration_type_present_even_when_country_unknown():
    """Every BOE is an import, so an unknown supplier country is still overseas."""
    doc = make_doc(
        [make_line(1, 0.05, 100, "KGS", "9")], seller_country=RawValue.missing()
    )
    assert _voucher(doc)["gstregistrationtype"] == "OIDAR"


def test_bug005_domestic_supplier_is_regular_not_oidar():
    doc = make_doc([make_line(1, 0.05, 100, "KGS", "9")], seller_country=rv("India"))
    assert _voucher(doc)["gstregistrationtype"] == "Regular"


def test_bug005_company_registration_type_is_overridable():
    doc = make_doc([make_line(1, 0.05, 100, "KGS", "9")])
    v = _voucher(doc, company=CompanyProfile(gst_registration_type="Composition"))
    assert v["cmpgstregistrationtype"] == "Composition"


# ---------------------------------------------------------------------------
# End-to-end: the exact voucher shape the bug reports describe
# ---------------------------------------------------------------------------
def test_all_five_fields_match_tally_on_the_excel_upload_path():
    doc = make_doc(
        [make_line(1, 0.05, 3145.357, "KGS", "9")], details=rv("CO-49 CTN-1100")
    )
    xlsx = ExcelGenerator(use_formulas=False).generate(doc, ReviewFlagSet([]))
    rebuilt = read_workbook(xlsx)

    # Excel carries no party identity; the stored profiles supply it.
    v = _voucher(
        rebuilt,
        company=CompanyProfile(gstin="27AAYFG7003K1ZW"),
        seller=SellerProfile(country="China"),
    )
    assert v["narration"] == "CO-49 CTN-1100 USD 31,453.57 @96.05"  # BUG-001
    assert v["countryofresidence"] == "China"                        # BUG-002
    assert v["placeofsupply"] == "Maharashtra"                       # BUG-003
    assert v["consigneestatename"] == "Maharashtra"                  # BUG-004
    assert v["gstregistrationtype"] == "OIDAR"                       # BUG-005


# ---------------------------------------------------------------------------
# GSTIN state decoding
# ---------------------------------------------------------------------------
@pytest.mark.parametrize(
    "gstin,state",
    [
        ("27AAYFG7003K1ZW", "Maharashtra"),
        ("24AAAAA0000A1Z5", "Gujarat"),
        ("07AAAAA0000A1Z5", "Delhi"),
        ("29AAAAA0000A1Z5", "Karnataka"),
        ("33AAAAA0000A1Z5", "Tamil Nadu"),
    ],
)
def test_state_from_gstin_decodes_the_state_code(gstin, state):
    assert state_from_gstin(gstin) == state


@pytest.mark.parametrize("bad", ["", None, "  ", "XXAAYFG7003K1ZW", "99AAAAA0000A1Z5"])
def test_state_from_gstin_returns_blank_for_undecodable_input(bad):
    assert state_from_gstin(bad) == ""


# ---------------------------------------------------------------------------
# Stored-record gap filling (what the Excel-upload path relies on)
# ---------------------------------------------------------------------------
class _Buyer:
    def __init__(self, gstin="", state="", pincode="", address_lines=(), name=""):
        self.gstin, self.state = gstin, state
        self.pincode, self.address_lines = pincode, list(address_lines)
        self.name = name


class _Seller:
    def __init__(self, country="", address_lines=()):
        self.country, self.address_lines = country, list(address_lines)


def _blank_header():
    return make_doc(
        [make_line(1, 0.05, 100, "KGS", "9")],
        buyer_gstin=RawValue.missing(),
        buyer_state=RawValue.missing(),
        buyer_pincode=RawValue.missing(),
        buyer_address=RawValue.missing(),
        seller_country=RawValue.missing(),
        seller_address=RawValue.missing(),
    ).header


def test_company_profile_fills_gaps_from_the_stored_buyer():
    prof = company_profile_for(
        _blank_header(),
        _Buyer(gstin="27AAYFG7003K1ZW", state="Maharashtra", pincode="400080"),
    )
    assert prof.gstin == "27AAYFG7003K1ZW"
    assert prof.state == "Maharashtra"
    assert prof.pincode == "400080"
    # No stored name -> nothing to override the document's with.
    assert prof.name is None


def test_stored_buyer_name_is_the_canonical_spelling():
    """A BOE shouts "M/S GEMINI UNICOM LLP"; Tally does not."""
    prof = company_profile_for(_blank_header(), _Buyer(name="Gemini Unicom LLP"))
    assert prof.name == "Gemini Unicom LLP"


def test_company_profile_never_overrides_what_the_boe_states():
    header = make_doc([make_line(1, 0.05, 100, "KGS", "9")]).header  # full identity
    prof = company_profile_for(header, _Buyer(gstin="24ZZZZZ0000Z1Z5", state="Gujarat"))
    assert prof.gstin is None and prof.state is None  # BOE wins


def test_seller_profile_fills_country_gap():
    assert seller_profile_for(_blank_header(), _Seller(country="China")).country == "China"


def test_seller_profile_never_overrides_a_stated_country():
    header = make_doc([make_line(1, 0.05, 100, "KGS", "9")]).header  # seller_country=China
    assert seller_profile_for(header, _Seller(country="Vietnam")).country is None


def test_profiles_are_empty_without_a_stored_record():
    assert company_profile_for(_blank_header(), None) == CompanyProfile()
    assert seller_profile_for(_blank_header(), None) == SellerProfile()


def test_gap_filled_profiles_produce_the_correct_voucher():
    """The full Excel-path chain: blank header + stored records -> Tally values."""
    doc = make_doc(
        [make_line(1, 0.05, 3145.357, "KGS", "9")],
        details=rv("CO-49 CTN-1100"),
        buyer_gstin=RawValue.missing(),
        buyer_state=RawValue.missing(),
        seller_country=RawValue.missing(),
    )
    v = _voucher(
        doc,
        company=company_profile_for(doc.header, _Buyer(gstin="27AAYFG7003K1ZW")),
        seller=seller_profile_for(doc.header, _Seller(country="China")),
    )
    assert v["narration"] == "CO-49 CTN-1100 USD 31,453.57 @96.05"
    assert v["countryofresidence"] == "China"
    assert v["placeofsupply"] == "Maharashtra"
    assert v["consigneestatename"] == "Maharashtra"
    assert v["cmpgststate"] == "Maharashtra"
    assert v["gstregistrationtype"] == "OIDAR"
    assert v["cmpgstregistrationtype"] == "Regular"
