"""Cost-centre handling on the BOE -> Tally pipeline.

The cost centre (the consignment code, e.g. ``CO-04 CTN 1255``) is what ties the
purchase voucher to every downstream sales voucher for the same consignment - in
the CO-04 data all ten sales vouchers carry the purchase's cost centre. It is an
operator-supplied code: it appears nowhere on the Bill of Entry, and the parser
leaves ``header.details`` missing on the PDF path.

Two defects this covers:
  1. On the direct PDF path nothing could supply it, so the voucher went out with
     ``costcentrename: ""`` *and* allocations named ``""`` - Tally being asked to
     allocate to a cost centre with no name.
  2. Tally's own voucher allocates the party ledger to the cost centre; ours did
     not.

Allocation shape, from Tally's own export: entries carrying stock items allocate
per *item*; entries without (party, IGST purchase/payable) allocate at entry
level; Custom Duty Payable is not cost-centre tracked at all.
"""

from __future__ import annotations

import pytest

from boe_converter.models import RawValue
from boe_converter.tally_exporter import TallyExporter

from tests.test_tally_export import make_doc, make_line, rv

CC = "CO-04 CTN 1255"


def _doc(details=CC, rates=(0.05, 0.0)):
    return make_doc(
        [make_line(i + 1, r, 100, "KGS", "9") for i, r in enumerate(rates)],
        details=rv(details) if details else RawValue.missing(),
    )


def _v(doc, **kw):
    return TallyExporter().build(doc, 93.8, **kw)["tallymessage"][0]


def _alloc_names(voucher) -> list[str]:
    """Every cost-centre name the voucher allocates to, at any depth."""
    names: list[str] = []

    def walk(node):
        if isinstance(node, dict):
            for k, v in node.items():
                if k == "costcentreallocations":
                    names.extend(a.get("name", "") for a in v)
                else:
                    walk(v)
        elif isinstance(node, list):
            for x in node:
                walk(x)

    walk(voucher)
    return names


# ---------------------------------------------------------------------------
# The operator can supply the cost centre the document does not carry
# ---------------------------------------------------------------------------
def test_cost_centre_can_be_supplied_when_the_document_lacks_it():
    """The PDF path leaves `details` missing, so it must come from the caller."""
    v = _v(_doc(details=None), cost_centre=CC)
    assert v["costcentrename"] == CC
    assert set(_alloc_names(v)) == {CC}


def test_supplied_cost_centre_overrides_the_documents():
    v = _v(_doc(details="OLD CODE"), cost_centre=CC)
    assert v["costcentrename"] == CC
    assert set(_alloc_names(v)) == {CC}


def test_document_cost_centre_is_used_when_none_is_supplied():
    assert _v(_doc())["costcentrename"] == CC


def test_supplied_cost_centre_reaches_the_narration():
    v = _v(_doc(details=None), cost_centre=CC)
    assert v["narration"].startswith(CC + " USD ")


# ---------------------------------------------------------------------------
# Never ask Tally to allocate to a cost centre with no name
# ---------------------------------------------------------------------------
def test_no_cost_centre_means_no_allocations_at_all():
    v = _v(_doc(details=None))
    assert v["costcentrename"] == ""
    assert _alloc_names(v) == []
    assert v["iscostcentre"] is False


def test_blank_cost_centre_string_is_treated_as_absent():
    v = _v(_doc(details=None), cost_centre="   ")
    assert _alloc_names(v) == []
    assert v["iscostcentre"] is False


def test_cost_centre_present_sets_the_flag():
    assert _v(_doc())["iscostcentre"] is True


# ---------------------------------------------------------------------------
# Allocation shape matches Tally's own voucher
# ---------------------------------------------------------------------------
def test_party_ledger_is_allocated_to_the_cost_centre():
    """Tally's own voucher allocates the party ledger; ours did not."""
    party = next(e for e in _v(_doc())["allledgerentries"] if e.get("ispartyledger"))
    alloc = party["categoryallocations"][0]["costcentreallocations"][0]
    assert alloc["name"] == CC
    assert alloc["amount"] == party["amount"]


def test_stock_bearing_ledgers_allocate_per_item_not_at_entry_level():
    for e in _v(_doc())["allledgerentries"]:
        if e.get("inventoryallocations"):
            assert "categoryallocations" not in e
            for item in e["inventoryallocations"]:
                assert item["categoryallocations"][0]["costcentreallocations"][0][
                    "name"
                ] == CC


def test_custom_duty_is_not_cost_centre_tracked():
    duty = next(
        e for e in _v(_doc())["allledgerentries"] if "Custom Duty" in e["ledgername"]
    )
    assert "categoryallocations" not in duty


def test_igst_ledgers_allocate_at_entry_level():
    igst = [e for e in _v(_doc())["allledgerentries"] if "IGST" in e["ledgername"]]
    assert igst
    for e in igst:
        assert e["categoryallocations"][0]["costcentreallocations"][0]["name"] == CC


# ---------------------------------------------------------------------------
# Adding the party allocation must not unbalance the voucher
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("details", [CC, None])
def test_voucher_still_balances_exactly(details):
    from decimal import Decimal

    entries = _v(_doc(details=details))["allledgerentries"]
    assert sum(Decimal(e["amount"]) for e in entries) == Decimal("0.00")
