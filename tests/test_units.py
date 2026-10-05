"""Unit tests for the shared unit-conversion rules in ``boe_converter.units``.

Validates: Requirements 14.6 (piece conversion); MTS -> KGS and the tape piece
override (PR #2).
"""

from __future__ import annotations

import pytest

from boe_converter.units import (
    UNIT_TO_PCS,
    convert_to_kgs,
    kgs_factor,
    pcs_factor,
    stock_quantity,
    tape_pcs_override,
)


def test_pcs_factor_known_units():
    assert pcs_factor("DOZ") == 12.0
    assert pcs_factor("GRS") == 144.0
    assert pcs_factor("THD") == 1000.0


@pytest.mark.parametrize("unit", ["NOS", "KGS", "PCS", "MTR", "XYZ"])
def test_pcs_factor_non_piece_units_return_none(unit):
    assert pcs_factor(unit) is None


@pytest.mark.parametrize("unit", [None, "", "   "])
def test_pcs_factor_empty_or_none_returns_none(unit):
    assert pcs_factor(unit) is None


@pytest.mark.parametrize(
    "unit,expected",
    [
        ("doz", 12.0),
        ("Doz", 12.0),
        ("  DOZ  ", 12.0),
        (" grs ", 144.0),
        ("thd", 1000.0),
    ],
)
def test_pcs_factor_is_trimmed_and_case_insensitive(unit, expected):
    assert pcs_factor(unit) == expected


def test_unit_to_pcs_table_contents():
    assert UNIT_TO_PCS == {"DOZ": 12.0, "GRS": 144.0, "THD": 1000.0}


# ---------------------------------------------------------------------------
# MTS -> KGS
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("unit", ["MTS", "mts", "Mts", "  mts "])
def test_kgs_factor_is_trimmed_and_case_insensitive(unit):
    assert kgs_factor(unit) == 1000.0


@pytest.mark.parametrize("unit", ["KGS", "KG", "PCS", "DOZ", None, "", "   "])
def test_kgs_factor_other_units_return_none(unit):
    assert kgs_factor(unit) is None


def test_convert_to_kgs_converts_mts_only():
    assert convert_to_kgs(1.024, "mts") == (1024.0, "KGS")
    assert convert_to_kgs(456.0, "KGS") == (456.0, "KGS")
    assert convert_to_kgs(10.0, "DOZ") == (10.0, "DOZ")


@pytest.mark.parametrize("mts,kgs", [(1.005, 1005.0), (1.001, 1001.0), (0.289, 289.0)])
def test_convert_to_kgs_has_no_float_artefacts(mts, kgs):
    """``1.005 * 1000`` is ``1004.9999999999999`` in binary floating point."""
    assert convert_to_kgs(mts, "MTS") == (kgs, "KGS")


# ---------------------------------------------------------------------------
# Tape piece-count override
# ---------------------------------------------------------------------------
@pytest.mark.parametrize(
    "description",
    [
        "White Tape (6600pc)",
        "White Tape (6600 pc)",
        "White Tape (6600PC)",
        "White Tape (6600 PCS)",
        "White Tape (6,600pc)",
    ],
)
def test_tape_override_accepts_bracket_spellings(description):
    assert tape_pcs_override(description, "KGS") == ("White Tape", 6600.0)


def test_tape_override_keeps_the_rest_of_the_name():
    """The real BOE line this rule was written for."""
    assert tape_pcs_override(
        "PTFE TEFLON TAPE (26880 pc)(O/T REPUTED BRAND)", "KGS"
    ) == ("PTFE TEFLON TAPE (O/T REPUTED BRAND)", 26880.0)


@pytest.mark.parametrize("unit", ["KG", "KGS", "kgs", " Kgs ", "MTS", "GMS"])
def test_tape_override_applies_to_weight_units(unit):
    assert tape_pcs_override("White Tape (6500pc)", unit) == ("White Tape", 6500.0)


@pytest.mark.parametrize("unit", ["DOZ", "PCS", "NOS", "ROL", "SET", None, ""])
def test_tape_override_needs_a_weight_unit(unit):
    """A line already declared in pieces keeps its quantity: the bracket may be
    a pack size (``MASKING TAPE (12PCS)``, 200 DOZ), not the consignment total."""
    assert tape_pcs_override("MASKING TAPE (12PCS)", unit) is None


@pytest.mark.parametrize(
    "description",
    [
        "TAPERED ROLLER BEARING (50PCS)",
        "TAPESTRY WALL HANGING (20pc)",
        "STAPLER (100pc)",
    ],
)
def test_tape_override_matches_the_whole_word_only(description):
    assert tape_pcs_override(description, "KGS") is None


@pytest.mark.parametrize("description", ["ADHESIVE TAPES (500pcs)", "bopp tape (72pc)"])
def test_tape_override_word_is_case_insensitive_and_allows_plural(description):
    assert tape_pcs_override(description, "KGS") is not None


@pytest.mark.parametrize(
    "description",
    [
        "50pc HAND GLOVES",
        "BEARING (23988M)",
        "White Tape 6600 PC",
        "HAND GLOVES (50pc)",
        "White Tape (0pc)",
        "(6500pc)",
        None,
        "",
    ],
)
def test_tape_override_leaves_other_descriptions_alone(description):
    assert tape_pcs_override(description, "KGS") is None


# ---------------------------------------------------------------------------
# stock_quantity: the one place a line's booking quantity is decided
# ---------------------------------------------------------------------------
def test_stock_quantity_is_none_for_an_ordinary_line():
    assert stock_quantity("KEYCHAIN", 100.0, "PCS") is None
    assert stock_quantity("SALO TAPE", 5.0, "DOZ") is None


def test_stock_quantity_converts_mts_to_kgs():
    assert stock_quantity("SILICONE SEALANT", 1.024, "MTS") == (1024.0, "KGS", None)


def test_stock_quantity_books_a_weighed_tape_line_in_pieces():
    assert stock_quantity("White Tape (6500pc)", 456.0, "KGS") == (6500.0, "PCS", "White Tape")


def test_stock_quantity_tape_wins_over_mts():
    assert stock_quantity("White Tape (6500pc)", 0.456, "MTS") == (6500.0, "PCS", "White Tape")


def test_stock_quantity_needs_a_numeric_quantity_to_convert_mts():
    """An unreadable quantity cannot be multiplied; unit and quantity stay together."""
    assert stock_quantity("SILICONE SEALANT", None, "MTS") is None
