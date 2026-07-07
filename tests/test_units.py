"""Unit tests for the shared piece-conversion table in ``boe_converter.units``.

Validates: Requirements 14.6
"""

from __future__ import annotations

import pytest

from boe_converter.units import UNIT_TO_PCS, pcs_factor


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
