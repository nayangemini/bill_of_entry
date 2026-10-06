"""Shared unit-conversion tables.

A single source of truth for converting quantities:
- dozens/gross/thousand into pieces (PCS).
- metric tons (MTS) into kilograms (KGS).
- tape items carrying their piece count in the name, e.g.
  ``White Tape (6500pc)``, into PCS quantities.

Consumers (e.g. the Tally exporter, Excel writer) should import the tables
and helpers rather than defining their own copies.
"""

from __future__ import annotations

import re
from typing import NamedTuple

# Conversion factors to pieces (PCS). A unit not listed is not a
# piece-equivalent unit and is left unchanged by callers.
UNIT_TO_PCS: dict[str, float] = {"DOZ": 12.0, "GRS": 144.0, "THD": 1000.0}

# Conversion factors to kilograms (KGS). MTS (metric tons, case-insensitive)
# converts at 1000 KGS per MTS. A unit not listed is not a weight-equivalent
# unit and is left unchanged by callers.
UNIT_TO_KGS: dict[str, float] = {"MTS": 1000.0}


def pcs_factor(unit: str | None) -> float | None:
    """Return the pieces-per-unit factor for ``unit``, else ``None``.

    The unit is trimmed and upper-cased before lookup. Returns ``None`` when the
    unit is not a piece-equivalent unit (DOZ/GRS/THD) or is ``None``/empty.
    """
    if not unit:
        return None
    return UNIT_TO_PCS.get(unit.strip().upper())


def kgs_factor(unit: str | None) -> float | None:
    """Return the kilograms-per-unit factor for ``unit``, else ``None``.

    The unit is trimmed and upper-cased before lookup, so ``MTS``/``mts``/
    ``Mts`` all match. Returns ``None`` when the unit is not a
    weight-equivalent unit (MTS) or is ``None``/empty.
    """
    if not unit:
        return None
    return UNIT_TO_KGS.get(unit.strip().upper())


def convert_to_kgs(qty: float, unit: str) -> tuple[float, str]:
    """Convert a (qty, unit) to kilograms when the unit is MTS.

    Returns ``(converted_qty, converted_unit)``; unchanged for other units. The
    product is rounded to 6 decimals so binary floating point cannot leak into
    the quantity (``1.005 * 1000`` is ``1004.9999999999999``).
    """
    factor = kgs_factor(unit)
    if factor:
        return round(qty * factor, 6), "KGS"
    return qty, unit


# Units that declare a line by weight: the ICEGATE UQC weight codes plus the
# ``KG`` spelling supplier invoices use. The tape override below only applies
# to these - a line already declared in pieces keeps its BOE quantity.
WEIGHT_UNITS: frozenset[str] = frozenset({"KG", "KGS", "GMS", "MTS", "QTL", "TON"})

# Tape override: the BOE prints the weight (e.g. 456 KG) while the actual sale
# quantity rides in the item name in brackets, e.g. ``White Tape (6500pc)``.
# Tape-only, brackets-only: ``6600 PC`` without brackets never triggers, and
# ``tape`` must be a whole word (``TAPERED ROLLER BEARING`` is not tape).
_TAPE_WORD_RE = re.compile(r"\btapes?\b", re.IGNORECASE)
_TAPE_PCS_RE = re.compile(r"\(\s*(\d[\d,]*(?:\.\d+)?)\s*pcs?\s*\)", re.IGNORECASE)


def tape_pcs_override(
    description: str | None, unit: str | None
) -> tuple[str, float] | None:
    """Piece override for a weighed tape line named like ``White Tape (6500pc)``.

    Returns ``(base_name, qty)`` when ``description`` contains the word
    ``tape`` (case-insensitive) AND a bracketed ``(NNNpc)`` count AND ``unit``
    is a weight unit; else ``None``. ``base_name`` is the description with that
    bracket removed and whitespace collapsed; ``qty`` is the bracket number.
    Callers book ``qty`` PCS with ``rate = amount / qty`` and keep every
    monetary amount unchanged.
    """
    if not description or not unit or unit.strip().upper() not in WEIGHT_UNITS:
        return None
    if not _TAPE_WORD_RE.search(description):
        return None
    match = _TAPE_PCS_RE.search(description)
    if not match:
        return None
    try:
        qty = float(match.group(1).replace(",", ""))
    except ValueError:
        return None
    if qty <= 0:
        return None
    base = _TAPE_PCS_RE.sub(" ", description, count=1)
    base = re.sub(r"\s+", " ", base).strip(" -–—:,;")
    if not base:
        return None
    return base, qty


class StockQuantity(NamedTuple):
    """The quantity a line is booked under when it differs from the BOE's."""

    qty: float
    unit: str
    name: str | None = None  # replacement description (tape base name), if any


def stock_quantity(
    description: str | None, qty: float | None, unit: str | None
) -> StockQuantity | None:
    """Booking quantity for a line, or ``None`` when the BOE's own applies.

    The single place this is decided (the calculator stores the result on the
    ``ComputedLine``), so the Excel workbook, the Tally voucher and the Step-2
    editor cannot disagree:

    - a weighed tape line with a bracketed piece count books that count in PCS
      under the base name (:func:`tape_pcs_override`); this wins over MTS;
    - MTS books as KGS at 1000 per MTS (needs a numeric quantity);
    - anything else keeps the BOE quantity and unit (DOZ/GRS/THD are converted
      to pieces separately, via :func:`pcs_factor`).
    """
    tape = tape_pcs_override(description, unit)
    if tape is not None:
        name, pcs = tape
        return StockQuantity(pcs, "PCS", name)
    if qty is not None and kgs_factor(unit) is not None:
        return StockQuantity(*convert_to_kgs(qty, unit))
    return None
