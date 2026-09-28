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

    Returns ``(converted_qty, converted_unit)``; unchanged for other units.
    """
    factor = kgs_factor(unit)
    if factor:
        return qty * factor, "KGS"
    return qty, unit


# Tape override: the BOE prints the weight (e.g. 456 KG) while the actual sale
# quantity rides in the item name in brackets, e.g. ``White Tape (6500pc)``.
# Tape-only, brackets-only: ``6600 PC`` without brackets never triggers.
_TAPE_PCS_RE = re.compile(r"\(\s*(\d+(?:\.\d+)?)\s*pcs?\s*\)", re.IGNORECASE)


def tape_pcs_override(description: str | None) -> tuple[str, float] | None:
    """Piece override for tape items named like ``White Tape (6500pc)``.

    Returns ``(base_name, qty)`` when ``description`` contains ``tape``
    (case-insensitive) AND a bracketed ``(NNNpc)`` count; else ``None``.
    ``base_name`` is the description with that bracket removed and whitespace
    collapsed; ``qty`` is the bracket number. Callers display ``qty`` as PCS
    with ``rate = amount / qty`` and keep every monetary amount unchanged.
    """
    if not description or not description.strip():
        return None
    if "tape" not in description.lower():
        return None
    match = _TAPE_PCS_RE.search(description)
    if not match:
        return None
    try:
        qty = float(match.group(1).replace(",", ""))
    except ValueError:
        return None
    base = (_TAPE_PCS_RE.sub("", description, count=1))
    base = re.sub(r"\s+", " ", base).strip(" -–—:,;")
    base = re.sub(r"\s+", " ", base).strip()
    if not base:
        return None
    return base, qty
