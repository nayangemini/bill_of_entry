"""Shared unit-conversion table for piece-equivalent units.

A single source of truth for converting quantities expressed in
dozens/gross/thousand into pieces (PCS). Consumers (e.g. the Tally exporter)
should import :data:`UNIT_TO_PCS` and :func:`pcs_factor` rather than defining
their own copies.
"""

from __future__ import annotations

# Conversion factors to pieces (PCS). A unit not listed is not a
# piece-equivalent unit and is left unchanged by callers.
UNIT_TO_PCS: dict[str, float] = {"DOZ": 12.0, "GRS": 144.0, "THD": 1000.0}


def pcs_factor(unit: str | None) -> float | None:
    """Return the pieces-per-unit factor for ``unit``, else ``None``.

    The unit is trimmed and upper-cased before lookup. Returns ``None`` when the
    unit is not a piece-equivalent unit (DOZ/GRS/THD) or is ``None``/empty.
    """
    if not unit:
        return None
    return UNIT_TO_PCS.get(unit.strip().upper())
