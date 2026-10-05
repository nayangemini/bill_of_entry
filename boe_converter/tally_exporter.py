"""Build a Tally *Purchase voucher* JSON from a computed BOE document.

This turns a :class:`~boe_converter.models.ComputedDocument` (the same in-memory
result that drives the Excel workbook) into the JSON shape Tally imports - a
single ``tallymessage`` Purchase voucher whose line items are **grouped by IGST
rate** into purchase ledgers, each carrying its stock items as
``inventoryallocations`` plus the matching ``IGST Purchase``/``IGST Payable``
ledgers, a ``Custom Duty Payable`` ledger and the supplier *party* ledger.

Buyer (importer) and seller (supplier) identity are populated **from the BOE**
(the ``HeaderBlock`` buyer_*/seller_* fields the parser now extracts), with an
optional :class:`CompanyProfile` / :class:`SellerProfile` override for any field
the user wants to correct from stored data.

Ledger names come from a per-company :class:`LedgerBook` - they are hand-made
inside each company's Tally and cannot be derived (one company's Tally holds
``Igst Purchase @18.00%`` where a naming convention would give
``IGST Purchase @ 18.00 %``, and importing the wrong spelling creates a duplicate
ledger rather than posting to the real one). When a company has nothing stored the
conventional names (``Factory Purchase (Import 5%)``, ``IGST Purchase @ 5.00 %``,
``IGST Payable @ 5%``, ``Custom Duty Payable``, ``Tax Free (Purchases)``) are used
as a fallback and reported by :meth:`TallyExporter.derived_ledger_names`.

Accounting model (reverse-engineered from the reference voucher and expressed
purely in terms of computed per-line fields):

- **Purchase ledger** (grouped by IGST rate), debit: sum of
  ``land_cost_excl_gst`` for its lines; each stock item's amount is that line's
  ``land_cost_excl_gst`` and its rate is ``purchase_rate_per_unit``.
- **IGST Purchase** (debit) and **IGST Payable** (credit): equal and opposite,
  each the sum of ``igst_amount`` for the group (zero-rate groups have none).
- **Party** ledger (supplier), credit: total ``purchase_inr`` (the CIF value
  paid to the supplier = USD invoice x rate).
- **Custom Duty Payable**, credit: total ``total_customs_duty``.

The identity ``sum(land_cost_excl_gst) == sum(purchase_inr) + sum(total_customs_duty)``
keeps the voucher balanced.

Unit conversion: quantities printed in dozens/gross/thousand are converted to
pieces on the inventory allocation (DOZ x12, GRS x144, THD x1000) so Tally
receives PCS. The line amount is preserved; only qty/unit/rate change.
"""

from __future__ import annotations

import uuid
from dataclasses import dataclass, field, replace

from boe_converter.models import ComputedDocument, ComputedLine, HeaderBlock, RawValue
from boe_converter.units import UNIT_TO_PCS, convert_to_kgs, pcs_factor

# ---------------------------------------------------------------------------
# Pure Tally structural constants (never printed on a BOE)
# ---------------------------------------------------------------------------
NOT_APPLICABLE = "\u0004 Not Applicable"
COST_CATEGORY = "Primary Cost Category"
GODOWN = "Main Location"
BATCH = "Primary Batch"
DEFAULT_ENTERED_BY = "boe-converter"
CUSTOM_DUTY_LEDGER = "Custom Duty Payable"
TAX_FREE_LEDGER = "Tax Free (Purchases)"

# Unit conversion factors to pieces (PCS). A unit not listed is left unchanged.
# Sourced from the shared table in ``boe_converter.units``.
_UNIT_TO_PCS = UNIT_TO_PCS

# GST registration types Tally accepts on the voucher. A Bill of Entry is an
# *import* document, so its supplier is an overseas party - Tally records those
# as ``OIDAR``. Only a supplier explicitly located in India is ``Regular``.
REG_TYPE_OVERSEAS = "OIDAR"
REG_TYPE_REGULAR = "Regular"

# The first two digits of a GSTIN are the registered state's GST state code.
# Reading them back is *decoding* an existing value, not inventing a missing one,
# which is what lets the voucher carry a place of supply when the BOE/Excel does
# not spell the state out. Names match Tally's own state list.
_GSTIN_STATE_CODES: dict[str, str] = {
    "01": "Jammu & Kashmir",
    "02": "Himachal Pradesh",
    "03": "Punjab",
    "04": "Chandigarh",
    "05": "Uttarakhand",
    "06": "Haryana",
    "07": "Delhi",
    "08": "Rajasthan",
    "09": "Uttar Pradesh",
    "10": "Bihar",
    "11": "Sikkim",
    "12": "Arunachal Pradesh",
    "13": "Nagaland",
    "14": "Manipur",
    "15": "Mizoram",
    "16": "Tripura",
    "17": "Meghalaya",
    "18": "Assam",
    "19": "West Bengal",
    "20": "Jharkhand",
    "21": "Odisha",
    "22": "Chhattisgarh",
    "23": "Madhya Pradesh",
    "24": "Gujarat",
    "25": "Daman & Diu",
    "26": "Dadra & Nagar Haveli & Daman & Diu",
    "27": "Maharashtra",
    "28": "Andhra Pradesh",
    "29": "Karnataka",
    "30": "Goa",
    "31": "Lakshadweep",
    "32": "Kerala",
    "33": "Tamil Nadu",
    "34": "Puducherry",
    "35": "Andaman & Nicobar Islands",
    "36": "Telangana",
    "37": "Andhra Pradesh",
    "38": "Ladakh",
    "97": "Other Territory",
}


def state_from_gstin(gstin: str | None) -> str:
    """The registered state encoded in a GSTIN's first two digits ("" if none)."""
    if not gstin:
        return ""
    code = gstin.strip()[:2]
    return _GSTIN_STATE_CODES.get(code, "") if code.isdigit() else ""


def _party_gst_registration_type(seller_country: str) -> str:
    """The supplier's GST registration type for the voucher.

    Every Bill of Entry is an import, so an unknown supplier country still means
    an overseas party; only a supplier explicitly in India is ``Regular``.
    """
    if seller_country.strip().lower() == "india":
        return REG_TYPE_REGULAR
    return REG_TYPE_OVERSEAS


# ---------------------------------------------------------------------------
# Filling identity a document cannot carry
# ---------------------------------------------------------------------------
# A re-uploaded CTN workbook has no GSTIN / state / country cells - the sheet is
# a costing table, not an identity record - so on that path the voucher's place
# of supply, ship-to state and supplier country came out blank. These build a
# profile from a stored buyer/seller record that fills **only the gaps**: any
# field the document itself states is left alone, so the BOE always wins.
def company_profile_for(header: HeaderBlock, record) -> CompanyProfile:
    """A :class:`CompanyProfile` filling what ``header`` does not already state.

    ``record`` is any object exposing ``gstin``/``state``/``pincode``/
    ``address_lines`` (e.g. a ``tally_store.BuyerRecord``). The record is matched
    *by* the document's own company name, so its name is the same entity spelled
    canonically - and a BOE shouts ("M/S GEMINI UNICOM LLP") where Tally does
    not. The stored spelling therefore wins.
    """
    if record is None:
        return CompanyProfile()
    lines = tuple(getattr(record, "address_lines", ()) or ())
    return CompanyProfile(
        name=(getattr(record, "name", "") or "").strip() or None,
        gstin=_gap(header.buyer_gstin, getattr(record, "gstin", "")),
        state=_gap(header.buyer_state, getattr(record, "state", "")),
        pincode=_gap(header.buyer_pincode, getattr(record, "pincode", "")),
        address_lines=lines or None if not _text(header.buyer_address) else None,
        # No document carries these, so they always come from the record.
        tax_unit=(getattr(record, "tax_unit", "") or "").strip() or None,
        tally_name=(getattr(record, "tally_name", "") or "").strip() or None,
    )


def seller_profile_for(header: HeaderBlock, record) -> SellerProfile:
    """A :class:`SellerProfile` filling what ``header`` does not already state.

    As with the buyer, the stored supplier name is the canonical one - it *is*
    the party ledger's name in this company's Tally - so it wins over the BOE's.
    """
    if record is None:
        return SellerProfile()
    lines = tuple(getattr(record, "address_lines", ()) or ())
    return SellerProfile(
        name=(getattr(record, "name", "") or "").strip() or None,
        country=_gap(header.seller_country, getattr(record, "country", "")),
        address_lines=lines or None if not _text(header.seller_address) else None,
    )


def _gap(field, stored: str) -> str | None:
    """``stored`` when the document's ``field`` is absent, else ``None``."""
    if _text(field):
        return None
    return (stored or "").strip() or None


@dataclass(frozen=True)
class CompanyProfile:
    """Buyer/company identity - the Tally company the voucher is imported into.

    Every field defaults to ``None`` meaning "use the value extracted from the
    BOE". Supply a value only to override the BOE (e.g. from stored data). This
    is how the buyer is populated *from the document* while still allowing a
    manual correction.
    """

    name: str | None = None
    gstin: str | None = None
    state: str | None = None
    pincode: str | None = None
    address_lines: tuple[str, ...] | None = None
    entered_by: str = DEFAULT_ENTERED_BY
    # The importing company's own GST registration type (``cmpgstregistrationtype``).
    gst_registration_type: str = REG_TYPE_REGULAR
    # The name of this company's GST registration (tax unit) in Tally, e.g.
    # "Maharashtra Registration". Defaults to "<state> Registration", which is
    # how Tally names a state registration unless the company renamed it.
    tax_unit: str | None = None
    # The Tally *company* name as it appears on the voucher's buyer line, which
    # may carry a financial-year suffix (e.g. "Gemini Unicom LLP (F.Y. 2026-27)")
    # while the mailing name stays plain. Defaults to the buyer name.
    tally_name: str | None = None


@dataclass(frozen=True)
class SellerProfile:
    """Supplier/seller identity override (defaults to the BOE-extracted values)."""

    name: str | None = None
    address_lines: tuple[str, ...] | None = None
    country: str | None = None


# ---------------------------------------------------------------------------
# Number / quantity formatting (match the reference voucher's string forms)
# ---------------------------------------------------------------------------
def _amt(x: float) -> str:
    """Format a monetary amount as a 2-decimal string (e.g. ``-282818.99``)."""
    return f"{x:.2f}"


def _qty(q: float, unit: str) -> str:
    """Format a quantity as `` 4451.00 KGS`` (leading space, 2 decimals)."""
    return f" {q:.2f} {unit}"


def _rate(r: float, unit: str) -> str:
    """Format a unit rate as ``63.54/KGS``."""
    return f"{r:.2f}/{unit}"


def _pct(rate_fraction: float) -> float:
    """Convert a stored IGST fraction (0.05) to a percent number (5.0)."""
    return round(rate_fraction * 100, 2)


def _pct_label(rate_fraction: float) -> str:
    """A human percent label: ``5`` for whole, ``2.5`` for fractional."""
    p = _pct(rate_fraction)
    return str(int(p)) if float(p).is_integer() else ("%g" % p)


def _guid() -> str:
    return f"{uuid.uuid4()}-{uuid.uuid4().hex[:8]}"


def _convert_to_pcs(qty: float, unit: str) -> tuple[float, str]:
    """Convert a (qty, unit) to pieces when the unit is DOZ/GRS/THD.

    Returns ``(converted_qty, converted_unit)``; unchanged for other units.
    """
    factor = pcs_factor(unit)
    if factor:
        return qty * factor, "PCS"
    return qty, unit


def _convert_qty_unit(qty: float, unit: str) -> tuple[float, str]:
    """Convert MTS to KGS first, then DOZ/GRS/THD to PCS.

    MTS (any case) becomes KGS at 1000 per MTS so Tally receives the stock
    unit; all other units keep existing behaviour. Amount is preserved by
    callers; only qty/unit/rate change.
    """
    qty, unit = convert_to_kgs(qty, unit)
    return _convert_to_pcs(qty, unit)


# ---------------------------------------------------------------------------
# Ledger-name conventions (deterministic; no master file needed)
# ---------------------------------------------------------------------------
def _purchase_ledger_name(rate_fraction: float) -> str:
    return f"Factory Purchase (Import {_pct_label(rate_fraction)}%)"


def _igst_purchase_ledger_name(rate_fraction: float) -> str:
    return f"IGST Purchase @ {_pct(rate_fraction):.2f} %"


def _igst_payable_ledger_name(rate_fraction: float) -> str:
    return f"IGST Payable @ {_pct_label(rate_fraction)}%"


def _cost_centre_block(cost_centre: str, amount: float, deemed_positive: bool) -> dict:
    """``{"categoryallocations": [...]}`` for ``cost_centre``, or ``{}`` when absent.

    Returning an empty dict lets callers splat this into an entry, so a voucher
    with no cost centre carries no allocations at all rather than allocating to
    one with a blank name - which is not something Tally can accept.
    """
    if not cost_centre:
        return {}
    return {
        "categoryallocations": [
            {
                "category": COST_CATEGORY,
                "isdeemedpositive": deemed_positive,
                "costcentreallocations": [
                    {"name": cost_centre, "amount": _amt(amount)}
                ],
            }
        ]
    }


def _party_ledger_name(supplier: str) -> str:
    """Tally party ledger name: title-cased supplier name (matches reference)."""
    return supplier.title()


# Ledger kinds a purchase voucher needs. ``PURCHASE``/``IGST_PURCHASE``/
# ``IGST_PAYABLE`` exist once per IGST rate; the other two are single ledgers.
KIND_PURCHASE = "purchase"
KIND_IGST_PURCHASE = "igst_purchase"
KIND_IGST_PAYABLE = "igst_payable"
KIND_CUSTOM_DUTY = "custom_duty"
KIND_TAX_FREE = "tax_free"

LEDGER_KINDS = (
    KIND_PURCHASE,
    KIND_IGST_PURCHASE,
    KIND_IGST_PAYABLE,
    KIND_CUSTOM_DUTY,
    KIND_TAX_FREE,
)


def _bp(rate_fraction: float) -> int:
    """An IGST rate as integer basis points (0.05 -> 500).

    Ledgers are keyed on this rather than the float so a lookup never depends on
    floating-point equality.
    """
    return round(rate_fraction * 10000)


@dataclass(frozen=True)
class LedgerBook:
    """The ledger names a *specific* Tally company actually uses.

    Ledger names cannot be derived. One company's Tally holds
    ``Igst Purchase @18.00%`` where the naming convention would produce
    ``IGST Purchase @ 18.00 %``; importing the derived name creates a second,
    duplicate ledger instead of posting to the real one. So every name is looked
    up here first, keyed by ``(kind, basis points)``.

    ``names`` maps ``(kind, rate_bp)`` to the exact Tally ledger name; the two
    single ledgers use a rate of ``0``. Anything absent falls back to the naming
    convention so a company with nothing stored still produces a voucher - but
    :meth:`TallyExporter.derived_ledger_names` reports every such fallback so a
    derived name is a visible choice, never a silent duplicate.
    """

    names: dict[tuple[str, int], str] = field(default_factory=dict)

    def get(self, kind: str, rate_fraction: float = 0.0) -> str | None:
        return self.names.get((kind, _bp(rate_fraction)))

    def resolve(self, kind: str, rate_fraction: float = 0.0) -> str:
        """The stored name, else the conventional one."""
        return self.get(kind, rate_fraction) or _conventional_ledger_name(
            kind, rate_fraction
        )

    @classmethod
    def from_records(cls, records) -> "LedgerBook":
        """Build from anything exposing ``kind``, ``rate_bp`` and ``name``."""
        return cls(
            {
                (r.kind, int(r.rate_bp)): r.name
                for r in records
                if getattr(r, "name", "").strip()
            }
        )


def _conventional_ledger_name(kind: str, rate_fraction: float) -> str:
    """The name Tally's own defaults would give a ledger of this kind."""
    if kind == KIND_PURCHASE:
        return _purchase_ledger_name(rate_fraction)
    if kind == KIND_IGST_PURCHASE:
        return _igst_purchase_ledger_name(rate_fraction)
    if kind == KIND_IGST_PAYABLE:
        return _igst_payable_ledger_name(rate_fraction)
    if kind == KIND_CUSTOM_DUTY:
        return CUSTOM_DUTY_LEDGER
    if kind == KIND_TAX_FREE:
        return TAX_FREE_LEDGER
    raise ValueError(f"unknown ledger kind: {kind!r}")


# ---------------------------------------------------------------------------
# Value extraction helpers
# ---------------------------------------------------------------------------
def _num(rv) -> float | None:
    """Numeric interpretation of a RawValue-like field, else None."""
    if rv is None or getattr(rv, "is_missing", False) or getattr(rv, "is_unparseable", False):
        return None
    parsed = getattr(rv, "parsed", None)
    if isinstance(parsed, bool):
        return None
    if isinstance(parsed, (int, float)):
        return float(parsed)
    return None


def _text(rv) -> str | None:
    """Text interpretation of a RawValue-like field, else None."""
    if rv is None or getattr(rv, "is_missing", False):
        return None
    parsed = getattr(rv, "parsed", None)
    if isinstance(parsed, str) and parsed.strip():
        return parsed.strip()
    raw = getattr(rv, "raw_text", None)
    return raw.strip() if isinstance(raw, str) and raw.strip() else None


def _split_address(text: str | None) -> list[str]:
    """Split a comma-joined address into individual lines (empty when absent)."""
    if not text:
        return []
    return [part.strip() for part in text.split(",") if part.strip()]


def _line_igst_fraction(line: ComputedLine) -> float:
    """The IGST rate fraction for a line (0.0 when missing/unreadable)."""
    r = _num(line.source.igst_rate)
    return r if r is not None else 0.0


def _stock_name(line: ComputedLine) -> str:
    """Stock item name: the BOE description (Excel col D is filled by a human).

    When a mapped Tally name is later supplied via the Excel upload path it can
    override this; from the in-memory document the verbatim description is used.
    The name is never rewritten here: by this point it may be a Tally master
    name chosen in Step 2.
    """
    return _text(line.source.description) or f"ITEM {line.source.item_serial}"


def _booking_qty_unit(line: ComputedLine) -> tuple[float, str]:
    """The ``(qty, unit)`` a line is booked under in Tally.

    A line the calculator re-expressed (MTS -> KGS, a weighed tape line -> its
    piece count in PCS) carries that on ``stock_qty``/``stock_unit``. It is read
    from the line, not re-derived from the description, because Step 2 replaces
    the description with the mapped Tally name. Any other line books its BOE
    quantity, converting MTS to KGS and dozens/gross/thousand to pieces.
    """
    if line.stock_qty is not None and line.stock_unit:
        return line.stock_qty, line.stock_unit
    raw_unit = _text(line.source.unit) or "NOS"
    raw_qty = _num(line.source.quantity) or 0.0
    return _convert_qty_unit(raw_qty, raw_unit)


# ---------------------------------------------------------------------------
# Exporter
# ---------------------------------------------------------------------------
class TallyExporter:
    """Builds a Purchase voucher ``tallymessage`` document.

    Buyer/seller identity comes from the BOE ``HeaderBlock`` by default; pass a
    :class:`CompanyProfile` / :class:`SellerProfile` to override individual
    fields (e.g. from stored data).
    """

    def __init__(
        self,
        company: CompanyProfile | None = None,
        seller: SellerProfile | None = None,
        ledgers: LedgerBook | None = None,
    ) -> None:
        self.company = company or CompanyProfile()
        self.seller = seller or SellerProfile()
        # This company's real Tally ledger names; empty means "use the
        # convention", which derived_ledger_names() then reports.
        self.ledgers = ledgers or LedgerBook()

    def _ledger(self, kind: str, rate_fraction: float = 0.0) -> str:
        return self.ledgers.resolve(kind, rate_fraction)

    # -- public API ---------------------------------------------------------
    def required_ledger_names(self, computed: ComputedDocument) -> list[str]:
        """Every ledger name this voucher will reference (for UI display)."""
        names: list[str] = [self._party_name(computed)]
        for r in _distinct_rates(computed):
            if r <= 0:
                names.append(self._ledger(KIND_TAX_FREE))
                continue
            names.append(self._ledger(KIND_PURCHASE, r))
            names.append(self._ledger(KIND_IGST_PURCHASE, r))
            names.append(self._ledger(KIND_IGST_PAYABLE, r))
        names.append(self._ledger(KIND_CUSTOM_DUTY))
        seen: set[str] = set()
        out: list[str] = []
        for n in names:
            if n not in seen:
                seen.add(n)
                out.append(n)
        return out

    def derived_ledger_names(self, computed: ComputedDocument) -> list[tuple[str, str]]:
        """``(kind, name)`` for ledgers whose name was *derived*, not stored.

        A derived name may not exist in the target company's Tally, in which case
        the import silently creates a duplicate ledger. Surfacing these lets the
        user store the real name instead.
        """
        wanted: list[tuple[str, float]] = []
        for r in _distinct_rates(computed):
            if r <= 0:
                wanted.append((KIND_TAX_FREE, 0.0))
                continue
            wanted.append((KIND_PURCHASE, r))
            if any(abs(l.igst_amount or 0.0) > 0 for l in computed.lines
                   if _same_rate(_line_igst_fraction(l), r)):
                wanted.append((KIND_IGST_PURCHASE, r))
                wanted.append((KIND_IGST_PAYABLE, r))
        wanted.append((KIND_CUSTOM_DUTY, 0.0))

        out: list[tuple[str, str]] = []
        seen: set[str] = set()
        for kind, rate in wanted:
            if self.ledgers.get(kind, rate) is not None:
                continue
            name = _conventional_ledger_name(kind, rate)
            if name not in seen:
                seen.add(name)
                out.append((kind, name))
        return out

    def build(
        self,
        computed: ComputedDocument,
        usd_rate: float,
        voucher_date: str | None = None,
        cost_centre: str | None = None,
    ) -> dict:
        """Build the full Tally import document for ``computed``.

        ``voucher_date`` (``YYYYMMDD``) is the date the entry is *booked* under,
        which is a bookkeeping decision rather than a fact on the document; it
        defaults to the BE date. The supplier invoice's own date always stays on
        ``referencedate``.

        ``cost_centre`` is the consignment code (e.g. ``CO-04 CTN 1255``) that
        ties this purchase to the sales vouchers raised against the same
        consignment. It appears nowhere on the Bill of Entry, so on the PDF path
        only the caller can supply it; it defaults to the workbook's ``Details``
        cell. When it is unknown the voucher carries no cost-centre allocations
        at all rather than allocating to a nameless one.
        """
        party_ledger = self._party_ledger(computed)
        cost_centre = (cost_centre or _text(computed.header.details) or "").strip()
        usd_total = computed.totals.total_amount_usd
        duty_total = sum(l.total_customs_duty or 0.0 for l in computed.lines)

        narration = self._narration(cost_centre, usd_total, usd_rate)

        # 1) Purchase + IGST ledgers, grouped by IGST rate. These come first
        #    because the party amount is derived from them (see below), even
        #    though the party entry is listed first in the voucher.
        ledger_entries: list[dict] = []
        for rate in _distinct_rates(computed):
            group = [l for l in computed.lines if _same_rate(_line_igst_fraction(l), rate)]
            if rate <= 0:
                ledger_entries.append(self._tax_free_entry(group, cost_centre))
                continue
            ledger_entries.append(self._purchase_entry(rate, group, cost_centre))
            igst_sum = sum(l.igst_amount or 0.0 for l in group)
            if abs(igst_sum) > 0:
                ledger_entries.append(self._igst_purchase_entry(rate, igst_sum, cost_centre))
                ledger_entries.append(self._igst_payable_entry(rate, igst_sum, cost_centre))

        # 2) Custom Duty Payable - credit.
        duty_entry = self._simple_credit(
            self._ledger(KIND_CUSTOM_DUTY), duty_total, cost_centre=None
        )
        ledger_entries.append(duty_entry)

        # 3) Party (supplier) ledger - credit, carries the bill reference.
        #    Its amount is derived from the entries above rather than summed
        #    independently. ``land_cost_excl_gst == purchase_inr + duty`` holds in
        #    full precision, but rounding each aggregate to paise on its own let
        #    the two sides drift a paisa apart and the voucher did not balance -
        #    invalid double entry. The purchase ledgers cannot absorb the residue
        #    (each must equal the sum of its own stock-item amounts) and the duty
        #    is a declared customs figure, so the supplier payable - which the
        #    identity defines anyway - is the balancing figure. IGST purchase and
        #    payable are equal and opposite, so they never affect this.
        party_total = -sum(float(e["amount"]) for e in ledger_entries)
        ledger_entries.insert(
            0,
            self._party_entry(party_ledger, party_total, computed, cost_centre),
        )

        voucher = self._voucher_shell(
            computed, party_ledger, narration, cost_centre, voucher_date
        )
        voucher["allledgerentries"] = ledger_entries
        return {"tallymessage": [voucher]}

    # -- buyer / seller resolution -----------------------------------------
    def _buyer_name(self, computed: ComputedDocument) -> str:
        if self.company.name:
            return self.company.name
        name = computed.header.company_name or ""
        # The BOE importer name carries an "M/S " prefix for the Excel sheet;
        # strip it for the Tally buyer so it matches the company master.
        if name[:4].upper() == "M/S ":
            name = name[4:].strip()
        return name

    def _buyer_gstin(self, computed: ComputedDocument) -> str:
        return self.company.gstin or _text(computed.header.buyer_gstin) or ""

    def _buyer_state(self, computed: ComputedDocument) -> str:
        """The buyer's state: stated if known, else decoded from their GSTIN.

        The Excel-upload path carries no state column, so without the GSTIN
        fallback ``placeofsupply``/``consigneestatename``/``cmpgststate`` all came
        out blank. When neither a state nor a GSTIN is known this still returns
        "" - a missing value is never fabricated.
        """
        return (
            self.company.state
            or _text(computed.header.buyer_state)
            or state_from_gstin(self._buyer_gstin(computed))
        )

    def _buyer_pincode(self, computed: ComputedDocument) -> str:
        return self.company.pincode or _text(computed.header.buyer_pincode) or ""

    def _tax_unit(self, computed: ComputedDocument) -> str:
        """This company's GST registration (tax unit) name in Tally.

        Tally binds a voucher numbering series to a GST registration, so a
        voucher that declares no tax unit cannot be numbered - that is the
        "numbering series is already used for another GST Registration" import
        exception. Defaults to Tally's own "<state> Registration" naming and is
        overridable per company; returns "" when the state is unknown, in which
        case no registration is asserted at all.
        """
        if self.company.tax_unit:
            return self.company.tax_unit
        state = self._buyer_state(computed)
        return f"{state} Registration" if state else ""

    def _voucher_number(self, computed: ComputedDocument) -> str:
        """The number Tally files this voucher under: the BE number.

        Tally's own vouchers for these consignments are numbered with the Bill
        of Entry number (falling back to the invoice number), entered manually.
        """
        return (
            _text(computed.header.be_no) or _text(computed.header.invoice_no) or ""
        )

    def _buyer_address_lines(self, computed: ComputedDocument) -> list[str]:
        if self.company.address_lines is not None:
            return list(self.company.address_lines)
        return _split_address(_text(computed.header.buyer_address))

    def _seller_address_lines(self, computed: ComputedDocument) -> list[str]:
        if self.seller.address_lines is not None:
            return list(self.seller.address_lines)
        return _split_address(_text(computed.header.seller_address))

    def _seller_country(self, computed: ComputedDocument) -> str:
        return self.seller.country or _text(computed.header.seller_country) or ""

    # -- party --------------------------------------------------------------
    def _party_name(self, computed: ComputedDocument) -> str:
        return self.seller.name or _text(computed.header.party_name) or "Unknown Supplier"

    def _party_ledger(self, computed: ComputedDocument) -> str:
        """The supplier's ledger name in this company's Tally.

        A stored seller name is used verbatim - it *is* the ledger name, and
        companies differ on casing (one Tally holds
        "Prayan Impex Company Limited", another "DIA IMPEX COMPANY LIMITED").
        Only without a stored record is the BOE's name title-cased.
        """
        if self.seller.name:
            return self.seller.name
        return _party_ledger_name(self._party_name(computed))

    def _party_entry(
        self,
        ledger: str,
        total: float,
        computed: ComputedDocument,
        cost_centre: str = "",
    ) -> dict:
        bill_ref = _text(computed.header.invoice_no) or _text(computed.header.be_no) or "Ref"
        return {
            "oldauditentryids": [{"metadata": True, "type": "Number"}, "-1"],
            "ledgername": ledger,
            "gstclass": NOT_APPLICABLE,
            "isdeemedpositive": False,
            "ledgerfromitem": False,
            "removezeroentries": False,
            "ispartyledger": True,
            "amount": _amt(total),
            "vatexpamount": _amt(total),
            **_cost_centre_block(cost_centre, total, False),
            "billallocations": [
                {
                    "name": bill_ref,
                    "billtype": "New Ref",
                    "tdsdeducteeisspecialrate": False,
                    "amount": _amt(total),
                }
            ],
        }

    # -- purchase (taxable, grouped) ---------------------------------------
    def _purchase_entry(self, rate: float, group: list[ComputedLine], cost_centre: str) -> dict:
        ledger = self._ledger(KIND_PURCHASE, rate)
        total = sum(l.land_cost_excl_gst or 0.0 for l in group)
        return {
            "oldauditentryids": [{"metadata": True, "type": "Number"}, "-1"],
            "ledgername": ledger,
            "gstclass": NOT_APPLICABLE,
            "gstovrdnineligibleitc": NOT_APPLICABLE,
            "gstovrdnisrevchargeappl": NOT_APPLICABLE,
            "gstovrdntypeofsupply": "Goods",
            "gstrateinferapplicability": "As per Masters/Company",
            "gsthsninferapplicability": "As per Masters/Company",
            "isdeemedpositive": True,
            "ledgerfromitem": False,
            "removezeroentries": False,
            "ispartyledger": False,
            "islastdeemedpositive": True,
            "amount": _amt(-total),
            "vatexpamount": _amt(-total),
            "inventoryallocations": [
                self._inventory(l, rate, cost_centre) for l in group
            ],
        }

    def _inventory(self, line: ComputedLine, rate: float, cost_centre: str) -> dict:
        name = _stock_name(line)
        hsn = _text(line.source.cth_hsn) or ""
        amount = line.land_cost_excl_gst or 0.0
        # Booking quantity (MTS -> KGS, tape piece count, dozens/gross/thousand
        # -> pieces); keep the amount, adjust rate.
        qty, unit = _booking_qty_unit(line)
        unit_rate = (amount / qty) if qty else 0.0
        pct = _pct(rate)
        half = round(pct / 2, 2)
        alloc = {
            "stockitemname": name,
            "gstovrdnineligibleitc": NOT_APPLICABLE,
            "gstovrdnisrevchargeappl": NOT_APPLICABLE,
            "gstovrdntaxability": "Taxable",
            "gstsourcetype": "Stock Item",
            "gstitemsource": name,
            "hsnsourcetype": "Stock Item",
            "hsnitemsource": name,
            "gstovrdntypeofsupply": "Goods",
            "gstrateinferapplicability": "As per Masters/Company",
            "gsthsnname": hsn,
            "gsthsninferapplicability": "As per Masters/Company",
            "isdeemedpositive": True,
            "islastdeemedpositive": True,
            "isautonegate": False,
            "iscustomsclearance": False,
            "rate": _rate(unit_rate, unit),
            "amount": _amt(-amount),
            "actualqty": _qty(qty, unit),
            "billedqty": _qty(qty, unit),
            **_cost_centre_block(cost_centre, -amount, True),
            "batchallocations": [
                {
                    "godownname": GODOWN,
                    "batchname": BATCH,
                    "indentno": NOT_APPLICABLE,
                    "orderno": NOT_APPLICABLE,
                    "trackingnumber": NOT_APPLICABLE,
                    "dynamiccstiscleared": False,
                    "amount": _amt(-amount),
                    "actualqty": _qty(qty, unit),
                    "billedqty": _qty(qty, unit),
                }
            ],
            "ratedetails": [
                {"gstratedutyhead": "CGST", "gstratevaluationtype": "Based on Value", "gstrate": f" {half}"},
                {"gstratedutyhead": "SGST/UTGST", "gstratevaluationtype": "Based on Value", "gstrate": f" {half}"},
                {"gstratedutyhead": "IGST", "gstratevaluationtype": "Based on Value", "gstrate": f" {_pct_label(rate)}"},
                {"gstratedutyhead": "Cess", "gstratevaluationtype": NOT_APPLICABLE},
                {"gstratedutyhead": "State Cess", "gstratevaluationtype": "Based on Value"},
            ],
        }
        return alloc

    # -- tax free (zero-rated) ---------------------------------------------
    def _tax_free_entry(self, group: list[ComputedLine], cost_centre: str) -> dict:
        ledger = self._ledger(KIND_TAX_FREE)
        total = sum(l.land_cost_excl_gst or 0.0 for l in group)
        return {
            "oldauditentryids": [{"metadata": True, "type": "Number"}, "-1"],
            "ledgername": ledger,
            "gstclass": NOT_APPLICABLE,
            "gstovrdnineligibleitc": NOT_APPLICABLE,
            "gstovrdnisrevchargeappl": NOT_APPLICABLE,
            "gstovrdntaxability": "Nil Rated",
            "gstsourcetype": "Ledger",
            "gstledgersource": ledger,
            "gstovrdntypeofsupply": "Services",
            "isdeemedpositive": True,
            "ledgerfromitem": False,
            "removezeroentries": False,
            "ispartyledger": False,
            "islastdeemedpositive": True,
            "amount": _amt(-total),
            "vatexpamount": _amt(-total),
            "inventoryallocations": [
                self._inventory_nil(l, cost_centre) for l in group
            ],
        }

    def _inventory_nil(self, line: ComputedLine, cost_centre: str) -> dict:
        name = _stock_name(line)
        hsn = _text(line.source.cth_hsn) or ""
        amount = line.land_cost_excl_gst or 0.0
        qty, unit = _booking_qty_unit(line)
        unit_rate = (amount / qty) if qty else 0.0
        return {
            "stockitemname": name,
            "gstovrdntaxability": "Nil Rated",
            "gstsourcetype": "Stock Item",
            "gstitemsource": name,
            "hsnsourcetype": "Stock Item",
            "hsnitemsource": name,
            "gstovrdntypeofsupply": "Goods",
            "gsthsnname": hsn,
            "isdeemedpositive": True,
            "islastdeemedpositive": True,
            "rate": _rate(unit_rate, unit),
            "amount": _amt(-amount),
            "actualqty": _qty(qty, unit),
            "billedqty": _qty(qty, unit),
            **_cost_centre_block(cost_centre, -amount, True),
            "batchallocations": [
                {
                    "godownname": GODOWN,
                    "batchname": BATCH,
                    "amount": _amt(-amount),
                    "actualqty": _qty(qty, unit),
                    "billedqty": _qty(qty, unit),
                }
            ],
        }

    # -- IGST purchase / payable -------------------------------------------
    def _igst_purchase_entry(self, rate: float, igst: float, cost_centre: str) -> dict:
        ledger = self._ledger(KIND_IGST_PURCHASE, rate)
        return self._tax_entry(ledger, -igst, deemed_positive=True, cost_centre=cost_centre)

    def _igst_payable_entry(self, rate: float, igst: float, cost_centre: str) -> dict:
        ledger = self._ledger(KIND_IGST_PAYABLE, rate)
        return self._tax_entry(ledger, igst, deemed_positive=False, cost_centre=cost_centre)

    def _tax_entry(self, ledger: str, amount: float, deemed_positive: bool, cost_centre: str) -> dict:
        return {
            "oldauditentryids": [{"metadata": True, "type": "Number"}, "-1"],
            "ledgername": ledger,
            "gstclass": NOT_APPLICABLE,
            "isdeemedpositive": deemed_positive,
            "ledgerfromitem": False,
            "removezeroentries": False,
            "ispartyledger": False,
            "islastdeemedpositive": deemed_positive,
            "amount": _amt(amount),
            "vatexpamount": _amt(amount),
            **_cost_centre_block(cost_centre, amount, deemed_positive),
        }

    def _simple_credit(self, ledger: str, amount: float, cost_centre: str | None) -> dict:
        return {
            "oldauditentryids": [{"metadata": True, "type": "Number"}, "-1"],
            "ledgername": ledger,
            "gstclass": NOT_APPLICABLE,
            "isdeemedpositive": False,
            "ledgerfromitem": False,
            "removezeroentries": False,
            "ispartyledger": False,
            "islastdeemedpositive": False,
            "amount": _amt(amount),
            "vatexpamount": _amt(amount),
        }

    # -- voucher shell ------------------------------------------------------
    def _narration(self, cost_centre: str, usd_total: float, usd_rate: float) -> str:
        """``CO-49 CTN-1100 USD 31,453.57 @96.05`` - matches Tally's own wording.

        The USD total is thousands-grouped exactly as Tally writes it; a bare
        ``31453.57`` did not match the narration typed against the same sheet.
        """
        base = cost_centre or ""
        return f"{base} USD {usd_total:,.2f} @{_fmt_num(usd_rate)}".strip()

    def _string_block(self, lines: list[str]) -> list:
        """A Tally multi-line string block: ``[{metadata..}, line, line, ...]``."""
        return [{"metadata": True, "type": "String"}, *lines]

    def _voucher_shell(
        self,
        computed: ComputedDocument,
        party_ledger: str,
        narration: str,
        cost_centre: str,
        voucher_date: str | None = None,
    ) -> dict:
        date = voucher_date or _tally_date(_text(computed.header.be_date))
        guid = _guid()
        party_name = self._party_name(computed)
        buyer_name = self._buyer_name(computed)
        buyer_gstin = self._buyer_gstin(computed)
        buyer_state = self._buyer_state(computed)
        buyer_pincode = self._buyer_pincode(computed)
        buyer_addr = self._buyer_address_lines(computed)
        seller_addr = self._seller_address_lines(computed)
        seller_country = self._seller_country(computed)

        shell = {
            "metadata": {
                "type": "Voucher",
                "guid": guid,
                "vchtype": "Purchase",
                "action": "Create",
                "objview": "Accounting Voucher View",
            },
            "date": date,
            # ``reference`` is the supplier invoice number, so ``referencedate``
            # is that invoice's date - not the BE date. Tally's own vouchers
            # pair them this way; it falls back to the BE date when the invoice
            # date is absent.
            "referencedate": _tally_date(_text(computed.header.invoice_date)) or date,
            "effectivedate": date,
            "guid": guid,
            "vatdealertype": "Regular",
            "narration": narration,
            "enteredby": self.company.entered_by,
            # The supplier's country. A Bill of Entry is an import document, so
            # defaulting this to India was always wrong; when the country is
            # unknown it is left blank rather than asserted incorrectly.
            "countryofresidence": seller_country,
            # Party (supplier) vs company GST registration type. Tally shows the
            # former on the bill's party details; omitting it left it blank.
            "gstregistrationtype": _party_gst_registration_type(seller_country),
            "cmpgstregistrationtype": self.company.gst_registration_type,
            "vouchertypename": "Purchase",
            "partyname": party_name,
            "partyledgername": party_ledger,
            "partymailingname": party_name,
            "basicbasepartyname": party_name,
            "basicbuyername": self.company.tally_name or buyer_name,
            "placeofsupply": buyer_state,
            "cmpgststate": buyer_state,
            "consigneestatename": buyer_state,
            "consigneecountryname": "India",
            "cmpgstin": buyer_gstin,
            "consigneegstin": buyer_gstin,
            "consigneepincode": buyer_pincode,
            "consigneemailingname": buyer_name,
            "reference": _text(computed.header.invoice_no) or "",
            "costcentrename": cost_centre,
            "vchentrymode": "As Voucher",
            "iscostcentre": bool(cost_centre),
            "persistedview": "Accounting Voucher View",
            "isinvoice": False,
            "isdeemedpositive": False,
            "iseligibleforitc": True,
        }
        # Voucher numbering. Tally binds each *automatic* numbering series to one
        # GST registration, so asking it to auto-number a voucher that names no
        # tax unit raises "the voucher numbering series selected is already used
        # for another GST Registration". Tally's own vouchers for these
        # consignments are numbered manually with the BE number, so we do the
        # same and never touch the automatic series. With no number to use at
        # all, numbering is left to Tally rather than sending an empty one.
        number = self._voucher_number(computed)
        if number:
            shell["numberingstyle"] = "Manual"
            shell["vouchernumber"] = number
            shell["vouchernumberseries"] = "Default"
        else:
            shell["numberingstyle"] = "Auto"

        # The GST registration (tax unit) this voucher belongs to. Without it
        # Tally cannot bind the voucher to a registration. Emitted only when the
        # registration is actually known - never asserted on a guess.
        tax_unit = self._tax_unit(computed)
        if tax_unit:
            shell["gstregistration"] = {
                "value": tax_unit,
                "taxtype": "GST",
                "taxregistration": buyer_gstin,
            }
            shell["vchstatustaxunit"] = tax_unit
            shell["vchstatusvouchertype"] = "Purchase"
            shell["vchstatustaxadjustment"] = "Default"
            shell["vchstatusdate"] = date

        # Seller (supplier) address block and buyer address block, populated
        # from the BOE so Tally shows the correct parties.
        if seller_addr:
            shell["address"] = self._string_block(seller_addr)
        if buyer_addr:
            shell["basicbuyeraddress"] = self._string_block(buyer_addr)
        return shell


# ---------------------------------------------------------------------------
# Module-level helpers
# ---------------------------------------------------------------------------
def _fmt_num(x: float) -> str:
    return str(int(x)) if float(x).is_integer() else ("%g" % x)


def _same_rate(a: float, b: float) -> bool:
    return round(a * 10000) == round(b * 10000)


def _distinct_rates(computed: ComputedDocument) -> list[float]:
    """Distinct IGST rate fractions present, ascending."""
    rates: list[float] = []
    for line in computed.lines:
        r = _line_igst_fraction(line)
        if not any(_same_rate(r, x) for x in rates):
            rates.append(r)
    return sorted(rates)


def apply_stock_names(computed: ComputedDocument, names: dict[int, str]) -> ComputedDocument:
    """Return a copy of ``computed`` with mapped "as per Tally" stock names.

    ``names`` maps a line's serial number to the canonical Tally stock-item name
    chosen in Step 2. The name replaces that line's ``description`` (the field
    the exporter reads for ``stockitemname``); lines without a mapping are left
    unchanged. This lets the in-memory JSON path use the Step 2 selections
    without going through an Excel round-trip.
    """
    if not names:
        return computed
    new_lines = []
    for line in computed.lines:
        mapped = names.get(line.source.item_serial)
        if mapped and mapped.strip():
            new_source = replace(
                line.source,
                description=RawValue(raw_text=mapped.strip(), parsed=mapped.strip()),
            )
            new_lines.append(replace(line, source=new_source))
        else:
            new_lines.append(line)
    return replace(computed, lines=new_lines)


def tally_date(text: str | None) -> str:
    """Public alias of :func:`_tally_date` (a BOE date -> Tally ``YYYYMMDD``)."""
    return _tally_date(text)


_MONTH_NAMES = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


def _tally_date(be_date: str | None) -> str:
    """Convert a BOE date to Tally ``YYYYMMDD``.

    Handles both forms a Bill of Entry prints: numeric (``14/04/2026``,
    ``14-04-2026``) and the alphabetic-month form its *invoice* date uses
    (``28-MAR-26``). Falls back to an empty string when unparseable (Tally then
    uses the import date); never guesses a wrong date.
    """
    if not be_date:
        return ""
    import re

    m = re.search(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})", be_date)
    if m:
        d, mo, y = m.groups()
        month = int(mo)
    else:
        m = re.search(r"(\d{1,2})[/\-.\s]*([A-Za-z]{3,})[/\-.\s]*(\d{2,4})", be_date)
        if not m:
            return ""
        d, name, y = m.groups()
        month = _MONTH_NAMES.get(name[:3].lower(), 0)
        if not month:
            return ""
    if len(y) == 2:
        y = "20" + y
    return f"{int(y):04d}{month:02d}{int(d):02d}"
