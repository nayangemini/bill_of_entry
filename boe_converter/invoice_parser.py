"""Invoice / Packing-List parser: per-line carton (CTN) counts.

The Bill of Entry PDF does not carry a per-line carton count (only a document
level total). The supplier's *commercial invoice* does: its line table has a
``TOTAL CTNS`` column keyed by ``SR NO`` that maps 1:1 to the BOE line serial.

This parser reads that ``TOTAL CTNS`` column from the **invoice** pages of an
"invoice cum packing list" PDF and returns a ``{serial: cartons}`` mapping that
the orchestrator attaches to the matching BOE line items (so the Excel CTN
column, ``G``, is populated per line). It is intentionally tolerant: cells that
are blank in the invoice yield no entry (that line stays blank), and the more
granular *packing list* pages (a different SL-NO breakdown that would not align
1:1 with the BOE) are skipped.

Extraction is positional: the ``TOTAL CTNS`` column sits at a stable horizontal
band, so a blank carton cell is simply the absence of a token in that band -
which a text-only parse could not distinguish from the quantity column.
"""

from __future__ import annotations

import re

from boe_converter.models import RawValue

# Units seen in the invoice's quantity column; used to recognize a data row and
# to bound the carton column on its right.
_UNIT_RE = re.compile(r"^[A-Za-z]{2,4}$")
_UNITS = {
    "PCS", "PC", "DOZ", "KGS", "KG", "GRS", "THD", "SET", "NOS", "EA",
    "UNT", "MTR", "PRS", "BOX", "MTS", "MT", "CTN", "PKT", "BAG", "ROL",
    "YDS", "LTR", "SQM", "PAIR",
}

# Horizontal half-width (points) of the carton column band around the ``CTNS``
# header token's centre. The invoice's CTNS values sit within ~15pt of the
# header centre; the quantity column is ~55pt to the right, well outside.
_CTNS_BAND = 22.0

# Header spellings accepted for the carton column (case-insensitive, trailing
# punctuation ignored): suppliers print TOTAL CTNS, TOTAL CTN, CTNS. etc.
_CTNS_TOKENS = {"CTN", "CTNS"}

# Words expected in the table's header row alongside the carton header. Used to
# prefer the real column header over a stray "CTNS" in title/total text.
_HEADER_HINTS = {"SR", "NO", "DESCRIPTION", "QTY", "PRICE", "AMOUNT", "UNIT", "UQC"}

# Max gap (points) between two in-band digit tokens for them to count as one
# value split by the extractor (e.g. "20" read as "2" + "0"). pdfplumber can
# split wide-tracked digits into separate words; joining only near-adjacent
# tokens keeps two genuinely separate numbers from merging.
_DIGIT_JOIN_GAP = 6.0


def _header_token(text: str) -> str | None:
    """The carton-header token in ``text``, else ``None``."""
    cleaned = text.strip().upper().strip(".,:;")
    return cleaned if cleaned in _CTNS_TOKENS else None

# A row's SR NO sits in the far-left column.
_SR_NO_MAX_X = 70.0


def _serial_text(text: str) -> int | None:
    """A serial number from a far-left token (tolerating ``1.``), else ``None``."""
    cleaned = text.strip().rstrip(".").strip()
    return int(cleaned) if cleaned.isdigit() else None


def _unit_token(text: str) -> str | None:
    """A known unit hiding in ``text``, else ``None``.

    Accepts the bare token (``THD``) plus punctuated/merged extractions the PDF
    reader produces (``THD.``, ``CTN:``, ``THD$0.12``): a leading 2-4 letter run
    in the unit table counts. Tokens starting with a digit (``2PCS``,
    ``12CM``) never count, so description fragments can't qualify a row.
    """
    cleaned = text.strip()
    if _UNIT_RE.match(cleaned) and cleaned.upper() in _UNITS:
        return cleaned.upper()
    m = re.match(r"^([A-Za-z]{2,4})[^A-Za-z\s]", cleaned)
    if m and m.group(1).upper() in _UNITS:
        return m.group(1).upper()
    return None


class InvoicePackingListParser:
    """Extracts per-line carton counts from an invoice / packing-list PDF."""

    ROW_TOLERANCE = 3.0

    def parse_cartons(self, doc) -> dict[int, RawValue]:
        """Return ``{serial: cartons}`` read from the invoice pages of ``doc``.

        Backward-compatible wrapper over :meth:`parse_line_details`.
        """
        return {
            serial: detail["cartons"]
            for serial, detail in self.parse_line_details(doc).items()
            if detail.get("cartons") is not None
        }

    def parse_line_details(self, doc) -> dict[int, dict]:
        """Return ``{serial: {"cartons", "description"}}`` per line.

        Reads the invoice line table (packing-list pages and the totals row are
        skipped). For each line the per-line carton count (``TOTAL CTNS`` column)
        and the ``DESCRIPTION OF GOODS`` text are captured, keyed by ``SR NO``
        (which maps 1:1 to the BOE line serial). Blank cells yield ``None`` for
        that field. ``doc`` may be a path, bytes buffer, or an opened
        ``pdfplumber.PDF``.
        """
        handle, pages, should_close = self._resolve(doc)
        try:
            details: dict[int, dict] = {}
            for page in pages:
                text = page.extract_text() or ""
                if not self._is_invoice_page(text):
                    continue
                self._parse_page(page, details)
            return details
        finally:
            if should_close and handle is not None:
                try:
                    handle.close()
                except Exception:
                    pass

    # ------------------------------------------------------------------
    @staticmethod
    def _resolve(doc):
        if hasattr(doc, "pages"):
            return doc, list(doc.pages), False
        try:
            import pdfplumber
        except ImportError as exc:
            raise ValueError(
                "pdfplumber is required to read an invoice PDF; "
                "install it (pip install pdfplumber) or pass an opened "
                "document with a .pages attribute."
            ) from exc
        handle = pdfplumber.open(doc)
        return handle, list(handle.pages), True

    @staticmethod
    def _is_invoice_page(text: str) -> bool:
        """True for an invoice line-table page (not a packing-list page).

        The invoice header carries ``UNIT PRICE`` / ``TOTAL AMOUNT``; the packing
        list carries ``QTY PER CTNS`` / ``PACKINGLIST`` and a different SL-NO
        breakdown that must not be read as per-line cartons.
        """
        low = text.lower()
        if "packinglist" in low or "qty per" in low:
            return False
        return "amount" in low and "price" in low

    def _rows(self, page) -> list[list[dict]]:
        """Group the page's words into geometric rows (top-to-bottom)."""
        words = page.extract_words()
        words.sort(key=lambda w: (round(float(w["top"]), 1), float(w["x0"])))
        rows: list[list[dict]] = []
        current: list[dict] = []
        cur_top: float | None = None
        for w in words:
            top = float(w["top"])
            if cur_top is None or abs(top - cur_top) <= self.ROW_TOLERANCE:
                current.append(w)
                cur_top = top if cur_top is None else (cur_top + top) / 2.0
            else:
                rows.append(current)
                current = [w]
                cur_top = top
        if current:
            rows.append(current)
        return rows

    @staticmethod
    def _center(w: dict) -> float:
        return (float(w["x0"]) + float(w["x1"])) / 2.0

    def _ctns_center(self, rows: list[list[dict]]) -> float | None:
        """Locate the carton column header's horizontal centre on the page.

        Accepts ``CTN``/``CTNS`` in any case with trailing punctuation (suppliers
        vary the label). When several candidates exist, the one sharing its row
        with table-header words (SR/NO/DESCRIPTION/QTY/...) wins, so a stray
        "TOTAL CTNS: 500" in title text cannot hijack the band; otherwise the
        first candidate is used.
        """
        fallback: float | None = None
        for row in rows:
            hits = [w for w in row if _header_token(w["text"])]
            if not hits:
                continue
            if fallback is None:
                fallback = self._center(hits[0])
            row_words = {
                w["text"].strip().upper().strip(".,:;") for w in row
            }
            if row_words & _HEADER_HINTS:
                return self._center(hits[0])
        return fallback

    def _parse_page(self, page, details: dict[int, dict]) -> None:
        rows = self._rows(page)
        ctns_center = self._ctns_center(rows)
        if ctns_center is None:
            return
        lo, hi = ctns_center - _CTNS_BAND, ctns_center + _CTNS_BAND
        # The description column sits between the SR NO column and the CTNS band.
        desc_hi = ctns_center - _CTNS_BAND

        for row in rows:
            ordered = sorted(row, key=lambda w: float(w["x0"]))
            # SR NO: an integer in the far-left column (tolerating a trailing
            # dot, e.g. "1." as some invoices print it).
            sr_word = next(
                (w for w in ordered
                 if self._center(w) < _SR_NO_MAX_X and _serial_text(w["text"]) is not None),
                None,
            )
            if sr_word is None:
                continue
            # Require a unit token (e.g. PCS/DOZ) so totals/footer rows are
            # ignored - they have a serial-like number but no unit. The match
            # is punctuation-tolerant (``THD.``, ``CTN:``, ``THD$0.12``) since
            # tight supplier layouts merge the unit with its neighbour.
            has_unit = any(
                _unit_token(w["text"]) is not None for w in ordered
            )
            if not has_unit:
                continue

            serial = _serial_text(sr_word["text"])
            assert serial is not None
            entry = details.setdefault(serial, {"cartons": None, "description": None})

            # Carton count: numeric token(s) whose centre falls in the CTNS
            # band. Adjacent digit tokens are joined first: extractors can
            # split "20" into "2" + "0", and taking the first token turned
            # 20 into 2. Only near-adjacent tokens merge, so genuinely
            # separate numbers never combine.
            ctn_words = [
                w for w in ordered
                if lo <= self._center(w) <= hi
                and self._is_number(w["text"])
            ]
            if ctn_words:
                entry["cartons"] = self._carton_value(
                    self._join_number_tokens(ctn_words)
                )

            # Description: all tokens between the SR NO column and the CTNS band,
            # left-to-right (e.g. "SLIDERS (GARMENT ACCESSORY)").
            desc_words = [
                w for w in ordered
                if _SR_NO_MAX_X <= self._center(w) < desc_hi
            ]
            desc = " ".join(w["text"].strip() for w in desc_words).strip()
            if desc:
                entry["description"] = RawValue(raw_text=desc, parsed=desc)

    @staticmethod
    def _is_number(text: str) -> bool:
        return bool(re.fullmatch(r"\d+(?:\.\d+)?", text.strip().replace(",", "")))

    @staticmethod
    def _join_number_tokens(words: list[dict]) -> str:
        """Join in-band number tokens split by the extractor into one value.

        Tokens are ordered left-to-right and merged while the gap between one
        token's end and the next token's start stays within
        ``_DIGIT_JOIN_GAP``. The first group wins: a row carries a single
        carton value, so extra far-apart groups are stray marks, not data.
        """
        ranked = sorted(words, key=lambda w: (float(w["x0"]) + float(w["x1"])) / 2.0)
        groups: list[list[dict]] = [[ranked[0]]]
        for w in ranked[1:]:
            prev = groups[-1][-1]
            if float(w["x0"]) - float(prev["x1"]) <= _DIGIT_JOIN_GAP:
                groups[-1].append(w)
            else:
                groups.append([w])
        return "".join(w["text"].strip() for w in groups[0])

    @staticmethod
    def _carton_value(text: str) -> RawValue:
        """Wrap a carton count verbatim, parsing it as an int when whole."""
        cleaned = text.replace(",", "")
        try:
            number = float(cleaned)
        except ValueError:
            return RawValue(raw_text=text, parsed=text)
        parsed: float | int = int(number) if number.is_integer() else number
        return RawValue(raw_text=text, parsed=parsed)
