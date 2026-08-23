"""Neon (serverless Postgres) store of Tally reference data.

Replaces the old Master.json flow with a small, user-curated store backed by a
**Neon Postgres** database (free tier). On an ephemeral host such as Streamlit
Community Cloud a local file would be wiped on every restart, so the durable
store lives in Neon instead. There is **no JSON/file fallback**: if the database
is unreachable or misconfigured the store raises :class:`StoreError` so the UI
can flag it clearly - the mapping / JSON steps cannot work without it.

Tables (created on demand by :meth:`TallyStore.init_schema`):

- ``stock_items(name)``            - canonical "as per Tally" stock-item names,
  offered as the Step-2 dropdown so a BOE description maps to the exact name.
- ``buyers(name, gstin, state, pincode, address_lines)`` - saved buyer records.
- ``sellers(name, country, address_lines)``              - saved seller records.

The connection string is read from the ``DATABASE_URL`` environment variable /
Streamlit secret (a Neon ``postgresql://...?sslmode=require`` URL) unless passed
explicitly. Connections are short-lived (opened per operation) which suits
Neon's serverless, autosuspending free tier.
"""

from __future__ import annotations

import json
import os
from contextlib import contextmanager
from dataclasses import dataclass, field

try:  # psycopg2 is the Postgres driver; a clear error is raised if it is absent.
    import psycopg2
    from psycopg2.extras import RealDictCursor
except Exception:  # pragma: no cover - import guard
    psycopg2 = None
    RealDictCursor = None


class StoreError(RuntimeError):
    """Raised when the Neon/Postgres store is unavailable or a query fails.

    The app surfaces this directly to the user (no silent fallback) because the
    mapping and JSON-generation steps depend on the stored data.
    """


@dataclass
class BuyerRecord:
    """A saved buyer (importer) identity."""

    name: str = ""
    gstin: str = ""
    state: str = ""
    pincode: str = ""
    address_lines: list[str] = field(default_factory=list)
    # Per-company Tally identity no document carries.
    tax_unit: str = ""      # GST registration name, e.g. "Maharashtra Registration"
    tally_name: str = ""    # Tally company name, e.g. "Acme LLP (F.Y. 2026-27)"


@dataclass
class LedgerRecord:
    """A saved Tally ledger name for one company.

    ``kind`` is one of ``boe_converter.tally_exporter.LEDGER_KINDS``. ``rate_bp``
    is the IGST rate in integer basis points (5% -> 500) so a lookup never turns
    on floating-point equality; the two rate-less ledgers (custom duty, tax free)
    use 0.
    """

    company: str = ""
    kind: str = ""
    rate_bp: int = 0
    name: str = ""


@dataclass
class SellerRecord:
    """A saved seller (supplier) identity."""

    name: str = ""
    country: str = ""
    address_lines: list[str] = field(default_factory=list)


def _dsn(explicit: str | None) -> str:
    dsn = explicit or os.environ.get("DATABASE_URL") or os.environ.get("NEON_DATABASE_URL")
    if not dsn:
        raise StoreError(
            "No database connection string configured. Set DATABASE_URL to your "
            "Neon Postgres URL (Manage app → Settings → Secrets)."
        )
    return dsn


class TallyStore:
    """Neon/Postgres-backed store for stock names and buyer/seller records."""

    def __init__(self, dsn: str | None = None) -> None:
        if psycopg2 is None:
            raise StoreError(
                "The Postgres driver (psycopg2) is not installed. Add "
                "'psycopg2-binary' to requirements.txt and redeploy."
            )
        self.dsn = _dsn(dsn)

    # -- connection ---------------------------------------------------------
    def _connect(self):
        try:
            return psycopg2.connect(self.dsn, connect_timeout=10)
        except Exception as exc:  # pragma: no cover - network dependent
            raise StoreError(
                f"Could not connect to the Neon database: {exc}. Check DATABASE_URL "
                "and that the Neon project is active."
            ) from exc

    @contextmanager
    def _cursor(self, *, dict_rows: bool = False):
        """Yield a cursor with per-transaction timeout guards, ALWAYS closing the
        connection afterwards.

        psycopg2's ``with connection`` only ends the transaction, not the socket;
        leaving it open leaks connections and can exhaust Neon's free-tier limit.
        Timeouts are applied with ``SET LOCAL`` (the Neon pooler rejects them as
        startup ``options``) so a query fails fast instead of hanging on a lock.
        """
        conn = self._connect()
        try:
            factory = RealDictCursor if dict_rows else None
            with conn.cursor(cursor_factory=factory) as cur:
                cur.execute(
                    "SET LOCAL statement_timeout=15000; "
                    "SET LOCAL lock_timeout=8000; "
                    "SET LOCAL idle_in_transaction_session_timeout=15000"
                )
                yield conn, cur
                conn.commit()
        finally:
            try:
                conn.close()
            except Exception:
                pass

    def ping(self) -> None:
        """Raise :class:`StoreError` unless a trivial query succeeds."""
        try:
            with self._cursor() as (_conn, cur):
                cur.execute("SELECT 1")
                cur.fetchone()
        except StoreError:
            raise
        except Exception as exc:  # pragma: no cover - network dependent
            raise StoreError(f"Database health check failed: {exc}") from exc

    def init_schema(self) -> None:
        """Create the tables/indexes if they do not already exist."""
        ddl = """
        CREATE TABLE IF NOT EXISTS stock_items (
            id SERIAL PRIMARY KEY,
            name TEXT NOT NULL,
            company TEXT NOT NULL DEFAULT '',
            hs_code TEXT NOT NULL DEFAULT ''
        );
        -- Migrate an existing (pre-company / pre-hs_code) table and switch to a
        -- per-company unique index so the same stock name can exist for
        -- different companies.
        ALTER TABLE stock_items ADD COLUMN IF NOT EXISTS company TEXT NOT NULL DEFAULT '';
        ALTER TABLE stock_items ADD COLUMN IF NOT EXISTS hs_code TEXT NOT NULL DEFAULT '';
        DROP INDEX IF EXISTS stock_items_lname;
        CREATE UNIQUE INDEX IF NOT EXISTS stock_items_company_lname
            ON stock_items (lower(company), lower(name));
        CREATE TABLE IF NOT EXISTS buyers (
            id SERIAL PRIMARY KEY,
            name TEXT NOT NULL,
            gstin TEXT NOT NULL DEFAULT '',
            state TEXT NOT NULL DEFAULT '',
            pincode TEXT NOT NULL DEFAULT '',
            address_lines JSONB NOT NULL DEFAULT '[]'::jsonb
        );
        CREATE UNIQUE INDEX IF NOT EXISTS buyers_lname ON buyers (lower(name));
        CREATE TABLE IF NOT EXISTS sellers (
            id SERIAL PRIMARY KEY,
            name TEXT NOT NULL,
            country TEXT NOT NULL DEFAULT '',
            address_lines JSONB NOT NULL DEFAULT '[]'::jsonb
        );
        CREATE UNIQUE INDEX IF NOT EXISTS sellers_lname ON sellers (lower(name));
        -- Per-company Tally identity that no document carries: the GST
        -- registration (tax unit) name and the Tally company name, which may
        -- differ from the mailing name by a financial-year suffix.
        ALTER TABLE buyers ADD COLUMN IF NOT EXISTS tax_unit TEXT NOT NULL DEFAULT '';
        ALTER TABLE buyers ADD COLUMN IF NOT EXISTS tally_name TEXT NOT NULL DEFAULT '';
        -- Ledger names are hand-made inside each company's Tally and cannot be
        -- derived (one company holds "Igst Purchase @18.00%" where the naming
        -- convention gives "IGST Purchase @ 18.00 %"), so they are stored per
        -- company and looked up by (kind, rate).
        CREATE TABLE IF NOT EXISTS ledgers (
            id SERIAL PRIMARY KEY,
            company TEXT NOT NULL DEFAULT '',
            kind TEXT NOT NULL,
            rate_bp INTEGER NOT NULL DEFAULT 0,
            name TEXT NOT NULL
        );
        CREATE UNIQUE INDEX IF NOT EXISTS ledgers_company_kind_rate
            ON ledgers (lower(company), kind, rate_bp);
        """
        self._exec(ddl)

    # -- low-level exec helpers --------------------------------------------
    def _exec(self, sql: str, params: tuple = ()) -> None:
        try:
            with self._cursor() as (_conn, cur):
                cur.execute(sql, params)
        except StoreError:
            raise
        except Exception as exc:
            raise StoreError(f"Database write failed: {exc}") from exc

    def _query(self, sql: str, params: tuple = ()) -> list[dict]:
        try:
            with self._cursor(dict_rows=True) as (_conn, cur):
                cur.execute(sql, params)
                return list(cur.fetchall())
        except StoreError:
            raise
        except Exception as exc:
            raise StoreError(f"Database read failed: {exc}") from exc

    def _exec_returning(self, sql: str, params: tuple = ()) -> int:
        """Execute a write and return the number of affected rows."""
        try:
            with self._cursor() as (_conn, cur):
                cur.execute(sql, params)
                return cur.rowcount
        except StoreError:
            raise
        except Exception as exc:
            raise StoreError(f"Database write failed: {exc}") from exc

    # -- stock items (scoped by company) -----------------------------------
    def list_companies(self) -> list[str]:
        """Distinct company names that have stored stock items."""
        rows = self._query(
            "SELECT DISTINCT company FROM stock_items "
            "WHERE company <> '' ORDER BY company"
        )
        return [r["company"] for r in rows]

    def list_stock_items(self, company: str | None = None) -> list[str]:
        """Stock names, optionally filtered to a single company."""
        if company is None:
            rows = self._query("SELECT name FROM stock_items ORDER BY lower(name)")
        else:
            rows = self._query(
                "SELECT name FROM stock_items WHERE lower(company) = lower(%s) "
                "ORDER BY lower(name)",
                (company.strip(),),
            )
        return [r["name"] for r in rows]

    def list_stock_items_full(self, company: str | None = None) -> list[tuple[str, str]]:
        """``(name, hs_code)`` pairs, optionally filtered to a single company.

        The HS code is reference-only (shown beside the name in the dropdown);
        only the name is ever used as the mapped value.
        """
        if company is None:
            rows = self._query(
                "SELECT name, hs_code FROM stock_items ORDER BY lower(name)"
            )
        else:
            rows = self._query(
                "SELECT name, hs_code FROM stock_items WHERE lower(company) = lower(%s) "
                "ORDER BY lower(name)",
                (company.strip(),),
            )
        return [(r["name"], r.get("hs_code") or "") for r in rows]

    def add_stock_item(self, name: str, company: str = "", hs_code: str = "") -> bool:
        name = name.strip()
        if not name:
            return False
        n = self._exec_returning(
            "INSERT INTO stock_items (name, company, hs_code) VALUES (%s, %s, %s) "
            "ON CONFLICT (lower(company), lower(name)) DO NOTHING",
            (name, company.strip(), hs_code.strip()),
        )
        return n > 0

    def add_stock_items(self, names, company: str = "") -> int:
        """Bulk-add stock names for a company in ONE connection.

        ``names`` may be an iterable of plain names (assigned to ``company``),
        ``(name, company)`` pairs, or ``(name, company, hs_code)`` triples (the
        ``company`` arg is the default when an entry omits it). Case-insensitive
        de-dup per company. Returns the number of rows actually inserted.
        """
        from psycopg2.extras import execute_values

        seen: set[tuple[str, str]] = set()
        rows: list[tuple[str, str, str]] = []
        for entry in names:
            hs = ""
            if isinstance(entry, (tuple, list)):
                name = str(entry[0]).strip() if entry and entry[0] is not None else ""
                comp = (str(entry[1]).strip() if len(entry) > 1 and entry[1] else company).strip()
                hs = str(entry[2]).strip() if len(entry) > 2 and entry[2] else ""
            elif isinstance(entry, str):
                name, comp = entry.strip(), company.strip()
            else:
                continue
            key = (comp.lower(), name.lower())
            if name and key not in seen:
                seen.add(key)
                rows.append((name, comp, hs))
        if not rows:
            return 0
        try:
            with self._cursor() as (_conn, cur):
                execute_values(
                    cur,
                    "INSERT INTO stock_items (name, company, hs_code) VALUES %s "
                    "ON CONFLICT (lower(company), lower(name)) DO NOTHING",
                    rows,
                )
                count = cur.rowcount
                return count if count and count > 0 else 0
        except StoreError:
            raise
        except Exception as exc:
            raise StoreError(f"Bulk stock insert failed: {exc}") from exc

    def delete_stock_item(self, name: str, company: str | None = None) -> bool:
        if company is None:
            n = self._exec_returning(
                "DELETE FROM stock_items WHERE lower(name) = lower(%s)", (name.strip(),)
            )
        else:
            n = self._exec_returning(
                "DELETE FROM stock_items WHERE lower(name) = lower(%s) "
                "AND lower(company) = lower(%s)",
                (name.strip(), company.strip()),
            )
        return n > 0

    def replace_company_stock(self, company: str, entries) -> int:
        """Replace a company's entire stock list with ``entries`` in ONE txn.

        ``entries`` is an iterable of ``(name, hs_code)`` pairs. Every existing
        row for ``company`` is deleted first, then the new set is inserted, so
        renames, HS-code edits and deletions all take effect. Returns the number
        of rows written.
        """
        from psycopg2.extras import execute_values

        company = company.strip()
        seen: set[str] = set()
        rows: list[tuple[str, str, str]] = []
        for entry in entries:
            if isinstance(entry, (tuple, list)):
                name = str(entry[0]).strip() if entry and entry[0] is not None else ""
                hs = str(entry[1]).strip() if len(entry) > 1 and entry[1] else ""
            else:
                name, hs = str(entry).strip(), ""
            if name and name.lower() not in seen:
                seen.add(name.lower())
                rows.append((name, company, hs))
        try:
            with self._cursor() as (_conn, cur):
                cur.execute(
                    "DELETE FROM stock_items WHERE lower(company) = lower(%s)",
                    (company,),
                )
                if rows:
                    execute_values(
                        cur,
                        "INSERT INTO stock_items (name, company, hs_code) VALUES %s "
                        "ON CONFLICT (lower(company), lower(name)) DO NOTHING",
                        rows,
                    )
                return len(rows)
        except StoreError:
            raise
        except Exception as exc:
            raise StoreError(f"Replace company stock failed: {exc}") from exc

    # -- buyers -------------------------------------------------------------
    def list_buyers(self) -> list[BuyerRecord]:
        rows = self._query(
            "SELECT name, gstin, state, pincode, address_lines, tax_unit, tally_name "
            "FROM buyers ORDER BY lower(name)"
        )
        return [_row_to_buyer(r) for r in rows]

    def add_buyer(self, buyer: BuyerRecord) -> bool:
        if not buyer.name.strip():
            return False
        self._exec(
            "INSERT INTO buyers (name, gstin, state, pincode, address_lines, "
            "tax_unit, tally_name) "
            "VALUES (%s, %s, %s, %s, %s::jsonb, %s, %s) "
            "ON CONFLICT (lower(name)) DO UPDATE SET "
            "gstin = EXCLUDED.gstin, state = EXCLUDED.state, "
            "tax_unit = EXCLUDED.tax_unit, tally_name = EXCLUDED.tally_name, "
            "pincode = EXCLUDED.pincode, address_lines = EXCLUDED.address_lines",
            (
                buyer.name.strip(),
                buyer.gstin,
                buyer.state,
                buyer.pincode,
                json.dumps(buyer.address_lines),
                buyer.tax_unit,
                buyer.tally_name,
            ),
        )
        return True

    def add_buyers(self, buyers) -> int:
        """Bulk add/replace buyer records in ONE connection. Returns count saved."""
        rows = [
            (
                b.name.strip(), b.gstin, b.state, b.pincode,
                json.dumps(b.address_lines), b.tax_unit, b.tally_name,
            )
            for b in buyers if b.name.strip()
        ]
        if not rows:
            return 0
        try:
            with self._cursor() as (_conn, cur):
                cur.executemany(
                    "INSERT INTO buyers (name, gstin, state, pincode, "
                    "address_lines, tax_unit, tally_name) "
                    "VALUES (%s, %s, %s, %s, %s::jsonb, %s, %s) "
                    "ON CONFLICT (lower(name)) DO UPDATE SET "
                    "gstin = EXCLUDED.gstin, state = EXCLUDED.state, "
                    "tax_unit = EXCLUDED.tax_unit, tally_name = EXCLUDED.tally_name, "
                    "pincode = EXCLUDED.pincode, address_lines = EXCLUDED.address_lines",
                    rows,
                )
                return len(rows)
        except StoreError:
            raise
        except Exception as exc:
            raise StoreError(f"Bulk buyer insert failed: {exc}") from exc

    def delete_buyer(self, name: str) -> bool:
        n = self._exec_returning(
            "DELETE FROM buyers WHERE lower(name) = lower(%s)", (name.strip(),)
        )
        return n > 0

    def find_buyer(self, name: str) -> BuyerRecord | None:
        rows = self._query(
            "SELECT name, gstin, state, pincode, address_lines, tax_unit, tally_name "
            "FROM buyers WHERE lower(name) = lower(%s)",
            (name.strip(),),
        )
        return _row_to_buyer(rows[0]) if rows else None

    # -- ledgers ------------------------------------------------------------
    def list_ledgers(self, company: str | None = None) -> list[LedgerRecord]:
        """Stored Tally ledger names, optionally for a single company."""
        if company is None:
            rows = self._query(
                "SELECT company, kind, rate_bp, name FROM ledgers "
                "ORDER BY lower(company), kind, rate_bp"
            )
        else:
            rows = self._query(
                "SELECT company, kind, rate_bp, name FROM ledgers "
                "WHERE lower(company) = lower(%s) ORDER BY kind, rate_bp",
                (company.strip(),),
            )
        return [_row_to_ledger(r) for r in rows]

    def add_ledger(self, ledger: LedgerRecord) -> bool:
        if not ledger.kind.strip() or not ledger.name.strip():
            return False
        self._exec(
            "INSERT INTO ledgers (company, kind, rate_bp, name) "
            "VALUES (%s, %s, %s, %s) "
            "ON CONFLICT (lower(company), kind, rate_bp) DO UPDATE SET "
            "name = EXCLUDED.name",
            (
                ledger.company.strip(),
                ledger.kind.strip(),
                int(ledger.rate_bp),
                ledger.name.strip(),
            ),
        )
        return True

    def add_ledgers(self, ledgers) -> int:
        """Bulk upsert ledger names in ONE connection. Returns rows written."""
        from psycopg2.extras import execute_values

        rows = [
            (l.company.strip(), l.kind.strip(), int(l.rate_bp), l.name.strip())
            for l in ledgers
            if l.kind.strip() and l.name.strip()
        ]
        if not rows:
            return 0
        with self._cursor() as (_conn, cur):
            execute_values(
                cur,
                "INSERT INTO ledgers (company, kind, rate_bp, name) VALUES %s "
                "ON CONFLICT (lower(company), kind, rate_bp) DO UPDATE SET "
                "name = EXCLUDED.name",
                rows,
            )
        return len(rows)

    def delete_ledger(self, company: str, kind: str, rate_bp: int) -> None:
        self._exec(
            "DELETE FROM ledgers WHERE lower(company) = lower(%s) "
            "AND kind = %s AND rate_bp = %s",
            (company.strip(), kind.strip(), int(rate_bp)),
        )

    # -- sellers ------------------------------------------------------------
    def list_sellers(self) -> list[SellerRecord]:
        rows = self._query(
            "SELECT name, country, address_lines FROM sellers ORDER BY lower(name)"
        )
        return [_row_to_seller(r) for r in rows]

    def add_seller(self, seller: SellerRecord) -> bool:
        if not seller.name.strip():
            return False
        self._exec(
            "INSERT INTO sellers (name, country, address_lines) "
            "VALUES (%s, %s, %s::jsonb) "
            "ON CONFLICT (lower(name)) DO UPDATE SET "
            "country = EXCLUDED.country, address_lines = EXCLUDED.address_lines",
            (seller.name.strip(), seller.country, json.dumps(seller.address_lines)),
        )
        return True

    def add_sellers(self, sellers) -> int:
        """Bulk add/replace seller records in ONE connection. Returns count saved."""
        rows = [
            (s.name.strip(), s.country, json.dumps(s.address_lines))
            for s in sellers if s.name.strip()
        ]
        if not rows:
            return 0
        try:
            with self._cursor() as (_conn, cur):
                cur.executemany(
                    "INSERT INTO sellers (name, country, address_lines) "
                    "VALUES (%s, %s, %s::jsonb) "
                    "ON CONFLICT (lower(name)) DO UPDATE SET "
                    "country = EXCLUDED.country, address_lines = EXCLUDED.address_lines",
                    rows,
                )
                return len(rows)
        except StoreError:
            raise
        except Exception as exc:
            raise StoreError(f"Bulk seller insert failed: {exc}") from exc

    def delete_seller(self, name: str) -> bool:
        n = self._exec_returning(
            "DELETE FROM sellers WHERE lower(name) = lower(%s)", (name.strip(),)
        )
        return n > 0

    def find_seller(self, name: str) -> SellerRecord | None:
        rows = self._query(
            "SELECT name, country, address_lines FROM sellers "
            "WHERE lower(name) = lower(%s)",
            (name.strip(),),
        )
        return _row_to_seller(rows[0]) if rows else None


def _as_lines(value) -> list[str]:
    """Coerce a JSONB address_lines column into a list[str]."""
    if isinstance(value, list):
        return [str(x) for x in value]
    if isinstance(value, str):
        try:
            parsed = json.loads(value)
            return [str(x) for x in parsed] if isinstance(parsed, list) else []
        except ValueError:
            return []
    return []


def _row_to_buyer(r: dict) -> BuyerRecord:
    return BuyerRecord(
        name=r.get("name", ""),
        gstin=r.get("gstin", "") or "",
        state=r.get("state", "") or "",
        pincode=r.get("pincode", "") or "",
        address_lines=_as_lines(r.get("address_lines")),
        tax_unit=r.get("tax_unit", "") or "",
        tally_name=r.get("tally_name", "") or "",
    )


def _row_to_ledger(r: dict) -> LedgerRecord:
    return LedgerRecord(
        company=r.get("company", "") or "",
        kind=r.get("kind", "") or "",
        rate_bp=int(r.get("rate_bp", 0) or 0),
        name=r.get("name", "") or "",
    )


def _row_to_seller(r: dict) -> SellerRecord:
    return SellerRecord(
        name=r.get("name", ""),
        country=r.get("country", "") or "",
        address_lines=_as_lines(r.get("address_lines")),
    )
