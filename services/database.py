"""SQLite persistence for completed deal analyses.

Deals are stored in outputs/deals.db so they survive server restarts.
The in-memory `jobs` dict is the primary store during a session;
the DB is the secondary store for cross-restart persistence.

Schema (one table: deals)
    job_id          TEXT PRIMARY KEY
    created_at      TEXT
    property_name   TEXT
    address         TEXT
    property_type   TEXT
    purchase_price  REAL
    total_units     INT
    levered_irr     REAL
    equity_multiple REAL
    dscr_yr1        REAL
    going_in_cap    REAL
    recommendation  TEXT
    full_json       TEXT   (JSON blob of the entire job dict minus deal object)
"""

import json
import logging
import os
import sqlite3
from datetime import datetime

logger = logging.getLogger(__name__)

_DB_PATH = os.path.join(os.path.dirname(__file__), "..", "outputs", "deals.db")


def _get_conn() -> sqlite3.Connection:
    os.makedirs(os.path.dirname(_DB_PATH), exist_ok=True)
    conn = sqlite3.connect(_DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Create the deals table if it doesn't exist."""
    with _get_conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS deals (
                job_id          TEXT PRIMARY KEY,
                created_at      TEXT NOT NULL,
                property_name   TEXT,
                address         TEXT,
                property_type   TEXT,
                purchase_price  REAL,
                total_units     INTEGER,
                levered_irr     REAL,
                equity_multiple REAL,
                dscr_yr1        REAL,
                going_in_cap    REAL,
                recommendation  TEXT,
                full_json       TEXT
            )
        """)
        conn.commit()
    logger.info(f"Database initialized at {_DB_PATH}")


def _safe_json(obj):
    """Serialize a job dict to JSON, skipping un-serializable objects."""
    def default(o):
        if hasattr(o, "__dict__"):
            return o.__dict__
        return str(o)
    try:
        return json.dumps(obj, default=default)
    except Exception:
        return "{}"


def save_deal(job_id: str, job: dict):
    """Persist a completed deal analysis to SQLite."""
    try:
        pf = job.get("results", {})
        m = pf.get("metrics", {})
        deal_dict = pf.get("inputs", {}).get("deal", {})
        derived = pf.get("inputs", {}).get("derived", {})
        rec = job.get("recommendation", {}) or {}

        # Exclude the deal object (not JSON serializable) from the blob
        slim_job = {k: v for k, v in job.items()
                    if k not in ("deal",) and not callable(v)}

        with _get_conn() as conn:
            conn.execute("""
                INSERT OR REPLACE INTO deals
                    (job_id, created_at, property_name, address, property_type,
                     purchase_price, total_units, levered_irr, equity_multiple,
                     dscr_yr1, going_in_cap, recommendation, full_json)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                job_id,
                datetime.utcnow().isoformat(),
                derived.get("property_name", ""),
                deal_dict.get("address", ""),
                deal_dict.get("property_type", ""),
                deal_dict.get("purchase_price"),
                deal_dict.get("total_units"),
                m.get("levered_irr"),
                m.get("equity_multiple"),
                m.get("dscr_yr1"),
                m.get("going_in_cap_rate"),
                rec.get("recommendation", ""),
                _safe_json(slim_job),
            ))
            conn.commit()
        logger.info(f"Deal {job_id} saved to database")
    except Exception as e:
        logger.error(f"Failed to save deal {job_id}: {e}")


def load_all_deals() -> list[dict]:
    """Load all saved deals as a list of summary dicts (no full_json)."""
    try:
        with _get_conn() as conn:
            rows = conn.execute("""
                SELECT job_id, created_at, property_name, address, property_type,
                       purchase_price, total_units, levered_irr, equity_multiple,
                       dscr_yr1, going_in_cap, recommendation
                FROM deals ORDER BY created_at DESC
            """).fetchall()
        return [dict(r) for r in rows]
    except Exception as e:
        logger.error(f"Failed to load deals: {e}")
        return []


def load_deal(job_id: str) -> dict | None:
    """Load the full JSON blob for a single deal."""
    try:
        with _get_conn() as conn:
            row = conn.execute(
                "SELECT full_json FROM deals WHERE job_id = ?", (job_id,)
            ).fetchone()
        if row:
            return json.loads(row["full_json"])
        return None
    except Exception as e:
        logger.error(f"Failed to load deal {job_id}: {e}")
        return None


def delete_deal(job_id: str):
    """Delete a deal from the database."""
    try:
        with _get_conn() as conn:
            conn.execute("DELETE FROM deals WHERE job_id = ?", (job_id,))
            conn.commit()
    except Exception as e:
        logger.error(f"Failed to delete deal {job_id}: {e}")
