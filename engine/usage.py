"""Token usage log: one row per model call, in out/registry.db (table usage).

purpose is what the call was for: "chat" (a question in the Ask tab), "identify" (target / valuation date)
or "summary" (the written change summary). session is the chat conversation id from the browser.
"""
import sqlite3
import threading
import time
from pathlib import Path

from pricing import ESTIMATES, cost

DB = Path(__file__).resolve().parent.parent / "out" / "registry.db"
_lock = threading.Lock()


def _conn() -> sqlite3.Connection:
    DB.parent.mkdir(exist_ok=True)
    db = sqlite3.connect(DB, check_same_thread=False)
    db.execute("""CREATE TABLE IF NOT EXISTS usage(ts REAL, purpose TEXT, model TEXT, session TEXT, file_id INT,
                  input_tokens INT, cached_tokens INT, output_tokens INT, cost_usd REAL)""")
    if "cost_usd" not in [r[1] for r in db.execute("PRAGMA table_info(usage)")]:  # log created before prices
        db.execute("ALTER TABLE usage ADD COLUMN cost_usd REAL")
        rows = db.execute("SELECT rowid, model, input_tokens, cached_tokens, output_tokens FROM usage").fetchall()
        db.executemany("UPDATE usage SET cost_usd=? WHERE rowid=?", [(cost(m, i, c, o), rid) for rid, m, i, c, o in rows])
        db.commit()
    return db


def record(model: str, u, purpose: str, file_id: int | None = None, session: str | None = None) -> dict:
    """Log one response's usage (the `usage` object from the Responses API). Returns the counts."""
    row = {"input_tokens": getattr(u, "input_tokens", 0) or 0, "output_tokens": getattr(u, "output_tokens", 0) or 0,
           "cached_tokens": getattr(getattr(u, "input_tokens_details", None), "cached_tokens", 0) or 0}
    with _lock, _conn() as db:
        db.execute("INSERT INTO usage VALUES (?,?,?,?,?,?,?,?,?)",
                   (time.time(), purpose, model, session, file_id, row["input_tokens"], row["cached_tokens"],
                    row["output_tokens"], cost(model, row["input_tokens"], row["cached_tokens"], row["output_tokens"])))
    return row


_SUMS = ("COUNT(*) AS calls, COALESCE(SUM(input_tokens),0) AS input_tokens, "
         "COALESCE(SUM(cached_tokens),0) AS cached_tokens, COALESCE(SUM(output_tokens),0) AS output_tokens, "
         "COALESCE(SUM(cost_usd),0) AS cost_usd, SUM(cost_usd IS NULL) AS unpriced_calls")


def summary(session: str | None = None) -> dict:
    with _conn() as db:
        db.row_factory = sqlite3.Row
        q = lambda sql, *a: [dict(r) for r in db.execute(sql, a)]
        return {
            "estimated_models": sorted(ESTIMATES),
            "total": q(f"SELECT {_SUMS}, MIN(ts) AS since FROM usage")[0],
            "session": q(f"SELECT {_SUMS} FROM usage WHERE session = ?", session)[0] if session else None,
            "by_model": q(f"SELECT model, {_SUMS} FROM usage GROUP BY model ORDER BY SUM(input_tokens) DESC"),
            "by_purpose": q(f"SELECT purpose, {_SUMS} FROM usage GROUP BY purpose ORDER BY SUM(input_tokens) DESC"),
            "by_day": q(f"""SELECT date(ts, 'unixepoch', 'localtime') AS day, {_SUMS} FROM usage
                            GROUP BY day ORDER BY day DESC LIMIT 30"""),
            "sessions": q(f"""SELECT u.session, MIN(u.ts) AS started, MAX(u.ts) AS last, u.file_id, f.filename,
                              GROUP_CONCAT(DISTINCT u.model) AS models, {_SUMS}
                              FROM usage u LEFT JOIN files f ON f.id = u.file_id
                              WHERE u.session IS NOT NULL GROUP BY u.session ORDER BY last DESC LIMIT 25"""),
        }
