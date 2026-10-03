"""Registry of uploaded workbooks: fingerprint, process once (row map, external links), identify.

Each file is keyed by the SHA-256 of its bytes, so the same file uploaded again (under any name) is not
processed twice. Each version gets its own folder, out/<stem>__<sha8>/. Processing runs on one background
worker thread, one file at a time.
"""
import hashlib
import json
import queue
import re
import shutil
import sqlite3
import threading
import time
import traceback
from pathlib import Path

import build_map
import identify as identmod

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "out"
UPLOADS = ROOT / "uploads"
REGISTRY = OUT / "registry.db"
SUMMARY_MODEL = "gpt-6-luna"  # identifies a file when the upload didn't say which model
SUPPORTED = (".xlsx", ".xlsm")

_lock = threading.Lock()
_jobs: "queue.Queue[int]" = queue.Queue()

SCHEMA = """
CREATE TABLE IF NOT EXISTS files(
  id INTEGER PRIMARY KEY, sha256 TEXT UNIQUE, filename TEXT, size INT, uploaded_at REAL, source_path TEXT,
  status TEXT, step TEXT, pct REAL, error TEXT, out_dir TEXT, db_path TEXT, processed_at REAL,
  sheets INT, line_items INT, build_secs REAL,
  target_name TEXT, project_name TEXT, valuation_date TEXT, identity_json TEXT, identity_confirmed INT DEFAULT 0,
  previous_id INT, diff_json TEXT, diff_summary TEXT);
"""
LIST_COLS = ("id, sha256, filename, size, uploaded_at, status, step, pct, error, processed_at, started_at, sheets, line_items, "
             "build_secs, target_name, project_name, valuation_date, identity_confirmed, previous_id, "
             "json_array_length(diff_json, '$.warnings') AS n_warnings")


def _conn() -> sqlite3.Connection:
    OUT.mkdir(exist_ok=True)
    db = sqlite3.connect(REGISTRY, check_same_thread=False, timeout=30)
    db.row_factory = sqlite3.Row
    db.executescript(SCHEMA)
    have = {r[1] for r in db.execute("PRAGMA table_info(files)")}
    if "started_at" not in have:  # registries from before timings
        db.execute("ALTER TABLE files ADD COLUMN started_at REAL")
    if "model" not in have:  # ... and before the model was chosen on the page
        db.execute("ALTER TABLE files ADD COLUMN model TEXT")
    return db


def _update(fid: int, **fields) -> None:
    with _lock, _conn() as db:
        db.execute(f"UPDATE files SET {', '.join(f'{k}=?' for k in fields)} WHERE id=?", (*fields.values(), fid))


def get(fid: int, full: bool = False) -> dict | None:
    with _conn() as db:
        r = db.execute(f"SELECT {'*' if full else LIST_COLS} FROM files WHERE id=?", (fid,)).fetchone()
    if not r:
        return None
    d = dict(r)
    if full:
        for k in ("identity_json", "diff_json"):
            d[k.removesuffix("_json")] = json.loads(d.pop(k) or "null")
    return d


def all_files() -> list[dict]:
    with _conn() as db:
        return [dict(r) for r in db.execute(f"SELECT {LIST_COLS} FROM files ORDER BY uploaded_at DESC")]


def by_sha(sha: str) -> dict | None:
    with _conn() as db:
        r = db.execute("SELECT id FROM files WHERE sha256=?", (sha,)).fetchone()
    return get(r["id"]) if r else None


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def add_upload(tmp_path: Path, filename: str, sha: str, model: str | None = None) -> tuple[str, dict]:
    """Register an uploaded file. Returns ("duplicate", existing) or ("queued", new record). model: the model the
    page has selected, for identifying the file (SUMMARY_MODEL if none)."""
    existing = by_sha(sha)
    if existing:
        tmp_path.unlink(missing_ok=True)
        return "duplicate", existing
    if not filename.lower().endswith(SUPPORTED):
        tmp_path.unlink(missing_ok=True)
        raise ValueError(f"{filename}: only .xlsx and .xlsm can be read (formulas can't be read from .xlsb/.xls). "
                         "Save a copy as .xlsm from Excel and upload that.")
    dest = UPLOADS / sha[:12] / Path(filename).name
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp_path.replace(dest)
    return "queued", _register(dest, filename, sha, model=model)


def _register(path: Path, filename: str, sha: str, out_dir: Path | None = None, model: str | None = None) -> dict:
    out_dir = out_dir or OUT / f"{Path(filename).stem}__{sha[:8]}"
    with _lock, _conn() as db:
        cur = db.execute("""INSERT INTO files(sha256, filename, size, uploaded_at, source_path, status, step, pct,
                            out_dir, db_path, model) VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                         (sha, filename, path.stat().st_size, time.time(), str(path), "queued", "Waiting to start",
                          0, str(out_dir), str(out_dir / "model.db"), model))
        fid = cur.lastrowid
    _jobs.put(fid)
    return get(fid)


# ---- versions ---------------------------------------------------------------------------------------

_VERSION_BITS = re.compile(r"^copy of\b|\b(v\d+[a-z]*|vsent|final|draft|updated?|rev\d*|clean)\b|"
                           r"\b\d{4} \d{2} \d{2}\b|\b\d{6,8}\b|\(\d+\)", re.I)


def family(filename: str) -> str:
    """File name with dates, 'Copy of', v2/final/draft etc. removed: '200401 Acme Valuation model_vSent' -> 'acme valuation model'."""
    s = Path(filename).stem.replace("_", " ").replace("-", " ")
    s = _VERSION_BITS.sub(" ", s)
    return re.sub(r"\s+", " ", s).strip().lower()


def _norm(s: str | None) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (s or "").lower()).strip()


def remove(fid: int) -> None:
    """Forget a file: registry row, its model.db folder and the uploaded copy. Later versions get re-linked."""
    f = get(fid, full=True)
    if not f:
        return
    if f["status"] in ("queued", "processing"):
        raise ValueError("wait for processing to finish before removing this file")
    with _lock, _conn() as db:
        db.execute("DELETE FROM files WHERE id=?", (fid,))
    out_dir, src = Path(f["out_dir"]), Path(f["source_path"])
    if out_dir.is_relative_to(OUT) and out_dir.name.endswith(f"__{f['sha256'][:8]}"):
        shutil.rmtree(out_dir, ignore_errors=True)  # never delete folders built before the registry
    if src.is_relative_to(UPLOADS):
        shutil.rmtree(src.parent, ignore_errors=True)


# ---- processing ---------------------------------------------------------------------------------------

def _stage(fid: int, lo: float, hi: float):
    return lambda frac, msg: _update(fid, pct=round(lo + (hi - lo) * frac, 3), step=msg)


def process(fid: int) -> None:
    import calllog
    with calllog.tag(workbook=fid, step="process the workbook"):
        _process(fid)


def _process(fid: int) -> None:
    f = get(fid, full=True)
    try:
        _update(fid, status="processing", step="Opening workbook", pct=0, error=None, started_at=time.time())
        stats = build_map.main(f["source_path"], f["out_dir"], _stage(fid, 0.0, 0.8))
        _update(fid, sheets=stats["sheets"], line_items=stats["line_items"], build_secs=stats["secs"])
        _update(fid, step="Reading external links", pct=0.81)
        import extlinks  # now, while nothing else reads this model.db: later they'd be written under readers' feet
        extlinks.build(f["source_path"], f["db_path"])
        _update(fid, step="Reading its structure", pct=0.815)
        try:  # rows' kinds, blocks and their copies (structure.py), in the same window; worked out later if not
            import structure
            structure.build(f["db_path"])
        except Exception:
            traceback.print_exc()

        model = f.get("model") or SUMMARY_MODEL  # the model selected on the page that uploaded it
        _update(fid, step=f"Identifying target and valuation date ({model})", pct=0.82)
        ident = identmod.identify(f["db_path"], f["filename"], model=model, file_id=fid)
        _update(fid, target_name=ident.get("target_name"), project_name=ident.get("project_name"),
                valuation_date=ident.get("valuation_date"), identity_json=json.dumps(ident, default=str))

        _update(fid, status="done", step="Done", pct=1.0, processed_at=time.time())
    except Exception as e:
        traceback.print_exc()
        _update(fid, status="error", step="Failed", error=f"{type(e).__name__}: {e}")


def rebuild(fid: int, model: str | None = None) -> dict:
    """Process a file again from scratch, done or failed (its model.db is rebuilt for everyone using it); model:
    the model now selected on the page, else the one it was processed with."""
    f = get(fid)
    if not f:
        raise ValueError("no such file")
    if f["status"] in ("queued", "processing"):
        raise ValueError(f"{f['filename']} is already being processed")
    _update(fid, status="queued", step="Waiting to start", pct=0, error=None, **({"model": model} if model else {}))
    _jobs.put(fid)
    return get(fid)


def retry(fid: int) -> dict:
    """Process a failed file again (uploading the same bytes again would only find the failed record)."""
    f = get(fid)
    if not f:
        raise ValueError("no such file")
    if f["status"] != "error":
        raise ValueError("only a file that failed can be retried")
    _update(fid, status="queued", step="Waiting to start", pct=0, error=None)
    _jobs.put(fid)
    return get(fid)


def confirm_identity(fid: int, by: str, why: list[str] | None = None) -> dict:
    """The valuation date as it stands, confirmed (by "you" or "agents", with why)."""
    rec = get(fid, full=True)
    ident = (rec or {}).get("identity") or {}
    ident["confirmed"] = {"by": by, "why": why or [], "at": time.time()}
    _update(fid, identity_confirmed=1, identity_json=json.dumps(ident, default=str))
    return get(fid, full=True)


def note_identity(fid: int, **fields) -> None:
    """Keep something about the identity with it (identity_json), e.g. the agents' check of the date."""
    rec = get(fid, full=True)
    ident = (rec or {}).get("identity") or {}
    ident.update(fields)
    _update(fid, identity_json=json.dumps(ident, default=str))


def set_identity(fid: int, target_name: str | None, project_name: str | None, valuation_date: str | None) -> dict:
    """A person's confirmation or correction of the target and valuation date."""
    _update(fid, target_name=target_name or None, project_name=project_name or None,
            valuation_date=valuation_date or None, identity_confirmed=1)
    return get(fid, full=True)


def _worker() -> None:
    while True:
        fid = _jobs.get()
        try:
            process(fid)
        finally:
            _jobs.task_done()


def start() -> None:
    """Start the worker; re-queue files interrupted by a restart."""
    import ratelimit
    threading.Thread(target=ratelimit.load_capacities, daemon=True, name="rate-limits").start()
    with _lock, _conn() as db:
        stuck = [r["id"] for r in db.execute("SELECT id FROM files WHERE status IN ('queued','processing')")]
    for fid in stuck:
        _jobs.put(fid)
    threading.Thread(target=_worker, daemon=True, name="library-worker").start()
