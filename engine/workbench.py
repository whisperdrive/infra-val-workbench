"""The workbench's store and steps: an engagement's four files, from upload to last year's value rebuilt in Python
and rolled forward. The orchestrator (orchestrator.py) decides what runs when; the steps here do the work.

  files      workbooks go through the library (library.py: row map, external links, target and valuation date);
             reports are read from their text layer (docingest.py, no model calls)
  facts      keyfacts.py: the few key facts, extracted by one model, checked in code, reviewed by another, the loop
             between them for what's open; the agents' agreement approves a fact, a person can change it
  roles      roles.py suggests prior report / prior client model / prior overlay / current client model from the
             workbooks' likeness, dates, links and where the report's figures sit, with a second opinion; the
             orchestrator confirms them with that evidence, or asks
  rebuild    overlay.py: the overlay sheets compiled to a Python module, checked cell by cell against Excel, fed
             from last year's client model, and rolled forward onto this year's (rows found by rowfind.py and the
             row agents, rowagent.py)
  map        linkmap.py: report -> overlay cells, overlay -> prior client rows -> current rows

State is in out/workbench.db; reports are copied to uploads/<sha12>/ and read into out/docs/<stem>__<sha8>/
(all git-ignored). Model calls use the engagement's models and go to the usage log and the call log.
"""
import json
import re
import shutil
import sqlite3
import threading
import time
import traceback
from collections import Counter
from contextlib import closing
from pathlib import Path

import calllog
import docingest
import extlinks
import keyfacts
import lessons
import library
import linkmap
import rodb
import roles as rolesmod
import usage

ROOT = Path(__file__).resolve().parent.parent
OUT, UPLOADS = ROOT / "out", ROOT / "uploads"
DB, DOCS = OUT / "workbench.db", OUT / "docs"
DEFAULT_MODEL = "gpt-6-luna"      # extraction
DEFAULT_REVIEWER = "gpt-6-sol"    # review, second opinions, the orchestrator's decisions
DEFAULT_ARBITER = "gpt-6-sol"     # settles what the fact review loop can't
REPORT_TYPES = (".pdf", ".pptx")

_lock = threading.Lock()

SCHEMA = """
CREATE TABLE IF NOT EXISTS engagements(id INTEGER PRIMARY KEY, name TEXT, created_at REAL, updated_at REAL,
  model TEXT, reviewer_model TEXT, arbiter_model TEXT, roles_suggested TEXT, profile_json TEXT,
  map_status TEXT, map_step TEXT, map_error TEXT, map_json TEXT, map_started_at REAL, map_secs REAL,
  overlay_status TEXT, overlay_step TEXT, overlay_error TEXT, overlay_json TEXT, overlay_pct REAL,
  overlay_started_at REAL, overlay_secs REAL,
  rows_status TEXT, rows_step TEXT, rows_error TEXT, rows_started_at REAL, rows_secs REAL,
  doctor_status TEXT, doctor_step TEXT, doctor_error TEXT, doctor_started_at REAL, doctor_secs REAL,
  result_json TEXT);
CREATE TABLE IF NOT EXISTS eng_files(engagement_id INT, file_id INT, added_at REAL, PRIMARY KEY(engagement_id, file_id));
CREATE TABLE IF NOT EXISTS documents(id INTEGER PRIMARY KEY, engagement_id INT, sha256 TEXT, filename TEXT, kind TEXT,
  size INT, uploaded_at REAL, source_path TEXT, out_dir TEXT, status TEXT, step TEXT, pct REAL, error TEXT,
  pages INT, n_tables INT, n_flagged INT, doc_json TEXT, processed_at REAL, started_at REAL, doc_secs REAL,
  facts_status TEXT, facts_step TEXT, facts_error TEXT, facts_notes TEXT, review_summary TEXT, loop_json TEXT,
  facts_started_at REAL, facts_secs REAL, UNIQUE(engagement_id, sha256));
CREATE TABLE IF NOT EXISTS facts(id INTEGER PRIMARY KEY, engagement_id INT, document_id INT, n INT, category TEXT,
  key TEXT, label TEXT, value_text TEXT, low_text TEXT, high_text TEXT, value REAL, unit TEXT, basis TEXT, page INT,
  quote TEXT, origin TEXT, check_json TEXT, review_json TEXT, status TEXT DEFAULT 'pending', final_json TEXT,
  updated_at REAL, agent_json TEXT, decided_by TEXT, visual_json TEXT);
CREATE TABLE IF NOT EXISTS roles(engagement_id INT, role TEXT, kind TEXT, ref_id INT, sheets_json TEXT, why_json TEXT,
  confirmed INT DEFAULT 0, confirmed_by TEXT, evidence_json TEXT, PRIMARY KEY(engagement_id, role));
CREATE TABLE IF NOT EXISTS stages(engagement_id INT, stage TEXT, status TEXT, inputs TEXT, started_at REAL,
  finished_at REAL, note TEXT, data_json TEXT, PRIMARY KEY(engagement_id, stage));
CREATE TABLE IF NOT EXISTS runlog(id INTEGER PRIMARY KEY, engagement_id INT, at REAL, stage TEXT, event TEXT,
  issue TEXT, inputs TEXT, text TEXT, data_json TEXT);
CREATE INDEX IF NOT EXISTS runlog_eid ON runlog(engagement_id, stage, event);
"""
FACT_FIELDS = ("category", "key", "label", "value_text", "low_text", "high_text", "value", "unit", "basis", "page", "quote")


def _declared() -> dict[str, list[tuple[str, str]]]:
    """Each table's columns as SCHEMA declares them: {table: [(name, type)]}."""
    out = {}
    for table, body in re.findall(r"CREATE TABLE IF NOT EXISTS (\w+)\((.*?)\);", SCHEMA, re.S):
        parts, depth, cur = [], 0, ""
        for ch in body:
            depth += (ch == "(") - (ch == ")")
            if ch == "," and depth == 0:
                parts.append(cur)
                cur = ""
            else:
                cur += ch
        parts.append(cur)
        cols = [p.split(None, 1) for p in (x.strip() for x in parts) if p and not re.match(r"(PRIMARY|UNIQUE)\b", p)]
        out[table] = [(c[0], c[1] if len(c) > 1 else "") for c in cols]
    return out


_MIGRATED = set()


def _conn() -> sqlite3.Connection:
    OUT.mkdir(exist_ok=True)
    db = sqlite3.connect(DB, check_same_thread=False, timeout=30)
    db.row_factory = sqlite3.Row
    db.executescript(SCHEMA)
    if str(DB) not in _MIGRATED:  # a database from an earlier version gets the columns added since
        for table, cols in _declared().items():
            have = {r[1] for r in db.execute(f"PRAGMA table_info({table})")}
            for name, kind in cols:
                if name not in have:
                    db.execute(f"ALTER TABLE {table} ADD COLUMN {name} {kind}")
        _MIGRATED.add(str(DB))
    return db


def _q(sql: str, *args) -> list[dict]:
    with _conn() as db:
        return [dict(r) for r in db.execute(sql, args)]


def _exec(sql: str, *args) -> int:
    with _lock, _conn() as db:
        return db.execute(sql, args).lastrowid


def _set(table: str, rid: int, **fields) -> None:
    with _lock, _conn() as db:
        db.execute(f"UPDATE {table} SET {', '.join(f'{k}=?' for k in fields)} WHERE id=?", (*fields.values(), rid))


def _session(eid: int) -> str:
    return f"engagement-{eid}"


def _logger(eid: int):
    return lambda model, u, purpose: usage.record(model, u, purpose, None, _session(eid))


def friendly(e: Exception) -> str:
    name = type(e).__name__
    if name in ("AuthenticationRequiredError", "ClientAuthenticationError"):
        return ("Azure sign-in needed: run `uv run python engine/llm.py` in a terminal, sign in with the device code, "
                "then retry.")
    if name == "APITimeoutError":
        import llm
        return (f"the model didn't answer within {llm.READ_TIMEOUT:.0f} s, {llm.MAX_RETRIES + 1} tries "
                "(LLM_READ_TIMEOUT in .env): retry")
    return f"{name}: {e}"


# ---- engagements --------------------------------------------------------------------------------------------

def create(name: str) -> dict:
    now = time.time()
    eid = _exec("INSERT INTO engagements(id, name, created_at, updated_at, model, reviewer_model) VALUES (?,?,?,?,?,?)",
                _new_id(), name.strip() or "Untitled engagement", now, now, DEFAULT_MODEL, DEFAULT_REVIEWER)
    return get(eid)


def _new_id() -> int:
    """An engagement id nothing is left under: past every engagement, any row left under an id (a stage the
    orchestrator wrote as one was deleted) and every overlay folder, so a new engagement can't inherit a deleted
    one's leftovers. An id a delete cleared completely can be used again: nothing of it remains."""
    used = [r["n"] or 0 for r in _q("SELECT MAX(id) AS n FROM engagements")]
    used += [r["n"] or 0 for t in ("stages", "runlog", "facts", "roles", "eng_files", "documents")
             for r in _q(f"SELECT MAX(engagement_id) AS n FROM {t}")]
    used += [int(p.name[1:]) for p in (OUT / "overlays").glob("e*") if p.name[1:].isdigit()]
    return max(used, default=0) + 1


def all_engagements() -> list[dict]:
    rows = _q("""SELECT e.id, e.name, e.created_at, e.updated_at,
                 (SELECT COUNT(*) FROM documents d WHERE d.engagement_id = e.id) AS n_docs,
                 (SELECT COUNT(*) FROM eng_files f WHERE f.engagement_id = e.id) AS n_workbooks,
                 (SELECT COUNT(*) FROM roles r WHERE r.engagement_id = e.id AND r.confirmed) AS n_roles
                 FROM engagements e ORDER BY e.updated_at DESC""")
    return rows


def update(eid: int, **fields) -> dict:
    fields = {k: v for k, v in fields.items() if k in ("name", "model", "reviewer_model", "arbiter_model") and v}
    if fields:
        _set("engagements", eid, **fields, updated_at=time.time())
    return get(eid)


def workbooks(eid: int) -> list[dict]:
    ids = [r["file_id"] for r in _q("SELECT file_id FROM eng_files WHERE engagement_id=? ORDER BY added_at", eid)]
    out = []
    for fid in ids:
        rec = library.get(fid, full=True)
        if rec:
            w = {k: rec.get(k) for k in ("id", "filename", "size", "uploaded_at", "status", "step", "pct", "error",
                                         "sheets", "line_items", "target_name", "project_name", "valuation_date",
                                         "identity_confirmed",
                                         "db_path", "source_path", "identity", "processed_at", "started_at",
                                         "build_secs")}
            w["sheet_names"] = _sheet_names(w) if w["status"] == "done" and w["db_path"] else []
            out.append(w)
    return out


_SHEET_NAMES: dict[tuple, list[str]] = {}


def _sheet_names(w: dict) -> list[str]:
    """A processed workbook's sheets, read once: the page polls every couple of seconds, and a model.db can be
    busy for a moment (its link tables being built); a busy file shows no sheet list rather than failing the page."""
    key = (w["db_path"], w.get("processed_at"))
    if key not in _SHEET_NAMES:
        try:
            with rodb.connect(w["db_path"], timeout=2) as m:
                _SHEET_NAMES[key] = [r[0] for r in m.execute("SELECT sheet FROM sheets ORDER BY rowid")]
        except sqlite3.Error:
            return []  # not cached: tried again on the next poll
    return _SHEET_NAMES[key]


def documents(eid: int) -> list[dict]:
    cols = ("id, engagement_id, filename, kind, size, uploaded_at, status, step, pct, error, pages, n_tables, "
            "processed_at, facts_status, facts_step, facts_error, facts_notes, review_summary, loop_json, started_at, "
            "doc_secs, facts_started_at, facts_secs")
    out = _q(f"SELECT {cols} FROM documents WHERE engagement_id=? ORDER BY uploaded_at", eid)
    for d in out:
        d["loop"] = json.loads(d.pop("loop_json") or "null")
    return out


def facts(eid: int) -> list[dict]:
    out = []
    for f in _q("SELECT * FROM facts WHERE engagement_id=? ORDER BY document_id, n", eid):
        for k in ("check_json", "review_json", "final_json", "agent_json", "visual_json"):
            f[k.removesuffix("_json")] = json.loads(f.pop(k) or "null")
        out.append(f)
    return out


def reference(eid: int) -> list[dict]:
    """The facts to navigate by: approved ones as approved (with edits); if none approved yet, those that pass
    the code checks and aren't waiting for a person, marked unapproved."""
    fs = facts(eid)
    approved = [{**{k: f[k] for k in ("id", *FACT_FIELDS)}, **(f["final"] or {}), "approved": True}
                for f in fs if f["status"] == "approved"]
    if approved:
        return approved
    return [{**{k: f[k] for k in ("id", *FACT_FIELDS)}, "approved": False} for f in fs
            if f["status"] != "rejected" and (f["check"] or {}).get("ok")
            and (f["agent"] or {}).get("status") not in ("withdrawn", "escalated")]  # one a person must settle: not yet


def roles(eid: int) -> dict:
    out = {}
    for r in _q("SELECT * FROM roles WHERE engagement_id=?", eid):
        out[r["role"]] = {"kind": r["kind"], "id": r["ref_id"], "sheets": json.loads(r["sheets_json"] or "null"),
                          "why": json.loads(r["why_json"] or "[]"), "confirmed": bool(r["confirmed"]),
                          "by": r["confirmed_by"], "evidence": json.loads(r["evidence_json"] or "null")}
    return out


def get(eid: int) -> dict | None:
    rows = _q("SELECT * FROM engagements WHERE id=?", eid)
    if not rows:
        return None
    e = rows[0]
    for k in ("roles_suggested", "map_json", "overlay_json", "profile_json", "result_json"):
        e[k.removesuffix("_json")] = json.loads(e.pop(k) or "null")
    wbs = workbooks(eid)
    for w in wbs:
        w.pop("db_path", None)
        w.pop("source_path", None)
        ident = w.pop("identity", None) or {}
        w["identity_notes"] = ident.get("notes")
        conf = ident.get("confirmed") or {}
        w["identity_by"] = conf.get("by") or ("you" if w.get("identity_confirmed") else None)
        w["identity_why"] = conf.get("why") or []
        w["identity_check"] = ident.get("auto_check")  # the agents' check of the date, where it didn't confirm
    import orchestrator  # what runs, and what it's waiting for: the orchestrator's view (reading it starts nothing)
    rl = roles(eid)
    terminal = terminal_view(eid, rl)
    return {**e, "documents": documents(eid), "workbooks": wbs, "facts": facts(eid), "roles": rl,
            "session": _session(eid), "now": time.time(), "run": orchestrator.view(eid),
            "dates": dates(eid, e.get("result"), wbs, rl), "terminal": terminal}


def terminal_view(eid: int, rl: dict | None = None) -> dict | None:
    """How last year's report works out its terminal value (keyfacts.terminal_method), with what it says about it
    (context.terminal: searched once per report, kept with its loop notes)."""
    import context
    rl = rl if rl is not None else roles(eid)
    did = (rl.get("prior_report") or {}).get("id") if (rl.get("prior_report") or {}).get("kind") == "document" else None
    d = next((x for x in documents(eid) if x["id"] == did), None) if did else \
        next((x for x in documents(eid) if x.get("facts_status") == "done"), None)
    if not d or d.get("facts_status") != "done":
        return None
    passages = (d.get("loop") or {}).get("terminal")
    if passages is None:
        doc = (document(d["id"]) or {}).get("doc") or {}
        passages = context.terminal(doc.get("markdown") or "")
        _note_loop(d["id"], terminal=passages)
    return keyfacts.terminal_method(reference(eid), passages)


def dates(eid: int, res: dict | None, wbs: list[dict], rl: dict) -> list[dict]:
    """The valuation dates across the files, in one line, the report's the anchor: the report (last year's) → last
    year's overlay (the cell its discount factors read, once they're traced; else the date found by its label) →
    last year's client model (a day or so off is a note: a model can carry its start or balance date) → this year's
    client model (after last year's, a year on, confirmed by whom) → the date the roll-forward runs to. Each step:
    {"step", "date", "where", "how", "ok" (True / False / None: a note), "note"}. Shows what the checks found; the
    rules that confirm a date are the orchestrator's."""
    rep = next((f for f in reference(eid) if f["key"] == "valuation_date" and f.get("value")), None)
    anchor = f"{int(rep['value']) // 10000:04d}-{int(rep['value']) // 100 % 100:02d}-{int(rep['value']) % 100:02d}" \
        if rep else None
    seen = next((f.get("visual") or {} for f in facts(eid) if rep and f["id"] == rep["id"]), {})
    out = [{"step": "The report", "date": anchor, "where": f"page {rep['page']}" if rep and rep.get("page") else None,
            "how": rep.get("value_text") if rep else None, "ok": True if rep else None,
            "note": ("seen on its image" if str(seen.get("status", "")).startswith("confirmed") else
                     "corrected on its image" if seen.get("status") == "corrected" else "read from its text")
            if rep else "last year's valuation date isn't among the key facts yet"}]
    wb_of = lambda role: next((w for w in wbs if (rl.get(role) or {}).get("kind") == "workbook"
                               and w["id"] == rl[role]["id"]), None)
    traced = sorted({(a.get("valuation_date"), a.get("valuation_date_source"), bool(a.get("valuation_date_sourced")))
                     for e in ("low", "high") for a in ((res or {}).get("assumptions") or {}).get(e) or []
                     if a.get("valuation_date")})
    ovw = wb_of("prior_overlay")
    if traced:
        d, cell, src = traced[0]
        out.append({"step": "Last year's overlay", "date": d, "where": cell,
                    "how": "the cell its discount factors read" if src else "found by its label: the factors don't read it",
                    "ok": d == anchor if anchor else None,
                    "note": None if len(traced) == 1 else "its discountings read different dates: " +
                                                          ", ".join(f"{x[0]} ({x[1]})" for x in traced)})
    else:
        d = (ovw or {}).get("valuation_date")
        out.append({"step": "Last year's overlay", "date": d, "where": None,
                    "how": "found by its label; its discountings aren't traced yet", "ok": (d == anchor) if d and anchor else None})
    pm = wb_of("prior_model")
    d = (pm or {}).get("valuation_date")
    out.append({"step": "Last year's client model", "date": d, "where": None, "how": "its own date",
                "ok": True if d and d == anchor else None,
                "note": None if not d or not anchor or d == anchor else
                "not the report's: a client model can carry its start or balance date, so this is a note"})
    cm = wb_of("current_model")
    d = (cm or {}).get("valuation_date")
    chk = (cm or {}).get("identity_check") or {}
    later = bool(d and anchor and d > anchor)
    year_on = bool(d and anchor and d == _year_on(anchor))
    who = (cm or {}).get("identity_by")
    out.append({"step": "This year's client model", "date": d, "where": None,
                "how": ("a year after the report's" if year_on else "after the report's, but not a year on" if later
                        else "not after the report's" if d and anchor else "its own date"),
                "ok": year_on if d and anchor and (year_on or not later) else None,
                "note": "; ".join(filter(None, [
                    f"confirmed by {'the agents' if who == 'agents' else who}" if who else "not confirmed yet",
                    "agrees: " + "; ".join(chk["agree"]) if chk.get("agree") else None,
                    "disagrees: " + "; ".join(chk["disagree"]) if chk.get("disagree") else None]))})
    mine = this_year_date(eid)
    used = ((res or {}).get("bridges") or {}).get("valuation_date") or mine
    out.append({"step": "The roll-forward runs to", "date": used, "where": None,
                "how": "the date you set" if mine and used == mine else "this year's client model's" if used else None,
                "ok": None, "note": None if used else "not rolled forward yet"})
    return out


def delete(eid: int) -> dict | None:
    """Delete an engagement and everything worked out for it, so the same files can be uploaded and run again from
    the start: its reports (their pages, table images and key facts), the models only it uses (their row maps,
    links, identity and the dates confirmed on them; a model another engagement also uses stays), the overlay's
    module and what was picked or held on it (out/overlays/e<eid>), its stages, run log and model-call log. The
    rules learned from reports stay: they're the app's, not the engagement's. Refused while anything of it is queued
    or running. -> {"name", "reports", "models", "kept": [models another engagement uses]}, None if there's none."""
    import orchestrator
    rows = _q("SELECT name FROM engagements WHERE id=?", eid)
    if not rows:
        return None
    docs = documents(eid)
    linked = _q("SELECT file_id FROM eng_files WHERE engagement_id=?", eid)
    wbs = [w for w in (library.get(r["file_id"], full=True) for r in linked) if w]
    busy = orchestrator.busy_with(eid) + \
        [d["filename"] for d in docs if d["status"] in ("queued", "processing") or d["facts_status"] in ("queued", "running")] + \
        [w["filename"] for w in wbs if w["status"] in ("queued", "processing")]
    if busy:
        raise ValueError(f"wait for what's running to finish before deleting it ({', '.join(busy[:4])})")
    kept = [w for w in wbs if _q("SELECT 1 FROM eng_files WHERE file_id=? AND engagement_id<>?", w["id"], eid)]
    gone = [w for w in wbs if w not in kept]
    _exec("DELETE FROM engagements WHERE id=?", eid)  # first: the orchestrator stops looking at it
    _let_go({eid}, {w["db_path"] for w in gone})
    for d in docs:
        remove_document(d["id"])
    for w in gone:
        library.remove(w["id"])
        _DATE_CHECKED.difference_update({k for k in _DATE_CHECKED if k[0] == w["id"]})
    with _lock, _conn() as db:
        for t, col in (("facts", "engagement_id"), ("roles", "engagement_id"), ("eng_files", "engagement_id"),
                       ("stages", "engagement_id"), ("runlog", "engagement_id"), ("engagements", "id")):
            db.execute(f"DELETE FROM {t} WHERE {col}=?", (eid,))
    shutil.rmtree(OUT / "overlays" / f"e{eid}", ignore_errors=True)
    calllog.forget(eid, [w["id"] for w in gone])
    usage.forget(_session(eid), [w["id"] for w in gone])
    return {"name": rows[0]["name"], "reports": len(docs), "models": len(gone), "kept": [w["filename"] for w in kept]}


# ---- uploads ------------------------------------------------------------------------------------------------

def _model(eid: int) -> str:
    """The engagement's model, as the header's Models selector has it."""
    rows = _q("SELECT model FROM engagements WHERE id=?", eid)
    return (rows[0]["model"] if rows else None) or DEFAULT_MODEL


def role_fits(role: str, filename: str) -> None:
    """A file placed in a role by hand must be the kind that role takes: last year's report a PDF or PPTX, the others
    a workbook. Raises ValueError, before anything is stored."""
    if role not in rolesmod.ROLES:
        raise ValueError(f"unknown role {role}")
    want = REPORT_TYPES if role == "prior_report" else library.SUPPORTED
    if Path(filename).suffix.lower() not in want:
        what = "last year's report" if role == "prior_report" else "a model"
        raise ValueError(f"{what} goes in as {' or '.join(x.lstrip('.').upper() for x in want)}")


def add_upload(eid: int, tmp: Path, filename: str, sha: str, role: str | None = None) -> dict:
    """Workbooks -> the shared library (deduplicated across the app); reports -> this engagement's documents. With a
    role, the file is placed in it by you (a re-upload of the same file too): the orchestrator fills in the others."""
    if role:
        try:
            role_fits(role, filename)
        except ValueError:
            tmp.unlink(missing_ok=True)
            raise
        got = add_upload(eid, tmp, filename, sha)
        confirm_roles(eid, {role: {"kind": got["kind"], "id": got["id"], "sheets": None}}, "you")
        return {**got, "role": role}
    ext = Path(filename).suffix.lower()
    if ext in library.SUPPORTED:  # identified and summarised with the engagement's model (the Models selector)
        status, rec = library.add_upload(tmp, filename, sha, _model(eid))
        with _lock, _conn() as db:
            db.execute("INSERT OR IGNORE INTO eng_files VALUES (?,?,?)", (eid, rec["id"], time.time()))
        _touch(eid)
        return {"status": status, "kind": "workbook", "id": rec["id"], "filename": rec["filename"]}
    if ext not in REPORT_TYPES:
        tmp.unlink(missing_ok=True)
        raise ValueError(f"{filename}: upload reports as .pdf or .pptx and models as .xlsx or .xlsm "
                         "(save .xlsb / .xls / .ppt / .docx in one of those formats first)")
    have = _q("SELECT id FROM documents WHERE engagement_id=? AND sha256=?", eid, sha)
    if have:
        tmp.unlink(missing_ok=True)
        return {"status": "duplicate", "kind": "document", "id": have[0]["id"], "filename": filename}
    dest = UPLOADS / sha[:12] / Path(filename).name
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp.replace(dest)
    out_dir = DOCS / f"{Path(filename).stem}__{sha[:8]}"
    did = _exec("""INSERT INTO documents(engagement_id, sha256, filename, kind, size, uploaded_at, source_path, out_dir,
                   status, step, pct) VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                eid, sha, filename, ext.lstrip("."), dest.stat().st_size, time.time(), str(dest), str(out_dir),
                "queued", "Waiting to start", 0)
    _touch(eid)  # queued: the orchestrator reads it
    return {"status": "queued", "kind": "document", "id": did, "filename": filename}


# ---- this year's valuation date, checked by the agents ---------------------------------------------------------
# identify.py reads a workbook's valuation date off one cell. Before a person is asked to check it, the agents weigh
# the evidence the files already hold (no model calls): other cells labelled like it, the file name, a year on from
# last year's valuation date, the model's own financial-year end. Confirmed when two or more agree and none
# disagrees; otherwise the person is asked, and told why.

_DATE_CHECKED: set = set()  # (workbook, date, last year's): weighed in this run
_MONTH = r"(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)"
_FILE_MONTH = re.compile(rf"(?i)(?<![a-z]){_MONTH}(?![a-z])[\s_.-]*(\d{{4}}|\d{{2}})(?!\d)")


def _file_months(name: str) -> set[tuple[int, int]]:
    """(year, month) a file name gives in words ("Jun 26", "Jun-26", "June 2026"): a date stamp in digits
    (20260521) is when the file was saved, not its valuation date, so it doesn't count."""
    out = set()
    for m in _FILE_MONTH.finditer(name or ""):
        y = int(m[2])
        out.add((y + 2000 if y < 100 else y, [x[:3] for x in (s.lower() for s in MONTHS)].index(m[1][:3].lower()) + 1))
    return out


def _year_on(iso: str) -> str:
    """A date a year later, a month's end kept at the month's end."""
    import calendar
    y, m, d = map(int, iso[:10].split("-"))
    last = d == calendar.monthrange(y, m)[1]
    return f"{y + 1:04d}-{m:02d}-{calendar.monthrange(y + 1, m)[1] if last else min(d, calendar.monthrange(y + 1, m)[1]):02d}"


def _date_evidence(eid: int, w: dict, cited: str | None) -> dict:
    """What agrees and what disagrees with a workbook's valuation date, besides the cell identify cited:
    {"agree": [...], "disagree": [...]}. A row whose label names another date ("roll forward to 30/9/2025") counts
    neither way; the financial-year end counts only where the model shows it or a person set it (the profile's
    fallback is last year's valuation date's month, which the year-on test already weighs)."""
    import calendar
    import chartdata
    import identify
    date = w["valuation_date"]
    y, m, d = map(int, date.split("-"))
    agree, disagree = [], []
    others = [x for x in identify.candidates(w["db_path"])["valuation_date"] if x["where"] != cited]
    same = [x for x in others if x["value"] == date]
    if same:
        agree.append(f"{same[0]['where']} ({same[0]['why']}) holds it too" + (f", and {len(same) - 1} more" if len(same) > 1 else ""))
    against = [x for x in others if x.get("rank", 1) <= 1 and x.get("value") and x["value"] != date]
    if against:
        disagree.append(f"{against[0]['where']} ({against[0]['why']}) holds {against[0]['value']}")
    months = _file_months(w["filename"])
    if (y, m) in months:
        agree.append(f"the file name says {MONTHS[m - 1][:3]} {y}")
    elif months:
        disagree.append("the file name says " + ", ".join(f"{MONTHS[mm - 1][:3]} {yy}" for yy, mm in sorted(months)))
    prior = _prior_vd(eid)
    if prior and date <= prior[:10]:
        disagree.append(f"not after last year's valuation date ({prior[:10]})")
    elif prior and date == _year_on(prior):
        agree.append(f"a year after last year's valuation date ({prior[:10]})")
    fy = _profile(eid).get("fy_end_month")
    if not fy:
        with closing(rodb.connect(w["db_path"])) as db:
            fy, _ = chartdata.fy_end_detect(db)
    if fy and m == fy and d == calendar.monthrange(y, m)[1]:
        agree.append(f"the model's financial year ends in {MONTHS[fy - 1]}")
    return {"agree": agree, "disagree": disagree}


def _agents_check_dates(eid: int, wbs: list[dict]) -> bool:
    """This year's client model's valuation date, weighed once per date and last year's date (in this run, and kept
    with the workbook's identity across a restart): confirmed by the agents where two or more signals agree and
    none disagrees. Waits for last year's valuation date: without it a model's own stale date can't be told.
    True if it confirmed one."""
    cur = roles(eid).get("current_model") or {}
    w = next((x for x in wbs if x["id"] == cur.get("id") and cur.get("kind") == "workbook"), None)
    if not w or w.get("status") != "done" or w.get("identity_confirmed") or not w.get("valuation_date") or not w.get("db_path"):
        return False
    prior = _prior_vd(eid)
    key = (w["id"], w["valuation_date"], prior)
    if not prior or key in _DATE_CHECKED:
        return False
    _DATE_CHECKED.add(key)
    ident = (library.get(w["id"], full=True) or {}).get("identity") or {}
    done = ident.get("auto_check") or {}
    if (done.get("date"), done.get("prior")) == (w["valuation_date"], prior):
        return False
    try:
        ev = _date_evidence(eid, w, ident.get("valuation_date_evidence"))
    except Exception:  # the check is a bonus: the person is asked as before
        traceback.print_exc()
        return False
    library.note_identity(w["id"], auto_check={"date": w["valuation_date"], "prior": prior, **ev, "at": time.time()})
    if len(ev["agree"]) >= 2 and not ev["disagree"]:
        library.confirm_identity(w["id"], "agents", ev["agree"])
        return True
    return False


def confirm_date(eid: int, fid: int, valuation_date: str) -> dict:
    """A person's check of a workbook's valuation date (the roll-forward runs from last year's to this year's)."""
    w = next((w for w in workbooks(eid) if w["id"] == fid), None)
    if not w:
        raise ValueError("that workbook isn't in this engagement")
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", valuation_date or ""):
        raise ValueError("give the date as YYYY-MM-DD")
    if valuation_date == w.get("valuation_date"):  # right as it is: confirmed, nothing re-linked
        library.confirm_identity(fid, "you")
    else:
        library.set_identity(fid, w.get("target_name"), w.get("project_name"), valuation_date)
    _touch(eid)
    return get(eid)


def remove_workbook(eid: int, fid: int) -> None:
    """Unlink from this engagement only; the workbook stays in the shared library."""
    with _lock, _conn() as db:
        db.execute("DELETE FROM eng_files WHERE engagement_id=? AND file_id=?", (eid, fid))
        db.execute("DELETE FROM roles WHERE engagement_id=? AND kind='workbook' AND ref_id=?", (eid, fid))


def remove_document(did: int) -> None:
    d = _doc(did)
    if not d:
        return
    if d["status"] in ("queued", "processing") or d["facts_status"] == "running":
        raise ValueError("wait for this document to finish processing before removing it")
    with _lock, _conn() as db:
        db.execute("DELETE FROM facts WHERE document_id=?", (did,))
        db.execute("DELETE FROM roles WHERE kind='document' AND ref_id=?", (did,))
        db.execute("DELETE FROM documents WHERE id=?", (did,))
    out_dir, src = Path(d["out_dir"]), Path(d["source_path"])
    if out_dir.is_relative_to(DOCS):
        shutil.rmtree(out_dir, ignore_errors=True)
    others = _q("SELECT id FROM documents WHERE sha256=?", d["sha256"])
    if src.is_relative_to(UPLOADS) and not others and not library.by_sha(d["sha256"]):
        shutil.rmtree(src.parent, ignore_errors=True)


def _touch(eid: int) -> None:
    _set("engagements", eid, updated_at=time.time())


def _doc(did: int) -> dict | None:
    rows = _q("SELECT * FROM documents WHERE id=?", did)
    return rows[0] if rows else None


def document(did: int) -> dict | None:
    """Pages (Markdown) and tables (with checks) of a processed report."""
    d = _doc(did)
    if not d:
        return None
    doc = json.loads(d.pop("doc_json") or "null")
    d.pop("source_path")
    d["loop"] = json.loads(d.pop("loop_json", None) or "null")
    return {**d, "doc": doc}


def _save_doc(did: int, doc: dict) -> None:
    md = docingest.render(doc)
    d = _doc(did)
    Path(d["out_dir"], "document.md").write_text(md, encoding="utf-8")
    _set("documents", did, doc_json=json.dumps(doc, default=str), n_flagged=0,
         n_tables=sum(t.get("status") != "figure" for t in doc["tables"]), pages=len(doc["pages"]))


# ---- facts --------------------------------------------------------------------------------------------------


def set_fact(fact_id: int, action: str, fields: dict | None = None) -> dict:
    rows = _q("SELECT * FROM facts WHERE id=?", fact_id)
    if not rows:
        raise ValueError("no such fact")
    f = rows[0]
    if action not in ("approve", "use_suggestion", "edit", "reject", "reset"):
        raise ValueError(f"unknown action {action}")
    rv = json.loads(f["review_json"] or "null") or {}
    now = time.time()
    if action == "reset" and f["decided_by"] == "agents":  # a person undid the agents' decision: it's theirs now
        a = json.loads(f["agent_json"] or "null") or {}
        _set("facts", fact_id, agent_json=json.dumps({**a, "held": True}))
    _set("facts", fact_id, decided_by=None if action == "reset" else "you")
    if action == "approve":
        _set("facts", fact_id, status="approved", final_json=None, updated_at=now)
    elif action == "use_suggestion":
        s = rv.get("suggestion")
        if not s:
            raise ValueError("the reviewer made no correction to use")
        _set("facts", fact_id, status="approved", final_json=json.dumps({k: s.get(k) for k in FACT_FIELDS}), updated_at=now)
    elif action == "edit":
        final = {k: f[k] for k in FACT_FIELDS}
        final.update({k: v for k, v in (fields or {}).items() if k in FACT_FIELDS})
        if "value" not in (fields or {}):
            keyfacts.settle_value(final)
        d = _doc(f["document_id"])
        pg = keyfacts.pages(json.loads(d["doc_json"])["markdown"])
        final["check"] = keyfacts.check(final, pg)
        _set("facts", fact_id, status="approved", final_json=json.dumps(final), updated_at=now)
    elif action == "reject":
        _set("facts", fact_id, status="rejected", updated_at=now)
    elif action == "reset":
        _set("facts", fact_id, status="pending", final_json=None, updated_at=now)
    else:
        raise ValueError(f"unknown action {action}")
    _touch(f["engagement_id"])
    return next(x for x in facts(f["engagement_id"]) if x["id"] == fact_id)


def agreed(f: dict) -> bool:
    """Both models agree on the fact and the code checks pass (the review loop's outcome, or, for facts from
    before the loop, the reviewer's accept)."""
    a = f.get("agent")
    ok = (f.get("check") or {}).get("ok")
    return bool(ok and (a["status"] == "agreed" if a else (f.get("review") or {}).get("verdict") == "accept"))


def auto_decide(did: int) -> dict:
    """The agents' agreement is the decision: facts they agree on (and that pass the checks) are approved, facts
    they agree to withdraw are rejected, both marked as decided by the agents. Facts a person has decided, or
    whose agents' decision a person undid, are left alone."""
    n = {"approved": 0, "rejected": 0}
    now = time.time()
    for f in _q("SELECT id, check_json, agent_json FROM facts WHERE document_id=? AND status='pending'", did):
        a = json.loads(f["agent_json"] or "null") or {}
        if a.get("held"):
            continue
        if a.get("status") == "agreed" and (json.loads(f["check_json"] or "null") or {}).get("ok"):
            _set("facts", f["id"], status="approved", final_json=None, decided_by="agents", updated_at=now)
            n["approved"] += 1
        elif a.get("status") == "withdrawn":
            _set("facts", f["id"], status="rejected", decided_by="agents", updated_at=now)
            n["rejected"] += 1
    return n


# ---- the review loop and its lessons --------------------------------------------------------------------------

def _names(eid: int, did: int | None = None) -> list[str]:
    """What identifies this engagement, for the lessons' anonymity check: its name, file names, targets and the
    report's identity facts."""
    e = _q("SELECT name FROM engagements WHERE id=?", eid)
    out = [e[0]["name"]] if e else []
    out += [d["filename"] for d in documents(eid)]
    for w in workbooks(eid):
        out += [w["filename"], w.get("target_name"), w.get("project_name")]
    out += [f["value_text"] for f in facts(eid) if f["category"] == "identity"
            and re.search(r"name|client|target|project|asset|company|vendor|purchaser|owner", f["key"] or "")]
    return [x for x in out if x]


def _learn(did: int, scope: str, episodes: list[dict]) -> dict | None:
    """Distil a loop's episodes into lessons (reviewer model), and note the result on the document."""
    if not episodes:
        return None
    d = _doc(did)
    eid = d["engagement_id"]
    e = _q("SELECT reviewer_model FROM engagements WHERE id=?", eid)[0]
    md = json.loads(d["doc_json"] or "{}").get("markdown", "")
    try:
        got = lessons.distil(scope, episodes, _names(eid, did), md, e["reviewer_model"] or DEFAULT_REVIEWER,
                             _logger(eid), engagement=eid)
    except Exception as ex:  # learning is a bonus: never fail the step for it
        traceback.print_exc()
        got = {"error": friendly(ex)}
    _note_loop(did, **{f"lessons_{scope}": got})
    return got


def _note_loop(did: int, **fields) -> None:
    d = _q("SELECT loop_json FROM documents WHERE id=?", did)
    loop = json.loads((d[0]["loop_json"] if d else None) or "{}")
    loop.update(fields)
    _set("documents", did, loop_json=json.dumps(loop, default=str))


def calls_view(eid: int) -> dict:
    """Tokens, cost and time for an engagement: model calls by file and by step (calllog.py), plus how long each
    job took on the clock (a job's time also counts waiting for rate-limit room and work without model calls)."""
    wbs, docs = workbooks(eid), documents(eid)
    e = _q("SELECT * FROM engagements WHERE id=?", eid)
    if not e:
        raise ValueError("no such engagement")
    e = e[0]
    jobs = []
    for d in docs:
        jobs += [{"file": d["filename"], "document": d["id"], "job": "read the report's text", "secs": d["doc_secs"]},
                 {"file": d["filename"], "document": d["id"], "job": "key facts and their review loop", "secs": d["facts_secs"]}]
    for w in wbs:
        took = (w["processed_at"] - w["started_at"]) if w.get("processed_at") and w.get("started_at") and \
            w["processed_at"] > w["started_at"] else None
        jobs.append({"file": w["filename"], "workbook": w["id"], "job": "process the workbook", "secs": took,
                     "build_secs": w.get("build_secs")})
    jobs += [{"file": None, "job": label, "secs": e.get(f"{k}_secs")} for k, label in
             (("overlay", "python overlay"), ("rows", "row agents"), ("doctor", "overlay doctor"), ("map", "map"))]
    return {"calls": calllog.breakdown(eid, [w["id"] for w in wbs]), "jobs": [j for j in jobs if j["secs"]],
            "names": {"documents": {d["id"]: d["filename"] for d in docs}, "workbooks": {w["id"]: w["filename"] for w in wbs}},
            "log_file": str(calllog.DB.relative_to(ROOT))}


def call_view(eid: int, cid: int) -> dict | None:
    return calllog.call(cid, eid, [w["id"] for w in workbooks(eid)])


def lessons_view() -> dict:
    return {"curated": lessons.curated(), "learned": lessons.load()["lessons"], "rules_file": "docs/report_rules.md"}


# ---- roles --------------------------------------------------------------------------------------------------

def _wb_inputs(eid: int) -> list[dict]:
    out = []
    for w in workbooks(eid):
        if w["status"] == "done" and w["db_path"] and Path(w["db_path"]).exists():
            extlinks.ensure(w["source_path"], w["db_path"])
            out.append(w)
    return out


def busy_doc(d: dict) -> bool:
    return d["status"] in ("queued", "processing") or d["facts_status"] in ("queued", "running")


def _roles_key(eid: int, wbs: list[dict], docs: list[dict]) -> str:
    """What a suggestion depends on: the processed files and the report facts it navigates by."""
    ref = [(f["id"], f["value_text"], f["page"], f["approved"]) for f in reference(eid)]
    return json.dumps([sorted(w["id"] for w in wbs if w["status"] == "done"),
                       sorted(d["id"] for d in docs if d["status"] == "done"), ref], default=str)


_SUGGESTING = threading.Lock()  # one suggestion at a time: two at once wrote the same model.db and locked each other out


def suggest_roles(eid: int) -> dict:
    """The rules' suggestion (roles.py) and the reviewer model's second opinion on the same evidence. Stored so the
    page can show it; roles not yet confirmed take the suggestion. Returns the suggestion."""
    with _SUGGESTING, calllog.tag(engagement=eid, step="roles"):
        return _suggest_roles(eid, _roles_key(eid, workbooks(eid), documents(eid)))


def _suggest_roles(eid: int, key: str) -> dict:
    docs = [d for d in documents(eid) if d["status"] == "done"]
    n_facts = {d["id"]: sum(1 for f in facts(eid) if f["document_id"] == d["id"]) for d in docs}
    docs_in = [{**d, "n_facts": n_facts[d["id"]]} for d in docs]
    res = rolesmod.suggest(docs_in, _wb_inputs(eid), reference(eid))
    res.update(key=key, at=time.time())
    e = _q("SELECT reviewer_model FROM engagements WHERE id=?", eid)[0]
    try:  # a second opinion from the reviewer model on the same evidence
        res["second_opinion"] = rolesmod.second_opinion(docs_in, _wb_inputs(eid), reference(eid), res,
                                                        e["reviewer_model"] or DEFAULT_REVIEWER, _logger(eid))
    except Exception as ex:
        res["second_opinion"] = {"error": friendly(ex)}
    _set("engagements", eid, roles_suggested=json.dumps(res, default=str), updated_at=time.time())
    have = roles(eid)
    with _lock, _conn() as db:
        for role, r in res["roles"].items():
            if not have.get(role, {}).get("confirmed"):
                db.execute("INSERT OR REPLACE INTO roles(engagement_id, role, kind, ref_id, sheets_json, why_json, "
                           "confirmed) VALUES (?,?,?,?,?,?,0)",
                           (eid, role, r["kind"], r["id"], json.dumps(r["sheets"]), json.dumps(r["why"])))
    return res


def confirm_roles(eid: int, assignments: dict, by: str = "you", evidence: dict | None = None) -> dict:
    """assignments: {role: {"kind", "id", "sheets"} or None to clear}. Everything given is confirmed, by "you" or
    by "orchestrator" (with the evidence it confirmed on). The map, the Python overlay and the result were built
    on the old roles, so they go."""
    have = roles(eid)
    with _lock, _conn() as db:
        for role, a in assignments.items():
            if role not in rolesmod.ROLES:
                raise ValueError(f"unknown role {role}")
            if not a:
                db.execute("DELETE FROM roles WHERE engagement_id=? AND role=?", (eid, role))
                continue
            why = have.get(role, {}).get("why") if have.get(role, {}).get("id") == a["id"] else ["set by you"]
            db.execute("INSERT OR REPLACE INTO roles(engagement_id, role, kind, ref_id, sheets_json, why_json, confirmed, "
                       "confirmed_by, evidence_json) VALUES (?,?,?,?,?,?,1,?,?)",
                       (eid, role, a["kind"], int(a["id"]), json.dumps(a.get("sheets")), json.dumps(why or []), by,
                        json.dumps((evidence or {}).get(role) if evidence else None)))
        changed = any((have.get(r) or {}).get("id") != (a or {}).get("id") or
                      sorted((have.get(r) or {}).get("sheets") or []) != sorted((a or {}).get("sheets") or [])
                      for r, a in assignments.items())
        if changed:
            db.execute("""UPDATE engagements SET updated_at=?, map_status=NULL, map_json=NULL, map_error=NULL,
                          overlay_status=NULL, overlay_json=NULL, overlay_error=NULL, result_json=NULL WHERE id=?""",
                       (time.time(), eid))
    if changed:
        _SESSIONS.pop(eid, None)
    return get(eid)


def _role_wb(eid: int, role: str) -> dict | None:
    r = roles(eid).get(role)
    if not r or r["kind"] != "workbook":
        return None
    w = next((w for w in workbooks(eid) if w["id"] == r["id"]), None)
    if not w or w["status"] != "done":
        return None
    return {**w, "sheets": set(r["sheets"]) if r["sheets"] else None, "confirmed": r["confirmed"]}


# ---- the map ------------------------------------------------------------------------------------------------

def _map(eid: int) -> None:
    import valuation
    ov, prior, cur = _role_wb(eid, "prior_overlay"), _role_wb(eid, "prior_model"), _role_wb(eid, "current_model")
    ref = reference(eid)
    step = lambda msg: _set("engagements", eid, map_status="running", map_step=msg)
    step("Finding the report's figures in the overlay")
    report_overlay = linkmap.match_facts(ov["db_path"], ref, ov["sheets"])
    step("Recomputing the overlay's DCFs in Python")
    try:
        cat = [a for a in valuation.catalogue(ov["db_path"]) if not ov["sheets"] or a["cell"].split("!")[0] in ov["sheets"]]
    except Exception:
        cat = []
        traceback.print_exc()
    anchors_at = {}
    for fm in report_overlay:
        for m in (x for x in fm["matches"] if x.get("located")):
            anchors_at.setdefault(f"{m['sheet']}!{m['addr']}", []).append(f"{fm['label'] or fm['key']} {fm['value_text']}")
    python = [{"cell": a["cell"], "label": a["label"], "value": a["value"], "python": a.get("total"),
               "reproduced": bool(a.get("matches")), "ok": a.get("ok"), "reason": a.get("reason"),
               "report": anchors_at.get(a["cell"], [])} for a in cat]
    step("Following the overlay's links into the prior client model")
    copy = _wiring(eid).get("client_sheets") if ov and prior else None
    if copy:  # the overlay reads its own copy of the client sheets; that copy's rows are the prior model's rows
        ov_client = linkmap.overlay_to_client({**ov, "sheets": set(ov["sheets"] or [])}, {**ov, "sheets": set(copy)})
        ov_client["mode"] = "inside a copy of the prior client model"
    else:
        ov_client = linkmap.overlay_to_client(ov, prior) if ov else {"links": [], "books": []}
    alignment = {}
    if prior and cur:
        step("Lining up those rows with the current client model")
        refs = [(ln["client_sheet"], ln["client_row"]) for ln in ov_client["links"] if ln.get("client_row")
                and (ov_client.get("client_link") is None or ln.get("link") == ov_client.get("client_link"))]
        alignment = linkmap.align_rows(prior["db_path"], cur["db_path"], refs)
    out = {"report_overlay": report_overlay, "python": python, "overlay_client": {**ov_client, "links": None},
           "chain": linkmap.chain(ov_client, alignment), "reference_approved": all(f["approved"] for f in ref) and bool(ref),
           "files": {k: (v["filename"] if v else None) for k, v in
                     (("overlay", ov), ("prior_model", prior), ("current_model", cur))}}
    if prior and cur and prior["id"] != cur["id"]:
        step("Comparing last year's and this year's client models cell by cell")
        try:  # what changed, in code (no model calls): counts and the biggest moves, for the Map's cards
            import diff as diffmod
            d = diffmod.diff(prior["db_path"], cur["db_path"])
            ours = set(ov["sheets"] or []) if ov and ov["id"] == prior["id"] else set()
            d["sheets"]["removed"] = [x for x in d["sheets"]["removed"] if x not in ours]
            out["changes"] = {"counts": d.get("counts"), "sheets": d.get("sheets"),
                              **{k: (d.get(k) or [])[:15] for k in ("key_outputs", "inputs", "rows_added", "rows_removed",
                                                                    "formula_rows")}}
        except Exception as ex:
            traceback.print_exc()
            out["changes"] = {"error": friendly(ex)}
    _set("engagements", eid, map_json=json.dumps(out, default=str), map_status="done", map_step="Done",
         updated_at=time.time())


# ---- model dashboards (map step) ----------------------------------------------------------------------------

_ROW_REF = re.compile(r"^'?(.+?)'?!r(\d+)$")
_CELL_REF = re.compile(r"^'?(.+?)'?!\$?[A-Z]{1,3}\$?(\d+)$")


def _engagement_wb(eid: int, fid: int) -> dict:
    w = next((w for w in workbooks(eid) if w["id"] == fid), None)
    if not w:
        raise ValueError("that workbook isn't in this engagement")
    if w["status"] != "done" or not w["db_path"] or not Path(w["db_path"]).exists():
        raise ValueError(f"{w['filename']} hasn't finished processing")
    return w


def model_marks(eid: int, fid: int) -> dict[tuple[str, int], list[str]]:
    """What the map and the Python overlay say about this workbook's rows: report figures found there, overlay rows
    reading the client model, client rows the overlay reads, this year's matches, levers and outputs."""
    rl = roles(eid)
    has = lambda role: (rl.get(role) or {}).get("kind") == "workbook" and rl[role]["id"] == fid
    e = _q("SELECT map_json, overlay_json FROM engagements WHERE id=?", eid)[0]
    m = json.loads(e["map_json"] or "null") or {}
    o = json.loads(e["overlay_json"] or "null") or {}
    marks: dict[tuple[str, int], list[str]] = {}

    def mark(ref, pattern, text):
        hit = pattern.match(ref or "")
        if hit:
            tags = marks.setdefault((hit[1], int(hit[2])), [])
            if text not in tags:
                tags.append(text)

    chain = m.get("chain") or []
    if has("prior_overlay"):
        for x in m.get("report_overlay") or []:
            if x["matches"] and x["matches"][0].get("located", True):  # a value-only match isn't where the figure is
                b = x["matches"][0]
                mark(f"{b['sheet']}!r{b['row']}", _ROW_REF, f"report: {x.get('label') or x['key']} {x['value_text']}")
        for c in chain:
            mark(c["overlay"], _ROW_REF, f"reads client {c['client']}")
        for lv in o.get("levers") or []:
            mark(lv["cell"], _CELL_REF, f"lever: {lv['label']}")
        for x in o.get("outputs") or []:
            mark(x["cell"], _CELL_REF, f"output: {x['label']}")
    if has("prior_model"):
        for c in chain:
            mark(c["client"], _ROW_REF, f"read by overlay {c['overlay']}")
    if has("current_model"):
        for c in chain:
            if c.get("current"):
                mark(c["current"], _ROW_REF, f"this year's {c['client']} (read by overlay {c['overlay']})")
    return marks


def model_dashboard(eid: int, fid: int) -> dict:
    import modeldash
    w = _engagement_wb(eid, fid)
    rl = roles(eid)
    as_roles = [k for k, r in rl.items() if r["kind"] == "workbook" and r["id"] == fid]
    ov = rl.get("prior_overlay") or {}
    marks = model_marks(eid, fid)
    per_sheet: dict[str, int] = {}
    for s, _ in marks:
        per_sheet[s] = per_sheet.get(s, 0) + 1
    return {**modeldash.summary(w["db_path"]), "id": fid, "filename": w["filename"], "roles": as_roles,
            "target_name": w["target_name"], "valuation_date": w["valuation_date"],
            "overlay_sheets": (ov.get("sheets") or []) if ov.get("id") == fid else [],
            "marked": {"rows": len(marks), "by_sheet": per_sheet}}


def model_rows(eid: int, fid: int, sheet: str | None, q: str | None, mapped: bool, limit: int = 200) -> list[dict]:
    import modeldash
    w = _engagement_wb(eid, fid)
    marks = model_marks(eid, fid)
    out = modeldash.rows(w["db_path"], sheet or None, q or None, list(marks) if mapped else None, limit)
    for r in out:
        r["marks"] = marks.get((r["sheet"], r["row"]), [])
    return out


# ---- the Python overlay -------------------------------------------------------------------------------------

_SESSIONS: dict[int, tuple] = {}  # engagement -> (overlay.Session, summary): the compiled module, loaded and wired


def _wiring(eid: int) -> dict:
    """Which files and sheets the overlay module reads, from the confirmed (or suggested) roles."""
    ov, prior, cur = _role_wb(eid, "prior_overlay"), _role_wb(eid, "prior_model"), _role_wb(eid, "current_model")
    names = {w["id"]: w["sheet_names"] for w in workbooks(eid)}
    sheets = [s for s in names[ov["id"]] if not ov["sheets"] or s in ov["sheets"]]
    same_file = bool(prior) and prior["id"] == ov["id"]
    # the overlay in a copy of the client model, the client's own file assigned as the prior model: the overlay
    # reads its copy's sheets, which are fed from that file
    copy = [s for s in names[ov["id"]] if s not in sheets and s in set(names.get(prior["id"], []))] \
        if prior and not same_file and ov["sheets"] else []
    client_link = None
    if prior and not same_file and not copy:
        extlinks.ensure(ov["source_path"], ov["db_path"])
        client_link = linkmap.overlay_to_client({**ov, "sheets": None}, {**prior, "sheets": None}).get("client_link")
    vd = next((f for f in reference(eid) if f["key"] == "valuation_date" and f.get("value")), None)
    prior_vd = None
    if vd:
        v = int(vd["value"])
        prior_vd = f"{v // 10000:04d}-{v // 100 % 100:02d}-{v % 100:02d}"
    elif library.get(ov["id"]) and library.get(ov["id"])["valuation_date"]:
        prior_vd = library.get(ov["id"])["valuation_date"]
    return {"overlay": {"db_path": ov["db_path"], "filename": ov["filename"], "sheets": sheets,
                        "source_path": ov.get("source_path")},
            "prior": {"db_path": prior["db_path"], "filename": prior["filename"],
                      "sheets": sorted(prior["sheets"]) if prior["sheets"] else None,
                      "valuation_date": (library.get(prior["id"]) or {}).get("valuation_date")} if prior else None,
            "current": {"db_path": cur["db_path"], "filename": cur["filename"], "sheets": None,
                        "valuation_date": (library.get(cur["id"]) or {}).get("valuation_date")} if cur else None,
            "client_link": client_link, "prior_valuation_date": prior_vd, "same_file": same_file,
            "client_sheets": copy or None}


def _overlay(eid: int) -> None:
    import overlay as ovmod
    e = _q("SELECT name FROM engagements WHERE id=?", eid)[0]
    step = lambda frac, msg: _set("engagements", eid, overlay_status="running", overlay_step=msg,
                                  overlay_pct=round(frac, 3))
    step(0, "Reading the roles")
    w = _wiring(eid)
    _SESSIONS.pop(eid, None)
    summary, sess = ovmod.build(OUT / "overlays" / f"e{eid}", w["overlay"], w["prior"], w["current"], reference(eid),
                                e["name"], w["client_link"], w["prior_valuation_date"], step, client_sheets=w["client_sheets"])
    summary["wiring"] = w
    summary["reference_approved"] = all(f["approved"] for f in reference(eid)) and bool(reference(eid))
    ovmod.deep(_load_holds, eid, sess)
    _load_rowpicks(eid, sess)
    _sync_roll(eid, sess, summary)
    _SESSIONS[eid] = (sess, summary)
    _set("engagements", eid, overlay_json=json.dumps(summary, default=str), overlay_status="done", overlay_step="Done",
         updated_at=time.time())


def _this_year_file(eid: int) -> Path:
    return OUT / "overlays" / f"e{eid}" / "this_year.json"


def this_year_date(eid: int) -> str | None:
    """This year's valuation date as set for the engagement (a client model's own date is the model's, which a
    fresh model built for another date doesn't make this year's), or None."""
    f = _this_year_file(eid)
    try:
        return json.loads(f.read_text(encoding="utf-8")).get("valuation_date") if f.exists() else None
    except (OSError, ValueError):
        return None


def set_this_year_date(eid: int, valuation_date: str | None) -> dict:
    """Set (or clear, with None) this year's valuation date for the engagement: the roll-forward runs to it."""
    if valuation_date and not re.match(r"^\d{4}-\d{2}-\d{2}$", valuation_date):
        raise ValueError("give the date as YYYY-MM-DD")
    last = _dates(eid)["dates"]
    base = last.get("overlay") or last.get("prior_client")
    if valuation_date and base and valuation_date <= base[:10]:
        raise ValueError(f"this year's valuation date must be after last year's ({base[:10]})")
    f = _this_year_file(eid)
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps({"valuation_date": valuation_date}), encoding="utf-8")
    if eid in _SESSIONS:
        sess, summary = _SESSIONS[eid]
        _sync_roll(eid, sess, summary)
    _touch(eid)  # the roll moved: the orchestrator runs the row agents and the result again
    return {"valuation_date": valuation_date}


def _dates(eid: int) -> dict:
    """The valuation dates the roll-forward uses, as the files and the report give them now: last year's (the
    report's, else the overlay file's), last year's and this year's client models', this year's as set for the
    engagement, and which are confirmed."""
    vd = next((f for f in reference(eid) if f["key"] == "valuation_date" and f.get("value")), None)
    out, confirmed = {}, {}
    for role, key in (("prior_overlay", "overlay"), ("prior_model", "prior_client"), ("current_model", "current_client")):
        w = _role_wb(eid, role)
        rec = library.get(w["id"]) if w else None
        out[key] = (rec or {}).get("valuation_date")
        confirmed[key] = bool((rec or {}).get("identity_confirmed"))
    if vd:
        v = int(vd["value"])
        out["overlay"] = f"{v // 10000:04d}-{v // 100 % 100:02d}-{v % 100:02d}"
        confirmed["overlay"] = bool(vd.get("approved"))
    out["this_year"] = this_year_date(eid)
    confirmed["this_year"] = bool(out["this_year"])
    return {"dates": out, "confirmed": confirmed}


def held_file(eid: int) -> Path:
    return OUT / "overlays" / f"e{eid}" / "held.json"


def held_values(eid: int) -> dict:
    """This year's figures a person set for the inputs held at last year's: {cell: {"value", "label", "by", "from",
    "was", "at"}}."""
    try:
        return json.loads(held_file(eid).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def set_held(eid: int, cell: str, value: float | None, source: str = "typed") -> dict:
    """This year's figure for an input held at last year's (value None: back to last year's). Only an input the
    result lists as held; source: "suggestion" (accepted from this year's model) or "typed"."""
    res = (get(eid) or {}).get("result") or {}
    item = next((x for x in res.get("held") or [] if x["cell"] == cell), None)
    if not item:
        raise ValueError(f"{cell} isn't an input held at last year's")
    if source not in ("suggestion", "typed"):
        raise ValueError("source is suggestion or typed")
    vals = held_values(eid)
    if value is None:
        vals.pop(cell, None)
    else:
        vals[cell] = {"value": float(value), "label": item["label"], "by": "you", "from": source, "was": item["value"],
                      "at": time.time()}
    f = held_file(eid)
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps(vals, indent=1), encoding="utf-8")
    if eid in _SESSIONS:
        _SESSIONS[eid][1]["held_values"] = vals
    _touch(eid)  # the result's inputs changed: the orchestrator works it out again
    return {"cell": cell, "label": item["label"], "value": value, "was": item["value"], "from": source}


def rate_file(eid: int) -> Path:
    return OUT / "overlays" / f"e{eid}" / "this_year_rate.json"


def this_year_rate(eid: int) -> dict:
    """This year's discount rate as a person set it: {"low", "high" (the low end at the higher rate), "by", "at"},
    or {} for last year's."""
    try:
        return json.loads(rate_file(eid).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _rate_in(v) -> float | None:
    """A discount rate as typed: 8.9, "8.9%", 0.089 -> 0.089."""
    if v is None or (isinstance(v, str) and not v.strip()):
        return None
    t = str(v).replace(",", "").strip()
    x = float(t.rstrip("%"))
    x = x / 100 if t.endswith("%") or x >= 1 else x
    if not 0 < x < 0.5:
        raise ValueError(f"{v} isn't a discount rate")
    return x


def set_this_year_rate(eid: int, low, high=None) -> dict:
    """This year's discount rate (low None: back to last year's): a range's two ends, or one rate for both. The low
    end of the value is at the higher rate, whichever order they're given in. The roll-forward runs at it, and the
    bridge has a step of its own for it."""
    a, b = _rate_in(low), _rate_in(high)
    f = rate_file(eid)
    if a is None and b is None:
        f.unlink(missing_ok=True)
        got = {}
    else:
        a, b = (a, b if b is not None else a) if a is not None else (b, b)
        got = {"low": max(a, b), "high": min(a, b), "by": "you", "at": time.time()}
        f.parent.mkdir(parents=True, exist_ok=True)
        f.write_text(json.dumps(got), encoding="utf-8")
    if eid in _SESSIONS:
        _SESSIONS[eid][1]["this_year_rate"] = got
    _touch(eid)  # the result's inputs changed: the orchestrator works it out again
    return got


def method_file(eid: int) -> Path:
    return OUT / "overlays" / f"e{eid}" / "method.json"


def preferred_method(eid: int) -> str | None:
    """The method a person prefers for this year's value (methods.py), or None for the default."""
    try:
        return json.loads(method_file(eid).read_text(encoding="utf-8")).get("key")
    except (OSError, ValueError):
        return None


def set_method(eid: int, key: str | None) -> dict:
    """The method this year's value is worked out by (None: the default). The bridge then has a step of its own
    for the move from the default to it."""
    import methods
    f = method_file(eid)
    if key in (None, "", methods.DEFAULT):
        f.unlink(missing_ok=True)
        key = None
    elif key not in methods.LABEL:
        raise ValueError(f"no method {key}: one of {', '.join(methods.LABEL)}")
    else:
        f.parent.mkdir(parents=True, exist_ok=True)
        f.write_text(json.dumps({"key": key, "by": "you", "at": time.time()}), encoding="utf-8")
    if eid in _SESSIONS:
        _SESSIONS[eid][1]["method"] = key
    _touch(eid)  # the result's inputs changed: the orchestrator works it out again
    return {"key": key or methods.DEFAULT, "label": methods.LABEL[key or methods.DEFAULT], "default": key is None}


def _sync_roll(eid: int, sess, summary: dict) -> None:
    """The roll-forward from the dates as they are now (a file's valuation date can be corrected after the build),
    last year's valuation date and the discountings' cut-off set on the session."""
    import overlay as ovmod
    want = _profile(eid).get("horizon")  # set in the engagement's profile, else worked out
    if getattr(sess, "horizon_set", None) != want:
        sess.horizon_set = want
        sess._pshift.clear()
    summary["held_values"] = held_values(eid)  # this year's figures a person set, on this year's feed
    summary["this_year_rate"] = this_year_rate(eid)  # and this year's discount rate (result.this_year_rate)
    summary["method"] = preferred_method(eid)  # and the method this year's value is worked out by (methods.py)
    roll = summary.get("roll")
    if not roll or not summary["wiring"].get("current"):
        return
    if "valuation_date_reads" not in roll or "cutoff" not in roll:
        # built before every date the discountings read was moved, or before their periods were cut off: work them out
        lever = next((l for l in summary.get("levers") or [] if l["key"] == "valuation_date"), None)
        roll.update(ovmod.deep(ovmod.date_cells, summary["wiring"]["overlay"]["db_path"], summary.get("outputs") or [],
                               summary.get("sheets"), lever))
    sess.cutoffs = [(*ovmod.parse_a1(c), ovmod.serial(ovmod.date.fromisoformat(d))) for c, d in roll.get("cutoff") or []]
    now = _dates(eid)
    d = now["dates"]
    if d != roll.get("dates") or roll.get("plan") != ovmod.ROLL_PLAN or roll.get("horizon_set") != want:
        w = summary["wiring"]
        fresh = ovmod.deep(ovmod.plan_roll, sess, w.get("prior"), w["overlay"], w.get("same_file"), d["overlay"],
                           d["prior_client"], d["current_client"], d.get("this_year"))
        roll.update(fresh)
    elif roll.get("prior_valuation_date"):
        sess.base_vd = ovmod.serial(ovmod.date.fromisoformat(roll["prior_valuation_date"][:10]))
    roll["confirmed"] = now["confirmed"]


def overlay_session(eid: int):
    """The live module for an engagement, loaded from its saved build after a restart."""
    import overlay as ovmod
    if eid in _SESSIONS:
        return _SESSIONS[eid]
    rows = _q("SELECT overlay_json FROM engagements WHERE id=?", eid)
    summary = json.loads(rows[0]["overlay_json"] or "null") if rows else None
    if not summary or not Path(summary["module"]).exists():
        raise ValueError("build the Python overlay first")
    w = summary["wiring"]
    sess = ovmod.Session(summary["module"], w["overlay"]["db_path"], w["overlay"]["sheets"],
                         None if w["same_file"] else (w["prior"] or {}).get("db_path"), (w["current"] or {}).get("db_path"),
                         w["client_link"], w.get("client_sheets") or ((w["prior"] or {}).get("sheets") if w["same_file"] else None))
    ovmod.deep(_load_holds, eid, sess)
    _load_rowpicks(eid, sess)  # the agents' picks from an older row finder are set aside (the rows stage runs again)
    _sync_roll(eid, sess, summary)
    _SESSIONS[eid] = (sess, summary)
    return _SESSIONS[eid]


def _live(eid: int, mode: str, changes: dict | None) -> tuple:
    """The engagement's session and summary, after checking the feed exists; changes with dates as serials."""
    import overlay as ovmod
    sess, summary = overlay_session(eid)
    if mode not in ("workbook", "prior", "current"):
        raise ValueError("the feed must be workbook, prior or current")
    if mode == "current" and not summary["wiring"].get("current"):
        raise ValueError("assign the current client model (Roles) and rebuild in Python to roll forward")
    if mode == "prior" and not summary["wiring"].get("prior"):
        raise ValueError("assign the prior client model (Roles) to feed from it")
    clean = {}
    for cell, v in (changes or {}).items():
        if isinstance(v, str) and re.match(r"^\d{4}-\d{2}-\d{2}$", v):
            v = ovmod.serial(ovmod.date.fromisoformat(v))
        elif isinstance(v, str):
            v = float(v.replace(",", "").rstrip("%")) / (100 if v.strip().endswith("%") else 1)
        clean[cell] = v
    return sess, summary, clean


def overlay_run(eid: int, mode: str, changes: dict, valuation_date: str | None, months: int | None) -> dict:
    import overlay as ovmod
    sess, summary, clean = _live(eid, mode, changes)
    _sync_roll(eid, sess, summary)
    return ovmod.deep(ovmod.scenario, sess, summary, mode, clean, valuation_date, months)


def overlay_valuation(eid: int, cell: str | None) -> dict:
    """The Model Desk's Validate step on the overlay: a DCF found in the workbook, recomputed step by step from the
    values Excel saved."""
    import overlay as ovmod
    import valuation
    _, summary = overlay_session(eid)
    anchors = ovmod.dcf_anchors(summary)
    usable = [a for a in anchors if a.get("ok")]
    if not usable:
        return {"anchors": valuation._listing(anchors), "selected": None}
    pick = next((a for a in usable if a["cell"] == cell), usable[0])
    v = valuation.validation(summary["wiring"]["overlay"]["db_path"], pick["cell"], anchors=anchors)
    v["anchors"] = valuation._listing(anchors)
    return v


def _fy_hint(eid: int):
    """The engagement's financial-year end (chartdata.fy_end_hint), for totalling cash flows by financial year: the
    profile's if a person set it (it wins over what a model says), else last year's valuation date's month for a
    model that shows none."""
    import chartdata
    mine = _profile(eid).get("fy_end_month")
    if mine:
        return chartdata.fy_end_hint(mine, forced=True)
    vd = _prior_vd(eid)
    return chartdata.fy_end_hint(int(vd[5:7]) if vd else None)


def _prior_vd(eid: int) -> str | None:
    """Last year's valuation date (ISO): the report's fact, else the roll's from the overlay build."""
    vd = next((f for f in reference(eid) if f["key"] == "valuation_date" and f.get("value")), None)
    if vd:
        v = int(vd["value"])
        return f"{v // 10000:04d}-{v // 100 % 100:02d}-{v % 100:02d}"
    rows = _q("SELECT overlay_json FROM engagements WHERE id=?", eid)
    return ((json.loads((rows[0]["overlay_json"] if rows else None) or "null") or {}).get("roll") or {}).get("prior_valuation_date")


# ---- the engagement's profile -------------------------------------------------------------------------------
# The facts about an engagement's models that each step would otherwise guess on its own, in one place: detected,
# with how, and set by a person where the detection is wrong. The financial-year end and the horizon drive the
# cash-flow chart and the roll-forward; the period frequency, the units and the discounting convention are shown
# (the convention is read back from each DCF's own factors).

MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
          "November", "December"]


def _profile(eid: int) -> dict:
    """A person's settings: {"fy_end_month": 6, "horizon": "fixed"}."""
    rows = _q("SELECT profile_json FROM engagements WHERE id=?", eid)
    return json.loads((rows[0]["profile_json"] if rows else None) or "null") or {}


def set_profile(eid: int, fields: dict) -> dict:
    """Set (or, with None, clear) the financial-year end month (1-12) or the horizon ("fixed" / "rolling")."""
    mine = _profile(eid)
    for k, v in fields.items():
        if k == "fy_end_month" and v not in (None, ""):
            if not (isinstance(v, int) and 1 <= v <= 12):
                raise ValueError("the financial year ends in a month, 1 to 12")
        elif k == "horizon" and v not in (None, ""):
            if v not in ("fixed", "rolling"):
                raise ValueError('the horizon is "fixed" or "rolling"')
        elif k not in ("fy_end_month", "horizon"):
            raise ValueError(f"{k} isn't set here: it's read from the models")
        if v in (None, ""):
            mine.pop(k, None)
        else:
            mine[k] = v
    _set("engagements", eid, profile_json=json.dumps(mine), updated_at=time.time())
    return profile_view(eid)


class _Timelines:
    """A model's period dates by sheet, read from each sheet's timeline row alone (overlay.Workbook reads whole
    sheets: on a large model, the profile would read most of the workbook)."""

    def __init__(self, path: str):
        self.db = rodb.connect(path)

    def timeline(self, s: str) -> dict[int, float]:
        from xlruntime import from_db
        lay = self.db.execute("SELECT layout FROM sheets WHERE sheet=?", (s,)).fetchone()
        hr = json.loads((lay[0] if lay else None) or "{}").get("header_row")
        out = {}
        for c, v in self.db.execute("SELECT col, value FROM cells WHERE sheet=? AND row=?", (s, hr)) if hr else []:
            x = from_db(v)
            if isinstance(x, float) and 3000 < x < 120000:
                out[c] = x
        return out

    def close(self):
        self.db.close()


def _frequency(wb) -> tuple[str | None, str]:
    """The periods' length most sheets with a timeline have: monthly, quarterly, half-yearly or annual."""
    kinds = Counter()
    for (s,) in wb.db.execute("SELECT sheet FROM sheets"):
        ends = sorted(set(wb.timeline(s).values()))
        if len(ends) < 3:
            continue
        gap = sorted(b - a for a, b in zip(ends, ends[1:]))[(len(ends) - 1) // 2]
        kinds["monthly" if gap <= 35 else "quarterly" if gap <= 100 else "half-yearly" if gap <= 200
              else "annual" if gap <= 400 else "longer than a year"] += 1
    if not kinds:
        return None, "no sheet with a timeline"
    k, n = kinds.most_common(1)[0]
    return k, f"{n} of {sum(kinds.values())} sheet(s) with a timeline" + (
        f" (also {', '.join(f'{x}: {m}' for x, m in kinds.items() if x != k)})" if len(kinds) > 1 else "")


def profile_view(eid: int) -> dict:
    """The profile: each fact detected (value, how), a person's setting where there is one (set), what it drives.
    Reads the models' own files; no model calls."""
    import chartdata
    import overlay as ovmod
    mine = _profile(eid)
    prior, cur, ovw = _role_wb(eid, "prior_model"), _role_wb(eid, "current_model"), _role_wb(eid, "prior_overlay")
    out = {}
    # the financial year: what the model shows, else last year's valuation date's month
    src = cur or prior
    month, how = None, "no client model assigned yet"
    if src:
        with closing(rodb.connect(src["db_path"])) as db:  # closed: Windows won't rebuild an open file
            month, how = chartdata.fy_end_detect(db)
        how = f"{how}, in {src['filename']}" if month else f"{src['filename']} shows none: {how}"
    vd = _prior_vd(eid)
    if not month and vd:
        month, how = int(vd[5:7]), f"last year's valuation date ({vd}); {how}"
    out["fy_end_month"] = {"label": "Financial year ends", "value": mine.get("fy_end_month") or month, "detected": month,
                           "shown": MONTHS[(mine.get("fy_end_month") or month) - 1] if (mine.get("fy_end_month") or month) else None,
                           "how": how if month else f"{how}; December is assumed", "set": "fy_end_month" in mine,
                           "drives": "the cash-flow chart: each model's periods are totalled by financial year"}
    # the horizon: this year's client model ends its periods where last year's did, or later
    kind, n = None, {"fixed": 0, "rolling": 0}
    live = _SESSIONS.get(eid)
    if live and live[0].current and live[0].rowmap:
        sess = live[0]
        kind, n = ovmod.horizon(sess.prior or sess.ov, sess.current, sess.client_sheets or {k[1] for k in sess.ext_cached},
                                sess.rowmap.sheet_for)
    elif (prior or ovw) and cur:
        base = prior or ovw
        a, b = _Timelines(base["db_path"]), _Timelines(cur["db_path"])
        try:
            sheets = base["sheets"] or [s for (s,) in a.db.execute("SELECT sheet FROM sheets")]
            kind, n = ovmod.horizon(a, b, sheets)
        finally:
            a.close()
            b.close()
    how = (f"this year's client model ends its periods where last year's did on {n['fixed']} sheet(s), later on "
           f"{n['rolling']}" if kind else "needs last year's and this year's client models, with timelines")
    out["horizon"] = {"label": "Horizon", "value": mine.get("horizon") or kind, "detected": kind, "how": how,
                      "set": "horizon" in mine,
                      "shown": {"fixed": "Fixed: the periods end on the same date each year",
                                "rolling": "Rolling: the periods move on each year"}.get(mine.get("horizon") or kind),
                      "drives": "the roll-forward: on a fixed horizon the periods stay and only the valuation date moves"}
    # shown, not set: the period frequency, the units, the discounting convention
    freq, how = (None, "no client model assigned yet")
    if src:
        w = _Timelines(src["db_path"])
        try:
            freq, how = _frequency(w)
        finally:
            w.close()
    out["frequency"] = {"label": "Periods", "value": freq, "shown": freq, "how": how, "set": False,
                        "drives": "shown: each sheet's own periods are used"}
    units = next((f for f in reference(eid) if f["key"] == "currency_units"), None)
    out["units"] = {"label": "Currency and units", "value": units and units.get("value_text"),
                    "shown": units and units.get("value_text"), "set": False,
                    "how": f"the report's facts (page {units.get('page')})" if units else "not among the report's facts",
                    "drives": "shown: each figure keeps its own units"}
    conv, how = _conventions(eid)
    out["discounting"] = {"label": "Discounting", "value": conv, "shown": conv, "how": how, "set": False,
                          "drives": "shown: read back from each DCF's factors"}
    return {"fields": out, "settable": ["fy_end_month", "horizon"]}


def _conventions(eid: int) -> tuple[str | None, str]:
    """How the DCFs under the overlay's figures discount (timing, day count), read back from their factors."""
    import dcftrace
    rows = _q("SELECT overlay_json FROM engagements WHERE id=?", eid)
    summary = json.loads((rows[0]["overlay_json"] if rows else None) or "null") or {}
    if not summary.get("outputs"):
        return None, "build the Python overlay: it's read from the DCFs under the figures"
    seen = Counter()
    with closing(rodb.connect(summary["wiring"]["overlay"]["db_path"])) as db:
        for o in summary["outputs"][:8]:
            try:
                for c in dcftrace.cores(dcftrace.trace(db, o["cell"])):
                    m = c.get("method")
                    if m:
                        seen[f"{m['timing']} of period, {m['day_count']}"] += 1
            except ValueError:
                continue
    if not seen:
        return None, "no DCF under the figures whose factors could be read back"
    k, n = seen.most_common(1)[0]
    return k, f"read back from the factors of {sum(seen.values())} discounting(s) under the figures" + (
        f" (others: {', '.join(x for x in seen if x != k)})" if len(seen) > 1 else "")


def overlay_value_trace(eid: int, start: str | None = None) -> dict:
    """How a report figure is built in the overlay, from the cell it was matched to down to the discounting
    (dcftrace.py), with the Python overlay's value for every cell on the way."""
    import dcftrace
    import overlay as ovmod
    import rodb
    sess, summary = overlay_session(eid)
    starts = ovmod.trace_starts(summary)
    if start and start not in {x["cell"] for x in starts}:
        starts.append({"cell": start, "label": None, "report": None, "key": None, "ties": False, "value": None})
    if not starts:
        return {"starts": [], "selected": None}
    pick = next((x for x in starts if x["cell"] == start), starts[0])
    db = rodb.connect(summary["wiring"]["overlay"]["db_path"])
    tree = dcftrace.trace(db, pick["cell"])
    nodes, sheets = [], set(summary["sheets"])

    def walk(n):
        if not n.get("again"):
            nodes.append(n)
        for c in n.get("children", []):
            walk(c)
    walk(tree)
    mine = [n for n in nodes if n["cell"].rsplit("!", 1)[0].strip("'") in sheets]
    keys = [ovmod.parse_a1(n["cell"]) for n in mine]

    def run():
        sess.configure("workbook")
        return sess.values(keys)
    for n, v in zip(mine, ovmod.deep(run)):  # numbers only: a date cell is text in model.db and a serial in Python
        n["python"] = v if isinstance(v, float) and isinstance(n.get("value"), (int, float)) else None
    return {"starts": starts, "selected": pick["cell"], "start": pick, "tree": tree, "cores": dcftrace.cores(tree),
            "text": dcftrace.text(tree)}


def overlay_module(eid: int) -> Path:
    _, summary = overlay_session(eid)
    return Path(summary["module"])


# ---- the overlay doctor (doctor.py) ------------------------------------------------------------------------

def _doctor_file(eid: int) -> Path:
    return OUT / "overlays" / f"e{eid}" / "doctor.json"


def _holds_file(eid: int) -> Path:
    return OUT / "overlays" / f"e{eid}" / "holds.json"


def _load_holds(eid: int, sess) -> None:
    """The cells a person chose to hold at Excel's saved value, while each is still the same cell: the same
    formula and the same saved value as when it was held (a rebuilt workbook that changed it drops the hold)."""
    import overlay as ovmod
    f = _holds_file(eid)
    held = json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
    keep = {}
    with rodb.connect(sess.ov.path) as db:
        for cell, h in held.items():
            k = ovmod.parse_a1(cell)
            row = db.execute("SELECT formula FROM cells WHERE sheet=? AND row=? AND col=?", k).fetchone()
            if row and row[0] == h.get("formula") and ovmod.same(sess.ov.value(*k), h.get("value")):
                keep[k] = h["value"]
    sess.holds = keep
    sess.configure("workbook")


def _rowpicks_file(eid: int) -> Path:
    return OUT / "overlays" / f"e{eid}" / "rowpicks.json"


def _load_rowpicks(eid: int, sess) -> int:
    """A person's choices of this year's row for last year's rows ({"CF!r11": "CF!r15"}), and the agents', for the
    roll-forward. The agents' picks made by another version of the row finder are left out (it may now find those
    rows itself, and better): returns how many, so the agents can look again."""
    if not sess.rowmap:
        return 0
    import rowfind
    stale = 0
    for a, b in _read_rowpicks(eid).items():
        try:  # "[1]Sheet!r9" (a row of the linked client model) is Sheet row 9, as when it was picked
            s, r = _row_ref(a)
            to, by = (b.get("to"), b.get("by", "you")) if isinstance(b, dict) else (b, "you")
            if by == "agent" and b.get("v") != rowfind.VERSION:
                stale += 1
                continue
            sess.rowmap.pick(s, r, rowfind.STAND_IN if to == "-" else _row_ref(to) if to else None, by)
        except ValueError:
            continue
    return stale


def _read_rowpicks(eid: int) -> dict:
    """{"Sheet!r9": "Sheet!r12" | "-" | {"to", "by": "you" | "agent", "why", "checked_by"}}: a person's picks were
    saved as plain strings before the agents made picks too."""
    f = _rowpicks_file(eid)
    try:
        return json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
    except (OSError, ValueError):
        return {}


def _write_rowpicks(eid: int, picks: dict) -> None:
    f = _rowpicks_file(eid)
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps(picks, indent=1), encoding="utf-8")


def _row_ref(text: str) -> tuple[str, int]:
    m = re.match(r"^(?:\[\d+\])?(.+)!r(\d+)$", text or "")  # [1]Sheet!r9: a row of the linked client model
    if not m:
        raise ValueError(f"not a row: {text!r} (Sheet!rN)")
    return m[1], int(m[2])


def _num(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def row_context(sess, s: str, r: int, origins=()) -> dict:
    """What one of last year's rows is, for finding it this year: the heading it sits under and the rows beside it,
    what its formula adds up or reads (a "Total" says what it totals), last year's figures for its first periods,
    which overlay rows read it and whether it feeds the discounted cash flows under the value (origins: those rows),
    and this year's candidates with their own figures for the same periods."""
    import dcf
    from xlruntime import to_date
    prior = sess.prior or sess.ov
    labels, vals, tl = prior.labels(), prior.sheet(s), prior.timeline(s)
    cols = sorted(tl, key=tl.get)
    figures = lambda rr: [vals.get((rr, c)) for c in cols] if cols else [v for (r2, _c), v in vals.items() if r2 == rr]
    # the timeline's own rows (period dates, financial-year labels) are neither headings nor line items
    dated = lambda rr: cols and sum(1 for c in cols if _num(vals.get((rr, c))) and round(vals[(rr, c)]) == round(tl[c])) > len(cols) // 2
    has_figures = lambda rr: any(_num(v) for v in figures(rr)) and not dated(rr)
    empty = lambda rr: not any(v not in (None, "") for v in figures(rr))
    words = lambda rr: labels.get((s, rr)) or next((v for (r2, c), v in sorted(vals.items()) if r2 == rr and c not in tl
                                                    and isinstance(v, str) and v.strip()), None)
    heading = next((words(rr) for rr in range(r - 1, max(0, r - 80), -1)
                    if words(rr) and empty(rr)), None)  # a label (or a title) with nothing across the periods
    near = lambda rng: next(({"row": f"{s}!r{rr}", "label": labels[(s, rr)]} for rr in rng
                             if labels.get((s, rr)) and has_figures(rr)), None)
    out = {"heading": heading, "above": near(range(r - 1, max(0, r - 12), -1)), "below": near(range(r + 1, r + 12))}
    # its formula, where it has one: what it adds up, or what it reads
    got = next(((c, f) for c in (cols or sorted({c for (r2, c) in vals if r2 == r}))
                for (f,) in [prior.db.execute("SELECT formula FROM cells WHERE sheet=? AND row=? AND col=?", (s, r, c))
                             .fetchone() or (None,)] if f), None)
    if got:
        c0, f = got
        refs = [dcf._ref(m[0], s) for m in dcf._FREF.finditer(re.sub(r'"[^"]*"', "", f))]
        refs = [x for x in refs if x and not x[0].startswith("[")]
        rows = []
        for sh, r1, _c1, r2, _c2 in refs:
            for rr in range(r1, min(r2, r1 + 60) + 1):
                if (sh, rr) != (s, r) and (sh, rr) not in [(x["sheet"], x["r"]) for x in rows]:
                    rows.append({"sheet": sh, "r": rr, "row": f"{sh}!r{rr}", "label": prior.labels().get((sh, rr), "")})
        total = bool(re.fullmatch(r"=?\s*SUM\(\s*\$?[A-Z]{1,3}\$?\d+\s*:\s*\$?[A-Z]{1,3}\$?\d+\s*\)\s*", f, re.I))
        out["formula"] = {"text": f if f.startswith("=") else "=" + f, "cell": f"{s}!{dcf._addr(c0, r)}",
                          "adds_up": total, "reads": [{k: x[k] for k in ("row", "label")} for x in rows if x["label"]][:12],
                          "n": len([x for x in rows if x["label"]])}
    # last year's figures for its first periods with one
    out["values"] = [{"period": to_date(tl[c]).isoformat(), "value": vals.get((r, c))} for c in cols
                     if _num(vals.get((r, c)))][:5]
    # the overlay rows that read it
    sheets = list(sess.sheets or [])
    rx = re.compile(rf"(?:\[\d+\])?'?{re.escape(s)}'?!\$?[A-Z]{{1,3}}\$?{r}(?!\d)")
    readers = {}
    if sheets:
        for sh, rr, f in sess.ov.db.execute(
                f"SELECT sheet, row, formula FROM cells WHERE formula LIKE ? AND sheet IN ({','.join('?' * len(sheets))})",
                (f"%{s}%", *sheets)):
            if rx.search(f or "") and (sh, rr) != (s, r):
                readers[(sh, rr)] = readers.get((sh, rr), 0) + 1
    ov_labels = sess.ov.labels()
    out["read_by"] = [{"row": f"{sh}!r{rr}", "label": ov_labels.get((sh, rr), ""), "cells": n}
                      for (sh, rr), n in sorted(readers.items(), key=lambda kv: -kv[1])[:5]]
    out["feeds_dcf"] = f"{s}!r{r}" in set(origins or [])
    return out


def _this_year_figures(sess, row: str, periods: list[str]) -> list:
    """This year's figures of a row of this year's model (Sheet!rN) at these period ends (ISO), None where it has none."""
    from xlruntime import serial
    from datetime import date as _d
    s2, r2 = _row_ref(row)
    tl = sess.current.timeline(s2)
    by = {round(v): c for c, v in tl.items()}
    return [sess.current.value(s2, r2, by[round(serial(_d.fromisoformat(p)))]) if round(serial(_d.fromisoformat(p))) in by
            else None for p in periods]


def row_found(sess, s: str, r: int, origins=()) -> dict:
    """What this year's model has for one of last year's rows: the row found, how, the evidence, the
    alternatives, and a few periods of both years' values side by side."""
    ex = sess.rowmap.explain(s, r)
    prior = sess.prior or sess.ov
    out = {"row": f"{s}!r{r}", "label": prior.labels().get((s, r), ""), "found": None, "how": ex["how"],
           "evidence": [f"{n}: {t}" for n, t in ex["evidence"]], "confidence": ex["confidence"],
           "alternatives": ex["alternatives"], "picked": (s, r) in sess.rowmap.picks, "copies": ex.get("copies") or sess.rowmap.copies(s, r),
           "stand_in": bool(ex.get("stand_in")), "confident": sess.rowmap.confident(s, r)}
    if ex["found"]:
        s2, r2 = ex["found"]
        out.update(found=f"{s2}!r{r2}", found_label=sess.current.labels().get((s2, r2), ""))
        tl_p, tl_c = prior.timeline(s), sess.current.timeline(s2)
        both = sorted(set(tl_p.values()) & set(tl_c.values()))[:6]
        inv_p, inv_c = {v: c for c, v in tl_p.items()}, {v: c for c, v in tl_c.items()}
        from xlruntime import to_date
        out["side_by_side"] = [{"period": to_date(w).isoformat(), "last_year": prior.value(s, r, inv_p[w]),
                                "this_year": sess.current.value(s2, r2, inv_c[w])} for w in both]
    out["context"] = row_context(sess, s, r, origins)
    periods = [x["period"] for x in out["context"]["values"]]
    out["candidates"] = [{"row": x["row"], "label": x["label"], "why": x.get("why") or "; ".join(x.get("evidence") or []),
                          "figures": _this_year_figures(sess, x["row"], periods)}
                         for x in ([{"row": out["found"], "label": out.get("found_label"), "why": "what the finder found"}]
                                   if out["found"] else []) + list(out["alternatives"] or [])][:4]
    return out


def row_pick(eid: int, prior_row: str, current_row: str | None) -> dict:
    """A person's choice of this year's row for one of last year's (None to go back to what was found)."""
    import overlay as ovmod
    sess, summary = overlay_session(eid)
    if not sess.rowmap:
        raise ValueError("assign this year's client model (Roles) and rebuild in Python first")
    import rowfind
    s, r = _row_ref(prior_row)
    keep = current_row == "-"  # keep last year's values for the row, on purpose
    to = rowfind.STAND_IN if keep else _row_ref(current_row) if current_row else None
    if to and not keep and not sess.current.db.execute("SELECT 1 FROM rows WHERE sheet=? AND row=?", to).fetchone():
        raise ValueError(f"{current_row} isn't a line item in this year's model")
    picks = _read_rowpicks(eid)
    if to:
        picks[prior_row] = {"to": current_row, "by": "you"}
    else:
        picks.pop(prior_row, None)
    _write_rowpicks(eid, picks)
    ovmod.deep(sess.rowmap.pick, s, r, to)
    return ovmod.deep(row_found, sess, s, r)


def row_info(eid: int, prior_row: str) -> dict:
    """What this year's model has for one of last year's rows (Sheet!rN), with the alternatives: for picking it."""
    import overlay as ovmod
    sess, _ = overlay_session(eid)
    if not sess.rowmap:
        raise ValueError("assign this year's client model (Roles) and rebuild in Python first")
    got = _q("SELECT json_extract(result_json, '$.figures.gaps.dcf_origins') AS o FROM engagements WHERE id=?", eid)
    origins = json.loads((got[0]["o"] if got else None) or "[]")  # the rows the discounted cash flows come from
    return ovmod.deep(row_found, sess, *_row_ref(prior_row), origins)


def _rowagent_file(eid: int) -> Path:
    return OUT / "overlays" / f"e{eid}" / "rowagent.json"


def _rows_job(eid: int, cells: list[str]) -> None:
    """The row agents on the rows this year's value of these cells (the equity value's) is waiting on."""
    import rowagent
    import rowfind
    step = lambda msg: _set("engagements", eid, rows_status="running", rows_step=msg)
    step("Loading the Python overlay")
    sess, summary = overlay_session(eid)
    _sync_roll(eid, sess, summary)
    reader, why = None, None
    try:  # luna proposes, sol checks: the engagement's model and its reviewer
        e = _q("SELECT model, reviewer_model FROM engagements WHERE id=?", eid)[0]
        reader = docingest.Reader(e["model"] or DEFAULT_MODEL, e["reviewer_model"] or DEFAULT_REVIEWER, _logger(eid))
    except Exception as ex:
        why = friendly(ex)
    res = rowagent.run(sess, summary, cells, step, reader)
    if reader is not None:
        res["models"] = {"proposes": reader.model, "checks": reader.reviewer_model}
    else:
        res["models_error"] = why
    # the agents' picks, beside a person's (a person's always win; the agents' from an earlier run are replaced)
    picks = {k: v for k, v in _read_rowpicks(eid).items() if not (isinstance(v, dict) and v.get("by") == "agent")}
    for d in res["decisions"]:
        if d.get("decision") and d["row"] not in picks:
            picks[d["row"]] = {"to": d["decision"], "by": "agent", "v": rowfind.VERSION, "why": d.get("why"),
                               "checked_by": "the numbers" if d.get("how") == "numbers" else d.get("review") or d.get("how")}
    _write_rowpicks(eid, picks)
    res["at"], res["v"] = time.time(), rowfind.VERSION
    f = _rowagent_file(eid)
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps(res, default=str), encoding="utf-8")
    _set("engagements", eid, rows_status="done", rows_step="Done")


def rows_view(eid: int) -> dict:
    rows = _q("SELECT rows_status, rows_step, rows_error, rows_secs, overlay_started_at FROM engagements WHERE id=?", eid)
    if not rows:
        raise ValueError("no such engagement")
    e = rows[0]
    f = _rowagent_file(eid)
    res = json.loads(f.read_text(encoding="utf-8")) if f.exists() else None
    import rowfind
    built = max([(_role_wb(eid, k) or {}).get("processed_at") or 0 for k in ("prior_overlay", "prior_model", "current_model")]
                + [e["overlay_started_at"] or 0])
    # made before the Python overlay or a workbook was built again, or by another version of the row finder
    stale = bool(res and (res["at"] < built or res.get("v") != rowfind.VERSION))
    return {"status": e["rows_status"], "step": e["rows_step"], "error": e["rows_error"], "secs": e["rows_secs"],
            "result": res, "stale": stale}


def _doctor_job(eid: int) -> None:
    import doctor
    import overlay as ovmod
    step = lambda frac, msg: _set("engagements", eid, doctor_status="running", doctor_step=msg)
    step(0, "Loading the Python overlay")
    sess, summary = overlay_session(eid)
    evidence = ovmod.deep(doctor.examine, sess, summary, step)
    step(0.9, "Writing up the diagnosis")
    e = _q("SELECT model, reviewer_model FROM engagements WHERE id=?", eid)[0]
    res = {"at": time.time(), "evidence": evidence, "diagnosis": None, "diagnosis_error": None}
    try:
        reader = docingest.Reader(e["model"] or DEFAULT_MODEL, e["reviewer_model"] or DEFAULT_REVIEWER, _logger(eid))
        res["diagnosis"] = doctor.diagnose(reader, evidence)
        res["model"] = reader.reviewer_model
    except Exception as ex:  # the evidence stands on its own
        traceback.print_exc()
        res["diagnosis_error"] = friendly(ex)
    res["text"] = doctor.report_text(res)
    f = _doctor_file(eid)
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps(res, default=str), encoding="utf-8")
    _set("engagements", eid, doctor_status="done", doctor_step="Done")


def doctor_view(eid: int) -> dict:
    rows = _q("SELECT overlay_status, overlay_started_at, doctor_status, doctor_step, doctor_error, doctor_secs "
              "FROM engagements WHERE id=?", eid)
    if not rows:
        raise ValueError("no such engagement")
    e = rows[0]
    f = _doctor_file(eid)
    res = json.loads(f.read_text(encoding="utf-8")) if f.exists() else None
    held = json.loads(_holds_file(eid).read_text(encoding="utf-8")) if _holds_file(eid).exists() else {}
    return {"status": e["doctor_status"], "step": e["doctor_step"], "error": e["doctor_error"], "secs": e["doctor_secs"],
            "overlay_status": e["overlay_status"], "result": res,
            "stale": bool(res and e["overlay_started_at"] and res["at"] < e["overlay_started_at"]),
            "held": [{"cell": c, **h} for c, h in held.items()]}


def doctor_holds(eid: int, cells: list[str] | None, release: bool = False) -> dict:
    """Hold the doctor's safe cells at Excel's saved value on every feed (cells: all of them when None), or
    release every hold. Only cells the doctor found safe can be held."""
    import overlay as ovmod
    f = _holds_file(eid)
    held = json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
    if release:
        held = {}
    else:
        res = json.loads(_doctor_file(eid).read_text(encoding="utf-8")) if _doctor_file(eid).exists() else None
        safe = {h["cell"]: h for h in ((res or {}).get("evidence", {}).get("holds") or {}).get("safe") or []}
        if not safe:
            raise ValueError("the doctor found nothing that can be held: run it first")
        pick = [c for c in (cells or list(safe)) if c in safe]
        sess, summary = overlay_session(eid)
        with rodb.connect(sess.ov.path) as db:
            for c in pick:
                k = ovmod.parse_a1(c)
                row = db.execute("SELECT formula FROM cells WHERE sheet=? AND row=? AND col=?", k).fetchone()
                held[c] = {"value": safe[c]["value"], "formula": row[0] if row else None, "why": safe[c]["title"],
                           "at": time.time()}
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps(held, default=str), encoding="utf-8")
    if eid in _SESSIONS:
        import overlay as ovmod
        ovmod.deep(_load_holds, eid, _SESSIONS[eid][0])
    return doctor_view(eid)


def _process_doc(did: int) -> None:
    """Read a report's text layer (docingest.py with no model: pages and tables as the file's own text)."""
    d = _doc(did)
    prog = lambda frac, msg: _set("documents", did, pct=round(frac, 3), step=msg)
    _set("documents", did, status="processing", step="Reading the pages", pct=0, error=None, started_at=time.time())
    doc = docingest.process(d["source_path"], d["out_dir"], progress=prog, read=False)
    _save_doc(did, doc)
    d = _doc(did)
    _set("documents", did, status="done", step="Done", pct=1.0, processed_at=time.time(),
         doc_secs=round(time.time() - (d["started_at"] or time.time()), 1))


def _process_facts(did: int) -> None:
    """The key facts: the key tables read from their images first (visual.read_tables), then the facts extracted,
    checked, reviewed and looped on (keyfacts.py), then each one looked up on the image of where it sits
    (visual.confirm); the agents' agreement decides (auto_decide); what the loops taught goes to the lessons."""
    import visual
    d = _doc(did)
    eid = d["engagement_id"]
    e = _q("SELECT model, reviewer_model, arbiter_model FROM engagements WHERE id=?", eid)[0]
    doc = json.loads(d["doc_json"])
    t0 = time.time()
    step = lambda msg: _set("documents", did, facts_step=msg)
    _set("documents", did, facts_status="running", facts_step="Reading the key tables from their images", facts_error=None,
         facts_started_at=t0)
    arbiter = e["arbiter_model"] or DEFAULT_ARBITER
    note = {}
    try:
        reader = docingest.Reader(e["model"] or DEFAULT_MODEL, e["reviewer_model"] or DEFAULT_REVIEWER, _logger(eid), arbiter)
    except Exception as ex:  # no sign-in: the facts come from the text layer alone, and say so
        reader, note["error"] = None, "the images weren't read: " + friendly(ex)
    if reader is not None:
        note["tables"] = visual.read_tables(doc, d["out_dir"], reader, lambda frac, msg: step(msg), arbiter)
        _save_doc(did, doc)
    res = keyfacts.run(doc["markdown"], e["model"] or DEFAULT_MODEL, e["reviewer_model"] or DEFAULT_REVIEWER, _logger(eid),
                       lambda frac, msg: step(msg), arbiter_model=arbiter)
    if reader is not None:
        step("Checking each fact on the report's images")
        try:
            note["facts"] = visual.confirm(doc, res["facts"], reader, d["out_dir"], d["source_path"], lambda frac, msg: step(msg))
        except Exception as ex:  # the facts stand as the text layer gave them, and say they weren't looked at
            traceback.print_exc()
            note["error"] = "the facts weren't checked on the images: " + friendly(ex)
    now = time.time()
    with _lock, _conn() as db:
        db.execute("DELETE FROM facts WHERE document_id=?", (did,))
        for f in res["facts"]:
            if f.get("waivers"):
                f.setdefault("agent", {"status": "agreed", "round": 0, "thread": []})["waivers"] = f["waivers"]
            db.execute(f"""INSERT INTO facts(engagement_id, document_id, n, {', '.join(FACT_FIELDS)}, origin, check_json,
                           review_json, status, updated_at, agent_json, visual_json)
                           VALUES ({', '.join('?' * (len(FACT_FIELDS) + 10))})""",
                       (eid, did, f["id"], *[f.get(k) for k in FACT_FIELDS], f["origin"], json.dumps(f["check"]),
                        json.dumps(f.get("review")), "pending", now, json.dumps(f.get("agent")),
                        json.dumps(f.get("visual"))))
    _set("documents", did, facts_notes=res.get("notes"), review_summary=res.get("review_summary"))
    import context
    _note_loop(did, visual={**note, "at": now}, terminal=context.terminal(doc["markdown"]))  # what it says of its TV
    if res.get("loop"):
        _note_loop(did, facts={**res["loop"]["summary"], **auto_decide(did), "at": now}, lessons_facts=None)
        step("Writing down what the loop taught")
        _learn(did, "facts", res["loop"]["episodes"])
    _set("documents", did, facts_status="done", facts_step="Done", facts_secs=round(time.time() - t0, 1))
    _touch(eid)


def document_image(did: int, rel: str) -> Path | None:
    """One of a report's images (a table's, or a page's), kept in its folder."""
    d = _doc(did)
    if not d or not re.fullmatch(r"(tables|pages)/[A-Za-z0-9_-]+\.png", rel or ""):
        return None
    p = Path(d["out_dir"]) / rel
    return p if p.is_file() else None


def rebuild_document(did: int) -> dict:
    """Read a report again from scratch; its key facts go and are extracted again from the new reading."""
    d = _doc(did)
    if not d:
        raise ValueError("no such document")
    if d["status"] in ("queued", "processing") or d["facts_status"] in ("queued", "running"):
        raise ValueError(f"{d['filename']} is already being worked on")
    with _lock, _conn() as db:
        db.execute("DELETE FROM facts WHERE document_id=?", (did,))
    _set("documents", did, facts_status=None, facts_step=None, facts_error=None, loop_json=None)
    retry_document(did)
    return document(did)


def rebuild_workbook(eid: int, fid: int) -> dict:
    """Process a workbook again from scratch. Its model.db is shared (other engagements, the Model Desk), so every
    live Python overlay reading it lets go of it first: the build deletes the file, which Windows won't do while
    it's open."""
    return library.rebuild(_release(eid, fid)["id"], _model(eid))


def retry_workbook(eid: int, fid: int) -> dict:
    """Process a failed workbook again. It may have failed because a live Python overlay held its model.db (Windows
    won't delete an open file), so those let go of it first, as for a rebuild."""
    return library.retry(_release(eid, fid)["id"])


def _release(eid: int, fid: int) -> dict:
    """Every live Python overlay reading the workbook's model.db lets go of it (its build deletes the file)."""
    w = next((w for w in workbooks(eid) if w["id"] == fid), None)
    if not w:
        raise ValueError("that workbook isn't in this engagement")
    _let_go(set(), {w["db_path"]})
    for key in [k for k in _SHEET_NAMES if k[0] == w["db_path"]]:
        _SHEET_NAMES.pop(key, None)
    return w


def _let_go(eids: set[int], paths: set[str]) -> None:
    """The live Python overlays of these engagements, and any other reading one of these model.db files, let go of
    their files; then what's left over is collected. Windows won't delete an open file, and a sqlite connection or a
    read-only openpyxl workbook no longer used keeps its file open until the garbage collector gets to it."""
    import gc
    import overlay as ovmod
    for other, (sess, _) in list(_SESSIONS.items()):
        if other in eids or sess.paths() & paths:
            try:
                ovmod.deep(sess.close)
            finally:
                _SESSIONS.pop(other, None)
    gc.collect()


def retry_document(did: int) -> None:
    """Read the report again (the orchestrator picks it up)."""
    _set("documents", did, status="queued", step="Waiting to start", pct=0, error=None)


