"""The orchestrator: an engagement's work from its four files to last year's value rolled forward, with no one
starting each step, and a run log so it doesn't go round in circles.

Plain code runs the stages in order; each runs when what it depends on is ready, and again only when that changes:
  files    the report's text read (workbooks are built by the library as they arrive)
  facts    the report's key facts: extracted, checked, reviewed, the loop, the agents' approvals (keyfacts.py)
  roles    who's who: the rules' suggestion and a second opinion; confirmed here when both agree and the checks give
           evidence for it, else gpt-sol decides between them on that evidence (checked in code), else a person
  rebuild  the overlay compiled to Python and checked against Excel (cells that differ: the doctor looks, and what it
           can safely hold is held, once); this year's valuation date weighed; the report's equity value found
           (gpt-sol picks among the cells that hold it where the pairing can't tell)
  rows     this year's rows found for last year's (the row agents)
  result   the bridge from the report's value to this year's, and the cash-flow chart (result.py)
  review   gpt-sol reads the run end to end and says what looks wrong
  map      the links between the files, for the Map's cards
The run log (table runlog) is the memory: every start, outcome, decision and escalation, with the inputs it was
made on. Code, not the prompt, keeps it from spinning: a stage that failed doesn't run again until its inputs change
(or a person says try again); gpt-sol decides an issue once per set of inputs, sees what was tried before, and its
choice is checked in code before anything acts on it. A person's action changes the inputs, so the stages it
affects run again by themselves.
"""
import hashlib
import json
import queue
import re
import threading
import time
import traceback
from pathlib import Path

import calllog
import workbench as wb

STAGES = ("files", "facts", "roles", "rebuild", "rows", "result", "review", "map")
LABEL = {"files": "Read the files", "facts": "Report key facts", "roles": "Roles", "rebuild": "Rebuild last year",
         "rows": "Roll forward", "result": "Value bridge", "review": "Review", "map": "Map"}
LANE = {"doc": "report", "facts": "report", "review": "report", "roles": "models", "rebuild": "models",
        "rows": "models", "result": "models", "map": "models"}
UPSTREAM = {"facts": "files", "roles": "facts", "rebuild": "roles", "rows": "rebuild", "result": "rows",
            "review": "result", "map": "rebuild"}
BUDGET = 1  # gpt-sol decisions per issue on one set of inputs: after that, a person
MOVING, SETTLED = ("queued", "running"), ("done", "attention")
TICK = 2.0  # seconds between looks when nothing pokes it

_lock = threading.RLock()
_active: dict[tuple, str] = {}  # (engagement, stage or ("doc", id)) -> "queued" / "running"
_queues = {lane: queue.Queue() for lane in set(LANE.values())}
_wake = threading.Event()
_started = False


# ---- the store: stages and the run log ---------------------------------------------------------------------------

def stage(eid: int, name: str) -> dict:
    rows = wb._q("SELECT * FROM stages WHERE engagement_id=? AND stage=?", eid, name)
    r = rows[0] if rows else {"status": None, "inputs": None, "note": None, "data_json": None, "started_at": None,
                              "finished_at": None}
    return {**r, "data": json.loads(r.get("data_json") or "null") or {}}


def _put(eid: int, name: str, **fields) -> None:
    if "data" in fields:
        fields["data_json"] = json.dumps(fields.pop("data"), default=str)
    with wb._lock, wb._conn() as db:
        db.execute("INSERT OR IGNORE INTO stages(engagement_id, stage) VALUES (?,?)", (eid, name))
        db.execute(f"UPDATE stages SET {', '.join(f'{k}=?' for k in fields)} WHERE engagement_id=? AND stage=?",
                   (*fields.values(), eid, name))


def log(eid: int, name: str, event: str, text: str, issue: str | None = None, inputs: str | None = None,
        data=None) -> None:
    """One line of the run log: start, done, attention, blocked, failed, decide, verify, escalate, person, note."""
    wb._exec("INSERT INTO runlog(engagement_id, at, stage, event, issue, inputs, text, data_json) VALUES (?,?,?,?,?,?,?,?)",
             eid, time.time(), name, event, issue, inputs, text, json.dumps(data, default=str) if data is not None else None)


def history(eid: int, name: str | None = None, issue: str | None = None, limit: int = 60) -> list[dict]:
    where, args = ["engagement_id=?"], [eid]
    if name:
        where.append("stage=?")
        args.append(name)
    if issue:
        where.append("issue=?")
        args.append(issue)
    rows = wb._q(f"SELECT * FROM runlog WHERE {' AND '.join(where)} ORDER BY id DESC LIMIT ?", *args, limit)
    for r in rows:
        r["data"] = json.loads(r.pop("data_json") or "null")
    return rows


def _h(obj) -> str:
    return hashlib.sha1(json.dumps(obj, sort_keys=True, default=str).encode()).hexdigest()[:12]


# ---- what each stage depends on ----------------------------------------------------------------------------------

def _critical(facts: list[dict]) -> list:
    """The facts the rebuild reads (valuation date, equity value, discount rate): value, texts and decision."""
    keep = ("valuation_date", "equity_value", "equity_value_ex", "equity_value_cum", "discount_rate", "currency_units")
    return sorted([f["key"], f.get("value_text"), f.get("low_text"), f.get("high_text"), f.get("status"),
                   (f.get("final") or {}).get("value_text")] for f in facts if f["key"] in keep)


def _file_state(path: Path) -> str | None:
    try:
        return path.read_text(encoding="utf-8") if path.exists() else None
    except OSError:
        return None


def equity_file(eid: int) -> Path:
    return wb.OUT / "overlays" / f"e{eid}" / "equity.json"


def equity_pick(eid: int) -> dict | None:
    """The cells picked for last year's equity value (by gpt-sol or a person), where the pairing couldn't tell."""
    t = _file_state(equity_file(eid))
    return json.loads(t) if t else None


def set_equity_pick(eid: int, pick: dict | None, by: str, why: str = "") -> None:
    f = equity_file(eid)
    if pick is None:
        f.unlink(missing_ok=True)
        return
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps({**pick, "by": by, "why": why, "at": time.time()}), encoding="utf-8")


def inputs(eid: int, name: str, snap: dict, holds: bool = True) -> str:
    """A fingerprint of what a stage reads: it runs again when this changes, and only then. holds=False: the
    rebuild's without the cells held at Excel's value (the doctor runs once per rebuild of the same files)."""
    rl = snap["roles"]
    roles = sorted([k, r["kind"], r["id"], sorted(r["sheets"] or []), r["confirmed"]] for k, r in rl.items())
    ids = {r["id"] for r in rl.values() if r["kind"] == "workbook"}
    built = sorted([w["id"], w.get("processed_at"), w.get("valuation_date"), w.get("identity_confirmed")]
                   for w in snap["workbooks"] if w["id"] in ids)
    ov_dir = wb.OUT / "overlays" / f"e{eid}"
    if name == "roles":
        return _h([wb._roles_key(eid, snap["workbooks"], snap["documents"]),
                   [(k, r["id"], r["sheets"]) for k, r in rl.items() if r["confirmed"] and r.get("by") == "you"]])
    rebuild = [roles, built, _critical(snap["facts"]), _file_state(ov_dir / "holds.json") if holds else None,
               (snap.get("profile") or {}).get("horizon"), equity_pick(eid)]
    if name in ("rebuild", "map"):
        return _h(rebuild)
    picks = {k: v for k, v in wb._read_rowpicks(eid).items() if not (isinstance(v, dict) and v.get("by") == "agent")}
    rows = [rebuild, wb.this_year_date(eid), picks]
    if name == "rows":
        return _h(rows)
    res = [rows, equity_pick(eid), (snap.get("profile") or {}).get("fy_end_month"),
           sorted([f["key"], f.get("value_text"), f.get("status")] for f in snap["facts"]),
           _file_state(wb.held_file(eid)),  # this year's figures a person set for held inputs
           _file_state(wb.rate_file(eid)),  # and this year's discount rate
           _file_state(wb.method_file(eid)),  # and the method this year's value is worked out by
           _file_state(wb.terms_file(eid)),  # and the new terms a person confirmed belong in it
           _file_state(wb.acks_file(eid)),  # and the checks a person acknowledged, with the reason
           _result_version()]  # and the result's rules: a result worked out by older rules is worked out again
    return _h(res) if name == "result" else _h(["review", result_digest(eid)])


def _result_version() -> int:
    import result
    return result.VERSION


REVIEWED = ("head", "where", "tie", "values", "bridges", "chart", "reconcile", "inputs", "assumptions")


def result_digest(eid: int) -> str | None:
    """What the saved result says (the parts the review reads), as a fingerprint: a rerun that works out the same
    result doesn't need reviewing again."""
    rows = wb._q("SELECT result_json FROM engagements WHERE id=?", eid)
    res = json.loads((rows[0]["result_json"] if rows else None) or "null")
    return _h({k: res.get(k) for k in REVIEWED}) if res else None


# ---- the engagement as it stands ---------------------------------------------------------------------------------

def _snapshot(eid: int) -> dict:
    return {"workbooks": wb.workbooks(eid), "documents": wb.documents(eid), "facts": wb.facts(eid),
            "roles": wb.roles(eid), "profile": wb._profile(eid)}


def _files_state(snap: dict) -> tuple[str, str]:
    docs, wbs = snap["documents"], snap["workbooks"]
    busy = [x["filename"] for x in docs + wbs if x["status"] in MOVING or x["status"] == "processing"]
    bad = [x["filename"] for x in docs + wbs if x["status"] == "error"]
    if busy:
        return "running", f"reading {', '.join(busy[:3])}"
    if bad:
        return "failed", f"couldn't read {', '.join(bad)}"
    if not docs or len(wbs) < 2:
        want = (["last year's report"] if not docs else []) + (
            ["the client models and last year's overlay"] if len(wbs) < 2 else [])
        return "waiting", "upload " + " and ".join(want)
    return "done", f"{len(docs)} report(s) and {len(wbs)} workbook(s) read"


def _report(snap: dict) -> dict | None:
    """Last year's report: the confirmed or suggested role, else the only report."""
    r = snap["roles"].get("prior_report")
    docs = [d for d in snap["documents"] if d["status"] == "done"]
    return next((d for d in docs if r and d["id"] == r["id"]), None) or (docs[0] if len(docs) == 1 else None)


def _facts_state(snap: dict) -> tuple[str, str, list[dict]]:
    d = _report(snap) or next((d for d in snap["documents"] if d["status"] == "done"), None)
    if not d:
        return "waiting", "waiting for the report to be read", []
    if d["facts_status"] in MOVING:
        return "running", d.get("facts_step") or "extracting the key facts", []
    if d["facts_status"] == "error":
        return "failed", d.get("facts_error") or "the key facts failed", []
    if d["facts_status"] != "done":
        return "queued", "waiting to start", []
    mine = [f for f in snap["facts"] if f["document_id"] == d["id"]]
    live = [f for f in mine if f["status"] != "rejected"]
    needs = []
    for f in mine:
        if f["status"] == "pending" and (f.get("agent") or {}).get("status") == "escalated":
            crit = f["key"] in ("valuation_date", "equity_value", "discount_rate")
            needs.append({"id": f"fact-{f['id']}", "stage": "facts", "severity": "check" if crit else "info",
                          "title": f"Check {f.get('label') or f['key']} ({f.get('value_text') or f.get('low_text') or ''})",
                          "detail": ((f.get("agent") or {}).get("open") or {}).get("reason") or "the agents couldn't settle it",
                          "go": {"step": "report", "anchor": f"fact-{f['id']}"}})
    have = {f["key"] for f in live}
    for key, what in (("equity_value", "the equity value"), ("discount_rate", "the discount rate"),
                      ("valuation_date", "the valuation date")):
        if key not in have and not (key == "equity_value" and have & {"equity_value_ex", "equity_value_cum"}):
            needs.append({"id": f"missing-{key}", "stage": "facts", "severity": "check",
                          "title": f"The report's {what} wasn't found", "go": {"step": "report"},
                          "detail": "add it on The report page (page and quote), or check the report is last year's final"})
    n_ok = sum(f["status"] == "approved" for f in mine)
    return ("attention" if any(n["severity"] != "info" for n in needs) else "done"), \
        f"{n_ok} of {len(mine)} facts approved" + (f", {len(needs)} to look at" if needs else ""), needs


# ---- the loop ----------------------------------------------------------------------------------------------------

def busy_with(eid: int) -> list[str]:
    """What of an engagement is queued or running now (its stages, its reports being read)."""
    with _lock:
        jobs = [job for (e, job), _ in _active.items() if e == eid]
    return [LABEL.get(j, j) if isinstance(j, str) else f"report {j[1]}" for j in jobs]


def poke() -> None:
    _wake.set()


def _submit(eid: int, job, key: str | None = None) -> None:
    lane = LANE[job if isinstance(job, str) else job[0]]
    with _lock:
        if (eid, job) in _active:
            return
        _active[(eid, job)] = "queued"
    if isinstance(job, str):
        _put(eid, job, status="queued", note="waiting to start")
    _queues[lane].put((eid, job, key))


def tick(eid: int) -> None:
    """Look at one engagement and start whatever is ready. Starts nothing that's already moving, and nothing whose
    inputs are the same as when it last finished (or failed)."""
    if not wb._q("SELECT 1 FROM engagements WHERE id=?", eid):  # deleted since the loop listed it
        return
    snap = _snapshot(eid)
    for d in snap["documents"]:  # reports waiting to be read
        if d["status"] == "queued":
            _submit(eid, ("doc", d["id"]))
    st, note = _files_state(snap)
    _put(eid, "files", status=st, note=note)
    ready = {"files": st in SETTLED}
    # the key facts of every report read (usually the one)
    for d in snap["documents"]:
        if d["status"] == "done" and not d["facts_status"] and (eid, ("facts", d["id"])) not in _active:
            wb._set("documents", d["id"], facts_status="queued", facts_step="Waiting to start")
            _submit(eid, ("facts", d["id"]))
    st, note, needs = _facts_state(snap)
    if not ready["files"] and st == "waiting":
        note = "waiting for the files"
    said = {h["issue"] for h in history(eid, "facts", limit=200) if h["event"] == "escalate"}
    for n in needs:  # into the run log once: a fact the agents couldn't settle goes to a person
        if n["severity"] != "info" and n["id"] not in said:
            log(eid, "facts", "escalate", n["title"] + (f": {n['detail']}" if n.get("detail") else ""), issue=n["id"])
    _put(eid, "facts", status=st, note=note, data={"needs": needs})
    # facts that failed don't hold the roles up (structure alone can place the files); the rest wait for them
    ready["facts"] = ready["files"] and (st in SETTLED or st == "failed")
    has_current = bool(snap["roles"].get("current_model"))
    for name in ("roles", "rebuild", "rows", "result", "review", "map"):
        up = UPSTREAM[name]
        if not ready.get(up):
            rec = stage(eid, name)
            if rec["status"] not in MOVING and (eid, name) not in _active:
                data = rec["data"]
                if rec["status"] in ("done", "attention", "blocked", "failed") and rec["inputs"]:
                    # its outcome stands if its inputs come out the same once the stage before settles again
                    data = {**data, "prev": {"status": rec["status"], "note": rec["note"]}}
                _put(eid, name, status="waiting", note=f"waiting for {LABEL[up].lower()}", data=data)
            ready[name] = False
            continue
        rec = stage(eid, name)
        if (eid, name) in _active:
            ready[name] = False
            continue
        if name == "rows" and not has_current:
            _put(eid, name, status="done", note="no client model for this year yet: last year only", inputs=None)
            ready[name] = True
            continue
        key = inputs(eid, name, snap)
        if rec["inputs"] != key:
            _submit(eid, name, key)
            ready[name] = False
            continue
        if rec["status"] == "waiting":  # the stage before ran again and changed nothing this stage reads
            prev = rec["data"].get("prev")
            if not prev:
                _submit(eid, name, key)
                ready[name] = False
                continue
            _put(eid, name, status=prev["status"], note=prev["note"],
                 data={k: v for k, v in rec["data"].items() if k != "prev"})
            log(eid, name, "note", f"{LABEL[name]}: unchanged by the rerun before it, so its outcome stands", inputs=key)
            rec["status"] = prev["status"]
        ready[name] = rec["status"] in SETTLED


def _loop() -> None:
    while True:
        _wake.wait(TICK)
        _wake.clear()
        for e in wb._q("SELECT id FROM engagements ORDER BY updated_at DESC"):
            try:
                tick(e["id"])
            except Exception:
                traceback.print_exc()


def _worker(lane: str) -> None:
    q = _queues[lane]
    while True:
        eid, job, key = q.get()
        with _lock:
            _active[(eid, job)] = "running"
        try:
            _run(eid, job, key)
        except Exception:
            traceback.print_exc()
        finally:
            with _lock:
                _active.pop((eid, job), None)
            q.task_done()
            poke()


def _run(eid: int, job, key: str | None) -> None:
    if isinstance(job, tuple):  # a report: read it, or its key facts
        kind, did = job
        name = "files" if kind == "doc" else "facts"
        d = wb._doc(did)
        with calllog.tag(engagement=eid, document=did, step="read the report" if kind == "doc" else "key facts"):
            log(eid, name, "start", f"{'reading' if kind == 'doc' else 'the key facts of'} {d['filename']}")
            try:
                (wb._process_doc if kind == "doc" else wb._process_facts)(did)
                log(eid, name, "done", f"{d['filename']}: " + ("read" if kind == "doc" else _facts_line(did)))
            except Exception as e:
                traceback.print_exc()
                msg = wb.friendly(e)
                if kind == "doc":
                    wb._set("documents", did, status="error", step="Failed", error=msg)
                else:
                    wb._set("documents", did, facts_status="error", facts_step="Failed", facts_error=msg)
                log(eid, name, "failed", f"{d['filename']}: {msg}")
        return
    t0 = time.time()
    _put(eid, job, status="running", note="starting", started_at=t0)
    log(eid, job, "start", f"{LABEL[job]}: started", inputs=key)
    try:
        with calllog.tag(engagement=eid, step=f"orchestrator: {job}"):
            status, note, data = STAGE_JOBS[job](eid, key)
    except Exception as e:
        traceback.print_exc()
        status, note, data = "failed", wb.friendly(e), {"needs": [{
            "id": f"failed-{job}", "stage": job, "severity": "block", "title": f"{LABEL[job]} failed",
            "detail": wb.friendly(e), "retry": job}]}
    _put(eid, job, status=status, note=note, data=data, inputs=key, finished_at=time.time())
    log(eid, job, {"done": "done", "attention": "attention", "blocked": "blocked", "failed": "failed"}.get(status, "note"),
        f"{LABEL[job]}: {note}", inputs=key)


def _facts_line(did: int) -> str:
    loop = (wb.document(did) or {}).get("loop") or {}
    s = loop.get("facts") or {}
    return (f"{s.get('agreed', 0)} agreed, {s.get('withdrawn', 0)} withdrawn, {s.get('escalated', 0)} for a person"
            if s else "extracted")


def start() -> None:
    """Pick up what a restart interrupted, then start the loop and the lanes."""
    global _started
    if _started:
        return
    _started = True
    wb._conn().close()
    for r in wb._q("SELECT id FROM documents WHERE status='processing'"):
        wb._set("documents", r["id"], status="queued", step="Waiting to start")
    for r in wb._q("SELECT id FROM documents WHERE facts_status IN ('queued','running')"):
        wb._set("documents", r["id"], facts_status=None, facts_step=None)
    with wb._lock, wb._conn() as db:  # interrupted stages run again: their inputs are forgotten
        db.execute("UPDATE stages SET status='waiting', inputs=NULL, note='interrupted by a restart' "
                   "WHERE status IN ('queued','running')")
    for lane in _queues:
        threading.Thread(target=_worker, args=(lane,), daemon=True, name=f"orchestrator-{lane}").start()
    threading.Thread(target=_loop, daemon=True, name="orchestrator").start()
    poke()


# ---- gpt-sol at the decision points ------------------------------------------------------------------------------

def _decisions(eid: int, name: str, issue: str, key: str) -> int:
    return len(wb._q("SELECT 1 FROM runlog WHERE engagement_id=? AND stage=? AND event='decide' AND issue=? AND inputs=?",
                     eid, name, issue, key))


def ask(eid: int, name: str, issue: str, key: str, prompt: str, schema: dict) -> dict | None:
    """One decision by gpt-sol (the engagement's reviewer model) on an issue, with what was tried before on it (the run
    log) in the prompt. None when the budget for this issue on these inputs is spent: a person decides."""
    if _decisions(eid, name, issue, key) >= BUDGET:
        log(eid, name, "escalate", "decided once already on these inputs: a person decides", issue=issue, inputs=key)
        return None
    from llm import client, create
    model = wb._q("SELECT reviewer_model FROM engagements WHERE id=?", eid)[0]["reviewer_model"] or wb.DEFAULT_REVIEWER
    before = [{"when": time.strftime("%Y-%m-%d %H:%M", time.localtime(h["at"])), "event": h["event"], "text": h["text"],
               "decision": (h.get("data") or {}).get("choice")} for h in history(eid, name, issue, 12)]
    full = (prompt + "\n\nWhat was tried before on this (the run log, newest first; don't repeat what failed):\n"
            + (json.dumps(before, indent=1) if before else "nothing yet"))
    r = create(client(interactive=False), model, input=full, text={"format": schema}, max_output_tokens=3000,
               purpose=f"orchestrator-{name}")
    if r.usage:
        wb._logger(eid)(model, r.usage, f"orchestrator-{name}")
    out = json.loads(r.output_text)
    log(eid, name, "decide", f"{model}: {out.get('choice')} — {out.get('reason', '')}", issue=issue, inputs=key,
        data={**out, "model": model})
    return out


_S = {"type": "string"}


def _schema(name: str, choices: list[str], extra: dict | None = None) -> dict:
    props = {"choice": {"type": "string", "enum": choices}, "reason": _S, "question": _S, **(extra or {})}
    return {"type": "json_schema", "name": name, "strict": True, "schema": {
        "type": "object", "additionalProperties": False, "required": list(props), "properties": props}}


# ---- roles -------------------------------------------------------------------------------------------------------

ROLES_PROMPT = """You are the orchestrator of a recurring infrastructure valuation. Before last year's valuation can be
rebuilt, each file must be given its part: prior_report (last year's report), prior_model (the client's model behind
it), prior_overlay (our valuation workings: a workbook, or sheets inside a copy of the client model) and
current_model (this year's client model). Rules suggested an assignment and a second model gave its opinion; they
don't agree, or the checks don't give enough evidence. Decide on the evidence:
- "rules": the rules' assignment is right
- "second_opinion": the second opinion's assignment is right
- "escalate": neither is safe; say in question what a person should look at
Only choose an assignment the evidence supports; when in doubt, escalate.

Checks on the rules' assignment:
{checks}

The rules' assignment, with their reasons:
{rules}

The second opinion:
{second}

The evidence both saw:
{evidence}"""


def _assignment(res: dict) -> dict:
    return {k: {"kind": r["kind"], "id": r["id"], "sheets": r["sheets"]} for k, r in res["roles"].items()}


def _from_second(eid: int, so: dict, res: dict) -> dict | None:
    """The second opinion's assignment by file name, as ids; None if a name doesn't match a file."""
    docs = {d["filename"]: d["id"] for d in wb.documents(eid)}
    wbs = {w["filename"]: w for w in wb.workbooks(eid)}
    out = {}
    for role in ("prior_report", "prior_model", "prior_overlay", "current_model"):
        name = so.get(role) or ""
        if not name:
            continue
        if role == "prior_report":
            if name not in docs:
                return None
            out[role] = {"kind": "document", "id": docs[name], "sheets": None}
        else:
            w = wbs.get(name)
            if not w:
                return None
            sheets = None
            if role == "prior_overlay":
                sheets = [s for s in so.get("overlay_sheets") or [] if s in set(w["sheet_names"])] or None
                if so.get("overlay_sheets") and not sheets:
                    return None
            elif role == "prior_model" and (res["roles"].get("prior_model") or {}).get("id") == w["id"]:
                sheets = res["roles"]["prior_model"]["sheets"]
            out[role] = {"kind": "workbook", "id": w["id"], "sheets": sheets}
    return out


def verify_roles(eid: int, a: dict) -> list[str]:
    """What's wrong with an assignment, in code: every part given, the right kinds, this year's model not last
    year's, dates the right way round where the files say, the report's valuation date the overlay's. [] if nothing."""
    bad = []
    for role, kind in (("prior_report", "document"), ("prior_model", "workbook"), ("prior_overlay", "workbook"),
                       ("current_model", "workbook")):
        if not a.get(role):
            bad.append(f"no {role.replace('_', ' ')}")
        elif a[role]["kind"] != kind:
            bad.append(f"the {role.replace('_', ' ')} must be a {kind}")
    if bad:
        return bad
    w = {x["id"]: x for x in wb.workbooks(eid)}
    pm, cm, ovw = w.get(a["prior_model"]["id"]), w.get(a["current_model"]["id"]), w.get(a["prior_overlay"]["id"])
    if a["current_model"]["id"] in (a["prior_model"]["id"], a["prior_overlay"]["id"]):
        bad.append("this year's client model is also one of last year's files")
    if pm and cm and pm.get("valuation_date") and cm.get("valuation_date") and cm["valuation_date"] <= pm["valuation_date"]:
        bad.append(f"this year's client model's valuation date ({cm['valuation_date']}) isn't after last year's "
                   f"({pm['valuation_date']})")
    rep = next((f for f in wb.reference(eid) if f["key"] == "valuation_date" and f.get("value")), None)
    if rep and ovw and ovw.get("valuation_date"):
        v = int(rep["value"])
        iso = f"{v // 10000:04d}-{v // 100 % 100:02d}-{v % 100:02d}"
        if iso != ovw["valuation_date"] and not (pm and pm.get("valuation_date") == iso):
            bad.append(f"the report's valuation date ({iso}) is neither the overlay's ({ovw['valuation_date']}) nor "
                       "last year's client model's")
    if a["prior_overlay"]["id"] == a["prior_model"]["id"] and not a["prior_overlay"].get("sheets"):
        bad.append("the overlay is inside the client model but its sheets aren't named")
    elif a["prior_overlay"]["id"] == a["prior_model"]["id"] and ovw and ovw.get("sheet_names") and \
            set(ovw["sheet_names"]) <= set(a["prior_overlay"]["sheets"]):
        bad.append("the overlay is inside the client model but takes every sheet: none is left as the client's")
    return bad


def _roles_job(eid: int, key: str):
    rl = wb.roles(eid)
    if len(rl) == 4 and all(r["confirmed"] and r.get("by") == "you" for r in rl.values()):
        return "done", "confirmed by you", {}
    res = wb.suggest_roles(eid)
    a = _assignment(res)
    mine = {k: {"kind": r["kind"], "id": r["id"], "sheets": r["sheets"]} for k, r in rl.items()
            if r["confirmed"] and r.get("by") == "you"}
    filled = {}
    for k, m in mine.items():  # a file placed by hand, with no sheets named: the suggestion's, where it picked the same file
        if m["sheets"] is None and (a.get(k) or {}).get("id") == m["id"] and a[k]["sheets"]:
            m["sheets"] = a[k]["sheets"]
            filled[k] = m
    a.update(mine)  # a person's choices stand; the rest is the suggestion

    def settle(pick: dict, ev: dict) -> None:
        """The orchestrator confirms the roles a person didn't place; a person's stay theirs (with the sheets filled in)."""
        if filled:
            wb.confirm_roles(eid, filled, by="you")
        wb.confirm_roles(eid, {k: v for k, v in pick.items() if k not in mine}, by="orchestrator", evidence=ev)
    checks = res.get("checks") or []
    so = res.get("second_opinion") or {}
    yes, no = [c["text"] for c in checks if c["ok"] is True], [c["text"] for c in checks if c["ok"] is False]
    agree = not so.get("error") and not so.get("differs") and so.get("confidence") in ("high", "medium")
    evidence = lambda extra: {role: {"why": (res["roles"].get(role) or {}).get("why") or [], "checks": yes, **extra}
                              for role in a}
    if len(a) == 4 and not no and len(yes) >= 2 and agree and not verify_roles(eid, a):
        settle(a, evidence({"second_opinion": f"{so.get('model')} agrees ({so.get('confidence')} confidence)"}))
        return "done", f"confirmed on {len(yes)} checks and a second opinion that agrees", {"evidence": yes}
    why = ([f"{len(no)} check(s) against: {'; '.join(no)}"] if no else []) + \
        ([f"only {len(yes)} check(s) for"] if len(yes) < 2 else []) + \
        ([f"the second opinion differs on {', '.join(so.get('differs') or [])}"] if so.get("differs") else []) + \
        (["no second opinion: " + so["error"]] if so.get("error") else []) + \
        ([f"not every part was found ({', '.join(sorted(set(wb.rolesmod.ROLES) - set(a)))} missing)"] if len(a) < 4 else [])
    log(eid, "roles", "note", "not confirmed by the rules alone: " + "; ".join(why), issue="roles", inputs=key)
    names = {w["id"]: w["filename"] for w in wb.workbooks(eid)} | {d["id"]: d["filename"] for d in wb.documents(eid)}
    rules = {k: {"file": names.get(r["id"]), "sheets": r["sheets"], "why": (res["roles"].get(k) or {}).get("why")}
             for k, r in a.items()}
    docs = [d for d in wb.documents(eid) if d["status"] == "done"]
    ev = wb.rolesmod.evidence([{**d, "n_facts": 0} for d in docs], wb._wb_inputs(eid), wb.reference(eid), res)
    out = None
    if not so.get("error"):
        out = ask(eid, "roles", "roles", key, ROLES_PROMPT.format(
            checks=json.dumps(checks, indent=1), rules=json.dumps(rules, indent=1, default=str),
            second=json.dumps({k: so.get(k) for k in ("prior_report", "prior_model", "prior_overlay", "overlay_sheets",
                                                      "current_model", "confidence", "reasons")}, indent=1),
            evidence=json.dumps(ev, indent=1, default=str)[:40000]), _schema("roles_decision", ["rules", "second_opinion", "escalate"]))
    if out and out["choice"] in ("rules", "second_opinion"):
        pick = a if out["choice"] == "rules" else _from_second(eid, so, res)
        if pick is not None:
            pick = {**pick, **mine}  # the second opinion doesn't overrule a person either
        bad = verify_roles(eid, pick) if pick else ["the second opinion names files that aren't here"]
        if not bad:
            log(eid, "roles", "verify", "the checks in code pass on the chosen assignment", issue="roles", inputs=key)
            settle(pick, evidence({"decided_by": f"{out.get('reason')}"}))
            return "done", f"confirmed: gpt-sol chose the {'rules' if out['choice'] == 'rules' else 'second opinion'}'s " \
                           f"assignment ({out['reason']})", {"decision": out}
        log(eid, "roles", "verify", "the chosen assignment fails the checks in code: " + "; ".join(bad), issue="roles", inputs=key)
        why.append("gpt-sol's choice failed the checks: " + "; ".join(bad))
    question = (out or {}).get("question") or "Which file is which? The suggestion is filled in on Roles; confirm or change it."
    return "blocked", "a person confirms the roles", {"needs": [{
        "id": "roles", "stage": "roles", "severity": "block", "title": "Confirm the roles",
        "detail": question + " (" + "; ".join(why) + ")", "go": {"step": "workbench", "anchor": "rolesCard"}}]}


# ---- the rebuild -------------------------------------------------------------------------------------------------

TIE_PROMPT = """You are the orchestrator of a recurring infrastructure valuation, rebuilding last year's valuation in Python.
The report concludes an equity value as a range (low and high; the mid is their midpoint), and we need the overlay
cells holding the low and the high. Code found the cells below whose saved values equal the report's figures, but
couldn't pair a low with a high on its own. Pick the pair: the cells that are the equity value (cash flows to equity
at the cost of equity, or free cash flows at a WACC less net debt; after any distribution), on the report's basis (ex-distribution unless the model only gives cum-distribution), low and
high of the same range. Or escalate if none is right.

The report's equity value: {head}

Cells (cell, its line item's label, saved value, which report figure it equals, basis, units scale):
{cands}"""


def _tie_schema() -> dict:
    return _schema("equity_cells", ["pick", "escalate"], {"low_cell": _S, "high_cell": _S})


def _report_md(eid: int) -> str:
    d = _report(_snapshot(eid))
    return (json.loads((wb._doc(d["id"]) or {}).get("doc_json") or "{}").get("markdown") or "") if d else ""


def _rebuild_job(eid: int, key: str):
    import keyfacts
    import result
    needs, notes = [], []
    wb._overlay(eid)
    sess, summary = wb.overlay_session(eid)
    val = summary["validation"]
    off = val["cells"] - val["matched"] - val.get("text", 0)
    if off > 0:
        # The doctor's gate is keyed on the rebuild's inputs WITHOUT the held cells, not on the stage's own key: holding
        # cells changes the stage's key (so the rebuild runs again), and gating on it would run the doctor again after
        # every hold, and hold again, for ever. Once per rebuild of the same files, roles and facts.
        core = inputs(eid, "rebuild", _snapshot(eid), holds=False)
        if not [h for h in history(eid, "rebuild", "validation") if h["inputs"] == core and h["event"] == "note"]:
            log(eid, "rebuild", "note", f"{off} of {val['cells']:,} cells differ from Excel: the doctor looks at them",
                issue="validation", inputs=core)
            wb._doctor_job(eid)
            doc = wb.doctor_view(eid)
            safe = (((doc.get("result") or {}).get("evidence") or {}).get("holds") or {}).get("safe") or []
            new = [h for h in safe if h["cell"] not in {x["cell"] for x in doc.get("held") or []}]
            if new:
                wb.doctor_holds(eid, None)
                log(eid, "rebuild", "note", f"the doctor found {len(new)} cell(s) it can safely hold at Excel's value: "
                    "held, and the rebuild runs again", issue="validation", inputs=core)
                return "done", "holding cells the doctor found safe; rebuilding again", {"rerun": True}
        needs.append({"id": "validation", "stage": "rebuild", "severity": "info",
                      "title": f"{off} of {val['cells']:,} overlay cells differ from Excel in Python",
                      "detail": "shown on Rebuild; they matter only if they're under the equity value",
                      "go": {"step": "rebuild", "anchor": "mismatches"}})
    if wb._agents_check_dates(eid, wb.workbooks(eid)):
        log(eid, "rebuild", "note", "this year's valuation date confirmed on the files' evidence", issue="date", inputs=key)
    cur = wb._role_wb(eid, "current_model")
    if cur and not cur.get("identity_confirmed") and not wb.this_year_date(eid):
        w = next((x for x in wb.workbooks(eid) if x["id"] == cur["id"]), {})
        needs.append({"id": "date", "stage": "rebuild", "severity": "check",
                      "title": f"Confirm this year's valuation date ({cur.get('valuation_date') or 'not found'})",
                      "detail": "; ".join(((w.get("identity_check") or {}).get("disagree") or [])) or
                                "the files don't say it clearly enough; the roll-forward uses it",
                      "go": {"step": "workbench", "anchor": f"wb-{cur['id']}"}})
    facts = wb.reference(eid)
    head = keyfacts.conclusion(facts, _report_md(eid))
    where = None
    if not head:
        needs.append({"id": "equity", "stage": "rebuild", "severity": "block", "title": "No equity value from the report",
                      "detail": "the bridge starts from it: approve or add it on The report", "go": {"step": "report"}})
    else:
        pick = equity_pick(eid)
        where = pick or result.locate(summary, head)
        if where:
            notes.append(f"the equity value is at {where['low']} / {where['high']} ({where['how']})")
        else:
            cands = result.candidates(summary, head)
            out = ask(eid, "rebuild", "equity", key, TIE_PROMPT.format(
                head=json.dumps({k: head[k] for k in ("basis", "low", "high", "mid", "texts", "units")}, default=str),
                cands=json.dumps([{k: c[k] for k in ("cell", "label", "value", "part", "basis", "scale")} for c in cands],
                                 indent=1, default=str)), _tie_schema()) if cands else None
            if out and out["choice"] == "pick":
                got = {c["cell"]: c for c in cands}
                lo, hi = got.get(out["low_cell"]), got.get(out["high_cell"])
                ok = lo and hi and lo["part"] in ("low", "value") and hi["part"] in ("high", "value") \
                    and (lo["scale"], lo["sign"]) == (hi["scale"], hi["sign"])
                if ok:
                    where = {"low": lo["cell"], "high": hi["cell"], "mid": None, "scale": lo["scale"], "sign": lo["sign"],
                             "basis": lo["basis"], "label": lo["label"], "how": f"picked by gpt-sol: {out['reason']}"}
                    set_equity_pick(eid, where, "orchestrator", out["reason"])
                    log(eid, "rebuild", "verify", f"the pick holds the report's figures: {lo['cell']} / {hi['cell']}",
                        issue="equity", inputs=key)
                    notes.append(f"the equity value is at {where['low']} / {where['high']} (picked by gpt-sol)")
                else:
                    log(eid, "rebuild", "verify", "the pick isn't a low and a high of the report's range", issue="equity",
                        inputs=key)
        if not where:
            needs.append({"id": "equity", "stage": "rebuild", "severity": "block",
                          "title": "Where is last year's equity value in the overlay?",
                          "detail": (out or {}).get("question") or "no pair of cells holding the report's low and high was "
                                    "found: pick them on Rebuild", "go": {"step": "rebuild", "anchor": "equityPick"},
                          "candidates": cands})
    blocked = any(n["severity"] == "block" for n in needs)
    moved = (summary.get("feeds") or {}).get("prior")
    if moved and not moved.get("same"):
        needs.append({"id": "prior-feed", "stage": "rebuild", "severity": "info",
                      "title": "Last year's client model gives slightly different values than the overlay saved",
                      "detail": "the client file may not be the version the overlay was built on: the bridge shows the "
                                "difference as its own step", "go": {"step": "result"}})
    status = "blocked" if blocked else "attention" if any(n["severity"] == "check" for n in needs) else "done"
    return status, (f"{val['matched']:,} of {val['cells']:,} cells reproduced" + (f"; {notes[0]}" if notes else "")), \
        {"needs": needs, "where": where, "validation": {k: val[k] for k in ("cells", "matched")}}


# ---- rows, result, review, map -----------------------------------------------------------------------------------

def equity_cells(eid: int) -> list[str]:
    """The cells holding last year's equity value (the rebuild found them, or gpt-sol or a person picked them)."""
    w = equity_pick(eid) or stage(eid, "rebuild")["data"].get("where") or {}
    return list(dict.fromkeys(x for x in (w.get("low"), w.get("high"), w.get("mid")) if x))


def _rows_job(eid: int, key: str):
    cells = equity_cells(eid)
    if not cells:
        return "blocked", "no equity value cells to roll forward", {"needs": [{
            "id": "equity-rows", "stage": "rows", "severity": "block", "title": "Pick last year's equity value cells first",
            "detail": "the rows to roll forward are the ones under the equity value", "go": {"step": "rebuild", "anchor": "equityPick"}}]}
    wb._rows_job(eid, cells)
    res = (wb.rows_view(eid).get("result") or {})
    n = len([d for d in res.get("decisions") or [] if d.get("decision")])
    return "done", f"the row agents found {n} row(s) the finder wasn't sure of" if n else "this year's rows found", {}


def _time_text(x: dict, t: float) -> str:
    return (f"{x['cell']} grows {x['ratio'] - 1:+.1%} over {t:.2f} year(s), {x['implied']:.2%} a year against its rate "
            f"of {x['rate']:.2%}")


def _term_text(x: dict, how: str) -> str:
    return (f"{x['row']} ({x['label']}) {how} {x['under_label'] or x['under'] or x['last_year']}: {x['periods']} "
            f"period(s), {x['total']:,.1f} in total" + (f" (last year's; now at {x['now']}, outside the sum)"
                                                        if x.get("now") else " (last year's)" if how != "in" else ""))


def _acked(n: dict, key: str, ack: dict | None) -> dict:
    """A need a person can acknowledge with a reason (key: the figures it found). Acknowledged on these figures, it
    stays on the list as a note with the reason; on other figures it stands again."""
    n = {**n, "ack": key}
    if ack and ack.get("key") == key:
        when = time.strftime("%d %B %Y", time.localtime(ack.get("at") or time.time())).lstrip("0")
        n.update(severity="info", acked=ack, title="Acknowledged: " + n["title"],
                 detail=(n.get("detail") or "") + f" — acknowledged by {ack.get('by') or 'you'} on {when}"
                        + (f": {ack['reason']}" if ack.get("reason") else ""))
    return n


def _need_of(h: dict, go: dict, stage_: str = "result", **more) -> dict:
    """A check's finding (result.hold) as a need: blocks hold the value back until acknowledged."""
    return _acked({"id": h["id"], "stage": stage_, "severity": h["severity"], "title": h["title"],
                   "detail": h["detail"], "go": go, **more}, h["key"], h.get("acked"))


def _result_job(eid: int, key: str):
    import overlay as ovmod
    import result
    sess, summary = wb.overlay_session(eid)
    wb._sync_roll(eid, sess, summary)
    fy = wb.profile_view(eid)["fields"]["fy_end_month"]["value"] or 12
    with wb._fy_hint(eid):
        res = result.compute(sess, summary, wb.reference(eid), _report_md(eid), fy, equity_pick(eid))
    if res.get("stop"):
        return "blocked", res["why"], {"needs": [{"id": "equity", "stage": "result", "severity": "block",
                                                  "title": res["why"], "go": {"step": "rebuild", "anchor": "equityPick"}}]}
    res["inputs_key"] = key  # the inputs it was worked out on: shown as stale when they change
    wb._set("engagements", eid, result_json=json.dumps(res, default=str), updated_at=time.time())
    needs = []
    g = res["figures"].get("gaps")
    dc = (g or {}).get("date_cells") or {}
    for h in (g or {}).get("holds") or []:  # this year's cash flows against last year's (cashflows.py)
        anchor = ("bridgeCard" if h["id"].startswith(("cf-split", "cf-sign")) else "compareCard"
                  if h["id"] in ("basis", "interest", "interest-two") else "flowsCard")
        needs.append(_need_of(h, {"step": "result", "anchor": anchor}))
    if g and not g["reliable"]:
        if g.get("no_reads"):
            needs.append({"id": "no-reads", "stage": "result", "severity": "block",
                          "title": "This year's value is held back: the overlay reads nothing of this year's client model",
                          "detail": "which workbook is last year's client model, and which of the overlay's sheets are its "
                                    "own? With none of the client's sheets left to feed, this year's model can't reach the "
                                    "value (it would be last year's figures, rolled by date alone)",
                          "go": {"step": "workbench", "anchor": "rolesCard"}})
        if dc.get("off"):
            off = dc["off"]
            needs.append({"id": "dates-roll", "stage": "result", "severity": "block",
                          "title": "This year's value is held back: a discounting still reads another valuation date",
                          "detail": "; ".join(f"{x['cell']} reads {x['date']}" for x in off[:4]) +
                                    f", not {dc.get('to')}: its date isn't one the roll-forward moves",
                          "go": {"step": "workbench", "anchor": "datesCard"}})
        if g.get("time_off"):
            t = (res["figures"].get("time") or {}).get("t") or 0
            needs.append({"id": "time", "stage": "result", "severity": "block",
                          "title": "This year's value is held back: the roll doesn't move a discounting on by its rate",
                          "detail": "; ".join(_time_text(x, t) for x in g["time_off"][:4]) + ": its cash flow dates and "
                                    "its discount periods don't move together (one counted from the valuation date's input, "
                                    "the other from a copy of it, say), or its date isn't one the roll moves",
                          "go": {"step": "workbench", "anchor": "datesCard"}})
        held = [x for x in g.get("new_terms") or [] if x["hold"]]
        gone = [x for x in g.get("gone_terms") or [] if x["hold"]]
        if held or gone:
            needs.append({"id": "new-terms", "stage": "result", "severity": "block",
                          "title": "This year's value is held back: this year's model " + (
                              "adds and drops" if held and gone else "adds" if held else "drops") +
                                   " terms in the sums the value reads",
                          "detail": "; ".join([_term_text(x, "in") for x in held[:4]] + [_term_text(x, "gone from")
                                                                                        for x in gone[:4]])
                                    + (". Last year's model had no such term, and the value takes it in: confirm each "
                                       "belongs in this year's value" if held else "")
                                    + (". Last year's value took these in, and this year's model hasn't them: confirm "
                                       "each is gone" if gone else ""),
                          "go": {"step": "result", "anchor": "termsCard"}, "rows": [x["key"] for x in held + gone]})
        rows = [x["row"] + (f" ({x['label']})" if x.get("label") else "") for x in
                g["dcf_missing"] + g["blank_rows"] + g["weak_rows"] + g["timing_open"]]
        zero = [c for c, x in g["by_cell"].items() if not x["zero_roll"]["ok"]]
        if rows or zero or g.get("date_check") or not needs:
            needs.append({"id": "rows", "stage": "result", "severity": "block",
                          "title": "This year's value is held back: rows to find in this year's model",
                          "detail": (f"{len(rows)} row(s): {', '.join(rows[:6])}" if rows else "") +
                                    (f"; the zero-roll check fails on {', '.join(zero)}" if zero else "") +
                                    ("; this year's valuation date isn't known" if g.get("date_check") else ""),
                          "go": {"step": "result", "anchor": "rowsCard"}, "rows": rows})
    tc = res["figures"].get("time") or {}
    near = [x for x in tc.get("discountings") or [] if not x["ok"] and x not in ((g or {}).get("time_off") or [])]
    if near:
        needs.append({"id": "time-check", "stage": "result", "severity": "check",
                      "title": "The roll moves a discounting on by a little more or less than its rate",
                      "detail": "; ".join(_time_text(x, tc.get("t") or 0) for x in near[:4]) + ": a convention that "
                                "doesn't move evenly (a part-period stub), or a date that doesn't quite move with the rest",
                      "go": {"step": "workbench", "anchor": "datesCard"}})
    if g and tc and not tc.get("measured"):
        needs.append({"id": "time-none", "stage": "result", "severity": "info",
                      "title": "The roll's time value isn't measured",
                      "detail": (tc.get("why") or "") + ": that each discounting moves on by its rate is a person's to check",
                      "go": {"step": "result", "anchor": "bridgeCard"}})
    if g and not dc.get("moved"):
        needs.append({"id": "dates-none", "stage": "result", "severity": "check",
                      "title": "No valuation date cell was found to move",
                      "detail": "the discountings under the value couldn't be traced and no cell is labelled as the valuation "
                                "date, so this year's value may still be discounted to last year's",
                      "go": {"step": "workbench", "anchor": "datesCard"}})
    if g and g.get("rebuilt"):
        rr = g.get("rebuilt_rows") or []
        needs.append({"id": "rebuilt", "stage": "result", "severity": "check",
                      "title": f"This year's model looks rebuilt: {100 * g['family']:.0f}% of its line items are last year's",
                      "detail": (f"{len(rr)} row(s) the value reads were found by their numbers or words, not their labels: "
                                 "check they're the right ones" if rr else "the rows the value reads were found by their labels")
                                + ("" if g["reliable"] else "; this year's value waits on the rows listed"
                                   + (" and the terms to confirm" if g.get("terms_held") else "")),
                      "go": {"step": "result", "anchor": "rowsCard"}})
    if g and dc.get("by_label"):
        needs.append({"id": "dates-label", "stage": "result", "severity": "check",
                      "title": f"The valuation date moved is the one labelled so ({dc['moved'][0]})",
                      "detail": "the discountings under the value couldn't be traced, so it isn't confirmed that they read it",
                      "go": {"step": "workbench", "anchor": "datesCard"}})
    moved = [(c, x["zero_roll"]["ratio"]) for c, x in ((g or {}).get("by_cell") or {}).items()
             if x["zero_roll"]["ok"] and x["zero_roll"]["ratio"] is not None
             and not ovmod.ZERO_ROLL_CHECK[0] <= x["zero_roll"]["ratio"] <= ovmod.ZERO_ROLL_CHECK[1]]
    if moved:
        needs.append({"id": "zero-roll", "stage": "result", "severity": "check",
                      "title": f"At last year's valuation date, this year's model gives {moved[0][1]:.2f}× last year's value",
                      "detail": "; ".join(f"{c}: {r:.2f}×" for c, r in moved) + ": a move of that size is usually the new "
                                "forecast, but a row matched wrongly looks the same. Check the new forecast's step",
                      "go": {"step": "result", "anchor": "bridgeCard"}})
    out_terms = [x for x in (g or {}).get("new_terms") or [] if not x["in_value"]]
    out_gone = [x for x in (g or {}).get("gone_terms") or [] if not x["in_value"]]
    if out_terms or out_gone:
        needs.append({"id": "new-terms-out", "stage": "result", "severity": "check",
                      "title": "This year's model " + ("adds and drops" if out_terms and out_gone else "adds" if out_terms
                                                       else "drops") + " terms in sums beside the rows the overlay reads",
                      "detail": "; ".join([_term_text(x, "in") for x in out_terms[:4]] +
                                          [_term_text(x, "gone from") for x in out_gone[:4]]) +
                                ". The overlay reads the sum's other terms, not the sum, so this year's value doesn't "
                                "move with them: check whether it should",
                      "go": {"step": "result", "anchor": "termsCard"}, "rows": [x["key"] for x in out_terms + out_gone]})
    lines = (g or {}).get("new_lines") or []
    if lines:
        needs.append({"id": "new-lines", "stage": "result", "severity": "check",
                      "title": "This year's model has cash-flow lines the overlay doesn't read",
                      "detail": "; ".join(f"{x['row']} ({x['label']}): {x['periods']} period(s), {x['total']:,.1f} in total"
                                          for x in lines[:4]) + (f"; {len(lines) - 4} more" if len(lines) > 4 else "") +
                                ". Last year's overlay had nothing of them to discount, so the roll leaves them out: check "
                                "whether this year's value should take them in",
                      "go": {"step": "result", "anchor": "linesCard"}, "rows": [x["row"] for x in lines]})
    sc = res.get("scenario") or {}
    odd = [x for x in sc.get("selectors") or [] if x["status"] != "same"]
    if odd:
        shown = lambda v: "–" if v is None else f"{v:g}" if isinstance(v, float) else str(v)

        def said(x):
            if x["status"] == "unmatched":
                return f"{x['cell']} ({x['label']}) is {shown(x['this_year'])}, with none like it in last year's model"
            if x["status"] == "gone":
                return f"last year's {x['last_cell']} ({x['label']}) was {shown(x['last_year'])}, not found this year"
            return (f"{x['cell']} ({x['label']}) is {shown(x['this_year'])}, last year's {shown(x['last_year'])}"
                    + (f" ({shown(x['overlay'])} in the overlay's copy)" if x.get("overlay") is not None else ""))
        now, then = ((sc.get("saved") or {}).get(k) or {} for k in ("this_year", "last_year"))
        when = ", ".join(x for x in (f"this year's model was saved {now['date']}" if now.get("date") else "",
                                     f"last year's {'' if now.get('date') else 'model was saved '}{then['date']}"
                                     if then.get("date") else "") if x)
        needs.append({"id": "scenario", "stage": "result", "severity": "check",
                      "title": "Confirm the scenario this year's client model is saved on: " + (
                          "it isn't last year's" if any(x["status"] == "differs" for x in odd) else
                          "it can't be matched to last year's"),
                      "detail": "; ".join(said(x) for x in odd[:4]) + (f"; {len(odd) - 4} more" if len(odd) > 4 else "") +
                                (f"; {when}" if when else "") + ". The overlay reads the model's saved values, so this "
                                "year's value takes the scenario the model was saved on",
                      "go": {"step": "result", "anchor": "scenarioCard"}})
    for r in (res.get("reconcile") or {}).get("rows") or []:
        if r.get("ok") is False:
            bad = [f"{e}: report {r['report'][e]}, Python {r['python'][e]:,.1f}" for e, ok in r["ties"].items()
                   if ok is False and r["python"].get(e) is not None]
            needs.append({"id": f"reconcile-{r['key']}", "stage": "result", "severity": "check",
                          "title": f"{r['label']} doesn't reconcile to the report",
                          "detail": "; ".join(bad) or "Python couldn't split the value this way",
                          "go": {"step": "rebuild", "anchor": "reconcileCard"}})
    ex = (res.get("head") or {}).get("basis") != "cum"
    for i, h in enumerate(x for x in res.get("held") or [] if x["held"]):
        sg = h.get("suggestion") or {}
        if ex and re.search(r"distribution|dividend", h.get("label") or "", re.I) and h.get("value"):
            sg = {**sg, "text": (sg.get("text") or "") + " This year's value is ex-distribution: the distribution to "
                  "deduct is the one declared at this year's date, not last year's."}
        needs.append({"id": f"held-{i}", "stage": "result", "severity": "check",
                      "title": f"{h['label']}: {h['value']:,.1f} held at last year's" + (
                          f"; {sg['value']:,.1f} in this year's model" + (", checked" if sg["status"] == "checked" else "")
                          if sg.get("value") is not None else ""),
                      "detail": sg.get("text") or "", "go": {"step": "result", "anchor": "heldCard"}})
    inv = res.get("methods") or {}
    if inv.get("methods") and not inv.get("ties"):
        needs.append({"id": "method-ties", "stage": "result", "severity": "check",
                      "title": "The recomputed methods aren't like for like: the recompute doesn't give the overlay's figure",
                      "detail": "each discounting recomputed in code, as the overlay discounts, should give the default's "
                                "figure to the cent; it doesn't, so the recomputed methods (mid-period, mid-year, the other "
                                "day count) differ from the default by more than their convention",
                      "go": {"step": "result", "anchor": "methodsCard"}})
    if inv.get("asked") and inv.get("asked") != inv.get("preferred"):
        m = next((x for x in inv.get("methods") or [] if x["key"] == inv["asked"]), {})
        needs.append({"id": "method", "stage": "result", "severity": "check",
                      "title": "The preferred method can't be worked out here: this year's value is the default's",
                      "detail": m.get("why") or inv.get("error") or "", "go": {"step": "result", "anchor": "methodsCard"}})
    rt = ((res.get("inputs") or {}).get("rate") or {}).get("this_year") or {}
    if rt and not rt.get("applied"):
        needs.append({"id": "rate-this-year", "stage": "result", "severity": "check",
                      "title": "This year's discount rate isn't applied: the roll-forward is at last year's",
                      "detail": rt.get("why") or "", "go": {"step": "result", "anchor": "rateCard"}})
    names = {"rate": "discount rate", "growth": "terminal growth rate", "multiple": "exit multiple",
             "franking": "franking credit utilisation"}
    for key, what in names.items():
        for end, r in (((res.get("inputs") or {}).get(key) or {}).get("ends") or {}).items():
            if r.get("ok") is not False:
                continue
            at = ((f"{r['value']:.2f}x" if key == "multiple" else f"{100 * r['value']:.2f}%") if r.get("value") is not None
                  else "") + (f" in {r['cell']}" if r.get("cell") else "")
            bad = [c["text"] for c in r.get("checks") or [] if c["ok"] is False]
            needs.append({"id": f"{key}-{end}", "stage": "result", "severity": "check",
                          "title": f"The {end} end's {what} ({at.strip()}) " + (
                              "isn't sourced to a cell" if not r.get("sourced") else "doesn't check out"),
                          "detail": "; ".join(bad) or r.get("note") or "", "go": {"step": "rebuild", "anchor": "inputsCard"}})
    said = next((f for f in wb.reference(eid) if f["key"] == "valuation_date" and f.get("value")), None)
    if said:
        v = int(said["value"])
        iso = f"{v // 10000:04d}-{v // 100 % 100:02d}-{v % 100:02d}"
        off = sorted({(a.get("valuation_date"), a.get("valuation_date_source")) for e in ("low", "high")
                      for a in (res.get("assumptions") or {}).get(e) or [] if a.get("valuation_date")
                      and a["valuation_date"] != iso})
        if off:
            needs.append({"id": "date-overlay", "stage": "result", "severity": "check",
                          "title": f"The overlay discounts to {off[0][0]} ({off[0][1]}), the report says {said.get('value_text')}",
                          "detail": "the valuation date the discount factors read isn't the report's",
                          "go": {"step": "workbench", "anchor": "datesCard"}})
    for end in ("low", "high"):
        t = res["tie"][end]
        if not t["ok"]:
            needs.append({"id": f"tie-{end}", "stage": "result", "severity": "check",
                          "title": f"The overlay's {end} doesn't round to the report's ({t['report_text']})",
                          "detail": f"Excel saved {t['saved']}", "go": {"step": "rebuild"}})
    v = res["values"]
    line = (f"last year {v['report']['mid']:,.1f} → this year {v['this_year']['mid']:,.1f} (mid)"
            if v.get("this_year") and v["this_year"].get("mid") is not None else
            f"last year {v['report']['mid']:,.1f} (mid) rebuilt; this year's held back")
    status = "blocked" if any(n["severity"] == "block" for n in needs) else "attention" if needs else "done"
    return status, line, {"needs": needs}


REVIEW_PROMPT = """You are the reviewing partner's assistant on a recurring infrastructure valuation. The workbench has
rebuilt last year's equity value in Python from last year's overlay and client model, tied it to the report, and rolled
it forward onto this year's client model at last year's discount rate. Read the run below end to end and say whether
anything looks wrong or implausible: a bridge step out of proportion (the time value should be about the equity value
times the discount rate over the years between the dates; the cash flows paid about last year's first year's cash
flow), a figure that doesn't tie, a key fact that doesn't fit the model, a date that doesn't follow, a cash flow profile
that jumps. Be specific and brief. choice: "ok" if nothing needs a person, else "concerns", with each concern in
concerns (title, detail, severity: "check" for something a person should look at, "info" for a note; years: the
financial years it's about, as cash_flows_by_year labels them (FY2045), else empty; step: the bridge step it's about,
one of {steps}, else "").

How the bridge is built, so you don't flag what follows from it: the primary approach discounts cash flows to equity
at the cost of equity (some overlays discount free cash flows at a WACC and deduct net debt). Every discounting under
the equity value moves (the main cash flows with the terminal value, and any other stream the overlay discounts, such
as franking credits). The inputs typed in the overlay outside the discountings (a net debt, a cash balance, a declared
distribution, an adjustment) stay at last year's figures unless a person set this year's: see held_inputs, and the
"held" step of the bridge where they were set. So the time value is about the discount rate times the discounted
streams (not the equity value) over the years between the dates, and the cash flows paid are the first year of every
stream (the periods ending on or before the new valuation date: they're cut off this year's discounting). The time
value, the cash flows and the new forecast are at last year's discount rate; where a person set this year's
(discount_rate_this_year), the "rate" step is the move from last year's rate to it, else it's nil. methods is this
year's value worked out other ways (the period ending on the date kept, mid-period, mid-year, the other day count,
the mid at the midpoint rate), each against the default; where a person prefers one, the "method" step is the move
to it. A method far from the default is a convention, not an error: say so only where the preferred one looks wrong.

The run:
{run}"""


def _review_job(eid: int, key: str):
    e = wb.get(eid)
    res = e.get("result") or {}
    if not res:
        return "done", "nothing to review yet", {}
    ch = res.get("chart") or {}
    streams = {e: [{"stream": x.get("label"), "present_value": x.get("pv"), "rate": x.get("rate"),
                    "undiscounted": x.get("undiscounted"), "terminal_value": x.get("terminal_value")}
                   for x in rows if not x.get("error")] for e, rows in (res.get("assumptions") or {}).items()}
    run = {"discounted_streams_under_the_value": streams,
           "facts": [{k: f.get(k) for k in ("key", "value_text", "low_text", "high_text", "basis", "status", "decided_by")}
                     for f in e["facts"]],
           "equity_value": {k: (res.get("head") or {}).get(k) for k in ("basis", "low", "mid", "high", "units", "why")},
           "where_in_overlay": res.get("where"), "tie": res.get("tie"), "values": res.get("values"),
           "bridge_mid": [(s["label"], s["value"]) for s in (res.get("bridges") or {}).get("mid", {}).get("steps", [])],
           "valuation_dates": {"last_year": ((e.get("overlay") or {}).get("roll") or {}).get("prior_valuation_date"),
                               "this_year": (res.get("bridges") or {}).get("valuation_date")},
           "cash_flows_by_year": ch.get("series"), "this_year_gaps_reliable": ((res.get("figures") or {}).get("gaps") or {}).get("reliable"),
           "assumptions": res.get("assumptions"),
           "discount_rate_this_year": {k: (((res.get("inputs") or {}).get("rate") or {}).get("this_year") or {}).get(k)
                                       for k in ("low", "high", "was", "applied", "why")},
           "methods": {"preferred": (res.get("methods") or {}).get("preferred"),
                       "each": [{k: m.get(k) for k in ("key", "label", "mid", "vs_default", "why")}
                                for m in (res.get("methods") or {}).get("methods") or []]},
           "terminal_value_basis": {k: (res.get("terminal") or {}).get(k) for k in ("label", "phrase", "page", "multiple")},
           "held_inputs": [{"input": h["label"], "last_year": h["value"], "this_year": h.get("this_year"),
                            "still_held": h["held"], "suggested_from_this_years_model": (h.get("suggestion") or {}).get("value")}
                           for h in res.get("held") or []],
           "run_log": [f"{h['stage']}: {h['event']}: {h['text']}" for h in history(eid, limit=40)][::-1]}
    import llm
    model = e.get("reviewer_model") or wb.DEFAULT_REVIEWER
    steps = [s["key"] for s in (res.get("bridges") or {}).get("mid", {}).get("steps", [])]
    done = [h for h in history(eid, "review", "review", 20) if h["event"] == "decide" and h.get("data")]
    same = next((h for h in done if h["inputs"] == key), None)
    if same and _decisions(eid, "review", "review", key) >= BUDGET:
        # reviewed once already on what the result says now: the same points stand (a "try again", or a rerun that
        # worked out the same result, doesn't buy another review, as it doesn't buy another decision)
        out = same["data"]
        log(eid, "review", "note", "reviewed once already on this result: the same points stand", issue="review", inputs=key)
    else:
        schema = _schema("run_review", ["ok", "concerns"], {"concerns": {"type": "array", "items": {
            "type": "object", "additionalProperties": False, "required": ["title", "detail", "severity", "years", "step"],
            "properties": {"title": _S, "detail": _S, "severity": {"type": "string", "enum": ["check", "info"]},
                           "years": {"type": "array", "items": _S}, "step": _S}}}})
        before = [{"title": c["title"], "detail": c["detail"], "severity": c["severity"]}
                  for c in (done[0]["data"].get("concerns") or [])] if done else []
        r = llm.create(llm.client(interactive=False), model, input=REVIEW_PROMPT.format(
            steps=", ".join(steps) or "(none)", run=json.dumps(run, indent=1, default=str)[:60000]) + (
            "\n\nYour points on the run before this one (its result has changed since): keep a point's title where it "
            "still stands, so a person can follow it, and drop what the new result settles:\n" + json.dumps(before, indent=1)
            if before else ""), text={"format": schema}, max_output_tokens=3000, purpose="orchestrator-review")
        if r.usage:
            wb._logger(eid)(model, r.usage, "orchestrator-review")
        out = json.loads(r.output_text)
        log(eid, "review", "decide", f"{model}: {out['choice']} — {out['reason']}", issue="review", inputs=key, data=out)
    # a point's years and step are checked against the run: only years the chart has, only a step the bridge has
    years = set(ch.get("years") or [])
    needs = [{"id": f"review-{i}", "stage": "review", "severity": c["severity"], "title": c["title"], "detail": c["detail"],
              "years": [y for y in dict.fromkeys(c.get("years") or []) if y in years],
              "step": c.get("step") if c.get("step") in steps else None,
              "go": {"step": "result", "anchor": f"review-{i}"}} for i, c in enumerate(out.get("concerns") or [])]
    return ("attention" if any(n["severity"] == "check" for n in needs) else "done"), \
        out["reason"] if out["choice"] == "ok" else f"{len(needs)} point(s) to look at", {"needs": needs, "review": out}


def _map_job(eid: int, key: str):
    wb._map(eid)
    m = wb.get(eid).get("map") or {}
    found = sum(1 for x in m.get("report_overlay") or [] if x["matches"] and x["matches"][0].get("located", True))
    return "done", f"{found} report figure(s) located; {len(m.get('chain') or [])} overlay row(s) traced to the client model", {}


STAGE_JOBS = {"roles": _roles_job, "rebuild": _rebuild_job, "rows": _rows_job, "result": _result_job,
              "review": _review_job, "map": _map_job}


# ---- what the page shows, and a person's actions ----------------------------------------------------------------

_STEP_COLS = {"rebuild": "overlay", "rows": "rows", "map": "map"}
PAGE = {"files": "workbench", "facts": "report", "roles": "workbench", "rebuild": "rebuild", "rows": "result",
        "result": "result", "review": "result", "map": "result"}
PAGE_ANCHOR = {"workbench": "filesCard", "report": "reportCard", "rebuild": "tieCard", "result": "bridgeCard"}
# What a need asks of a person, by its id (the first that fits): its kind, and how the page names it
KINDS = (("fact-", "confirm-fact", "A fact to confirm"), ("missing-", "add-fact", "A fact to add"),
         ("roles", "confirm-roles", "The roles to confirm"), ("no-reads", "check-roles", "The roles to check"),
         ("dates-", "check-date", "A date to check"), ("zero-roll", "check-forecast", "A move to check"),
         ("held-", "check-held", "An input held at last year's"), ("rebuilt", "confirm-rows", "Rows to confirm"),
         ("new-lines", "check-lines", "Cash-flow lines to check"), ("scenario", "check-scenario", "A scenario to confirm"),
         ("date-overlay", "check-date", "A date to check"),
         ("date", "confirm-date", "A date to confirm"), ("equity", "pick-cells", "Cells to pick"),
         ("rows", "find-rows", "Rows to find"), ("reconcile-", "check-reconcile", "A reconciliation to check"),
         ("rate-", "check-input", "A model input to check"), ("growth-", "check-input", "A model input to check"),
         ("franking-", "check-input", "A model input to check"), ("tie-", "check-tie", "A tie to check"),
         ("review-", "review-point", "A review point"), ("failed-", "retry", "A step that failed"),
         ("method", "check-method", "A method to check"), ("cf-", "check-flows", "Cash flows to check"),
         ("interest", "check-interest", "The interest valued"), ("basis", "check-basis", "The basis to confirm"))


def dress(n: dict) -> dict:
    """A need with its kind (what a person is asked to do) and where it lands: its page and the card on it."""
    i = n["id"]
    kind, what = next(((k, w) for p, k, w in KINDS if i == p.rstrip("-") or i.startswith(p if p.endswith("-") else p + "-")),
                      ("note", "A note"))
    go = dict(n.get("go") or {})
    if kind != "retry":
        go.setdefault("step", PAGE.get(n["stage"], "workbench"))
        go.setdefault("anchor", PAGE_ANCHOR.get(go["step"]))
    return {**n, "kind": kind, "kind_label": what, "go": go}


_TYPICAL: dict = {}


def typical(eid: int) -> dict:
    """How long each stage takes: its last run on this engagement, else the median of its last ten anywhere (from the
    run log's starts and ends), so the page can say how long is left. {stage: seconds or None}"""
    mx = (wb._q("SELECT MAX(id) AS m FROM runlog") or [{}])[0].get("m")
    if eid in _TYPICAL and _TYPICAL[eid][0] == mx:  # nothing logged since: as worked out last time
        return _TYPICAL[eid][1]
    rows = wb._q("SELECT engagement_id AS e, stage, event, at FROM runlog WHERE event IN "
                 "('start', 'done', 'attention', 'blocked', 'failed') ORDER BY id DESC LIMIT 4000")
    began, mine, anywhere = {}, {}, {}
    for r in reversed(rows):
        k = (r["e"], r["stage"])
        if r["event"] == "start":
            began[k] = r["at"]
        elif k in began:
            secs = r["at"] - began.pop(k)
            if r["event"] != "failed" and secs >= 0:
                anywhere.setdefault(r["stage"], []).append(secs)
                if r["e"] == eid:
                    mine[r["stage"]] = secs
    mid = lambda xs: sorted(xs[-10:])[len(xs[-10:]) // 2]
    out = {s: mine.get(s) or (mid(anywhere[s]) if anywhere.get(s) else None) for s in STAGES}
    _TYPICAL[eid] = (mx, out)
    return out


def view(eid: int) -> dict:
    """The stages (status, note, what's running now), the needs-you list, and the latest of the run log. Reading it
    starts nothing."""
    e = wb._q("SELECT * FROM engagements WHERE id=?", eid)
    e = e[0] if e else {}
    stages, needs, took = [], [], typical(eid)
    for name in STAGES:
        rec = stage(eid, name)
        live = rec["note"]
        if rec["status"] == "running" and name in _STEP_COLS:
            live = e.get(f"{_STEP_COLS[name]}_step") or live
        if rec["status"] == "running" and name == "rebuild" and e.get("doctor_status") == "running":
            live = "the doctor: " + (e.get("doctor_step") or "")
        stages.append({"stage": name, "label": LABEL[name], "status": rec["status"] or "waiting", "note": live,
                       "started_at": rec["started_at"], "finished_at": rec["finished_at"], "typical_secs": took.get(name)})
        if rec["status"] != "waiting":  # a stage waiting on the one before: its needs may no longer stand
            needs += [dress(n) for n in rec["data"].get("needs") or []]
    sev = {"block": 0, "check": 1, "info": 2}
    needs.sort(key=lambda n: sev.get(n["severity"], 3))
    running = [{"job": j if isinstance(j, str) else j[0], "state": st} for (x, j), st in list(_active.items()) if x == eid]
    return {"stages": stages, "needs": needs, "log": history(eid, limit=40), "running": running,
            "busy": bool(running) or any(s["status"] in MOVING for s in stages)}


def overview() -> list[dict]:
    """Every engagement at a glance, for the list: where it is (empty, needs you, running, finished, waiting), a word
    on it, how many things are for a person, and its equity value (the saved result's few fields the list shows, read
    by SQLite rather than the whole result). Reading it starts nothing."""
    pick = {"values": "$.values", "units": "$.bridges.units", "head_units": "$.head.units", "basis": "$.head.basis",
            "rolled_to": "$.bridges.valuation_date", "held": "$.bridges.held"}
    cols = ", ".join(f"json_extract(result_json, '{path}') AS \"{k}\"" for k, path in pick.items())
    saved = {r["id"]: r for r in wb._q(f"SELECT id, {cols} FROM engagements WHERE result_json IS NOT NULL")}
    out = []
    for e in wb.all_engagements():
        eid = e["id"]
        v = view(eid)
        st = {s["stage"]: s for s in v["stages"]}
        needs = [n for n in v["needs"] if n["severity"] != "info"]
        blocks = [n for n in needs if n["severity"] == "block"]
        moving = next((s for s in v["stages"] if s["status"] == "running"), None) or \
            next((s for s in v["stages"] if s["status"] == "queued"), None)
        r = saved.get(eid) or {}
        values = json.loads(r["values"]) if r.get("values") else None
        files = e["n_docs"] + e["n_workbooks"]
        if not files:
            state, note = "empty", "no files yet"
        elif blocks:
            state, note = "needs you", blocks[0]["title"]
        elif moving:
            state, note = "running", f"{LABEL[moving['stage']]}: {moving['note'] or 'in line'}"
        elif (values or {}).get("this_year") and st["result"]["status"] in SETTLED:
            state, note = "finished", f"{len(needs)} to check" if needs else "up to date with the files"
        else:
            waiting = next((s for s in v["stages"] if s["status"] not in SETTLED), None)
            state = "waiting"
            note = f"{files} of 4 files in" if files < 4 else \
                f"{LABEL[waiting['stage']]}: {waiting['note'] or 'waiting for the step before it'}" if waiting else "up to date"
        out.append({"id": eid, "name": e["name"], "updated_at": e["updated_at"], "files": files, "state": state,
                    "note": note, "needs": len(needs), "blocks": len(blocks), "values": values,
                    "units": r.get("units") or r.get("head_units"), "basis": r.get("basis"),
                    "valuation_date": r.get("rolled_to"), "held": bool(r.get("held"))})
    return out


def retry(eid: int, name: str) -> dict:
    """A person's "try again": the stage forgets its inputs, so it runs on the next look."""
    if name not in STAGES:
        raise ValueError(f"no stage {name}")
    if name == "files":
        for d in wb.documents(eid):
            if d["status"] == "error":
                wb.retry_document(d["id"])
        for w in wb.workbooks(eid):
            if w["status"] == "error":
                wb.retry_workbook(eid, w["id"])
    elif name == "facts":
        d = _report(_snapshot(eid))
        if d:
            wb._set("documents", d["id"], facts_status=None, facts_error=None)
    else:
        _put(eid, name, inputs=None, status="waiting", note="trying again")
    log(eid, name, "person", "try again")
    poke()
    return view(eid)


def person(eid: int, name: str, text: str) -> None:
    """A person's action, for the run log (the stage it touches reruns by itself when its inputs change)."""
    log(eid, name, "person", text)
    poke()
