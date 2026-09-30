"""Infra Val Workbench: last year's valuation report, overlay and client models to this year's value.
    uv run uvicorn app.server:app --port 8003      then open http://localhost:8003

Upload last year's report (PDF / PPTX), last year's client model and overlay, and this year's client model. The
orchestrator (engine/orchestrator.py) does the rest: the report's key facts, the roles, last year's value rebuilt in
Python and tied to the report, the roll-forward onto this year's model, the value bridge and the cash-flow chart.
A person steps in where it asks, and can change anything it decided.
"""
import os
import sys
import tempfile
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from pydantic import BaseModel
from starlette.concurrency import run_in_threadpool

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "engine"))
import library  # noqa: E402
import orchestrator  # noqa: E402
import usage  # noqa: E402
import workbench  # noqa: E402
import xlruntime  # noqa: E402

# Model deployments offered. gpt-6-luna extracts; gpt-6-sol reviews, gives second opinions and decides for the
# orchestrator.
MODELS = ["gpt-6-luna", "gpt-6-sol", "gpt-4o", "gpt-4o-mini", "gpt-5-nano"]


@asynccontextmanager
async def lifespan(app):
    library.start()       # builds uploaded workbooks
    orchestrator.start()  # everything after that
    yield


app = FastAPI(lifespan=lifespan, title="Infra Val Workbench")


def _run(fn, *args, **kw):
    """Run in a worker thread; ValueError -> 400, missing -> 404."""
    async def go():
        try:
            out = await run_in_threadpool(lambda: fn(*args, **kw))
        except ValueError as e:
            raise HTTPException(400, str(e))
        if out is None:
            raise HTTPException(404, "not found")
        return out
    return go()


@app.get("/")
def index():
    return FileResponse(Path(__file__).parent / "index.html")


@app.get("/charts.js")
def charts_js():  # the chart drawing (web/charts.js)
    return FileResponse(ROOT / "web" / "charts.js", media_type="text/javascript")


# ---- the logo in the header's top-left corner: your firm's, from the git-ignored brand/ folder, else the mascot ----
BRAND = ROOT / "brand"
LOGO_TYPES = {".svg": "image/svg+xml", ".png": "image/png", ".webp": "image/webp", ".jpg": "image/jpeg", ".jpeg": "image/jpeg"}
LOGO_MAX = 2 * 1024 * 1024
LOGO_SPEC = {"folder": "brand", "names": ["logo.svg", "logo.png"], "also": ["logo.webp", "logo.jpg"],
             "height_px": 144, "shown_px": 36, "max_width_px": 640, "max_bytes": LOGO_MAX,
             "note": "transparent background, light artwork (the header is dark)"}


def _logo() -> Path | None:
    return next((BRAND / f"logo{ext}" for ext in LOGO_TYPES if (BRAND / f"logo{ext}").exists()), None)


@app.get("/brand/logo")
def brand_logo():
    p = _logo()
    headers = {"Cache-Control": "no-cache", "Content-Security-Policy": "script-src 'none'"}  # an SVG never runs scripts
    if p:
        return FileResponse(p, media_type=LOGO_TYPES[p.suffix.lower()], headers=headers)
    return FileResponse(Path(__file__).parent / "mascot.svg", media_type="image/svg+xml", headers=headers)


@app.get("/api/brand")
def brand_info():
    p = _logo()
    return {"custom": bool(p), "file": f"brand/{p.name}" if p else None, "spec": LOGO_SPEC}


@app.post("/api/brand/logo")
async def brand_upload(file: UploadFile = File(...)):
    """Your firm's logo for the header: SVG, PNG, WebP or JPEG, kept in brand/ (never committed)."""
    ext = Path(file.filename or "").suffix.lower()
    data = await file.read(LOGO_MAX + 1)
    if ext not in LOGO_TYPES:
        raise HTTPException(400, "upload the logo as .svg or .png (or .webp / .jpg)")
    if len(data) > LOGO_MAX:
        raise HTTPException(400, "the logo must be 2 MB or less")
    ok = {".png": data[:8] == b"\x89PNG\r\n\x1a\n", ".jpg": data[:3] == b"\xff\xd8\xff", ".jpeg": data[:3] == b"\xff\xd8\xff",
          ".webp": data[:4] == b"RIFF" and data[8:12] == b"WEBP", ".svg": b"<svg" in data[:4096].lower()}[ext]
    if not ok:
        raise HTTPException(400, f"that file isn't a {ext[1:].upper()} image")
    BRAND.mkdir(exist_ok=True)
    for other in LOGO_TYPES:
        (BRAND / f"logo{other}").unlink(missing_ok=True)
    (BRAND / f"logo{'.jpg' if ext == '.jpeg' else ext}").write_bytes(data)
    return brand_info()


@app.delete("/api/brand/logo")
def brand_reset():
    for ext in LOGO_TYPES:
        (BRAND / f"logo{ext}").unlink(missing_ok=True)
    return brand_info()


@app.get("/api/config")
def config():
    return {"models": MODELS, "default_model": workbench.DEFAULT_MODEL, "default_reviewer": workbench.DEFAULT_REVIEWER,
            "default_arbiter": workbench.DEFAULT_ARBITER, "roles": workbench.rolesmod.ROLES,
            "stages": [{"stage": s, "label": orchestrator.LABEL[s]} for s in orchestrator.STAGES],
            "runtime": xlruntime.RUNTIME}


# ---- engagements and files ----------------------------------------------------------------------------------------

class NewEngagement(BaseModel):
    name: str


class EngagementPatch(BaseModel):
    name: str | None = None
    model: str | None = None
    reviewer_model: str | None = None
    arbiter_model: str | None = None


@app.get("/api/engagements")
def list_engagements():
    return workbench.all_engagements()


@app.post("/api/engagements")
async def new_engagement(body: NewEngagement):
    return await _run(workbench.create, body.name)


@app.get("/api/engagements/{eid}")
async def get_engagement(eid: int):
    return await _run(workbench.get, eid)


@app.patch("/api/engagements/{eid}")
async def patch_engagement(eid: int, body: EngagementPatch):
    for m in (body.model, body.reviewer_model, body.arbiter_model):
        if m and m not in MODELS:
            raise HTTPException(400, f"unknown model {m}")
    return await _run(workbench.update, eid, **body.model_dump())


@app.delete("/api/engagements/{eid}")
async def delete_engagement(eid: int):
    await _run(lambda: workbench.delete(eid) or True)
    return {"ok": True}


@app.post("/api/engagements/{eid}/files")
async def upload(eid: int, file: UploadFile = File(...)):
    if not workbench._q("SELECT 1 FROM engagements WHERE id=?", eid):
        raise HTTPException(404, "no such engagement")
    library.UPLOADS.mkdir(exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(dir=library.UPLOADS, suffix=".part")
    os.close(fd)
    tmp = Path(tmp_name)
    h = library.hashlib.sha256()
    with open(tmp, "wb") as out:
        while chunk := await file.read(1 << 20):
            h.update(chunk)
            out.write(chunk)
    got = await _run(workbench.add_upload, eid, tmp, file.filename, h.hexdigest())
    orchestrator.person(eid, "files", f"uploaded {file.filename}")
    return got


@app.delete("/api/engagements/{eid}/workbooks/{fid}")
async def unlink_workbook(eid: int, fid: int):
    await _run(lambda: workbench.remove_workbook(eid, fid) or True)
    orchestrator.person(eid, "files", f"removed workbook {fid}")
    return {"ok": True}


class DateCheck(BaseModel):
    valuation_date: str


@app.post("/api/engagements/{eid}/workbooks/{fid}/date")
async def confirm_date(eid: int, fid: int, body: DateCheck):
    """Your check of a workbook's valuation date: the roll-forward runs between last year's and this year's."""
    got = await _run(workbench.confirm_date, eid, fid, body.valuation_date)
    orchestrator.person(eid, "rebuild", f"valuation date of workbook {fid}: {body.valuation_date}")
    return got


@app.post("/api/engagements/{eid}/workbooks/{fid}/retry")
async def retry_workbook(eid: int, fid: int):
    await _run(lambda: workbench.retry_workbook(eid, fid) and True)
    orchestrator.person(eid, "files", f"retry workbook {fid}")
    return {"ok": True}


@app.post("/api/engagements/{eid}/workbooks/{fid}/rebuild")
async def rebuild_workbook(eid: int, fid: int):
    """Process a workbook again from scratch."""
    await _run(lambda: workbench.rebuild_workbook(eid, fid) and True)
    orchestrator.person(eid, "files", f"rebuild workbook {fid}")
    return {"ok": True}


@app.get("/api/documents/{did}/image/{kind}/{name}")
def document_image(did: int, kind: str, name: str):
    """One of a report's images: a table's (tables/...) or a page's (pages/...)."""
    p = workbench.document_image(did, f"{kind}/{name}")
    if not p:
        raise HTTPException(404, "no such image")
    return FileResponse(p, media_type="image/png")


@app.get("/api/documents/{did}")
async def get_document(did: int):
    return await _run(workbench.document, did)


@app.delete("/api/documents/{did}")
async def delete_document(did: int):
    await _run(lambda: workbench.remove_document(did) or True)
    orchestrator.poke()
    return {"ok": True}


@app.post("/api/documents/{did}/retry")
async def retry_document(did: int):
    await _run(lambda: workbench.retry_document(did) or True)
    orchestrator.poke()
    return {"ok": True}


@app.post("/api/documents/{did}/rebuild")
async def rebuild_document(did: int):
    """Read a report again from scratch; its key facts are extracted again."""
    await _run(lambda: workbench.rebuild_document(did) and True)
    orchestrator.poke()
    return {"ok": True}


# ---- the run: the orchestrator's stages, the needs-you list, the run log ----------------------------------------

@app.get("/api/engagements/{eid}/run")
async def run_view(eid: int):
    return await _run(orchestrator.view, eid)


@app.post("/api/engagements/{eid}/run/retry/{stage}")
async def run_retry(eid: int, stage: str):
    """A person's "try again" on a stage: it runs on the orchestrator's next look."""
    return await _run(orchestrator.retry, eid, stage)


@app.get("/api/engagements/{eid}/runlog")
async def run_log(eid: int, stage: str | None = None, limit: int = 200):
    return await _run(orchestrator.history, eid, stage, None, limit)


@app.get("/api/engagements/{eid}/calls")
async def get_calls(eid: int):
    """Tokens, cost and time by file and by step, and the latest calls (the call log, calllog.py)."""
    return await _run(workbench.calls_view, eid)


@app.get("/api/engagements/{eid}/calls/{cid}")
async def get_call(eid: int, cid: int):
    """One call in full: what was sent and what came back."""
    return await _run(workbench.call_view, eid, cid)


@app.get("/api/usage")
def get_usage(session: str | None = None):
    s = usage.summary(session)
    return {"total": s["total"], "session": s["session"], "by_model": s["by_model"], "by_purpose": s["by_purpose"]}


# ---- a person's decisions ---------------------------------------------------------------------------------------

class FactDecision(BaseModel):
    action: str  # approve | use_suggestion | edit | reject | reset
    fields: dict | None = None


@app.put("/api/facts/{fact_id}")
async def put_fact(fact_id: int, body: FactDecision):
    f = await _run(workbench.set_fact, fact_id, body.action, body.fields)
    orchestrator.person(f["engagement_id"], "facts", f"{body.action} {f.get('label') or f['key']}")
    return f


class Roles(BaseModel):
    roles: dict


@app.post("/api/engagements/{eid}/roles/suggest")
async def suggest_roles(eid: int):
    """Look again: the orchestrator suggests and weighs the roles on its next look."""
    return await _run(orchestrator.retry, eid, "roles")


@app.put("/api/engagements/{eid}/roles")
async def put_roles(eid: int, body: Roles):
    got = await _run(workbench.confirm_roles, eid, body.roles, "you")
    orchestrator.person(eid, "roles", "confirmed " + ", ".join(k.replace("_", " ") for k in body.roles))
    return got


class Profile(BaseModel):
    fields: dict  # {"fy_end_month": 6 | None, "horizon": "fixed" | "rolling" | None}


@app.get("/api/engagements/{eid}/profile")
async def get_profile(eid: int):
    """The engagement's profile: financial-year end, horizon, periods, units, discounting; detected, and set."""
    return await _run(workbench.profile_view, eid)


@app.put("/api/engagements/{eid}/profile")
async def put_profile(eid: int, body: Profile):
    got = await _run(workbench.set_profile, eid, body.fields)
    orchestrator.person(eid, "result", "profile: " + ", ".join(f"{k} {v}" for k, v in body.fields.items()))
    return got


class Equity(BaseModel):
    low: str | None = None   # "Sheet!C9": the cell holding last year's equity value, the low end (None: clear)
    high: str | None = None
    basis: str | None = None  # "ex" / "cum"


@app.put("/api/engagements/{eid}/equity")
async def put_equity(eid: int, body: Equity):
    """Your pick of the overlay cells holding last year's equity value (low and high)."""
    def pick():
        if not body.low:
            orchestrator.set_equity_pick(eid, None, "you")
            return {"ok": True}
        import result
        _, summary = workbench.overlay_session(eid)
        cands = {c["cell"]: c for c in result.candidates(summary, _head(eid), limit=60)}
        lo, hi = (cands.get(x) or _cell(summary, x) for x in (body.low, body.high or body.low))
        orchestrator.set_equity_pick(eid, {"low": lo["cell"], "high": hi["cell"], "mid": None, "scale": lo["scale"],
                                           "sign": lo["sign"], "basis": body.basis or lo.get("basis") or "ex",
                                           "label": lo.get("label"), "how": "picked by you"}, "you")
        return {"ok": True}
    got = await _run(pick)
    orchestrator.person(eid, "rebuild", f"equity value cells: {body.low} / {body.high}")
    return got


def _head(eid: int) -> dict:
    import keyfacts
    head = keyfacts.conclusion(workbench.reference(eid), orchestrator._report_md(eid))
    if not head:
        raise ValueError("the report's equity value isn't among the key facts")
    return head


def _cell(summary: dict, ref: str) -> dict:
    """A cell given by hand: it must be on the overlay's sheets; its units are the report's (scale 1)."""
    import overlay as ov
    s, r, c = ov.parse_a1(ref)
    if s not in summary["sheets"]:
        raise ValueError(f"{ref} isn't on the overlay's sheets ({', '.join(summary['sheets'])})")
    return {"cell": ref, "scale": 1.0, "sign": 1, "label": None}


class RowPick(BaseModel):
    prior: str
    current: str | None = None


@app.post("/api/engagements/{eid}/overlay/rowpick")
async def row_pick(eid: int, body: RowPick):
    """Your choice of this year's row for one of last year's rows ("-": keep last year's values; none: back to what
    was found)."""
    got = await _run(workbench.row_pick, eid, body.prior, body.current)
    orchestrator.person(eid, "rows", f"{body.prior} -> {body.current or 'as found'}")
    return got


@app.get("/api/engagements/{eid}/overlay/row")
async def overlay_row(eid: int, row: str):
    """What this year's model has for one of last year's rows (Sheet!rN), with the alternatives."""
    return await _run(workbench.row_info, eid, row)


class ThisYearDate(BaseModel):
    valuation_date: str | None = None


@app.post("/api/engagements/{eid}/roll/this_year_date")
async def this_year_date(eid: int, body: ThisYearDate):
    """This year's valuation date for the engagement (None clears it): the roll-forward runs to it."""
    got = await _run(workbench.set_this_year_date, eid, body.valuation_date)
    orchestrator.person(eid, "rows", f"this year's valuation date: {body.valuation_date or 'as the model says'}")
    return got


# ---- what the pages read ------------------------------------------------------------------------------------------

@app.get("/api/engagements/{eid}/models/{fid}")
async def model_dashboard(eid: int, fid: int):
    """A workbook's dashboard, for the Map's details."""
    return await _run(workbench.model_dashboard, eid, fid)


@app.get("/api/engagements/{eid}/models/{fid}/rows")
async def model_rows(eid: int, fid: int, sheet: str | None = None, q: str | None = None, mapped: bool = False,
                     limit: int = 200):
    return await _run(workbench.model_rows, eid, fid, sheet, q, mapped, limit)


@app.get("/api/engagements/{eid}/doctor")
async def doctor_view(eid: int):
    """The overlay doctor's last diagnosis and the cells held at Excel's values."""
    return await _run(workbench.doctor_view, eid)


@app.get("/api/engagements/{eid}/rows")
async def rows_view(eid: int):
    """The row agents: what they decided for each row this year's value was waiting on."""
    return await _run(workbench.rows_view, eid)


@app.get("/api/engagements/{eid}/overlay/module.py")
async def overlay_module(eid: int):
    path = await _run(workbench.overlay_module, eid)
    return FileResponse(path, media_type="text/x-python; charset=utf-8", filename=f"overlay_e{eid}.py",
                        content_disposition_type="inline")


@app.get("/api/engagements/{eid}/candidates")
async def equity_candidates(eid: int):
    """Overlay cells holding the report's equity value figures, for a person's pick."""
    def go():
        import result
        _, summary = workbench.overlay_session(eid)
        return result.candidates(summary, _head(eid), limit=30)
    return await _run(go)
