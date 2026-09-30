"""The workbench end to end on the synthetic pack (tests/make_pack.py), every model call stubbed:

  run       four files uploaded; the orchestrator reads the report's text, reads its key tables from their images (the
            discount rate is only in a pasted picture), extracts and reviews the key facts, checks each on the image of
            where it sits (the extraction swaps the rate's low and high, as a scrambled layout would; the image
            reading and the reviewer put it right),
            confirms the roles on the evidence (the rules' checks and a second opinion that agrees), rebuilds last
            year's overlay in Python (every formula cell as Excel saved it), finds the report's equity value (low
            and high on one row, ex-distribution), rolls forward onto this year's model, and writes the bridge
            (report -> rebuilt -> time value -> cash flows paid -> new forecast -> this year) and the cash-flow chart
            (last year's and this year's undiscounted forecast, the terminal value left out); gpt-sol reviews the run
  gating    a stage that fails isn't run again on the same inputs (only a person's "try again", or new inputs);
            gpt-sol decides an issue once per set of inputs, then a person does
  roles     where the second opinion disagrees, gpt-sol picks between the two on the evidence and the pick is
            checked in code before it's confirmed; a pick that fails the checks goes to a person
  inside    the same run with the overlay inside a copy of the client model and the report as slides
  workpaper the Excel workpaper built from the run, in memory: its sheets, the bridge's mid ends at this year's value,
            a row for every financial year of the cash flows, the disclaimer on its summary; none before the bridge
  overview  every engagement at a glance: where it is (a finished one, an empty one), its values
  ranges    a dash between two figures is a range in the fact checks ("7.25% - 7.75%"), not a minus; a bracket
            or a dash after a word still is

    uv run python tests/make_pack.py && uv run python tests/check_workbench.py
"""
import json
import re
import shutil
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "engine"))
PACK = ROOT / "tests" / "pack"

import calllog  # noqa: E402
import lessons  # noqa: E402
import library  # noqa: E402
import llm  # noqa: E402
import ratelimit  # noqa: E402
import usage  # noqa: E402
import workbench as wb  # noqa: E402
import workpaper  # noqa: E402
import orchestrator as orc  # noqa: E402


# ---- a sandbox: every file this run writes goes to a temp folder ----------------------------------------------------

def sandbox() -> Path:
    tmp = Path(tempfile.mkdtemp(prefix="workbench_"))
    wb.OUT, wb.UPLOADS, wb.DB, wb.DOCS = tmp / "out", tmp / "uploads", tmp / "out" / "workbench.db", tmp / "out" / "docs"
    library.OUT, library.UPLOADS, library.REGISTRY = tmp / "out", tmp / "uploads", tmp / "out" / "registry.db"
    usage.DB, calllog.DB, lessons.STORE = tmp / "out" / "registry.db", tmp / "out" / "calls.db", tmp / "out" / "lessons.json"
    (tmp / "out").mkdir(parents=True)
    ratelimit.load_capacities = lambda *a, **k: None
    return tmp


# ---- the models, stubbed -------------------------------------------------------------------------------------------

CALLS: list[str] = []
PROMPTS: dict[str, str] = {}  # the last prompt of each kind of call
BEHAVIOUR = {"second_opinion": "agree", "roles_decision": "rules", "verdict": "image", "verdict_figures": {}}
TABLES: dict[int, str] = {}  # page -> the table on it, as a reader of its image would transcribe it
TRUTH = {"equity_value": {"value": "A$2,507.9m", "low": "A$2,345.5m", "high": "A$2,670.3m"},
         "equity_value_cum": {"value": "A$2,532.9m"}, "valuation_date": {"value": "30 June 2025"},
         "discount_rate": {"low": "7.25%", "high": "7.75%"}, "terminal_growth_rate": {"value": "2.50%"},
         "franking_utilisation": {"value": "50%"}, "franking_credits_value": {"value": "A$389.8m"},
         "franking_credits_share": {"value": "15.5%"}, "terminal_value": {"value": "A$4,968.9m"},
         "pv_forecast": {"value": "A$1,819.3m"}, "pv_terminal_value": {"value": "A$1,173.7m"}}


class Reply:
    def __init__(self, out):
        self.output_text, self.usage = json.dumps(out), None


def _section(text: str, head: str, end: str | None) -> str:
    i = text.index(head) + len(head)
    return text[i:text.index(end, i)] if end else text[i:]


def _quote(page: str, needle: str) -> str:
    """The sentence (or table row, its cells separated by spaces) on the page holding the needle."""
    for line in page.splitlines():
        line = re.sub(r"\s+", " ", re.sub(r"\s*\|\s*", " ", line)).strip()
        if needle in line:
            for part in re.split(r"(?<=\.) (?=[A-Z])", line):
                if needle in part:
                    return part.strip()
    raise AssertionError(f"{needle!r} not on the page")


def fake_facts(doc: str) -> list[dict]:
    pages = {int(n): t for n, t in re.findall(r"<!-- page (\d+) -->\n(.*?)(?=<!-- page \d+ -->|\Z)", doc, re.S)}
    flat = lambda t: re.sub(r"\s+", " ", re.sub(r"\s*\|\s*", " ", t))
    where = lambda needle: next(n for n, t in pages.items() if needle in flat(t))

    def f(cat, key, label, needle, value_text="", low="", high="", unit="A$m", basis="", value=None):
        n = where(needle)
        return {"category": cat, "key": key, "label": label, "value_text": value_text, "low_text": low, "high_text": high,
                "value": value, "unit": unit, "basis": basis, "page": n, "quote": _quote(pages[n], needle)}
    return [f("identity", "target_name", "Target", "equity in Asset A Pty Ltd", "Asset A Pty Ltd", unit="text"),
            f("identity", "client", "Client", "engaged by Holdco A Pty Ltd", "Holdco A Pty Ltd", unit="text"),
            f("identity", "valuation_date", "Valuation date", "as at 30 June 2025 (the Valuation Date)", "30 June 2025",
              unit="date", value=20250630),
            f("conclusion", "equity_value", "Equity value", "A$2,345.5m", "A$2,507.9m", "A$2,345.5m", "A$2,670.3m",
              basis="ex-distribution"),
            f("conclusion", "equity_value_cum", "Equity value (cum-distribution)", "the midpoint would be A$2,532.9m",
              "A$2,532.9m", basis="cum-distribution"),
            # the low and the high the wrong way round, as a scrambled text layer can give them
            f("assumption", "discount_rate", "Discount rate", "Discount rate 7.25% - 7.75%", "", "7.75%", "7.25%",
              unit="%", basis="Post-tax nominal WACC"),
            f("assumption", "terminal_growth_rate", "Terminal growth rate", "Terminal growth rate 2.50%", "2.50%", unit="%"),
            f("assumption", "franking_utilisation", "Franking credit utilisation", "Franking credit utilisation 50%", "50%",
              unit="%", basis="Gamma"),
            f("conclusion", "franking_credits_value", "Value of franking credits", "franking credits of A$389.8m",
              "A$389.8m"),
            f("conclusion", "franking_credits_share", "Franking credits, share of equity value", "15.5% of the equity value",
              "15.5%", unit="%"),
            f("conclusion", "terminal_value", "Terminal value", "the terminal value is A$4,968.9m", "A$4,968.9m"),
            f("conclusion", "pv_forecast", "PV of the forecast", "A$1,819.3m", "A$1,819.3m"),
            f("conclusion", "pv_terminal_value", "PV of the terminal value", "A$1,173.7m", "A$1,173.7m")]


def fake_create(client, model, input, text=None, max_output_tokens=None, purpose=None, **kw):
    name = text["format"]["name"]
    CALLS.append(name)
    if isinstance(input, list):  # a call with an image: its prompt is the text part
        input = next(c["text"] for c in input[0]["content"] if c["type"] == "input_text")
    PROMPTS[name] = input
    if name == "table_read":
        page = int(re.search(r"(?:page|slide) (\d+)", input)[1])
        md = TABLES.get(page, "")
        return Reply({"is_table": bool(md), "title": "", "markdown": md, "description": "" if md else "a picture"})
    if name == "visual_read":
        items = json.loads(_section(input, "Items:\n", None))
        return Reply({"reads": [{"id": x["id"], "found": x["key"] in TRUTH, "column": "", "row": x["item"],
                                 **{k: TRUTH.get(x["key"], {}).get(k, "") for k in ("value", "low", "high")}} for x in items]})
    if name == "visual_verdict":
        items = json.loads(_section(input, "Items:\n", None))
        fig = BEHAVIOUR["verdict_figures"]
        return Reply({"verdicts": [{"id": x["id"], "choice": BEHAVIOUR["verdict"], "value": fig.get("value", ""),
                                    "low": fig.get("low", ""), "high": fig.get("high", ""),
                                    "reason": "the image shows 7.25% under Low and 7.75% under High"} for x in items]})
    if name == "identity":
        raise RuntimeError("no model in the test: identify falls back to the cells it found")
    if name == "report_facts":
        return Reply({"facts": fake_facts(_section(input, "Document:\n", None)), "notes": ""})
    if name == "facts_review":
        facts = json.loads(_section(input, "Facts:\n", "\n\nDocument:"))
        return Reply({"reviews": [{"id": f["id"], "verdict": "accept", "reason": "as the report states it", "value_text": "",
                                   "low_text": "", "high_text": "", "basis": "", "page": None, "quote": ""} for f in facts],
                      "missing": [], "summary": "all accepted"})
    if name == "roles_second_opinion":
        sug = json.loads(_section(input, "Suggested assignment:\n", None))
        if BEHAVIOUR["second_opinion"] == "agree":
            return Reply({"agree": True, **{k: sug.get(k, "") for k in ("prior_report", "prior_model", "prior_overlay",
                                                                         "current_model")},
                          "overlay_sheets": sug.get("overlay_sheets") or [], "confidence": "high", "reasons": ["as suggested"]})
        return Reply({"agree": False, **{k: sug.get(k, "") for k in ("prior_report", "prior_model", "prior_overlay")},
                      "current_model": "no such file.xlsx", "overlay_sheets": sug.get("overlay_sheets") or [],
                      "confidence": "medium", "reasons": ["a different current model"]})
    if name == "roles_decision":
        return Reply({"choice": BEHAVIOUR["roles_decision"], "reason": "the dates and links support it",
                      "question": "Which file is this year's client model?"})
    if name == "run_review":  # one point, about two years the chart has and one it doesn't, and a bridge step
        return Reply({"choice": "concerns", "reason": "one note", "question": "", "concerns": [
            {"title": "Terminal-year cash flow", "detail": "the new final year's cash flow jumps", "severity": "info",
             "years": ["FY45", "FY46", "FY99"], "step": "forecast"}]})
    if name == "test_decision":
        return Reply({"choice": "pick", "reason": "test", "question": ""})
    raise AssertionError(f"unexpected model call {name}")


def stub_models() -> None:
    llm.client = lambda *a, **k: None
    llm.create = fake_create
    import docingest  # what a reader of each page's table image sees: the slides' own tables, which are the same
    slides = docingest.process(str(PACK / "AssetA_valuation_report_FY25.pptx"), tempfile.mkdtemp(), read=False)
    TABLES.update({t["page"]: t["markdown"] for t in slides["tables"]})


# ---- helpers -------------------------------------------------------------------------------------------------------

def upload(eid: int, name: str, role: str | None = None) -> dict:
    src = PACK / name
    tmp = Path(tempfile.mkdtemp()) / name
    shutil.copy(src, tmp)
    return wb.add_upload(eid, tmp, name, library.sha256_file(tmp), role)


def wait(eid: int, done, what: str, timeout: float = 300) -> dict:
    t0 = time.time()
    while time.time() - t0 < timeout:
        v = orc.view(eid)
        if done(v):
            return v
        time.sleep(0.5)
    v = orc.view(eid)
    raise AssertionError(f"timed out waiting for {what}: " + json.dumps({s["stage"]: (s["status"], s["note"]) for s in v["stages"]},
                                                                        indent=1) + "\n" + "\n".join(
        f"{h['stage']}: {h['event']}: {h['text']}" for h in v["log"][:15]))


def status(v: dict) -> dict:
    return {s["stage"]: s["status"] for s in v["stages"]}


# ---- the checks ----------------------------------------------------------------------------------------------------

PACK_A = ("AssetA_valuation_report_FY25.pdf", "AssetA_BP25_client_model_Jun25.xlsx", "Alpha_valuation_overlay_FY25.xlsx",
          "AssetA_BP26_client_model_Jun26.xlsx")
PACK_B = ("AssetA_valuation_report_FY25.pptx", "AssetA_BP25_with_overlay.xlsx", "AssetA_BP26_client_model_Jun26.xlsx")


def run_check(files=PACK_A, name="Asset A, FY26") -> int:
    e = wb.create(name)
    eid = e["id"]
    n0 = dict((c, CALLS.count(c)) for c in set(CALLS))
    for f in files:
        upload(eid, f)
    v = wait(eid, lambda v: status(v).get("review") in (*orc.SETTLED, "blocked", "failed") or
             any(s["status"] in ("blocked", "failed") for s in v["stages"]), "the run", 600)
    st = status(v)
    assert all(st[s] in orc.SETTLED for s in ("files", "facts", "roles", "rebuild", "rows", "result", "review")), \
        (st, [(s["stage"], s["note"]) for s in v["stages"]], v["needs"])
    e = wb.get(eid)
    rl = e["roles"]
    assert len(rl) == 4 and all(r["confirmed"] and r["by"] == "orchestrator" and r["evidence"] for r in rl.values()), rl
    names = {w["id"]: w["filename"] for w in e["workbooks"]}
    assert names[rl["prior_overlay"]["id"]] == files[2 if files == PACK_A else 1] and \
        names[rl["current_model"]["id"]].startswith("AssetA_BP26"), (names, rl)
    if files == PACK_B:
        assert rl["prior_model"]["id"] == rl["prior_overlay"]["id"] and \
            {"Val_Inputs", "DCF", "Summary"} <= set(rl["prior_overlay"]["sheets"]) and \
            "CashFlow" in rl["prior_model"]["sheets"], rl
    approved = {f["key"] for f in e["facts"] if f["status"] == "approved" and f["decided_by"] == "agents"}
    assert {"equity_value", "discount_rate", "valuation_date", "franking_utilisation"} <= approved, approved
    res = e["result"]
    assert res["head"]["basis"] == "ex" and abs(res["head"]["mid"] - 2507.9) < 1e-9, res["head"]
    assert res["tie"]["low"]["ok"] and res["tie"]["high"]["ok"] and res["tie"]["low"]["rebuilt_ok"], res["tie"]
    v = res["values"]
    assert abs(v["rebuilt"]["mid"] - 2507.876) < 0.01 and v["this_year"]["mid"] > v["rebuilt"]["mid"], v
    steps = res["bridges"]["mid"]["steps"]
    keys = [s["key"] for s in steps]
    assert keys == ["report", "rounding", "prior_feed", "rebuilt", "time", "cash", "forecast", "rate", "this_year"], keys
    moves = sum(s["value"] for s in steps if not s.get("total") and s["key"] not in ("rounding", "prior_feed"))
    assert abs(steps[3]["value"] + moves - steps[-1]["value"]) < 1e-6, "the steps add up"
    ch = res["chart"]
    assert ch["left_out"] == ["Terminal value"] and len(ch["series"]["last_year"]) == 20 and \
        len(ch["series"]["this_year"]) == 20, ch
    assert list(ch["series"]["last_year"])[0] == "FY26" and list(ch["series"]["this_year"])[0] == "FY27"
    rate = next(f for f in e["facts"] if f["key"] == "discount_rate")
    if files == PACK_A:  # read from the picture, the swap put right on the image, and the picture's table read
        assert (rate["low_text"], rate["high_text"]) == ("7.25%", "7.75%") and rate["visual"]["status"] == "corrected", rate
        assert rate["status"] == "approved", rate
        seen = {f["key"]: (f.get("visual") or {}).get("status") for f in e["facts"]}
        assert seen["equity_value"] == "confirmed" and seen["terminal_value"] == "confirmed", seen
        doc = wb.document(rate["document_id"])["doc"]
        pic = next(t for t in doc["tables"] if t["source"] == "picture")
        assert pic["status"] == "verified" and "Discount rate" in pic["markdown"], pic
    else:
        assert rate["visual"]["status"] == "native", rate
    rec = {r["key"]: r for r in res["reconcile"]["rows"]}
    for key in ("terminal_value", "pv_forecast", "pv_terminal_value", "franking_credits_value", "franking_credits_share"):
        assert rec[key].get("ok") is True, (key, rec[key])
    this = res["reconcile"]["this_year"]["mid"]
    assert this["pv_tv"] and this["pv_forecast"] and abs(this["pv"] - (this["pv_tv"] + this["pv_forecast"])) < 1e-6, this
    rates = {e: next(a["rate_source"] for a in res["assumptions"][e] if a.get("parts")) for e in ("low", "high")}
    assert rates == {"low": "Val_Inputs!C5", "high": "Val_Inputs!E5"}, rates
    # each end's rate sighted, sourced by following the factors' formulas, and checked again against the report
    ins = res["inputs"]
    rc = ins["rate"]["ends"]
    assert ins["rate"]["ok"] and [(rc[e]["cell"], rc[e]["sourced"], rc[e]["ties"], rc[e]["report"], rc[e]["sighted"]["heading"],
                                   rc[e]["sighted"]["input"]) for e in ("low", "high")] == \
        [("Val_Inputs!C5", True, True, "7.75%", "Low", True), ("Val_Inputs!E5", True, True, "7.25%", "High", True)], rc
    # the terminal growth rate: the cell the terminal value reads as g, the terminal value at each end's own rate
    gr = ins["growth"]["ends"]
    assert ins["growth"]["ok"] and all(gr[e]["cell"] == "Val_Inputs!C6" and gr[e]["ties"] and
                                       all(c["ok"] for c in gr[e]["checks"]) for e in ("low", "high")), gr
    # franking credit utilisation: the fraction the franking credits read; at nil, the equity value falls by exactly
    # the value of franking credits in the Python overlay, and is back where it was after
    fk = ins["franking"]["ends"]
    assert ins["franking"]["ok"] and all(fk[e]["cell"] == "Val_Inputs!C9" and fk[e]["ties"] and fk[e]["rerun"]["ok"] and
                                         fk[e]["rerun"]["restored"] for e in ("low", "high")), fk
    assert abs(fk["low"]["rerun"]["drop"] - 381.1203152392057) < 1e-6, fk["low"]["rerun"]
    # the date line: the report's date the anchor, the overlay's the cell its discount factors read
    dl = {d["step"]: d for d in wb.get(eid)["dates"]}
    assert (dl["The report"]["date"], dl["Last year's overlay"]["date"], dl["Last year's overlay"]["where"],
            dl["Last year's overlay"]["ok"], dl["Last year's client model"]["ok"], dl["This year's client model"]["date"],
            dl["This year's client model"]["ok"], dl["The roll-forward runs to"]["date"]) == \
        ("2025-06-30", "2025-06-30", "Val_Inputs!C4", True, True, "2026-06-30", True, "2026-06-30"), dl
    view = orc.view(eid)
    assert not [n for n in view["needs"] if n["severity"] == "block"]
    # every need says what it asks of a person and lands on a card; a review point names its years and step, checked
    # against the run (a year the chart doesn't have is dropped)
    assert all(n.get("kind") and (n["kind"] == "retry" or n["go"].get("anchor")) for n in view["needs"]), view["needs"]
    rp = next(n for n in view["needs"] if n["id"] == "review-0")
    assert (rp["kind"], rp["years"], rp["step"], rp["go"]["anchor"]) == ("review-point", ["FY45", "FY46"], "forecast", "review-0"), rp
    # how long each stage took, so the page can say how long is left
    took = {x["stage"]: x["typical_secs"] for x in view["stages"]}
    assert all(took[k] is not None for k in ("facts", "roles", "rebuild", "rows", "result", "review")), took
    log = orc.history(eid, limit=200)
    assert any(h["stage"] == "roles" and h["event"] == "done" for h in log) and \
        any(h["stage"] == "review" and h["event"] == "decide" for h in log)
    calls = {c: CALLS.count(c) - n0.get(c, 0) for c in set(CALLS)}
    assert calls.get("report_facts") == 1 and calls.get("facts_review") == 1 and not calls.get("roles_decision"), calls
    if files == PACK_A:
        assert calls.get("visual_read") == 3 and calls.get("visual_verdict") == 1, calls
        # what the report says around the figures reached every model that read or judged one
        for kind in ("visual_read", "visual_verdict", "table_read"):
            assert "Dear Directors" in PROMPTS[kind] and "context only" in PROMPTS[kind], kind
        assert "Scope of our work" in PROMPTS["visual_verdict"] and "Established so far" in PROMPTS["visual_verdict"]
    print(f"{'run' if files == PACK_A else 'inside'}: ok (the terminal value, the PV split and the franking credits "
          f"reconcile to the report; facts agreed, roles confirmed on {len(rl['current_model']['evidence']['checks'])} checks and a "
          f"second opinion; {res['where']['low']} / {res['where']['high']} tie; mid {v['report']['mid']:,.1f} -> "
          f"{v['this_year']['mid']:,.1f}; chart {len(ch['series']['last_year'])} + {len(ch['series']['this_year'])} years; "
          f"model calls {calls})")
    return eid


def gating_check(eid: int) -> None:
    runs = []

    def broken(eid_, key):
        runs.append(key)
        raise RuntimeError("broken on purpose")
    keep = orc.STAGE_JOBS["map"]
    orc.STAGE_JOBS["map"] = broken
    try:
        orc.retry(eid, "map")
        wait(eid, lambda v: status(v)["map"] == "failed", "the map to fail")
        for _ in range(4):
            orc.tick(eid)
            time.sleep(0.3)
        assert len(runs) == 1, f"a failed stage ran again on the same inputs: {len(runs)}"
        assert any(n["id"] == "failed-map" for n in orc.view(eid)["needs"])
        orc.retry(eid, "map")
        wait(eid, lambda v: len(runs) == 2 and status(v)["map"] == "failed", "the person's try again")
    finally:
        orc.STAGE_JOBS["map"] = keep
    # the roles run again and come out the same: the rebuild and what follows keep their outcome, without running again
    before = {n: orc.stage(eid, n)["started_at"] for n in ("rebuild", "rows", "result")}
    orc.retry(eid, "roles")
    wait(eid, lambda v: all(status(v)[n] in orc.SETTLED for n in ("roles", "rebuild", "rows", "result")), "the rerun")
    after = {n: orc.stage(eid, n)["started_at"] for n in ("rebuild", "rows", "result")}
    assert before == after, (before, after)
    schema = orc._schema("test_decision", ["pick", "escalate"])
    n0 = CALLS.count("test_decision")
    a = orc.ask(eid, "map", "test", "k1", "decide", schema)
    b = orc.ask(eid, "map", "test", "k1", "decide", schema)
    c = orc.ask(eid, "map", "test", "k2", "decide", schema)
    assert a and b is None and c and CALLS.count("test_decision") == n0 + 2, (a, b, c)
    assert any(h["event"] == "escalate" and h["issue"] == "test" for h in orc.history(eid, "map"))
    orc.retry(eid, "map")
    wait(eid, lambda v: status(v)["map"] == "done", "the map")
    print("gating: ok (a failed stage waits for new inputs or a person's try again; a stage the rerun before it didn't "
          "change keeps its outcome; gpt-sol decides an issue once per set of inputs, then escalates)")


def roles_check() -> None:
    BEHAVIOUR.update(second_opinion="differ", roles_decision="rules")
    e = wb.create("Asset A, FY26 (roles)")
    eid = e["id"]
    for name in ("AssetA_valuation_report_FY25.pdf", "AssetA_BP25_client_model_Jun25.xlsx",
                 "Alpha_valuation_overlay_FY25.xlsx", "AssetA_BP26_client_model_Jun26.xlsx"):
        upload(eid, name)
    wait(eid, lambda v: status(v)["roles"] in (*orc.SETTLED, "blocked", "failed"), "the roles")
    rl = wb.roles(eid)
    assert all(r["by"] == "orchestrator" and "decided_by" in (r["evidence"] or {}) for r in rl.values()), rl
    assert any(h["event"] == "verify" for h in orc.history(eid, "roles"))
    # gpt-sol takes the second opinion, which names a file that isn't here: the checks in code stop it
    BEHAVIOUR.update(roles_decision="second_opinion")
    with wb._lock, wb._conn() as db:
        db.execute("DELETE FROM roles WHERE engagement_id=?", (eid,))
    orc.retry(eid, "roles")
    v = wait(eid, lambda v: status(v)["roles"] == "blocked", "the roles to go to a person")
    need = next(n for n in v["needs"] if n["id"] == "roles")
    assert need["severity"] == "block" and "Which file" in need["detail"], need
    BEHAVIOUR.update(second_opinion="agree", roles_decision="rules")
    print("roles: ok (a disagreement is decided by gpt-sol on the evidence and checked in code; a pick that fails the "
          "checks goes to a person)")


def escalate_check() -> None:
    """Where the image reading and the reviewer agree with neither each other nor the extraction, the fact goes to a
    person: not approved, the agents' status escalated, on the needs-you list with the image."""
    BEHAVIOUR.update(verdict="neither", verdict_figures={"low": "9.00%", "high": "9.50%"})
    try:
        e = wb.create("Asset A, FY26 (the image disagrees)")
        eid = e["id"]
        for f in PACK_A:
            upload(eid, f)
        v = wait(eid, lambda v: status(v)["facts"] in (*orc.SETTLED, "failed"), "the facts")
        rate = next(f for f in wb.facts(eid) if f["key"] == "discount_rate")
        assert rate["status"] == "pending" and rate["agent"]["status"] == "escalated" and \
            rate["visual"]["status"] == "escalated" and rate["agent"]["open"]["image"].startswith("tables/"), rate
        need = next(n for n in v["needs"] if n["id"] == f"fact-{rate['id']}")
        assert need["severity"] == "check" and "on the image" in need["detail"], need
    finally:
        BEHAVIOUR.update(verdict="image", verdict_figures={})
    print("escalate: ok (where the reads of the image agree with nothing, the fact goes to a person, with the image)")


def place_check() -> None:
    """Files placed in their roles by hand while uploading stay the person's; the orchestrator places the rest (last
    year's client model here) and confirms it on the evidence. A file of the wrong kind for its role is refused before
    anything is stored."""
    e = wb.create("Asset A, FY26 (placed by hand)")
    eid = e["id"]
    bad = Path(tempfile.mkdtemp()) / PACK_A[0]
    shutil.copy(PACK / PACK_A[0], bad)
    try:
        wb.add_upload(eid, bad, PACK_A[0], library.sha256_file(bad), "current_model")
        raise AssertionError("a report was taken as this year's client model")
    except ValueError as ex:
        assert "a model goes in as XLSX or XLSM" in str(ex) and not bad.exists() and not wb.documents(eid), ex
    for name, role in zip(PACK_A, ("prior_report", None, "prior_overlay", "current_model")):
        upload(eid, name, role)
    wait(eid, lambda v: status(v)["roles"] in (*orc.SETTLED, "blocked", "failed"), "the roles")
    rl = wb.roles(eid)
    assert {k: (r["confirmed"], r["by"]) for k, r in rl.items()} == {
        "prior_report": (True, "you"), "prior_overlay": (True, "you"), "current_model": (True, "you"),
        "prior_model": (True, "orchestrator")}, rl
    print("place: ok (files placed by hand while uploading stay yours; the orchestrator places the rest; a file of the "
          "wrong kind for its role is refused before it's stored)")


def overview_check(eid: int) -> None:
    """The list of every engagement: where each is, and the value the result page shows."""
    e = wb.create("Asset A, nothing yet")
    try:
        ov = {x["id"]: x for x in orc.overview()}
        x, empty = ov[eid], ov[e["id"]]
        want = wb.get(eid)["result"]["values"]
        assert x["state"] == "finished" and x["values"] == want and x["units"] and x["basis"] == "ex" and \
            x["valuation_date"], x
        assert x["needs"] == len([n for n in wb.get(eid)["run"]["needs"] if n["severity"] != "info"]), x
        assert empty["state"] == "empty" and empty["values"] is None and empty["files"] == 0, empty
    finally:
        wb.delete(e["id"])
    print("overview: ok (a finished engagement with its values and what's for a person; an empty one as empty)")


def workpaper_check(eid: int) -> None:
    import io
    import openpyxl
    book = openpyxl.load_workbook(io.BytesIO(workpaper.build(eid)), data_only=True)
    want = ["Summary", "Bridge", "Cash flows", "Inputs", "Reconciliation", "Key facts", "Files and roles", "Review", "Run log"]
    assert book.sheetnames == want, book.sheetnames
    cells = lambda name: [[c for c in row] for row in book[name].iter_rows(values_only=True)]
    res = wb.get(eid)["result"]
    tot = [r for r in cells("Bridge") if r[0] and str(r[0]).startswith("This year, rolled forward")]
    assert tot and abs(tot[-1][2] - res["values"]["this_year"]["mid"]) < 0.05, (tot, res["values"]["this_year"])
    assert not any(re.search(r"\b\d{4}-\d{2}-\d{2}\b", str(c)) for r in cells("Bridge") + cells("Summary") for c in r if c), \
        "a date left as ISO"
    fy = [r for r in cells("Cash flows") if isinstance(r[0], str) and re.fullmatch(r"FY\d{2}", r[0])]
    assert len(fy) == len(res["chart"]["years"]) == 21, len(fy)
    assert any(c and str(c).startswith(workpaper.DISCLAIMER[:40]) for r in cells("Summary") for c in r), "no disclaimer"
    rate = [r for r in cells("Inputs") if r[0] and str(r[0]).startswith("Low (at the higher rate)")]
    assert rate and abs(rate[0][1] - res["inputs"]["rate"]["ends"]["low"]["value"]) < 1e-12 and "!" in rate[0][3], rate
    e = wb.create("Asset A, nothing yet")
    try:
        assert not workpaper.ready(wb.get(e["id"]))
        workpaper.build(e["id"])
        raise AssertionError("a workpaper before the bridge")
    except ValueError:
        pass
    finally:
        wb.delete(e["id"])
    assert workpaper.build(10 ** 6) is None
    print("workpaper: ok (9 sheets; the bridge's mid ends at this year's value; 21 financial years; the disclaimer; "
          "the rate's figure then its cell; none before the bridge)")


def dates_check() -> None:
    """The image shows a different year from the text layer's valuation date, with the same day: the date is compared
    whole, so it's a difference (its first number, the day, alone would confirm it), and two agreeing reads of the
    image correct it, its YYYYMMDD with it."""
    TRUTH["valuation_date"]["value"] = "30 June 2024"
    try:
        e = wb.create("Asset A, FY26 (the image dates it a year earlier)")
        eid = e["id"]
        for f in PACK_A:
            upload(eid, f)
        wait(eid, lambda v: status(v)["facts"] in (*orc.SETTLED, "failed"), "the facts")
        vd = next(f for f in wb.facts(eid) if f["key"] == "valuation_date")
        assert vd["visual"]["status"] == "corrected" and vd["value_text"] == "30 June 2024" and \
            int(vd["value"]) == 20240630, vd
    finally:
        TRUTH["valuation_date"]["value"] = "30 June 2025"
    print("dates: ok (a date on its image is compared whole, day, month and year; a correction carries its YYYYMMDD)")


def delete_check(first: int) -> None:
    """Deleting an engagement takes everything worked out for it: its reports, rows, overlay folder and call log;
    a model another engagement uses stays until the last of them goes. Then the same files, uploaded again, run from
    the start (every file read afresh) to the same value."""
    import calllog
    want = wb.get(first)["result"]["values"]["this_year"]["mid"]
    eids = [e["id"] for e in wb.all_engagements()]
    files = [library.get(f, full=True) for f in {r["file_id"] for r in wb._q("SELECT file_id FROM eng_files")}]
    docs = [d for e in eids for d in wb._q("SELECT out_dir, source_path FROM documents WHERE engagement_id=?", e)]
    kept_once = False
    for eid in eids:
        t0 = time.time()
        while True:  # the checks before may have left an engagement still running: deleting waits for it
            try:
                r = wb.delete(eid)
                break
            except ValueError as e:
                assert "wait for" in str(e) and time.time() - t0 < 300, e
                time.sleep(0.5)
        kept_once |= bool(r["kept"])
        # every table but stages, which the orchestrator's loop may be writing for it that very moment (an id a
        # new engagement then doesn't take)
        assert wb.get(eid) is None and not (wb.OUT / "overlays" / f"e{eid}").exists() and not any(
            wb._q(f"SELECT 1 FROM {t} WHERE engagement_id=?", eid) for t in ("runlog", "facts", "roles", "eng_files", "documents"))
        with calllog._conn() as db:
            assert not db.execute("SELECT 1 FROM calls WHERE engagement=?", (eid,)).fetchone(), eid
    assert kept_once, "a model several engagements use was deleted with the first of them"
    assert not any(library.get(f["id"]) for f in files) and not any(Path(f["out_dir"]).exists() or
                                                                    Path(f["source_path"]).exists() for f in files)
    assert not any(Path(d["out_dir"]).exists() or Path(d["source_path"]).exists() for d in docs)
    t_gone = time.time()
    again = run_check(PACK_A, "Asset A, FY26 (again, from the start)")
    got = wb.get(again)["result"]["values"]["this_year"]["mid"]
    assert abs(got - want) < 1e-6 and all(w["processed_at"] > t_gone for w in wb.workbooks(again)), (got, want)
    print(f"delete: ok ({len(eids)} engagements deleted with what was worked out for them, a shared model kept until "
          f"the last; the same files run again from the start to {got:,.1f})")


def ranges_check() -> None:
    import keyfacts
    import context
    import docingest
    for kind in ("pdf", "pptx"):
        d = docingest.process(str(PACK / f"AssetA_valuation_report_FY25.{kind}"), tempfile.mkdtemp(), read=False)
        c = context.find(d["markdown"])
        assert c["letter"]["pages"] == [2] and (kind == "pptx" or c["scope"]["pages"] == [2]), (kind, c)
    n = lambda t: keyfacts.numbers(keyfacts._unrange(t))
    assert n("Discount rate 7.25% - 7.75% Post-tax") == ["7.25%", "7.75%"]
    assert n("7.25%–7.75%") == ["7.25%", "7.75%"]
    assert n("Less: net debt (850.0) (850.0)") == ["-850.0", "-850.0"]
    assert n("Net debt at valuation - 850.0") == ["-850.0"], n("Net debt at valuation - 850.0")
    got = n("FY25 - 30.0")
    assert got == ["25", "30.0"], got  # a year then a figure reads as a range: the check only looks for the fact's own numbers
    import visual
    from datetime import date
    for t in ("as at 30 June 2025 (the", "30 Jun 25", "30-Jun-25", "June 30, 2025", "30/06/2025", "2025-06-30",
              "30th of June 2025"):
        assert keyfacts.date_of(t) == date(2025, 6, 30), t
    assert keyfacts.date_of("30 and 2025") is None and keyfacts.date_of("FY25 7.25%") is None
    vd = {"key": "valuation_date", "unit": "date", "value_text": "30 June 2025"}
    assert visual._same(vd, {"value": "30/06/2025"}) and not visual._same(vd, {"value": "30 September 2025"}) and \
        not visual._same(vd, {"value": "30 June 2024"}) and not visual._same(vd, {"value": ""})
    got = keyfacts.settle_value({**vd, "value": 20240630})["value"]
    assert got == 20250630, got  # the model's YYYYMMDD gives way to the date in the text
    chk = keyfacts.check({**vd, "value_text": "30 July 2025", "value": 20250730, "page": 2,
                          "quote": "as at 30 June 2025 (the"}, {2: "as at 30 June 2025 (the"})
    assert not chk["ok"] and "30 July 2025 is not a date in the quote" in [i["text"] for i in chk["items"]], chk
    print("ranges: ok (a dash between figures is a range; a bracket, or a dash after a word, is a minus; the letter and "
          "the scope are found; a date is read and checked whole)")


def upgrade_check() -> None:
    """A database from an earlier version gets the columns added since (a Windows install upgrades in place)."""
    import sqlite3
    db_path = Path(tempfile.mkdtemp()) / "old.db"
    with sqlite3.connect(db_path) as old:
        old.execute("CREATE TABLE facts(id INTEGER PRIMARY KEY, engagement_id INT, key TEXT)")
    keep = wb.DB
    wb.DB = db_path
    try:
        cols = {r[1] for r in wb._conn().execute("PRAGMA table_info(facts)")}
    finally:
        wb.DB = keep
    assert {"visual_json", "agent_json", "status"} <= cols, cols
    print("upgrade: ok (an older database gets the columns added since)")


def main() -> None:
    ranges_check()
    upgrade_check()
    if not PACK.exists():
        sys.exit("run tests/make_pack.py first")
    sandbox()
    stub_models()
    library.start()
    orc.start()
    eid = run_check()
    workpaper_check(eid)
    overview_check(eid)
    gating_check(eid)
    roles_check()
    escalate_check()
    dates_check()
    place_check()
    other = run_check(PACK_B, "Asset A, FY26 (overlay inside)")
    a, b = (wb.get(x)["result"]["values"]["this_year"]["mid"] for x in (eid, other))
    assert abs(a - b) < 1e-6, f"the same files give different values in the two layouts: {a} vs {b}"
    print("layouts: ok (the overlay standalone and inside the client model give the same value this year: the roll "
          "moves the valuation date the discountings read)")
    delete_check(eid)


if __name__ == "__main__":
    main()
