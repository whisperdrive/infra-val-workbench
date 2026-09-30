"""The workbench end to end on the synthetic pack (tests/make_pack.py), every model call stubbed:

  run       four files uploaded; the orchestrator reads the report's text, extracts and reviews the key facts,
            confirms the roles on the evidence (the rules' checks and a second opinion that agrees), rebuilds last
            year's overlay in Python (every formula cell as Excel saved it), finds the report's equity value (low
            and high on one row, ex-distribution), rolls forward onto this year's model, and writes the bridge
            (report -> rebuilt -> time value -> cash flows paid -> new forecast -> this year) and the cash-flow chart
            (last year's and this year's undiscounted forecast, the terminal value left out); gpt-sol reviews the run
  gating    a stage that fails isn't run again on the same inputs (only a person's "try again", or new inputs);
            gpt-sol decides an issue once per set of inputs, then a person does
  roles     where the second opinion disagrees, gpt-sol picks between the two on the evidence and the pick is
            checked in code before it's confirmed; a pick that fails the checks goes to a person

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
BEHAVIOUR = {"second_opinion": "agree", "roles_decision": "rules"}


class Reply:
    def __init__(self, out):
        self.output_text, self.usage = json.dumps(out), None


def _section(text: str, head: str, end: str | None) -> str:
    i = text.index(head) + len(head)
    return text[i:text.index(end, i)] if end else text[i:]


def _quote(page: str, needle: str) -> str:
    """The sentence (or table row) on the page holding the needle, copied from the page."""
    for line in page.splitlines():
        if needle in line:
            for part in re.split(r"(?<=\.) (?=[A-Z])", line):
                if needle in part:
                    return part.strip()
    raise AssertionError(f"{needle!r} not on the page")


def fake_facts(doc: str) -> list[dict]:
    pages = {int(n): t for n, t in re.findall(r"<!-- page (\d+) -->\n(.*?)(?=<!-- page \d+ -->|\Z)", doc, re.S)}
    where = lambda needle: next(n for n, t in pages.items() if needle in t)

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
            f("assumption", "discount_rate", "Discount rate", "Discount rate 7.25% - 7.75%", "", "7.25%", "7.75%",
              unit="%", basis="Post-tax nominal WACC"),
            f("assumption", "terminal_growth_rate", "Terminal growth rate", "Terminal growth rate 2.50%", "2.50%", unit="%"),
            f("assumption", "franking_utilisation", "Franking credit utilisation", "Franking credit utilisation 50%", "50%",
              unit="%", basis="Gamma"),
            f("conclusion", "franking_credits_value", "Value of franking credits", "franking credits of A$389.8m",
              "A$389.8m"),
            f("conclusion", "franking_credits_share", "Franking credits, share of equity value", "15.5% of the equity value",
              "15.5%", unit="%"),
            f("conclusion", "terminal_value", "Terminal value", "the terminal value is A$4,968.9m", "A$4,968.9m")]


def fake_create(client, model, input, text=None, max_output_tokens=None, purpose=None, **kw):
    name = text["format"]["name"]
    CALLS.append(name)
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
    if name == "run_review":
        return Reply({"choice": "ok", "reason": "the bridge steps are in proportion", "question": "", "concerns": []})
    if name == "test_decision":
        return Reply({"choice": "pick", "reason": "test", "question": ""})
    raise AssertionError(f"unexpected model call {name}")


def stub_models() -> None:
    llm.client = lambda *a, **k: None
    llm.create = fake_create


# ---- helpers -------------------------------------------------------------------------------------------------------

def upload(eid: int, name: str) -> dict:
    src = PACK / name
    tmp = Path(tempfile.mkdtemp()) / name
    shutil.copy(src, tmp)
    return wb.add_upload(eid, tmp, name, library.sha256_file(tmp))


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

def run_check() -> int:
    e = wb.create("Asset A, FY26")
    eid = e["id"]
    for name in ("AssetA_valuation_report_FY25.pdf", "AssetA_BP25_client_model_Jun25.xlsx",
                 "Alpha_valuation_overlay_FY25.xlsx", "AssetA_BP26_client_model_Jun26.xlsx"):
        upload(eid, name)
    v = wait(eid, lambda v: status(v).get("review") in (*orc.SETTLED, "blocked", "failed") or
             any(s["status"] in ("blocked", "failed") for s in v["stages"]), "the run", 600)
    st = status(v)
    assert all(st[s] in orc.SETTLED for s in ("files", "facts", "roles", "rebuild", "rows", "result", "review")), \
        (st, [(s["stage"], s["note"]) for s in v["stages"]], v["needs"])
    e = wb.get(eid)
    rl = e["roles"]
    assert len(rl) == 4 and all(r["confirmed"] and r["by"] == "orchestrator" and r["evidence"] for r in rl.values()), rl
    names = {w["id"]: w["filename"] for w in e["workbooks"]}
    assert names[rl["prior_overlay"]["id"]].startswith("Alpha_") and names[rl["current_model"]["id"]].startswith("AssetA_BP26")
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
    assert not [n for n in orc.view(eid)["needs"] if n["severity"] == "block"]
    log = orc.history(eid, limit=200)
    assert any(h["stage"] == "roles" and h["event"] == "done" for h in log) and \
        any(h["stage"] == "review" and h["event"] == "decide" for h in log)
    calls = {c: CALLS.count(c) for c in set(CALLS)}
    assert calls.get("report_facts") == 1 and calls.get("facts_review") == 1 and not calls.get("roles_decision"), calls
    print(f"run: ok (facts agreed, roles confirmed on {len(rl['current_model']['evidence']['checks'])} checks and a "
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
    schema = orc._schema("test_decision", ["pick", "escalate"])
    n0 = CALLS.count("test_decision")
    a = orc.ask(eid, "map", "test", "k1", "decide", schema)
    b = orc.ask(eid, "map", "test", "k1", "decide", schema)
    c = orc.ask(eid, "map", "test", "k2", "decide", schema)
    assert a and b is None and c and CALLS.count("test_decision") == n0 + 2, (a, b, c)
    assert any(h["event"] == "escalate" and h["issue"] == "test" for h in orc.history(eid, "map"))
    orc.retry(eid, "map")
    wait(eid, lambda v: status(v)["map"] == "done", "the map")
    print("gating: ok (a failed stage waits for new inputs or a person's try again; gpt-sol decides an issue once per "
          "set of inputs, then escalates)")


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


def main() -> None:
    if not PACK.exists():
        sys.exit("run tests/make_pack.py first")
    sandbox()
    stub_models()
    library.start()
    orc.start()
    eid = run_check()
    gating_check(eid)
    roles_check()


if __name__ == "__main__":
    main()
