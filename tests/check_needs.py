"""Every need the run can raise, held to the page and the export: what it asks of a person (its kind), the card its
"Go" lands on (on the page the need names, listing it or showing it in its own rows), and the words the diagnostics
export may carry. Read from the code itself: the needs the orchestrator builds, and the checks' holds (result.hold), with
the ids built from a figure declared below. No sandbox, nothing run:
    .venv/bin/python tests/check_needs.py
"""
import ast
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "engine"))
import diagnostics  # noqa: E402
import orchestrator as orc  # noqa: E402

HTML = (ROOT / "app" / "index.html").read_text()

# the ids built from a figure (an end, a file, a model, a count): what each can be, or one id standing for the rest
EXPAND = {
    "{key}-{end}": [f"{k}-{e}" for k in ("rate", "growth", "multiple", "franking") for e in ("low", "high")],
    "tv-base-{end}": ["tv-base-low", "tv-base-high"], "fact-{f['id']}": ["fact-7"], "missing-{key}": ["missing-discount_rate"],
    "failed-{job}": ["failed-result"], "held-{i}": ["held-0"], "held-lost-{i}": ["held-lost-0"],
    "held-moved-{i}": ["held-moved-0"], "pick-lost-{i}": ["pick-lost-0"], "pick-moved-{i}": ["pick-moved-0"],
    "review-{i}": ["review-0"], "reconcile-{r['key']}": ["reconcile-pv_forecast"],
    "cf-split-{end}": ["cf-split-low", "cf-split-high"], "cf-sign-{end}": ["cf-sign-low", "cf-sign-high"],
    "cf-flip-{c['cell']}": ["cf-flip-DCF!C4"], "cf-midperiod-{c['cell']}": ["cf-midperiod-DCF!C4"],
    "cf-ondate-{c['cell']}": ["cf-ondate-DCF!C4"], "cf-uncut-{c['cell']}": ["cf-uncut-DCF!C4"],
    "cf-untimed-{c['cell']}": ["cf-untimed-DCF!C4"], "cf-xnpv-{c['cell']}": ["cf-xnpv-DCF!C4"],
    "equity-typed-{e}": ["equity-typed-low", "equity-typed-high"], "equity-unmoved-{e}": ["equity-unmoved-low", "equity-unmoved-high"],
    "errors-{which}": ["errors-current", "errors-prior"], "unsaved-{which}": ["unsaved-current", "unsaved-prior"],
    "rebuild-{e}": ["rebuild-low", "rebuild-high"], "rebuild-error-{e}": ["rebuild-error-low", "rebuild-error-high"],
    "tie-{e}": ["tie-low", "tie-high"], "tie-loose-{e}": ["tie-loose-low", "tie-loose-high"],
}
# the card each need can land on, by page (jumpTo opens the page, then the card)
PAGE_OF = {"filesCard": "workbench", "rolesCard": "workbench", "datesCard": "workbench", "reportCard": "report",
           "tieCard": "rebuild", "inputsCard": "rebuild", "reconcileCard": "rebuild", "equityPick": "rebuild",
           "mismatches": "rebuild", "doctorCard": "rebuild", "bridgeCard": "result", "rateCard": "result",
           "methodsCard": "result", "heldCard": "result", "rowsCard": "result", "flowsCard": "result",
           "compareCard": "result", "linesCard": "result", "termsCard": "result", "scenarioCard": "result"}
# needs a card shows in its own rows, not in its list of needs (a full match): the held inputs' table, the terms and lines
# tables, the inputs' evidence, the reconciliation's ticks, the scenario's selectors, this year's rate, the cells that
# differ, the bridge's step for last year's client model
SHOWN = {"heldCard": r"held-\d+", "termsCard": r"new-terms", "linesCard": r"new-lines",
         "inputsCard": r"(rate|growth|multiple|franking)-(low|high)", "reconcileCard": r"reconcile-\w+",
         "scenarioCard": r"scenario", "rateCard": r"rate-this-year", "mismatches": r"validation", "bridgeCard": r"prior-feed"}
# anchors built from a figure: the element the page gives each (the workbook's slot, the fact's row, the review's point)
# and its page; the element is the need's own subject
BUILT = {"wb-": ('key = (doc ? "doc-" : "wb-") + f.id', "workbench"), "fact-": ('<tr id="fact-${f.id}">', "report"),
         "review-": ("`review-${i}`", "result")}


def _id(node, src) -> str | None:
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        return node.value
    if isinstance(node, ast.JoinedStr):
        return ast.get_source_segment(src, node)[2:-1]
    return None


def needs_built() -> list[dict]:
    """The needs the orchestrator builds: each dict literal with an "id" and a "stage", its severity and its "go"."""
    src = (ROOT / "engine" / "orchestrator.py").read_text()
    out = []
    for node in ast.walk(ast.parse(src)):
        if not isinstance(node, ast.Dict):
            continue
        d = {k.value: v for k, v in zip(node.keys, node.values) if isinstance(k, ast.Constant)}
        if "id" not in d or "stage" not in d or _id(d["id"], src) is None:
            continue  # (a check's finding as a need, _need_of: the holds, below)
        go = d.get("go")
        go = {k.value: _id(v, src) for k, v in zip(go.keys, go.values)} if isinstance(go, ast.Dict) else None
        sev = d.get("severity")
        out.append({"id": _id(d["id"], src), "stage": _id(d["stage"], src) or "result", "go": go, "line": node.lineno,
                    "severity": sev.value if isinstance(sev, ast.Constant) else "?"})
    return out


def holds_built() -> list[dict]:
    """The checks' findings: each hold(summary, id, ...) call, its severity ("block" unless said) and its export id."""
    out = []
    for f in ("result.py", "cashflows.py", "interest.py"):
        src = (ROOT / "engine" / f).read_text()
        for node in ast.walk(ast.parse(src)):
            if not (isinstance(node, ast.Call) and getattr(node.func, "id", getattr(node.func, "attr", None)) == "hold"):
                continue
            kw = {k.arg: k.value for k in node.keywords}
            sev = kw.get("severity")
            out.append({"id": _id(node.args[1], src), "file": f, "line": node.lineno,
                        "severity": sev.value if isinstance(sev, ast.Constant) else "block" if sev is None else "?",
                        "check": kw["check"].value if isinstance(kw.get("check"), ast.Constant) else None})
    return out


def card_lists() -> dict:
    """The needs each card lists, from the page's own table (CARD_NEEDS): {card: [id, or a prefix ending in "-"]}."""
    m = re.search(r"^const CARD_NEEDS = (\{.*?\});", HTML, re.M | re.S)
    assert m, "the page has no CARD_NEEDS table"
    return json.loads(m.group(1))


def lists(pats, i: str) -> bool:
    return any(i.startswith(p) if p.endswith("-") else i == p for p in pats)


def table() -> list[dict]:
    rows = []
    for n in needs_built():
        for i in EXPAND.get(n["id"], [n["id"]]):
            rows.append({**n, "id": i, "pattern": n["id"], "from": f"orchestrator.py:{n['line']}"})
    for h in holds_built():
        for i in EXPAND.get(h["id"], [h["id"]]):
            rows.append({**h, "id": i, "pattern": h["id"], "stage": "result", "go": orc.hold_go(i),
                         "from": f"{h['file']}:{h['line']}"})
    for r in rows:
        d = orc.dress({"id": r["id"], "stage": r["stage"], "go": r["go"]})
        r.update(kind=d["kind"], kind_label=d["kind_label"], at=d["go"])
    return rows


def check() -> None:
    rows, cards = table(), card_lists()
    ids = set(re.findall(r'id="([A-Za-z][\w-]*)"', HTML)) | set(re.findall(r', "([A-Za-z]+Card)"\)', HTML))
    words = diagnostics._vocabulary()
    bad = []
    for r in rows:
        i, at, where = r["id"], r["at"] or {}, f"{r['id']} ({r['from']})"
        if r["kind"] == "note" and r["severity"] != "info":
            bad.append(f"{where}: a {r['severity']} with no kind (KINDS)")
        if r["kind"] == "retry":
            continue  # a step to try again: no card
        a = at.get("anchor") or ""
        built = next((b for b in BUILT if a.startswith(b) and "{" in a), None)
        if built:
            snippet, page = BUILT[built]
            if snippet not in HTML or at.get("step") != page:
                bad.append(f"{where}: lands on {a} ({at.get('step')}), which the page doesn't build there")
            continue
        if a not in ids:
            bad.append(f"{where}: lands on {a or 'nothing'}, which isn't on the page")
        elif PAGE_OF.get(a) != at.get("step"):
            bad.append(f"{where}: lands on {a} on the {at.get('step')} page; it's on the {PAGE_OF.get(a)} page")
        if not lists(cards.get(a, []), i) and not re.fullmatch(SHOWN.get(a, r"(?!)"), i):
            bad.append(f"{where}: lands on {a}, which neither lists it (CARD_NEEDS) nor shows it")
    for h in holds_built():  # the export carries a check's id where it's the app's word (a cell is never one)
        for i in EXPAND.get(h["id"], [h["id"]]):
            if (h["check"] or i) not in words:
                bad.append(f"{i} ({h['file']}:{h['line']}): not a word the diagnostics export may carry")
    every = {r["id"] for r in rows}
    for card, pats in cards.items():
        if card not in ids or card not in PAGE_OF:
            bad.append(f"CARD_NEEDS: {card} isn't a card a need lands on")
        for p in pats:
            hit = [r for r in rows if lists([p], r["id"])]
            if not hit:
                bad.append(f"CARD_NEEDS: {card}'s {p} lists no need")
            bad += [f"CARD_NEEDS: {card} lists {r['id']}, which lands on {(r['at'] or {}).get('anchor')}"
                    for r in hit if (r["at"] or {}).get("anchor") != card]
    assert not bad, "\n".join(bad)
    print(f"needs: ok ({len(every)} needs, {len({r['id'] for r in rows if r.get('file')})} of them the checks' holds: "
          f"each says what it asks of a person and lands on a card on its page that lists it or shows it; "
          f"{len(cards)} cards list only what lands on them; the holds' ids are words the export may carry)")


if __name__ == "__main__":
    if sys.argv[1:] == ["--table"]:
        for r in table():
            print(f"{r['id']:28} {r['severity']:6} {r['kind']:16} {str((r['at'] or {}).get('step')):10} "
                  f"{str((r['at'] or {}).get('anchor')):16} {r['from']}")
        sys.exit()
    check()
