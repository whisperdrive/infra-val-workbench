"""Synthetic pairs of client models with known answers: last year's model and this year's, changed one way each (the
ways real models change between valuations), and which of this year's rows is each of last year's. The row tools are
measured on them: how many rows each settles rightly, leaves open for the models, or gets wrong.

    pairs(out_dir) -> [{"name", "last", "this", "truth": {(sheet, row): (sheet, row) or None}}]
    measure(pairs) -> {name: {"right", "open", "wrong", "rows": [...]}}

The changes: revised figures only; rows renamed; a downside case inserted above the base; the sheet renamed; a
prior-forecast block (last year's figures, by formulas) beside renamed rows; rows reordered in their block; a row
dropped; and, from the second review, copies of the block whose headings don't name last year's (the base renamed,
no headings, a P90 above a P50, the base moved to its own sheet), a rebuilt model with a low case beside a revised
central one, and two rows merged. KNOWN_WRONG lists the kinds the tools still get wrong, until they're fixed. All
names and figures are made up.
"""
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "engine"))
import xlsxwriter  # noqa: E402
from xlsxwriter.utility import xl_col_to_name as COL  # noqa: E402

import build_map  # noqa: E402

ENDS = [date(2026 + k, 6, 30) for k in range(8)]
# kinds of change the row tools still settle wrongly (the second review, 3 October 2026; docs/hardening.md S1, S9):
# taken off as each is fixed, so the measure stays honest and a new wrong row anywhere else fails the check
KNOWN_WRONG = {"rebuilt, low case beside a revised central", "costs and tax merged"}
# kinds where leaving the rows open (for the models or a person) is the right answer: nothing in the model says which
OPEN_OK = {"downside above, no headings, no value row", "downside above, base renamed, the value reads both",
           "copies swapped, no headings, no value row"}
LINES = (("Revenue", 100.0), ("Operating costs", -40.0), ("Tax paid", -15.0))  # then Distributions, their sum


def _book(path: Path, sheet="CF", blocks=(("Base case", 1.0, None),), prior_block=False, order=None, drop=None,
          merge=None, value="base") -> dict:
    """A client model: blocks of lines under headings, each ending in Distributions (the sum), and an Equity value
    reading the base case's. blocks: (key, growth on last year's, labels {line: label} or None[, heading shown (None:
    no heading; default the key)[, sheet (default the model's sheet)]]). merge: (line, line, label): the two lines as
    one. value: the Equity value reads the base case's distributions ("base"), every block's ("both"), or there's none
    (None). -> {(key, line): (sheet, row)} (1-based)."""
    wb = xlsxwriter.Workbook(path)
    dt = wb.add_format({"num_format": "dd-mmm-yy"})
    sheets, next_row = {}, {}

    def ws_for(name):
        if name not in sheets:
            ws = sheets[name] = wb.add_worksheet(name)
            ws.write(1, 1, "Period ending")
            for k, e in enumerate(ENDS):
                ws.write_datetime(1, 3 + k, e, dt)
            next_row[name] = 3
        return sheets[name]
    ws_for(sheet)
    where = {}
    for b in blocks:
        key, g, names = b[:3]
        heading = b[3] if len(b) > 3 else key
        on = b[4] if len(b) > 4 and b[4] else sheet
        ws, r = ws_for(on), next_row.get(on, 3)
        if heading:
            ws.write(r, 0, heading)
        top = r + 1
        lines = [x for x in LINES if x[0] != drop]
        if merge:
            a, b2, label = merge
            both = sum(v for n, v in lines if n in (a, b2))
            lines = [(label, both) if n == a else (n, v) for n, v in lines if n != b2]
        if order:
            lines = [next(x for x in lines if x[0] == n) for n in order if any(x[0] == n for x in lines)]
        for i, (line, base) in enumerate(lines):
            ws.write(top + i, 1, (names or {}).get(line, line))
            for k in range(len(ENDS)):
                ws.write_number(top + i, 3 + k, base * g * 1.03 ** k)
            where[(key, line)] = (on, top + i + 1)
        dr = top + len(lines)
        ws.write(dr, 1, (names or {}).get("Distributions", "Distributions"))
        for k in range(len(ENDS)):
            c = COL(3 + k)
            ws.write_formula(dr, 3 + k, f"=SUM({c}{top + 1}:{c}{dr})", None, sum(v for _, v in lines) * g * 1.03 ** k)
        where[(key, "Distributions")] = (on, dr + 1)
        next_row[on] = dr + 2
    ws, r = sheets[sheet], next_row[sheet]
    reads = [where[(b[0], "Distributions")] for b in blocks if value == "both" or b[0] == "Base case"] if value else []
    if reads:
        refs = ",".join(f"{'' if bs == sheet else repr(bs).replace(chr(34), chr(39))+'!'}D{bd}:"
                        f"{'' if bs == sheet else repr(bs).replace(chr(34), chr(39))+'!'}{COL(3 + len(ENDS) - 1)}{bd}"
                        for bs, bd in reads)
        ws.write(r, 1, "Equity value")
        ws.write_formula(r, 2, f"=SUM({refs})", None, 1.0)
    if prior_block:  # last year's figures, by formulas from a sheet of them: a prior-forecast comparison
        pf = wb.add_worksheet("Prior")
        ws.write(r + 2, 0, "Last valuation's forecast")
        for i, (line, b) in enumerate(LINES + (("Distributions", sum(x for _, x in LINES)),)):
            ws.write(r + 3 + i, 1, f"{line} per last valuation")
            for k in range(len(ENDS)):
                c = COL(3 + k)
                pf.write_number(2 + i, 3 + k, b * 1.03 ** k)
                ws.write_formula(r + 3 + i, 3 + k, f"=Prior!{c}{3 + i}", None, b * 1.03 ** k)
    wb.close()
    return where


def pairs(out: Path) -> list[dict]:
    out.mkdir(parents=True, exist_ok=True)
    last_where = _book(out / "last.xlsx")
    last = build_map.main(str(out / "last.xlsx"), str(out / "db_last"))["db"]
    renames = {"Distributions": "Cash to equity", "Tax paid": "Income tax", "Operating costs": "Opex"}
    specs = [
        ("revised", {"blocks": (("Base case", 1.04, None),)}),
        ("renamed", {"blocks": (("Base case", 1.04, renames),)}),
        ("downside inserted", {"blocks": (("Downside case", 0.9, None), ("Base case", 1.04, None))}),
        ("sheet renamed", {"sheet": "CashFlows", "blocks": (("Base case", 1.04, None),)}),
        ("prior-forecast block", {"blocks": (("Base case", 1.04, renames),), "prior_block": True}),
        ("reordered", {"blocks": (("Base case", 1.04, None),), "order": ("Tax paid", "Revenue", "Operating costs")}),
        ("row dropped", {"blocks": (("Base case", 1.04, None),), "drop": "Tax paid"}),
        ("renamed, downside inserted", {"blocks": (("Downside case", 0.95, renames), ("Base case", 1.04, renames))}),
        ("base and upside, renamed", {"blocks": (("Upside case", 1.08, renames), ("Base case", 1.04, renames))}),
        ("everything renamed", {"blocks": (("Base case", 1.04, {**renames, "Revenue": "Sales"}),)}),
        ("prior-forecast block, downside", {"blocks": (("Downside case", 0.95, None), ("Base case", 1.04, None)),
                                            "prior_block": True}),
        # the second review's (3 October 2026): the row tools get these wrong today, listed in KNOWN_WRONG until fixed
        ("downside above, base renamed", {"blocks": (("Downside case", 0.93, None),
                                                      ("Base case", 1.04, None, "Management forecast"))}),
        ("downside above, no headings", {"blocks": (("Downside case", 0.93, None, None), ("Base case", 1.04, None, None))}),
        ("P90 above, base as P50", {"blocks": (("P90 case", 0.85, None), ("Base case", 1.04, None, "P50 case"))}),
        ("base moved to its own sheet", {"blocks": (("Downside case", 0.93, None), ("Base case", 1.04, None, "Base case",
                                                                                     "Base"))}),
        ("rebuilt, low case beside a revised central", {"blocks": (
            ("Low case", 0.90, {**renames, "Revenue": "Sales"}, "Low"),
            ("Base case", 1.20, {**renames, "Revenue": "Sales"}, "Central"))}),
        ("costs and tax merged", {"blocks": (("Base case", 1.04, None),),
                                  "merge": ("Operating costs", "Tax paid", "Operating costs and tax")}),
        # copies nothing in the model tells apart (no heading names a case; no value, or every copy feeds it): left
        # open for the models or a person, never the first copy (OPEN_OK)
        ("downside above, no headings, no value row", {"blocks": (("Downside case", 0.93, None, None),
                                                                   ("Base case", 1.04, None, None)), "value": None}),
        ("downside above, base renamed, the value reads both", {"blocks": (
            ("Downside case", 0.93, None), ("Base case", 1.04, None, "Management forecast")), "value": "both"}),
        # the copies both years, swapped this year, with nothing to tell them apart: the figures say the first
        # occurrence isn't last year's, so open (where they couldn't, it would still be matched by order: see S1)
        ("copies swapped, no headings, no value row", {"blocks": (("Downside case", 0.93, None, None),
                                                                   ("Base case", 1.04, None, None)), "value": None,
                                                        "last": {"blocks": (("Base case", 1.0, None, None),
                                                                            ("Downside case", 0.9, None, None)),
                                                                 "value": None}}),
    ]
    out_pairs = []
    for name, kw in specs:
        slug = "".join(c if c.isalnum() else "_" for c in name)
        kw = dict(kw)
        lw, ldb = last_where, last
        if "last" in kw:  # a last year's model of its own
            lw = _book(out / f"last_{slug}.xlsx", **kw.pop("last"))
            ldb = build_map.main(str(out / f"last_{slug}.xlsx"), str(out / f"db_last_{slug}"))["db"]
        where = _book(out / f"{slug}.xlsx", **kw)
        db = build_map.main(str(out / f"{slug}.xlsx"), str(out / f"db_{slug}"))["db"]
        truth = {}
        for (key, line), k in lw.items():
            truth[k] = where.get((key, line))  # None where this year's model has no such row (merged, dropped)
        out_pairs.append({"name": name, "last": ldb, "this": db, "truth": truth})
    return out_pairs


def measure(ps: list[dict]) -> dict:
    """Each of last year's rows through the row finder, then the agents' first pass (by meaning, no model): right,
    open (left for the models or a person: fine), or wrong (a row settled on the wrong row: the failure to avoid)."""
    import overlay as ov
    import rowagent
    import rowfind
    res = {}
    for p in ps:
        last, this = ov.Workbook(p["last"]), ov.Workbook(p["this"])
        f = rowfind.RowFinder(ov.RowMap(last, this), last, this)
        f.since = ov.serial(date(2025, 6, 30))
        rows, tally = [], {"right": 0, "open": 0, "wrong": 0}
        for k, want in sorted(p["truth"].items()):
            ex = f.explain(*k)
            got, how = (ex["found"], "finder") if f.confident(*k) else (None, None)
            if got is None:
                m = rowagent.by_meaning(f, *k)
                got, how = (m["to"], "meaning") if m else (None, "open")
            verdict = "open" if got is None and want is not None else "right" if got == want or (got is None and want is None) \
                else "wrong"
            if got is None and want is None:
                verdict = "right"
            tally[verdict] += 1
            rows.append({"row": f"{k[0]}!r{k[1]}", "want": want, "got": got, "how": how, "verdict": verdict,
                         "agreed": ex.get("agreed")})
        res[p["name"]] = {**tally, "rows": rows}
    return res


if __name__ == "__main__":
    import tempfile
    got = measure(pairs(Path(tempfile.mkdtemp(prefix="variants_"))))
    for name, x in got.items():
        print(f"{name:22s} right {x['right']}  open {x['open']}  wrong {x['wrong']}")
        for r in x["rows"]:
            if r["verdict"] != "right":
                print(f"    {r['verdict']:5s} {r['row']}: wanted {r['want']}, got {r['got']} ({r['how']})")
