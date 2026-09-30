"""The inputs the equity value reads that are typed into the overlay itself, outside its discountings: a net debt, a
cash balance, a declared distribution, an adjustment. The roll-forward feeds this year's client model through the
overlay and moves the valuation date; nothing tells these to change, so they stay at last year's figures ("held")
until a person sets this year's.

For each: where it is, the line of the equity value it feeds, and a suggestion from this year's client model, found by
checking last year's first: the row of last year's client model holding last year's figure at last year's valuation
date (its label agreeing), then that row in this year's model at this year's valuation date. A suggestion is only a
suggestion: it's applied once a person accepts it, or types another figure (workbench.set_held, held.json).
    find(db, [(end, cell)], sheets) -> [{"cell", "label", "value", "lines", "ends"}]
    suggest(sess, item, prior_vd, this_vd) -> {"status", "value", "last", "this", "text"}
"""
import re

import dcf
import dcftrace
import overlay as ov
import valuation

SCALES = (1.0, 1e3, 1e-3)  # a client model in A$'000 against an overlay in A$m, and the other way
SAME = {"borrowings": "debt", "borrowing": "debt", "loan": "debt", "loans": "debt", "debts": "debt",
        "dividend": "distribution", "dividends": "distribution", "distributions": "distribution"}
STOP = {"less", "add", "plus", "the", "and", "at", "date", "valuation", "value", "unpaid", "declared", "balance",
        "total", "for", "net", "from", "per"}


def _words(text: str | None) -> set[str]:
    return {SAME.get(w, w) for w in re.findall(r"[a-z]{3,}", (text or "").lower()) if w not in STOP}


def _label(db, sheet: str, row: int) -> str:
    """The row's label as typed (its first text cell), else the row map's."""
    got = db.execute("SELECT value FROM cells WHERE sheet=? AND row=? AND typeof(value)='text' AND formula IS NULL "
                     "ORDER BY col LIMIT 1", (sheet, row)).fetchone() or \
        db.execute("SELECT label FROM rows WHERE sheet=? AND row=?", (sheet, row)).fetchone()
    return (got[0] if got else "") or ""


def _number(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def find(db, cells: list[str], sheets) -> list[dict]:
    """The typed inputs the equity value cells read outside their discountings, on the overlay's own sheets (a cell
    on a client sheet, or behind a link to another workbook, is fed from this year's model: not held). A discounting's
    own inputs (its rate, dates, growth, franking utilisation) are the sourced inputs' business, not these.
    -> [{"cell", "label", "value", "lines": [{"cell", "label"}] (the lines of the equity value it feeds), "ends"}]."""
    own, out = set(sheets or []), {}
    for end, cell in cells:
        try:
            t = dcftrace.trace(db, cell)
        except ValueError:
            continue
        theirs = set()  # the single cells the discountings read as inputs
        for c in dcftrace.cores(t):
            for v in (c.get("inputs") or {}).values():
                r = dcf._ref(v, "") if isinstance(v, str) else None
                if r and r[1:3] == r[3:5]:
                    theirs.add(ov._a1(r[0], r[1], r[2]))
        leaves = []

        def walk(n):
            kids = n.get("children") or []
            if kids:
                for k in kids:
                    walk(k)
            elif not n.get("cores"):
                leaves.append(n)
        walk(t)
        for leaf in leaves:
            s, r, c = ov.parse_a1(leaf["cell"])
            got = db.execute("SELECT formula, value FROM cells WHERE sheet=? AND row=? AND col=?", (s, r, c)).fetchone()
            typed = {leaf["cell"]: {"sheet": s, "row": r, "col": c, "formula": None, "value": got[1]}} \
                if got and not got[0] else valuation.reads(db, cells=[(s, r, c)])
            for ref, x in typed.items():
                if x["formula"] or not _number(x["value"]) or x["sheet"] not in own or ref in theirs or "[" in ref:
                    continue
                item = out.setdefault(ref, {"cell": ref, "label": _label(db, x["sheet"], x["row"]) or ref,
                                            "value": float(x["value"]), "lines": [], "ends": []})
                if leaf["cell"] not in [ln["cell"] for ln in item["lines"]]:
                    item["lines"].append({"cell": leaf["cell"], "label": leaf.get("label") or leaf["cell"]})
                if end not in item["ends"]:
                    item["ends"].append(end)
    return list(out.values())


def _column(wb, sheet: str, when: str | None):
    """The column of a sheet's timeline at a date (a balance at it), or None."""
    if not when:
        return None
    want = ov.serial(ov.date.fromisoformat(when[:10]))
    return next((c for c, d in wb.timeline(sheet).items() if round(d) == want), None)


def suggest(sess, item: dict, prior_vd: str | None, this_vd: str | None) -> dict:
    """This year's figure for a held input, from this year's client model, found by checking last year's first.
    -> {"status": "checked" (last year's figure found in last year's model under a label that agrees, and that row
    this year) / "figure only" (the figure found, its label doesn't agree: check it) / "not found" / "no row this year" /
    "no date this year" / "nil", "value" (the suggestion, in the overlay's units, or None), "last", "this" (where:
    {"cell", "label", "value"}), "text" (how it was found, figure first)}."""
    v = item["value"]
    if not v:
        return {"status": "nil", "value": None, "last": None, "this": None,
                "text": "last year's figure is nil: nothing to look for; set this year's if there is one"}
    base = sess.prior or sess.ov  # last year's client model (the overlay's own workbook, where it sits inside it)
    if not sess.current or not prior_vd:
        return {"status": "not found", "value": None, "last": None, "this": None,
                "text": "no client models to look in" if not sess.current else "last year's valuation date isn't known"}
    want = _words(item["label"] + " " + " ".join(ln["label"] for ln in item["lines"]))
    hits = []
    for (s,) in base.db.execute("SELECT sheet FROM sheets"):
        if base is sess.ov and s in set(sess.sheets):
            continue
        col = _column(base, s, prior_vd)
        if col is None:
            continue
        labels = base.labels()
        for (r, c), x in base.sheet(s).items():
            if c != col or not _number(x) or not x:
                continue
            k = next((k for k in SCALES if abs(abs(x) * k - abs(v)) <= max(5e-4 * abs(v), 0.05)), None)
            if k is not None:
                lab = labels.get((s, r), "")
                hits.append({"sheet": s, "row": r, "col": c, "value": float(x), "scale": k, "sign": 1 if (x > 0) == (v > 0) else -1,
                             "label": lab, "agrees": bool(want & _words(lab))})
    if not hits:
        return {"status": "not found", "value": None, "last": None, "this": None,
                "text": f"{v:,.1f} isn't in last year's client model at {prior_vd[:10]}: set this year's figure"}
    hits.sort(key=lambda h: (not h["agrees"], h["sheet"], h["row"]))
    h = hits[0]
    last = {"cell": ov._a1(h["sheet"], h["row"], h["col"]), "label": h["label"], "value": h["value"]}
    said = f"{h['value']:,.1f} in last year's client model at {prior_vd[:10]} ({last['cell']}, {h['label'] or 'no label'})"
    others = f"; {len(hits) - 1} other cell(s) hold the same figure" if len(hits) > 1 else ""
    found = sess.rowmap.explain(h["sheet"], h["row"])["found"] if sess.rowmap else None
    if not found:
        return {"status": "no row this year", "value": None, "last": last, "this": None,
                "text": f"{said}; that row isn't found in this year's model: set this year's figure{others}"}
    s2, r2 = found
    col2 = _column(sess.current, s2, this_vd)
    if col2 is None:
        return {"status": "no date this year", "value": None, "last": last, "this": None,
                "text": f"{said}; this year's model has the row ({s2}!r{r2}) but no column at "
                        f"{(this_vd or 'this year’s valuation date')[:10]}{others}"}
    y = sess.current.value(s2, r2, col2)
    if not _number(y):
        return {"status": "no date this year", "value": None, "last": last, "this": None,
                "text": f"{said}; this year's model's cell {ov._a1(s2, r2, col2)} holds no figure{others}"}
    this = {"cell": ov._a1(s2, r2, col2), "label": sess.current.labels().get((s2, r2), ""), "value": float(y)}
    value = abs(float(y)) * h["scale"] * (1 if v > 0 else -1) * (1 if (y > 0) == (h["value"] > 0) else -1)
    # checked only where one row holds last year's figure under a label that agrees: two such rows, a person picks
    status = "checked" if h["agrees"] and sum(x["agrees"] for x in hits) == 1 else "figure only"
    how = (f"{value:,.1f} in this year's client model at {this_vd[:10]} ({this['cell']}): the row that holds last year's "
           f"figure as typed here, {said}" if status == "checked" else
           f"{value:,.1f} in this year's client model at {this_vd[:10]} ({this['cell']}), the row of {said}; " + (
               "another row with a label like this input's holds the same figure, so check it's the right one"
               if h["agrees"] else "its label doesn't match this input's, so check it's the same thing"))
    return {"status": status, "value": value, "last": last, "this": this, "text": how + others}
