"""Last year's equity value, confirmed in Python and rolled forward: the workbench's output.

  locate   the overlay cells holding the report's equity value, its low and its high (the mid is their midpoint,
           the convention); ex-distribution unless only the cum-distribution figure is in the model
  tie      Excel's saved values round to the report's figures, and the Python rebuild on last year's client model
           gives them again (a difference there is the client file not being the version the overlay was built on)
  bridge   for the low and the high: the report's figure -> rounding -> the rebuild on last year's client model ->
           time value -> last year's forecast cash flows up to the new valuation date -> this year's client model
           (new forecast, rolled forward) -> the discount rate (last year's, unchanged by default) -> this year;
           the mid's bridge is the average of the two, step by step
  chart    the undiscounted discrete forecast cash flows under the value (the terminal value left out), last year's
           and this year's, totalled by financial year
  gaps     a this-year value only where the rows its cash flows come from were found in this year's model (and the
           zero-roll check holds): otherwise the rows to find are listed, and there's no this-year value yet
Every figure is shown in the report's units (the match to the report says how the overlay's units compare).
"""
import re
from collections import Counter
from datetime import date

import dcf
import dcftrace
import keyfacts
import linkmap
import overlay as ov
import rodb

TV_WORDS = re.compile(r"terminal|continuing value|residual|perpetuity|gordon|exit value", re.I)


# ---- where the report's equity value is in the overlay ------------------------------------------------------------

def _fact_for(texts: dict, units: str | None) -> dict:
    return {"id": 0, "key": "equity_value", "label": "Equity value", "category": "conclusion", "unit": units or "",
            **{k: texts.get(k) or "" for k in ("value_text", "low_text", "high_text")}}


def candidates(summary: dict, head: dict, limit: int = 12) -> list[dict]:
    """Overlay cells whose saved value is one of the report's equity value figures (low, high, the printed mid;
    either basis), best first: for the orchestrator's decision when the pairing below finds nothing clear."""
    path, sheets = summary["wiring"]["overlay"]["db_path"], set(summary["sheets"])
    with rodb.connect(path) as db:
        nums, labels = linkmap.Numbers(db, sheets), linkmap._labels(db)
    out = []
    for basis, texts in ((head["basis"] or "ex", head["texts"]), *(((head["other"]["basis"], head["other"]["texts"]),)
                                                                  if head.get("other") else ())):
        for m in linkmap.match_fact(_fact_for(texts, head.get("units")), nums, labels, set(), limit=40)["matches"]:
            out.append({"cell": f"{m['sheet']}!{m['addr']}", "label": m["label"], "value": m["value"], "part": m["part"],
                        "basis": basis, "scale": m["scale"], "sign": m["sign"], "formula": m["formula"],
                        "label_match": m["label_match"], "score": m["score"]})
    out.sort(key=lambda x: -x["score"])
    return out[:limit]


def locate(summary: dict, head: dict) -> dict | None:
    """The cells for the low and the high (or the one figure), paired on one row where they can be:
    {"low", "high", "mid" (a cell holding the midpoint, if there is one), "scale", "sign", "basis", "how"}.
    The report's basis first; the other basis only where the first isn't in the model (a model whose cash flows are
    cum-distribution gives the cum-distribution value)."""
    path, sheets = summary["wiring"]["overlay"]["db_path"], set(summary["sheets"])
    with rodb.connect(path) as db:
        nums, labels = linkmap.Numbers(db, sheets), linkmap._labels(db)
        tries = [(head["basis"], head["texts"])]
        if head.get("other"):
            tries.append((head["other"]["basis"], head["other"]["texts"]))
        for k, (basis, texts) in enumerate(tries):
            ms = linkmap.match_fact(_fact_for(texts, head.get("units")), nums, labels, set(), limit=200)["matches"]
            got = _pair(ms)
            if got:
                got["basis"] = basis
                got["how"] = (got["how"] + (f"; the report's {head['basis'] or 'ex'}-distribution figure isn't in the overlay, "
                                            f"its {basis}-distribution one is (the model gives {basis}-distribution values)"
                                            if k else ""))
                got["switched"] = bool(k)
                return got
    return None


def _pair(ms: list[dict]) -> dict | None:
    cell = lambda m: f"{m['sheet']}!{m['addr']}"
    lows, highs, mids = ([m for m in ms if m["part"] == p] for p in ("low", "high", "value"))
    best = None
    for a in lows:
        for b in highs:
            if (a["sheet"], a["row"], a["scale"], a["sign"]) != (b["sheet"], b["row"], b["scale"], b["sign"]):
                continue
            score = a["score"] + b["score"] + 2 * (a["formula"] and b["formula"]) + (a["label_match"] and b["label_match"])
            if best is None or score > best[0]:
                best = (score, a, b)
    if best:
        _, a, b = best
        mid = next((cell(m) for m in mids if (m["sheet"], m["row"]) == (a["sheet"], a["row"])), None)
        return {"low": cell(a), "high": cell(b), "mid": mid, "scale": a["scale"], "sign": a["sign"], "label": a["label"],
                "how": f"low and high on one row ({a['label'] or a['sheet']})"}
    located = lambda xs: sorted((m for m in xs if m["located"]), key=lambda m: -m["score"])
    lo, hi = located(lows), located(highs)
    if lo and hi and (lo[0]["scale"], lo[0]["sign"]) == (hi[0]["scale"], hi[0]["sign"]):
        return {"low": cell(lo[0]), "high": cell(hi[0]), "mid": None, "scale": lo[0]["scale"], "sign": lo[0]["sign"],
                "label": lo[0]["label"], "how": "low and high found by their labels, on different rows"}
    one = located(mids)
    if one and not lows and not highs:  # the report gives one figure
        m = one[0]
        return {"low": cell(m), "high": cell(m), "mid": cell(m), "scale": m["scale"], "sign": m["sign"],
                "label": m["label"], "how": "the report's one figure"}
    return None


# ---- values on each feed, and whether this year's can be trusted -------------------------------------------------

def _read(sess, summary: dict, feed: str, cells: list[tuple], vd: str | None = None, months: int | None = None):
    defaults, roll, months = ov._feed(summary, feed, vd, months)
    sess.configure(feed, defaults, months or 0)
    return dict(zip(cells, sess.values(cells))), roll, defaults, months


def figures(sess, summary: dict, cells: list[str]) -> dict:
    """Each cell as Excel saved it, rebuilt in Python on last year's client model (the values saved in the overlay
    without one), and on this year's model rolled forward; with this year's gaps (overlay.summary_table's rules,
    for these cells): {"saved", "rebuilt", "this_year" ({cell: value}), "roll", "gaps", "feeds"}."""
    w = summary["wiring"]
    keys = [ov.parse_a1(c) for c in cells]
    base_feed = "prior" if w.get("prior") else "workbook"
    out = {"saved": {c: ov._show(sess.ov.value(*k)) for c, k in zip(cells, keys)}, "feeds": {"rebuilt": base_feed}}
    got, _, _, _ = _read(sess, summary, base_feed, keys)
    out["rebuilt"] = {c: ov._show(got[k]) for c, k in zip(cells, keys)}
    out["this_year"], out["roll"], out["gaps"] = None, None, None
    if w.get("current") and sess.rowmap:
        out["feeds"]["this_year"] = "current"
        out["gaps"] = _gaps(sess, summary, cells)
        got, out["roll"], _, _ = _read(sess, summary, "current", keys)
        out["this_year"] = {c: ov._show(got[k]) for c, k in zip(cells, keys)}
        if out["roll"]:
            tl = sess.rolled_timeline(out["roll"]["months"] or 0)
            firsts = sorted(tl.values())
            out["roll"].update(first_period=ov.to_date(firsts[0]).isoformat() if firsts else None,
                               last_period=ov.to_date(firsts[-1]).isoformat() if firsts else None)
    sess.configure("workbook")
    return out


def _gaps(sess, summary: dict, cells: list[str]) -> dict:
    """Whether this year's value can be trusted, for each cell: the rows its discountings' cash flows come from are
    found in this year's model (or picked, or kept on purpose), none of the rows read is found but blank or found
    with little confidence, the timing rows are found or worked out, this year's valuation date is known, and the
    zero-roll check holds (this year's model at last year's date gives about last year's value)."""
    by_fig = ov.dcf_origins(sess, summary, cells)
    origins = sorted({k for x in by_fig.values() for k in x["amounts"]})
    timing = sorted({k for x in by_fig.values() for k in x["timing"]} - {(s_, r_, "") for s_, r_ in origins})
    sess.derived = {}
    src = sess.prior or sess.ov
    for s_, r_, _w in timing:
        if not sess.rowmap.confident(s_, r_):
            rule = ov.timing_rule(src, s_, r_, sess.base_vd)
            if rule:
                sess.derived[(s_, r_)] = rule
    keys = [ov.parse_a1(c) for c in cells]
    this, _, _, _ = _read(sess, summary, "current", keys)
    by_row = Counter((s_, r_) for (s_, r_, _c) in sess.unmatched)
    read_by_row = Counter((s_, r_) for (s_, r_, _c) in sess.client_reads)
    labels = src.labels()
    ex = sess.rowmap.explain
    kept = lambda k: ex(*k).get("stand_in") and ex(*k).get("by", "you") in ("you", "code")
    missing = [k for k in origins if not kept(k) and (sess.rowmap.locate(*k) is None
               or by_row.get(k, 0) > 0.5 * max(1, read_by_row.get(k, 0)))]
    blank_by_row, blank_at = Counter(), {}
    for (s_, r_, _c), at in sess.blank.items():
        blank_by_row[(s_, r_)] += 1
        blank_at.setdefault((s_, r_), at)
    blank_rows = [k for k, n in blank_by_row.most_common() if n > 0.2 * max(1, read_by_row.get(k, 0)) and k not in missing]
    timing_rows = {(s_, r_) for s_, r_, _w in timing}
    weak_rows = [k for k in sorted(read_by_row) if k not in missing and k not in blank_rows
                 and k not in timing_rows and not sess.rowmap.confident(*k)]
    agents_kept = lambda k: ex(*k).get("stand_in") and ex(*k).get("by") == "agent"
    timing_open = [k for k in sorted(timing_rows) if k in read_by_row and k not in sess.derived
                   and (not sess.rowmap.confident(*k) or agents_kept(k))]
    reads = len(sess.client_reads)
    share = 1 - len(sess.unmatched) / reads if reads else 1.0
    date_hold = bool((summary.get("roll") or {}).get("date_check"))
    pvd = (summary.get("roll") or {}).get("prior_valuation_date")
    zero = this if date_hold else (_read(sess, summary, "current", keys, pvd, 0)[0] if pvd else None)
    last, _, _, _ = _read(sess, summary, "prior" if summary["wiring"].get("prior") else "workbook", keys)
    by_cell = {}
    for c, k in zip(cells, keys):
        mine = (by_fig.get(c) or {}).get("amounts") or []
        gone = [x for x in mine if x in missing]
        z, l0 = (zero or {}).get(k), last.get(k)
        ratio = z / l0 if isinstance(z, float) and isinstance(l0, float) and l0 else None
        zr = {"value": ov._show(z), "ratio": None if ratio is None else round(ratio, 4),
              "ok": ratio is not None and ov.ZERO_ROLL[0] <= ratio <= ov.ZERO_ROLL[1], "valuation_date": pvd}
        by_cell[c] = {"basis": "dcf" if mine else "share", "dcf_rows": len(mine),
                      "missing": [f"{s_}!r{r_}" for s_, r_ in gone], "zero_roll": zr,
                      "reliable": (not gone if mine else share >= 0.5) and not blank_rows and not weak_rows
                      and not timing_open and not date_hold and zr["ok"]}

    def why_missing(k):
        if ex(*k).get("stand_in"):
            return "the agents couldn't find it this year (last year's values stand in)"
        if sess.rowmap.locate(*k) is None:
            return "not found this year"
        if blank_by_row.get(k, 0) > 0.5 * by_row.get(k, 0):
            return f"found at {blank_at[k][0]}!r{blank_at[k][1]}, but blank there in most of its periods"
        return "found, but most of its periods aren't in this year's model"

    found = lambda k: (lambda h: f"{h[0]}!r{h[1]}" if h else None)(sess.rowmap.locate(*k))
    ov_labels = sess.ov.labels()
    zero_off = [{"cell": c, "label": ov_labels.get(ov.parse_a1(c)[:2], "") or c, **x["zero_roll"]}
                for c, x in by_cell.items() if not x["zero_roll"]["ok"]]
    return {"reliable": all(x["reliable"] for x in by_cell.values()), "by_cell": by_cell, "date_check": date_hold,
            "zero_roll_off": zero_off, "read_rows": [f"{s_}!r{r_}" for s_, r_ in sorted(read_by_row)],
            "timing": [{"row": f"{s_}!r{r_}", "label": labels.get((s_, r_), ""), "why": why, "found": found((s_, r_)),
                        "derived": (sess.derived.get((s_, r_)) or {}).get("text"), "open": (s_, r_) in timing_open}
                       for s_, r_, why in timing],
            "found_share": round(share, 3), "values": len(sess.unmatched), "reads": reads,
            "dcf_rows": len(origins), "dcf_origins": [f"{s_}!r{r_}" for s_, r_ in origins],
            "dcf_missing": [{"row": f"{s_}!r{r_}", "label": labels.get((s_, r_), ""), "why": why_missing((s_, r_))}
                            for s_, r_ in missing],
            "blank_rows": [{"row": f"{s_}!r{r_}", "label": labels.get((s_, r_), ""),
                            "found": f"{blank_at[(s_, r_)][0]}!r{blank_at[(s_, r_)][1]}",
                            "blank": blank_by_row[(s_, r_)], "of": read_by_row.get((s_, r_), 0)} for s_, r_ in blank_rows],
            "weak_rows": [{"row": f"{s_}!r{r_}", "label": labels.get((s_, r_), ""), "found": found((s_, r_)),
                           "confidence": ex(s_, r_)["confidence"], "how": ex(s_, r_)["how"]} for s_, r_ in weak_rows],
            "timing_open": [{"row": f"{s_}!r{r_}", "label": labels.get((s_, r_), ""), "found": found((s_, r_))}
                            for s_, r_ in timing_open]}


# ---- the bridge --------------------------------------------------------------------------------------------------

def _traced(db, cell: str) -> tuple[dict, list[dict]] | None:
    try:
        t = dcftrace.trace(db, cell)
    except ValueError:
        return None
    return t, [c for c in dcftrace.cores(t) if c.get("inputs")]


def _need(db, t: dict, cs: list[dict]) -> set[tuple]:
    need = set()
    for c in cs:
        need |= ov._dcf_cells(db, c["inputs"])

    def walk(n):
        if not n.get("again"):
            x = dcf._ref(n["cell"], "")
            need.add((x[0], x[1], x[2]))
        for ch in n.get("children", []):
            walk(ch)
    walk(t)
    return need


def _patched(sess, summary: dict, db, feed: str, need: set, roll: dict | None = None):
    """The overlay's model.db with the module's values on a feed in place (dcf.py and dcftrace read a model.db)."""
    if feed == "current":
        defaults, _, months = ov._feed(summary, "current", None, None)
        sess.configure("current", defaults, months or 0)
    else:
        sess.configure(feed)
    got = rodb.patched(summary["wiring"]["overlay"]["db_path"], ov._module_values(db, sess, need)) if need else db
    sess.configure("workbook")
    return got


def _steps(prior_db, traced, v0: float, v2: float, vd1: str | None) -> tuple[list[dict], str | None]:
    """The roll-forward from last year's rebuilt value v0 to this year's v2, split where the formulas above the
    discountings can be recomputed: time value, last year's cash flows to the new date, this year's forecast."""
    if not traced or not traced[1]:
        return [{"key": "roll", "label": "Roll-forward onto this year's model (time, cash flows and forecast)", "value": v2 - v0}], \
            "no discounting found under it, so the roll-forward is one step"
    if not vd1:
        return [{"key": "roll", "label": "Roll-forward onto this year's model", "value": v2 - v0}], \
            "no valuation date for this year"
    t, cs = traced
    try:
        base = {c["cell"]: dcf.compute(prior_db, **{**c["inputs"], "compare_to": None}, fix=False) for c in cs}
        again = dcftrace.recompute(prior_db, t, {k: x["total"] for k, x in base.items()})
        if again is None or not dcf._close(again, v0):
            return [{"key": "roll", "label": "Roll-forward onto this year's model (time, cash flows and forecast)",
                     "value": v2 - v0}], "the formulas above its discountings can't be recomputed here, so the roll-forward is one step"
        vd1d = date.fromisoformat(vd1[:10])
        grown = {k: x["total"] * (1 + x["rate"]) ** dcf.yearfrac(x["valuation_date"], vd1d, x["day_count"])
                 for k, x in base.items()}
        rolled = {c["cell"]: dcf.compute(prior_db, **{**c["inputs"], "compare_to": None, "valuation_date": vd1},
                                         fix=False)["total"] for c in cs}
        vu, v1 = dcftrace.recompute(prior_db, t, grown), dcftrace.recompute(prior_db, t, rolled)
        if vu is None or v1 is None:
            raise ValueError("a formula on the way up can't be evaluated here")
        years = dcf.yearfrac(next(iter(base.values()))["valuation_date"], vd1d, "actual/actual")
        return [{"key": "time", "label": f"Time value: {years:.2f} years of unwind at the discount rate", "value": vu - v0},
                {"key": "cash", "label": f"Last year's forecast cash flows up to {vd1[:10]}", "value": v1 - vu},
                {"key": "forecast", "label": "This year's client model (new forecast, rolled forward)", "value": v2 - v1}], None
    except (ValueError, ZeroDivisionError, OverflowError) as e:
        return [{"key": "roll", "label": "Roll-forward onto this year's model (time, cash flows and forecast)",
                 "value": v2 - v0}], f"couldn't split the roll-forward: {e}"


def bridges(sess, summary: dict, head: dict, where: dict, figs: dict) -> dict:
    """The low's and the high's bridges, and the mid's (their average), in the report's units."""
    path = summary["wiring"]["overlay"]["db_path"]
    db = rodb.connect(path)
    unit = lambda v: v / (where["scale"] or 1.0) * (where["sign"] or 1) if isinstance(v, float) else None
    vd1 = (figs.get("roll") or {}).get("valuation_date")
    ends, notes = {}, []
    held = figs.get("gaps") and not figs["gaps"]["reliable"]
    for end in ("low", "high"):
        cell = where[end]
        rep = head[end]
        s0, r0 = unit(figs["saved"].get(cell)), unit(figs["rebuilt"].get(cell))
        v2 = unit((figs.get("this_year") or {}).get(cell)) if figs.get("this_year") and not held else None
        steps = [{"key": "report", "label": "Last year, per the report", "value": rep, "total": True}]
        if s0 is not None and rep is not None:
            steps.append({"key": "rounding", "label": "Rounding in the report (Excel's saved value)", "value": s0 - rep})
        if r0 is not None and s0 is not None:
            steps.append({"key": "prior_feed", "label": "Rebuilt in Python on last year's client model", "value": r0 - s0})
        steps.append({"key": "rebuilt", "label": "Last year, rebuilt in Python", "value": r0, "total": True})
        if v2 is not None and r0 is not None:
            traced = _traced(db, cell)
            need = _need(db, *traced) if traced and traced[1] else set()
            prior_db = _patched(sess, summary, db, figs["feeds"]["rebuilt"], need)
            # the steps in the overlay's units, then shown in the report's
            raw, note = _steps(prior_db, traced, figs["rebuilt"][cell], figs["this_year"][cell], vd1)
            steps += [{**x, "value": unit(x["value"])} for x in raw]
            steps.append({"key": "rate", "label": "Discount rate: last year's, unchanged", "value": 0.0})
            steps.append({"key": "this_year", "label": f"This year, rolled forward{f' to {vd1[:10]}' if vd1 else ''}",
                          "value": v2, "total": True})
            if note:
                notes.append(f"{end}: {note}")
        ends[end] = {"cell": cell, "steps": steps}
    # the mid: the average of the low's and the high's, step by step (the same steps where both split alike)
    keys = [[s["key"] for s in ends[e]["steps"]] for e in ("low", "high")]
    if keys[0] != keys[1]:  # one end split, the other didn't: both as one roll-forward step
        for e in ("low", "high"):
            st = ends[e]["steps"]
            roll = [s for s in st if s["key"] in ("time", "cash", "forecast", "roll")]
            if roll:
                i = st.index(roll[0])
                st[i:i + len(roll)] = [{"key": "roll", "label": "Roll-forward onto this year's model (time, cash flows and forecast)",
                                        "value": sum(s["value"] for s in roll)}]
    mid = []
    for a, b in zip(ends["low"]["steps"], ends["high"]["steps"]):
        both = [x for x in (a["value"], b["value"]) if isinstance(x, float)]
        mid.append({**a, "value": sum(both) / 2 if len(both) == 2 else None})
    if mid and mid[0]["key"] == "report":
        mid[0]["value"] = head["mid"]
    return {"low": ends["low"], "high": ends["high"], "mid": {"steps": mid}, "notes": notes, "valuation_date": vd1,
            "units": head.get("units"), "held": bool(held)}


# ---- the cash-flow chart ------------------------------------------------------------------------------------------

def _fy(d: date, end_month: int) -> str:
    y = d.year if d.month <= end_month else d.year + 1
    return f"FY{y % 100:02d}"


def _discrete_rows(core: dict) -> tuple[list[str], list[str]]:
    """The cash-flow ranges of a discounting without its terminal value: the parts of its cash-flow row that aren't
    labelled or built like a terminal value, over the same columns. -> (ranges, the parts left out)."""
    rng = core["inputs"]["cashflow"]
    parts = core.get("parts") or []
    tv = [p for p in parts if TV_WORDS.search(p.get("label") or "") or re.search(r"/\s*\(.*-.*growth", p.get("words") or "", re.I)]
    keep = [p for p in parts if p not in tv]
    if not tv or not keep or len(rng) != 1:
        return list(rng), []
    m = re.match(r"^(.+?)!\$?([A-Z]{1,3})\$?\d+:\$?([A-Z]{1,3})\$?\d+$", rng[0])
    if not m:
        return list(rng), []
    out = []
    for p in keep:
        pm = re.match(r"^(?:\[\d+\])?(.+)!r(\d+)$", p["row"])
        if not pm or pm[1] != m[1].strip("'"):
            return list(rng), []
        out.append(f"{m[1]}!{m[2]}{pm[2]}:{m[3]}{pm[2]}")
    return out, [p.get("label") or p["row"] for p in tv]


def chart(sess, summary: dict, where: dict, figs: dict, fy_end: int) -> dict:
    """The undiscounted discrete forecast cash flows under the value, by financial year: last year's (rebuilt on
    last year's client model, periods after last year's valuation date) and this year's (rolled forward, after the
    new date). The discounting with the largest value under the low's cell (the enterprise value, rather than a
    side stream like franking credits), its terminal value left out."""
    path = summary["wiring"]["overlay"]["db_path"]
    db = rodb.connect(path)
    traced = _traced(db, where["low"])
    if not traced or not traced[1]:
        return {"why": "no discounting found under the equity value"}
    t, cs = traced
    main = max(cs, key=lambda c: abs(c.get("pv") or 0.0))
    ranges, left_out = _discrete_rows(main)
    inputs = {**main["inputs"], "cashflow": ranges}
    need = ov._dcf_cells(db, inputs)
    unit = lambda v: v / (where["scale"] or 1.0) * (where["sign"] or 1)
    series = {}
    for key, feed, vd in (("last_year", figs["feeds"]["rebuilt"], (summary.get("roll") or {}).get("prior_valuation_date")),
                          ("this_year", "current" if figs.get("this_year") else None, (figs.get("roll") or {}).get("valuation_date"))):
        if not feed:
            continue
        pdb = _patched(sess, summary, db, feed, need)
        flows, ends = ov._flows(pdb, inputs)
        vd_d = date.fromisoformat(vd[:10]) if vd else None
        by_fy = {}
        for c, v in flows.items():
            e = ends.get(c)
            e = e if isinstance(e, date) else None
            if e is None or (vd_d and e <= vd_d):
                continue
            by_fy.setdefault(_fy(e, fy_end), [e, 0.0])
            by_fy[_fy(e, fy_end)][1] += v
        series[key] = {k: unit(v) for k, (_, v) in sorted(by_fy.items(), key=lambda kv: kv[1][0])}
    years = sorted({y for s in series.values() for y in s}, key=lambda y: int(y[2:]) + (100 if int(y[2:]) < 50 else 0))
    return {"years": years, "series": series, "rows": ranges, "left_out": left_out, "core": main["cell"],
            "label": next((p.get("label") for p in (main.get("parts") or []) if p.get("label") not in left_out), None),
            "held": bool(figs.get("gaps") and not figs["gaps"]["reliable"]), "units": None}


# ---- the model's assumptions under the value ---------------------------------------------------------------------

def assumptions(summary: dict, where: dict) -> dict:
    """For the low and the high: each discounting under the cell, as the workbook saved it (dcf.compute on the
    saved values: rate, valuation date, convention, cut-off, periods, undiscounted cash flows), and where each
    comes from. The Model assumptions card."""
    path = summary["wiring"]["overlay"]["db_path"]
    db = rodb.connect(path)
    out = {}
    for end in ("low", "high"):
        traced = _traced(db, where[end])
        rows = []
        for c in (traced[1] if traced else []):
            try:
                r = dcf.compute(db, **{**c["inputs"], "compare_to": None}, fix=False)
            except ValueError as e:
                rows.append({"cell": c["cell"], "error": str(e)})
                continue
            tv = [p for p in (c.get("parts") or []) if TV_WORDS.search(p.get("label") or "")]
            rows.append({"cell": c["cell"], "kind": c.get("kind"), **ov._brief(r),
                         "rate_source": c["inputs"].get("rate") if isinstance(c["inputs"].get("rate"), str) else
                         (c.get("method") or {}).get("rate_note") or "a constant",
                         "valuation_date_source": c["inputs"].get("valuation_date"),
                         "cashflow": c["inputs"]["cashflow"], "parts": [p.get("label") for p in c.get("parts") or []],
                         "terminal_value": tv[0]["total"] if tv else None, "terminal_value_row": tv[0]["row"] if tv else None})
        out[end] = rows
    return out


# ---- all of it --------------------------------------------------------------------------------------------------

def compute(sess, summary: dict, facts: list[dict], markdown: str, fy_end: int, override: dict | None = None) -> dict:
    """The workbench's result, or {"stop": why, ...} where it can't go on (the orchestrator decides what next)."""
    head = keyfacts.conclusion(facts, markdown)
    if not head:
        return {"stop": "no_equity_value", "why": "the report's equity value isn't among the key facts"}
    where = override or locate(summary, head)
    if not where:
        return {"stop": "not_located", "head": head, "candidates": candidates(summary, head),
                "why": "the report's equity value (low and high) wasn't found in the overlay"}
    cells = list(dict.fromkeys(x for x in (where["low"], where["high"], where.get("mid")) if x))
    figs = ov.deep(figures, sess, summary, cells)
    if where.get("basis") == "cum" and head["basis"] != "cum" and head.get("other"):
        head = {**head, **{k: head["other"][k] for k in ("low", "high")}, "basis": "cum",
                "texts": head["other"]["texts"], "why": head["why"] + ["the overlay gives the cum-distribution value"]}
        if head["low"] is not None and head["high"] is not None:
            head["mid"] = (head["low"] + head["high"]) / 2
        else:
            head["low"] = head["high"] = head["mid"] = head["other"]["value"]
    unit = lambda v: v / (where["scale"] or 1.0) * (where["sign"] or 1) if isinstance(v, float) else None
    tie = {}
    for end in ("low", "high"):
        text = head["texts"].get(f"{end}_text") or head["texts"].get("value_text")
        t = ov.tie(figs["saved"].get(where[end]), text, where["scale"], where["sign"]) if text else None
        tie[end] = {"report": head[end], "report_text": text, "saved": unit(figs["saved"].get(where[end])),
                    "rebuilt": unit(figs["rebuilt"].get(where[end])), "ok": bool(t and t["ok"]),
                    "rebuilt_ok": bool((ov.tie(figs["rebuilt"].get(where[end]), text, where["scale"], where["sign"]) or {})
                                       .get("ok")) if text else False}
    br = ov.deep(bridges, sess, summary, head, where, figs)
    ch = ov.deep(chart, sess, summary, where, figs, fy_end)
    this = {e: unit((figs.get("this_year") or {}).get(where[e])) for e in ("low", "high")} if figs.get("this_year") else None
    return {"head": head, "where": where, "tie": tie, "figures": figs, "bridges": br, "chart": ch,
            "assumptions": assumptions(summary, where),
            "values": {"report": {e: head[e] for e in ("low", "mid", "high")},
                       "rebuilt": {"low": tie["low"]["rebuilt"], "high": tie["high"]["rebuilt"],
                                   "mid": (tie["low"]["rebuilt"] + tie["high"]["rebuilt"]) / 2
                                   if tie["low"]["rebuilt"] is not None and tie["high"]["rebuilt"] is not None else None},
                       "this_year": ({**this, "mid": (this["low"] + this["high"]) / 2
                                      if this["low"] is not None and this["high"] is not None else None}
                                     if this and not br["held"] else None)}}
