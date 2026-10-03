"""This year's cash flows against last year's, period by period: the checks that the roll-forward took this year's
model in, and the bridge's new-forecast step split into what moved.

  layer    every discounting under the low's and the high's cells, per period (its end, the cash flow, the terminal
           value's part of it, the factor): last year's (rebuilt on last year's client model, at last year's date),
           last year's rolled to the new date (the bridge's "cash flows" end point), and this year's (this year's
           feed at last year's rate)
  stale    this year's discounted cash flows are last year's, period for period, while this year's client model's
           rows behind them changed: the overlay didn't take this year's model in (a link, a pick, a cached value)
  split    the new-forecast step as revisions to the periods both years have, periods added or dropped at the end,
           the terminal value's change, and what's left (moves outside the discountings): a large residual, or a step
           of the other sign to the cash flows' change, is a point to check or a hold
  horizon  this year's model forecasting past the last period the overlay reads (where last year's didn't), a
           terminal value reading nil this year, last year's figures standing in for periods this year's model lacks
Each finding is a result.hold: a person can acknowledge it with a reason.
"""
from datetime import date

import dcf
import dcftrace
import overlay as ov
import rodb

EXACT = 1e-9          # a cash flow the same as last year's, relative
ZERO_EXACT = 1e-4     # a zero-roll ratio this close to 1 while the model's rows changed: too exact to be a revision
SPLIT_CHECK = 0.005   # what's left of the new-forecast step after the cash flows' change, as a share of last year's value
SPLIT_HOLD = 0.02     # ... and this much: held until a person says why
SIGN_FLIP = 0.01      # a discounting's present value changing sign, both years at least this share of the value: held
SIGN_MIN = 0.005      # the cash flows' change, as a share of last year's value, from which its sign must be the step's
MIN_PERIODS = 2       # periods both years have after the new date, at least, to call the cash flows the same
# the checks' ids (the diagnostics may carry them: they're the app's words)
IDS = ("cf-stale", "cf-same", "cf-exact", "cf-split-low", "cf-split-high", "cf-sign-low", "cf-sign-high", "cf-tv-nil",
       "cf-short", "cf-horizon", "cf-standin", "cf-error", "cf-uncut", "cf-untimed", "cf-xnpv", "cf-ondate", "cf-flip",
       "cf-midperiod")
KINDS = ("sumproduct", "pv row", "xnpv", "npv", "unknown")


def _iso(d) -> str | None:
    return d.isoformat() if isinstance(d, date) else None


def _periods(pdb, inputs: dict, discrete: list[str], r: dict) -> dict:
    """{period end ISO: {"col", "cf", "tv", "factor"}} of a discounting on a model.db, at r's rate, date and convention:
    cf the whole cash flow, tv the part of it that's the terminal value (the cash-flow rows less the discrete ones)."""
    fl, ends = ov._flows(pdb, inputs)
    disc = fl if list(discrete) == list(inputs["cashflow"]) else ov._flows(pdb, {**inputs, "cashflow": discrete})[0]
    fac = dcf.factors(ends, r["valuation_date"], r["rate"], r["timing"], r["day_count"], r.get("terminal_date"))
    out = {}
    for c, e in ends.items():
        if isinstance(e, date):
            out[e.isoformat()] = {"col": c, "cf": fl.get(c, 0.0), "tv": fl.get(c, 0.0) - disc.get(c, 0.0),
                                  "factor": fac.get(c, 0.0)}
    return out


def _run(pdb, inputs: dict, vd=None) -> dict:
    kw = {**inputs, "compare_to": None}
    if vd:
        kw["valuation_date"] = vd
    return dcf.compute(pdb, **kw, fix=False)


def _raw(sess, summary: dict, db, c: dict, feed: str) -> dict:
    """A discounting the app doesn't recompute (XNPV, NPV, factors of its own), on this year's feed as the overlay
    works it out: {period end ISO: {"cf", "factor"}} (the factor None where it can't be read: an XNPV's or an NPV's
    is its function's), and an XNPV's first date."""
    import result
    look = {"cashflow": [c["cashflow"]], "mask": c.get("mask"), "dates": c.get("dates")}
    extra = c.get("factor_row") or c.get("pv_row")
    need = ov._dcf_cells(db, {**look, "cashflow": [c["cashflow"]] + ([extra] if extra else [])})
    pdb = result._patched(sess, summary, db, feed, need)
    fl, ends = ov._flows(pdb, look)
    fac = {}
    if extra:
        s, r, cols = dcf._row_range(pdb, extra)
        cs = dcf._row_range(pdb, c["cashflow"])[2]
        vals = dict(pdb.execute("SELECT col, value FROM cells WHERE sheet=? AND row=? AND col BETWEEN ? AND ?",
                                (s, r, cols[0], cols[-1])))
        for k, col in enumerate(cs):
            v = dcf._num(vals.get(cols[k])) if k < len(cols) else None
            if v is not None:
                fac[col] = v / fl[col] if c.get("pv_row") and fl.get(col) else v if not c.get("pv_row") else None
    out = {e.isoformat(): {"cf": fl.get(col, 0.0), "factor": fac.get(col)} for col, e in ends.items() if isinstance(e, date)}
    first = None
    if c.get("kind") == "xnpv" and c.get("dates"):
        s, r, cols = dcf._row_range(pdb, c["dates"])
        ds = [dcf._as_date(v) or ov.to_date(v) if isinstance(v, (int, float)) else dcf._as_date(v) for (v,) in
              pdb.execute("SELECT value FROM cells WHERE sheet=? AND row=? AND col BETWEEN ? AND ? ORDER BY col",
                          (s, r, cols[0], cols[-1]))]
        ds = [d for d in ds if isinstance(d, date)]
        first = ds[0].isoformat() if ds else None
    return {"periods": out, "first_date": first}


def layer(sess, summary: dict, where: dict, figs: dict) -> dict:
    """{"cores": [{"cell", "kind", "what", "label", "ends", "form", "last", "rolled", "this"}], "trees": {end: ...},
    "vd0", "vd1"}: each discounting under the value, its periods on last year's feed, rolled to the new date, and on
    this year's feed at last year's rate. A discounting the app can't recompute is listed with its form and why."""
    import result
    path = summary["wiring"]["overlay"]["db_path"]
    db = rodb.connect(path)
    vd0 = (summary.get("roll") or {}).get("prior_valuation_date")
    vd1 = (figs.get("roll") or {}).get("valuation_date")
    this_feed = "current" if figs.get("this_year") else None
    cores, trees, seen = [], {}, {}
    for end in ("low", "high"):
        traced = result._traced(db, where[end])
        if not traced:
            continue
        t, cs = traced
        trees[end] = {"tree": t, "cores": [c["cell"] for c in cs]}
        for c in dcftrace.cores(t):
            if c["cell"] in seen:
                seen[c["cell"]]["ends"].append(end)
                continue
            entry = {"cell": c["cell"], "kind": c.get("kind"), "what": c.get("what"), "ends": [end],
                     "label": c.get("cashflow_label"), "form": form(c), "pv": c.get("pv")}
            seen[c["cell"]] = entry
            cores.append(entry)
            if not c.get("inputs"):
                entry["why"] = ("its convention isn't one the app recomputes: its rate and date are read from its formulas"
                                if c.get("loose") else "the app can't recompute it period by period")
                if this_feed and c.get("cashflow"):  # this year's cash flows and factors as the overlay works them out
                    try:
                        entry["raw"] = _raw(sess, summary, db, c, this_feed)
                    except (ValueError, ZeroDivisionError, OverflowError) as e:
                        entry["raw_why"] = str(e)
                continue
            inputs = c["inputs"]
            discrete, left_out = result._discrete_rows(c)
            entry["tv_rows"] = left_out
            # the rows the discrete forecast is made of too, not only the whole cash-flow row: else they'd keep the
            # values Excel saved last year on every feed, and the terminal value's part come out wrong
            need = ov._dcf_cells(db, inputs) | (ov._dcf_cells(db, {**inputs, "cashflow": discrete})
                                                if list(discrete) != list(inputs["cashflow"]) else set())
            try:
                last_db = result._patched(sess, summary, db, figs["feeds"]["rebuilt"], need)
                r0 = _run(last_db, inputs)
                entry["last"] = {"rate": r0["rate"], "valuation_date": _iso(r0["valuation_date"]), "timing": r0["timing"],
                                 "day_count": r0["day_count"], "pv": r0["total"],
                                 "periods": _periods(last_db, inputs, discrete, r0)}
                if vd1:
                    r1 = _run(last_db, inputs, vd1)
                    entry["rolled"] = {"pv": r1["total"], "valuation_date": _iso(r1["valuation_date"]),
                                       "periods": _periods(last_db, inputs, discrete, r1)}
                if this_feed:
                    this_db = result._patched(sess, summary, db, this_feed, need, rates=False)
                    r2 = _run(this_db, inputs)
                    entry["this"] = {"rate": r2["rate"], "valuation_date": _iso(r2["valuation_date"]),
                                     "timing": r2["timing"], "day_count": r2["day_count"], "pv": r2["total"],
                                     "periods": _periods(this_db, inputs, discrete, r2)}
            except (ValueError, ZeroDivisionError, OverflowError) as e:
                entry["why"] = f"couldn't read it period by period: {e}"
    return {"cores": cores, "trees": trees, "vd0": vd0, "vd1": vd1}


def form(c: dict) -> dict:
    """How a discounting is built, for a person: its function (XNPV, NPV, SUMPRODUCT, a present-value row), the
    convention its factors fit (end of period, mid-period, mid-year; the day count) or that they don't fit one the
    app recomputes, and what that means for the roll-forward (cut off and time-checked, or not)."""
    kind = c.get("kind") or "unknown"
    words = {"xnpv": "XNPV", "npv": "NPV", "sumproduct": "SUMPRODUCT", "pv row": "Present values summed"}.get(kind, kind)
    inp, loose = c.get("inputs") or {}, c.get("loose")
    if kind == "xnpv":
        conv = "actual/365 from the first date (XNPV)"
    elif kind == "npv":
        conv = "one period per column, end of period (NPV: no dates)"
    elif inp:
        conv = {"end": "end of period", "mid": "mid-period", "mid-year": "mid-year"}.get(inp.get("timing"), inp.get("timing")) \
            + f", {inp.get('day_count')}"
    elif loose:
        conv = "its own (factors that don't fit a convention the app recomputes)"
    else:
        conv = "not read"
    exact = bool(inp) and kind not in ("xnpv", "npv")
    return {"function": words, "kind": kind, "convention": conv, "recomputed": exact,
            "rolled": "cut off and time-checked" if exact else "not cut off or time-checked by the app"}


def _same(a: float, b: float) -> bool:
    return abs(a - b) <= EXACT * max(1.0, abs(a), abs(b))


def _changed_rows(sess, summary: dict, cells: list[str], vd1: str | None) -> dict:
    """The client rows the value's cash flows come from: how many changed between last year's and this year's model
    in the periods after the new date (and how many could be compared)."""
    out = {"rows": 0, "compared": 0, "changed": 0, "examples": []}
    if not sess.rowmap or not vd1:
        return out
    since = ov.serial(date.fromisoformat(vd1[:10]))
    by_fig = ov.dcf_origins(sess, summary, cells)
    origins = sorted({k for x in by_fig.values() for k in x["amounts"]})
    out["rows"] = len(origins)
    rm = sess.rowmap
    for s, r in origins:
        hit = rm.locate(s, r)
        if not hit:
            continue
        a, b = rm._series(rm.prior, s, r), rm._series(rm.current, *hit)
        both = [w for w in a if w > since and w in b]
        if not both:
            continue
        out["compared"] += 1
        diff = [w for w in both if not _same(a[w], b[w])]
        if diff:
            out["changed"] += 1
            if len(out["examples"]) < 3:
                out["examples"].append(f"{s}!r{r} → {hit[0]}!r{hit[1]}: {len(diff)} of {len(both)} periods revised")
    return out


def _tail_zeros(periods: dict, after: str | None) -> int:
    xs = [p["cf"] - p["tv"] for e, p in sorted(periods.items()) if not after or e > after]
    n = 0
    for v in reversed(xs):
        if abs(v) > EXACT:
            break
        n += 1
    return n if n < len(xs) else 0


def checks(sess, summary: dict, where: dict, figs: dict, fl: dict, unit) -> dict:
    """The cash-flow checks on the layer: {"holds": [result.hold], "split": {end: {...}}, "compare": {...}}."""
    import result
    holds, split = [], {}
    vd1 = fl.get("vd1")
    cells = [c for c in (where.get("low"), where.get("high")) if c]
    rows = _changed_rows(sess, summary, cells, vd1)
    have = [c for c in fl["cores"] if c.get("last") and c.get("this")]

    # a discounting whose present value turns the other way (a sign convention changed in a row it reads: tax shown
    # positive, a credit stream going negative): both material, last year's and this year's at last year's rate
    ref0 = max([abs(v) for v in (figs.get("rebuilt") or {}).values() if isinstance(v, float)] + [0.0]) or 1.0
    for c in have:
        a, b = c["last"].get("pv"), c["this"].get("pv")
        if isinstance(a, float) and isinstance(b, float) and a * b < 0 and min(abs(a), abs(b)) > SIGN_FLIP * ref0:
            holds.append(result.hold(summary, f"cf-flip-{c['cell']}", [c["cell"], a, b],
                                     f"{c.get('label') or c['cell']} turns the other way: {unit(a):,.1f} last year, "
                                     f"{unit(b):,.1f} this year at last year's rate",
                                     f"{c['cell']}'s present value changed sign: a row it reads may have changed its sign "
                                     "convention (tax shown positive, a credit as a cost). Check the rows it reads, or "
                                     "acknowledge why it's right", check="cf-flip"))

    # this year's date inside a period of a discounting (a mid-year date on annual periods): the period is kept whole,
    # though part of it was earned before the date. A point to check, once
    for c in have:
        ends = sorted(c["this"]["periods"])
        if not vd1 or vd1[:10] in ends:
            continue
        before, after = [e for e in ends if e < vd1[:10]], [e for e in ends if e > vd1[:10]]
        d = lambda x: date.fromisoformat(x[:10])
        if not before and len(after) >= 2:  # this year's model starts after the period's start: its length from the next
            months = round((d(after[1]) - d(after[0])).days / 30.44)
            before = [ov.to_date(ov.add_months(ov.serial(d(after[0])), -months)).isoformat()]
        # (a date within a few days of the period's start is its start: a calendar's end of month, not mid-period)
        if before and after and (d(vd1) - d(before[-1])).days > 3 and (d(after[0]) - d(vd1)).days > 3:
            e0, e1 = before[-1], after[0]
            frac = (d(vd1) - d(e0)).days / max(1, (d(e1) - d(e0)).days)
            p = c["this"]["periods"][e1]
            holds.append(result.hold(summary, f"cf-midperiod-{c['cell']}", [c["cell"], vd1[:10], e1],
                                     f"This year's valuation date ({vd1[:10]}) falls inside the period ending {e1}",
                                     f"{frac:.0%} of the period from {e0} to {e1} is before the date, and its whole cash flow "
                                     f"({unit(p['cf']):,.1f}) is kept in this year's value, the part earned before the date "
                                     "with it. Check whether it should be prorated (the methods card's mid-period is the "
                                     "nearest)", severity="check", check="cf-midperiod"))
            break

    # stale: this year's discounted cash flows are last year's while the model's rows behind them changed
    same_cores = []
    for c in have:
        lp, tp = c["last"]["periods"], c["this"]["periods"]
        # the discrete cash flows, the terminal value left out: on a rolling horizon it moves a column, and a forecast
        # left wholly at last year's would never look the same with it in
        d = lambda p: p["cf"] - p["tv"]
        shared = [e for e in tp if e in lp and (not vd1 or e > vd1[:10]) and (abs(d(lp[e])) > EXACT or abs(d(tp[e])) > EXACT)]
        if len(shared) >= MIN_PERIODS and all(_same(d(lp[e]), d(tp[e])) for e in shared):
            same_cores.append((c, len(shared)))
    if same_cores and rows["changed"]:
        c, n = same_cores[0]
        holds.append(result.hold(summary, "cf-stale", [x["cell"] for x, _ in same_cores],
                                 "This year's value discounts last year's cash flows: this year's model changed but the "
                                 "overlay's cash flows didn't",
                                 f"{c['cell']} ({c.get('label') or c['form']['function']}) has last year's figures in all "
                                 f"{n} periods both years have after {vd1}, while {rows['changed']} of the {rows['compared']} "
                                 f"client rows behind it were revised ({'; '.join(rows['examples'])}). A link still on last "
                                 "year's file, a row matched wrongly or a value held at Excel's would look like this",
                                 cores=[x["cell"] for x, _ in same_cores]))
    elif same_cores and rows["compared"] and not rows["changed"]:
        c, n = same_cores[0]
        holds.append(result.hold(summary, "cf-same", [x["cell"] for x, _ in same_cores],
                                 "This year's client model has last year's forecast for the rows the value reads",
                                 f"{c['cell']}'s cash flows are last year's in all {n} periods both years have after {vd1}, "
                                 f"and so are the {rows['compared']} client rows behind them: check this year's model is "
                                 "the updated one", severity="check"))

    # a zero roll exactly 1 while the rows changed: the value didn't take this year's model in
    for cell, x in (((figs.get("gaps") or {}).get("by_cell")) or {}).items():
        ratio = (x.get("zero_roll") or {}).get("ratio")
        if ratio is not None and abs(ratio - 1) < ZERO_EXACT and rows["changed"]:
            holds.append(result.hold(summary, "cf-exact", [cell, ratio],
                                     f"At last year's date, this year's model gives exactly last year's value ({ratio:.4f}×)",
                                     f"{cell}: {rows['changed']} of the {rows['compared']} client rows behind it were "
                                     "revised, so a revision would move it. Exactly the same value means the revisions "
                                     "aren't reaching it"))
            break

    # a discounting the app doesn't cut off or recompute (XNPV, NPV, factors of its own): the overlay's own formulas
    # roll it, so they're checked here: an XNPV counting from this year's date, nothing on or before the new date still
    # discounted, and its time measured where it can be
    timed = {x["cell"]: x for x in ((figs.get("time") or {}).get("discountings") or [])}
    for c in fl["cores"]:
        if (c.get("form") or {}).get("recomputed") or not vd1:
            continue
        raw, fn = c.get("raw") or {}, (c.get("form") or {}).get("function") or "a discounting"
        if not raw and c.get("this"):  # read period by period (an XNPV has exact inputs): this year's as the layer has it
            raw = {"periods": {e: {"cf": p["cf"], "factor": None} for e, p in c["this"]["periods"].items()},
                   "first_date": c["this"].get("valuation_date") if c.get("kind") == "xnpv" else None}
        ps = raw.get("periods") or {}
        past = {e: p for e, p in ps.items() if e <= vd1[:10] and abs(p["cf"]) > EXACT
                and (p["factor"] is None or abs(p["factor"]) > EXACT)}
        if c.get("kind") == "xnpv":
            first = raw.get("first_date")
            if first and first != vd1[:10]:
                holds.append(result.hold(summary, f"cf-xnpv-{c['cell']}", [c["cell"], first, vd1],
                                         f"The XNPV at {c['cell']} counts from {first}, not this year's valuation date",
                                         f"XNPV discounts every cash flow from the first date in its range: on this year's "
                                         f"model that's {first}, so this year's value is discounted to it, not to {vd1}", check="cf-xnpv"))
            elif first and abs((ps.get(first) or {}).get("cf") or 0.0) > EXACT:
                holds.append(result.hold(summary, f"cf-ondate-{c['cell']}", [c["cell"], ps[first]["cf"]],
                                         f"The XNPV at {c['cell']} takes the cash flow on this year's valuation date in "
                                         "whole", f"{unit(ps[first]['cf']):,.1f} on {first}, undiscounted: in the value as "
                                         "at the date (cum-distribution), where the default cuts the period ending on it",
                                         severity="check", check="cf-ondate"))
            continue
        if past:
            tot = sum(p["cf"] for p in past.values())
            holds.append(result.hold(summary, f"cf-uncut-{c['cell']}", [c["cell"], sorted(past), tot],
                                     f"{fn} at {c['cell']} still discounts {len(past)} period(s) ending on or before "
                                     "this year's valuation date",
                                     f"{', '.join(sorted(past)[:4])}: {unit(tot):,.1f} in total, in this year's value. The "
                                     f"app can't cut this discounting off ({(c.get('form') or {}).get('convention', 'its own')}), and its own "
                                     "formulas don't: the past is in this year's value (or, for the period ending on the "
                                     "date, it's cum-distribution on purpose: say so)", check="cf-uncut"))
        elif c["cell"] not in timed:
            holds.append(result.hold(summary, f"cf-untimed-{c['cell']}", [c["cell"], fn],
                                     f"{fn} at {c['cell']}: the roll's time isn't measured",
                                     f"the app can't work out its factors on another date ({(c.get('form') or {}).get('convention', 'its own')}), "
                                     "so that the roll moves it on by its rate is a person's to check",
                                     severity="check", check="cf-untimed"))

    # the new-forecast step, split: revisions, periods added and dropped, the terminal value, what's left
    v0s, before = figs.get("rebuilt") or {}, (figs.get("this_year_held") or figs.get("this_year_last_rate")
                                              or figs.get("this_year") or {})
    unmoved = figs.get("this_year_unmoved") or {}  # the balances at the valuation date left at last year's date
    db = None
    for end in ("low", "high"):
        cell, tr = where.get(end), (fl.get("trees") or {}).get(end)
        mine = [c for c in have if end in c["ends"] and c.get("rolled")]
        if not tr or not mine or cell not in before or len(mine) != len(tr["cores"]):
            continue
        db = db or rodb.connect(summary["wiring"]["overlay"]["db_path"])
        try:
            t = tr["tree"]
            need = result._need(db, t, [c for c in dcftrace.cores(t) if c.get("inputs")])
            prior_db = result._patched(sess, summary, db, figs["feeds"]["rebuilt"], need)
            base = {c["cell"]: c["rolled"]["pv"] for c in mine}
            v1 = dcftrace.recompute(prior_db, t, base)
            v_cf = dcftrace.recompute(prior_db, t, {c["cell"]: c["this"]["pv"] for c in mine})
            if v1 is None or v_cf is None:
                continue
            cats = {"revised": {}, "added": {}, "dropped": {}, "terminal": {}}
            for c in mine:
                rp, tp = c["rolled"]["periods"], c["this"]["periods"]
                live = lambda ps, e: (not vd1 or e > vd1[:10]) and ps[e]["factor"]
                rev = add = drop = tv = 0.0
                for e, p in tp.items():
                    if not live(tp, e):
                        continue
                    if e in rp:
                        rev += ((p["cf"] - p["tv"]) - (rp[e]["cf"] - rp[e]["tv"])) * p["factor"]
                        tv += (p["tv"] - rp[e]["tv"]) * p["factor"]
                    else:
                        add += (p["cf"] - p["tv"]) * p["factor"]
                        tv += p["tv"] * p["factor"]
                for e, p in rp.items():
                    if live(rp, e) and e not in tp:
                        drop -= (p["cf"] - p["tv"]) * p["factor"]
                        tv -= p["tv"] * p["factor"]
                for k, v in (("revised", rev), ("added", add), ("dropped", drop), ("terminal", tv)):
                    cats[k][c["cell"]] = v
            parts = {}
            for k, by in cats.items():
                moved = dcftrace.recompute(prior_db, t, {c: base[c] + by.get(c, 0.0) for c in base})
                parts[k] = (moved - v1) if moved is not None else None
            d_cf, step = v_cf - v1, before[cell] - v1
            # the balances at the valuation date read at this year's date (a net debt, a cash balance): known, not
            # something else moving outside the discountings
            bal = before[cell] - unmoved[cell] if isinstance(unmoved.get(cell), float) else 0.0
            known = sum(v for v in parts.values() if v is not None)
            ref = abs(v0s.get(cell) or 0.0) or abs(v1) or 1.0
            out = {"step": unit(step), "cash_flows": unit(d_cf), "balances": unit(bal), "outside": unit(step - d_cf - bal),
                   "convention": unit(d_cf - known), **{k: unit(v) if v is not None else None for k, v in parts.items()},
                   "share_outside": (step - d_cf - bal) / ref}
            split[end] = out
            off = abs(step - d_cf - bal) / ref
            if off > SPLIT_CHECK:
                holds.append(result.hold(summary, f"cf-split-{end}", [cell, step, d_cf],
                                         f"The {end} end's new-forecast step isn't the cash flows' change: "
                                         f"{off:.1%} of last year's value moves outside the discountings",
                                         f"step {out['step']:,.1f}, of which this year's discounted cash flows "
                                         f"{out['cash_flows']:,.1f} (revised {out['revised'] or 0:,.1f}, added "
                                         f"{out['added'] or 0:,.1f}, dropped {out['dropped'] or 0:,.1f}, terminal value "
                                         f"{out['terminal'] or 0:,.1f})" + (f"; balances at the valuation date read at this "
                                         f"year's date {out['balances']:,.1f}" if bal else "") + f"; {out['outside']:,.1f} "
                                         "is something else this year's "
                                         "model moves (an input read from it outside the discountings) or a roll the cash "
                                         "flows don't explain",
                                         severity="block" if off > SPLIT_HOLD else "check"))
            if abs(d_cf) / ref > SIGN_MIN and step * d_cf < 0:
                holds.append(result.hold(summary, f"cf-sign-{end}", [cell, step, d_cf],
                                         f"The {end} end's new-forecast step goes the other way to the cash flows' change",
                                         f"this year's discounted cash flows move the value {out['cash_flows']:+,.1f}, the "
                                         f"step is {out['step']:+,.1f}", severity="check"))
        except (ValueError, ZeroDivisionError, OverflowError, TypeError) as e:
            split[end] = {"why": f"couldn't split the step: {e}"}

    # the horizon: this year's model past the overlay's last period, a terminal value reading nil, stand-ins
    rm = sess.rowmap
    for c in have:
        lp, tp = c["last"]["periods"], c["this"]["periods"]
        tv_last = sum(p["tv"] for p in lp.values() if p["factor"])
        tv_this = sum(p["tv"] for p in tp.values() if p["factor"])
        if abs(tv_last) > EXACT and abs(tv_this) <= EXACT * max(1.0, abs(tv_last)) * 1e3:
            holds.append(result.hold(summary, "cf-tv-nil", [c["cell"], tv_last],
                                     "This year's terminal value reads nil",
                                     f"{c['cell']}'s terminal value ({', '.join(c.get('tv_rows') or []) or 'its last period'}) "
                                     f"was {unit(tv_last):,.1f} last year and is nil on this year's model: this year's "
                                     "forecast may end before the overlay's last column, which the terminal value grows "
                                     "from"))
        n0, n1 = _tail_zeros(lp, fl.get("vd0")), _tail_zeros(tp, vd1)
        if n1 > n0:
            holds.append(result.hold(summary, "cf-short", [c["cell"], n0, n1],
                                     "This year's forecast ends before the overlay's last column",
                                     f"{c['cell']}: the last {n1} period(s) read nil this year (last year {n0}): the "
                                     "overlay's columns run past this year's forecast", severity="check"))
    if rm and vd1 and have:
        last_end = max((e for c in have for e, p in c["this"]["periods"].items() if p["factor"]), default=None)
        last_end0 = max((e for c in have for e, p in c["last"]["periods"].items() if p["factor"]), default=None)
        if last_end and last_end0:
            s1, s0 = ov.serial(date.fromisoformat(last_end)), ov.serial(date.fromisoformat(last_end0))
            past, total = [], 0.0
            by_fig = ov.dcf_origins(sess, summary, cells)
            for s, r in sorted({k for x in by_fig.values() for k in x["amounts"]}):
                hit = rm.locate(s, r)
                if not hit:
                    continue
                now = {w: v for w, v in rm._series(rm.current, *hit).items() if w > s1 and abs(v) > EXACT}
                then = {w: v for w, v in rm._series(rm.prior, s, r).items() if w > s0 and abs(v) > EXACT}
                if now and not then:
                    past.append(f"{hit[0]}!r{hit[1]} ({len(now)} period(s) to {ov.to_date(max(now)).isoformat()})")
                    total += sum(now.values())
            if past:
                holds.append(result.hold(summary, "cf-horizon", [past, total],
                                         "This year's model forecasts past the last period the overlay reads",
                                         f"{'; '.join(past[:4])}: {total:,.1f} in total after {last_end}, where last "
                                         "year's model had nothing past the overlay's columns. The roll leaves those "
                                         "years out of this year's value (an extended life or a longer forecast?)"))
    stood = (figs.get("feed") or {}).get("stood_in") or []
    g = figs.get("gaps") or {}
    origin = set(g.get("dcf_origins") or []) - {x["row"] for x in g.get("dcf_missing") or []}
    kept = set()
    if rm:  # a row a person (or the code) kept at last year's on purpose isn't standing in by accident
        for row in origin:
            s_, r_ = row.rsplit("!r", 1)
            ex = rm.explain(s_, int(r_))
            if ex.get("stand_in") and ex.get("by", "you") in ("you", "code"):
                kept.add(row)
    mine = [x for x in stood if x["row"] in origin and x["row"] not in kept]
    if mine and any(abs(x["value"]) > EXACT for x in mine):
        tot = sum(x["value"] for x in mine)
        holds.append(result.hold(summary, "cf-standin", [sorted({x["row"] for x in mine}), len(mine), tot],
                                 f"Last year's figures stand in for {len(mine)} period(s) this year's model doesn't have",
                                 f"on {', '.join(sorted({x['row'] for x in mine})[:4])}: {tot:,.1f} in total of last year's "
                                 "forecast in this year's value, where this year's model has no period (or a blank) for "
                                 "it"))
    return {"holds": holds, "split": split, "rows": rows}
