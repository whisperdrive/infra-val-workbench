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
           zero-roll check holds): otherwise the rows to find are listed, and there's no this-year value yet; and
           the cash-flow lines this year's model has that the overlay doesn't read, to check (they don't hold it)
Every figure is shown in the report's units (the match to the report says how the overlay's units compare).

A check that holds this year's value back is a hold (hold()): its id, a key from the figures it found, and what it
says. A person can acknowledge it with a reason (workbench.acknowledge); the acknowledgement is keyed by the figures,
so it lapses by itself when they change.
"""
import hashlib
import json
import re
from collections import Counter
from datetime import date

import cashflows
import dcf
import dcftrace
import forward
import held
import keyfacts
import linkmap
import overlay as ov
import rodb
import rowfind
import scenarios
import sourced
import valuation

TV_WORDS = re.compile(r"terminal|continuing value|residual|perpetuity|gordon|exit value", re.I)
FRANKING = re.compile(r"frank|imputation|gamma", re.I)
# a line of cash flows to or from equity, by its label; not a rate, a ratio or a flag about one
FLOW_WORDS = re.compile(r"equity|injection|contribution|distribution|dividend|capital return|return of capital|"
                        r"buy-?back|redemption|subscription|capital rais|shareholder loan", re.I)
NOT_A_FLOW = re.compile(r"cost of|return on|\birr\b|\brates?\b|ratio|%|gearing|\bbeta\b|premium|multiple|yield|"
                        r"margin|\bflags?\b|factor", re.I)
PV_WORDS = re.compile(r"\bn?pv\b|present value|discounted", re.I)
VERSION = 2  # the result's rules: a result worked out by older rules is worked out again (the stage's inputs)


def fingerprint(obj) -> str:
    """A short key for the figures a check found (rounded, so a rerun's float noise doesn't change it)."""
    def rnd(x):
        if isinstance(x, float):
            return float(f"{x:.6g}")
        if isinstance(x, dict):
            return {k: rnd(v) for k, v in x.items()}
        if isinstance(x, (list, tuple)):
            return [rnd(v) for v in x]
        return x
    return hashlib.sha1(json.dumps(rnd(obj), sort_keys=True, default=str).encode()).hexdigest()[:12]


def hold(summary: dict, hid: str, figures, title: str, detail: str, severity: str = "block", **more) -> dict:
    """A check's finding: {"id", "key", "title", "detail", "severity", "acked"}. A "block" holds this year's value back
    unless a person acknowledged this id on these figures (summary["acks"]); a "check" is a point to look at."""
    key = fingerprint(figures)
    ack = (summary.get("acks") or {}).get(hid)
    acked = ack if ack and ack.get("key") == key else None
    return {"id": hid, "key": key, "title": title, "detail": detail, "severity": severity, "acked": acked, **more}


def holding(holds: list[dict]) -> bool:
    """Whether any of these findings holds the value back (a block nobody acknowledged)."""
    return any(h["severity"] == "block" and not h.get("acked") for h in holds)


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

def _read(sess, summary: dict, feed: str, cells: list[tuple], vd: str | None = None, months: int | None = None,
          held: bool = True, rates: bool = True):
    defaults, roll, months = ov._feed(summary, feed, vd, months, held, rates)
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
        # the roll moves each discounting on by its rate: where one is far off, its dates don't move together
        out["time"] = time_check(sess, summary, cells)
        if out["time"] and out["time"]["hold"]:
            out["gaps"]["reliable"] = False
            out["gaps"]["time_off"] = [x for x in out["time"]["discountings"] if abs(x["off"]) > TIME_HOLD]
        got, out["roll"], _, _ = _read(sess, summary, "current", keys)
        out["this_year"] = {c: ov._show(got[k]) for c, k in zip(cells, keys)}
        # last year's figures standing in on this year's feed, and the periods past this year's forecast (read nil)
        out["feed"] = {"stood_in": [{"row": f"{s_}!r{r_}", "col": c_, "value": float(v)}
                                    for (s_, r_, c_), v in sess.stood_in.items() if isinstance(v, (int, float))],
                       "beyond": [{"row": f"{s_}!r{r_}", "col": c_, "why": str(w)} for (s_, r_, c_), w in sess.beyond.items()]}
        out["cut_off"] = len(sess.cut)  # the discountings' periods before the new date, cut off on this year's feed
        if summary.get("rate_values"):
            got, _, _, _ = _read(sess, summary, "current", keys, rates=False)  # at last year's rate: the rate's own step
            out["this_year_last_rate"] = {c: ov._show(got[k]) for c, k in zip(cells, keys)}
        if any(x.get("value") is not None for x in (summary.get("held_values") or {}).values()):
            # the inputs a person set left at last year's, and the rate too: the roll-forward's steps end here
            got, _, _, _ = _read(sess, summary, "current", keys, held=False, rates=False)
            out["this_year_held"] = {c: ov._show(got[k]) for c, k in zip(cells, keys)}
        if out["roll"]:
            tl = sess.rolled_timeline(out["roll"]["months"] or 0)
            firsts = sorted(tl.values())
            out["roll"].update(first_period=ov.to_date(firsts[0]).isoformat() if firsts else None,
                               last_period=ov.to_date(firsts[-1]).isoformat() if firsts else None)
    sess.configure("workbook")
    return out


FOUND_MIN = 0.9  # a figure not under a discounting: this share of its client reads found in this year's model
TIME_CHECK = 0.005  # a discounting moving on by a rate this far from its own (a year): a point to check
TIME_HOLD = 0.015  # and this far: its dates don't move together, and this year's value is held back
NEW_LINES = 12  # the cash-flow lines the overlay doesn't read: at most this many, the largest first


def _rows_read(tree: dict) -> set[tuple]:
    """Every row a formula in the tree reads, long ranges too (an operand row, not followed cell by cell)."""
    out = set()

    def walk(n):
        here = dcf._ref(n["cell"], "")[0] if n.get("cell") else ""
        for m in dcf._FREF.finditer(n.get("formula") or ""):
            r = dcf._ref(m[0], here)
            if r:
                out.update((r[0], row) for row in range(r[1], r[3] + 1))
        for ch in n.get("children") or []:
            walk(ch)
    walk(tree)
    return out


def time_check(sess, summary: dict, cells: list[str]) -> dict | None:
    """Whether the roll moves each discounting on by its own rate: this year's model at last year's date (the zero
    roll, the periods to the new date cut off) against the roll at last year's rate, each discounting's present value
    of the same cash flows. Rolled a year at 10%, a discounting's present value grows by 10%; where it doesn't, its
    dates don't move together (cash flow dates counted from one cell, discount periods from a copy of it) and the
    value is wrong, however plausible it looks. The same cash flows in both runs, matched by their amounts (this
    year's model gives them either way, in whichever column the overlay puts them); a terminal value worked out again
    from the periods before it, or a period only one run has, isn't matched. The discountings under the figures, and
    another output's of the rows their formulas read; not an XNPV or an NPV (they count from their own first date or
    column, which the roll doesn't move), nor factors worked out inside a formula. The session is left on the
    workbook feed. -> {"t", "discountings": [{"cell", "rate", "ratio", "implied", "off", "ok"}],
    "measured", "ok", "hold"}, or None where there's no roll."""
    roll = summary.get("roll") or {}
    months = roll.get("months") or 0
    if not months or sess.base_vd is None or not sess.current:
        return None
    db = rodb.connect(summary["wiring"]["overlay"]["db_path"])
    traced = [x for x in (_traced(db, c) for c in cells) if x]
    read = set().union(set(), *(_rows_read(t) for t, _ in traced))
    row_of = lambda rng: (lambda x: (x[0], x[1]))(dcf._row_range(db, rng)) if rng else None
    # the discountings under the figures, and another output's of the rows their formulas read (a sum of the present
    # values the figure adds up with a SUMIF): not a discounting of its own elsewhere (a value bridge's)
    # every discounting under the figures, those whose convention the app doesn't recompute too: their factors are
    # read as the overlay works them out on each feed, at the rate their formulas read (annualised)
    cores = {(c["cell"], c["call"]): c for t, _ in traced for c in dcftrace.cores(t)}
    for c in ov._cores(db, summary.get("outputs") or []):
        if c.get("inputs") and row_of(c.get("pv_row") or c.get("factor_row")) in read:
            cores.setdefault((c["cell"], c["call"]), c)
    plan = []
    for c in cores.values():
        i, loose = c.get("inputs"), c.get("loose") or {}
        if c.get("kind") in ("xnpv", "npv") or not (c.get("pv_row") or c.get("factor_row")) or not (i or loose):
            continue
        try:
            rate = dcf.resolve(db, i["rate"])[0] if i else loose.get("rate_value")
            cfs = i["cashflow"] if i else [c["cashflow"]] if c.get("cashflow") else []
            if not cfs:
                continue
            if c.get("pv_row"):
                terms = [[dcf._row_range(db, c["pv_row"])]]
            else:
                rows = [dcf._row_range(db, c["factor_row"])] + [dcf._row_range(db, m) for m in c.get("mask") or []]
                terms = [[dcf._row_range(db, cf)] + rows for cf in cfs]
            flows = [dcf._row_range(db, cf) for cf in cfs]
        except (ValueError, KeyError):
            continue
        if isinstance(rate, (int, float)) and not isinstance(rate, bool):
            plan.append((c, float(rate), terms, flows))
    if not plan:
        return {"measured": 0, "ok": None, "hold": False, "discountings": [],
                "why": "no discounting under the figures is read exactly, so the roll's time can't be measured"}
    keys = {(s, r, col) for _, _, terms, flows in plan for rows in terms + [flows] for s, r, cols in rows for col in cols}

    def run(d, m, cut_to=None):
        try:
            sess.configure("current", d, m)
            if cut_to is not None:
                sess.cut |= sess._cut_off(cut_to)
                sess.B.range_cache.clear()
            return dict(zip(keys, sess.values(list(keys))))
        finally:
            sess.configure("workbook")

    def measure(got, terms, flows):
        """Each column's (cash flow, present value), in column order, where both are there."""
        num = lambda v: float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0
        n = len(terms[0][0][2])
        pv = [0.0] * n
        for rows in terms:
            for k in range(n):
                x = 1.0
                for s, r, cols in rows:
                    x *= num(got.get((s, r, cols[k])))
                pv[k] += x
        cf = [sum(num(got.get((s, r, cols[k]))) for s, r, cols in flows if k < len(cols)) for k in range(n)]
        return [(cf[k], pv[k]) for k in range(n) if cf[k] and pv[k]]

    def matched(a, b):
        """The same cash flows in both runs, by their amounts (this year's model gives them either way, in whichever
        column the overlay puts them), an amount met more than once matched in column order: (pv a, pv b) pairs. A
        terminal value worked out again, or a period one run has and the other doesn't, isn't matched."""
        key = lambda x: float(f"{x:.9g}")
        pool = {}
        for cf, pv in a:
            pool.setdefault(key(cf), []).append(pv)
        out = []
        for cf, pv in b:
            got = pool.get(key(cf))
            if got:
                out.append((got.pop(0), pv))
        return out

    t = (ov.add_months(sess.base_vd, months) - sess.base_vd) / 365.0
    then, _, _ = ov._feed(summary, "current", ov.to_date(sess.base_vd).isoformat(), 0, held=False, rates=False)
    now, _, m = ov._feed(summary, "current", None, None, held=False, rates=False)
    z, v = run(then, 0, months), run(now, m or 0)
    out = []
    for c, rate, terms, flows in plan:
        pairs = matched(measure(z, terms, flows), measure(v, terms, flows))
        pz, pv = sum(p for p, _ in pairs), sum(q for _, q in pairs)
        if not pairs or not pz or pv / pz <= 0:
            continue
        ratio = pv / pz
        implied = ratio ** (1 / t) - 1
        out.append({"cell": c["cell"], "label": c.get("cashflow_label") or None, "periods": len(pairs),
                    "rate": rate, "ratio": ratio, "implied": implied, "off": implied - rate,
                    "ok": abs(implied - rate) <= TIME_CHECK, "loose": not c.get("inputs")})
    return {"t": t, "discountings": out, "measured": len(out), "ok": all(x["ok"] for x in out) if out else None,
            "hold": any(abs(x["off"]) > TIME_HOLD for x in out),
            "why": None if out else "no discounting under the figures has a present value after the new date to measure"}


def _moves(sess, summary: dict, keys: list[tuple], base: dict, row: tuple, cols: set) -> bool:
    """Whether a client row's figures move these cells on this year's feed (at last year's rate): its cells as read,
    each nudged (a tenth more, plus one), and the cells worked out again. A row the discountings' formulas reach but
    only a check or a label reads doesn't, and isn't one to wait for. The session is left on the workbook feed."""
    if not cols:
        return True
    nudge = lambda v: v * 1.1 + 1 if isinstance(v, (int, float)) and not isinstance(v, bool) else 1.0
    defaults, _, months = ov._feed(summary, "current", None, None, rates=False)
    try:
        sess.configure("current", defaults, months or 0)
        # nudged where the feed hands the row's values over: read in the workbook or through a link alike
        f0, e0 = sess.B.feed, sess.B.ext
        hit = lambda s, r, c: (s, r) == tuple(row) and c in cols
        sess.B.feed = lambda s, r, c: nudge(f0(s, r, c)) if hit(s, r, c) else f0(s, r, c)
        sess.B.ext = lambda i, s, r, c: nudge(e0(i, s, r, c)) if hit(s, r, c) else e0(i, s, r, c)
        got = dict(zip(keys, sess.values(keys)))
    finally:
        sess.configure("workbook")
    return any(not ov.same(got[k], base.get(k)) for k in keys)


def _after(wb, sheet: str, when: float) -> list[int]:
    """The columns of a sheet's timeline whose periods end after a date (a timeline of period starts, all on the
    1st, ends the day before the next start)."""
    tl = wb.timeline(sheet)
    ds = sorted(set(tl.values()))
    if len(ds) < 2:
        return []
    gaps = sorted(b - a for a, b in zip(ds, ds[1:]))
    plen = max(1, round(gaps[len(gaps) // 2] / 30.44))
    starts = all(ov.to_date(d).day == 1 for d in ds)
    return sorted(c for c, d in tl.items() if (ov.add_months(d, plen) - 1 if starts else d) > when)


def _new_lines(sess, read_rows: list[tuple], vd: str | None, prior_vd: str | None) -> list[dict]:
    """The cash-flow lines of this year's model the overlay doesn't read, on the client sheets it reads: rows
    labelled as cash flows to or from equity (an injection, a contribution, a distribution, a dividend, a capital
    return) with figures after this year's valuation date, that match no row of last year's model (rowfind, the
    other way round) or match one with nothing after last year's date, and that the rows the overlay reads don't
    read (a line they add up is in the value already).
    Last year's overlay had nothing of them to discount, so the roll can't take them in: a point to check, not a
    hold. A present-value row of one of them goes with it, not on its own. Figures in this year's model's units."""
    cur, src = sess.current, sess.prior or sess.ov
    if not read_rows or not cur or not sess.rowmap or not vd:
        return []
    here = {sess.rowmap.locate(*k) for k in read_rows} - {None}
    sheets = ({sess.rowmap.sheet_for(s) for s, _r in read_rows} | {s for s, _r in here}) - {None}
    reads = sess.rowmap._index()["edges"][1][0]
    fed = {x for k in here for x in reads.get(k, ())}  # what the rows the overlay reads add up or work out from
    new_vd = ov.serial(date.fromisoformat(vd[:10]))
    cols = {s: _after(cur, s, new_vd) for s in sheets}
    found = {}
    for (s, r), lab in sorted(cur.labels().items()):
        if not cols.get(s) or (s, r) in here or (s, r) in fed or not FLOW_WORDS.search(lab) or NOT_A_FLOW.search(lab):
            continue
        vals, tl = cur.sheet(s), cur.timeline(s)
        xs = [(c, vals.get((r, c))) for c in cols[s]]
        xs = [(c, v) for c, v in xs if isinstance(v, float) and abs(v) > 1e-9]
        if xs and not all(0 < v <= 1 for _c, v in xs):  # not a flag or a factor
            found[(s, r)] = {"row": f"{s}!r{r}", "label": lab, "periods": len(xs), "total": sum(v for _c, v in xs),
                             "from": ov.to_date(tl[xs[0][0]]).isoformat(), "to": ov.to_date(tl[xs[-1][0]]).isoformat()}
    if not found:
        return []
    back = rowfind.RowFinder(ov.RowMap(cur, src), cur, src, only={s for s, _r in read_rows})
    then = ov.serial(date.fromisoformat(prior_vd[:10])) if prior_vd else None
    for k, x in list(found.items()):
        was = back.locate(*k) if back.confident(*k) else None
        if not was:
            x["why"] = "no row like it in last year's model"
            continue
        last = src.sheet(was[0])
        if then is None or any(isinstance(last.get((was[1], c)), float) and abs(last[(was[1], c)]) > 1e-9
                               for c in _after(src, was[0], then)):
            found.pop(k)  # last year's model had it, with figures to discount: the overlay left it out on purpose
        else:
            x["why"] = f"last year's {was[0]}!r{was[1]} had nothing after last year's valuation date"
    for k, x in list(found.items()):  # the model's own present value of a line goes with the line
        if PV_WORDS.search(x["label"]):
            of = next((j for j in sorted(reads.get(k, ())) if j in found and not PV_WORDS.search(found[j]["label"])), None)
            found.pop(k)
            if of:
                found[of]["pv"] = {"row": x["row"], "label": x["label"], "total": x["total"]}
    unit = lambda k: (cur.db.execute("SELECT units FROM rows WHERE sheet=? AND row=?", k).fetchone() or ("",))[0] or ""
    out = [{**x, "units": unit(k), "pv": x.get("pv")} for k, x in found.items()]
    return sorted(out, key=lambda x: -abs(x["total"]))[:NEW_LINES]


TERM_UP, TERM_DOWN = 2, 3  # the sums looked at for terms: rows this far above and below the rows the overlay reads
GONE = "was:"  # a confirmation's key for one of last year's terms gone this year, before last year's row


def _figures_after(wb, k: tuple, when: float) -> list[tuple]:
    """A row's figures in the periods ending after a date, other than nil: [(col, value)]; none for a row of flags,
    factors or dates."""
    vals = wb.sheet(k[0])
    xs = [(c, vals.get((k[1], c))) for c in _after(wb, k[0], when)]
    xs = [(c, v) for c, v in xs if isinstance(v, float) and abs(v) > 1e-9]
    if not xs or all(0 < v <= 1 for _c, v in xs) or all(v.is_integer() and 30000 <= v <= 80000 for _c, v in xs):
        return []
    return xs


def _term_changes(sess, read_rows: list[tuple], vd: str | None, prior_vd: str | None, confirmed: set,
                  lines: set) -> tuple[list[dict], list[dict]]:
    """The terms this year's model adds to, or drops from, the sums the overlay's rows sit in: for last year's rows
    the overlay reads, the rows they add up or work out from (TERM_DOWN deep) and the rows that add them up
    (TERM_UP), each found this year and its terms set against last year's (rowfind.terms: paired by the row finder,
    else by the trace's step; a term regrouped under a subtotal of the sum is neither).
    A term this year's model has and last year's didn't, with figures after this year's valuation date:
      in the value   it feeds a row the overlay reads (a new cost under the cash flow it reads): the value moves by
                     it, so it holds the value back until a person confirms it belongs (confirmed: this year's row)
      left out      it feeds none of them (a new line beside the ones the overlay reads): the value doesn't take it
                     in, a point to check, not a hold; one the cash-flow lines list has (lines) isn't listed again
    A term last year's model had and this year's doesn't, with figures after last year's valuation date: in the
    value where last year's rows the overlay reads took it in (the value no longer has it: held until a person
    confirms it's gone, GONE + last year's row), else beside them, a point to check. One the overlay read itself
    isn't listed: that's a row to find. -> (added, dropped), each figures in its own model's units."""
    cur, rm, src = sess.current, sess.rowmap, sess.prior or sess.ov
    if not read_rows or not cur or not rm or not vd:
        return [], []
    (p_reads, p_by), (c_reads, _) = rm._index()["edges"]

    def reach(starts, graph, depth):
        seen, frontier = set(starts), list(starts)
        for _ in range(depth):
            frontier = [x for k in frontier for x in graph.get(k, ()) if x not in seen]
            seen |= set(frontier)
        return seen
    sums = reach(read_rows, p_reads, TERM_DOWN) | reach(read_rows, p_by, TERM_UP)
    here = {rm.locate(*k) for k in read_rows} - {None}
    inside = reach(here, c_reads, 12)  # what the rows the overlay reads take in, this year
    took_then = reach(read_rows, p_reads, 12)  # and last year
    new_vd = ov.serial(date.fromisoformat(vd[:10]))
    then = ov.serial(date.fromisoformat(prior_vd[:10])) if prior_vd else None
    plab, labels, added, dropped = src.labels(), cur.labels(), {}, {}
    span = lambda wb, k, xs: (ov.to_date(wb.timeline(k[0])[xs[0][0]]).isoformat() if xs[0][0] in wb.timeline(k[0]) else None,
                              ov.to_date(wb.timeline(k[0])[xs[-1][0]]).isoformat() if xs[-1][0] in wb.timeline(k[0]) else None)
    for p in sorted(sums):
        got = rm.terms(*p)
        if not got:
            continue
        under = rm.locate(*p) if rm.confident(*p) else rm._anchor(p)
        sums_at = {"under": f"{under[0]}!r{under[1]}" if under else None, "under_label": labels.get(under, "") if under
                   else "", "last_year": f"{p[0]}!r{p[1]}", "last_year_label": plab.get(p, "")}
        for k in got[0]:
            if k in added or k in here:
                continue
            xs = _figures_after(cur, k, new_vd)
            row, took = f"{k[0]}!r{k[1]}", k in inside
            if not xs or (not took and row in lines):
                continue
            a, b = span(cur, k, xs)
            added[k] = {"row": row, "key": row, "label": labels.get(k, ""), "periods": len(xs),
                        "total": sum(v for _c, v in xs), "from": a, "to": b, **sums_at, "in_value": took,
                        "confirmed": row in confirmed, "hold": took and row not in confirmed}
        for x in got[1]:
            if x in dropped or x in read_rows or then is None:
                continue
            xs = _figures_after(src, x, then)
            if not xs:
                continue
            row, took, now = f"{x[0]}!r{x[1]}", x in took_then, rm.locate(*x) if rm.confident(*x) else None
            a, b = span(src, x, xs)
            dropped[x] = {"row": row, "key": GONE + row, "label": plab.get(x, ""), "periods": len(xs),
                          "total": sum(v for _c, v in xs), "from": a, "to": b, **sums_at, "in_value": took,
                          "now": f"{now[0]}!r{now[1]}" if now else None, "confirmed": GONE + row in confirmed,
                          "hold": took and GONE + row not in confirmed}
    unit = lambda wb, k: (wb.db.execute("SELECT units FROM rows WHERE sheet=? AND row=?", k).fetchone() or ("",))[0] or ""
    order = lambda xs: sorted(xs, key=lambda x: (not x["in_value"], -abs(x["total"])))[:NEW_LINES]
    return (order([{**x, "units": unit(cur, k)} for k, x in added.items()]),
            order([{**x, "units": unit(src, k)} for k, x in dropped.items()]))


def _gaps(sess, summary: dict, cells: list[str]) -> dict:
    """Whether this year's value can be trusted, for each cell: this year's client model is read at all (an overlay
    that reads nothing of it would give last year's figures, rolled by date alone), the rows its discountings' cash
    flows come from are found in this year's model (or picked, or kept on purpose), none of the rows read is found
    but blank or found with little confidence, the timing rows are found or worked out, this year's valuation date is
    known and every discounting reads it, and the zero-roll check holds (this year's model at last year's date gives
    about last year's value)."""
    by_fig = ov.dcf_origins(sess, summary, cells)
    origins = sorted({k for x in by_fig.values() for k in x["amounts"]})
    timing = sorted({k for x in by_fig.values() for k in x["timing"]} - {(s_, r_, "") for s_, r_ in origins})
    sess.derived = {}
    src = sess.prior or sess.ov
    # not found surely, or kept at last year's by the agents (a stand-in reads last year's at the rolled period, which
    # for a row of periods of its own is another period's): worked out where a rule fits
    open_ = lambda k: not sess.rowmap.confident(*k) or (sess.rowmap.explain(*k).get("stand_in")
                                                        and sess.rowmap.explain(*k).get("by") == "agent")
    for s_, r_, _w in timing:
        if open_((s_, r_)):
            rule = ov.timing_rule(src, s_, r_, sess.base_vd)
            if rule:
                sess.derived[(s_, r_)] = rule
    # a row the discounted amounts read that carries periods of its own (year numbers, dates: an annual block under a
    # quarterly header) is timing too: worked out, not hunted for in this year's model
    for s_, r_ in origins:
        if (s_, r_) not in sess.derived and open_((s_, r_)):
            tl = src.timeline(s_)
            vals = [v for v in (src.value(s_, r_, c) for c in sorted(tl)) if isinstance(v, float)]
            rule = ov.own_rule(vals)
            if rule:
                sess.derived[(s_, r_)] = rule
    for (s_, r_), rule in sess.derived.items():  # year numbers move with a row of dates on the same sheet
        if rule["kind"] == "own_year" and not rule.get("ends"):
            rule["ends"] = next((x["ends"] for (s2, _r), x in sess.derived.items() if s2 == s_ and x.get("ends")), [])
    keys = [ov.parse_a1(c) for c in cells]
    # at last year's discount rate, wherever a person set this year's: a new rate isn't a row matched wrongly
    this, info, _, _ = _read(sess, summary, "current", keys, rates=False)
    roll = summary.get("roll") or {}
    date_hold = bool(roll.get("date_check"))
    # every discounting under the figures reads this year's valuation date on this year's feed (not one of them
    # left on a date cell of its own, discounting to last year's)
    vd_reads, vd_off = roll.get("valuation_date_reads") or [], []
    if vd_reads and not date_hold and (info or {}).get("valuation_date"):
        want = ov.serial(date.fromisoformat(info["valuation_date"][:10]))
        for c, v in zip(vd_reads, sess.values([ov.parse_a1(c) for c in vd_reads])):
            if not isinstance(v, (int, float)) or isinstance(v, bool) or round(v) != want:
                vd_off.append({"cell": c, "date": ov.to_date(v).isoformat() if isinstance(v, (int, float)) and v > 0
                               else str(v)})
    by_row = Counter((s_, r_) for (s_, r_, _c) in sess.unmatched)
    read_by_row = Counter((s_, r_) for (s_, r_, _c) in sess.client_reads)
    labels = src.labels()
    ex = sess.rowmap.explain
    kept = lambda k: ex(*k).get("stand_in") and ex(*k).get("by", "you") in ("you", "code")
    missing = [k for k in origins if not kept(k) and k not in sess.derived and (sess.rowmap.locate(*k) is None
               or by_row.get(k, 0) > 0.5 * max(1, read_by_row.get(k, 0)))]
    blank_by_row, blank_at = Counter(), {}
    for (s_, r_, _c), at in sess.blank.items():
        blank_by_row[(s_, r_)] += 1
        blank_at.setdefault((s_, r_), at)
    blank_rows = [k for k, n in blank_by_row.most_common() if n > 0.2 * max(1, read_by_row.get(k, 0)) and k not in missing]
    timing_rows = {(s_, r_) for s_, r_, _w in timing}
    weak_rows = [k for k in sorted(read_by_row) if k not in missing and k not in blank_rows
                 and k not in timing_rows and k not in sess.derived and not sess.rowmap.confident(*k)]
    agents_kept = lambda k: ex(*k).get("stand_in") and ex(*k).get("by") == "agent"
    timing_open = [k for k in sorted(timing_rows) if k in read_by_row and k not in sess.derived
                   and (not sess.rowmap.confident(*k) or agents_kept(k))]
    reads = len(sess.client_reads)
    share = 1 - len(sess.unmatched) / reads if reads else 0.0  # nothing read of this year's model: nothing found
    read_cols = {}
    for s_, r_, c_ in sess.client_reads:
        read_cols.setdefault((s_, r_), set()).add(c_)
    idle = [k for k in missing if not _moves(sess, summary, keys, this, k, read_cols.get(k) or set())]
    missing = [k for k in missing if k not in idle]
    pvd = roll.get("prior_valuation_date")
    zero = this if date_hold else (_read(sess, summary, "current", keys, pvd, 0, rates=False)[0] if pvd else None)
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
                      "reliable": (not gone if mine else share >= FOUND_MIN) and reads > 0 and not blank_rows
                      and not weak_rows and not timing_open and not date_hold and not vd_off and zr["ok"]}

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
    # a rebuilt model (few line-item labels in common): the rows found other than by their own label are listed to
    # check, though the value stands where the checks pass (a pasted copy of last year's figures is never a match)
    fam = sess.rowmap.family()
    by_label = lambda k: any(n == "label" for n, _t in (ex(*k).get("evidence") or []))
    rebuilt_rows = [{"row": f"{s_}!r{r_}", "label": labels.get((s_, r_), ""), "found": found((s_, r_)),
                     "how": ex(s_, r_).get("how")} for s_, r_ in sorted(read_by_row)
                    if fam < ov.REBUILT and found((s_, r_)) and not by_label((s_, r_))]
    try:  # beside the value, not in its way
        lines = _new_lines(sess, sorted(read_by_row), None if date_hold else (info or {}).get("valuation_date"), pvd)
        lines_error = None
    except Exception as ex_:
        lines, lines_error = [], f"{type(ex_).__name__}: {ex_}"
    try:  # a term the value takes in that last year's model didn't have holds it until a person confirms it
        terms, gone = _term_changes(sess, sorted(read_by_row), None if date_hold else (info or {}).get("valuation_date"),
                                    pvd, set(summary.get("terms_confirmed") or []), {x["row"] for x in lines})
        terms_error = None
    except Exception as ex_:
        terms, gone, terms_error = [], [], f"{type(ex_).__name__}: {ex_}"
    terms_held = [x for x in terms + gone if x["hold"]]
    return {"reliable": all(x["reliable"] for x in by_cell.values()) and not terms_held, "by_cell": by_cell,
            "date_check": date_hold, "new_lines": lines, "new_lines_error": lines_error,
            "new_terms": terms, "gone_terms": gone, "new_terms_error": terms_error,
            "terms_held": [x["key"] for x in terms_held],
            "no_reads": reads == 0, "family": fam, "rebuilt": fam < ov.REBUILT, "rebuilt_rows": rebuilt_rows,
            "date_cells": {"moved": roll.get("valuation_date_cells") or [], "read": vd_reads, "off": vd_off,
                           "by_label": bool(roll.get("valuation_date_by_label")), "to": (info or {}).get("valuation_date")},
            "zero_roll_off": zero_off, "read_rows": [f"{s_}!r{r_}" for s_, r_ in sorted(read_by_row)],
            "timing": [{"row": f"{s_}!r{r_}", "label": labels.get((s_, r_), ""), "why": why, "found": found((s_, r_)),
                        "derived": (sess.derived.get((s_, r_)) or {}).get("text"), "open": (s_, r_) in timing_open}
                       for s_, r_, why in timing],
            "found_share": round(share, 3), "values": len(sess.unmatched), "reads": reads,
            "dcf_rows": len(origins), "dcf_origins": [f"{s_}!r{r_}" for s_, r_ in origins],
            "dcf_missing": [{"row": f"{s_}!r{r_}", "label": labels.get((s_, r_), ""), "why": why_missing((s_, r_))}
                            for s_, r_ in missing],
            "idle_rows": [{"row": f"{s_}!r{r_}", "label": labels.get((s_, r_), ""),
                           "why": "not found this year, but its figures don't move the value (tried): last year's stand in"}
                          for s_, r_ in idle],
            "blank_rows": [{"row": f"{s_}!r{r_}", "label": labels.get((s_, r_), ""),
                            "found": f"{blank_at[(s_, r_)][0]}!r{blank_at[(s_, r_)][1]}",
                            "blank": blank_by_row[(s_, r_)], "of": read_by_row.get((s_, r_), 0)} for s_, r_ in blank_rows],
            "weak_rows": [{"row": f"{s_}!r{r_}", "label": labels.get((s_, r_), ""), "found": found((s_, r_)),
                           "confidence": ex(s_, r_)["confidence"], "how": ex(s_, r_)["how"]} for s_, r_ in weak_rows],
            "timing_open": [{"row": f"{s_}!r{r_}", "label": labels.get((s_, r_), ""), "found": found((s_, r_))}
                            for s_, r_ in timing_open]}


# ---- this year's horizon -----------------------------------------------------------------------------------------

END_WORDS = re.compile(r"\bend\b|terminal|horizon|expir|concession|final|last|maturity", re.I)


def _last_flow(db, cores: list[dict]) -> date | None:
    """The end of the last period any of these discountings has a cash flow in (flags in), on db."""
    last = None
    for c in cores:
        try:
            flows, ends = ov._flows(db, c["inputs"])
        except (ValueError, KeyError):
            continue
        for col, v in flows.items():
            if v and isinstance(ends.get(col), date) and (last is None or ends[col] > last):
                last = ends[col]
    return last


def _label_left(db, s: str, r: int, c: int) -> str:
    """A cell's row label; where the layout kept none (an inputs sheet's one-date row read as a heading), the text
    typed to its left on the row."""
    lab = dcf._row_label(db, s, r)
    if lab:
        return lab
    return " ".join(v for (v,) in db.execute("SELECT value FROM cells WHERE sheet=? AND row=? AND col<? AND formula IS NULL "
                                              "ORDER BY col", (s, r, c)) if isinstance(v, str) and not dcf._as_date(v))


def _windows(db, trees: list[dict]) -> list[dict]:
    """The sums in these trees of a row up to a date typed in a cell (SUMIF(dates, "<"&end+1, present values), or
    SUMIFS with that criterion: dcftrace.windows): an overlay's own horizon, the periods after it left to its terminal
    value. -> [{"sum": (sheet, row, c1, c2), "dates": (sheet, row, c1, c2), "end": (sheet, row, col)}]."""
    out, names = [], dcftrace._names(db)

    def walk(n):
        f = n.get("formula") or ""
        if "SUMIF" in f.upper() and n.get("cell"):
            here = dcf._ref(n["cell"], "")[0]
            for call in dcftrace._calls(dcftrace._expand(db, f, names)):
                out.extend({"sum": sm, "dates": d, "end": e} for sm, d, e in dcftrace.windows(db, call, here))
        for ch in n.get("children") or []:
            walk(ch)
    for t in trees:
        walk(t)
    return out


def _window_last(db, w: dict) -> date | None:
    """The date of the last period a window's row has a value in, on db (inside the window or after it)."""
    s, r, c1, c2 = w["sum"]
    ds, dr, d1, d2 = w["dates"]
    vals = dcftrace._values(db, s, r, c1, c2)
    when = {c - d1 + c1: v for c, v in db.execute("SELECT col, value FROM cells WHERE sheet=? AND row=? AND col BETWEEN ? AND ?",
                                                    (ds, dr, d1, d2))}
    last = None
    for c, v in vals.items():
        x = when.get(c)
        d = ov.to_date(x) if isinstance(x, (int, float)) and 3000 < x < 120000 else dcf._as_date(x) if isinstance(x, str) else None
        if v and isinstance(d, date) and (last is None or d > last):
            last = d
    return last


def this_year_horizon(sess, summary: dict, where: dict) -> dict | None:
    """Where this year's model forecasts to another date than last year's (a rolling horizon that moved by other than
    the roll): the overlay's own copy of the forecast's end date goes to the end of the last period this year's cash
    flows reach, on this year's feed (summary["horizon_values"], read by overlay._feed). Its copy is a date typed on
    its sheets holding the end of the last period last year's cash flows reach: labelled as an end, a terminal or a
    horizon date; or the end a sum up to a date reads (SUMIF(dates, "<"&end+1, present values), the periods after it
    left to a terminal value), where last year's present values stopped on it. Without it, a terminal value placed by
    the overlay at its model's end stays where last year's model ended, and this year's last periods drop out.
    -> {"was", "now", "cells"} or None where nothing moves."""
    summary["horizon_values"] = {}
    path = summary["wiring"]["overlay"]["db_path"]
    db = rodb.connect(path)
    traced = [x for x in (_traced(db, where[e]) for e in ("low", "high")) if x]
    cores = [c for _, cs in traced for c in cs]
    wins = _windows(db, [t for t, _ in traced])
    if not cores and not wins:
        return None
    need = set().union(set(), *(ov._dcf_cells(db, c["inputs"]) for c in cores))
    for w in wins:
        for s, r, c1, c2 in (w["sum"], w["dates"]):
            need |= {(s, r, c) for c in range(c1, c2 + 1)}
    now_db = _patched(sess, summary, db, "current", need)
    sheets = set(summary["sheets"])
    moved = set(summary.get("roll", {}).get("valuation_date_cells") or [])
    out, was_now = {}, None
    was, now = (_last_flow(db, cores), _last_flow(now_db, cores)) if cores else (None, None)
    if was and now and was != now:
        iso = was.isoformat()
        for s, r, c, v in db.execute("SELECT sheet, row, col, value FROM cells WHERE formula IS NULL AND value=?", (iso,)):
            if s in sheets and END_WORDS.search(_label_left(db, s, r, c)) and ov._a1(s, r, c) not in moved:
                out[ov._a1(s, r, c)] = ov.serial(now)
                was_now = was_now or (was, now)
    for w in wins:
        root = ov._typed_in(db, *w["end"], sheets)
        s, r, c = ov.parse_a1(root)
        f, v = dcftrace._cell(db, s, r, c)
        end = dcf._as_date(v) if isinstance(v, str) and not f else None
        a, b = _window_last(db, w), _window_last(now_db, w)
        if end and a == end and b and b != a and s in sheets and root not in moved and root not in out:
            out[root] = ov.serial(b)
            was_now = was_now or (a, b)
    if not out:
        return None
    summary["horizon_values"] = out
    return {"was": was_now[0].isoformat(), "now": was_now[1].isoformat(), "cells": list(out)}


# ---- this year's discount rate -----------------------------------------------------------------------------------

def this_year_rate(summary: dict, asm: dict, facts: list[dict]) -> dict | None:
    """This year's discount rate where a person set it (workbench.set_this_year_rate: the low end at the higher rate),
    put on the cells each end's discountings read for it (sourced.rate, followed to the input), and so on this year's
    feed (summary["rate_values"], read by overlay._feed). All or nothing: where an end's rate isn't sourced to a cell,
    or the low and the high read one cell, it isn't applied, and says why. None where none is set.
    -> {"low", "high", "was" (last year's, by end), "cells" ({cell: this year's rate}), "applied", "why"}."""
    want = summary.get("this_year_rate") or {}
    summary["rate_values"] = {}
    if want.get("low") is None or want.get("high") is None:
        return None
    ends = sourced.rate(asm, facts)["ends"]
    cells, was, why = {}, {}, []
    for e in ("low", "high"):
        x = ends.get(e) or {}
        if not x.get("sourced"):
            why.append(f"the {e} end's discount rate isn't sourced to a cell")
            continue
        if x.get("ok") is False:  # last year's rate there doesn't check out against the report: not the cell to set
            why.append(f"the {e} end's discount rate ({x['cell']}) doesn't check out against the report's")
            continue
        was[e] = x["value"]
        for c in x.get("cells") or [x["cell"]]:
            if c in cells and abs(cells[c] - want[e]) > 1e-12:
                why.append(f"the low and the high read one discount rate cell ({c})")
            cells[c] = want[e]
    if why:
        cells = {}
    summary["rate_values"] = cells
    return {"low": want["low"], "high": want["high"], "was": was, "cells": cells, "applied": bool(cells),
            "why": "; ".join(why) or None, "by": want.get("by")}


def _range_pct(lo: float | None, hi: float | None) -> str:
    """Two ends' rates as the report prints a range, the lower first: 7.90% to 8.90%."""
    xs = sorted(x for x in (lo, hi) if x is not None)
    return " to ".join(f"{100 * x:.2f}%" for x in dict.fromkeys(xs))


def rate_label(rt: dict) -> str:
    was = rt.get("was") or {}
    return (f"Discount rate: this year's, {_range_pct(rt.get('low'), rt.get('high'))}"
            + (f" (last year's {_range_pct(was.get('low'), was.get('high'))})" if was else ""))


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


def _patched(sess, summary: dict, db, feed: str, need: set, roll: dict | None = None, rates: bool = True):
    """The overlay's model.db with the module's values on a feed in place (dcf.py and dcftrace read a model.db).
    rates=False: this year's feed at last year's discount rate (where a person set this year's)."""
    if feed == "current":
        defaults, _, months = ov._feed(summary, "current", None, None, rates=rates)
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
    every = dcftrace.cores(t)
    if len(cs) < len(every):  # its unwind can't be split out, and would land in the new forecast
        return [{"key": "roll", "label": "Roll-forward onto this year's model (time, cash flows and forecast)", "value": v2 - v0}], \
            (f"{len(every) - len(cs)} of the {len(every)} discountings under it can't be read here, so the roll-forward "
             "is one step")
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
        streams = f" (all {len(cs)} discounted streams)" if len(cs) > 1 else ""
        return [{"key": "time", "label": f"Time value: {years:.2f} years of unwind at the discount rate on the discounted "
                                         f"cash flows{streams}", "value": vu - v0},
                {"key": "cash", "label": f"Last year's forecast cash flows{streams} up to {vd1[:10]}", "value": v1 - vu},
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
            # the steps in the overlay's units, then shown in the report's; the inputs a person set and this year's
            # discount rate are steps of their own, after the roll-forward
            at_rate = (figs.get("this_year_last_rate") or figs["this_year"])[cell]
            before = (figs.get("this_year_held") or {}).get(cell, at_rate)
            raw, note = _steps(prior_db, traced, figs["rebuilt"][cell], before, vd1)
            steps += [{**x, "value": unit(x["value"])} for x in raw]
            if figs.get("this_year_held"):
                names = ", ".join(x.get("label") or c for c, x in (summary.get("held_values") or {}).items()
                                  if x.get("value") is not None)
                steps.append({"key": "held", "label": f"This year's figures for inputs held at last year's ({names})",
                              "value": unit(at_rate) - unit(before)})
            rt = figs.get("rate") or {}
            steps.append({"key": "rate", "label": rate_label(rt) if figs.get("this_year_last_rate") else
                          "Discount rate: last year's, unchanged", "value": v2 - unit(at_rate)})
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

def fy_year(d: date, end_month: int) -> int:
    """The financial year a period ending on d falls in (the year it ends in), as a four-digit year."""
    return d.year if d.month <= end_month else d.year + 1


def _fy(d: date, end_month: int) -> str:
    return f"FY{fy_year(d, end_month)}"


def _tv_parts(core: dict) -> list[dict]:
    return [p for p in core.get("parts") or []
            if TV_WORDS.search(p.get("label") or "") or re.search(r"/\s*\(.*-.*growth", p.get("words") or "", re.I)]


def _ranges(core: dict, parts: list[dict]) -> list[str] | None:
    """These parts of a discounting's cash-flow row as ranges over its columns (None if a part is on another sheet)."""
    rng = core["inputs"]["cashflow"]
    m = re.match(r"^(.+?)!\$?([A-Z]{1,3})\$?\d+:\$?([A-Z]{1,3})\$?\d+$", rng[0]) if len(rng) == 1 else None
    if not m:
        return None
    out = []
    for p in parts:
        pm = re.match(r"^(?:\[\d+\])?(.+)!r(\d+)$", p["row"])
        if not pm or pm[1] != m[1].strip("'"):
            return None
        out.append(f"{m[1]}!{m[2]}{pm[2]}:{m[3]}{pm[2]}")
    return out


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
    new date). The discounting with the largest value under the low's cell (the main cash flows, to equity or free
    cash flows, rather than a side stream like franking credits), its terminal value left out."""
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
    series, dropped = {}, {}
    for key, feed, vd in (("last_year", figs["feeds"]["rebuilt"], (summary.get("roll") or {}).get("prior_valuation_date")),
                          ("this_year", "current" if figs.get("this_year") else None, (figs.get("roll") or {}).get("valuation_date"))):
        if not feed:
            continue
        pdb = _patched(sess, summary, db, feed, need)
        flows, ends = ov._flows(pdb, inputs)
        vd_d = date.fromisoformat(vd[:10]) if vd else None
        by_fy, undated = {}, 0
        for c, v in flows.items():
            e = ends.get(c)
            e = e if isinstance(e, date) else None
            if e is None or e.year < 1950:  # no usable period date (a year number or a serial read as one)
                undated += 1 if v else 0
                continue
            if vd_d and e <= vd_d:
                continue
            by_fy[fy_year(e, fy_end)] = by_fy.get(fy_year(e, fy_end), 0.0) + v
        series[key] = {f"FY{y}": unit(v) for y, v in sorted(by_fy.items())}
        dropped[key] = undated
    years = [f"FY{y}" for y in sorted({int(y[2:]) for s in series.values() for y in s})]
    return {"years": years, "series": series, "rows": ranges, "left_out": left_out, "core": main["cell"], "undated": dropped,
            "label": next((p.get("label") for p in (main.get("parts") or []) if p.get("label") not in left_out), None),
            "held": bool(figs.get("gaps") and not figs["gaps"]["reliable"]), "units": None}


# ---- the report's disclosures, reconciled ------------------------------------------------------------------------
# What the report discloses of the value (the terminal value, the present values of the forecast and of the
# terminal value, the value of franking credits and its share of the equity value), against the same split of the
# overlay's own discountings: the one with the largest value (the main cash flows) split into its terminal value
# part and the rest, the others that are franking credits by their label. Each end separately; the mid their
# average, the convention for the equity value too. A figure ties when it's within half a unit of the report's
# last printed digit.

def _split(db, core: dict) -> dict:
    """One discounting on db: its present value, and where its cash-flow row has a terminal value part (added, not
    subtracted), that part undiscounted and discounted, and the rest (the discrete forecast)."""
    inputs = {**core["inputs"], "compare_to": None}
    r = dcf.compute(db, **inputs, fix=False)
    out = {"pv": r["pv"], "rate": r["rate"]}
    tv = _tv_parts(core)
    rng = _ranges(core, tv) if tv and all(p.get("sign", 1) == 1 for p in tv) else None
    if rng:
        t = dcf.compute(db, **{**inputs, "cashflow": rng}, fix=False)
        flows, _ = ov._flows(db, {"cashflow": rng, "dates": inputs.get("dates")})
        out.update(pv_tv=t["pv"], pv_forecast=r["pv"] - t["pv"], tv=sum(flows.values()))
    return out


def _streams(cs: list[dict]) -> tuple[dict | None, list[dict], list[dict]]:
    """(the main discounting, the franking credit ones, the others)."""
    if not cs:
        return None, [], []
    main = max(cs, key=lambda c: abs(c.get("pv") or 0.0))
    rest = [c for c in cs if c is not main]
    words = lambda c: " ".join([c.get("cashflow_label") or ""] + [p.get("label") or "" for p in c.get("parts") or []])
    fr = [c for c in rest if FRANKING.search(words(c))]
    return main, fr, [c for c in rest if c not in fr]


def _franking_part(db, summary: dict, facts: list[dict], cell: str, cs: list[dict]) -> dict | None:
    """Franking credits no discounting under the figure discounts on its own, found by tracing up from the report's
    utilisation to the figure (forward.py): the row added into a discounting's cash flows, and its sign there."""
    if _streams(cs)[1]:
        return None
    stated = sourced._stated(facts, "franking_utilisation")
    sheets = set(summary.get("sheets") or []) or None
    starts = forward.holding(db, [p for p, _ in stated], sourced.FRANKING, sheets) if stated else []
    fw = forward.franking(db, starts, cell, cs) if starts else None
    if not fw or fw["where"] != "part" or not fw["range"]:
        return None
    return {"core": fw["core"], "range": fw["range"], "sign": fw["sign"]}


def _tv_place(db, summary: dict, facts: list[dict], cell: str, cs: list[dict]) -> dict | None:
    """A terminal value the split can't see in the discounting's cash-flow rows, found by tracing up from the
    report's growth rate (else its exit multiple) to the figure (forward.py): inside the last cash flow, or added after
    the discounting. None where the cash-flow rows have it, or it isn't found."""
    main = _streams(cs)[0]
    if not main or _tv_parts(main):
        return None
    sheets = set(summary.get("sheets") or []) or None
    g, x = sourced._stated(facts, "terminal_growth_rate"), sourced._stated_x(facts)
    tries = [(forward.growth, forward.holding(db, [p for p, _ in g], sourced.GROWTH_WORDS, sheets) if g else []),
             (forward.multiple, forward.holding(db, [m for m, _ in x], sourced.MULTIPLE_WORDS, sheets, percent=False)
              if x else [])]
    for fn, starts in tries:
        fw = fn(db, starts, cell, cs) if starts else None
        if fw and ((fw["where"] == forward.LAST and fw["core"] == main["cell"] and fw["term"]) or
                   (fw["where"] == forward.AFTER and fw["pv_term"])):
            return {k: fw[k] for k in ("where", "core", "tv_cell", "term", "pv_cell", "pv_term")}
    return None


def _tv_cells(db, tv: dict) -> set:
    """The cells a terminal value placed by _tv_place is worked out from, to read on another feed."""
    out = set()
    for ref, term in ((tv["tv_cell"], tv["term"]), (tv["pv_cell"], tv["pv_term"])):
        if not ref:
            continue
        k = dcf._ref(ref, "")[:3]
        out.add(k)
        if term:
            out |= {(n["sheet"], n["row"], n["col"]) for n in valuation.reads(db, expr=term, here=k[0], depth=1).values()}
    return out


def _factor_at(db, core: dict, col: int) -> float:
    """A discounting's factor for one column of its cash flows: dcf.factors at its rate, date and convention, its
    flags in."""
    i = core["inputs"]
    sheet, _, cols = dcf._row_range(db, i["cashflow"][0])
    ends, _ = dcf.period_ends(db, sheet, cols, i.get("dates"))
    rate = dcf._num(dcf.resolve(db, i["rate"])[0])
    rate = rate / 100 if rate >= 1 else rate
    vd = dcf._as_date(dcf.resolve(db, i["valuation_date"])[0])
    td = dcf._as_date(dcf.resolve(db, i["terminal_date"])[0]) if i.get("terminal_date") else None
    f = dcf.factors(ends, vd, rate, i.get("timing") or "end", i.get("day_count") or "actual/actual", td).get(col, 0.0)
    for m in i.get("mask") or []:
        ms, mr, _ = dcf._row_range(db, m)
        f *= dcf._num(dcf._cell(db, ms, mr, col)) or 0.0
    return f


def _tv_split(db, main: dict, tv: dict, out: dict) -> None:
    """The split for a terminal value placed by _tv_place, into out: inside the last cash flow, its term there
    (undiscounted) times that column's factor; added after the discounting, the term that adds it (its present
    value), the discounting's own value then the forecast's."""
    k = dcf._ref(tv["tv_cell"], "")[:3]
    if tv["where"] == forward.LAST:
        t = float(dcftrace.evaluate(db, tv["term"], k[0]))
        p = t * _factor_at(db, main, k[2])
        out.update(tv=t, pv_tv=p, pv_forecast=out["pv"] - p)
    else:
        here = dcf._ref(tv["pv_cell"], "")[0]
        p = float(dcftrace.evaluate(db, tv["pv_term"], here))
        out.update(tv=dcf._num(dcf._cell(db, *k)), pv_tv=p, pv_forecast=out["pv"], pv=out["pv"] + p)


def _end_split(db, cs: list[dict], part: dict | None = None, tv: dict | None = None) -> dict:
    main, fr, other = _streams(cs)
    if not main:
        return {}
    pv = lambda c: dcf.compute(db, **{**c["inputs"], "compare_to": None}, fix=False)["pv"]
    out = _split(db, main)
    if tv and "pv_tv" not in out:  # inside the last cash flow, or added after the discounting (forward.py)
        _tv_split(db, main, tv, out)
    # a terminal value discounted on its own (its one period, at the forecast's end), its franking credits too: the
    # terminal value is theirs, undiscounted and discounted, and the franking credits the forecast's
    tvs = [c for c in fr + other if c.get("periods") == 1 and TV_WORDS.search(c.get("cashflow_label") or "")]
    if tvs and "pv_tv" not in out:
        flows = lambda c: ov._flows(db, {"cashflow": c["inputs"]["cashflow"], "dates": c["inputs"].get("dates")})[0]
        out.update(pv_tv=sum(pv(c) for c in tvs), tv=sum(sum(flows(c).values()) for c in tvs), pv_forecast=out["pv"])
        fr, other = [c for c in fr if c not in tvs], [c for c in other if c not in tvs]
    out["franking"] = sum(pv(c) for c in fr) if fr else None
    c = next((x for x in cs if part and x["cell"] == part["core"]), None)
    if c and not fr:  # franking credits added into a discounting's cash flows: that part, discounted with them
        fv = part["sign"] * dcf.compute(db, **{**c["inputs"], "cashflow": part["range"], "compare_to": None}, fix=False)["pv"]
        out["franking"] = fv
        if c is main and "pv_forecast" in out:
            out["pv_forecast"] -= fv
    out["other"] = [{"cell": c["cell"], "label": c.get("cashflow_label"), "pv": pv(c)} for c in other]
    out["main_label"] = main.get("cashflow_label")
    return out


def _ties(python: float | None, text: str | None) -> bool | None:
    nums = keyfacts.numbers(text or "")
    if python is None or not nums:
        return None
    x = float(nums[0].rstrip("%"))
    d = len(nums[0].rstrip("%").split(".")[1]) if "." in nums[0] else 0
    return abs(python - x) <= 0.5 * 10 ** -d + 1e-9


SPLIT_UNITS = (1.0, 1e-3, 1e-6, 1e3)  # the discountings' units against the figure's, the figure's first
REPORTED = (("terminal_value", "Terminal value", "tv"), ("pv_forecast", "PV of the discrete forecast", "pv_forecast"),
            ("pv_terminal_value", "PV of the terminal value", "pv_tv"),
            ("franking_credits_value", "Value of franking credits", "franking"),
            ("franking_credits_share", "Franking credits, % of the equity value", "franking_share"))


def reconcile(sess, summary: dict, where: dict, facts: list[dict], figs: dict) -> dict:
    """The report's disclosed split against Python's, last year (on the values Excel saved), with this year's split
    beside it (rolled forward). {"rows": [...], "last_year": {end: split}, "this_year": {end: split} | None}."""
    path = summary["wiring"]["overlay"]["db_path"]
    db = rodb.connect(path)
    unit = lambda v: v / (where["scale"] or 1.0) * (where["sign"] or 1) if isinstance(v, float) else None
    traced = {e: _traced(db, where[e]) for e in ("low", "high")}
    parts = {e: _franking_part(db, summary, facts, where[e], (traced[e] or (None, []))[1]) for e in ("low", "high")}
    tvs = {e: _tv_place(db, summary, facts, where[e], (traced[e] or (None, []))[1]) for e in ("low", "high")}
    last, this = {}, {}
    for e in ("low", "high"):
        cs = (traced[e] or (None, []))[1]
        try:
            last[e] = _end_split(db, cs, parts[e], tvs[e])
        except (ValueError, ZeroDivisionError) as ex:
            last[e] = {"error": str(ex)}
    if figs.get("this_year") and not (figs.get("gaps") and not figs["gaps"]["reliable"]):
        need = set()
        for e in ("low", "high"):
            for c in (traced[e] or (None, []))[1]:
                need |= ov._dcf_cells(db, c["inputs"])
                for rng in [_ranges(c, [p]) for p in c.get("parts") or []]:
                    if rng:
                        need |= ov._dcf_cells(db, {"cashflow": rng})
            if parts[e]:
                need |= ov._dcf_cells(db, {"cashflow": parts[e]["range"]})
            if tvs[e]:
                need |= _tv_cells(db, tvs[e])
        cur = _patched(sess, summary, db, "current", need)
        for e in ("low", "high"):
            try:
                this[e] = _end_split(cur, (traced[e] or (None, []))[1], parts[e], tvs[e])
            except (ValueError, ZeroDivisionError) as ex:
                this[e] = {"error": str(ex)}
    saved = {e: figs["saved"].get(where[e]) for e in ("low", "high")}
    thisv = {e: (figs.get("this_year") or {}).get(where[e]) for e in ("low", "high")}

    def view(split: dict, equity: dict, k: float = 1.0) -> dict:
        out = {}
        for e in ("low", "high"):
            x = split.get(e) or {}
            out[e] = {f: unit(x[f] * k) if isinstance(x.get(f), float) else None
                      for f in ("pv", "pv_forecast", "pv_tv", "tv", "franking")}
            eq = equity.get(e) if k == 1.0 else None  # in other units than the figure (a share of it, say): no share of it
            out[e]["franking_share"] = 100 * x["franking"] / eq if x.get("franking") is not None and isinstance(eq, float) \
                and eq else None
            out[e]["tv_share"] = 100 * x["pv_tv"] / x["pv"] if x.get("pv_tv") is not None and x.get("pv") else None
        out["mid"] = {k: (out["low"][k] + out["high"][k]) / 2 if out["low"][k] is not None and out["high"][k] is not None
                      else None for k in out["low"]}
        # a share at the mid is the mid's figure over the mid's total (as the report works it), not the average share
        m, eq = out["mid"], [unit(equity.get(e)) if k == 1.0 else None for e in ("low", "high")]
        eq_mid = (eq[0] + eq[1]) / 2 if None not in eq else None
        m["franking_share"] = 100 * m["franking"] / eq_mid if m["franking"] is not None and eq_mid else None
        m["tv_share"] = 100 * m["pv_tv"] / m["pv"] if m["pv_tv"] is not None and m["pv"] else None
        return out

    by_key = {}
    for f in facts:
        if f.get("status") != "rejected" and f.get("key") not in by_key:
            by_key[f["key"]] = f

    def rows_at(L: dict, T: dict | None) -> list[dict]:
        rows = []
        for key, label, field in REPORTED:
            f = by_key.get(key)
            row = {"key": key, "label": label, "python": {e: L[e][field] for e in ("low", "mid", "high")},
                   "this_year": {e: T[e][field] for e in ("low", "mid", "high")} if T else None,
                   "unit": "%" if field.endswith("share") else None}
            if f:
                v = f.get("final") or f
                texts = {"low": v.get("low_text"), "high": v.get("high_text"), "mid": v.get("value_text")}
                checks = {e: _ties(row["python"][e], t) for e, t in texts.items() if t and keyfacts.numbers(t)}
                row.update(report={e: t for e, t in texts.items() if t}, page=v.get("page"), ties=checks,
                           ok=all(checks.values()) if checks else None, basis=v.get("basis"))
            rows.append(row)
        return rows

    # the discountings can count in other units than the figure (thousands, under an equity value in millions, a
    # share of it): the report's split in the units most of its figures tie in, the figure's where none tie better
    best = None
    for k in SPLIT_UNITS:
        L, T = view(last, saved, k), view(this, thisv, k) if this else None
        rows = rows_at(L, T)
        n = sum(1 for r in rows if r["unit"] is None for t in (r.get("ties") or {}).values() if t)
        if best is None or n > best[0]:
            best = (n, k, L, T, rows)
    _, k, L, T, rows = best
    return {"rows": rows, "last_year": L, "this_year": T, "units": k,
            "streams": {e: {"main": last.get(e, {}).get("main_label"), "other": last.get(e, {}).get("other")}
                        for e in ("low", "high")}}


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
            m = c.get("method") or {}
            rows.append({"cell": c["cell"], "kind": c.get("kind"), "label": c.get("cashflow_label"), **ov._brief(r),
                         "form": cashflows.form(c),
                         "rate_source": c["inputs"].get("rate") if isinstance(c["inputs"].get("rate"), str) else
                         m.get("rate_note") or "a constant",
                         "rate_sighted": m.get("rate_source"), "rate_note": m.get("rate_note"),
                         "rate_cells": m.get("rate_cells") or [],
                         "valuation_date_source": c["inputs"].get("valuation_date"),
                         "valuation_date_sourced": m.get("valuation_date_sourced"),
                         "cashflow": c["inputs"]["cashflow"], "parts": [p.get("label") for p in c.get("parts") or []],
                         "terminal_value": tv[0]["total"] if tv else None, "terminal_value_row": tv[0]["row"] if tv else None})
        for c in (dcftrace.cores(traced[0]) if traced else []):
            L = c.get("loose")
            if c.get("inputs") or not L:
                continue
            # a convention the app doesn't recompute: its rate and date as its formulas read them, nothing recomputed
            rate = dcf._num(dcf._cell(db, *dcf._ref(L["rate"], "")[:3])) if L.get("rate") else L.get("rate_value")
            rows.append({"cell": c["cell"], "kind": c.get("kind"), "label": c.get("cashflow_label"), "rate": rate,
                         "form": cashflows.form(c), "period_months": L.get("period_months"), "period_rate": L.get("period_rate"),
                         "pv": c.get("pv"), "total": c.get("pv"), "anchor": None, "valuation_date": L.get("valuation_date_value"),
                         "timing": None, "day_count": None, "terminal_date": None, "bridge": [], "periods": c.get("periods"),
                         "first_period": c.get("first_period"), "last_period": c.get("last_period"),
                         "undiscounted": c.get("undiscounted"), "rate_source": L.get("rate"), "rate_sighted": L.get("rate_source"),
                         "rate_cells": L.get("rate_cells") or [],
                         "rate_note": None if L.get("rate") else "not sourced: no cell every period's factors read holds it",
                         "valuation_date_source": L.get("valuation_date"), "valuation_date_sourced": bool(L.get("valuation_date")),
                         "cashflow": [c["cashflow"]], "parts": [p.get("label") for p in c.get("parts") or []],
                         "terminal_value": None, "terminal_value_row": None, "loose": L["note"]})
        out[end] = rows
    return out


# ---- the inputs held at last year's ------------------------------------------------------------------------------

def held_list(sess, summary: dict, where: dict, figs: dict) -> list[dict]:
    """The inputs typed in the overlay outside its discountings (held.py), each with this year's figure where a person
    set it, else held at last year's, and the suggestion from this year's client model (checked against last
    year's). In the overlay's units, as typed."""
    cells = [(e, where[e]) for e in ("low", "high") if where.get(e)]
    with rodb.connect(summary["wiring"]["overlay"]["db_path"]) as db:
        items = held.find(db, cells, summary.get("sheets"))
    roll, set_ = figs.get("roll") or {}, summary.get("held_values") or {}
    prior_vd = (summary.get("roll") or {}).get("prior_valuation_date")
    out = []
    for it in items:
        mine = set_.get(it["cell"]) or {}
        out.append({**it, "suggestion": held.suggest(sess, it, prior_vd, roll.get("valuation_date")),
                    "this_year": mine.get("value"), "by": mine.get("by"), "from": mine.get("from"),
                    "held": mine.get("value") is None})
    sess.configure("workbook")
    return out


# ---- all of it --------------------------------------------------------------------------------------------------

def _flows_public(fl: dict, unit) -> dict:
    """The cash-flow layer as the page and the workpaper show it: per discounting, its form, and per period the cash
    flow last year and this year (the report's units), without the trace trees."""
    out = []
    for c in fl.get("cores") or []:
        x = {k: c.get(k) for k in ("cell", "kind", "what", "label", "ends", "form", "why", "tv_rows")}
        for key in ("last", "rolled", "this"):
            if c.get(key):
                x[key] = {**{k: v for k, v in c[key].items() if k != "periods"}, "pv": unit(c[key].get("pv")),
                          "periods": {e: {"cf": unit(p["cf"]), "tv": unit(p["tv"]), "factor": p["factor"]}
                                      for e, p in c[key]["periods"].items()}}
        out.append(x)
    return {"cores": out, "vd0": fl.get("vd0"), "vd1": fl.get("vd1"), "error": fl.get("error")}


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
    asm = assumptions(summary, where)
    rate_now = this_year_rate(summary, asm, facts)  # on this year's feed from here on, where it's applied
    horizon = ov.deep(this_year_horizon, sess, summary, where)  # and the overlay's own forecast end, where it moved
    figs = ov.deep(figures, sess, summary, cells)
    figs["rate"] = rate_now
    figs["horizon"] = horizon
    if where.get("basis") == "cum" and head["basis"] != "cum" and head.get("other"):
        head = {**head, **{k: head["other"][k] for k in ("low", "high")}, "basis": "cum",
                "texts": head["other"]["texts"], "why": head["why"] + ["the overlay gives the cum-distribution value"]}
        if head["low"] is not None and head["high"] is not None:
            head["mid"] = (head["low"] + head["high"]) / 2
        else:
            head["low"] = head["high"] = head["mid"] = head["other"]["value"]
    unit = lambda v: v / (where["scale"] or 1.0) * (where["sign"] or 1) if isinstance(v, float) else None
    import cashflows
    try:  # this year's cash flows against last year's, period by period, and the checks on them
        fl = ov.deep(cashflows.layer, sess, summary, where, figs)
        cfc = ov.deep(cashflows.checks, sess, summary, where, figs, fl, unit)
    except Exception as ex:  # beside the value: but the checks not running is itself a point to check
        fl, cfc = {"cores": [], "error": f"{type(ex).__name__}: {ex}"}, {"holds": [hold(
            summary, "cf-error", [type(ex).__name__], "The cash-flow checks couldn't run",
            f"{type(ex).__name__}: {ex}: this year's cash flows weren't compared with last year's", severity="check")],
            "split": {}}
    if figs.get("gaps") is not None:
        figs["gaps"]["holds"] = cfc["holds"]
        if holding(cfc["holds"]):
            figs["gaps"]["reliable"] = False
    tie = {}
    for end in ("low", "high"):
        text = head["texts"].get(f"{end}_text") or head["texts"].get("value_text")
        t = ov.tie(figs["saved"].get(where[end]), text, where["scale"], where["sign"]) if text else None
        tie[end] = {"report": head[end], "report_text": text, "saved": unit(figs["saved"].get(where[end])),
                    "rebuilt": unit(figs["rebuilt"].get(where[end])), "ok": bool(t and t["ok"]),
                    "rebuilt_ok": bool((ov.tie(figs["rebuilt"].get(where[end]), text, where["scale"], where["sign"]) or {})
                                       .get("ok")) if text else False}
    br = ov.deep(bridges, sess, summary, head, where, figs)
    try:
        rec = ov.deep(reconcile, sess, summary, where, facts, figs)
    except Exception as ex:  # the reconciliation is beside the value, not in its way
        rec = {"rows": [], "error": f"{type(ex).__name__}: {ex}"}
    try:
        ch = ov.deep(chart, sess, summary, where, figs, fy_end)
    except Exception as ex:  # the chart is beside the value, not in its way
        ch = {"why": f"couldn't draw the cash flows: {type(ex).__name__}: {ex}"}
    this = {e: unit((figs.get("this_year") or {}).get(where[e])) for e in ("low", "high")} if figs.get("this_year") else None
    with rodb.connect(summary["wiring"]["overlay"]["db_path"]) as db:
        traced = {e: (_traced(db, where[e]) or (None, []))[1] for e in ("low", "high")}
    import context
    terminal = keyfacts.terminal_method(facts, context.terminal(markdown))  # how the report works out its terminal value
    inputs = ov.deep(sourced.check, sess, summary, where, facts, asm, traced, unit, terminal)
    inputs["rate"]["this_year"] = rate_now  # last year's is sourced and checked; this year's is a person's
    held_inputs = ov.deep(held_list, sess, summary, where, figs)
    try:  # the scenario each client model was saved on, and when: beside the value, not in its way
        scenario = ov.deep(scenarios.settings, sess, summary)
    except Exception as ex:
        scenario = {"selectors": [], "saved": {}, "error": f"{type(ex).__name__}: {ex}"}
    this_year = ({**this, "mid": (this["low"] + this["high"]) / 2 if this["low"] is not None and this["high"] is not None
                  else None} if this and not br["held"] else None)
    inv = None
    if this_year and this_year["mid"] is not None:  # this year's value worked out the other ways (methods.py)
        import methods
        try:
            inv = ov.deep(methods.inventory, sess, summary, where, unit, rate_now, summary.get("method"))
            this_year = methods.apply(br, inv) or this_year  # the preferred method's, with its own bridge step
        except Exception as ex:  # beside the value, not in its way
            inv = {"methods": [], "default": methods.DEFAULT, "preferred": methods.DEFAULT, "asked": summary.get("method"),
                   "error": f"{type(ex).__name__}: {ex}"}
    return {"head": head, "where": where, "tie": tie, "figures": figs, "bridges": br, "chart": ch, "reconcile": rec,
            "flows": _flows_public(fl, unit), "flow_checks": {"split": cfc.get("split") or {}, "rows": cfc.get("rows")},
            "assumptions": asm, "inputs": inputs, "held": held_inputs, "terminal": terminal, "methods": inv,
            "scenario": scenario,
            "values": {"report": {e: head[e] for e in ("low", "mid", "high")},
                       "rebuilt": {"low": tie["low"]["rebuilt"], "high": tie["high"]["rebuilt"],
                                   "mid": (tie["low"]["rebuilt"] + tie["high"]["rebuilt"]) / 2
                                   if tie["low"]["rebuilt"] is not None and tie["high"]["rebuilt"] is not None else None},
                       "this_year": this_year}}
