"""The models side by side: one table comparing what each says, the way a shop compares products' specifications.

Columns: last year's report; last year's overlay as Excel saved it; last year rebuilt in Python on last year's client
model; last year's client model; this year's client model; this year rolled forward in Python. Rows: the valuation
date, the equity value (low, mid, high) and its basis, the interest valued, the discount rate, the terminal value's
basis and growth or multiple, franking, how the cash flows are discounted, the cash-flow periods, the undiscounted
cash flows and the last period's, the terminal value and its present value, the present value of the forecast, the
scenario and when each model was saved.

Only what's sighted: a client model gives its rows the value reads (their periods and figures, in its own units), its
date, its scenario and its save time, never a value worked out for it. The last-year columns should agree with each
other: a row where they don't is flagged; this year against last year is the move, shown, not flagged.
    build(...) -> {"columns": [...], "sections": [{"title", "rows": [{"label", "cells": [...], "flag", "note"}]}]}
Each cell: {"v": the value, "t": "num" / "n" (a count) / "pct" / "date" / "text" / "x" (a multiple), "src": where it's
from} or None.
"""
import statistics
from datetime import date

import overlay as ov

COLUMNS = (("report", "Last year's report"), ("overlay", "Last year's overlay, as saved"),
           ("python_last", "Last year, rebuilt in Python"), ("client_last", "Last year's client model"),
           ("client_this", "This year's client model"), ("python_this", "This year, rolled forward"))
LAST = (0, 1, 2)  # the columns that should agree: the report, the overlay as saved and the rebuild
AGREE = 0.001     # numbers in those columns this far apart, relative, are flagged


def _c(v, t="num", src=None):
    return None if v is None else {"v": v, "t": t, "src": src}


def _freq(serials: list[float]) -> str | None:
    if len(serials) < 2:
        return None
    gap = statistics.median(b - a for a, b in zip(serials, serials[1:]))
    return "monthly" if gap < 45 else "quarterly" if gap < 120 else "semi-annual" if gap < 250 else "annual" \
        if gap < 450 else "irregular"


def _client(sess, origins: list[str], vd: str | None, which: str) -> dict:
    """The client rows the value reads, in a client model (last year's or, by the row finder, this year's): their
    periods after its valuation date, summed: {"total", "last", "first_end", "last_end", "periods", "frequency"}."""
    rm = getattr(sess, "rowmap", None)
    if not rm or not origins:
        return {}
    wbk = rm.prior if which == "last" else rm.current
    since = ov.serial(date.fromisoformat(vd[:10])) if vd else None
    by = {}
    for ref in origins:
        s, r = ref.rsplit("!r", 1)
        at = (s, int(r)) if which == "last" else rm.locate(s, int(r))
        if not at:
            continue
        for w, v in rm._series(wbk, *at).items():
            if since is None or w > since:
                by[w] = by.get(w, 0.0) + v
    if not by:
        return {}
    ws = sorted(by)
    return {"total": sum(by.values()), "last": by[ws[-1]], "first_end": ov.to_date(ws[0]).isoformat(),
            "last_end": ov.to_date(ws[-1]).isoformat(), "periods": len(ws), "frequency": _freq(ws)}


def _periods(ps: dict, after: str | None) -> dict:
    """A discounting's periods after a date (the layer's): totals without the terminal value, the last period's."""
    es = sorted(e for e, p in ps.items() if (not after or e > after[:10]) and p.get("factor"))
    if not es:
        return {}
    disc = [ps[e]["cf"] - ps[e]["tv"] for e in es]
    return {"total": sum(disc), "last": disc[-1], "first_end": es[0], "last_end": es[-1], "periods": len(es),
            "frequency": _freq([ov.serial(date.fromisoformat(e)) for e in es]),
            "tv": sum(ps[e]["tv"] for e in es)}


def build(sess, summary: dict, facts: list[dict], head: dict, where: dict, tie: dict, figs: dict, flows: dict,
          asm: dict, inputs: dict, rec: dict, scenario: dict, inter: dict, this_year: dict | None, rate_now: dict | None,
          terminal: dict | None, methods: dict | None, unit=lambda v: v) -> dict:
    fact = {f["key"]: (f.get("final") or f) for f in facts if f.get("status") != "rejected"}
    said = lambda k: (lambda f: (f.get("value_text") or " – ".join(x for x in (f.get("low_text"), f.get("high_text")) if x))
                      if f else None)(fact.get(k))
    w = summary.get("wiring") or {}
    vd0, vd1 = (summary.get("roll") or {}).get("prior_valuation_date"), (figs.get("roll") or {}).get("valuation_date")
    main = next((c for c in flows.get("cores") or [] if c.get("last") and c.get("this")), None) or \
        next(iter(flows.get("cores") or []), None)
    A = next((a for a in (asm.get("low") or []) if not a.get("error") and a.get("parts")), None) or \
        next((a for a in (asm.get("low") or []) if not a.get("error")), None) or {}
    AH = next((a for a in (asm.get("high") or []) if not a.get("error") and a.get("parts")), None) or \
        next((a for a in (asm.get("high") or []) if not a.get("error")), None) or {}
    origins = ((figs.get("gaps") or {}).get("dcf_origins")) or []
    cl_last, cl_this = _client(sess, origins, vd0, "last"), _client(sess, origins, vd1, "this")
    p_last = _periods((main or {}).get("last", {}).get("periods") or {}, vd0) if main else {}
    p_this = _periods((main or {}).get("this", {}).get("periods") or {}, vd1) if main else {}
    p_saved = {"total": A.get("undiscounted"), "first_end": A.get("first_period"), "last_end": A.get("last_period"),
               "periods": A.get("periods")}
    ry, ty = rec.get("last_year") or {}, rec.get("this_year") or {}
    rmid, tmid = ry.get("mid") or {}, ty.get("mid") or {}
    rows = []

    def row(section, label, cells, note=None, kind="num"):
        cells = list(cells) + [None] * (len(COLUMNS) - len(cells))
        nums = [c["v"] for i, c in enumerate(cells) if i in LAST and c and isinstance(c["v"], (int, float))
                and c["t"] == kind and kind in ("num", "pct", "x", "n")]
        flag = len(nums) > 1 and (max(nums) - min(nums)) > AGREE * max(1.0, max(abs(x) for x in nums))
        dates = [c["v"] for i, c in enumerate(cells) if i in LAST and c and c["t"] == "date"]
        flag = flag or len(set(dates)) > 1
        rows.append((section, {"label": label, "cells": cells, "flag": flag, "note": note}))

    rt = (rate_now or {}) if (rate_now or {}).get("applied") else {}
    # the value
    row("The value", "Valuation date", [_c(_iso_fact(fact.get("valuation_date")), "date", "the report"),
                                        _c(A.get("valuation_date"), "date", A.get("valuation_date_source")),
                                        _c((main or {}).get("last", {}).get("valuation_date"), "date"),
                                        _c((w.get("prior") or {}).get("valuation_date"), "date", "its own date"),
                                        _c((w.get("current") or {}).get("valuation_date"), "date", "its own date"),
                                        _c(vd1, "date", "rolled to")], kind="date")
    for e, lab in (("low", "Equity value, low"), ("mid", "Equity value, mid"), ("high", "Equity value, high")):
        saved = (tie.get(e) or {}).get("saved") if e != "mid" else (
            ((tie["low"]["saved"] + tie["high"]["saved"]) / 2) if tie.get("low", {}).get("saved") is not None
            and tie.get("high", {}).get("saved") is not None else None)
        reb = (tie.get(e) or {}).get("rebuilt") if e != "mid" else (
            ((tie["low"]["rebuilt"] + tie["high"]["rebuilt"]) / 2) if tie.get("low", {}).get("rebuilt") is not None
            and tie.get("high", {}).get("rebuilt") is not None else None)
        row("The value", lab, [_c(head.get(e), "num", "the report"), _c(saved, "num", where.get(e) or where.get("low")),
                               _c(reb, "num"), None, None, _c((this_year or {}).get(e), "num")],
            note="the mid is the midpoint of the low and the high" if e == "mid" else None)
    basis = {"ex": "ex-distribution", "cum": "cum-distribution"}
    row("The value", "Basis", [_c(basis.get(head.get("basis") or "", "not stated (taken as ex)"), "text", "the report"),
                               _c(basis.get(where.get("cell_basis") or ""), "text", where.get("cell_basis_why")), None,
                               None, None, _c(basis.get(head.get("basis") or "ex") + (
                                   ", the period ending on the date kept" if (methods or {}).get("preferred") == "overlay_on_date"
                                   else ""), "text")], kind="text")
    m = (inter or {}).get("model") or {}
    row("The value", "Interest valued", [_c((inter or {}).get("report"), "pct", (inter or {}).get("report_text")),
                                         _c(m.get("value", 1.0 if inter else None), "pct",
                                            f"{m['cell']} ({m['where']} the discounting)" if m else "no share applied"),
                                         None, None, None, None], kind="pct")
    # the assumptions
    rates = sorted(x for x in (_pct_end(fact.get("discount_rate"), "low"), _pct_end(fact.get("discount_rate"), "high"))
                   if x is not None)
    row("Assumptions", "Discount rate, for the low value", [_c(rates[-1] if rates else None, "pct", said("discount_rate")),
                                                  _c(A.get("rate"), "pct", A.get("rate_source")), _c(A.get("rate"), "pct"),
                                                  None, None, _c(rt.get("low") or A.get("rate"), "pct",
                                                                 "this year's" if rt else "last year's")], kind="pct")
    row("Assumptions", "Discount rate, for the high value", [_c(rates[0] if rates else None, "pct", said("discount_rate")),
                                                   _c(AH.get("rate"), "pct", AH.get("rate_source")), _c(AH.get("rate"), "pct"),
                                                   None, None, _c(rt.get("high") or AH.get("rate"), "pct",
                                                                  "this year's" if rt else "last year's")], kind="pct")
    row("Assumptions", "Terminal value basis", [_c((terminal or {}).get("label") or said("terminal_value_method"), "text",
                                                   (terminal or {}).get("from"))], kind="text")
    g = ((inputs.get("growth") or {}).get("ends") or {}).get("low") or {}
    grows = sorted(x for x in (_pct_end(fact.get("terminal_growth_rate"), "low"),
                               _pct_end(fact.get("terminal_growth_rate"), "high")) if x is not None)
    row("Assumptions", "Terminal growth rate, for the low value", [_c(grows[0] if grows else None, "pct",
                                                                      said("terminal_growth_rate")), _c(g.get("value"), "pct", g.get("cell")),
                                                _c(g.get("value"), "pct"), None, None, _c(g.get("value"), "pct")], kind="pct")
    x = ((inputs.get("multiple") or {}).get("ends") or {}).get("low") or {}
    if x.get("value") is not None or fact.get("terminal_multiple"):
        row("Assumptions", "Exit multiple", [_c(_num_fact(fact.get("terminal_multiple")), "x", said("terminal_multiple")),
                                             _c(x.get("value"), "x", x.get("cell")), _c(x.get("value"), "x"), None, None,
                                             _c(x.get("value"), "x")], kind="x")
    fr = ((inputs.get("franking") or {}).get("ends") or {}).get("low") or {}
    if fr.get("value") is not None or fact.get("franking_utilisation"):
        row("Assumptions", "Franking credit utilisation", [_c(_pct_end(fact.get("franking_utilisation"), "low"), "pct",
                                                              said("franking_utilisation")),
                                                           _c(fr.get("value"), "pct", fr.get("cell")), _c(fr.get("value"), "pct"),
                                                           None, None, _c(fr.get("value"), "pct")], kind="pct")
    # the discounting
    form = (main or {}).get("form") or {}
    row("Discounting", "How the cash flows are discounted", [None, _c(f"{form.get('function', '')} · {form.get('convention', '')}"
                                                                      if form else None, "text", (main or {}).get("cell")),
                                                             _c(form.get("convention") if form.get("recomputed") else None, "text"),
                                                             None, None, _c(form.get("rolled"), "text")], kind="text")
    for lab, k, t in (("Cash-flow periods", "periods", "n"), ("Frequency", "frequency", "text"),
                      ("First period ending", "first_end", "date"), ("Last period ending", "last_end", "date")):
        row("Discounting", lab, [None, _c(p_saved.get(k), t) if k != "frequency" else None, _c(p_last.get(k), t),
                                 _c(cl_last.get(k), t, "the rows the value reads"), _c(cl_this.get(k), t, "the rows the value reads"),
                                 _c(p_this.get(k), t)], kind=t)
    # the cash flows
    note = "client models in their own units; the rows the value reads, summed" if cl_last or cl_this else None
    row("Cash flows", "Undiscounted forecast cash flows (no terminal value)",
        [None, None, _c(p_last.get("total"), "num", (main or {}).get("cell")), _c(cl_last.get("total"), "num"),
         _c(cl_this.get("total"), "num"), _c(p_this.get("total"), "num")], note=note)
    row("Cash flows", "The last period's cash flow", [None, None, _c(p_last.get("last"), "num"), _c(cl_last.get("last"), "num"),
                                                      _c(cl_this.get("last"), "num"), _c(p_this.get("last"), "num")])
    row("Cash flows", "Terminal value (the mid)", [_c(_num_fact(fact.get("terminal_value")), "num", said("terminal_value")),
                                         _c(unit((A["terminal_value"] + AH["terminal_value"]) / 2)
                                            if isinstance(A.get("terminal_value"), float) and isinstance(AH.get("terminal_value"), float)
                                            else None, "num", A.get("terminal_value_row")),
                                         _c(rmid.get("tv"), "num"), None, None, _c(tmid.get("tv"), "num")])
    row("Cash flows", "Present value of the terminal value (the mid)", [_c(_num_fact(fact.get("pv_terminal_value")), "num",
                                                                said("pv_terminal_value")), None, _c(rmid.get("pv_tv"), "num"),
                                                             None, None, _c(tmid.get("pv_tv"), "num")])
    row("Cash flows", "Present value of the forecast (the mid)", [_c(_num_fact(fact.get("pv_forecast")), "num", said("pv_forecast")),
                                                       None, _c(rmid.get("pv_forecast"), "num"), None, None,
                                                       _c(tmid.get("pv_forecast"), "num")])
    if rmid.get("franking") is not None or fact.get("franking_credits_value"):
        row("Cash flows", "Value of franking credits (the mid)", [_c(_num_fact(fact.get("franking_credits_value")), "num",
                                                           said("franking_credits_value")), None,
                                                        _c(rmid.get("franking"), "num"), None, None,
                                                        _c(tmid.get("franking"), "num")])
    # the models
    sel = [s for s in (scenario or {}).get("selectors") or []]
    if sel:
        s0 = sel[0]
        row("The models", f"Scenario ({s0.get('label') or s0.get('cell')})",
            [None, None, None, _c(s0.get("last_year"), "text", s0.get("last_cell")), _c(s0.get("this_year"), "text", s0.get("cell"))],
            kind="text")
    saved = (scenario or {}).get("saved") or {}
    row("The models", "Saved", [None, None, None, _c((saved.get("last_year") or {}).get("date"), "date"),
                                _c((saved.get("this_year") or {}).get("date"), "date")], kind="date")
    row("The models", "File", [None, _c((w.get("overlay") or {}).get("filename"), "text"), None,
                               _c((w.get("prior") or {}).get("filename"), "text"),
                               _c((w.get("current") or {}).get("filename"), "text")], kind="text")
    sections = []
    for sec, r in rows:
        if not sections or sections[-1]["title"] != sec:
            sections.append({"title": sec, "rows": []})
        if any(c for c in r["cells"]):
            sections[-1]["rows"].append(r)
    return {"columns": [{"key": k, "label": l} for k, l in COLUMNS], "sections": [s for s in sections if s["rows"]]}


def _iso_fact(f: dict | None) -> str | None:
    v = (f or {}).get("value")
    if isinstance(v, (int, float)) and v > 19000000:
        v = int(v)
        return f"{v // 10000:04d}-{v // 100 % 100:02d}-{v % 100:02d}"
    return None


def _pct_end(f: dict | None, end: str) -> float | None:
    """A % fact's end (low / high, else its one figure) as a fraction."""
    import keyfacts
    if not f:
        return None
    txt = f.get(f"{end}_text") or f.get("value_text") or ""
    n = keyfacts.numbers(keyfacts._unrange(txt))
    if not n:
        return None
    x = float(n[0].rstrip("%").replace(",", ""))
    return x / 100 if n[0].endswith("%") or x > 1 else x


def _num_fact(f: dict | None) -> float | None:
    import keyfacts
    if not f:
        return None
    txt = f.get("value_text") or f.get("low_text") or ""
    n = keyfacts.numbers(keyfacts._unrange(txt))
    try:
        return float(n[0].rstrip("%x").replace(",", "")) if n else None
    except ValueError:
        return None
