"""This year's equity value worked out more than one way: the methods inventory. A recurring valuation can be rolled
forward several defensible ways (which periods count, when in a period a cash flow lands, how the year is counted,
what the mid is), and the solved answer of a given engagement follows one of them. Each method here is worked out on
the same feed (this year's client model, rolled forward, at this year's discount rate where a person set it), for the
low and the high; the mid is their average, as for the value, except where the method is about the mid:

  overlay          the overlay's own formulas rolled forward (the default): every discounting's valuation date moved,
                   its periods ending on or before the new date cut off (overlay.Session._cut_off)
  overlay_on_date  the same, the period ending on the new date kept in, undiscounted (where a model's cash flow for the
                   period to the valuation date is received after it)
  own_flags        the overlay's own forecast flags rolled, in place of the cut-off: a typed 0 / 1 row on an overlay
                   sheet the cash flows read, 0 up to last year's date and 1 after it, set to 0 up to the new date (what
                   a valuer does by hand). The default's figure where the flags are the only way in; listed so the two
                   are seen to agree
  recompute        each discounting recomputed in code (dcf.compute) from this year's cash flows, as the overlay
                   discounts (its timing and day count), and the formulas above it again (dcftrace.recompute): the
                   default's figure, to the cent, where the inventory's own discounting is the overlay's
  mid_period       recomputed at the middle of each period
  mid_year         recomputed with the mid-year convention (half a year before each period's end)
  day_count        recomputed on the other day count (actual/actual for actual/365, and the other way)
  mid_rate         the mid with both ends at the midpoint of the two rates, rather than each at its own rate

A method that can't be worked out here says why. The preferred method (a person's, else the default) is this year's
value: the bridge then has a step of its own for the move from the default to it (apply).
"""
import dcf
import dcftrace
import overlay as ov
import valuation

DEFAULT = "overlay"
METHODS = (
    ("overlay", "The overlay's own formulas, rolled forward",
     "Every discounting's valuation date moved; its periods ending on or before the new date cut off"),
    ("overlay_on_date", "The overlay's own formulas, the period ending on the new date kept",
     "As the default, with the cash flow for the period ending on the valuation date in the value, undiscounted"),
    ("own_flags", "The overlay's own forecast flags rolled",
     "In place of the cut-off, the overlay's own 0 / 1 period flags set to 0 up to the new date, as a valuer does by hand"),
    ("recompute", "Recomputed: the overlay's convention",
     "Each discounting recomputed from this year's cash flows, as the overlay discounts, and the formulas above it again"),
    ("mid_period", "Recomputed: mid-period", "Each cash flow discounted from the middle of its period"),
    ("mid_year", "Recomputed: mid-year", "Each cash flow discounted from half a year before its period's end"),
    ("day_count", "Recomputed: the other day count", "Years counted the other way (actual/actual or actual/365)"),
    ("mid_rate", "The mid at the midpoint rate", "The mid with both ends at the midpoint of the two discount rates, "
                                                "rather than each end at its own rate"),
)
LABEL = {k: label for k, label, _ in METHODS}


def _ends(sess, summary: dict, where: dict, unit, extra: dict | None = None, keep_on_date: bool = False,
          cutoffs: list | None = None) -> dict:
    """The low's and the high's equity value on this year's feed, with extra overrides, the period ending on the new
    date kept or not, and another cut-off (None: the session's). The session is left as it was."""
    defaults, _, months = ov._feed(summary, "current", None, None)
    keep = sess.keep_on_date, sess.cutoffs
    sess.keep_on_date = keep_on_date
    if cutoffs is not None:
        sess.cutoffs = cutoffs
    try:
        sess.configure("current", {**defaults, **(extra or {})}, months or 0)
        vals = sess.values([ov.parse_a1(where[e]) for e in ("low", "high")])
    finally:
        sess.keep_on_date, sess.cutoffs = keep
        sess.configure("workbook")
    return {e: unit(v) if isinstance(v, float) else None for e, v in zip(("low", "high"), vals)}


def own_flags(sess, db, cores: list[dict]) -> list[tuple]:
    """The overlay's own forecast flags under the discountings: a row typed on an overlay sheet that their cash flows
    read, 0 or 1 in every period it has, nothing but 0 for the periods to last year's valuation date and 1 for the
    first after it. -> [(sheet, row)]"""
    base, out = sess.base_vd, []
    if base is None:
        return out
    for c in cores:
        try:
            sh, r, cols = dcf._row_range(db, c["inputs"]["cashflow"][0])
        except (ValueError, KeyError, IndexError):
            continue
        sample = cols[:: max(1, len(cols) // 8)]
        for n in valuation.reads(db, cells=[(sh, r, col) for col in sample], depth=4).values():
            s_, r_ = n["sheet"], n["row"]
            if n["formula"] or s_ not in sess.sheets or (s_, r_) in out:
                continue
            tl = sess.ov.timeline(s_)
            nums = {col: v for col in tl if isinstance(v := sess.ov.value(s_, r_, col), (int, float))
                    and not isinstance(v, bool)}
            if len(nums) < 2 or set(nums.values()) - {0.0, 1.0}:
                continue
            after = sorted((tl[col], nums.get(col)) for col in tl if tl[col] > base)
            if all(nums.get(col) in (0.0, None) for col in tl if tl[col] <= base) and after and after[0][1] == 1.0:
                out.append((s_, r_))
    return out


def _flags_rolled(sess, summary: dict, flags: list[tuple]) -> dict:
    """The flags for this year's feed: 0 for the periods ending on or before the new date, last year's after it."""
    _, _, months = ov._feed(summary, "current", None, None)
    keep = sess.shift
    sess.shift = months or 0
    try:
        new_vd = ov.add_months(sess.base_vd, months or 0)
        out = {}
        for s_, r_ in flags:
            step = sess.period_shift(sess.ov, s_)
            for col, d in sess.ov.timeline(s_).items():
                v = sess.ov.value(s_, r_, col)
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    out[(s_, r_, col)] = 0.0 if (ov.add_months(d, step) if step else d) <= new_vd else float(v)
        return out
    finally:
        sess.shift = keep


def _recomputed(sess, summary: dict, db, where: dict, unit, traced: dict) -> dict:
    """Each recompute method's low and high: every discounting under the end recomputed (dcf.compute) on this year's
    feed, the formulas above it again. {method: {"low", "high"} or {"why"}}."""
    import result
    out = {k: {} for k in ("recompute", "mid_period", "mid_year", "day_count")}
    for e in ("low", "high"):
        t, cs = traced.get(e) or (None, [])
        every = dcftrace.cores(t) if t else []
        if not cs or len(cs) < len(every):
            why = "no discounting under it can be recomputed here" if not cs else \
                f"{len(every) - len(cs)} of the {len(every)} discountings under it can't be read here"
            for k in out:
                out[k][e] = None
                out[k]["why"] = why
            continue
        cur = result._patched(sess, summary, db, "current", result._need(db, t, cs))
        for k in out:
            totals = {}
            try:
                for c in cs:
                    inp = {**c["inputs"], "compare_to": None}
                    if k == "mid_period":
                        inp["timing"] = "mid"
                    elif k == "mid_year":
                        inp["timing"] = "mid-year"
                    elif k == "day_count":
                        inp["day_count"] = "actual/actual" if inp.get("day_count") == "actual/365" else "actual/365"
                    totals[c["cell"]] = dcf.compute(cur, **inp, fix=False)["total"]
                v = dcftrace.recompute(cur, t, totals)
            except (ValueError, ZeroDivisionError, OverflowError) as ex:
                v, out[k]["why"] = None, f"couldn't recompute: {ex}"
            if v is None and not out[k].get("why"):
                out[k]["why"] = "a formula above the discountings can't be evaluated here"
            out[k][e] = unit(v) if isinstance(v, float) else None
    return out


def inventory(sess, summary: dict, where: dict, unit, rate_now: dict | None, preferred: str | None = None) -> dict:
    """Every method's low, mid and high, against the default's: {"methods": [{"key", "label", "what", "low", "mid",
    "high", "vs_default" ({"low", "mid", "high"}), "ok", "why"}], "default", "preferred", "ties" (the recompute
    reproduces the overlay's own formulas), "flags" (the overlay's own forecast flags found)}. Runs in overlay.deep."""
    import result
    import rodb
    db = rodb.connect(summary["wiring"]["overlay"]["db_path"])
    traced = {e: result._traced(db, where[e]) for e in ("low", "high")}
    got = {"overlay": _ends(sess, summary, where, unit),
           "overlay_on_date": _ends(sess, summary, where, unit, keep_on_date=True)}
    why = {}
    cores = [c for e in ("low", "high") for c in ((traced[e] or (None, []))[1])]
    flags = own_flags(sess, db, cores)
    if flags:
        got["own_flags"] = _ends(sess, summary, where, unit, _flags_rolled(sess, summary, flags), cutoffs=[])
    else:
        why["own_flags"] = "the overlay has no forecast flags of its own under its discountings"
    for k, x in _recomputed(sess, summary, db, where, unit, traced).items():
        if x.get("low") is not None and x.get("high") is not None:
            got[k] = {"low": x["low"], "high": x["high"]}
        else:
            why[k] = x.get("why") or "not worked out"
    # the mid at the midpoint rate: both ends' rate cells at it
    cells = (rate_now or {}).get("cells") if (rate_now or {}).get("applied") else None
    ends = {}
    if cells:
        mid_r = ((rate_now["low"]) + (rate_now["high"])) / 2
        ends = {c: mid_r for c in cells}
    else:
        rate = sourced_rates(sess, summary, where)
        if rate:
            mid_r = (rate["low"]["value"] + rate["high"]["value"]) / 2
            ends = {c: mid_r for e in ("low", "high") for c in rate[e]["cells"]}
    if ends:
        # each end at the midpoint rate, averaged: what else sets the ends apart (an exit multiple) stays averaged
        at = _ends(sess, summary, where, unit, {ov.parse_a1(c): v for c, v in ends.items()})
        if at["low"] is not None and at["high"] is not None:
            got["mid_rate"] = {"low": None, "high": None, "mid": (at["low"] + at["high"]) / 2}
    else:
        why["mid_rate"] = "the discount rate isn't sourced to a cell at both ends"
    base = got["overlay"]
    base_mid = (base["low"] + base["high"]) / 2 if None not in (base["low"], base["high"]) else None
    rows = []
    for key, label, what in METHODS:
        x = got.get(key)
        if not x:
            rows.append({"key": key, "label": label, "what": what, "ok": False, "why": why.get(key),
                         "low": None, "mid": None, "high": None, "vs_default": None})
            continue
        mid = x.get("mid") if "mid" in x else (x["low"] + x["high"]) / 2 if None not in (x["low"], x["high"]) else None
        vs = {"low": x["low"] - base["low"] if x["low"] is not None else None, "mid": mid - base_mid
              if mid is not None and base_mid is not None else None,
              "high": x["high"] - base["high"] if x["high"] is not None else None}
        rows.append({"key": key, "label": label, "what": what, "ok": True, "why": None, "low": x["low"], "mid": mid,
                     "high": x["high"], "vs_default": vs})
    rec = got.get("recompute")
    ties = bool(rec) and all(abs(rec[e] - base[e]) <= 1e-6 * max(1.0, abs(base[e])) for e in ("low", "high"))
    pick = preferred if preferred in got else DEFAULT
    return {"methods": rows, "default": DEFAULT, "preferred": pick, "asked": preferred, "ties": ties,
            "flags": [f"{s_}!r{r_}" for s_, r_ in flags]}


def sourced_rates(sess, summary: dict, where: dict) -> dict | None:
    """Last year's discount rate at each end, sourced to its cells (sourced.rate), or None."""
    import result
    import sourced
    asm = result.assumptions(summary, where)
    ends = sourced.rate(asm, [])["ends"]
    if all((ends.get(e) or {}).get("sourced") for e in ("low", "high")):
        return {e: {"value": ends[e]["value"], "cells": ends[e].get("cells") or [ends[e]["cell"]]} for e in ("low", "high")}
    return None


def apply(bridges: dict, inv: dict) -> dict | None:
    """The preferred method's figures, where it isn't the default: a bridge step of its own before this year's total,
    for the low, the mid and the high (the mid at the midpoint rate moves the mid only). -> {"low", "mid", "high"} of
    this year's value, or None where the default stands."""
    key = inv["preferred"]
    if key == DEFAULT:
        return None
    m = next(x for x in inv["methods"] if x["key"] == key)
    vs = m["vs_default"]
    out = {}
    for e in ("low", "mid", "high"):
        steps = bridges[e]["steps"]
        if not steps or steps[-1]["key"] != "this_year":
            return None
        move = vs[e] if vs[e] is not None else 0.0  # the mid at the midpoint rate leaves the ends where they are
        steps.insert(len(steps) - 1, {"key": "method", "label": f"Method: {m['label'].lower()}", "value": move})
        steps[-1] = {**steps[-1], "value": steps[-1]["value"] + move}
        out[e] = steps[-1]["value"]
    return out
