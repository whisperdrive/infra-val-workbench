"""The model's inputs under the equity value, sighted, sourced and checked again, not inferred.

For each end of last year's equity value (the low and the high), three inputs the report also states:
  the discount rate            the cell the discount factors read (valuation.read_factors follows their formulas
                               from the first and the last period's factor to it)
  the terminal growth rate     the cell the terminal value's formula reads as g, in X x (1 + g) / (r - g)
  franking credit utilisation  the fraction every period's franking credits read (the first and the last period's)
Each one is:
  sourced   found by following the formulas, not by looking for the number. A figure typed into a formula, or a
            cell found only because it holds the same number, isn't sourced: it's shown as not sourced and a person
            is asked
  sighted   the cell's value, its line item and column heading, and whether it's typed in or a formula (followed
            through cells that only pass it on, to the input)
  checked   recomputed from it: the discount factors at the rate; the terminal value from its growth and rate, and
            that rate the end's own discount rate; the equity value with utilisation at nil, which must fall by
            exactly the value of franking credits (the Python overlay, rerun). Then against the report's figure, at
            the precision the report prints it
"""
import re

import dcf
import keyfacts
import overlay as ov
import valuation

FRANKING = re.compile(r"frank|imputation|gamma|utili[sz]", re.I)
GROWTH_WORDS = re.compile(r"growth|\btgr\b|terminal|perpetu", re.I)
_LIT = re.compile(r"(?<![\w.$])(\d*\.\d+|\d+)(%?)(?![\w.])")


def _pct(v: float) -> str:
    return f"{100 * v:.2f}%"


def _as_pct(n: str) -> float:
    """A rate as printed, in percent: 7.25% is 7.25, and a fraction (0.80, a gamma) is 80."""
    x = float(n.rstrip("%"))
    return x if n.endswith("%") or x > 1 else 100 * x


def _stated(facts: list[dict], key: str) -> list[tuple[float, str]]:
    """The report's figures for a fact, in percent, lowest first: [(number, its text)]."""
    f = next((x for x in facts if x.get("key") == key and x.get("status") != "rejected"), None)
    v = (f.get("final") or f) if f else {}
    said = [(_as_pct(n), t) for t in (v.get("low_text"), v.get("high_text")) if t
            for n in keyfacts.numbers(keyfacts._unrange(t))[:1]]
    if not said and v.get("value_text"):
        said = [(_as_pct(n), n) for n in keyfacts.numbers(keyfacts._unrange(v["value_text"]))]
    return sorted(set(said))


def _tie(python_pct: float | None, rep) -> bool | None:
    """Does a rate (in percent) round to the report's, printed as a % or a fraction? Within half a unit of its last
    digit, in the units it's printed in."""
    if python_pct is None or not rep:
        return None
    n = (keyfacts.numbers(keyfacts._unrange(rep[1])) or [""])[0]
    if not n:
        return None
    d = len(n.rstrip("%").split(".")[1]) if "." in n else 0
    unit = 1.0 if n.endswith("%") or float(n.rstrip("%")) > 1 else 100.0
    return abs(python_pct - rep[0]) <= 0.5 * 10 ** -d * unit + 1e-9


def _key(ref: str) -> tuple:
    return dcf._ref(ref, "")[:3]


def _formula(db, ref: str) -> str | None:
    sh, r, c = _key(ref)
    got = db.execute("SELECT formula FROM cells WHERE sheet=? AND row=? AND col=?", (sh, r, c)).fetchone()
    return got[0] if got else None


def sight(db, ref: str, chain: list[str] | None = None) -> dict:
    """What a person sees at a cell: its line item, its column heading, and whether it's typed in."""
    sh, r, c = _key(ref)
    f = _formula(db, ref)
    return {"cell": ref, "input": not f, "formula": f, "label": dcf._row_label(db, sh, r),
            "heading": valuation._heading(db, sh, r, c), "chain": chain or [ref]}


def follow(db, ref: str, value: float) -> dict:
    """A cell, followed through the cells that only pass the same value on, to the input: sighted."""
    chain = [ref]
    for _ in range(6):
        if not _formula(db, chain[-1]):
            break
        w = valuation.reads(db, cells=[_key(chain[-1])], depth=1)
        same = [k for k, n in w.items() if dcf._num(n["value"]) is not None and abs(dcf._num(n["value"]) - value) < 1e-12]
        if len(same) != 1 or same[0] in chain:
            break
        chain.append(same[0])
    return sight(db, chain[-1], chain)


def _end(value, sighted, sourced, note, checks, report=None, ties=None, **more) -> dict:
    return {"value": value, "cell": sighted["cell"] if sighted else None, "sighted": sighted, "sourced": sourced,
            "note": note, "checks": [{"ok": ok, "text": t} for ok, t in checks], "report": report, "ties": ties,
            "ok": bool(sourced) and all(ok is not False for ok, _ in checks), **more}


def _vs_report(stated: list, rep, ties, what: str, side: str | None = None) -> tuple:
    if not rep:
        return None, f"The report's {what}: not among the key facts"
    of = f" (the {side} of {' and '.join(t for _, t in stated)})" if side and len(stated) > 1 else ""
    return ties, f"The report's {rep[1]}{of}" + (": doesn't match" if ties is False else "")


def _wrap(ends: dict, stated: list) -> dict:
    oks = [e["ok"] for e in ends.values() if e.get("ok") is not None]
    return {"ends": ends, "report": [t for _, t in stated], "ok": all(oks) if oks else None}


# ---- the discount rate ---------------------------------------------------------------------------------------------

def rate(asm: dict, facts: list[dict]) -> dict:
    """Each end's discount rate: the cell the factors read (sourced in valuation.read_factors), checked."""
    stated = _stated(facts, "discount_rate")
    want = {"low": stated[-1] if stated else None, "high": stated[0] if stated else None}  # the low value: the higher rate
    ends = {}
    for end in ("low", "high"):
        rows = [r for r in asm.get(end) or [] if not r.get("error") and r.get("rate") is not None]
        if not rows:
            ends[end] = {"ok": None, "why": "no discounting under it could be recomputed"}
            continue
        top = next((r for r in rows if r.get("parts")), rows[0])
        s = top.get("rate_sighted")
        sourced = all(r.get("rate_sighted") for r in rows)
        # the cell the factors read it from, and every input holding it they read (a blend of two rates, alike)
        cells = sorted({r["rate_sighted"]["cell"] for r in rows if r.get("rate_sighted")}
                       | {c for r in rows if r.get("rate_sighted") for c in r.get("rate_cells") or []})
        note = None if sourced else next(r.get("rate_note") for r in rows if not r.get("rate_sighted"))
        more = f" ({len(rows)} discountings" + (f", reading {' and '.join(cells)}" if len(cells) > 1 else "") + ")" \
            if len(rows) > 1 else ""
        rep = want[end]
        ties = _tie(100 * top["rate"], rep) if rep else None
        loose = any(r.get("loose") for r in rows)  # a convention the app doesn't recompute: the factors aren't
        ends[end] = _end(top["rate"], s, sourced, note, [
            (sourced, f"Sourced: the discount factors read this cell, the first and the last period's alike{more}"
             if sourced else f"Not sourced to a cell: {note}"),
            (None, f"The factors step by {_pct(top['rate'])} a year, but their convention isn't one the app recomputes, so "
                   "they aren't recomputed at it") if loose else
            (True, f"Every discount factor recomputed at {_pct(top['rate'])} matches the model's"),
            _vs_report(stated, rep, ties, "rate", "higher" if end == "low" else "lower")],
            rep[1] if rep else None, ties, cells=cells, discountings=len(rows))
    return _wrap(ends, stated)


# ---- the terminal growth rate --------------------------------------------------------------------------------------

def _fit(db, ref: str, tv: float, term: str | None = None) -> dict | None:
    """A terminal value's formula as X x (1 + g) / (r - g) (or X / (r - g)): g, r and X among the cells it reads,
    or figures typed into it (then not sourced). term: one term of the cell's formula, where the terminal value is
    added to the last cash flow in it (=W8 + W8 * (1 + g) / (r - g)), tv that term's value."""
    f = re.sub(r'"[^"]*"', "", term or _formula(db, ref) or "")
    w = valuation.reads(db, expr=f, here=_key(ref)[0], depth=1) if term else valuation.reads(db, cells=[_key(ref)], depth=1)
    cells = [(k, dcf._num(n["value"])) for k, n in w.items() if dcf._num(n["value"]) is not None]
    typed = [(None, float(m[1]) / (100 if m[2] else 1)) for m in _LIT.finditer(dcf._FREF.sub(" ", f))]
    small = [(k, v) for k, v in cells + typed if 0 <= v < 0.3]
    for gk, g in small:
        for rk, r in small:
            if (rk == gk and rk) or r <= g:
                continue
            for xk, x in cells:
                if xk in (gk, rk) or not x:
                    continue
                for grown, val in ((True, x * (1 + g) / (r - g)), (False, x / (r - g))):
                    if abs(val - tv) <= 1e-6 * max(1.0, abs(tv)):
                        sh, row, _ = _key(xk)
                        return {"tv_cell": ref, "tv": tv, "g": (gk, g), "r": (rk, r), "grown": grown,
                                "x": (xk, x, dcf._row_label(db, sh, row))}
    return None


def _terminal(db, core: dict, fit) -> dict | None:
    """The terminal value in a discounting's cash flows (its part labelled or built like one), fitted by fit(db, ref,
    tv): the cell itself, else a cell it reads holding the same value (a terminal value worked out elsewhere)."""
    import result
    try:
        sh, _, cols = dcf._row_range(db, core["inputs"]["cashflow"][0])
    except (ValueError, KeyError, IndexError):
        return None
    for p in result._tv_parts(core):
        m = re.match(r"^(.+)!r(\d+)$", p.get("row") or "")
        if not m or m[1].startswith("["):
            continue
        for c in reversed(cols):
            got = db.execute("SELECT formula, value FROM cells WHERE sheet=? AND row=? AND col=?",
                             (m[1], int(m[2]), c)).fetchone()
            tv = dcf._num(got[1]) if got else None
            if not tv or not got[0]:
                continue
            ref = f"{m[1]}!{dcf._addr(c, int(m[2]))}"
            for k in [ref] + [k for k, n in valuation.reads(db, cells=[(m[1], int(m[2]), c)], depth=3).items()
                              if n["formula"] and dcf._num(n["value"]) is not None and abs(dcf._num(n["value"]) - tv) < 1e-9]:
                got_fit = fit(db, k, tv)
                if got_fit:
                    return got_fit
            break  # the last non-zero terminal value cell is the one
    return None


def _gordon(db, core: dict) -> dict | None:
    """The terminal value in a discounting's cash flows, fitted as X x (1 + g) / (r - g)."""
    return _terminal(db, core, _fit)


def _steps(fw: dict) -> str:
    return " → ".join(f"{s['label'] or s['cell']} ({s['cell']})" for s in fw["path"][1:])


def growth(db, traced: dict, rates: dict, facts: list[dict], where: dict | None = None,
           sheets: set | None = None) -> dict:
    """Each end's terminal growth rate: the cell its terminal value's formula reads as g, checked. Found down from the
    discounting's cash flows, and up from the report's figure to the equity value (forward.py), which also finds a
    terminal value in the last cash flow itself or added after the discounting."""
    import forward
    import result
    stated = _stated(facts, "terminal_growth_rate")
    want = {"low": stated[0] if stated else None, "high": stated[-1] if stated else None}  # the low value: the lower growth
    starts = forward.holding(db, [p for p, _ in stated], GROWTH_WORDS, sheets) if stated and where else []
    ends = {}
    for end in ("low", "high"):
        rep = want[end]
        cs = traced.get(end) or []
        main, _, _ = result._streams(cs)
        fit = _gordon(db, main) if main else None
        fw = forward.growth(db, starts, where[end], cs) if starts and where.get(end) else None
        up = None
        if fw and fit:
            same = fw["tv_cell"] == fit["tv_cell"]
            up = (same, f"Traced up from {fw['cell']} to the equity value, through the same terminal value: {_steps(fw)}"
                  if same else f"Traced up from {fw['cell']}, the equity value is reached through {fw['tv_cell']}, not "
                               f"{fit['tv_cell']}: {_steps(fw)}")
        elif fw:
            fit = fw["fit"]
            up = (True, f"Found by tracing up from the report's figure ({fw['cell']}) to the equity value: the terminal "
                        f"value {fw['tv_cell']} is {fw['where']} ({_steps(fw)})")
        if not fit:
            ends[end] = {"ok": None, "why": "no terminal value built as X × (1 + g) / (r − g) found in its cash flows"
                         + (", nor on the way up from the report's figure to the equity value" if starts else "")}
            continue
        (gk, g), (rk, r), (xk, x, xl) = fit["g"], fit["r"], fit["x"]
        s = follow(db, gk, g) if gk else None
        r_in = follow(db, rk, r)["cell"] if rk else None
        own = (rates.get("ends") or {}).get(end) or {}
        by_cell = bool(own.get("cell") and r_in)
        if by_cell:
            at_rate = r_in == own["cell"]
        else:  # one of the two isn't a cell: only their values can be compared, and the card says so
            at_rate = own.get("value") is not None and abs(r - own["value"]) < 1e-12
        ties = _tie(100 * g, rep) if rep else None
        form = f"{xl or xk} × (1 + {_pct(g)}) / ({_pct(r)} − {_pct(g)})" if fit["grown"] else \
            f"{xl or xk} / ({_pct(r)} − {_pct(g)})"
        ends[end] = _end(g, s, bool(gk), None if gk else f"typed into the terminal value's formula ({fit['tv_cell']})", [
            (bool(gk), f"Sourced: the terminal value's formula ({fit['tv_cell']}) reads this cell" if gk else
             f"Not sourced: the growth rate is typed into the terminal value's formula ({fit['tv_cell']})"),
            (True, f"The terminal value recomputed as {form} = {fit['tv']:,.1f} matches the model's"),
            (at_rate if by_cell else None if at_rate else False,
             f"The terminal value is at this end's discount rate ({own['cell']})" if at_rate and by_cell else
             f"The terminal value's rate ({r_in or 'typed in'}) is {_pct(r)}, the same number as this end's discount "
             f"rate ({own.get('cell') or 'not sourced'}), matched by value only: one of them isn't a cell" if at_rate else
             f"The terminal value uses {r_in or 'a rate typed in'} ({_pct(r)}), not the rate the discount factors read "
             f"({own.get('cell') or '?'}, {_pct(own['value']) if own.get('value') is not None else '?'})"),
            *([up] if up else []),
            _vs_report(stated, rep, ties, "terminal growth rate")], rep[1] if rep else None, ties,
            terminal_value=fit["tv_cell"], traced_up={k: fw[k] for k in ("cell", "tv_cell", "where", "path")} if fw else None,
            base={"cell": xk, "label": xl, "value": x, "typed": bool(xk) and not _formula(db, xk)})
    return _wrap(ends, stated)


# ---- the exit multiple ---------------------------------------------------------------------------------------------

MULTIPLE_WORDS = re.compile(r"multiple|exit|ev\s*/|\bx\b|times", re.I)
METRIC = {"exit_ebitda": ("EBITDA", re.compile(r"ebitda", re.I)),
          "exit_rab": ("the RAB", re.compile(r"\brab\b|regulat\w* asset|asset base", re.I))}


def _fit_multiple(db, ref: str, tv: float, term: str | None = None) -> dict | None:
    """A terminal value's formula as M x X: the multiple M among the cells it reads (or typed into it: then not
    sourced) and the metric X (a cell), M the one labelled like a multiple, else the smaller. term: one term of the
    cell's formula (a terminal value added to the last cash flow in it), tv that term's value."""
    f = re.sub(r'"[^"]*"', "", term or _formula(db, ref) or "")
    w = valuation.reads(db, expr=f, here=_key(ref)[0], depth=1) if term else valuation.reads(db, cells=[_key(ref)], depth=1)
    cells = [(k, dcf._num(n["value"])) for k, n in w.items() if dcf._num(n["value"])]
    typed = [(None, float(m[1])) for m in _LIT.finditer(dcf._FREF.sub(" ", f)) if not m[2] and float(m[1])]
    label = lambda k: dcf._row_label(db, *_key(k)[:2]) if k else ""
    fits = []
    for mk, m in cells + typed:
        for xk, x in cells:
            if xk == mk or not 0 < m < 100 or abs(m * x - tv) > 1e-6 * max(1.0, abs(tv)):
                continue
            fits.append((not MULTIPLE_WORDS.search(label(mk) or ""), m, mk, xk, x))
    if not fits:
        return None
    _, m, mk, xk, x = min(fits)
    return {"tv_cell": ref, "tv": tv, "m": (mk, m), "x": (xk, x, label(xk))}


def _stated_x(facts: list[dict]) -> list[tuple[float, str]]:
    """The report's exit multiple, lowest first: [(number, its text)] (a multiple is as printed, 12.0x is 12.0)."""
    f = next((x for x in facts if x.get("key") == "terminal_multiple" and x.get("status") != "rejected"), None)
    v = {**(f or {}), **((f or {}).get("final") or {})}
    said = [(float(n), t) for t in (v.get("low_text"), v.get("high_text")) if t
            for n in keyfacts.numbers(keyfacts._unrange(t))[:1]]
    if not said and v.get("value_text"):
        said = [(float(n), n) for n in keyfacts.numbers(keyfacts._unrange(v["value_text"]))]
    return sorted(set(said))


def _tie_x(m: float | None, rep) -> bool | None:
    if m is None or not rep:
        return None
    n = (keyfacts.numbers(keyfacts._unrange(rep[1])) or [""])[0]
    d = len(n.split(".")[1]) if "." in n else 0
    return abs(m - rep[0]) <= 0.5 * 10 ** -d + 1e-9


def multiple(db, traced: dict, facts: list[dict], terminal: dict, where: dict | None = None,
             sheets: set | None = None) -> dict:
    """Each end's exit multiple, where the report's terminal value is one: the cell the terminal value's formula
    reads as the multiple in M x X, the terminal value recomputed from it, the metric X what the report says it's a
    multiple of (EBITDA, the RAB), and the report's multiple for that end (the low value at the lower multiple).
    Found down from the discounting's cash flows, and up from the report's figure to the equity value (forward.py)."""
    import forward
    import result
    stated = _stated_x(facts)
    want = {"low": stated[0] if stated else None, "high": stated[-1] if stated else None}
    of, like = METRIC.get(terminal.get("kind"), (None, None))
    starts = forward.holding(db, [m for m, _ in stated], MULTIPLE_WORDS, sheets, percent=False) \
        if stated and where else []
    ends = {}
    for end in ("low", "high"):
        rep = want[end]
        cs = traced.get(end) or []
        main, _, _ = result._streams(cs)
        fit = _terminal(db, main, _fit_multiple) if main else None
        fw = forward.multiple(db, starts, where[end], cs) if starts and where.get(end) else None
        up = None
        if fw and fit:
            same = fw["tv_cell"] == fit["tv_cell"]
            up = (same, f"Traced up from {fw['cell']} to the equity value, through the same terminal value: {_steps(fw)}"
                  if same else f"Traced up from {fw['cell']}, the equity value is reached through {fw['tv_cell']}, not "
                               f"{fit['tv_cell']}: {_steps(fw)}")
        elif fw:
            fit = fw["fit"]
            up = (True, f"Found by tracing up from the report's figure ({fw['cell']}) to the equity value: the terminal "
                        f"value {fw['tv_cell']} is {fw['where']} ({_steps(fw)})")
        if not fit:
            ends[end] = {"ok": None, "why": f"no terminal value built as a multiple × a metric found in its cash flows"
                                           + (", nor on the way up from the report's figure to the equity value"
                                              if starts else "") + f"; the report's is "
                                           f"{terminal['label'][0].lower() + terminal['label'][1:]}"}
            continue
        (mk, m), (xk, x, xl) = fit["m"], fit["x"]
        s = follow(db, mk, m) if mk else None
        ties = _tie_x(m, rep) if rep else None
        checks = [(bool(mk), f"Sourced: the terminal value's formula ({fit['tv_cell']}) reads this cell as the multiple"
                   if mk else f"Not sourced: the multiple is typed into the terminal value's formula ({fit['tv_cell']})"),
                  (True, f"The terminal value recomputed as {m:,.2f}x × {x:,.1f} ({xl or xk}) = {fit['tv']:,.1f} matches "
                         f"the model's")]
        if like is not None:
            checks.append((bool(like.search(xl or "")), f"A multiple of {of}, as the report says: the metric is {xl or xk}"
                           if like.search(xl or "") else f"The report says a multiple of {of}; the metric here is "
                                                         f"{xl or xk}"))
        if up:
            checks.append(up)
        checks.append((ties, f"The report's {rep[1]}" + (f" (the {'lower' if end == 'low' else 'higher'} of "
                                                         f"{' and '.join(t for _, t in stated)})" if len(stated) > 1 else "")
                       + (": doesn't match" if ties is False else "")) if rep else
                      (None, "The report's exit multiple: not among the key facts"))
        ends[end] = _end(m, s, bool(mk), None if mk else f"typed into the terminal value's formula ({fit['tv_cell']})",
                         checks, rep[1] if rep else None, ties, terminal_value=fit["tv_cell"],
                         metric={"cell": xk, "label": xl, "value": x},
                         traced_up={k: fw[k] for k in ("cell", "tv_cell", "where", "path")} if fw else None)
    return _wrap(ends, stated)


# ---- franking credit utilisation -----------------------------------------------------------------------------------

def _utilisation(db, cores: list[dict]) -> tuple[dict | None, str | None]:
    """The fraction every period's franking credits read (the first and the last period's with any), followed to
    its input: (sighted, None) or (None, why not)."""
    found = {}
    for c in cores:
        try:
            sh, row, cols = dcf._row_range(db, c["inputs"]["cashflow"][0])
        except (ValueError, KeyError, IndexError):
            continue
        live = [col for col in cols if dcf._num(dcf._cell(db, sh, row, col))]
        if not live:
            continue
        walks = [valuation.reads(db, cells=[(sh, row, col)]) for col in dict.fromkeys((live[0], live[-1]))]
        for k in set.intersection(*(set(w) for w in walks)):
            v = dcf._num(walks[0][k]["value"])
            if v is not None and 0 < v <= 1:
                s = follow(db, k, v)
                found[s["cell"]] = (s, v)
    if not found:
        return None, "no fraction is read by every period's franking credits"
    named = [x for x in found.values() if FRANKING.search(x[0]["label"] or "")]
    if len(named) == 1:
        return named[0][0], None
    if not named and len(found) == 1:
        s = next(iter(found.values()))[0]
        return {**s, "why": "the only fraction every period's franking credits read (its line item doesn't say)"}, None
    return None, "several fractions every period's franking credits read: " + ", ".join(sorted(found))


def franking(sess, summary: dict, db, traced: dict, where: dict, facts: list[dict], unit) -> dict:
    """Each end's franking credit utilisation: the fraction the franking credits read, checked by rerunning the
    Python overlay with it at nil (runs in overlay.deep; the session is left on the workbook feed)."""
    import forward
    import result
    stated = _stated(facts, "franking_utilisation")
    want = {"low": stated[0] if stated else None, "high": stated[-1] if stated else None}  # the low value: the lower utilisation
    starts = forward.holding(db, [p for p, _ in stated], FRANKING, set(summary.get("sheets") or []) or None) \
        if stated else []
    ends = {}
    for end in ("low", "high"):
        rep = want[end]
        cs = traced.get(end) or []
        _, fr, _ = result._streams(cs)
        fw = forward.franking(db, starts, where[end], cs) if starts else None
        up = None
        if fr:
            s, why = _utilisation(db, fr)
            fr_pv = result._end_split(db, cs).get("franking")
            if fw and s:
                same = follow(db, fw["cell"], dcf._num(dcf._cell(db, *_key(fw["cell"]))) or 0.0)["cell"] == s["cell"]
                up = (same, f"Traced up from {fw['cell']} to the equity value: {_steps(fw)}" if same else
                      f"Traced up from the report's figure, {fw['cell']} reaches the equity value, not {s['cell']}")
        elif fw and fw.get("pv") is not None:
            v0 = dcf._num(dcf._cell(db, *_key(fw["cell"])))
            s, why, fr_pv = follow(db, fw["cell"], v0 or 0.0), None, fw["pv"]
            what = (f"added into {fw['core']}'s cash flows and discounted with them" if fw["where"] == "part" else
                    f"discounted by {fw['core']}" if fw["where"] == "own" else
                    "applied to a present value of the gross credits")
            up = (True, f"Found by tracing up from the report's figure ({fw['cell']}) to the equity value: the franking "
                        f"credits used, {fw['used_label'] or fw['used_row']} ({fw['used_row']}), are "
                        + (f"{fw['gross_label'] or fw['gross_row']} ({fw['gross_row']}) × the utilisation, " if fw["gross"]
                           else "") + f"{what} ({_steps(fw)})")
        else:
            ends[end] = {"ok": None, "why": "no franking credits discounting found under it" + (
                ", and the report's figure, traced up, doesn't reach the equity value through franking credits it can "
                "value" if starts else "")}
            continue
        value = dcf._num(dcf._cell(db, *_key(s["cell"]))) if s else None
        rerun = None
        if s and fr_pv is not None:
            eq = ov.parse_a1(where[end])
            try:
                sess.configure("workbook")
                base = dcf._num(sess.B.get("", *eq))
                sess.configure("workbook", overrides={_key(s["cell"]): 0.0})
                nil = dcf._num(sess.B.get("", *eq))
            finally:
                sess.configure("workbook")
            after = dcf._num(sess.B.get("", *eq))
            if None not in (base, nil, after):
                drop = base - nil
                rerun = {"drop": unit(drop), "franking": unit(fr_pv), "restored": abs(after - base) < 1e-9,
                         "ok": abs(drop - fr_pv) <= 1e-6 * max(1.0, abs(fr_pv))}
        ties = _tie(100 * value, rep) if rep and value is not None else None
        checks = [(bool(s), ("Sourced: every period's franking credits read this cell, the first and the last period's "
                             "alike" if fr else "Sourced: the franking credits used read this cell, and through them the "
                                                "equity value") + (f"; {s['why']}" if s and s.get("why") else "")
                   if s else f"Not sourced: {why}")]
        if up:
            checks.append(up)
        if rerun:
            checks.append((rerun["ok"], f"Rerun with it at nil, Python's equity value at this end falls by "
                                        f"{rerun['drop']:,.1f}" + (", the value of franking credits" if rerun["ok"] else
                                                                   f", not the value of franking credits ({rerun['franking']:,.1f})")))
        checks.append(_vs_report(stated, rep, ties, "utilisation"))
        ends[end] = _end(value, s, bool(s), why, checks, rep[1] if rep else None, ties, rerun=rerun,
                         traced_up={k: fw[k] for k in ("cell", "used", "used_row", "used_label", "gross_row",
                                                       "gross_label", "where", "core", "path")} if fw else None)
    return _wrap(ends, stated)


EXITS = {"exit_ebitda", "exit_rab", "exit_other"}
NO_GROWTH = {"none"} | EXITS  # terminal values with no growth rate to source


def growth_applies(terminal: dict | None) -> str | None:
    """Why the terminal growth rate isn't one to source, where the report's terminal value has none (none at all, or
    an exit multiple); None where it applies."""
    if terminal and terminal.get("kind") in NO_GROWTH:
        return (f"not applicable: the report's terminal value is {terminal['label'][0].lower() + terminal['label'][1:]}"
                + (f" (p. {terminal['page']})" if terminal.get("page") else ""))
    return None


def check(sess, summary: dict, where: dict, facts: list[dict], asm: dict, traced: dict, unit,
          terminal: dict | None = None) -> dict:
    """{"rate", "growth", "franking"}: each {"ends": {"low", "high"}, "report", "ok"}. Runs in overlay.deep. terminal:
    how the report works out its terminal value (keyfacts.terminal_method): no growth rate to source where it's an
    exit multiple or there's none."""
    db = summary["wiring"]["overlay"]["db_path"]
    import rodb
    db = rodb.connect(db)
    out = {"rate": rate(asm, facts)}
    na = growth_applies(terminal)
    jobs = [("growth", lambda: {"ends": {e: {"ok": None, "why": na} for e in ("low", "high")}, "report": [],
                                "ok": None, "na": na} if na else growth(db, traced, out["rate"], facts, where,
                                                                        set(summary.get("sheets") or []) or None))]
    if (terminal or {}).get("kind") in EXITS:  # the report's terminal value is an exit multiple: source it
        jobs.append(("multiple", lambda: multiple(db, traced, facts, terminal, where,
                                                  set(summary.get("sheets") or []) or None)))
    for key, fn in jobs + [("franking", lambda: franking(sess, summary, db, traced, where, facts, unit))]:
        try:
            out[key] = fn()
        except Exception as ex:  # beside the value, not in its way
            out[key] = {"ends": {}, "report": [], "ok": None, "error": f"{type(ex).__name__}: {ex}"}
    return out
