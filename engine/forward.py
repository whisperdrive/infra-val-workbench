"""Forward traces: from an assumption to the figure, through the formulas that read it.

The tracer (dcftrace.py) goes down from the equity value to its discountings, and finds an assumption only where it
sits in a shape it knows: a terminal value in a row of its own added into the cash flows, franking credits
discounted on their own. Going the other way finds it wherever it is. From the assumption's cell (the report's
figure, on a row labelled like it), every formula that reads it, and every formula that reads those, up to the
figure: the path between them says how the assumption is used.
  franking credit utilisation   the first formula on the way reads it: the franking credits used, the gross
                                credits its other operand. Where that row is one a discounting's cash flows add
                                up, the franking credits are discounted with them, and their value is that part
                                discounted on its own; a cell outside the timeline is the value itself (the
                                utilisation applied to a present value of the gross credits)
  terminal growth rate          the first formula on the way built as X x (1 + g) / (r - g) is the terminal value,
                                and where it sits: a row of its own added into the cash flows, the last cash flow
                                itself, or added after the discounting (=XNPV(...) + TV x factor; the term of the
                                formula after it that reads it is its present value)
  exit multiple                 the same, the terminal value built as M x X
Plain code on the overlay's saved formulas and values, like the tracer: no model reads anything here.
"""
import os
import re
import threading
from collections import defaultdict, deque

import dcf
import dcftrace

LIMIT = 20000   # cells visited at most on the way up from an assumption
SPAN = 2000     # a range over more rows than this (a whole column) isn't a reference to follow
_CACHE: dict = {}
_LOCK = threading.Lock()


def _path_of(db) -> str:
    row = db.execute("PRAGMA database_list").fetchone()
    return row[2] if row else ""


def readers(db) -> dict:
    """{(sheet, row): [(first col, last col, (sheet, row, col) of a formula reading them)]}: who reads each cell,
    long ranges too (a SUMPRODUCT reads every cell of its rows). Cached per model.db."""
    path = _path_of(db)
    key = (path, os.path.getmtime(path) if path and os.path.exists(path) else None)
    with _LOCK:
        if key in _CACHE:
            return _CACHE[key]
    names = dcftrace._names(db)
    idx = defaultdict(list)
    for s, r, c, f in db.execute("SELECT sheet, row, col, formula FROM cells WHERE formula IS NOT NULL"):
        body = dcftrace._expand(db, dcftrace._STR.sub('""', f), names)
        for m in dcf._FREF.finditer(body):
            if body[m.end():m.end() + 1] == "(":
                continue  # a function's name, not a reference
            x = dcf._ref(m[0], s)
            if not x or x[3] - x[1] > SPAN:
                continue
            for rr in range(x[1], x[3] + 1):
                idx[(x[0], rr)].append((x[2], x[4], (s, r, c)))
    with _LOCK:
        _CACHE[key] = idx
    return idx


def path(db, start: tuple, goals: set) -> list[tuple] | None:
    """The shortest way up from a cell to one of the goal cells, through the formulas that read it:
    [start, ..., goal] as (sheet, row, col), or None where nothing that reads it reaches them."""
    idx = readers(db)
    parent, queue = {start: None}, deque([start])
    while queue and len(parent) < LIMIT:
        k = queue.popleft()
        if k in goals:
            out = []
            while k is not None:
                out.append(k)
                k = parent[k]
            return out[::-1]
        for c1, c2, by in idx.get((k[0], k[1]), ()):
            if c1 <= k[2] <= c2 and by not in parent:
                parent[by] = k
                queue.append(by)
    return None


def _a1(k: tuple) -> str:
    return f"{k[0]}!{dcf._addr(k[2], k[1])}"


def _label(db, k: tuple) -> str:
    lab = dcftrace._label(db, k[0], k[1])
    if lab:
        return lab
    left = db.execute("SELECT value FROM cells WHERE sheet=? AND row=? AND col<? AND typeof(value)='text' "
                      "ORDER BY col DESC LIMIT 1", (k[0], k[1], k[2])).fetchone()
    return (left[0] or "").strip() if left else ""


def _num(db, k: tuple):
    return dcf._num(dcf._cell(db, *k))


def holding(db, stated: list[float], words: re.Pattern, sheets: set | None = None, percent: bool = True) -> list[str]:
    """The cells holding one of the report's figures (percent: 2.5 is held as 0.025 or as 2.5; else as printed, a
    multiple of 10.0x as 10) on a row labelled like the assumption, typed inputs first."""
    want = {round(p, 12) for p in stated} | ({round(p / 100, 12) for p in stated} if percent else set())
    out = []
    for s, r, c, v, f in db.execute("SELECT sheet, row, col, value, formula FROM cells "
                                    "WHERE typeof(value) IN ('real', 'integer')"):
        if sheets and s not in sheets:
            continue
        if round(float(v), 12) in want and words.search(_label(db, (s, r, c))):
            out.append((bool(f), _a1((s, r, c))))
    return [a for _, a in sorted(out)]


def _rows(core: dict, db) -> dict:
    """{(sheet, row): (sign, how)}: the discounting's own cash-flow row, and the rows it adds up (any depth)."""
    out = {}
    for rng in core.get("inputs", {}).get("cashflow") or []:
        try:
            s, r, _ = dcf._row_range(db, rng)
            out[(s, r)] = (1, "own")
        except ValueError:
            pass

    def walk(ps, sign):
        for p in ps or []:
            m = re.match(r"^(.+)!r(\d+)$", p.get("row") or "")
            if m and not m[1].startswith("["):
                out.setdefault((m[1], int(m[2])), (sign * p.get("sign", 1), "part"))
            walk(p.get("parts"), sign * p.get("sign", 1))
    walk(core.get("parts"), 1)
    return out


def _part_range(db, core: dict, row: tuple) -> list[str] | None:
    """A row's cells over the discounting's columns, as a range dcf.compute takes."""
    try:
        _, _, cols = dcf._row_range(db, core["inputs"]["cashflow"][0])
    except (ValueError, KeyError, IndexError):
        return None
    return [f"{row[0]}!{dcf._addr(cols[0], row[1])}:{dcf._addr(cols[-1], row[1])}"]


def _in_timeline(db, k: tuple) -> bool:
    n = db.execute("SELECT COUNT(*) FROM cells WHERE sheet=? AND row=? AND typeof(value) IN ('real', 'integer')",
                   (k[0], k[1])).fetchone()[0]
    return n >= 3


def _steps(db, p: list[tuple]) -> list[dict]:
    return [{"cell": _a1(k), "label": _label(db, k), "value": _num(db, k)} for k in p]


def _where(db, row: tuple, cores: list[dict]) -> tuple[dict | None, str, int]:
    """Which discounting a row's figures go into, and how: (core, "own" | "part", sign), else (None, "", 1)."""
    for c in cores:
        hit = _rows(c, db).get(row)
        if hit:
            return c, hit[1], hit[0]
    return None, "", 1


def franking(db, starts: list[str], to: str, cores: list[dict]) -> dict | None:
    """From a utilisation cell up to the figure at to: the franking credits used, the gross credits, where they're
    discounted and their value there. The first start that reaches the figure."""
    goal = dcf._ref(to, "")[:3]
    for u in starts:
        uk = dcf._ref(u, "")[:3]
        p = path(db, uk, {goal})
        if not p or len(p) < 2:
            continue
        used = p[1]
        f = db.execute("SELECT formula FROM cells WHERE sheet=? AND row=? AND col=?", used).fetchone()
        text = dcftrace._STR.sub("", (f[0] if f else "") or "")
        others = [x for x in (dcf._ref(m[0], used[0]) for m in dcf._FREF.finditer(text)
                              if text[m.end():m.end() + 1] != "(") if x and x[:3] != uk and x[1] == x[3] and x[2] == x[4]]
        gross = next((x[:3] for x in others if _num(db, x[:3]) is not None), None)
        out = {"cell": u, "used": _a1(used), "used_row": f"{used[0]}!r{used[1]}", "used_label": _label(db, used),
               "gross": _a1(gross) if gross else None, "gross_row": f"{gross[0]}!r{gross[1]}" if gross else None,
               "gross_label": _label(db, gross) if gross else None, "path": _steps(db, p), "core": None,
               "where": None, "range": None, "sign": 1, "pv": None}
        core, how, sign = _where(db, used[:2], cores)
        if core:
            out.update(core=core["cell"], where=how, sign=sign)
            rng = _part_range(db, core, used[:2])
            if rng:
                try:
                    r = dcf.compute(db, **{**core["inputs"], "cashflow": rng, "compare_to": None}, fix=False)
                    out.update(range=rng, pv=sign * r["pv"])
                except (ValueError, KeyError, ZeroDivisionError):
                    pass
        elif not _in_timeline(db, used):
            out.update(where="after", pv=_num(db, used))  # the utilisation on a present value of the gross credits
        return out
    return None


def _tv_fit(db, k: tuple, fit) -> dict | None:
    """A cell as a terminal value: the whole of it, else one term of its formula (a terminal value added to the
    last cash flow in the same cell), each fitted by fit(db, ref, value, term)."""
    tv = _num(db, k)
    got = fit(db, _a1(k), tv) if tv else None
    if got:
        return got
    f = db.execute("SELECT formula FROM cells WHERE sheet=? AND row=? AND col=?", k).fetchone()
    if not f or not f[0]:
        return None
    import valuation
    body = dcftrace._STR.sub('""', f[0]).lstrip("=").lstrip("+")
    terms = valuation._split(body, "+-")
    if len(terms) < 2:
        return None
    for op, term in terms:
        try:
            v = float(dcftrace.evaluate(db, term, k[0]))
        except (ValueError, TypeError, ZeroDivisionError, OverflowError):
            continue
        got = fit(db, _a1(k), v, term) if v else None
        if got:
            return {**got, "term": term.strip(), "sign": -1 if op == "-" else 1}
    return None


LAST, PART, AFTER = "the last cash flow itself", "a row of its own added into the cash flows", "added after the discounting"


def _pv_term(db, k: tuple, tv: tuple) -> str | None:
    """The term of k's formula that reads the terminal value at tv (its present value: TV x the last factor)."""
    import valuation
    f = db.execute("SELECT formula FROM cells WHERE sheet=? AND row=? AND col=?", k).fetchone()
    body = dcftrace._STR.sub('""', (f[0] if f else "") or "").lstrip("=").lstrip("+")
    for _op, term in valuation._split(body, "+-"):
        refs = [dcf._ref(m[0], k[0]) for m in dcf._FREF.finditer(term) if term[m.end():m.end() + 1] != "("]
        if any(x and x[:3] == tv for x in refs):
            return term.strip()
    return None


def terminal(db, starts: list[str], to: str, cores: list[dict], fit, key: str) -> dict | None:
    """From an assumption's cell (a growth rate, an exit multiple) up to the figure at to: the terminal value, the
    first formula on the way that fit(db, ref, value[, term]) fits with fit[key] (the cell it reads as the
    assumption) on the way, and where it sits. The first start that reaches the figure through one. -> {"cell",
    "tv_cell", "tv_label", "fit", "core", "where", "term" (the terminal value's term of the last cash flow),
    "pv_cell" and "pv_term" (where it's added after the discounting), "path"}."""
    goal = dcf._ref(to, "")[:3]
    for a in starts:
        ak = dcf._ref(a, "")[:3]
        p = path(db, ak, {goal})
        if not p:
            continue
        on = {_a1(k) for k in p}
        for i, k in enumerate(p[1:], 1):
            got = _tv_fit(db, k, fit)
            if not got or (got[key][0] and got[key][0] not in on):
                continue
            core, how, _ = _where(db, k[:2], cores)
            where = LAST if how == "own" else PART if how == "part" else AFTER
            nxt = p[i + 1] if i + 1 < len(p) else None
            return {"cell": a, "tv_cell": _a1(k), "tv_label": _label(db, k), "fit": got,
                    "core": core["cell"] if core else None, "where": where, "term": got.get("term"),
                    "pv_cell": _a1(nxt) if where == AFTER and nxt else None,
                    "pv_term": _pv_term(db, nxt, k) if where == AFTER and nxt else None, "path": _steps(db, p)}
    return None


def growth(db, starts: list[str], to: str, cores: list[dict]) -> dict | None:
    """From a terminal growth cell up to the figure: the terminal value built as X x (1 + g) / (r - g)."""
    import sourced
    return terminal(db, starts, to, cores, sourced._fit, "g")


def multiple(db, starts: list[str], to: str, cores: list[dict]) -> dict | None:
    """From an exit multiple's cell up to the figure: the terminal value built as M x X."""
    import sourced
    return terminal(db, starts, to, cores, sourced._fit_multiple, "m")
