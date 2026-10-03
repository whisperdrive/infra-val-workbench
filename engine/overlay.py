"""The Python overlay: last year's valuation recreated as a live Python module, then re-fed from this year's model.

build() compiles the overlay sheets (xlcompile.py) into out/overlays/e<id>/overlay.py and checks it:
  1. every formula cell recomputed from the workbook's own inputs must equal the value Excel saved
  2. fed from the prior client model instead of the saved link values, the results must not move
  3. the results must tie to the report: conclusions at the printed precision
Then a Session runs it live: change levers (the report's assumptions, found in the overlay) or any cell, and
choose where the client-model values come from:
  workbook   the values saved in the overlay workbook (for an external link: the link's cached values)
  prior      the prior client model file
  current    the current client model, rolled forward: each overlay period moves on by the roll (default: the
             months between the two valuation dates), client values are read from the same line item (sheet +
             label) and the period with the rolled date, and the valuation date lever is set to the new date.
             Anything that can't be matched is listed, never silently zero.
From a terminal, after the engagement's overlay is built in the app:
    uv run python engine/overlay.py 1                                  # engagement 1, as saved
    uv run python engine/overlay.py 1 --mode current --set Val_Inputs!C5=0.075
"""
import json
import math
import re
import sqlite3
import sys
import threading
from collections import Counter, defaultdict, deque
from concurrent.futures import ThreadPoolExecutor
from datetime import date
from pathlib import Path

import rodb
import xlcompile
import xlruntime
from xlruntime import from_db, same, serial, to_date

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "out" / "overlays"
LEVER_KEYS = re.compile(r"discount|wacc|terminal|growth|multiple|rab|valuation_date|net_debt|cpi|tax|cost_of|gearing|"
                        r"inflation|beta|premium", re.I)

sys.setrecursionlimit(1_000_000)
# Windows accepts a thread stack strictly under 256 MB (and commits all of it up front); other systems want a whole
# number of pages. 255 MB suits both, and smaller sizes are tried if a system refuses it.
STACK = 255 * 1024 * 1024
STACK_FALLBACKS = (STACK, 192 * 1024 * 1024, 128 * 1024 * 1024, 64 * 1024 * 1024)
_EXEC = None


def _big_stack() -> int:
    """Set the next thread's stack to the largest size this system accepts; returns the size it had."""
    for size in STACK_FALLBACKS:
        try:
            return threading.stack_size(size)
        except ValueError:  # "size not valid" (Windows: 256 MB or more; some systems: not a page multiple)
            continue
    return threading.stack_size()


def deep(fn, *args, **kw):
    """Run fn on a thread with a 255 MB stack: a timeline recurrence can nest thousands of cells deep. One thread,
    so runs don't interleave: a Session's settings belong to the run that set them."""
    global _EXEC
    if threading.current_thread().name.startswith("overlay-eval"):
        return fn(*args, **kw)
    if _EXEC is None:
        old = _big_stack()
        try:
            _EXEC = ThreadPoolExecutor(1, thread_name_prefix="overlay-eval")
            _EXEC.submit(lambda: None).result()  # create the thread while the big stack size is set
        finally:
            threading.stack_size(old)
    return _EXEC.submit(fn, *args, **kw).result()


def _ro(path) -> sqlite3.Connection:
    return rodb.connect(path, check_same_thread=False)


def _a1(sheet, row, col) -> str:
    from openpyxl.utils import get_column_letter
    return f"{sheet}!{get_column_letter(col)}{row}"


def parse_a1(ref: str) -> tuple[str, int, int]:
    from openpyxl.utils import column_index_from_string
    sheet, addr = ref.rsplit("!", 1)
    m = re.match(r"^\$?([A-Z]{1,3})\$?(\d+)$", addr.upper())
    return sheet.strip("'"), int(m.group(2)), column_index_from_string(m.group(1))


def add_months(s: float, months: int) -> float:
    """Serial date moved by whole months; a month-end stays a month-end."""
    d = to_date(s)
    end = xlruntime._add_months(d, 0, end=True) == d
    return serial(xlruntime._add_months(d, months, end=end))


def months_between(a: str, b: str) -> int:
    da, db_ = date.fromisoformat(a[:10]), date.fromisoformat(b[:10])
    return (db_.year - da.year) * 12 + (db_.month - da.month)


# ---- workbook data ------------------------------------------------------------------------------------------

class Workbook:
    """Cell values of one model.db, loaded a sheet at a time, plus line-item labels and the timeline per sheet."""

    def __init__(self, db_path: str):
        self.path = db_path
        self.db = _ro(db_path)
        self.sheets = {}
        self._labels = None
        self._timeline = {}
        self._inverse = {}

    def close(self):
        self.db.close()

    def sheet(self, s):
        if s not in self.sheets:
            self.sheets[s] = {(r, c): from_db(v) for r, c, v in
                              self.db.execute("SELECT row, col, value FROM cells WHERE sheet=?", (s,))}
        return self.sheets[s]

    def value(self, s, r, c):
        return self.sheet(s).get((r, c))

    def labels(self):
        if self._labels is None:
            self._labels = {(s, r): (lab or "") for s, r, lab in self.db.execute("SELECT sheet, row, label FROM rows")}
        return self._labels

    def timeline(self, s) -> dict[int, float]:
        """col -> period date (serial) from the sheet's timeline row."""
        if s not in self._timeline:
            lay = self.db.execute("SELECT layout FROM sheets WHERE sheet=?", (s,)).fetchone()
            hr = json.loads(lay[0] or "{}").get("header_row") if lay else None
            tl = {}
            if hr:
                for (r, c), v in self.sheet(s).items():
                    if r == hr and isinstance(v, float) and 3000 < v < 120000:
                        tl[c] = v
            self._timeline[s] = tl
        return self._timeline[s]

    def column_of(self, s, when: float):
        """Column of a period date on the sheet's timeline, or None."""
        if s not in self._inverse:
            self._inverse[s] = {v: k for k, v in self.timeline(s).items()}
        return self._inverse[s].get(when)

    def header_row(self, s):
        lay = self.db.execute("SELECT layout FROM sheets WHERE sheet=?", (s,)).fetchone()
        return json.loads(lay[0] or "{}").get("header_row") if lay else None


class RowMap:
    """Prior (sheet, row) -> current (sheet, row), by line-item label:
      1. the n-th row with that label on the sheet (rows inserted above it don't matter)
      2. else the row with that label nearest its old position (the label now occurs more or fewer times)
      3. else, on a sheet laid out as before (most labelled rows still at the same row), the same row, which also
         covers rows without a label (flags, timing rows)
    why() says what failed for a row it can't map."""

    def __init__(self, prior: Workbook, current: Workbook):
        norm = lambda lab: re.sub(r"\s+", " ", (lab or "").strip().lower())

        def occ(wb):
            seen, out = defaultdict(int), {}
            for (s, r), lab in sorted(wb.labels().items()):
                k = (s, norm(lab))
                seen[k] += 1
                out[(s, r)] = (*k, seen[k])
            return out
        self.prior, self.current = occ(prior), occ(current)
        self.back = {v: k for k, v in self.current.items()}
        self.by_label = defaultdict(list)
        for (s, r), (_, lab, _) in self.current.items():
            if lab:
                self.by_label[(s, lab)].append(r)
        self.cur_sheets = {s for (s,) in current.db.execute("SELECT sheet FROM sheets")}
        cur_labels = {k: norm(v) for k, v in current.labels().items()}
        same, total = Counter(), Counter()
        for (s, r), (_, lab, _) in self.prior.items():
            if lab:
                total[s] += 1
                same[s] += cur_labels.get((s, r)) == lab
        self.same_layout = {s for s in total if s in self.cur_sheets and same[s] >= 0.8 * total[s]}

    def row(self, s, r):
        return self.match(s, r)[0]

    def match(self, s, r) -> tuple[int | None, str]:
        """(current row or None, how it was matched)."""
        k = self.prior.get((s, r))
        if k and k[1]:
            hit = self.back.get(k)
            if hit and hit[0] == s:
                return hit[1], "same label" if k[2] == 1 else f"same label, occurrence {k[2]} on the sheet"
            rows = self.by_label.get((s, k[1]))
            if rows:
                return min(rows, key=lambda x: abs(x - r)), "same label, the nearest of its rows (it occurs a different number of times now)"
        if s in self.same_layout:
            return r, "same row (no label to follow; the sheet's layout is unchanged)"
        return None, "unmatched"

    def why(self, s, r, labels: dict) -> str:
        lab = labels.get((s, r), "")
        if s not in self.cur_sheets:
            return f"sheet {s} isn't in the current model"
        if not lab:
            return f"row {r} of {s} has no line-item label, and the sheet's layout changed, so it can't be followed"
        return f"line item '{lab}' isn't on {s} in the current model"


# ---- a live session -----------------------------------------------------------------------------------------

HEAD_ROWS = 15  # rows above a single figure looked at for its column's heading


def _heading(wb: Workbook, s: str, r: int, c: int) -> tuple[str, int] | None:
    """The text heading a figure's column above it (the nearest text cell up the column, within HEAD_ROWS): a
    case's name over a column of inputs. (text, its row), or None."""
    vals = wb.sheet(s)
    for rr in range(r - 1, max(0, r - HEAD_ROWS), -1):
        v = vals.get((rr, c))
        if isinstance(v, str) and v.strip() and not re.match(r"^\d{4}-\d{2}-\d{2}", v):
            return v.strip(), rr
    return None


def case_column(prior: Workbook, cur: Workbook, s: str, r: int, c: int, s2: str, r2: int) -> tuple[int | None, str]:
    """This year's column for a single figure last year read in column c: c, unless the column's heading changed (a
    case column inserted before it): then the column in this year's heading row with last year's heading, or None
    where there's none. -> (column or None, why)."""
    was = _heading(prior, s, r, c)
    if not was:
        return c, ""
    # a heading compared without its years and period markers ("FY25 inputs" is "FY26 inputs"), word for word
    norm = lambda x: re.sub(r"[^a-z]+", " ", re.sub(r"\b(?:fy|cy|h[12]|q[1-4])?\s*'?\d{2,4}\b", " ", x.lower())).strip()
    now = _heading(cur, s2, r2, c)
    if now and norm(now[0]) == norm(was[0]):
        return c, ""
    row = now[1] if now else None
    vals = cur.sheet(s2)
    rows = [row] if row else range(r2 - 1, max(0, r2 - HEAD_ROWS), -1)
    for rr in rows:
        hits = [cc for (r_, cc), v in vals.items() if r_ == rr and isinstance(v, str) and norm(v) == norm(was[0])]
        if len(hits) == 1:
            return hits[0], (f"the column headed '{was[0]}' (column {c} last year, {hits[0]} this year: "
                             f"'{now[0] if now else ''}' is in column {c} now)")
    return None, f"the column's heading was '{was[0]}', it's '{now[0] if now else 'nothing'}' this year, and no column " \
                 f"is headed '{was[0]}'"


class Session:
    """A compiled overlay wired to its inputs and a feed. All evaluation goes through deep()."""

    def __init__(self, module_path: str, overlay_db: str, sheets: list[str], prior_db: str | None = None,
                 current_db: str | None = None, client_link: int | None = None, client_sheets: list[str] | None = None):
        src = Path(module_path).read_text(encoding="utf-8")
        self.g = {"__name__": "overlay"}
        exec(compile(src, str(module_path), "exec"), self.g)
        self.B = self.g["B"]
        self.sheets = sheets
        self.ov = Workbook(overlay_db)
        self.prior = Workbook(prior_db) if prior_db and prior_db != overlay_db else None
        self.current = Workbook(current_db) if current_db else None
        self.client_link = client_link
        self.client_sheets = set(client_sheets or [])  # client sheets inside the overlay workbook (combined case)
        self.ext_cached = {}
        with _ro(overlay_db) as db:
            have = {r[0] for r in db.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            if "extcells" in have:
                self.ext_cached = {(i, s, r, c): from_db(v) for i, s, r, c, v in
                                   db.execute("SELECT idx, sheet, row, col, value FROM extcells")}
        B = self.B
        for s in sheets:
            with _ro(overlay_db) as db:
                for r, c, v in db.execute("SELECT row, col, value FROM cells WHERE sheet=? AND formula IS NULL "
                                          "AND value IS NOT NULL", (s,)):
                    B.inputs[(s, r, c)] = from_db(v)
        self.formula_cells = []
        with _ro(overlay_db) as db:
            for s, r, c in db.execute(f"SELECT sheet, row, col FROM cells WHERE formula IS NOT NULL AND sheet IN "
                                      f"({','.join('?' * len(sheets))})", sheets):
                self.formula_cells.append((s, r, c))
        self.formula_cells.sort(key=lambda k: (k[2], k[1]))  # column by column: early periods are memoised first
        self.formula_set = set(self.formula_cells)
        B.cached = lambda s, r, c: self.ov.value(s, r, c)
        self.rowmap = None
        if self.current:  # last year's rows in this year's model: by label, history, words, neighbours, banner, trace
            import rowfind
            base = self.prior or self.ov
            self.rowmap = rowfind.RowFinder(RowMap(base, self.current), base, self.current)
        self.stood_in = {}  # (sheet, row, col) -> last year's value, used where this year's model has no match
        self.blank = {}  # (sheet, row, col) -> this year's cell found for it, blank where last year's had a value
        self.beyond = {}  # (sheet, row, col) -> the rolled period past the last this year's model has: nil there
        self.derived = {}  # (sheet, row) -> timing rule: a period flag or date row worked out from the period dates
        self.derived_used = {}  # (sheet, row, col) -> the rule, where a value was worked out on the current feed
        self.client_reads = set()  # (sheet, row, col) of last year's model read on the current feed
        self.base_vd = None  # last year's valuation date (serial): the roll's start
        self._pshift = {}
        self.horizon_set = None  # "fixed" / "rolling": the engagement's profile says, not worked out (fixed_horizon)
        self.mode = None
        self.holds = {}  # (sheet, row, col) -> value: cells held at Excel's value on every feed (the doctor's fixes)
        self.balances = {}  # client (sheet, row) -> {"cols", "series"}: read as a balance at the valuation date
        self.balances_at_last = False  # read them at last year's date instead: what their move adds (result.figures)
        self.moved, self.unmoved = {}, {}  # those read at this year's date on the current feed, and those not
        self.cutoffs = []  # (sheet, row, col, period end serial[, its date cell]): the discountings' per-period cells (cutoff_cells)
        self.cut = set()  # those of them cut off on the current feed (_cut_off)
        self.keep_on_date = False  # the period ending on the new valuation date stays in (a method: methods.py)
        self.configure("workbook")

    # feeds
    def configure(self, mode: str, overrides: dict | None = None, shift_months: int = 0, roll_dates: bool = True):
        """mode: workbook | prior | current. overrides: {(sheet, row, col): value}. shift_months: roll-forward."""
        B = self.B
        self.mode, self.shift = mode, shift_months
        self.unmatched = {}
        self.stood_in = {}
        self.blank = {}
        self.beyond = {}
        self.derived_used = {}
        self.client_reads = set()
        self.moved, self.unmoved = {}, {}
        self.other_reads = {}  # (link, sheet, row, col) -> value: another linked workbook's, last year's saved value
        B.overrides.clear()
        B.overrides.update(self.holds)
        B.reset()
        if mode == "workbook":
            B.feed = lambda s, r, c: self.ov.value(s, r, c)
            B.ext = lambda i, s, r, c: self.ext_cached.get((i, s, r, c))
        elif mode == "prior":
            src = self.prior or self.ov
            B.feed = lambda s, r, c: src.value(s, r, c) if s in self.client_sheets else self.ov.value(s, r, c)
            B.ext = lambda i, s, r, c: (self.prior.value(s, r, c) if self.prior and i == self.client_link
                                        else self.ext_cached.get((i, s, r, c)))
        elif mode == "current":
            if not self.current:
                raise ValueError("no current client model to feed from")
            prior = self.prior or self.ov
            B.feed = lambda s, r, c: self._rolled(prior, s, r, c) if s in self.client_sheets else self.ov.value(s, r, c)
            def ext(i, s, r, c):
                if i == self.client_link:
                    return self._rolled(prior, s, r, c)
                v = self.ext_cached.get((i, s, r, c))  # another workbook than the client model: last year's saved value
                self.other_reads[(i, s, r, c)] = v
                return v
            B.ext = ext
            if roll_dates and shift_months:
                for key, v in self.rolled_timeline(shift_months).items():
                    B.overrides[key] = v
        else:
            raise ValueError(f"unknown feed {mode}")
        for k, v in (overrides or {}).items():
            B.overrides[k] = v
        self.cut = self._cut_off(shift_months) if mode == "current" else set()
        B.range_cache.clear()

    def _cut_off(self, months: int) -> set:
        """On this year's feed, a discounting's periods that end on or before the new valuation date count as nil.
        Last year's overlay discounts what its client model forecast after last year's date, and needed no cut-off of
        its own where that model had nothing before it; rolled forward, the periods up to the new date are past (the
        bridge's cash flows paid), and a factor worked out from the date would compound them into the value instead.
        The period ending on the date is past too (the convention dcf.factors keeps, and a valuer zeroing the
        overlay's own period flags by hand does); keep_on_date leaves it in, undiscounted (methods.py). A cell a
        person or the feed set is left as set. A period's end is its date cell's on this feed where an overlay
        formula works it out (dates counted from the valuation date move with it: the column is then this year's
        period, not last year's), else last year's end moved by the sheet's periods. -> the cells cut off."""
        if not self.cutoffs or self.base_vd is None:
            return set()
        new_vd = add_months(self.base_vd, months)
        shift, cut, worked = {}, set(), False
        for s, r, c, end, *at in self.cutoffs:
            at = at[0] if at else None
            if at and at in self.formula_set:
                e = deep(lambda: self.value(*at))
                worked = True
                if isinstance(e, (int, float)) and not isinstance(e, bool) and 3000 < e < 120000:
                    if (e < new_vd or e == new_vd and not self.keep_on_date) and (s, r, c) not in self.B.overrides:
                        self.B.overrides[(s, r, c)] = 0.0
                        cut.add((s, r, c))
                    continue
            if s not in shift:
                shift[s] = self.period_shift(self.ov if s in self.sheets else (self.prior or self.ov), s)
            e = add_months(end, shift[s]) if shift[s] else end
            if (e < new_vd or e == new_vd and not self.keep_on_date) and (s, r, c) not in self.B.overrides:
                self.B.overrides[(s, r, c)] = 0.0
                cut.add((s, r, c))
        if worked:  # the dates were worked out before the cut: nothing worked out from them stays
            self.B.reset()
        return cut

    def _rolled(self, prior: Workbook, s, r, c):
        """Prior client cell -> the current model's value: the same line item (found by rowfind, wherever it is
        now), the period rolled on by the shift. Where this year's model has no match, last year's value stands
        in (its forecast for the same period), recorded in unmatched and stood_in: never a blank, which would
        read as zero. That includes a row and period that are found but blank this year where last year's cell
        had a value (a row matched on its history that stops where the history does): recorded in blank too."""
        cur = self.current
        self.client_reads.add((s, r, c))
        tl_p = prior.timeline(s)
        want = add_months(tl_p[c], self.period_shift(prior, s)) if c in tl_p else None
        rule = self.derived.get((s, r))
        if rule and rule["kind"].startswith("own"):  # a row with periods of its own: moved by them, not hunted for
            self.derived_used[(s, r, c)] = rule["text"]
            return self._own_value(rule, prior.value(s, r, c))
        if rule and want is not None:  # timing worked out from the rolled period date, not hunted for
            self.derived_used[(s, r, c)] = rule["text"]
            new_vd = add_months(self.base_vd, self.shift) if self.base_vd is not None else None
            return timing_value(rule, want, new_vd)
        hit = self.rowmap.locate(s, r)
        if hit is None:
            self.unmatched[(s, r, c)] = self.rowmap.why(s, r, prior.labels())
            return self._stand_in(prior, s, r, c, want)
        s2, r2 = hit
        c2 = c
        bal = self.balances.get((s, r))
        if bal and want is not None and self.base_vd is not None and c in tl_p:
            # a balance at last year's valuation date (balance_cells): this year's at this year's date. On a fixed
            # horizon the periods keep their dates, so nothing else moves a plain reference off last year's; an overlay
            # reading it by its date (INDEX/MATCH on the valuation date) moves it itself. Only the balance's own read:
            # where the row is also read as cash flows, those reads are the periods' (the same cell's value serves both
            # only for last year's date's column, whose period is cut off)
            old, new_vd = self.base_vd, add_months(self.base_vd, self.shift)
            single = c in bal["cols"]
            own = not single and not bal["series"] and round(new_vd) != round(old) and round(tl_p[c]) == round(new_vd)
            target, by = want, None
            if single or own:
                if self.balances_at_last:
                    target = old
                elif single and round(want) != round(new_vd):
                    target, by = new_vd, "the app"
                elif single and round(new_vd) != round(old):
                    by = "the periods"
                elif own:
                    by = "the overlay's own date"  # the overlay picked another column itself (INDEX/MATCH on its date)
            c3 = cur.column_of(s2, target) if single or own else None
            if not (single or own):
                pass
            elif c3 is None:
                # no column at this year's date: left at last year's date (a plain reference), or last year's forecast
                # standing in (read by its date): either way not this year's balance, which holds the value
                if not self.balances_at_last and round(target) == round(new_vd) and round(new_vd) != round(old):
                    self.unmoved[(s, r, c)] = f"this year's model has no column at {to_date(new_vd).isoformat()}"
            else:
                if by:
                    self.moved[(s, r, c)] = (old, target, s2, r2, c3, by)
                want = target
        if want is None and not (len(tl_p) > 1 and c > max(tl_p)):
            # a single figure, not on the timeline: from the column whose heading is last year's (a case column
            # inserted before it moves it), not the same column letter
            c2, why = case_column(prior, cur, s, r, c, s2, r2)
            if c2 is None:
                self.unmatched[(s, r, c)] = why
                return self._stand_in(prior, s, r, c, want)
            if c2 != c:
                self.derived_used[(s, r, c)] = why
        if want is None and len(tl_p) > 1 and c > max(tl_p):
            # to the right of last year's timeline: no period last year to roll, so not this year's same column (another
            # period there, counted twice); as last year had it, nothing
            self.beyond[(s, r, c)] = "past last year's timeline"
            return prior.value(s, r, c)
        if want is not None:
            c2 = cur.column_of(s2, want)
            tl_c = cur.timeline(s2)
            if c2 is None and tl_c and want > max(tl_c.values()) and rule is None:
                # past the last period this year's model has: its forecast has ended there, nothing to stand in for
                self.beyond[(s, r, c)] = to_date(want).isoformat()
                return 0.0 if isinstance(prior.value(s, r, c), (int, float)) else None
            if c2 is None:
                self.unmatched[(s, r, c)] = f"period {to_date(want).isoformat()} not in the current model"
                return self._stand_in(prior, s, r, c, want)
        v = cur.value(s2, r2, c2)
        if v is None and prior.value(s, r, c) is not None:
            self.blank[(s, r, c)] = (s2, r2, c2)
            self.unmatched[(s, r, c)] = (f"found at {s2}!r{r2}, but blank there"
                                         + (f" for {to_date(want).isoformat()}" if want is not None else ""))
            return self._stand_in(prior, s, r, c, want)
        return v

    def _own_value(self, rule: dict, v):
        """A cell of a row with periods of its own (own_rule), this year: moved on by its periods that ended between
        last year's valuation date and the new one (on a fixed horizon, none: the dates are the same dates)."""
        if not isinstance(v, (int, float)) or isinstance(v, bool) or self.base_vd is None:
            return v
        n = 0
        if not self.fixed_horizon():
            new_vd = add_months(self.base_vd, self.shift)
            n = sum(1 for e in rule.get("ends") or [] if self.base_vd < e <= new_vd)
        if not n:
            return v
        return add_months(v, n * rule["plen"]) if rule["kind"] == "own_date" else v + n * rule["plen"] / 12

    def _stand_in(self, prior: Workbook, s, r, c, want):
        """Last year's value for a client cell this year's model doesn't have: last year's forecast for the same
        period; where last year's model doesn't reach that period, its own cell (the period it read last year);
        the same cell where the row has no timeline. Blank only where last year's cell was blank too."""
        c0 = prior.column_of(s, want) if want is not None else None
        v = prior.value(s, r, c0 if c0 is not None else c)
        if v is not None:
            self.stood_in[(s, r, c)] = v
        return v

    def fixed_horizon(self) -> bool:
        """Whether rolling forward keeps the period dates: set in the engagement's profile (horizon_set), else
        worked out (horizon()). Cached until the roll is planned again."""
        if ("fixed",) not in self._pshift:
            set_ = getattr(self, "horizon_set", None)
            if set_ in ("fixed", "rolling"):
                self._pshift[("fixed",)] = set_ == "fixed"
            elif not self.current or not self.rowmap:
                self._pshift[("fixed",)] = False
            else:
                sheets = self.client_sheets or {k[1] for k in self.ext_cached}
                self._pshift[("fixed",)] = horizon(self.prior or self.ov, self.current, sheets, self.rowmap.sheet_for)[0] == "fixed"
        return self._pshift[("fixed",)]

    def period_shift(self, wb: Workbook, s: str) -> int:
        """How far a sheet's periods move when rolling forward, in months: the whole periods that ended between
        last year's valuation date and the new one. The valuation date can move three months while an annual
        sheet's periods stay put (FY2026 is still FY2026) and a quarterly sheet's move one quarter; moving an
        annual sheet by three months would land on dates it doesn't have. Nothing on a fixed horizon (this year's
        model ends where last year's did): the dates are the same dates, only the valuation date moves."""
        if not self.shift:
            return 0
        key = (wb.path, s, self.shift)
        if key not in self._pshift:
            tl = sorted(set(wb.timeline(s).values()))
            if self.fixed_horizon():
                out = 0
            elif len(tl) < 2:
                out = 0
            else:
                gaps = sorted(b - a for a, b in zip(tl, tl[1:]))
                plen = max(1, round(gaps[len(gaps) // 2] / 30.44))
                # a timeline of period starts (1 July, 1 October: all on the 1st) is counted by its periods' ends,
                # the day before the next start: a year that began in the window hasn't ended in it
                if all(to_date(d).day == 1 for d in tl):
                    ends = [add_months(d, plen) - 1 for d in tl]
                else:
                    ends = tl
                if self.base_vd is not None:
                    new = add_months(self.base_vd, self.shift)
                    out = plen * sum(1 for d in ends if self.base_vd < d <= new)
                else:
                    out = plen * (max(0, self.shift) // plen)
            self._pshift[key] = out
        return self._pshift[key]

    def rolled_timeline(self, months: int) -> dict:
        """The overlay's own period dates moved on: constants in each overlay sheet's timeline row, by whole
        periods (period_shift), not by the valuation date's move."""
        out = {}
        keep = self.shift
        self.shift = months
        try:
            for s in self.sheets:
                hr = self.ov.header_row(s)
                step = self.period_shift(self.ov, s)
                if not hr or not step:
                    continue
                for (r, c), v in self.ov.sheet(s).items():
                    if r == hr and (s, r, c) in self.B.inputs and isinstance(v, float) and 3000 < v < 120000:
                        out[(s, r, c)] = add_months(v, step)
        finally:
            self.shift = keep
        return out

    # evaluation
    def close(self):
        """Let go of the model.db files (a workbook being rebuilt deletes its model.db)."""
        for wb in (self.ov, self.prior, self.current):
            if wb:
                wb.close()

    def paths(self) -> set[str]:
        return {wb.path for wb in (self.ov, self.prior, self.current) if wb}

    def value(self, s, r, c):
        """A cell as the module sees it: overlay sheets computed, other sheets from the feed."""
        return self.B.get("", s, r, c)

    def values(self, cells):
        return deep(lambda: [self.value(*k) for k in cells])

    def evaluate_all(self):
        return deep(lambda: {k: self.value(*k) for k in self.formula_cells})

    def validate(self, limit: int = 60) -> dict:
        """Every formula cell vs the value Excel saved (feed = workbook)."""
        self.configure("workbook")
        xlruntime.xl.quirks.clear()
        res = self.evaluate_all()
        bad = []
        for k, v in res.items():
            w = self.ov.value(*k)
            if not same(v, w):
                bad.append((k, v, w))
        with _ro(self.ov.path) as db:
            forms = {}
            for k, _, _ in bad[:limit]:
                forms[k] = db.execute("SELECT formula FROM cells WHERE sheet=? AND row=? AND col=?", k).fetchone()[0]
        labels = self.ov.labels()
        return {"cells": len(res), "matched": len(res) - len(bad), "cycles": len(self.B.cycles), "runtime": xlruntime.RUNTIME,
                "text": sum(isinstance(w, str) for _, _, w in bad),  # label cells: Excel saved text (a caption)
                "unsupported": dict(self.B.unsupported), "quirks": dict(xlruntime.xl.quirks),
                "mismatches": [{"cell": _a1(*k), "label": labels.get(k[:2], ""), "python": _show(v), "workbook": _show(w),
                                "formula": forms.get(k)} for k, v, w in bad[:limit]]}

    def outputs(self, cells: list[tuple]) -> list:
        return self.values(cells)


def _show(v):
    if isinstance(v, xlruntime.XLError):
        return v.code
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return str(v)
    return v


# ---- building and checking ----------------------------------------------------------------------------------

def _num_in(text: str):
    m = re.search(r"-?\d[\d,]*(?:\.\d+)?", text or "")
    return (float(m.group(0).replace(",", "")), len(m.group(0).split(".")[1]) if "." in m.group(0) else 0) if m else None


def levers_and_outputs(overlay_db: str, sheets: list[str], facts: list[dict]) -> tuple[list, list]:
    """Levers: the report's assumptions found in the overlay (label agrees). Outputs: the overlay's DCF anchors and
    the cells holding the report's conclusions."""
    import linkmap
    import valuation
    fm = linkmap.match_facts(overlay_db, facts, set(sheets))
    by_id = {f["id"]: f for f in facts}
    with _ro(overlay_db) as db:
        formula = {(s, r, c) for s, r, c in db.execute(
            f"SELECT sheet, row, col FROM cells WHERE formula IS NOT NULL AND sheet IN ({','.join('?' * len(sheets))})", sheets)}
    from openpyxl.utils import column_index_from_string as ci
    levers, outputs, seen = [], [], set()
    for m in fm:
        f = by_id[m["fact_id"]]
        located = [x for x in m["matches"] if x["label_match"] or x["anchor"]]
        # the cell holding the report's own figure first: a range's low or high end isn't the rate the report quotes
        best = next((x for x in located if x.get("part") == "value"), None) or next(iter(located), None)
        if not best:
            continue
        cell = (best["sheet"], best["row"], ci(re.sub(r"\d", "", best["addr"])))
        if f["category"] in ("assumption", "identity") and LEVER_KEYS.search(f["key"] or "") and cell not in seen:
            seen.add(cell)
            levers.append({"key": f["key"], "label": f.get("label") or f["key"], "cell": _a1(*cell), "report": f.get("value_text"),
                           "unit": f.get("unit"), "scale": best.get("scale", 1.0), "formula": cell in formula,
                           "row_label": best["label"]})
        if f["category"] == "conclusion":
            outputs.append({"cell": _a1(*cell), "label": best["label"] or f.get("label"), "fact_id": f["id"],
                            "key": f["key"], "report": f.get("value_text"), "low": f.get("low_text"), "high": f.get("high_text"),
                            "unit": f.get("unit"), "scale": best.get("scale", 1.0), "sign": best.get("sign", 1)})
    try:
        for a in valuation.catalogue(overlay_db):
            if a.get("ok") and a["cell"].split("!")[0] in sheets and a["cell"] not in {o["cell"] for o in outputs}:
                outputs.append({"cell": a["cell"], "label": a["label"], "fact_id": None, "key": None, "report": None,
                                "unit": None, "scale": 1.0, "sign": 1})
    except Exception:
        pass
    return levers, outputs


def tie(value, report_text, scale=1.0, sign=1) -> dict | None:
    """Does a Python value round to the report's printed figure?"""
    n = _num_in(report_text)
    if n is None or not isinstance(value, float):
        return None
    x, d = n
    shown = sign * value / (scale or 1.0)
    return {"report": report_text, "python": round(shown, d + 2), "ok": abs(round(shown, d) - x) <= 0.5 * 10 ** -d + 1e-9}


def build(out_dir: Path, overlay: dict, prior: dict | None, current: dict | None, facts: list[dict],
          title: str, client_link: int | None, prior_val_date: str | None, progress=None,
          client_sheets: list[str] | None = None) -> dict:
    """overlay / prior / current: {"db_path", "filename", "sheets"}. Returns the summary saved to overlay.json.
    client_sheets: the overlay workbook's own copy of the client model's sheets, when the overlay sits in a copy of
    a client model that is also here as its own file (prior): those sheets are then fed from that file."""
    progress = progress or (lambda f, m: None)
    out_dir.mkdir(parents=True, exist_ok=True)
    sheets = overlay["sheets"]
    progress(0.05, "Compiling the overlay's formulas to Python")
    src, stats = deep(xlcompile.compile_overlay, overlay["db_path"], sheets, title,
                      lambda f, m: progress(0.05 + 0.3 * f, m), overlay.get("source_path"))
    module = out_dir / "overlay.py"
    module.write_text(src, encoding="utf-8")
    progress(0.4, "Loading the module")
    same_file = prior is not None and prior["db_path"] == overlay["db_path"]
    sess = Session(str(module), overlay["db_path"], sheets, None if same_file else (prior or {}).get("db_path"),
                   (current or {}).get("db_path"), client_link,
                   client_sheets or ((prior or {}).get("sheets") if same_file else None))
    relinked = None
    if prior and not same_file and link_unsaved(overlay["db_path"], client_link):
        progress(0.45, "The overlay's link to the client model was saved without values: filling it from last year's "
                       "client model and working its formulas out again")
        path2, relinked = deep(relink, sess, out_dir, client_link)
        sess.close()
        overlay = {**overlay, "db_path": path2}
        sess = Session(str(module), path2, sheets, prior["db_path"], (current or {}).get("db_path"), client_link,
                       client_sheets)
    progress(0.5, f"Recomputing {len(sess.formula_cells):,} formula cells and checking each against Excel")
    val = sess.validate()
    progress(0.7, "Finding levers and outputs")
    levers, outputs = levers_and_outputs(overlay["db_path"], sheets, facts)
    out_cells = [parse_a1(o["cell"]) for o in outputs]
    lever_cells = [parse_a1(l["cell"]) for l in levers]
    base = sess.values(out_cells + lever_cells)
    for o, v in zip(outputs, base[:len(outputs)]):
        o["value"] = _show(v)
        o["tie"] = tie(v, o["report"], o["scale"], o["sign"]) if o.get("report") else None
    for l, v in zip(levers, base[len(outputs):]):
        l["value"] = _show(v)
    feeds = {}
    if prior and not same_file:
        progress(0.78, "Feeding it from the prior client model")
        sess.configure("prior")
        pv = sess.values(out_cells)
        feeds["prior"] = {"outputs": [_show(v) for v in pv], "same": all(same(a, b) for a, b in zip(pv, base[:len(outputs)]))}
    sess.configure("workbook")
    roll = None
    if current:
        vd_lever = next((l for l in levers if l["key"] == "valuation_date"), None)
        roll = plan_roll(sess, prior, overlay, same_file, prior_val_date, (prior or {}).get("valuation_date"),
                         current.get("valuation_date"))
        # the dates the discountings under the figures read are the ones to move, every one of them: a cell labelled
        # like the valuation date elsewhere (a client inputs sheet holding the same date) would leave them
        # discounting to last year's, and so would moving one discounting's date and not another's
        roll.update(date_cells(overlay["db_path"], outputs, sheets, vd_lever))
    summary = {"module": str(module), "stats": {k: v for k, v in stats.items() if k != "not_compiled"},
               "not_compiled": stats["not_compiled"][:50], "validation": val, "levers": levers, "outputs": outputs,
               "feeds": feeds, "roll": roll, "sheets": sheets, "relinked": relinked,
               "files": {k: (v or {}).get("filename") for k, v in (("overlay", overlay), ("prior", prior), ("current", current))},
               "client_link": client_link, "same_file": same_file}
    (out_dir / "overlay.json").write_text(json.dumps(summary, default=str, indent=1), encoding="utf-8")
    progress(1.0, "Done")
    return summary, sess


def _cores(db, outputs: list[dict]) -> list[dict]:
    """The discountings under the outputs (dcftrace), each once, in the order they're met."""
    import dcftrace
    out, seen = [], set()
    for o in outputs[:8]:
        try:
            got = dcftrace.cores(dcftrace.trace(db, o["cell"]))
        except ValueError:
            continue
        for c in got:
            if (c["cell"], c["call"]) not in seen:
                seen.add((c["cell"], c["call"]))
                out.append(c)
    return out


def link_unsaved(path: str, idx: int | None) -> bool:
    """Whether the overlay's link to the client model was saved without the values it read (Excel's option to save
    external link values off, or links broken when it was saved), while its formulas read through it: then Excel's
    saved values of everything those formulas feed are stale, and only a recompute gives them."""
    if idx is None:
        return False
    with _ro(path) as db:
        have = {r[0] for r in db.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        if not {"extcells", "extrefs"} <= have:
            return False
        cached = db.execute("SELECT COUNT(*) FROM extcells WHERE idx=?", (idx,)).fetchone()[0]
        reads = db.execute("SELECT COALESCE(SUM(n_cells), 0) FROM extrefs WHERE idx=?", (idx,)).fetchone()[0]
    return not cached and reads > 0


def relink(sess: Session, out_dir: Path, idx: int) -> tuple[str, dict]:
    """A copy of the overlay's model.db with its link to the client model filled from last year's client model (the
    values every formula reads through it, as that file has them) and its formulas' saved values worked out again
    from them, where Excel saved the link without values (link_unsaved). Everything that reads the overlay's saved
    values (finding the report's figures in it, tracing its discountings, sourcing its inputs) then reads what the
    overlay gives on last year's model, not Excel's stale figures. -> (the copy's path, what was filled)."""
    import shutil
    from openpyxl.utils import get_column_letter
    dst = out_dir / "overlay_relinked.db"
    shutil.copyfile(sess.ov.path, dst)
    sess.configure("prior")
    sess.B.feed_log = {}
    try:
        vals = sess.evaluate_all()
        log = dict(sess.B.feed_log)
    finally:
        sess.B.feed_log = None
        sess.configure("workbook")
    stored = lambda v, saved: v.code if isinstance(v, xlruntime.XLError) else (
        to_date(v).isoformat() if isinstance(v, float) and isinstance(saved, str) and xlruntime._DATE.match(saved)
        and 0 < v < 2958466 else v)
    ext = [(i, s, f"{get_column_letter(c)}{r}", r, c, stored(v, None)) for (i, s, r, c), v in log.items()
           if i == idx and v is not None and not isinstance(v, xlruntime.Rng)]
    with sqlite3.connect(dst) as db:
        saved = {(s, r, c): v for s, r, c, v in db.execute("SELECT sheet, row, col, value FROM cells WHERE formula IS NOT NULL")}
        changed = [(stored(v, saved.get(k)), *k) for k, v in vals.items() if not isinstance(v, xlruntime.Rng)
                   and not same(v, from_db(saved.get(k)))]
        db.execute("DELETE FROM extcells WHERE idx=?", (idx,))
        db.executemany("INSERT INTO extcells(idx, sheet, addr, row, col, value) VALUES (?,?,?,?,?,?)", ext)
        db.executemany("UPDATE cells SET value=? WHERE sheet=? AND row=? AND col=?", changed)
    return str(dst), {"link": idx, "link_cells": len(ext), "cells": len(changed), "db_path": str(dst),
                      "why": "the overlay's link to the client model was saved without values: filled from last year's "
                             "client model, and the overlay's formulas worked out again from it"}


def discount_date_cells(path: str, outputs: list[dict], sheets=None,
                        cores: list[dict] | None = None) -> tuple[list[str], list[str]]:
    """The valuation dates the discountings under the outputs read (dcftrace), and the cells to move for them: each
    followed back through plain references (DCF_High!C3 = Inputs!C4) to the cell it's typed in, so a copy and
    everything else reading the date move together; a discounting with a date of its own has it moved too.
    cores: the discountings, where they're traced already. -> (the cells to move, the cells the discountings read)."""
    import dcf
    move, read = [], []
    with _ro(path) as db:
        for c in cores if cores is not None else _cores(db, outputs):
            # the date the discounting reads: fitted from its factors, else followed from their formulas (loose)
            v = (c.get("inputs") or {}).get("valuation_date") or (c.get("loose") or {}).get("valuation_date")
            r = dcf._ref(v, "") if isinstance(v, str) else None
            if not r:
                continue
            at, root = _a1(r[0], r[1], r[2]), _typed_in(db, r[0], r[1], r[2], sheets)
            read += [at] if at not in read else []
            move += [root] if root not in move else []
    return move, read


def cutoff_cells(path: str, outputs: list[dict], cores: list[dict] | None = None) -> list[list[str]]:
    """Each discounting's per-period cells under the outputs, with the date its period ends, for the roll's cut-off
    (Session._cut_off): [[cell, period end (ISO), the cell that date is in (where read from a row of ends)]]. Its present-value row where it sums one, else its factor row,
    else its cash flows (the factors worked out in its formula). Not an XNPV or an NPV: they count from their own
    first date or column, which the roll doesn't move."""
    import dcf
    out, seen = [], set()
    with _ro(path) as db:
        for c in cores if cores is not None else _cores(db, outputs):
            if c.get("kind") in ("xnpv", "npv") or not c.get("inputs"):
                continue
            try:
                sheet, r, cols = dcf._row_range(db, c.get("pv_row") or c.get("factor_row") or c["inputs"]["cashflow"][0])
                ends, note = dcf.period_ends(db, sheet, cols, c["inputs"].get("dates"))
            except (ValueError, KeyError, IndexError):
                continue
            # the row the ends were read from, for its cells on this year's feed (not where they were derived from starts)
            m = re.match(r"(.+?)!r(\d+)", note) if "derived" not in note else None
            for col in cols:
                if col in ends and (sheet, r, col) not in seen:
                    seen.add((sheet, r, col))
                    out.append([_a1(sheet, r, col), ends[col].isoformat()] + ([_a1(m[1], int(m[2]), col)] if m else []))
    return out


def _typed_in(db, sheet: str, row: int, col: int, sheets=None, hops: int = 8) -> str:
    """The cell a value is typed into, following plain references back (=Inputs!C4, =$C$4, =Val_Date: a name for one
    cell) within the overlay's own sheets; where a formula does more than refer (EOMONTH(...), another workbook),
    that cell itself. Moving the copy and not the cell it copies would leave everything else reading that cell on
    last year's date: an overlay's cash flow dates counted from it, its discount periods from the copy."""
    import dcf
    for _ in range(hops):
        f = db.execute("SELECT formula FROM cells WHERE sheet=? AND row=? AND col=?", (sheet, row, col)).fetchone()
        body = ((f[0] if f else None) or "").lstrip("=+ ").strip()
        cell = r"(?:(?:'[^'\[\]]+'|[A-Za-z0-9_.]+)!)?\$?[A-Z]{1,3}\$?\d+"
        if not re.fullmatch(cell, body) and re.fullmatch(r"[A-Za-z_\\][\w.]*", body):
            # a name: the cell it stands for, the sheet's own name first
            named = db.execute("SELECT ref FROM names WHERE lower(name)=lower(?) ORDER BY scope IS NOT ?, scope IS NOT NULL",
                               (body, sheet)).fetchone()
            body = (named[0] or "").lstrip("=+ ").strip() if named else ""
        if not re.fullmatch(cell, body):
            break
        r = dcf._ref(body, sheet)
        if not r or (sheets and r[0] not in sheets):
            break  # a client sheet's cell is fed from the client model, not moved in the overlay
        sheet, row, col = r[0], r[1], r[2]
    return _a1(sheet, row, col)


def date_cells(path: str, outputs: list[dict], sheets, lever: dict | None) -> dict:
    """The roll's valuation date cells: {"valuation_date_cells" (to move), "valuation_date_reads" (the cells the
    discountings read, checked on this year's feed), "valuation_date_cell" (the first, shown), "valuation_date_by_label"
    (none traced: the cell labelled as the valuation date is moved, unconfirmed), "cutoff" (each discounting's
    per-period cells and their period ends: cutoff_cells)}."""
    with _ro(path) as db:
        cores = _cores(db, outputs)
        move, read = discount_date_cells(path, outputs, sheets, cores)
        # the cell labelled as the valuation date moves too, followed to where it's typed: an overlay can keep one
        # date for its discountings and another for the rest (its period flags, its balances), both last year's
        root = _typed_in(db, *parse_a1(lever["cell"]), sheets) if lever else None
    if root and root not in move:
        move.append(root)
    return {"valuation_date_cells": move, "valuation_date_reads": read, "valuation_date_cell": move[0] if move else None,
            "valuation_date_by_label": bool(move) and not read, "cutoff": cutoff_cells(path, outputs, cores),
            "date_cells_plan": DATE_CELLS}


def horizon(prior: Workbook, current: Workbook, sheets, sheet_for=None) -> tuple[str | None, dict]:
    """How this year's client model's periods end against last year's: "fixed" where most client sheets with a
    timeline end on the same date both years (a concession or asset life with an end date: the period dates are the
    same dates, so rolling forward moves the valuation date, not the periods; moving them would read periods past
    the end and stand last year's values in for them), "rolling" where they end later, None with no sheet to tell
    by. sheet_for: last year's sheet -> this year's (rowfind), else the same name. -> (kind, {"fixed": n, "rolling": n})."""
    n = {"fixed": 0, "rolling": 0}
    for s in sorted(sheets or []):
        to = sheet_for(s) if sheet_for else s
        a, b = prior.timeline(s), current.timeline(to) if to else {}
        if len(a) >= 2 and len(b) >= 2:
            n["fixed" if max(a.values()) == max(b.values()) else "rolling"] += 1
    if not n["fixed"] and not n["rolling"]:
        return None, n
    return ("fixed" if n["fixed"] > n["rolling"] else "rolling"), n


def roll_months(sess: Session, prior: dict | None, overlay: dict, same_file: bool,
                dates: tuple = (None, None, None, None)) -> tuple[int, str, str | None]:
    """How far to roll forward: (months, how that was worked out, the new valuation date).
    dates: (last year's valuation date: the overlay's, from the report; last year's client model's; this year's
    client model's; this year's valuation date as set for the engagement).
    0. This year's valuation date set for the engagement: from last year's to it.
    1. From last year's valuation date to this year's client model's, when both are known and this year's is
       later: the new valuation date is this year's model's. (Not the move between the two client models: the
       overlay can sit on a copy of a client model of another date.)
    Where this year's model isn't dated after last year's valuation date, that date is likely the model's own,
    not this year's valuation date: nothing is rolled (0 months, "check:" in the basis) and the figures wait for
    this year's valuation date. Guessing a roll from anything else would put numbers on the page that aren't
    this year's.
    2. Else, only when last year's valuation date isn't known, the move between the two client models' dates.
    3. Else the client sheets' timelines: how far each sheet's first period moved (each of last year's sheets
       against the sheet it is this year), the most common move, the smaller on a tie; only forward, at most
       ROLL_MAX months, and from at least ROLL_SHEETS sheets.
    4. Else 12 months, flagged so the page asks for a check."""
    ov_vd, prior_vd, current_vd, this_vd = (list(dates) + [None] * 4)[:4]
    base = ov_vd or prior_vd
    if this_vd and base and this_vd[:10] > base[:10]:
        return (months_between(base[:10], this_vd[:10]),
                f"from last year's valuation date ({base[:10]}) to this year's ({this_vd[:10]}), set for the engagement",
                this_vd[:10])
    if ov_vd and current_vd and current_vd[:10] > ov_vd[:10]:
        return (months_between(ov_vd[:10], current_vd[:10]),
                f"from last year's valuation date ({ov_vd[:10]}) to this year's client model's ({current_vd[:10]})",
                current_vd[:10])
    if ov_vd and current_vd:
        return 0, (f"check: this year's model's date ({current_vd[:10]}) isn't after last year's valuation date "
                   f"({ov_vd[:10]}), so it's likely the model's own date, not this year's valuation date; nothing is "
                   "rolled until this year's valuation date is set"), ov_vd[:10]
    if not ov_vd and prior_vd and current_vd and current_vd[:10] > prior_vd[:10]:
        m = months_between(prior_vd[:10], current_vd[:10])
        return m, (f"from the valuation dates in the two client models ({prior_vd[:10]} to {current_vd[:10]}); last "
                   "year's own valuation date isn't known"), current_vd[:10]
    new = lambda m: to_date(add_months(serial(date.fromisoformat(base[:10])), m)).isoformat() if base else None
    src = sess.prior or sess.ov
    sheets = sess.client_sheets or ((prior or {}).get("sheets") if same_file else None)
    moves = Counter()
    for s in sorted(sheets or {k[1] for k in sess.ext_cached} or []):
        to = sess.rowmap.sheet_for(s) if sess.rowmap else s
        a, b = src.timeline(s), sess.current.timeline(to) if sess.current and to else {}
        if a and b:
            moves[months_between(to_date(min(a.values())).isoformat(), to_date(min(b.values())).isoformat())] += 1
    seen = sum(moves.values())
    forward = sorted(((n, -m) for m, n in moves.items() if 0 < m <= ROLL_MAX), reverse=True)
    if forward and seen >= ROLL_SHEETS:
        n, m = forward[0][0], -forward[0][1]
        return m, (f"from the client sheets' timelines (the first period moved {m} months on {n} of {seen} "
                   "sheet(s))"), new(m)
    return 12, ("assumed: the valuation dates don't give a roll and the timelines don't show one"
                + (f" (moves seen on {seen} sheet(s): {', '.join(f'{m:+d}' for m in sorted(moves))} months)" if moves else "")
                + "; check the roll-forward"), new(12)


ZERO_ROLL = (0.75, 1.33)  # this year's model at last year's date, against last year's figure: about the same
REBUILT = 0.5  # the two client models share fewer line-item labels than this (rowfind.family): this year's is rebuilt
ZERO_ROLL_CHECK = (0.87, 1.15)  # inside ZERO_ROLL but outside this, the value runs and a person is asked to confirm
                                # the move is the new forecast (a judgment call: forecasts move, a mismatched row too)
BALANCES = 2  # balance_cells' version: the cells a session found by older rules are found again when it loads


def balance_cells(sess: Session, summary: dict) -> list[dict]:
    """The client cells the value reads as a balance at last year's valuation date (a net debt, a cash balance, a
    distribution declared at the date): read by an overlay cell outside the overlay's periods (a summary cell, not a
    period column reading its own period), from a client row running over its timeline, in the column dated last
    year's valuation date. A row the discountings read as cash flows across its columns can be one too, where a
    separate cell reads it once at the date (the distribution declared at the date, deducted): only that read is the
    balance ("series": the row is also read as cash flows). Rolled forward, each is read at this year's date
    (Session._rolled). Walked down from the overlay's outputs on last year's feed. -> [{"cell": [sheet, row, col],
    "series", "reader", "date"}]."""
    if sess.base_vd is None:
        return []
    outs = [parse_a1(o["cell"]) for o in summary.get("outputs") or [] if o.get("cell")]
    if not outs:
        return []
    prior = sess.prior or sess.ov
    readers = defaultdict(set)  # client (sheet, row, col) -> the overlay cells reading it
    sess.configure("prior" if sess.prior else "workbook")
    try:
        sess.values(outs)
        B, seen, q = sess.B, set(), deque(("", s, r, c) for s, r, c in outs)
        while q and len(seen) < 60000:
            x = q.popleft()
            if x in seen:
                continue
            seen.add(x)
            src, s, r, c = x
            if src != "" or s in (sess.client_sheets or ()) or not B.is_formula(s, r, c):
                continue
            reads, _ = B.reads(s, r, c)
            for y in reads:
                ys, yr_, yc = y[1], y[2], y[3]
                if y[0] != "" or ys in (sess.client_sheets or ()):
                    if y[0] in ("", sess.client_link):
                        readers[(ys, yr_, yc)].add((s, r, c))
                elif y not in seen:
                    q.append(y)
    finally:
        sess.configure("workbook")
    tl_ov = {}

    def aligned(x, d):  # an overlay period column reading its own period: a cash flow, not a balance
        if x[0] not in tl_ov:
            tl_ov[x[0]] = sess.ov.timeline(x[0])
        return x[2] in tl_ov[x[0]] and round(tl_ov[x[0]][x[2]]) == round(d)
    rows = defaultdict(lambda: {"single": set(), "aligned": False})
    for (s, r, c), xs in readers.items():
        tl = prior.timeline(s)
        if c not in tl:
            continue
        for x in xs:
            if aligned(x, tl[c]):
                rows[(s, r)]["aligned"] = True
            else:
                rows[(s, r)]["single"].add((c, x))
    out = []
    for (s, r), v in rows.items():
        tl = prior.timeline(s)
        if sum(1 for cc in tl if isinstance(prior.value(s, r, cc), (int, float))) < 2:
            continue  # not a row running over its timeline
        if sum(1 for cc in tl if same(prior.value(s, r, cc), tl[cc])) >= 2:
            continue  # the period dates themselves (an INDEX/MATCH looking up the valuation date's column reads them)
        for c, x in sorted(v["single"]):
            if round(tl[c]) == round(sess.base_vd):
                out.append({"cell": [s, r, c], "series": v["aligned"], "reader": _a1(*x), "date": to_date(tl[c]).isoformat()})
    return sorted(out, key=lambda d: d["cell"])


def set_balances(sess: Session, cells: list) -> None:
    """The session's balances from balance_cells' list: {(sheet, row): {"cols", "series"}}."""
    sess.balances = {}
    for x in cells or []:
        if isinstance(x, dict):
            b = sess.balances.setdefault(tuple(x["cell"][:2]), {"cols": set(), "series": False})
            b["cols"].add(x["cell"][2])
            b["series"] = b["series"] or bool(x.get("series"))


ROLL_PLAN = 4  # the rules' version: a roll planned by older rules is planned again when a session loads
DATE_CELLS = 6  # date_cells' version: the cells a build traced with older rules are traced again when a session loads
ROLL_MAX = 24  # months: a move beyond it from the timelines isn't a roll-forward
ROLL_SHEETS = 5  # sheets: fewer can't show how far the timelines moved


def plan_roll(sess: Session, prior: dict | None, overlay: dict, same_file: bool, ov_vd: str | None,
              prior_vd: str | None, current_vd: str | None, this_vd: str | None = None) -> dict:
    """The roll-forward's settings (roll_months), and last year's valuation date set on the session."""
    months, basis, new_vd = roll_months(sess, prior, overlay, same_file, (ov_vd, prior_vd, current_vd, this_vd))
    sess.base_vd = serial(date.fromisoformat(ov_vd[:10])) if ov_vd else None
    sess._pshift.clear()
    if sess.rowmap and getattr(sess.rowmap, "since", None) != sess.base_vd:  # the finder compares rows from it on
        sess.rowmap.since = sess.base_vd
        sess.rowmap._cache.clear()
    return {"prior_valuation_date": ov_vd, "months": months, "months_basis": basis,
            "fixed_horizon": bool(sess.current) and sess.fixed_horizon(), "horizon_set": getattr(sess, "horizon_set", None),
            "months_assumed": "assumed:" in basis, "date_check": basis.startswith("check:"),
            "current_valuation_date": new_vd, "plan": ROLL_PLAN,
            "dates": {"overlay": ov_vd, "prior_client": prior_vd, "current_client": current_vd, "this_year": this_vd}}


def _feed(summary: dict, mode: str, valuation_date: str | None, months: int | None,
          held: bool = True, rates: bool = True) -> tuple[dict, dict | None, int]:
    """A feed's own settings, the base that a person's changes go on top of: rolled forward, the months to roll,
    the new valuation date on its lever, this year's figures a person set for inputs otherwise held at last
    year's (held.py; held=False leaves them at last year's, for the bridge's step), and this year's discount rate
    where a person set it, on the cells the discountings read for it (result.this_year_rate; rates=False leaves
    last year's, for the bridge's step and the zero-roll check). -> (overrides, roll info or None, months)."""
    if mode != "current":
        return {}, None, 0
    roll = summary.get("roll") or {}
    chosen = valuation_date is not None  # a date chosen on the page, rather than the plan's
    valuation_date = valuation_date or roll.get("current_valuation_date")
    pvd = roll.get("prior_valuation_date")
    if months is None:  # the plan's months; a date chosen on the page moves the periods as far as it moves from last year's
        months = months_between(pvd[:10], valuation_date[:10]) if chosen and pvd else roll.get("months", 12)
    vd_cells = roll.get("valuation_date_cells") or ([roll["valuation_date_cell"]] if roll.get("valuation_date_cell") else [])
    defaults = {parse_a1(c): serial(date.fromisoformat(valuation_date[:10])) for c in vd_cells} if valuation_date else {}
    if held:
        defaults.update({parse_a1(c): float(x["value"]) for c, x in (summary.get("held_values") or {}).items()
                         if x.get("value") is not None})
    if rates:
        defaults.update({parse_a1(c): float(v) for c, v in (summary.get("rate_values") or {}).items()})
    # the overlay's own copy of the forecast's end, where this year's model forecasts to another date (result.py)
    defaults.update({parse_a1(c): float(v) for c, v in (summary.get("horizon_values") or {}).items()})
    return defaults, {"months": months, "valuation_date": valuation_date, "valuation_date_cell": vd_cells[0] if vd_cells
                      else None, "valuation_date_cells": vd_cells}, months


def scenario(sess: Session, summary: dict, mode: str, changes: dict, valuation_date: str | None = None,
             months: int | None = None) -> dict:
    """Run the overlay with lever / cell changes ({"Sheet!A1": value}) on a feed. Returns outputs against the base
    (same feed, no changes) and the workbook, what the roll-forward changed, and feed cells it couldn't match."""
    outputs = summary["outputs"]
    cells = [parse_a1(o["cell"]) for o in outputs]
    extra = {parse_a1(k): v for k, v in changes.items()}
    defaults, roll_info, months = _feed(summary, mode, valuation_date, months)
    sess.configure(mode, defaults, months or 0)
    base = sess.values(cells)
    base_unmatched = dict(sess.unmatched)
    sess.configure(mode, {**defaults, **extra}, months or 0)
    got = sess.values(cells)
    unmatched = {**base_unmatched, **sess.unmatched}
    if mode == "current":
        tl = sess.rolled_timeline(months or 0)
        firsts = sorted(tl.values())
        roll_info.update(first_period=to_date(firsts[0]).isoformat() if firsts else None,
                         last_period=to_date(firsts[-1]).isoformat() if firsts else None, periods=len(set(firsts)))
    rows = []
    for o, b, v in zip(outputs, base, got):
        wb = o.get("value")
        rows.append({"cell": o["cell"], "label": o["label"], "workbook": wb, "base": _show(b), "value": _show(v),
                     "change": (v - b) if isinstance(v, float) and isinstance(b, float) else None,
                     "vs_workbook": (v - wb) if isinstance(v, float) and isinstance(wb, float) else None,
                     "report": o.get("report"), "tie": tie(v, o["report"], o["scale"], o["sign"]) if o.get("report") and mode != "current" else None})
    sess.configure("workbook")
    return {"mode": mode, "changes": changes, "outputs": rows, "roll": roll_info,
            "unmatched": [{"cell": _a1(*k), "why": w} for k, w in list(unmatched.items())[:100]],
            "n_unmatched": len(unmatched)}


# ---- this year's timing rows, worked out from the period dates ------------------------------------------------


def timing_rule(wb: Workbook, s: str, r: int, vd: float | None) -> dict | None:
    """How a timing row (period flags, period dates) follows from its sheet's period dates, if one relation holds
    for every period last year: the period's date as the header shows it, its end, its start, the year it ends
    in (plus a constant), 1 after the valuation date (or up to it) and 0 otherwise, or one constant. Then this
    year's value is worked out from the rolled period date instead of hunting for a row: the period grid is a
    property of the model."""
    tl = wb.timeline(s)
    vals = wb.sheet(s)
    pts = [(tl[c], vals.get((r, c))) for c in sorted(tl)]
    pts = [(t, float(v)) for t, v in pts if isinstance(v, (int, float)) and not isinstance(v, bool)]
    if len(pts) < 2:
        return None
    dates = sorted(set(tl.values()))
    gaps = sorted(b - a for a, b in zip(dates, dates[1:]))
    plen = max(1, round(gaps[len(gaps) // 2] / 30.44)) if gaps else 12
    starts = all(to_date(d).day == 1 for d in dates)
    base = {"plen": plen, "starts": starts}
    tries = [("timeline", "the period's date"), ("end", "the period's end"), ("start", "the period's start")]
    for kind, text in tries:
        rule = {**base, "kind": kind, "text": f"{text}, from the period dates"}
        if all(abs(timing_value(rule, t, vd) - v) < 0.5 for t, v in pts):
            return rule
    off = {v - to_date(_period_end(base, t)).year for t, v in pts}
    if len(off) == 1:
        o = off.pop()
        if abs(o) <= 1:
            return {**base, "kind": "year", "offset": o, "text": "the year the period ends in, from the period dates"}
    if vd is not None and {v for _, v in pts} <= {0.0, 1.0}:
        for kind, text in (("after", "1 for periods ending after the valuation date"),
                           ("upto", "1 for periods ending by the valuation date")):
            rule = {**base, "kind": kind, "text": f"{text}, from the period dates"}
            if all(timing_value(rule, t, vd) == v for t, v in pts):
                return rule
    if len({v for _, v in pts}) == 1:
        return {**base, "kind": "const", "value": pts[0][1], "text": f"the same in every period ({pts[0][1]:g})"}
    return own_rule([v for _, v in pts])


def own_rule(vals: list[float]) -> dict | None:
    """A row that carries its own periods, one a column, whatever the sheet's header says (an annual block laid out
    under a quarterly header): dates a whole number of months apart, on the same day or each a month's end (the last
    may be a short stub, a concession ending part-way through a year), or year numbers one apart. Rolled forward, it
    moves by its own periods that ended between the two valuation dates (Session._own_value), not by the header's."""
    if len(vals) < 3 or not all(float(v).is_integer() for v in vals):
        return None
    if all(1900 <= v <= 2200 for v in vals):
        if all(b - a == 1 for a, b in zip(vals, vals[1:])):
            return {"kind": "own_year", "plen": 12, "text": "year numbers, one a column, carried by the row"}
        return None
    if not all(3000 < v < 120000 for v in vals):
        return None
    ds = [to_date(v) for v in vals]
    gaps = [(b.year - a.year) * 12 + b.month - a.month for a, b in zip(ds, ds[1:])]
    body = gaps[1:-1] if len(gaps) > 2 else gaps  # the first and the last may be stubs (a concession's part-years)
    month_end = lambda d: xlruntime._add_months(d, 0, end=True) == d
    inner = ds[1:-1] if len(ds) > 3 else ds
    regular = all(month_end(d) for d in inner) or len({d.day for d in inner}) == 1
    if len(set(body)) != 1 or not body[0] > 0 or not all(0 < x <= body[0] for x in (gaps[0], gaps[-1])) or not regular:
        return None
    plen = body[0]
    starts = all(d.day == 1 for d in ds)
    ends = [add_months(v, plen) - 1 if starts else v for v in vals]
    stubs = [w for w, x in (("first", gaps[0]), ("last", gaps[-1])) if x < plen]
    return {"kind": "own_date", "plen": plen, "ends": ends,
            "text": f"dates {plen} months apart, one a column, carried by the row"
                    + (f" (the {' and the '.join(stubs)} a stub)" if stubs else "")}


def _period_end(rule: dict, t: float) -> float:
    return add_months(t, rule["plen"]) - 1 if rule["starts"] else t


def timing_value(rule: dict, t: float, vd: float | None) -> float:
    """A timing rule's value for the period whose header date is t, with valuation date vd."""
    end = _period_end(rule, t)
    k = rule["kind"]
    if k == "timeline":
        return t
    if k == "end":
        return end
    if k == "start":
        return t if rule["starts"] else add_months(t, -rule["plen"]) + 1
    if k == "year":
        return float(to_date(end).year + rule["offset"])
    if k == "after":
        return 1.0 if vd is not None and end > vd else 0.0
    if k == "upto":
        return 1.0 if vd is not None and end <= vd else 0.0
    return rule["value"]


def dcf_origins(sess: Session, summary: dict, cells: list[str]) -> dict:
    """For each figure, the client rows the discountings under it take their cash flows from (dcffacts.py):
    {cell: {"amounts": [(sheet, row)], "timing": [(sheet, row, why)]}}. The amounts are the rows this year's model
    must have for the figure to be this year's; the timing rows (period flags and dates) are what the discounting's
    timing depends on, listed but not asked for. Cached on the session. (Facts.run resets the feed: call it before
    reading a feed whose unmatched values are wanted.)"""
    import dcffacts
    cache = sess.__dict__.setdefault("_origins", {})
    out = {}
    for cell in cells:
        if cell not in cache:
            amounts, timing = set(), set()
            try:
                fx = dcffacts.Facts(sess, summary).run(cell)
                for c in fx["discountings"]:
                    for o in c["origins"]:
                        m = re.match(r"^(?:\[\d+\])?(.+)!r(\d+)$", o["row"])
                        if not m:
                            continue
                        if o.get("kind") == "timing":
                            timing.add((m[1], int(m[2]), o.get("kind_why") or ""))
                        else:
                            amounts.add((m[1], int(m[2])))
            except Exception:
                pass
            cache[cell] = {"amounts": sorted(amounts), "timing": sorted(timing)}
        out[cell] = cache[cell]
    return out


# ---- the DCF on the module's own numbers ----------------------------------------------------------------------
# dcf.py recomputes a discounting from the cash flows, dates and bridge amounts the module computes on a feed.
# dcf.py reads a model.db, so the module's values reach it through rodb.patched().

def _dcf_cells(db, inputs: dict) -> set[tuple]:
    """Every cell a DCF's inputs read: cash-flow rows, their period-end dates, the rate, valuation date and
    cut-off cells, bridge cells or rows, and the anchor cell itself."""
    import dcf
    cells: set[tuple] = set()

    def one(ref):
        if ref is None or isinstance(ref, (int, float)) or dcf._as_date(str(ref)):
            return
        t = str(ref).strip()
        named = db.execute("SELECT ref FROM names WHERE lower(name)=lower(?)", (t,)).fetchone()
        t = named[0] if named else t
        m = dcf._REF.match(t) or (dcf._REF.match(t.replace(" ", "")) if "'" not in t else None)
        if m and (not m["c2"] or (m["c2"], m["r2"]) == (m["c1"], m["r1"])):
            cells.add((m["sheet"].strip("'"), int(m["r1"]), dcf._col(m["c1"])))

    def row(ref):
        sheet, r, cols = dcf._row_range(db, ref)
        cells.update((sheet, r, c) for c in cols)
        _, src = dcf.period_ends(db, sheet, cols)
        m = re.match(r"^(.+?)!r(\d+)\b", src)
        if m:
            cells.update((m[1], int(m[2]), c) for c in cols)
        return sheet, cols

    for ref in [*inputs["cashflow"], *(inputs.get("mask") or [])]:
        row(ref)
    if inputs.get("dates"):
        row(inputs["dates"])
    for k in ("rate", "valuation_date", "terminal_date", "compare_to"):
        one(inputs.get(k))
    for a in inputs.get("adjustments") or []:
        v = a.get("value") if isinstance(a, dict) else a
        if isinstance(a, dict) and a.get("at_valuation_date"):
            try:
                row(str(v))
                continue
            except ValueError:
                pass
        one(v)
    return cells


def _module_values(db, sess: Session, cells: set[tuple]) -> dict:
    """The module's value for each cell, stored the way model.db stores it (dates as ISO text) so dcf.py reads it."""
    keys = sorted(cells)
    got = sess.values(keys)
    out = {}
    for k, v in zip(keys, got):
        saved = db.execute("SELECT value FROM main.cells WHERE sheet=? AND row=? AND col=?", k).fetchone()
        saved = saved[0] if saved else None
        if isinstance(v, xlruntime.XLError):
            v = v.code
        elif isinstance(v, float) and isinstance(saved, str) and xlruntime._DATE.match(saved) and 0 < v < 2958466:
            v = to_date(v).isoformat()
        out[k] = v
    return out


def _flows(db, inputs: dict) -> tuple[dict, dict]:
    """(cash flow by column, summed over the cash-flow rows; period end date by column) as db holds them."""
    import dcf
    flows, sheet0, cols0 = {}, None, None
    for ref in inputs["cashflow"]:
        sheet, r, cols = dcf._row_range(db, ref)
        sheet0, cols0 = sheet0 or sheet, cols0 or cols
        for c, v in db.execute("SELECT col, value FROM cells WHERE sheet=? AND row=? AND col BETWEEN ? AND ?",
                               (sheet, r, cols[0], cols[-1])):
            if dcf._num(v) is not None:
                flows[c] = flows.get(c, 0.0) + dcf._num(v)
    flows = dcf.apply_mask(db, flows, cols0, inputs.get("mask"))
    return flows, dcf.period_ends(db, sheet0, cols0, inputs.get("dates"))[0]


def _brief(r: dict) -> dict:
    iso = lambda d: d.isoformat() if d else None
    return {"rate": r["rate"], "pv": r["pv"], "total": r["total"], "anchor": r["compare_to"],
            "valuation_date": iso(r["valuation_date"]), "timing": r["timing"], "day_count": r["day_count"],
            "terminal_date": iso(r["terminal_date"]), "bridge": r["bridge"], "periods": r["periods"],
            "first_period": iso(r["first_period"]), "last_period": iso(r["last_period"]), "undiscounted": r["undiscounted"]}


def trace_starts(summary: dict) -> list[dict]:
    """Where to start tracing the valuation: the cells the report's conclusions were matched to, those that tie
    to the report at its printed precision first, then label matches with no report figure to tie to. A match
    that differs from the report is never used (it's likely the wrong cell)."""
    outs = [o for o in summary.get("outputs") or [] if o.get("fact_id")]
    tied = [o for o in outs if (o.get("tie") or {}).get("ok")]
    rest = [o for o in outs if not o.get("tie") and re.search(r"value|valuation|\bnpv\b", o.get("label") or "", re.I)]
    out, seen = [], set()
    for o in tied + rest:
        if o["cell"] not in seen:
            seen.add(o["cell"])
            out.append({"cell": o["cell"], "label": o.get("label"), "report": o.get("report"), "key": o.get("key"),
                        "ties": bool((o.get("tie") or {}).get("ok")), "value": o.get("value")})
    return out


def dcf_anchors(summary: dict) -> list[dict]:
    """The DCFs valuation.py finds on the overlay sheets (all of the workbook's if none are there), then the
    discountings traced down from the report's conclusions (dcftrace.py) that those don't already cover."""
    import dcftrace
    import valuation
    path = summary["wiring"]["overlay"]["db_path"]
    cat = valuation.catalogue(path)
    mine = [a for a in cat if a["cell"].split("!")[0].strip("'") in set(summary["sheets"])] or cat
    have = {a["cell"] for a in mine} | {a.get("pv_cell") for a in mine}
    try:
        traced = dcftrace.anchors(path, [s["cell"] for s in trace_starts(summary)])
    except Exception:  # the trace is extra: the label-based ones still stand
        traced = []
    return mine + [a for a in traced if a["cell"] not in have]


if __name__ == "__main__":
    import argparse
    sys.path.insert(0, str(Path(__file__).parent))
    ap = argparse.ArgumentParser(description="Run an engagement's Python overlay")
    ap.add_argument("engagement", type=int)
    ap.add_argument("--mode", default="workbook", choices=["workbook", "prior", "current"])
    ap.add_argument("--set", action="append", default=[], metavar="SHEET!A1=VALUE", help="change an input (repeatable)")
    ap.add_argument("--valuation-date", help="YYYY-MM-DD, for --mode current")
    a = ap.parse_args()
    import workbench
    changes = dict(x.split("=", 1) for x in a.set)
    res = workbench.overlay_run(a.engagement, a.mode, changes, a.valuation_date, None)
    if res.get("roll"):
        r = res["roll"]
        print(f"rolled forward {r['months']} months to {r['valuation_date']}: periods {r.get('first_period')} to "
              f"{r.get('last_period')} ({r.get('periods')}); {res['n_unmatched']} client value(s) unmatched")
    for o in res["outputs"]:
        ch = f"  change {o['change']:+,.2f}" if isinstance(o.get("change"), float) and changes else ""
        print(f"{o['cell']:<18} {o['label'][:40]:<40} {o['value']:>16,.2f}{ch}")
