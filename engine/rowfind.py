"""Last year's client-model rows in this year's model, when the model changed between valuations.

Rolling a valuation forward means reading, for every client row the overlay reads, the same line item in this
year's model. Models change between valuations: rows are inserted, renamed, restructured, sheets renamed. A row
is looked for in several independent ways, and the one the evidence supports best is taken:
  label        the same label on the same sheet (its n-th occurrence, else the nearest), or, where the sheet is
               gone (renamed), the only row with that label anywhere
  history      the same values in the periods both models have as history: actual years don't change between
               versions, so a row whose actuals equal last year's is the same line item whatever it's called now;
               forecast years that stay close (revised, not replaced) count too
  words        a label on the same sheet sharing most of its words ("Free cash flow" -> "Free cash flow after
               working capital"), with the forecast staying close
  neighbours   the same rows around it in the dependency graph (the edges table): it reads rows with the same
               labels (EBITDA, capital expenditure) and is read by the same ones
  banner       a model's summary cells in the first rows of its sheets (a total, a value at the valuation date,
               often repeated on several sheets) that read the row last year: the same banner this year reads
               the row to use
  trace        the rows both models share, by their label and their numbers (anchors), followed to it through
               the dependency graph: down from the ones that read it (an output: distributions, the cash flow
               available), up from the ones it reads (an input: a volume, a tariff), each step the row paired with
               last year's by its label, its numbers or its words, else as the one left once the others are paired.
               It follows formulas, not sheets, so a row moved to a new sheet, or with rows put in between, is
               still reached; two ways agreeing count for more, and a trace that leads elsewhere than the label
               leaves the row to be looked at
A candidate of another shape counts for less: last year's row calculated (formulas) and this one typed values,
or the other way round. A reconciliation sheet of pasted values ("LINKED EBITDA") has last year's history exactly
and a full series, and would otherwise win every row it copies.
A person's pick for a row always wins: a row, or last year's values kept on purpose (STAND_IN). explain() gives what was found for a row and why, with the alternatives.
A row with nothing to find (no label, no numbers, no formulas last year: a spacer inside a range a formula reads) is
settled by code: its blanks stand in, and no model is asked about it.
"""
import bisect
import re
from collections import defaultdict

BANNER_ROWS = 10
EXACT = 1e-9
SAME_SHEET = 0.35     # a sheet of the same name is the same sheet only if it shares this much of its labels
RENAMED_SHEET = 0.5   # a sheet of another name is a renamed one if it shares this much
RENAMED_TIE = 0.05    # ... and the next as like it within this: neither is taken (which one is a person's to say)
LINEAGE_CANDS = 12    # candidates whose lineage is compared with last year's row's, at most
# kinds of row that are what they are by their values: a row of one of these isn't a row of another kind
FIXED_KINDS = {"dates", "flags", "factors", "share", "index"}
# the kinds of evidence: who a row is (its label, its words), where it is (a banner reading it, its block, the same
# place), what it does (what it reads and what reads it, a trace through them), its numbers (last year's history)
FAMILY = {"label": "identity", "words": "identity", "banner": "place", "block": "place", "neighbours": "role",
          "trace": "role", "lineage": "role", "history": "numbers"}
CONFIDENT = 0.5      # a row found with less than this needs a person's look before its figures are this year's
CHECK_GAP = 0.15     # a row carries last year's numbers when its median difference from them is within this
CHECK_PERIODS = 3    # ... over at least this many periods side by side
SHAPE = 0.6           # a candidate whose share of formulas differs from last year's row's by this much is another shape
STAND_IN = "stand-in"  # a person's pick: keep last year's values for the row
LAYOUT = 0.9         # an unlabelled row found at its own row number, on a sheet laid out as before
TRACE_DEPTH = 6      # rows apart, at most, an anchor and the row the trace is for
TRACE_ANCHORS = 6    # the nearest anchors each way
TRACE_SCAN = 300     # rows looked at for anchors each way, at most
VERSION = 6          # bump when finding changes: the agents' picks made under another version are dropped and redone
# the rows a model's own valuation ends in: a copy of a block whose rows reach the row last year's reached is the case
# the model values (a downside or P90 copy beside it reaches none, or another)
VALUE_ROW = re.compile(r"\b(npv|net present value|present value|equity value|enterprise value|valuation|irr|dcf)\b", re.I)
NOT_VALUE_KINDS = ("dates", "factors", "flags", "share", "index")
                     # (4: the trace)
                     # (3: a pasted copy of last year's figures is no candidate)
                     # (2: unlabelled rows followed by the layout, blank rows settled by code)


def _norm(label: str) -> str:
    return re.sub(r"\s+", " ", (label or "").strip().lower())


def _words(label: str) -> set[str]:
    return set(re.findall(r"[a-z][a-z0-9]+", (label or "").lower())) - {"the", "and", "of", "for", "to", "in"}


def _jaccard(a: set, b: set) -> float:
    return len(a & b) / len(a | b) if a and b else 0.0


class RowFinder:
    def __init__(self, rowmap, prior, current, picks: dict | None = None, only: set | None = None):
        """rowmap: overlay.RowMap (the label matching); prior, current: overlay.Workbook; only: the sheets of current
        to look in (all of them by default)."""
        self.rowmap, self.prior, self.current = rowmap, prior, current
        self.only = set(only) if only is not None else None
        self.picks = dict(picks or {})   # {(sheet, row): (sheet, row) or STAND_IN}: a person's or the agents' choice
        self.pick_by: dict = {}          # {(sheet, row): "you" | "agent"}: whose choice it is (a person's by default)
        self._cache: dict = {}
        self._idx = None
        self._shapes: dict = {}
        self._numeric: dict = {}         # {sheet: {rows with a number other than 0}} in last year's model

    # ---- what each model has ---------------------------------------------------------------------------------
    def _index(self):
        if self._idx is not None:
            return self._idx
        cur = self.current
        sheets = [s for (s,) in cur.db.execute("SELECT sheet FROM sheets") if self.only is None or s in self.only]
        labels = cur.labels() if self.only is None else {k: v for k, v in cur.labels().items() if k[0] in self.only}
        by_label = defaultdict(list)
        for (s, r), lab in labels.items():
            if lab:
                by_label[_norm(lab)].append((s, r))
        dates = {s: {v: c for c, v in cur.timeline(s).items()} for s in sheets}
        by_word = defaultdict(list)
        for k, lab in labels.items():
            for w in _words(lab):
                by_word[w].append(k)
        self._idx = {"sheets": set(sheets), "labels": labels, "by_label": by_label, "dates": dates, "by_word": by_word,
                     "edges": (self._edges(self.prior), self._edges(cur)), "values": {}}
        return self._idx

    @staticmethod
    def _edges(wb) -> tuple[dict, dict]:
        """(reads, read_by): {(sheet, row): {(sheet, row)}} from the model's edges table."""
        reads, read_by = defaultdict(set), defaultdict(set)
        try:
            for s, r, ds, dr in wb.db.execute("SELECT src_sheet, src_row, dst_sheet, dst_row FROM edges "
                                              "WHERE kind IN ('direct', 'offset', 'active')"):
                if (s, r) != (ds, dr):
                    reads[(s, r)].add((ds, dr))
                    read_by[(ds, dr)].add((s, r))
        except Exception:
            pass
        return reads, read_by

    def sheet_for(self, s: str) -> str | None:
        """The sheet in this year's model that last year's sheet s is: the same name if it shares at least
        SAME_SHEET of its line-item labels (a rebuilt model can reuse a name for something else), else the sheet of
        another name sharing RENAMED_SHEET of them (renamed); None when no sheet does."""
        idx = self._index()
        idx.setdefault("sheet_map", {})
        if s in idx["sheet_map"]:
            return idx["sheet_map"][s]
        mine = {_norm(l) for (sh, _), l in self.prior.labels().items() if sh == s and l}
        labels_of = lambda sh: {_norm(l) for (x, _), l in idx["labels"].items() if x == sh and l}
        if s in idx["sheets"] and (not mine or _jaccard(mine, labels_of(s)) >= SAME_SHEET):
            idx["sheet_map"][s] = s
            return s
        scored = []
        for sh in sorted(idx["sheets"]):  # in order, so the same files give the same answer on every run
            if sh == s or self._has_prior_sheet(sh) or self._pasted_for(s, sh):
                continue
            scored.append((_jaccard(mine, labels_of(sh)), sh))
        scored.sort(key=lambda x: -x[0])
        best, share = (scored[0][1], scored[0][0]) if scored else (None, 0.0)
        # two of this year's sheets as like last year's (a nominal and a real copy, a 100% and a share): neither is
        # taken; its rows are left to find, both sheets for a person
        if len(scored) > 1 and scored[1][0] >= RENAMED_SHEET and scored[0][0] - scored[1][0] < RENAMED_TIE:
            best = None
        idx["sheet_map"][s] = best if share >= RENAMED_SHEET else None
        return idx["sheet_map"][s]

    def _has_prior_sheet(self, sheet: str) -> bool:
        return self.prior.db.execute("SELECT 1 FROM sheets WHERE sheet=?", (sheet,)).fetchone() is not None

    def _values_at(self, sheet: str, when: float) -> dict:
        """This year's model at one period date on one sheet: {rounded value: [rows]} (each sheet indexed once)."""
        idx = self._index()
        c = idx["dates"].get(sheet, {}).get(when)
        if c is None:
            return {}
        if sheet not in idx["values"]:
            by_col: dict = defaultdict(lambda: defaultdict(list))
            cols = set(idx["dates"][sheet].values())
            for (r, cc), v in self.current.sheet(sheet).items():
                if cc in cols and isinstance(v, float) and v:
                    by_col[cc][float(f"{v:.9g}")].append(r)
            idx["values"][sheet] = by_col
        return idx["values"][sheet].get(c, {})

    def _series(self, wb, sheet: str, row: int) -> dict:
        """A row's values by period date."""
        tl = wb.timeline(sheet)
        vals = wb.sheet(sheet)
        return {when: vals.get((row, c)) for c, when in tl.items() if isinstance(vals.get((row, c)), float)}

    def _same_series(self, a: tuple, b: tuple) -> bool:
        """Two of this year's rows hold the same figures in every period of the first (at least three of them) from
        the year to last year's valuation date on (since, set with the roll): what they did before it isn't valued."""
        x, y = self._series(self.current, *a), self._series(self.current, *b)
        since = getattr(self, "since", None)
        if since is not None:
            x = {w: v for w, v in x.items() if w > since - 366}
        if len(x) < 3 or not set(x) <= set(y):
            return False
        return all(abs(x[w] - y[w]) <= EXACT * max(1.0, abs(x[w])) for w in x)

    def _shape(self, wb, k) -> tuple | None:
        """(formulas, typed values) in a row, from the rows table (each workbook read once)."""
        if id(wb) not in self._shapes:
            try:
                self._shapes[id(wb)] = {(s, r): (nf or 0, nc or 0) for s, r, nf, nc in
                                        wb.db.execute("SELECT sheet, row, n_formula, n_const FROM rows")}
            except Exception:
                self._shapes[id(wb)] = {}
        return self._shapes[id(wb)].get(k)

    def _pasted_for(self, s: str, sh: str) -> bool:
        """Is this year's sheet sh typed values where last year's sheet s was formulas: a pasted copy (a
        reconciliation of last year's figures, say), which can't be the sheet s became, nor tell its horizon."""
        def share(wb, sheet):
            got = [self._shape(wb, (sheet, r)) for (x, r) in wb.labels() if x == sheet]
            f, c = sum(g[0] for g in got if g), sum(g[1] for g in got if g)
            return f / (f + c) if f + c else None
        a, b = share(self.prior, s), share(self.current, sh)
        return a is not None and b is not None and a >= 0.5 and b < 0.1

    def copies(self, s, r) -> list[str]:
        """The pasted copies of last year's row among the candidates its label and its history find ("Sheet!rN")."""
        got = set()
        for fn in (self._by_label, self._by_history):
            try:
                got |= {k for k, _sc, _t in fn(s, r) if k}
            except Exception:
                continue
        return [f"{k[0]}!r{k[1]}" for k in sorted(got) if self.is_copy(s, r, k)]

    def is_copy(self, s, r, k) -> bool:
        """Is candidate k a pasted copy of last year's row, not this year's line item: typed values where last
        year's row was formulas (other_shape), and last year's figures in every period both have, forecast
        included (forecasts are revised between valuations; a schedule that isn't is still built of formulas)."""
        if not self.other_shape(s, r, k):
            return False
        mine, theirs = self._series(self.prior, s, r), self._series(self.current, *k)
        both = [w for w in mine if w in theirs and mine[w]]
        return len(both) >= 3 and all(abs(theirs[w] - mine[w]) <= EXACT * max(1.0, abs(mine[w])) for w in both)

    def suspect_copy(self, s, r, k) -> str | None:
        """Why this year's row k looks like a copy of last year's figures rather than this year's line item, or None:
        last year's forecast exactly, in every period after last year's date, while other rows on last year's
        row's sheet were revised (a prior-forecast block, a reconciliation of last year's figures). Forecasts are
        revised between valuations; a whole row of them unchanged, beside rows that were, is a lookalike."""
        since = getattr(self, "since", None)
        mine, theirs = self._series(self.prior, s, r), self._series(self.current, *k)
        fut = [w for w in mine if (since is None or w > since) and w in theirs and mine[w]]
        same = lambda a, b: abs(a - b) <= EXACT * max(1.0, abs(a))
        if len(fut) < CHECK_PERIODS or not all(same(mine[w], theirs[w]) for w in fut):
            return None
        revised = looked = 0
        for (sh, rr), lab in sorted(self.prior.labels().items()):
            if sh != s or rr == r or not lab:
                continue
            hit = self.rowmap.match(sh, rr)[0]
            if not hit:
                continue
            a, b = self._series(self.prior, sh, rr), self._series(self.current, sh, hit)
            both = [w for w in a if (since is None or w > since) and w in b and a[w]]
            if len(both) >= CHECK_PERIODS:
                looked += 1
                revised += any(not same(a[w], b[w]) for w in both)
            if looked >= 12:
                break
        if not revised:
            return None
        return (f"last year's forecast exactly in all {len(fut)} periods after last year's date, where {revised} of "
                f"{looked} other rows on its sheet were revised: a copy of last year's figures?")

    def other_shape(self, s, r, k) -> str | None:
        """Why candidate k is another shape than last year's row (formulas against typed values), or None."""
        a, b = self._shape(self.prior, (s, r)), self._shape(self.current, k)
        if not a or not b or not sum(a) or not sum(b):
            return None
        fa, fb = a[0] / sum(a), b[0] / sum(b)
        if abs(fa - fb) < SHAPE:
            return None
        what = lambda x, f: f"formulas ({x[0]} of {sum(x)})" if f >= 0.5 else f"typed values ({x[1]} of {sum(x)})"
        return f"last year's row is {what(a, fa)}, this one {what(b, fb)}: another kind of row (a pasted copy?)"

    # ---- the strategies ----------------------------------------------------------------------------------------
    def _by_label(self, s, r) -> list[tuple]:
        idx = self._index()
        to = self.sheet_for(s)
        if to == s:
            r2, how = self.rowmap.match(s, r)
            if r2:
                unlabelled = not self.prior.labels().get((s, r)) and not self.current.labels().get((s, r2))
                return [((s, r2), 1.0 if how.startswith("same label") else
                         LAYOUT if how.startswith("same row") and unlabelled else 0.4, how)]
        lab = _norm(self.prior.labels().get((s, r), ""))
        if to and to != s and lab:  # the sheet was renamed: the label on the sheet it became, occurrence as before
            n = sum(1 for (sh, rr), l in self.prior.labels().items() if sh == s and rr <= r and _norm(l) == lab)
            there = sorted(rr for (sh, rr) in idx["by_label"].get(lab, []) if sh == to)
            if there:
                r2 = there[min(n, len(there)) - 1]
                return [((to, r2), 0.9, f"same label on sheet {to} (last year's {s}, renamed)")]
        hits = idx["by_label"].get(lab, []) if lab else []
        where = f"sheet {s} isn't in this year's model" if s not in idx["sheets"] else \
            f"this year's {s} isn't last year's (they share few labels)" if not to else f"not on {to}"
        if len(hits) == 1:
            return [(hits[0], 0.8, f"the only row labelled '{self.prior.labels().get((s, r))}' ({where})")]
        if 1 < len(hits) <= 5:
            return [(k, 0.5, f"one of {len(hits)} rows labelled '{self.prior.labels().get((s, r))}' ({where})") for k in hits]
        return []

    def _history(self, mine: dict, k: tuple) -> tuple | None:
        """One candidate's history against last year's row: (score, text, equal periods, differing history), or
        None when they share no period. A candidate blank in periods its sheet has dates for, where last year's
        row has values, scores less: an actuals sheet has the history exactly and no forecast, and would
        otherwise beat the row that carries both (a row with 9 of 51 periods can't outrank one with 51)."""
        theirs = self._series(self.current, *k)
        both = [w for w in mine if w in theirs and mine[w]]
        if not both:
            return None
        equal = [w for w in both if abs(theirs[w] - mine[w]) <= EXACT * max(1.0, abs(mine[w]))]
        later = [w for w in both if w not in equal]
        gaps = sorted(abs(theirs[w] / mine[w] - 1) for w in later)
        close = not gaps or gaps[len(gaps) // 2] < 0.25
        if not equal:
            return (0.0, "", 0, len(later))
        dates = self._index()["dates"].get(k[0], {})
        could = [w for w in mine if mine[w] and w in dates]
        cover = len(both) / len(could) if could else 1.0
        score = min(1.0, 0.5 * len(equal)) * (1.0 if close else 0.6) * min(1.0, cover / 0.8)
        return (score, f"{len(equal)} period(s) of history equal last year's"
                       + (f"; later years within {gaps[len(gaps) // 2]:.0%} (median)" if gaps else "")
                       + (f"; blank in {len(could) - len(both)} of the {len(could)} periods it has dates for"
                          if cover < 1 else ""), len(equal), len(later))

    def _by_history(self, s, r) -> list[tuple]:
        """Rows whose values equal last year's in the periods both models have (history), on any sheet."""
        mine = self._series(self.prior, s, r)
        if not mine:
            return []
        idx = self._index()
        exact = defaultdict(int)
        for sheet, dates in idx["dates"].items():
            for when, v in mine.items():
                if not v or when not in dates:
                    continue
                for r2 in self._values_at(sheet, when).get(float(f"{v:.9g}"), []):
                    exact[(sheet, r2)] += 1
        out = []
        for k in sorted(exact, key=lambda k: (-exact[k], k))[:40]:
            h = self._history(mine, k)
            if h and h[2] and (h[2] >= 2 or h[3] <= 20 * h[2]):
                out.append((k, h[0], h[1]))
        return out

    def _by_words(self, s, r) -> list[tuple]:
        """A label sharing most of its words, with the forecast staying close: on the sheet last year's sheet is
        this year, or anywhere when no sheet corresponds (then with more words in common)."""
        idx = self._index()
        mine_words = _words(self.prior.labels().get((s, r), ""))
        if not mine_words:
            return []
        to = self.sheet_for(s)
        floor = 0.4 if to else 0.5
        mine = self._series(self.prior, s, r)
        cands = {k for w in mine_words for k in idx["by_word"].get(w, ())}
        out = []
        for (s2, r2) in sorted(cands):
            if to and s2 != to:
                continue
            lab = idx["labels"].get((s2, r2), "")
            sim = _jaccard(mine_words, _words(lab))
            if sim < floor or _norm(lab) == _norm(self.prior.labels().get((s, r), "")):
                continue
            theirs = self._series(self.current, s2, r2)
            both = [w for w in mine if w in theirs and mine[w]]
            gaps = sorted(abs(theirs[w] / mine[w] - 1) for w in both)
            if both and gaps[len(gaps) // 2] < 0.1:
                out.append(((s2, r2), sim, f"label '{lab}' shares {sim:.0%} of its words; values within "
                                           f"{gaps[len(gaps) // 2]:.0%} (median) in {len(both)} period(s)"))
        return sorted(out, key=lambda x: (-x[1], x[0]))[:3]

    def _by_neighbours(self, s, r) -> list[tuple]:
        """Rows that read rows labelled as this one's inputs are and are read by rows labelled as its users."""
        idx = self._index()
        (p_reads, p_by), (c_reads, c_by) = idx["edges"]
        plab, clab = self.prior.labels(), idx["labels"]
        ins = {_norm(plab.get(x, "")) for x in p_reads.get((s, r), ())} - {""}
        outs = {_norm(plab.get(x, "")) for x in p_by.get((s, r), ())} - {""}
        if not ins and not outs:
            return []
        cands = set()
        for lab in ins:
            for x in idx["by_label"].get(lab, []):
                cands |= c_by.get(x, set())
        for lab in outs:
            for x in idx["by_label"].get(lab, []):
                cands |= c_reads.get(x, set())
        out = []
        for k in cands:
            ci = {_norm(clab.get(x, "")) for x in c_reads.get(k, ())} - {""}
            co = {_norm(clab.get(x, "")) for x in c_by.get(k, ())} - {""}
            parts = [j for j, a, b in ((_jaccard(ins, ci), ins, ci), (_jaccard(outs, co), outs, co)) if a]
            score = sum(parts) / len(parts) if parts else 0.0
            if score >= 0.34:
                out.append((k, score, f"reads {', '.join(sorted(ins & ci)) or 'nothing alike'}; read by "
                                      f"{', '.join(sorted(outs & co)) or 'nothing alike'}"))
        return sorted(out, key=lambda x: (-x[1], x[0]))[:3]

    def _by_banner(self, s, r) -> list[tuple]:
        """The model's banner (summary cells in its sheets' first rows) that read the row last year: the same
        banner in this year's model reads the row to use."""
        idx = self._index()
        (p_reads, p_by), (c_reads, c_by) = idx["edges"]
        plab = self.prior.labels()
        banners = [b for b in p_by.get((s, r), ()) if b[1] <= BANNER_ROWS and plab.get(b)]
        out = {}
        for b in banners:
            lab = _norm(plab[b])
            same = [x for x in idx["by_label"].get(lab, []) if x[1] <= BANNER_ROWS]
            for x in same:
                targets = [t for t in c_reads.get(x, ()) if t[1] > BANNER_ROWS]
                if len(targets) == 1:
                    k = targets[0]
                    seen = out.get(k, (k, 0.0, ""))
                    out[k] = (k, min(1.0, seen[1] + 0.7), f"the banner '{plab[b]}' ({x[0]}!r{x[1]}) reads it, as it read "
                                                          f"last year's ({b[0]}!r{b[1]})")
        return sorted(out.values(), key=lambda x: (-x[1], x[0]))

    # ---- deciding ----------------------------------------------------------------------------------------------
    # ---- the trace -------------------------------------------------------------------------------------------
    def _checked(self, a: tuple, c: tuple) -> bool:
        memo = self._index().setdefault("checked", {})
        if (a, c) not in memo:
            memo[(a, c)] = self.check(*a, c)["ok"]
        return memo[(a, c)]

    def _anchor(self, k: tuple) -> tuple | None:
        """This year's row for last year's row k by its label and its numbers alone, the trace's anchors: the same
        label where it was (or on the sheet it became), and last year's numbers within CHECK_GAP. None otherwise."""
        memo = self._index().setdefault("anchors", {})
        if k not in memo:
            memo[k] = None
            if self.prior.labels().get(k):
                hits = [x for x in self._by_label(*k) if x[1] >= 0.9]
                if len(hits) == 1 and self._checked(k, hits[0][0]) and not self.is_copy(*k, hits[0][0]):
                    memo[k] = hits[0][0]
        return memo[k]

    def _pair(self, a: tuple, c: tuple) -> float:
        """How surely this year's row c is last year's row a, by what they are: the same label 1, last year's
        numbers 0.8, most of the label's words 0.6; another kind of row, 0."""
        if self.other_shape(*a, c) or self.is_copy(*a, c):
            return 0.0
        la, lc = _norm(self.prior.labels().get(a, "")), _norm(self._index()["labels"].get(c, ""))
        if la and la == lc:
            return 1.0
        if self._checked(a, c):
            return 0.8
        return 0.6 if la and lc and _jaccard(_words(la), _words(lc)) >= 0.5 else 0.0

    def _step(self, x: tuple, sibs: set, cands: set, last: bool = True) -> tuple | None:
        """This year's row for last year's x among cands, the rows next to the path's previous row this year (the
        rows it reads, or is read by), sibs the same rows last year: the best paired, else (last) the one left once
        each of the others is paired. -> (row, quality, how) or None."""
        scored = sorted(((self._pair(x, c), c) for c in cands), key=lambda t: (-t[0], t[1]))
        if scored and scored[0][0] >= 0.6 and (len(scored) == 1 or scored[1][0] < scored[0][0]):
            q, c = scored[0]
            return c, q, "label" if q == 1.0 else "numbers" if q == 0.8 else "words"
        if not last:
            return None
        taken = set()
        for o in sorted(sibs - {x}):
            best = max(((self._pair(o, c), c) for c in cands - taken), default=(0.0, None))
            if best[0] < 0.8:
                return None
            taken.add(best[1])
        left = sorted(cands - taken)
        if len(left) == 1 and len(cands) == len(sibs):
            return left[0], 0.6, "the one left"
        return None

    def _by_trace(self, s, r) -> list[tuple]:
        """Last year's row reached from the anchors both models share, down from the rows that read it and up from
        the rows it reads, step by step through this year's model (_step). -> [(row, score, text)]."""
        (p_reads, p_by), (c_reads, c_by) = self._index()["edges"]
        key, plab = (s, r), self.prior.labels()
        found = defaultdict(list)
        for down, grow, sib_of, cand_of in ((True, p_by, p_reads, c_reads), (False, p_reads, p_by, c_by)):
            parent, frontier, anchors = {key: None}, [key], []
            for _ in range(TRACE_DEPTH):
                nxt = []
                for k in frontier:
                    for a in sorted(grow.get(k, ())):
                        if a in parent or len(parent) > TRACE_SCAN:
                            continue
                        parent[a] = k
                        nxt.append(a)
                        b = self._anchor(a)
                        if b and len(anchors) < TRACE_ANCHORS:
                            anchors.append((a, b))
                frontier = nxt
            for a, b in anchors:
                x, y, q, via = a, b, 1.0, []
                while x != key:
                    nx = parent[x]
                    st = self._step(nx, set(sib_of.get(x, ())), set(cand_of.get(y, ())))
                    if not st:
                        break
                    y, qq, how = st
                    q *= qq
                    via.append(f"'{plab.get(nx) or f'{nx[0]}!r{nx[1]}'}' ({how})")
                    x = nx
                else:
                    found[y].append((down, q, f"traced {'down' if down else 'up'} from '{plab.get(a)}' "
                                              f"({a[0]}!r{a[1]}; {b[0]}!r{b[1]} this year): " + " → ".join(via)))
        out = []
        for k, ways in found.items():
            if self.is_copy(s, r, k):
                continue
            best = {}
            for d, q, t in ways:  # the surest way each direction
                if d not in best or q > best[d][0]:
                    best[d] = (q, t)
            q = max(x[0] for x in best.values()) + 0.2 * (len(best) - 1) + 0.1 * (len(ways) - len(best))
            nums = self._checked(key, k)
            out.append((k, round(min(1.0, q + (0.2 if nums else 0.0)), 3),
                        "; ".join(t for _q, t in best.values()) + (f" (and {len(ways) - len(best)} more way(s))"
                                                                   if len(ways) > len(best) else "")
                        + ("; last year's numbers" if nums else "")))
        return sorted(out, key=lambda x: (-x[1], x[0]))[:3]

    def _below(self, k: tuple, depth: int = 3) -> set:
        """The rows of this year's model a row reads, and the rows they read, depth deep."""
        c_reads = self._index()["edges"][1][0]
        seen, frontier = set(), [k]
        for _ in range(depth):
            frontier = [x for y in frontier for x in c_reads.get(y, ()) if x not in seen]
            seen |= set(frontier)
        return seen

    def terms(self, s: str, r: int) -> tuple[list, list] | None:
        """How this year's row for last year's row (s, r) is made up against last year's: each row last year's read
        paired with one of this year's where it's found surely, else by its label, its numbers or its words (not as
        the one left: a term swapped for another is one gone and one new). -> (this year's rows it reads that last year's didn't, last year's rows
        it read that this year's doesn't), a term regrouped under another of the sum's (a subtotal of two of them)
        in neither. None where this year's row isn't settled, or where fewer than half of last year's terms pair
        (another make-up: an annual total of a quarterly row, a model rebuilt, not a term added or dropped)."""
        (p_reads, _), (c_reads, _) = self._index()["edges"]
        here = self.locate(s, r) if self.confident(s, r) else self._anchor((s, r))
        if not here:
            return None
        mine, theirs = set(p_reads.get((s, r), ())), set(c_reads.get(here, ()))
        sure = lambda x: self.locate(*x) if self.confident(*x) else None
        taken, left = set(), set()
        for x in sorted(mine):
            k = sure(x)
            if k in theirs and k not in taken:
                taken.add(k)
            else:
                left.add(x)
        for x in sorted(left):
            st = self._step(x, left, theirs - taken, last=False)
            if st:
                taken.add(st[0])
                left.discard(x)
        if not mine or 2 * (len(mine) - len(left)) < len(mine):
            return None
        under = self._below(here)
        moved = {sure(x) for x in left} - {None}
        new = [k for k in sorted(theirs - taken) if not (self._below(k) & moved)]
        gone = [x for x in sorted(left) if sure(x) not in under]
        return new, gone

    WEIGHTS = {"label": 0.3, "history": 0.35, "words": 0.2, "neighbours": 0.15, "banner": 0.25, "trace": 0.5,
               "shape": 0.0, "block": 0.3, "lineage": 0.2}

    def explain(self, s: str, r: int) -> dict:
        """{found: (sheet, row) or None, how, evidence: [(strategy, text)], confidence, alternatives}."""
        key = (s, r)
        if key in self._cache:
            return self._cache[key]
        if key in self.picks:
            k = self.picks[key]
            by = self.pick_by.get(key, "you")
            who = "the agents" if by == "agent" else "you"
            res = {"found": None if k == STAND_IN else k, "how": "the agents' pick" if by == "agent" else "your pick",
                   "confidence": 1.0, "alternatives": [], "in_place": False, "stand_in": k == STAND_IN, "by": by,
                   "evidence": [(who, "last year's values kept on purpose" if k == STAND_IN else f"picked by {who}")]}
            self._cache[key] = res
            return res
        found = defaultdict(dict)
        for name, fn in (("label", self._by_label), ("history", self._by_history), ("words", self._by_words),
                         ("neighbours", self._by_neighbours), ("banner", self._by_banner), ("trace", self._by_trace),
                         ("block", self._by_block)):
            try:
                for k, score, text in fn(s, r):
                    if k and (k not in found or name not in found[k] or found[k][name][0] < score):
                        found[k][name] = (score, text)
            except Exception:
                continue
        copies = [k for k in found if self.is_copy(s, r, k)]  # last year's figures pasted in: not a candidate at all
        for k in copies:
            found.pop(k)
        mine = self._series(self.prior, s, r)
        for k, ev in found.items():  # every candidate's own history, not only the closest lookalikes'
            if "history" not in ev and mine:
                h = self._history(mine, k)
                if h and h[2]:
                    ev["history"] = (h[0], h[1])
        for k, ev in list(found.items())[:LINEAGE_CANDS]:  # what each candidate is made of and feeds, two steps out
            lin = self._lineage(s, r, k)
            if lin and lin[0] >= 0.5:
                ev["lineage"] = lin
        ranked = []
        home = self.sheet_for(s)
        odd = set()
        for k, ev in found.items():
            total = sum(self.WEIGHTS[n] * sc for n, (sc, _) in ev.items())
            if home:  # on the sheet last year's sheet became, or not (the trace follows formulas onto a new sheet)
                total += 0.1 if k[0] == home else 0.0 if "trace" in ev else -0.1
            strong = any((n == "label" and sc >= 1.0) or (n == "history" and sc >= 0.5) or (n == "banner" and sc >= 0.7)
                         or (n == "neighbours" and sc >= 0.6) or (n == "label" and sc >= 0.8) or (n == "trace" and sc >= 0.7)
                         for n, (sc, _) in ev.items())
            # the same label where it was, and last year's numbers in the periods both have: nothing beats that
            sure = "label" in ev and ev["label"][0] >= 1.0 and "history" in ev
            why = self.other_shape(s, r, k) or self._kind_clash(s, r, k) or self.suspect_copy(s, r, k)
            if why:  # another kind of row: its evidence counts for less, and alone it isn't enough
                ev["shape"] = (0.0, why)
                total -= 0.2
                strong = sure = False
                odd.add(k)
            ranked.append((strong, total, k, ev, sure))
        ranked.sort(key=lambda x: (-x[4], -x[0], -x[1], x[2]))
        ranked = [x[:4] for x in ranked]
        res = {"found": None, "how": None, "evidence": [], "confidence": 0.0, "in_place": False,
               "copies": [f"{k[0]}!r{k[1]}" for k in copies],  # pasted copies of last year's figures, passed over
               "alternatives": [{"row": f"{k[0]}!r{k[1]}", "label": self.current.labels().get(k, ""),
                                 "score": round(t, 2), "evidence": [f"{n}: {tx}" for n, (_, tx) in ev.items()]}
                                for _, t, k, ev in ranked[:4]]}
        if ranked:
            strong, total, k, ev = ranked[0]
            # the history outranks a label that contradicts it: the same label, but other numbers in periods
            # both models have, where another row has last year's numbers
            hist = next((x for x in ranked if "history" in x[3] and x[3]["history"][0] >= 0.5 and x[2] not in odd), None)
            if hist and hist[2] != k and "history" not in ev and mine:
                h = self._history(mine, k)
                if h and h[3] and not h[2]:
                    strong, total, k, ev = hist
            if strong or total >= 0.3:
                # the same label in the same place, of the same kind: confident even where there's no history to
                # add (a single value, no timeline)
                res["in_place"] = k[0] == s and "label" in ev and k not in odd and (
                    ev["label"][0] >= 1.0 or (ev["label"][0] >= LAYOUT and k[1] == r))
                res.update(found=k, confidence=round(max(0.0, min(1.0, total)), 2),
                           scores={n: round(sc, 3) for n, (sc, _t) in ev.items()},
                           how=max(ev.items(), key=lambda kv: self.WEIGHTS[kv[0]] * kv[1][0])[0],
                           evidence=[(n, tx) for n, (_, tx) in sorted(ev.items(), key=lambda kv: -self.WEIGHTS[kv[0]] * kv[1][0])])
                res["alternatives"] = [a for a in res["alternatives"] if a["row"] != f"{k[0]}!r{k[1]}"]
                # last year's history, and every other candidate the same series in every period: which of them is
                # meant doesn't change a figure (a model carries one line on several sheets)
                others = [x[2] for x in ranked[1:4]]
                if res["confidence"] < CONFIDENT and "history" in ev and ev["history"][0] >= 0.5 and others \
                        and k not in odd and all(self._same_series(k, o) for o in others):
                    res["confidence"] = CONFIDENT
                    res["evidence"].append(("history", f"the other {len(others)} candidate(s) are the same series in every "
                                                       "period: whichever is meant, the figures are these"))
        # the trace, surely, to another row: the label or the numbers alone don't settle it (last year's numbers in
        # periods both have, under the same label, do)
        led = max(((ev["trace"][0], kk) for _s, _t, kk, ev in ranked if "trace" in ev), default=None)
        if res["found"] and led and led[0] >= 0.7 and led[1] != res["found"] and "trace" not in found[res["found"]] \
                and not ("label" in found[res["found"]] and "history" in found[res["found"]]):
            res["confidence"] = min(res["confidence"], round(CONFIDENT - 0.01, 2))
            res["in_place"] = False
            res["evidence"].append(("trace", f"but the trace leads to {led[1][0]}!r{led[1][1]}"))
        if res["found"]:  # copies of its block: which copy, by its heading (the trace may lead into the other copy)
            self._copies(s, r, res, found)
        if res["found"]:  # the kinds of evidence that agree on it: who it is, where it is, what it does, its numbers
            ev = found.get(res["found"]) or {}
            res["agreed"] = sorted({FAMILY[n] for n in ev if n in FAMILY} | ({"place"} if res["in_place"] else set()))
        if not (res["found"] and (res["confidence"] >= CONFIDENT or res["in_place"])) and self.blank(s, r):
            res.update(found=None, how="nothing to find", confidence=1.0, in_place=False, stand_in=True, blank=True,
                       by="code", evidence=[("code", "last year's row has no label, no numbers and no formulas: nothing "
                                                     "to find, so its blanks stand in")])
        self._cache[key] = res
        return res

    def _by_block(self, s, r) -> list[tuple]:
        """The row under last year's label in this year's block headed as last year's was (its heading's words);
        an unlabelled row at its place in a block of the same size."""
        import structure
        try:
            cur, pri = self._structure(self.current), self._structure(self.prior)
        except Exception:
            return []
        mine = structure.block_of(pri["blocks"], s, r)
        want = structure._words(mine["heading"]) if mine else set()
        if not want:
            return []
        lab = _norm(self.prior.labels().get((s, r), ""))
        clab = self._index()["labels"]
        i = mine["rows"].index(r)
        out = []
        for b in cur["blocks"]:
            if self.only is not None and b["sheet"] not in self.only:
                continue
            got = structure._words(b["heading"])
            j = len(want & got) / max(1, len(want | got))
            if j < 0.5:
                continue
            same = [x for x in b["rows"] if lab and _norm(clab.get((b["sheet"], x), "")) == lab]
            if len(same) == 1:
                out.append(((b["sheet"], same[0]), round(j, 2), f"under its label in the block headed '{b['heading']}', "
                                                                f"as last year's ('{mine['heading']}')"))
            elif not same and len(b["rows"]) == len(mine["rows"]):
                # its place in a block headed as last year's, of the same size: its label changed (or it has none)
                out.append(((b["sheet"], b["rows"][i]), round((0.5 if lab else 0.6) * j, 2),
                            f"at its place in the block headed '{b['heading']}', as last year's"
                            + (" (its label changed)" if lab else "")))
        return sorted(out, key=lambda x: (-x[1], x[0]))[:3]

    def _lineage(self, s, r, k) -> tuple | None:
        """How alike last year's row and this year's row k are by what they're made of and what they feed, two steps
        out each way (the rows' labels): (score, text), or None where last year's row reads and feeds nothing."""
        (p_reads, p_by), (c_reads, c_by) = self._index()["edges"]
        plab, clab = self.prior.labels(), self._index()["labels"]

        def out(graph, start, labels):
            one = set(graph.get(start, ()))
            two = {y for x in one for y in graph.get(x, ())} - {start}
            return {_norm(labels.get(x, "")) for x in one | two} - {""}
        sides = [(out(p_reads, (s, r), plab), out(c_reads, k, clab)), (out(p_by, (s, r), plab), out(c_by, k, clab))]
        parts = [_jaccard(a, b) for a, b in sides if a]
        if not parts:
            return None
        score = round(sum(parts) / len(parts), 3)
        (ua, ub), (da, db) = sides
        return score, (f"made of {', '.join(sorted(ua & ub)[:4]) or 'nothing alike'}; feeds "
                       f"{', '.join(sorted(da & db)[:4]) or 'nothing alike'} (two steps out)")

    def _kind_clash(self, s, r, k) -> str | None:
        """Another kind of row by its values: a row of dates, flags, factors, a share or an index isn't a cash flow."""
        try:
            a = self._structure(self.prior)["info"].get((s, r), {}).get("kind")
            b = self._structure(self.current)["info"].get(k, {}).get("kind")
        except Exception:
            return None
        if a and b and a != b and (a in FIXED_KINDS or b in FIXED_KINDS):
            return f"last year's row is {a}, this one {b}: another kind of row"
        return None

    # ---- a row's card: what it is, where it lives, what it's made of and feeds, roughly what it says ---------------
    def card(self, wb, k: tuple) -> dict:
        """A row's identity card in a workbook (this year's, for a pick): its label, its block's heading and its place
        in it, its kind, the labels around it, what it reads and what reads it, its first few figures, and the file."""
        import structure
        st = self._structure(wb)
        b = structure.block_of(st["blocks"], *k)
        labels = wb.labels()
        idx = self._index()
        reads, by = idx["edges"][1] if wb is self.current else idx["edges"][0]
        near = lambda d: [labels.get((k[0], k[1] + i), "") for i in d if labels.get((k[0], k[1] + i))]
        ser = self._series(wb, *k)
        return {"row": f"{k[0]}!r{k[1]}", "label": labels.get(k, ""), "heading": (b or {}).get("heading", ""),
                "position": b["rows"].index(k[1]) if b else None, "kind": st["info"].get(k, {}).get("kind"),
                "above": near((-1, -2)), "below": near((1, 2)),
                "reads": sorted({labels.get(x, "") for x in reads.get(k, ())} - {""})[:8],
                "fed": sorted({labels.get(x, "") for x in by.get(k, ())} - {""})[:8],
                "figures": [round(v, 6) for _w, v in sorted(ser.items())[:4]], "file": wb.path}

    def matches(self, card: dict, k: tuple) -> bool:
        """Whether this year's row k is the row a card describes: its label, heading and kind the same; and where its
        label is in more than one place this year (copies of a block), its place in its block and its figures too: a
        copy has the same label, heading words and kind."""
        now = self.card(self.current, k)
        same = lambda a, b: _norm(a or "") == _norm(b or "")
        if not (same(now["label"], card.get("label")) and same(now["heading"], card.get("heading"))
                and (not card.get("kind") or now["kind"] == card.get("kind"))):
            return False
        if len(self._index()["by_label"].get(_norm(card.get("label") or ""), [])) < 2:
            return True
        a, b = card.get("figures") or [], now["figures"]
        return (card.get("position") is None or now["position"] == card["position"]) and bool(a) and len(a) == len(b) \
            and all(abs(x - y) <= 1e-6 * max(1.0, abs(x)) for x, y in zip(a, b))

    def refind(self, card: dict) -> tuple[tuple | None, str]:
        """This year's row for a card, in a model that changed since the pick: the same label, heading and kind on the
        same sheet (the block moved), else the best match by label, heading, kind, the rows around it, what it reads
        and feeds, where two kinds of evidence agree and it's the only best one. -> (row or None, how)."""
        import structure
        st = self._structure(self.current)
        sheet = card["row"].rsplit("!r", 1)[0]
        lab = _norm(card.get("label") or "")
        same = lambda a, b: _norm(a or "") == _norm(b or "")
        here = [k for k in st["info"] if k[0] == sheet and lab and same(st["info"][k]["label"], card["label"])
                and self.matches(card, k)]
        if len(here) == 1:
            return here[0], f"moved to {here[0][0]}!r{here[0][1]} (the same label, heading and kind)"
        scored = []
        # the card's block is still in this model (its heading): a row under another heading is another case's, not it
        held = bool(card.get("heading")) and any(same(b["heading"], card["heading"]) for b in st["blocks"])
        for k, x in st["info"].items():
            if x["kind"] in ("empty", "text"):
                continue
            now = self.card(self.current, k)
            if held and not same(now["heading"], card["heading"]):
                continue
            fam = set()
            if lab and same(now["label"], card["label"]):
                fam.add("identity")
            if card.get("heading") and same(now["heading"], card["heading"]):
                fam.add("place")
            near = set(card.get("above", []) + card.get("below", []))
            if near and _jaccard({_norm(a) for a in near}, {_norm(a) for a in now["above"] + now["below"]}) >= 0.5:
                fam.add("place")
            lin = [_jaccard({_norm(a) for a in card.get(s_, [])}, {_norm(a) for a in now[s_]}) for s_ in ("reads", "fed")
                   if card.get(s_)]
            if lin and sum(lin) / len(lin) >= 0.5:
                fam.add("role")
            if card.get("kind") and now["kind"] != card["kind"] and (now["kind"] in FIXED_KINDS or card["kind"] in FIXED_KINDS):
                continue
            if len(fam) >= 2:
                scored.append((len(fam), k, fam))
        scored.sort(key=lambda x: (-x[0], x[1]))
        if scored and (len(scored) == 1 or scored[1][0] < scored[0][0]):
            n, k, fam = scored[0]
            return k, f"found again at {k[0]}!r{k[1]} by {', '.join(sorted(fam))}"
        return None, ("more than one row matches it as well" if scored else "no row of this model matches it")

    def _structure(self, wb):
        import structure
        return structure.load(wb.path)

    def _copies(self, s: str, r: int, res: dict, found: dict) -> None:
        """Last year's label in more than one place this year where it wasn't before (a downside or P90 case inserted
        above the base, the base renamed, no headings at all), or in copies of a block: the label, the neighbours, the
        trace and the lineage are the same in every copy, so none of them says which copy is meant. Two things can:
          the heading   the one whose heading names last year's case, surely and well ahead (structure.which_copy)
          the value     the one whose rows reach the row of the model's own valuation last year's row reached (the
                        equity value, the NPV), where exactly one does
        Else the row is left in doubt (copies_open), for the models or a person: never the first occurrence."""
        import structure
        k = res["found"]
        lab = _norm(self.prior.labels().get((s, r), ""))
        if not lab:
            return
        try:
            cur, pri = self._structure(self.current), self._structure(self.prior)
        except Exception:
            return
        cands = list(self._index()["by_label"].get(lab, []))
        if len(cands) < 2:
            return
        name = lambda b: f"{b['sheet']}!r{b['first']}:r{b['last']}"
        pairs = {frozenset((c["a"], c["b"])) for c in cur["copies"]}
        blocks = {x: structure.block_of(cur["blocks"], *x) for x in cands}
        paired = any(blocks[a] and blocks[b] and frozenset((name(blocks[a]), name(blocks[b]))) in pairs
                     for a in cands for b in cands if a != b)
        if not paired and len(cands) <= self._prior_count(lab):
            return  # as often as last year (and no copies of a block): the occurrences match in order
        res["in_place"] = False  # its place may be the copy's: a decider below, or doubt
        info = cur["info"]
        heading = lambda x: (blocks.get(x) or {}).get("heading") or info.get(x, {}).get("section") or x[0]
        mine = structure.block_of(pri["blocks"], s, r)
        want_bl = [{**mine, "heading": mine.get("heading") or pri["info"].get((s, r), {}).get("section") or s}] if mine \
            else [{"sheet": s, "rows": [r], "heading": pri["info"].get((s, r), {}).get("section") or s}]
        w = structure.which_copy(want_bl, cur["blocks"], cur["copies"], (s, r), cands, heading_of=heading)
        pick, why, how = w["pick"], w["why"], "block"
        if not pick:
            v = self._by_value(s, r, cands)
            if v:
                pick, why, how = v[0], v[1], "value"
        if pick:
            fam = "place" if how == "block" else "role"
            if pick == k:
                (found.get(k) or {})[how] = (1.0, why)
                res["evidence"] = [e for e in res["evidence"] if not (e[0] == "trace" and e[1].startswith("but the trace leads to")
                                                                       and any(e[1].endswith(f"{x[0]}!r{x[1]}") for x in cands))]
                res["evidence"].append((how, why))
                res["confidence"] = round(max(res["confidence"], CONFIDENT), 2)
            else:
                found.setdefault(pick, {})[how] = (1.0, why)
                found[pick].setdefault("label", (1.0, "the same label, in that copy"))
                res.update(found=pick, how=how, confidence=round(max(res["confidence"], CONFIDENT + 0.1), 2),
                           evidence=[(how, why), ("label", "the same label, in that copy")],
                           alternatives=[{"row": f"{k[0]}!r{k[1]}", "label": self.current.labels().get(k, ""), "score": 0.0,
                                          "evidence": ["found first, in another copy"]}] + res["alternatives"])
            res["copy_decided"] = fam
            return
        res["confidence"] = min(res["confidence"], round(CONFIDENT - 0.01, 2))
        res["copies_open"] = [f"{x[0]}!r{x[1]} '{heading(x)}'" for x in cands]
        res["evidence"].append(("block", f"the label is in {len(cands)} places this year (" + "; ".join(res["copies_open"][:4])
                                         + "): neither the headings nor the model's own value say which is last year's"))

    def _prior_count(self, lab: str) -> int:
        """How many of last year's rows have this label, on the sheets this year's are counted on."""
        if not hasattr(self, "_pcount"):
            self._pcount = defaultdict(int)
            for (s_, _r), l_ in self.prior.labels().items():
                if l_:
                    self._pcount[_norm(l_)] += 1
        return self._pcount.get(lab, 0)

    def _by_value(self, s: str, r: int, cands: list[tuple]) -> tuple | None:
        """The one candidate whose rows reach the same row of the model's own valuation (the equity value, the NPV:
        by its label) that last year's row reached in last year's model. -> (row, why) or None."""
        import structure
        (_p_reads, p_by), (_c_reads, c_by) = self._index()["edges"]
        plab, clab = self.prior.labels(), self._index()["labels"]
        try:
            pinfo, cinfo = self._structure(self.prior)["info"], self._structure(self.current)["info"]
        except Exception:
            return None

        def values(start, graph, labels, info):
            return {_norm(labels.get(x, "")) for x in structure.closure(start, graph) if x != start
                    and VALUE_ROW.search(labels.get(x, "") or "") and info.get(x, {}).get("kind") not in NOT_VALUE_KINDS}
        then = values((s, r), p_by, plab, pinfo) - {""}
        if not then:
            return None  # last year's row reached no valuation of the model's own: nothing to go by
        hits = [(x, then & values(x, c_by, clab, cinfo)) for x in cands]
        hits = [(x, v) for x, v in hits if v]
        if len(hits) != 1:
            return None
        x, v = hits[0]
        return x, f"the copy whose rows reach the model's own {sorted(v)[0]}, as last year's row did"

    def blank(self, s: str, r: int) -> bool:
        """Last year's row has nothing to find: no label, no number other than 0, no formulas."""
        if self.prior.labels().get((s, r)) or (self._shape(self.prior, (s, r)) or (0, 0))[0]:
            return False
        if s not in self._numeric:
            self._numeric[s] = {rr for (rr, _c), v in self.prior.sheet(s).items() if isinstance(v, float) and v}
        return r not in self._numeric[s]

    def family(self) -> float:
        """How alike the two models are: the share of line-item labels they have in common (a new version of a
        model shares most; a model rebuilt from the ground up, few)."""
        idx = self._index()
        if "family" not in idx:
            mine = {_norm(l) for l in self.prior.labels().values() if l}
            theirs = {_norm(l) for l in idx["labels"].values() if l}
            idx["family"] = round(_jaccard(mine, theirs), 3)
        return idx["family"]

    def locate(self, s: str, r: int) -> tuple | None:
        return self.explain(s, r)["found"]

    def why(self, s: str, r: int, labels: dict) -> str:
        if self.picks.get((s, r)) == STAND_IN:
            return "last year's values kept on purpose (your pick)"
        return self.rowmap.why(s, r, labels) + (" (and no other way found it: not by its history, words, neighbours, "
                                                 "banner or a trace from the rows both models share)")

    def confident(self, s: str, r: int) -> bool:
        """Settled well enough to roll on without a person looking: picked (a row, or last year's values kept; a
        person's pick or the agents'), nothing to find, a confidence of CONFIDENT or more, or found in place (the
        same label, or an unlabelled row at its own row number on a sheet laid out as before)."""
        if (s, r) in self.picks:
            return True
        ex = self.explain(s, r)
        return bool(ex.get("blank")) or (ex["found"] is not None and (ex["confidence"] >= CONFIDENT or ex["in_place"]))

    def pick(self, s: str, r: int, to, by: str = "you") -> None:
        """A choice for a row: (sheet, row), STAND_IN (keep last year's values), or None (back to what's found);
        by: "you" (a person) or "agent". The agents never replace a person's choice."""
        if by == "agent" and self.pick_by.get((s, r), "you" if (s, r) in self.picks else None) == "you":
            return
        if to:
            self.picks[(s, r)] = to
            self.pick_by[(s, r)] = by
        else:
            self.picks.pop((s, r), None)
            self.pick_by.pop((s, r), None)
        self._cache.pop((s, r), None)

    # ---- checking a match by its numbers, and finding rows by them -------------------------------------------
    def check(self, s: str, r: int, k: tuple) -> dict:
        """Whether this year's row k carries last year's row's numbers: over the periods both have (no roll), the
        median gap between them, how many are equal (history), and whether it's the same kind of row. Forecasts
        are revised between valuations, not replaced: ok when the median gap is within CHECK_GAP over at least
        CHECK_PERIODS periods, the kind agrees, and the row isn't blank where last year's has values."""
        mine = self._series(self.prior, s, r)
        theirs = self._series(self.current, *k)
        both = [w for w in mine if w in theirs and mine[w]]
        dates = self._index()["dates"].get(k[0], {})
        could = [w for w in mine if mine[w] and w in dates]
        gaps = sorted(abs(theirs[w] / mine[w] - 1) for w in both)
        med = gaps[len(gaps) // 2] if gaps else None
        equal = sum(1 for g in gaps if g <= 1e-9)
        shape = self.other_shape(s, r, k)
        cover = len(both) / len(could) if could else 0.0
        ok = len(both) >= CHECK_PERIODS and med is not None and med <= CHECK_GAP and not shape and cover >= 0.8
        return {"row": f"{k[0]}!r{k[1]}", "periods": len(both), "median_gap": None if med is None else round(med, 4),
                "equal": equal, "cover": round(cover, 3), "shape": shape, "ok": ok,
                "text": (f"{len(both)} period(s) side by side, median difference {med:.1%}" if med is not None else
                         "no periods side by side") + (f", {equal} equal" if equal else "")
                        + (f"; {shape}" if shape else "") + (f"; blank in {len(could) - len(both)} of {len(could)}"
                                                              if could and cover < 1 else "")}

    def _sorted_values(self, sheet: str, col: int) -> tuple[list, list]:
        idx = self._index()
        key = ("sorted", sheet, col)
        if key not in idx:
            vals = sorted((v, rr) for (rr, cc), v in self.current.sheet(sheet).items()
                          if cc == col and isinstance(v, float) and v)
            idx[key] = ([v for v, _ in vals], [rr for _, rr in vals])
        return idx[key]

    def near(self, s: str, r: int, limit: int = 8, tol: float = CHECK_GAP) -> list[tuple]:
        """Rows of this year's model, on any sheet, whose values are close to last year's row's (within tol) in
        several of the same periods: a row found by its numbers when its label says nothing (or there's none)."""
        mine = [(w, v) for w, v in self._series(self.prior, s, r).items() if v]
        if len(mine) < CHECK_PERIODS:
            return []
        step = max(1, len(mine) // 8)
        sample = mine[::step][:8]
        hits = defaultdict(int)
        for sheet, dates in self._index()["dates"].items():
            for w, v in sample:
                c = dates.get(w)
                if c is None:
                    continue
                vals, rows = self._sorted_values(sheet, c)
                lo, hi = sorted((v * (1 - tol), v * (1 + tol)))
                for i in range(bisect.bisect_left(vals, lo), bisect.bisect_right(vals, hi)):
                    hits[(sheet, rows[i])] += 1
        need = max(CHECK_PERIODS, len(sample) // 2)
        return sorted((k for k, n in hits.items() if n >= need), key=lambda k: (-hits[k], k))[:limit]
