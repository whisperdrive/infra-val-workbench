"""The engagement's workpaper, as an Excel file: what the workbench found, in the order a reviewer reads it.

  Summary          the equity value low / mid / high (the report's, rebuilt in Python, this year's, the move), the
                   model inputs, the valuation dates, the review in brief
  Bridge           last year to this year, each end, and a waterfall of the mid
  Cash flows       last year's model against this year's, by financial year
  Inputs           the discount rate, the terminal growth rate and franking: the cell each was sourced from and every
                   check on it; the discountings under the value
  Methods          this year's value worked out other ways (methods.py), against the default; the preferred one
  Reconciliation   what the report discloses of the value against the overlay's own split
  Key facts        the report's, with who decided each and how it was checked
  Files and roles  the four files and who placed each
  Review           the reviewer's points
  Run log          what ran, what it decided and why

Figures are values, as the workbench worked them out: nothing links back to the engagement's files. The figure
leads, the cell it came from follows. Dates read as 30 June 2025. Built in memory; nothing is written to disk.
"""
import io
import re
import time
from datetime import date

import xlsxwriter

import orchestrator
import workbench as wb

MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
          "November", "December"]
ISO = re.compile(r"\b(\d{4})-(\d{2})-(\d{2})\b")
DISCLAIMER = ("Prepared by software from the engagement's own files. Every figure here was read, rebuilt or rolled "
              "forward by the workbench, and is for the engagement team to check before relying on it. It is not a "
              "valuation report.")
STEPS = ("report", "rounding", "prior_feed", "rebuilt", "roll", "time", "cash", "forecast", "held", "rate", "this_year")
STEP_SHORT = {"report": "Report", "rounding": "Rounding", "prior_feed": "Client file", "rebuilt": "Rebuilt",
              "time": "Time value", "cash": "Cash flows paid", "forecast": "New forecast", "roll": "Roll-forward",
              "held": "Held inputs", "rate": "Discount rate", "this_year": "This year"}
ROLES = {"prior_report": "Last year's report", "prior_overlay": "Last year's overlay",
         "prior_model": "Last year's client model", "current_model": "This year's client model"}
INPUTS = {"rate": "Discount rate", "growth": "Terminal growth rate", "multiple": "Exit multiple",
          "franking": "Franking credit utilisation"}
SPLIT = (("pv", "Value of the discounted cash flows"), ("pv_forecast", "PV of the discrete forecast"),
         ("pv_tv", "PV of the terminal value"), ("tv", "Terminal value"), ("franking", "Value of franking credits"),
         ("franking_share", "Franking credits, % of the equity value"),
         ("tv_share", "Terminal value, % of the discounted cash flows' value"))
ACCENT, FILL = "#1a9afa", "#1478d0"


def nice(s):
    """ISO dates inside any text as 30 June 2025 (the page does the same)."""
    if not isinstance(s, str):
        return s
    return ISO.sub(lambda m: f"{m[3]} {MONTHS[int(m[2]) - 1]} {m[1]}"
                   if 1 <= int(m[2]) <= 12 and 1 <= int(m[3]) <= 31 else m[0], s)


def when(t: float | None) -> str:
    if not t:
        return ""
    lt = time.localtime(t)
    return f"{lt.tm_mday:02d} {MONTHS[lt.tm_mon - 1]} {lt.tm_year} {lt.tm_hour:02d}:{lt.tm_min:02d}"


def mark(ok) -> str:
    return "✓" if ok is True else "✗" if ok is False else "–"


def _num(v):
    return (0.0 if abs(v) < 1e-9 else float(v)) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


class _Sheet:
    """A worksheet written top to bottom: a row cursor, the workpaper's formats, text passed through nice()."""

    def __init__(self, book, fmt, name: str, widths: list[float]):
        self.ws, self.f, self.r = book.add_worksheet(name), fmt, 0
        for i, w in enumerate(widths):
            self.ws.set_column(i, i, w)
        self.ws.hide_gridlines(2)

    def cell(self, r: int, c: int, v, f=None):
        n = _num(v)
        if n is not None:
            self.ws.write_number(r, c, n, f)
        elif v is None or v == "":
            self.ws.write_blank(r, c, None, f)
        else:
            self.ws.write_string(r, c, nice(str(v)), f)

    def line(self, *vals, f=None, fs=None, height=None):
        for c, v in enumerate(vals):
            self.cell(self.r, c, v, (fs[c] if fs and c < len(fs) else None) or f)
        if height:
            self.ws.set_row(self.r, height)
        self.r += 1
        return self.r - 1

    def text(self, s: str, f=None, span: int = 6, height=None):
        if span > 1:
            self.ws.merge_range(self.r, 0, self.r, span - 1, nice(s or ""), f or self.f["wrap"])
        else:
            self.cell(self.r, 0, s, f)
        if height:
            self.ws.set_row(self.r, height)
        self.r += 1

    def title(self, s: str, sub: str | None = None, span: int = 6):
        self.line(s, f=self.f["title"])
        if sub:
            self.text(sub, self.f["sub"], span, height=30 if len(sub) > 110 else None)
        self.r += 1

    def head(self, s: str):
        self.line(s, f=self.f["h"])

    def header(self, *cols):
        return self.line(*cols, fs=[self.f["th"]] + [self.f["thn"] if c in ("Low", "Mid", "High") or
                                                        c.startswith(("Last year", "This year", "Change", "Python", "Value", "PV"))
                                                        else self.f["th"] for c in cols[1:]])

    def gap(self, n: int = 1):
        self.r += n


def _formats(book) -> dict:
    base = {"font_name": "Calibri", "font_size": 10, "valign": "top"}
    F = lambda **k: book.add_format({**base, **k})
    num, chg, pct = "#,##0.0", '+#,##0.0;-#,##0.0;"–"', "0.00%"
    return {"title": F(bold=True, font_size=16, font_color=FILL), "sub": F(italic=True, font_color="#555555", text_wrap=True),
            "h": F(bold=True, font_size=12, font_color=FILL), "th": F(bold=True, font_color="#ffffff", bg_color=FILL, text_wrap=True),
            "thn": F(bold=True, font_color="#ffffff", bg_color=FILL, text_wrap=True, align="right"),
            "wrap": F(text_wrap=True), "b": F(bold=True), "bwrap": F(bold=True, text_wrap=True),
            "muted": F(font_color="#666666", text_wrap=True), "mono": F(font_name="Consolas", font_color="#444444"),
            "num": F(num_format=num), "numb": F(num_format=num, bold=True), "chg": F(num_format=chg),
            "tot": F(num_format=num, bold=True, top=1), "totl": F(bold=True, top=1, text_wrap=True),
            "pct": F(num_format=pct), "pct1": F(num_format="0.0%"), "mult": F(num_format='0.00"x"'), "ok": F(font_color="#1e7b3c", bold=True, align="center"),
            "bad": F(font_color="#b3261e", bold=True, align="center"), "unk": F(font_color="#888888", align="center"),
            "disc": F(italic=True, text_wrap=True, border=1, border_color="#bbbbbb", bg_color="#f4f8fc"),
            "warn": F(text_wrap=True, bold=True, font_color="#8a5a00", bg_color="#fff4d6")}


def _okf(f: dict, ok):
    return f["ok"] if ok is True else f["bad"] if ok is False else f["unk"]


def _units(res: dict) -> str:
    return (res.get("bridges") or {}).get("units") or (res.get("head") or {}).get("units") or ""


def _basis(res: dict) -> str:
    return "cum-distribution" if (res.get("head") or {}).get("basis") == "cum" else "ex-distribution"


def _rates(rt: dict) -> str:
    """This year's discount rate as a range, the lower first: 7.90% to 8.90%."""
    import result
    return result._range_pct(rt.get("low"), rt.get("high"))


def _totals(steps: list[dict]) -> tuple:
    t = [s["value"] for s in steps if s.get("total") and s.get("value") is not None]
    return (t[0], t[-1]) if len(t) > 1 else (None, None)


# ---- the sheets ---------------------------------------------------------------------------------------------------

def _summary(S: _Sheet, g: dict, res: dict, review: dict):
    f, V, B = S.f, res.get("values") or {}, res.get("bridges") or {}
    S.ws.set_row(0, 24)
    S.line("Valuation workpaper", f=f["title"])
    S.line(g.get("name") or f"Engagement {g['id']}", f=f["b"])
    S.line(f"Prepared {date.today().isoformat()} by Infra Val Workbench", f=f["muted"])
    S.text(DISCLAIMER, f["disc"], 5, height=42)
    S.gap()
    S.head("The equity value")
    S.text(f"{_units(res)} · {_basis(res)} · the mid is the midpoint of the low and the high", f["muted"], 5)
    S.header("", "Low", "Mid", "High")
    row = lambda label, d, fl=None, fn=None: S.line(label, *[(d or {}).get(e) for e in ("low", "mid", "high")],
                                                   fs=[fl or f["wrap"]] + [fn or f["num"]] * 3)
    row("Last year, per the report", V.get("report"))
    tie = res.get("tie") or {}
    S.line("  the report as Excel saved it ties", *[mark((tie.get(e) or {}).get("ok")) if e != "mid" else "" for e in ("low", "mid", "high")],
           fs=[f["muted"]] + [_okf(f, (tie.get(e) or {}).get("ok")) for e in ("low", "mid", "high")])
    row("Last year, rebuilt in Python", V.get("rebuilt"))
    S.line("  the rebuild ties to the report", *[mark((tie.get(e) or {}).get("rebuilt_ok")) if e != "mid" else "" for e in ("low", "mid", "high")],
           fs=[f["muted"]] + [_okf(f, (tie.get(e) or {}).get("rebuilt_ok")) for e in ("low", "mid", "high")])
    to = f" to {B['valuation_date'][:10]}" if B.get("valuation_date") else ""
    if V.get("this_year"):
        row(f"This year, rolled forward{to}", V["this_year"], f["totl"], f["tot"])
        move = {e: (V["this_year"].get(e) - V["report"].get(e)) if V["this_year"].get(e) is not None
                and (V.get("report") or {}).get(e) is not None else None for e in ("low", "mid", "high")}
        row("The move: this year less last year per the report", move, None, f["chg"])
    else:
        S.line("This year", "held back: rows of this year's model still to be found (see the Result page)" if B.get("held")
               else "not rolled forward yet", fs=[f["totl"], f["muted"]])
    S.gap()

    I = res.get("inputs") or {}
    if I:
        S.head("Model inputs: sourced and checked")
        S.header("", "Low", "", "High", "Checked")
        for key, name in INPUTS.items():
            X = I.get(key) or {}
            if not X:
                continue
            ends = X.get("ends") or {}
            S.line(name, (ends.get("low") or {}).get("value"), "", (ends.get("high") or {}).get("value"),
                   "not applicable" if X.get("na") else "couldn't check" if X.get("error") else "✓ checked"
                   if X.get("ok") is True else "to look at",
                   fs=[f["wrap"], f["mult" if key == "multiple" else "pct"], None, f["mult" if key == "multiple" else "pct"],
                       _okf(f, None if X.get("error") or X.get("na") else X.get("ok") is True or False)])
        S.text("The low is at the higher discount rate. Each input's cell and every check on it: the Inputs sheet.", f["muted"], 5)
        S.gap()
    T = res.get("terminal") or {}
    if T:
        S.head("The terminal value")
        S.text(T.get("label") + (f": “{T['phrase']}” (p. {T.get('page')})" if T.get("phrase") else ""), f["wrap"], 5,
               height=30 if T.get("phrase") else None)
        S.gap()
    H = res.get("held") or []
    if H:
        still = [h for h in H if h["held"]]
        S.head("Inputs held at last year's")
        S.text(f"{len(still)} of {len(H)} still at last year's figure: " + ", ".join(h["label"] for h in still) if still else
               f"All {len(H)} set for this year: " + ", ".join(f"{h['label']} {h['this_year']:,.1f}" for h in H),
               f["warn"] if still else f["wrap"], 5)
        S.text("Typed into the overlay outside its discountings; the Inputs sheet has each, with this year's model's figure.",
               f["muted"], 5)
        S.gap()

    S.head("The valuation dates")
    S.header("Step", "Date", "Check", "Where", "How", "Note")
    for d in g.get("dates") or []:
        S.line(d["step"], d.get("date") or "–", mark(d.get("ok")), d.get("where") or "", d.get("how") or "", d.get("note") or "",
               fs=[f["wrap"], None, _okf(f, d.get("ok")), f["mono"], f["wrap"], f["muted"]])
    S.gap()

    S.head("The review")
    pts = review["points"]
    checks = [p for p in pts if p.get("severity") != "info"]
    S.text(f"{len(checks)} point(s) to check and {len(pts) - len(checks)} note(s): the Review sheet." if pts else
           (review["said"] or "The reviewer raised no concerns.") if review["ran"] else "Not reviewed yet.", f["wrap"], 5)
    for p in pts[:6]:
        S.text(f"• {p.get('title', '')}", f["wrap"], 5)


def _bridge(S: _Sheet, book, res: dict):
    f, B = S.f, res.get("bridges") or {}
    to = f" to {B['valuation_date'][:10]}" if B.get("valuation_date") else ""
    rt = ((res.get("inputs") or {}).get("rate") or {}).get("this_year") or {}
    at = f"at this year's discount rate, {_rates(rt)}" if rt.get("applied") else "at last year's discount rate"
    S.title("The value bridge: last year to this year",
            f"{_units(res)} · equity value, {_basis(res)} · {at}{to}. The mid is the midpoint of "
            f"the low and the high, step by step. A total is the value at that point; the other rows are the change.", 5)
    by = {e: {s["key"]: s for s in (B.get(e) or {}).get("steps") or []} for e in ("low", "mid", "high")}
    keys = [k for k in STEPS if any(k in by[e] for e in by)]
    keys += [k for e in by for k in by[e] if k not in keys]
    top = S.header("Step", "Low", "Mid", "High")
    helper = []  # the mid's waterfall: base (hidden), up, down, total
    run = None
    for k in keys:
        s = by["mid"].get(k) or by["low"].get(k) or by["high"].get(k)
        tot = bool(s.get("total"))
        vals = [(by[e].get(k) or {}).get("value") for e in ("low", "mid", "high")]
        r = S.line(s.get("label") or k, *vals, fs=[f["totl"] if tot else f["wrap"]] + [f["tot"] if tot else f["chg"]] * 3)
        v = vals[1]
        if v is None:
            continue
        if tot:
            helper.append((r, STEP_SHORT.get(k, k), 0.0, 0.0, 0.0, v))
            run = v
        elif run is not None:
            after = run + v
            helper.append((r, STEP_SHORT.get(k, k), min(run, after), max(v, 0.0), max(-v, 0.0), 0.0))
            run = after
    if B.get("held"):
        S.text("This year's value is held back until the rows of this year's model it reads are found: the bridge stops "
               "at last year's rebuilt value.", f["muted"], 4)
    for n in B.get("notes") or []:
        S.text(n, f["muted"], 4)
    if not helper:
        return
    # the chart's helper block, to the right: labelled, so a reader knows what it's for
    S.ws.write_string(top - 1, 6, "Chart helper: the mid's waterfall", f["muted"])
    for c, h in enumerate(("Step", "Base", "Up", "Down", "Total")):
        S.ws.write_string(top, 6 + c, h, f["th"])
    for i, (_, name, base, up, down, total) in enumerate(helper):
        S.ws.write_string(top + 1 + i, 6, name)
        for c, v in enumerate((base, up, down, total)):
            S.ws.write_number(top + 1 + i, 7 + c, v, f["num"])
    first, last = top + 1, top + len(helper)
    ch = book.add_chart({"type": "column", "subtype": "stacked"})
    for c, name, fill in ((7, "Base", None), (8, "Increase", "#2e9e5b"), (9, "Decrease", "#c0392b"), (10, "Total", ACCENT)):
        ch.add_series({"name": name, "categories": ["Bridge", first, 6, last, 6], "values": ["Bridge", first, c, last, c],
                       "gap": 60, **({"fill": {"color": fill}, "border": {"none": True}} if fill else
                                     {"fill": {"none": True}, "border": {"none": True}})})
    ch.set_title({"name": f"The mid: last year to this year ({_units(res)})", "name_font": {"size": 11}})
    ch.set_legend({"none": True})
    ch.set_y_axis({"num_format": "#,##0", "major_gridlines": {"visible": True, "line": {"color": "#e5e5e5"}}})
    ch.set_x_axis({"num_font": {"size": 8}})
    ch.set_size({"width": 720, "height": 360})
    S.ws.insert_chart(S.r + 1, 0, ch)


def _flows(S: _Sheet, book, res: dict):
    f, C = S.f, res.get("chart") or {}
    left = f" ({', '.join(C['left_out'])})" if C.get("left_out") else ""
    S.title("The forecast cash flows",
            f"{_units(res)} · undiscounted, by financial year, the terminal value left out{left}: last year's client "
            f"model against this year's.", 4)
    if not C.get("years"):
        S.text(C.get("why") or "No cash flows found under the value.", f["muted"], 4)
        return
    if C.get("label"):
        S.line("The line", C["label"], fs=[f["b"], f["wrap"]])
    if C.get("rows"):
        S.line("Read from", ", ".join(C["rows"]) + (f" (under {C['core']})" if C.get("core") else ""), fs=[f["b"], f["mono"]])
    S.gap()
    top = S.header("Financial year", "Last year's model", "This year's model", "Change")
    ly, ty = (C.get("series") or {}).get("last_year") or {}, (C.get("series") or {}).get("this_year") or {}
    for y in C["years"]:
        a, b = _num(ly.get(y)), _num(ty.get(y))
        S.line(y, a, b, b - a if a is not None and b is not None else None, fs=[None, f["num"], f["num"], f["chg"]])
    first, last = top + 1, top + len(C["years"])
    ch = book.add_chart({"type": "column"})
    for c, name, fill in ((1, "Last year's model", "#9aa7b4"), (2, "This year's model", ACCENT)):
        ch.add_series({"name": name, "categories": ["Cash flows", first, 0, last, 0],
                       "values": ["Cash flows", first, c, last, c], "fill": {"color": fill}, "border": {"none": True}, "gap": 80})
    ch.set_title({"name": f"Cash flows by financial year ({_units(res)})", "name_font": {"size": 11}})
    ch.set_legend({"position": "bottom"})
    ch.set_y_axis({"num_format": "#,##0", "major_gridlines": {"visible": True, "line": {"color": "#e5e5e5"}}})
    ch.set_x_axis({"num_font": {"size": 8, "rotation": -45}})
    ch.set_size({"width": 760, "height": 340})
    S.ws.insert_chart(top, 5, ch)


def _inputs(S: _Sheet, res: dict):
    f, I = S.f, res.get("inputs") or {}
    S.title("Model inputs: sourced and checked",
            "Each input is the value in the cell the model's formulas read, found by following them from the discount "
            "factors; recomputed from it; then checked against the report.", 9)
    for key, name in INPUTS.items():
        X = I.get(key)
        if not X:
            continue
        state = "not applicable" if X.get("na") else "couldn't check" if X.get("error") else "checked" \
            if X.get("ok") is True else "to look at"
        S.head(f"{name}: {state}")
        if X.get("report"):
            S.line("The report", " – ".join(X["report"]), fs=[f["b"], f["wrap"]])
        if X.get("error") or X.get("na"):
            S.text(X.get("na") or X["error"], f["muted"], 9)
            S.gap()
            continue
        S.header("End", "Value", "Ties", "Cell", "Label", "Heading", "Input or formula", "Read as", "Checks")
        for e in ("low", "high"):
            r = (X.get("ends") or {}).get(e) or {}
            end = {"low": "Low (at the higher rate)", "high": "High (at the lower rate)"}[e] if key == "rate" else e.title()
            if not r.get("checks"):
                S.line(end, r.get("why") or "not found", fs=[f["b"], f["muted"]])
                continue
            s = r.get("sighted") or {}
            checks = "\n".join(f"{mark(c.get('ok'))} {nice(c.get('text', ''))}" for c in r["checks"])
            S.line(end, r.get("value"), mark(r.get("ties")), r.get("cell") or "not sourced", s.get("label") or r.get("note") or "",
                   s.get("heading") or "", ("an input, typed in" if s.get("input") else f"a formula: {s.get('formula') or ''}") if s else "",
                   " → ".join(s.get("chain") or []), checks,
                   fs=[f["b"], f["mult" if key == "multiple" else "pct"], _okf(f, r.get("ties")), f["mono"], f["wrap"],
                       f["wrap"], f["wrap"], f["mono"], f["wrap"]],
                   height=max(15, 13 * (len(r["checks"]) + checks.count("\n") // 3 + 1)))
        rt = X.get("this_year") if key == "rate" else None
        if rt:
            S.line("This year", _rates(rt), "set by you" + ("" if rt.get("applied") else f"; not applied: {rt.get('why') or ''}"),
                   fs=[f["b"], f["wrap"], f["wrap"] if rt.get("applied") else f["warn"]])
        S.gap()

    H = res.get("held") or []
    if H:
        S.head("Inputs held at last year's")
        S.text("Typed into the overlay outside its discountings, these stay at last year's figures in the roll-forward "
               "until a person sets this year's. The suggestion is from this year's client model, found through the row "
               "that holds last year's figure in last year's model. Figures as typed in the overlay.", f["muted"], 9)
        S.header("Input", "Last year", "This year", "Status", "This year's model", "How found", "Cell", "Feeds")
        for h in H:
            sg = h.get("suggestion") or {}
            status = "held at last year's" if h["held"] else \
                f"this year's, {'accepted from the model' if h.get('from') == 'suggestion' else 'typed'} by a person"
            S.line(h["label"], h["value"], h.get("this_year") if not h["held"] else None, status, sg.get("value"),
                   ("checked: " if sg.get("status") == "checked" else "") + (sg.get("text") or ""), h["cell"],
                   ", ".join(dict.fromkeys(x["label"] for x in h.get("lines") or [])),
                   fs=[f["bwrap"], f["num"], f["num"], f["warn"] if h["held"] else f["wrap"], f["num"], f["wrap"], f["mono"],
                       f["wrap"]], height=45)
        S.gap()

    A = res.get("assumptions") or {}
    if any(A.get(e) for e in ("low", "high")):
        S.head("The discountings under the value")
        S.header("End", "PV", "Rate", "Valuation date", "Discounting", "Timing", "Day count", "Periods", "Terminal date", "Cell")
        for e in ("low", "high"):
            for a in A.get(e) or []:
                if a.get("error"):
                    S.line(e.title(), None, None, "", a.get("label") or "", a["error"], fs=[f["b"], None, None, None, f["wrap"], f["muted"]])
                    continue
                S.line(e.title(), a.get("pv"), a.get("rate"), a.get("valuation_date") or "", a.get("label") or a.get("kind") or "",
                       a.get("timing") or "", a.get("day_count") or "", a.get("periods"), a.get("terminal_date") or "", a.get("cell") or "",
                       fs=[f["b"], f["num"], f["pct"], None, f["wrap"], f["wrap"], None, None, None, f["mono"]])


def _methods(S: _Sheet, res: dict):
    f, M = S.f, res.get("methods") or {}
    S.title("Methods: this year's value worked out other ways",
            f"{_units(res)}. On the same rolled-forward model and discount rates. The default is the overlay's own formulas "
            f"with the periods ending on or before the new valuation date cut off; the preferred method is this year's "
            f"value, with a bridge step of its own for the move to it.", 7)
    if M.get("error") or not M.get("methods"):
        S.text(M.get("error") or "Worked out once this year's value is.", f["muted"], 7)
        return
    S.header("Method", "Low", "Mid", "High", "Mid vs the default", "", "How")
    for m in M["methods"]:
        tag = " (this year's value)" if m["key"] == M.get("preferred") else " (the default)" if m["key"] == M.get("default") else ""
        S.line(m["label"] + tag, m["low"], m["mid"], m["high"], (m.get("vs_default") or {}).get("mid"), "",
               m["what"] if m["ok"] else f"not worked out: {m.get('why') or ''}",
               fs=[f["bwrap"] if tag else f["wrap"], f["num"], f["num"], f["num"], f["chg"], None, f["wrap" if m["ok"] else "muted"]])
    S.gap()
    S.text(("The recompute ties to the overlay's own formulas." if M.get("ties") else
            "The recompute doesn't tie to the overlay's own formulas: read the recomputed methods as indicative.")
           + (f" The overlay's own forecast flags ({', '.join(M['flags'])}) were rolled as a method of their own."
              if M.get("flags") else ""), f["muted"], 7)


def _reconcile(S: _Sheet, res: dict):
    f, C = S.f, res.get("reconcile") or {}
    S.title("Reconciliation to the report",
            f"{_units(res)}. What the report discloses of the value, against the same split of the overlay's own "
            f"discountings as Excel saved them. The mid is the midpoint of the low and the high; a share at the mid is "
            f"the mid's figure over the mid's total. Ties within half a unit of the report's last digit.", 8)
    if C.get("error"):
        S.text(C["error"], f["muted"], 8)
        return
    S.header("Line item", "The report", "Python, low", "Mid", "High", "This year, mid", "Ties", "Report page")
    for r in C.get("rows") or []:
        pct = r.get("unit") == "%"
        v = lambda x: x / 100 if pct and isinstance(x, (int, float)) else x
        rep = " · ".join(f"{'' if e == 'mid' else e + ' '}{r['report'][e]}" for e in ("low", "mid", "high")
                         if (r.get("report") or {}).get(e)) or "not disclosed"
        p, t = r.get("python") or {}, r.get("ties") or {}
        ties = " ".join(f"{e} {mark(t[e])}" for e in ("low", "mid", "high") if e in t) or "–"
        n = f["pct1"] if pct else f["num"]
        S.line(r["label"], rep, v(p.get("low")), v(p.get("mid")), v(p.get("high")), v((r.get("this_year") or {}).get("mid")),
               ties, r.get("page"), fs=[f["bwrap"], f["wrap"], n, n, n, n, _okf(f, r.get("ok")), None])
    L, T = (C.get("last_year") or {}).get("mid") or {}, (C.get("this_year") or {}).get("mid") or {}
    if L:
        S.gap()
        S.head("The split at the mid")
        S.header("", "Last year", "This year")
        for k, label in SPLIT:
            pct = k.endswith("share")
            v = lambda x: x / 100 if pct and isinstance(x, (int, float)) else x
            S.line(label, v(L.get(k)), v(T.get(k)), fs=[f["wrap"]] + [f["pct1"] if pct else f["num"]] * 2)
    st = (C.get("streams") or {}).get("low") or {}
    if st.get("main"):
        S.gap()
        S.text(f"The largest discounted stream ({st['main']}) is split into its terminal value and the discrete forecast; "
               f"franking credits are found by their label.", f["muted"], 8)


def _midpoint(lo: str, hi: str) -> str | None:
    """The midpoint of a range's two ends, printed as the low end is (its prefix, unit and decimals)."""
    a, b = (re.search(r"-?\d[\d,]*(?:\.\d+)?", t or "") for t in (lo, hi))
    if not a or not b:
        return None
    d = max(len((m[0].split(".") + [""])[1]) for m in (a, b))
    m = (float(a[0].replace(",", "")) + float(b[0].replace(",", ""))) / 2
    return lo[:a.start()] + f"{m:,.{d}f}" + lo[a.end():]


def _facts(S: _Sheet, g: dict):
    f = S.f
    S.title("The report's key facts", "As the workbench read them from last year's report: who decided each, and how "
                                      "it was checked on its page and on the page's image.", 9)
    S.header("Fact", "Low", "Mid", "High", "Unit", "Basis", "Page", "Status", "Decided by", "On its page", "On its image")
    for x in g.get("facts") or []:
        v = {**x, **(x.get("final") or {})}
        lo, hi, mid = v.get("low_text") or "", v.get("high_text") or "", v.get("value_text") or ""
        if lo and hi and not mid:  # the report gives the ends only: their midpoint, marked as such
            mid = f"{_midpoint(lo, hi)} (midpoint)" if _midpoint(lo, hi) else ""
        if not lo and not hi and not re.search(r"\bmid|prefer", v.get("basis") or "", re.I):
            lo = hi = mid  # one figure: the same at every end (one the report gives at its mid stays the mid)
        seen = (x.get("visual") or {}).get("status") or ""
        S.line(x.get("label") or x.get("key"), lo, mid, hi, v.get("unit") or "", v.get("basis") or "",
               v.get("page") or x.get("page"), x.get("status") or "", x.get("decided_by") or "",
               mark((x.get("check") or {}).get("ok")), seen.replace("_", " "),
               fs=[f["bwrap"], f["wrap"], f["wrap"], f["wrap"], None, f["wrap"], None, None, None,
                   _okf(f, (x.get("check") or {}).get("ok")), f["wrap"]])
    T = (g.get("result") or {}).get("terminal") or g.get("terminal") or {}
    if T.get("passages"):
        S.gap()
        S.head(f"What the report says about the terminal value: {T.get('label')}")
        for x in T["passages"]:
            S.text(f"p. {x['page']}: {x['text']}", f["wrap"], 9, height=30 if len(x["text"]) > 150 else None)
    S.ws.freeze_panes(3, 1)


def _files(S: _Sheet, g: dict):
    f = S.f
    S.title("The files and their roles", "Which file plays which part, and who placed it there.", 7)
    S.header("Role", "File", "Valuation date", "Sheets used", "Placed by", "Confirmed", "Why")
    docs = {d["id"]: d for d in g.get("documents") or []}
    wbs = {w["id"]: w for w in g.get("workbooks") or []}
    used = set()
    for role, name in ROLES.items():
        r = (g.get("roles") or {}).get(role)
        if not r:
            S.line(name, "not placed yet", fs=[f["b"], f["muted"]])
            continue
        x = (docs if r.get("kind") == "document" else wbs).get(r.get("id")) or {}
        used.add((r.get("kind"), r.get("id")))
        who = {"orchestrator": "the orchestrator", "agents": "the agents"}.get(r.get("by"), r.get("by") or "")
        S.line(name, x.get("filename") or "", x.get("valuation_date") or "", ", ".join(r.get("sheets") or []) or "all",
               who, mark(r.get("confirmed")), "; ".join(r.get("why") or [])[:500],
               fs=[f["b"], f["wrap"], None, f["wrap"], None, _okf(f, r.get("confirmed")), f["muted"]])
    rest = [d["filename"] for d in docs.values() if ("document", d["id"]) not in used] + \
           [w["filename"] for w in wbs.values() if ("workbook", w["id"]) not in used]
    if rest:
        S.gap()
        S.line("Also uploaded", ", ".join(rest), fs=[f["b"], f["wrap"]])


def _review(S: _Sheet, review: dict):
    f = S.f
    S.title("The review", "The reviewer read the facts, the tie, the bridge, the cash flows and the run log end to "
                          "end, and said what looks wrong.", 6)
    if not review["ran"]:
        S.text("Not reviewed yet.", f["muted"], 6)
    elif review["said"]:
        S.text(review["said"], f["wrap"], 6, height=45 if len(review["said"]) > 200 else None)
        S.gap()
    if review["ran"] and not review["points"]:
        S.text("No concerns.", f["wrap"], 6)
    elif review["points"]:
        S.header("#", "Kind", "Point", "Detail", "Years", "Bridge step")
        for i, p in enumerate(review["points"], 1):
            S.line(i, "a note" if p.get("severity") == "info" else "to check", p.get("title") or "", p.get("detail") or "",
                   ", ".join(p.get("years") or []), STEP_SHORT.get(p.get("step"), p.get("step") or ""),
                   fs=[None, None, f["bwrap"], f["wrap"], f["wrap"], None], height=15 * max(1, len(p.get("detail") or "") // 70 + 1))
    # the checks a person acknowledged, with the reason, and those still open: part of the workpaper's record
    S.gap()
    S.text("Checks acknowledged by a person", f["bwrap"], 6)
    if review["acked"]:
        S.header("#", "By", "Check", "Reason", "When", "")
        for i, n in enumerate(review["acked"], 1):
            a = n.get("acked") or {}
            S.line(i, a.get("by") or "you", re.sub(r"^Acknowledged: ", "", n.get("title") or ""), a.get("reason") or "",
                   when(a.get("at")), "", fs=[None, None, f["bwrap"], f["wrap"], None, None],
                   height=15 * max(1, len(a.get("reason") or "") // 70 + 1))
    else:
        S.text("None.", f["muted"], 6)
    S.gap()
    S.text("Checks still open", f["bwrap"], 6)
    if review["open"]:
        S.header("#", "Kind", "Check", "Detail", "", "")
        for i, n in enumerate(review["open"], 1):
            S.line(i, "holds the value" if n.get("severity") == "block" else "to check", n.get("title") or "",
                   n.get("detail") or "", "", "", fs=[None, None, f["bwrap"], f["wrap"], None, None],
                   height=15 * max(1, len(n.get("detail") or "") // 70 + 1))
    else:
        S.text("None.", f["muted"], 6)


def _log(S: _Sheet, eid: int):
    f = S.f
    S.title("The run log", "Every start, outcome, decision and escalation, oldest first.", 5)
    S.header("When", "Stage", "Event", "What", "Issue")
    for h in reversed(orchestrator.history(eid, None, None, 2000)):
        text = h.get("text") or ""
        if h.get("event") == "decide":  # roles, not model names: "gpt-…: ok — why" as the reviewer's or the orchestrator's
            text = re.sub(r"^[\w.\-]+:\s*", "the reviewer: " if h.get("stage") == "review" else "the orchestrator: ", text)
        S.line(when(h.get("at")), orchestrator.LABEL.get(h.get("stage"), h.get("stage") or ""), h.get("event") or "",
               text, h.get("issue") or "", fs=[None, None, None, f["wrap"], f["muted"]])
    S.ws.freeze_panes(3, 0)


def _review_of(g: dict) -> dict:
    run = g.get("run") or {}
    points = [n for n in run.get("needs") or [] if n.get("stage") == "review"]
    rv = next((h for h in run.get("log") or [] if h.get("stage") == "review" and h.get("event") == "decide"), None)
    said = re.sub(r"^[\w.\-]+:\s*(ok|concerns)\s*[—–-]\s*", "", (rv or {}).get("text") or "", flags=re.I)
    st = next((s for s in run.get("stages") or [] if s.get("stage") == "review"), {})
    needs = run.get("needs") or []
    return {"points": points, "said": said, "ran": bool(rv) or st.get("status") in ("done", "attention"),
            "acked": [n for n in needs if n.get("acked")],
            "open": [n for n in needs if n.get("stage") != "review" and n.get("severity") in ("block", "check")]}


# ---- all of it ----------------------------------------------------------------------------------------------------

def ready(g: dict | None) -> bool:
    """The bridge is worked out, on the inputs as they are now (not an earlier result shown while it runs again)."""
    return bool(g and (g.get("result") or {}).get("bridges") and not g.get("result_stale"))


def filename(g: dict) -> str:
    """ASCII only: the name goes into a header."""
    name = re.sub(r"[^A-Za-z0-9 ._-]", "", g.get("name") or "").strip(" .") or f"engagement {g['id']}"
    return f"{name[:80]} workpaper.xlsx"


def build(eid: int) -> bytes | None:
    """The workpaper as .xlsx bytes; None if there's no such engagement. ValueError until there's a bridge."""
    g = wb.get(eid)
    if not g:
        return None
    if not ready(g):
        raise ValueError("the workpaper is ready once the bridge is worked out")
    res, review = g["result"], _review_of(g)
    buf = io.BytesIO()
    book = xlsxwriter.Workbook(buf, {"in_memory": True})
    book.set_properties({"title": "Valuation workpaper", "subject": g.get("name") or "",
                         "author": "Infra Val Workbench", "comments": DISCLAIMER})
    fmt = _formats(book)
    _summary(_Sheet(book, fmt, "Summary", [52, 16, 16, 16, 34, 46]), g, res, review)
    _bridge(_Sheet(book, fmt, "Bridge", [70, 14, 14, 14, 3, 3, 16, 12, 12, 12, 12]), book, res)
    _flows(_Sheet(book, fmt, "Cash flows", [16, 18, 18, 14]), book, res)
    _inputs(_Sheet(book, fmt, "Inputs", [26, 12, 8, 18, 34, 12, 26, 30, 70]), res)
    _methods(_Sheet(book, fmt, "Methods", [56, 14, 14, 14, 16, 3, 80]), res)
    _reconcile(_Sheet(book, fmt, "Reconciliation", [40, 30, 14, 14, 14, 14, 20, 10]), res)
    _facts(_Sheet(book, fmt, "Key facts", [30, 16, 18, 16, 8, 26, 6, 11, 12, 9, 16]), g)
    _files(_Sheet(book, fmt, "Files and roles", [26, 44, 16, 30, 18, 10, 60]), g)
    _review(_Sheet(book, fmt, "Review", [4, 10, 40, 80, 16, 14]), review)
    _log(_Sheet(book, fmt, "Run log", [22, 18, 12, 100, 16]), eid)
    book.close()
    return buf.getvalue()
