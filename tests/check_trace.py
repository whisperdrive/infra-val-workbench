"""The valuation tracer (engine/dcftrace.py) on a present-value row with quarters that have no cash flow.

  pv row    a figure = SUM of a present-value row (each column a cash flow times a factor), where some quarters have
            no cash flow: the factors are seen only where there is one (present value / cash flow), and the others
            aren't factors of 0. The rate, the valuation date and the convention are found from the ones seen, so
            the figure's discounting is known (the bridge's steps, the method selector, the DCF facts need it)
  mid-year  factors written (end - valuation date) / 365 - 0.5, the mid-year convention as models often have it,
            are read back as such

    uv run python tests/check_trace.py
"""
import sqlite3
import sys
import tempfile
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "engine"))
import xlsxwriter  # noqa: E402
from xlsxwriter.utility import xl_col_to_name as COL  # noqa: E402

import build_map  # noqa: E402
import dcftrace  # noqa: E402


def main() -> None:
    out = Path(tempfile.mkdtemp(prefix="trace_"))
    vd, rate, n = date(2025, 6, 30), 0.087, 24
    ends, y, m = [], 2025, 9
    for _ in range(n):
        ends.append(date(y + (m == 12), m % 12 + 1, 1) - timedelta(days=1))
        y, m = (y + 1, 3) if m == 12 else (y, m + 3)
    flows = [0.0 if k % 5 == 2 else 100.0 + 3 * k for k in range(n)]  # every fifth quarter has no cash flow
    df = [1 / (1 + rate) ** ((e - vd).days / 365) for e in ends]
    wb = xlsxwriter.Workbook(out / "pv.xlsx")
    dt = wb.add_format({"num_format": "dd-mmm-yy"})
    inp = wb.add_worksheet("Inputs")
    inp.write(3, 0, "Valuation date")
    inp.write_datetime(3, 2, vd, dt)
    inp.write(4, 0, "Discount rate")
    inp.write_number(4, 2, rate)
    fl = wb.add_worksheet("Flows")
    for r, label in ((2, "Period ending"), (9, "Cash flow"), (12, "Discount factor"), (13, "Present value"),
                     (15, "Equity value")):
        fl.write(r, 1, label)
    for k in range(n):
        c = COL(3 + k)
        fl.write_datetime(2, 3 + k, ends[k], dt)
        fl.write_number(9, 3 + k, flows[k])
        fl.write_formula(f"{c}13", f"=1/(1+Inputs!$C$5)^(({c}3-Inputs!$C$4)/365)", None, df[k])
        fl.write_formula(f"{c}14", f"={c}10*{c}13", None, flows[k] * df[k])
    fl.write_formula("C16", f"=SUM(D14:{COL(3 + n - 1)}14)", None, sum(f * x for f, x in zip(flows, df)))
    wb.close()
    db = sqlite3.connect(build_map.main(str(out / "pv.xlsx"), str(out / "db"))["db"])
    core = next(c for c in dcftrace.cores(dcftrace.trace(db, "Flows!C16")) if c.get("kind") == "pv row")
    m = core.get("method") or {}
    assert (m.get("rate"), m.get("valuation_date"), m.get("terminal_date")) == ("Inputs!C5", "Inputs!C4", None), m
    assert core.get("inputs", {}).get("cashflow") == ["Flows!D10:AA10"], core.get("inputs")
    assert core["periods"] == n - 5, core["periods"]
    print(f"pv row: ok (the discounting read from the {core['periods']} of {n} quarters with a cash flow: the rate "
          f"{m['rate']}, the valuation date {m['valuation_date']}, {m['timing']}, {m['day_count']}, no cut-off)")


def mid_year_check() -> None:
    out = Path(tempfile.mkdtemp(prefix="trace_mid_"))
    vd, rate, years = date(2025, 6, 30), 0.08, list(range(2026, 2036))
    wb = xlsxwriter.Workbook(out / "mid.xlsx")
    dt = wb.add_format({"num_format": "dd-mmm-yy"})
    inp = wb.add_worksheet("Inputs")
    inp.write(3, 0, "Valuation date")
    inp.write_datetime(3, 2, vd, dt)
    inp.write(4, 0, "Discount rate")
    inp.write_number(4, 2, rate)
    fl = wb.add_worksheet("Flows")
    for r, label in ((2, "Period ending"), (9, "Cash flow"), (12, "Discount factor"), (13, "Present value"),
                     (15, "Equity value")):
        fl.write(r, 1, label)
    total = 0.0
    for k, y in enumerate(years):
        c, end = COL(3 + k), date(y, 6, 30)
        f = 1 / (1 + rate) ** ((end - vd).days / 365 - 0.5)
        fl.write_datetime(2, 3 + k, end, dt)
        fl.write_number(9, 3 + k, 50.0 + k)
        fl.write_formula(f"{c}13", f"=1/(1+Inputs!$C$5)^(({c}3-Inputs!$C$4)/365-0.5)", None, f)
        fl.write_formula(f"{c}14", f"={c}10*{c}13", None, (50.0 + k) * f)
        total += (50.0 + k) * f
    fl.write_formula("C16", f"=SUM(D14:{COL(3 + len(years) - 1)}14)", None, total)
    wb.close()
    db = sqlite3.connect(build_map.main(str(out / "mid.xlsx"), str(out / "db"))["db"])
    core = next(c for c in dcftrace.cores(dcftrace.trace(db, "Flows!C16")) if c.get("kind") == "pv row")
    m = core.get("method") or {}
    assert (m.get("timing"), m.get("day_count"), m.get("rate")) == ("mid-year", "actual/365", "Inputs!C5"), m
    print("mid-year: ok (factors written (end - valuation date) / 365 - 0.5 read back as the mid-year convention)")


def sourcing_check() -> None:
    """The rate is the cell the factors read, followed through their formulas to the input; factors typed in as
    numbers read no cell, so the rate isn't sourced, even where a labelled rate cell holds the same number."""
    vd, rate, years = date(2025, 6, 30), 0.0725, list(range(2026, 2031))
    got = {}
    for how in ("via", "pasted"):
        out = Path(tempfile.mkdtemp(prefix=f"trace_{how}_"))
        wb = xlsxwriter.Workbook(out / f"{how}.xlsx")
        dt = wb.add_format({"num_format": "dd-mmm-yy"})
        inp = wb.add_worksheet("Inputs")
        inp.write(3, 0, "Valuation date")
        inp.write_datetime(3, 2, vd, dt)
        inp.write(4, 0, "Discount rate")
        inp.write_number(4, 2, rate)
        fl = wb.add_worksheet("Flows")
        for r, label in ((2, "Period ending"), (4, "Rate used"), (9, "Cash flow"), (12, "Discount factor"),
                         (13, "Present value"), (15, "Equity value")):
            fl.write(r, 1, label)
        fl.write_formula("C5", "=Inputs!C5", None, rate)
        total = 0.0
        for k, y in enumerate(years):
            c, end = COL(3 + k), date(y, 6, 30)
            f = 1 / (1 + rate) ** ((end - vd).days / 365)
            fl.write_datetime(2, 3 + k, end, dt)
            fl.write_number(9, 3 + k, 80.0 + k)
            if how == "via":
                fl.write_formula(f"{c}13", f"=1/(1+$C$5)^(({c}3-Inputs!$C$4)/365)", None, f)
            else:
                fl.write_number(12, 3 + k, f)
            fl.write_formula(f"{c}14", f"={c}10*{c}13", None, (80.0 + k) * f)
            total += (80.0 + k) * f
        fl.write_formula("C16", f"=SUM(D14:{COL(3 + len(years) - 1)}14)", None, total)
        wb.close()
        db = sqlite3.connect(build_map.main(str(out / f"{how}.xlsx"), str(out / "db"))["db"])
        got[how] = next(c for c in dcftrace.cores(dcftrace.trace(db, "Flows!C16")) if c.get("kind") == "pv row")["method"]
    via, pasted = got["via"], got["pasted"]
    assert via["rate"] == "Inputs!C5" and via["rate_source"]["input"] and \
        via["rate_source"]["chain"] == ["Flows!D14", "Flows!D13", "Flows!C5", "Inputs!C5"] and \
        via["valuation_date_sourced"], via
    assert isinstance(pasted["rate"], float) and abs(pasted["rate"] - rate) < 1e-9 and pasted["rate_source"] is None \
        and pasted["rate_note"].startswith("not sourced") and "Inputs!C5 holds the same number" in pasted["rate_note"] \
        and not pasted["valuation_date_sourced"], pasted
    print("sourcing: ok (the rate followed from the factors through Flows!C5 to the input Inputs!C5; factors typed in "
          "as numbers read no cell, so the rate isn't sourced, and a labelled cell holding the same number is only a hint)")


def growth_check() -> None:
    """The terminal growth rate is the cell the terminal value's formula reads as g in X x (1 + g) / (r - g), the
    terminal value recomputed from it at the end's own discount rate; a growth rate typed into the formula isn't
    sourced."""
    import sourced
    vd, rate, g, years = date(2025, 6, 30), 0.08, 0.025, list(range(2026, 2031))
    got = {}
    for how in ("cell", "typed"):
        out = Path(tempfile.mkdtemp(prefix=f"trace_tv_{how}_"))
        wb = xlsxwriter.Workbook(out / f"{how}.xlsx")
        dt = wb.add_format({"num_format": "dd-mmm-yy"})
        inp = wb.add_worksheet("Inputs")
        for r, label, v in ((4, "Discount rate", rate), (5, "Terminal growth rate", g)):
            inp.write(r, 0, label)
            inp.write_number(r, 2, v)
        inp.write(3, 0, "Valuation date")
        inp.write_datetime(3, 2, vd, dt)
        fl = wb.add_worksheet("Flows")
        for r, label in ((2, "Period ending"), (7, "Free cash flow"), (8, "Terminal value"), (9, "Valuation cash flow"),
                         (12, "Discount factor"), (13, "Present value"), (15, "Equity value")):
            fl.write(r, 1, label)
        total, last = 0.0, COL(3 + len(years) - 1)
        for k, y in enumerate(years):
            c, end = COL(3 + k), date(y, 6, 30)
            f, cf = 1 / (1 + rate) ** ((end - vd).days / 365), 100.0 + 5 * k
            tv = cf * (1 + g) / (rate - g) if k == len(years) - 1 else 0.0
            fl.write_datetime(2, 3 + k, end, dt)
            fl.write_number(7, 3 + k, cf)
            if tv:
                grow = "Inputs!$C$6" if how == "cell" else "0.025"
                fl.write_formula(f"{c}9", f"={c}8*(1+{grow})/(Inputs!$C$5-{grow})", None, tv)
            else:
                fl.write_number(8, 3 + k, 0.0)
            fl.write_formula(f"{c}10", f"={c}8+{c}9", None, cf + tv)
            fl.write_formula(f"{c}13", f"=1/(1+Inputs!$C$5)^(({c}3-Inputs!$C$4)/365)", None, f)
            fl.write_formula(f"{c}14", f"={c}10*{c}13", None, (cf + tv) * f)
            total += (cf + tv) * f
        fl.write_formula("C16", f"=SUM(D14:{last}14)", None, total)
        wb.close()
        db = sqlite3.connect(build_map.main(str(out / f"{how}.xlsx"), str(out / "db"))["db"])
        cs = [c for c in dcftrace.cores(dcftrace.trace(db, "Flows!C16")) if c.get("inputs")]
        rates = {"ends": {e: {"cell": "Inputs!C5", "value": rate} for e in ("low", "high")}}
        facts = [{"key": "terminal_growth_rate", "value_text": "2.50%", "status": "approved"}]
        got[how] = sourced.growth(db, {"low": cs, "high": cs}, rates, facts)["ends"]["low"]
    cell, typed = got["cell"], got["typed"]
    assert cell["sourced"] and cell["cell"] == "Inputs!C6" and cell["ok"] and cell["ties"], cell
    assert not typed["sourced"] and typed["cell"] is None and abs(typed["value"] - g) < 1e-12 and not typed["ok"] and \
        [c["ok"] for c in typed["checks"]] == [False, True, True, True], typed
    print("growth: ok (the terminal value's growth rate is the cell its formula reads, recomputed at the end's own "
          "discount rate; typed into the formula, it isn't sourced)")


def roll_dates_check() -> None:
    """The roll-forward moves every valuation date the discountings read. Three discountings: one reads the input
    (Inputs!C4), one a copy of it on its own sheet (=Inputs!C4), one a date typed on its own sheet. The input and
    the typed date are moved (the copy follows its input); on the rolled feed each discounting reads the new date,
    where moving the most common cell alone left the one with its own date discounting to last year's."""
    import overlay as ov
    import xlcompile
    from xlruntime import serial
    out = Path(tempfile.mkdtemp(prefix="trace_roll_"))
    vd, new, rate, years = date(2025, 6, 30), date(2026, 6, 30), 0.08, list(range(2026, 2036))
    wb = xlsxwriter.Workbook(out / "roll.xlsx")
    dt = wb.add_format({"num_format": "dd-mmm-yy"})
    inp = wb.add_worksheet("Inputs")
    inp.write(3, 0, "Valuation date")
    inp.write_datetime(3, 2, vd, dt)
    inp.write(4, 0, "Discount rate")
    inp.write_number(4, 2, rate)
    flows = [100.0 + 5 * k for k in range(len(years))]
    for name, date_ref in (("DCF_Low", "Inputs!$C$4"), ("DCF_High", "$C$4"), ("DCF_Own", "$C$4")):
        sh = wb.add_worksheet(name)
        sh.write(3, 1, "Valuation date")
        if name == "DCF_High":
            sh.write_formula("C4", "=Inputs!C4", dt, (vd - date(1899, 12, 30)).days)
        elif name == "DCF_Own":
            sh.write_datetime(3, 2, vd, dt)
        for r, label in ((5, "Period ending"), (9, "Cash flow"), (12, "Discount factor"), (13, "Present value"),
                         (15, "Equity value")):
            sh.write(r, 1, label)
        total = 0.0
        for k, y in enumerate(years):
            c, end = COL(3 + k), date(y, 6, 30)
            f = 1 / (1 + rate) ** ((end - vd).days / 365)
            total += flows[k] * f
            sh.write_datetime(5, 3 + k, end, dt)
            sh.write_number(9, 3 + k, flows[k])
            sh.write_formula(f"{c}13", f"=1/(1+Inputs!$C$5)^(({c}6-{date_ref})/365)", None, f)
            sh.write_formula(f"{c}14", f"={c}10*{c}13", None, flows[k] * f)
        sh.write_formula("C16", f"=SUM(D14:{COL(3 + len(years) - 1)}14)", None, total)
    wb.close()
    db = build_map.main(str(out / "roll.xlsx"), str(out / "db"))["db"]
    sheets = ["Inputs", "DCF_Low", "DCF_High", "DCF_Own"]
    outputs = [{"cell": f"{s}!C16"} for s in sheets[1:]]
    move, read = ov.discount_date_cells(db, outputs, sheets)
    assert sorted(move) == ["DCF_Own!C4", "Inputs!C4"], move
    assert sorted(read) == ["DCF_High!C4", "DCF_Own!C4", "Inputs!C4"], read
    src, _ = xlcompile.compile_overlay(db, sheets)
    mod = out / "overlay.py"
    mod.write_text(src)
    sess = ov.Session(str(mod), db, sheets)
    sess.configure("workbook", {ov.parse_a1(c): serial(new) for c in move})
    want = sum(x / (1 + rate) ** ((date(y, 6, 30) - new).days / 365) for x, y in zip(flows, years))
    got = sess.values([ov.parse_a1(o["cell"]) for o in outputs])
    assert all(abs(g - want) < 1e-6 for g in got), (got, want)
    assert all(round(v) == serial(new) for v in sess.values([ov.parse_a1(c) for c in read]))
    print(f"roll dates: ok (the input and a discounting's own date moved, the copy with its input: all three "
          f"discountings at {new:%d %B %Y}, {want:,.1f} each)")


def cutoff_check() -> None:
    """Rolled forward, the periods before the new valuation date are cut off the discounting. The overlay (inside
    last year's client model) discounts its client sheet's cash flows with a factor worked out from the date, and
    last year's model has nothing on or before last year's date, so the overlay never needed a cut-off. This year's
    model, on the same quarters (a fixed horizon), has the year between the two dates filled in: without the cut-off
    those quarters are compounded into this year's value; with it, the value is the quarters after the new date (the
    one ending on it is past too, the convention dcf keeps; kept, it's in undiscounted: a method of its own). And the
    roll moves the discounting on by its own rate (the time check); with the valuation date left where it was, it
    doesn't, and this year's value is held back."""
    import overlay as ov
    import result
    import xlcompile
    from xlruntime import serial
    out = Path(tempfile.mkdtemp(prefix="trace_cutoff_"))
    vd, new, rate = date(2025, 6, 30), date(2026, 6, 30), 0.085
    ends, y, m = [], 2024, 9
    for _ in range(24):  # quarters, 30 September 2024 to 30 June 2030
        ends.append(date(y + (m == 12), m % 12 + 1, 1) - timedelta(days=1))
        y, m = (y + 1, 3) if m == 12 else (y, m + 3)
    flows = {"prior": [0.0 if e <= vd else 50.0 + 2 * k for k, e in enumerate(ends)],
             "current": [0.0 if e <= vd else 52.0 + 2 * k for k, e in enumerate(ends)]}
    paths = {}
    for which in ("prior", "current"):
        wb = xlsxwriter.Workbook(out / f"{which}.xlsx")
        dt = wb.add_format({"num_format": "dd-mmm-yy"})
        cl = wb.add_worksheet("Client")
        cl.write(5, 1, "Period ending")
        cl.write(9, 1, "Distributions to equity")
        for k, e in enumerate(ends):
            cl.write_datetime(5, 3 + k, e, dt)
            cl.write_number(9, 3 + k, flows[which][k])
        if which == "prior":  # the overlay, inside last year's client model
            va = wb.add_worksheet("Val")
            va.write(3, 1, "Valuation date")
            va.write_datetime(3, 2, vd, dt)
            va.write(4, 1, "Discount rate")
            va.write_number(4, 2, rate)
            for r, label in ((5, "Period ending"), (9, "Cash flow"), (12, "Discount factor"), (13, "Present value"),
                             (15, "Equity value")):
                va.write(r, 1, label)
            total = 0.0
            for k, e in enumerate(ends):
                c = COL(3 + k)
                f = 1 / (1 + rate) ** ((e - vd).days / 365)
                total += flows["prior"][k] * f
                va.write_formula(f"{c}6", f"=Client!{c}6", dt, (e - date(1899, 12, 30)).days)
                va.write_formula(f"{c}10", f"=Client!{c}10", None, flows["prior"][k])
                va.write_formula(f"{c}13", f"=1/(1+$C$5)^(({c}6-$C$4)/365)", None, f)
                va.write_formula(f"{c}14", f"={c}10*{c}13", None, flows["prior"][k] * f)
            va.write_formula("C16", f"=SUM(D14:{COL(3 + len(ends) - 1)}14)", None, total)
        wb.close()
        paths[which] = build_map.main(str(out / f"{which}.xlsx"), str(out / f"db_{which}"))["db"]
    db = paths["prior"]
    src, _ = xlcompile.compile_overlay(db, ["Val"])
    mod = out / "overlay.py"
    mod.write_text(src)
    sess = ov.Session(str(mod), db, ["Val"], None, paths["current"], None, ["Client"])
    roll = ov.plan_roll(sess, None, {"sheets": ["Val"]}, True, vd.isoformat(), None, new.isoformat())
    roll.update(ov.date_cells(db, [{"cell": "Val!C16"}], ["Val"], None))
    assert roll["months"] == 12 and roll["fixed_horizon"] and roll["valuation_date_cells"] == ["Val!C4"], roll
    assert len(roll["cutoff"]) == len(ends) and roll["cutoff"][0] == ["Val!D14", ends[0].isoformat(), "Val!D6"], roll["cutoff"][:2]

    def at_new(cutoffs, keep=False):
        sess.cutoffs, sess.keep_on_date = cutoffs, keep
        sess.configure("current", {("Val", 4, 3): serial(new)}, roll["months"])
        v = sess.values([("Val", 16, 3)])[0]
        n_cut = len(sess.cut)
        sess.configure("workbook")
        sess.keep_on_date = False
        return v, n_cut
    pv = lambda keep: sum(x / (1 + rate) ** ((e - new).days / 365) for x, e in zip(flows["current"], ends) if keep(e))
    want, on, past = pv(lambda e: e > new), pv(lambda e: e == new), pv(lambda e: vd < e < new)
    cut = [(*ov.parse_a1(c), serial(date.fromisoformat(d)), *[ov.parse_a1(x) for x in at]) for c, d, *at in roll["cutoff"]]
    got, n_cut = at_new(cut)
    kept, n_kept = at_new(cut, keep=True)
    without, _ = at_new([])
    assert abs(got - want) < 1e-6, (got, want)
    assert abs(kept - (want + on)) < 1e-6 and on > 0 and n_kept == n_cut - 1, (kept, want, on)
    assert abs(without - (want + on + past)) < 1e-6 and past > 0, (without, want, past)
    assert n_cut == sum(1 for e in ends if e <= new), n_cut  # the quarter ending on the new date is cut too
    sess.base_vd = serial(vd)  # at last year's date (the zero-roll check) the cut-off is last year's: nothing of it
    sess.cutoffs = cut
    sess.configure("current", {}, 0)
    assert {k[2] for k in sess.cut} == {c for (_, _, c, e, *_) in cut if e <= serial(vd)}, sorted(sess.cut)
    sess.configure("workbook")
    summary = {"wiring": {"overlay": {"db_path": db}}, "sheets": ["Val"], "roll": roll, "held_values": {},
               "outputs": [{"cell": "Val!C16"}]}
    tc = ov.deep(result.time_check, sess, summary, ["Val!C16"])
    assert tc["measured"] == 1 and tc["ok"] and not tc["hold"] and abs(tc["discountings"][0]["off"]) < 1e-6, tc
    stuck = ov.deep(result.time_check, sess, {**summary, "roll": {**roll, "valuation_date_cells": [], "valuation_date_cell": None}},
                    ["Val!C16"])
    assert stuck["hold"] and abs(stuck["discountings"][0]["implied"]) < 1e-6, stuck
    print(f"cut-off: ok (rolled a year onto a fixed horizon, the {sum(1 for e in ends if vd < e <= new)} quarters to the "
          f"new date are cut off this year's discounting: {got:,.1f}, not {without:,.1f} with them compounded in; the one "
          f"ending on the date kept, {kept:,.1f}; moved on by its rate, and held back where the date isn't moved)")


def masked_check() -> None:
    """A SUMPRODUCT of flags, cash flows and factors (two windows of years, each its own SUMPRODUCT), the factors on a
    convention the app doesn't recompute (mid-year, the first part-year at its midpoint, the rate read from a row
    that copies the input): each window is a discounting, its flags a mask dcf.compute multiplies in; its rate and
    valuation date are the cells the factors' formulas read (loose), not fitted. And rows that carry periods of their
    own (an annual block under a quarterly header: year numbers, period ends with stubs) are told from amounts."""
    import dcf
    import overlay as ov
    from xlruntime import serial
    out = Path(tempfile.mkdtemp(prefix="trace_masked_"))
    vd, rate, years = date(2025, 9, 30), 0.095, list(range(2026, 2041))
    wb = xlsxwriter.Workbook(out / "masked.xlsx")
    dt = wb.add_format({"num_format": "dd-mmm-yy"})
    fl = wb.add_worksheet("Val")
    fl.write(3, 1, "Valuation date")
    fl.write_datetime(3, 2, vd, dt)
    fl.write(4, 1, "Cost of equity")
    fl.write_number(4, 2, rate)
    for r, label in ((5, "Period ending"), (7, "Early window"), (8, "Late window"), (9, "Distributions"), (10, "Discount rate"),
                     (11, "Discount period"), (12, "Discount factor"), (15, "PV early"), (16, "PV late")):
        fl.write(r, 1, label)
    pv = {"early": 0.0, "late": 0.0}
    last = COL(3 + len(years) - 1)
    for k, y in enumerate(years):
        c, end = COL(3 + k), date(y, 6, 30)
        yf = dcf.yearfrac(vd, end, "actual/actual")
        t = yf / 2 if yf < 1 else yf - 0.5
        f, cf, early = 1 / (1 + rate) ** t, 100.0 + 4 * k, 1.0 if k < 5 else 0.0
        fl.write_datetime(5, 3 + k, end, dt)
        fl.write_number(7, 3 + k, early)
        fl.write_formula(f"{c}9", f"=1-{c}8", None, 1 - early)
        fl.write_number(9, 3 + k, cf)
        fl.write_formula(f"{c}11", "=$C$5", None, rate)
        fl.write_formula(f"{c}12", f"=IF(YEARFRAC($C$4,{c}6,1)<1,YEARFRAC($C$4,{c}6,1)/2,YEARFRAC($C$4,{c}6,1)-0.5)", None, t)
        fl.write_formula(f"{c}13", f"=1/(1+{c}11)^{c}12", None, f)
        pv["early" if early else "late"] += cf * f
    fl.write_formula("C16", f"=SUMPRODUCT(D8:{last}8,D10:{last}10,D13:{last}13)", None, pv["early"])
    fl.write_formula("C17", f"=SUMPRODUCT(D9:{last}9,D10:{last}10,D13:{last}13)", None, pv["late"])
    fl.write_formula("C18", "=C16+C17", None, pv["early"] + pv["late"])
    wb.close()
    path = build_map.main(str(out / "masked.xlsx"), str(out / "db"))["db"]
    db = sqlite3.connect(path)
    cs = dcftrace.cores(dcftrace.trace(db, "Val!C18"))
    assert len(cs) == 2 and all(c["kind"] == "sumproduct" and len(c["mask"]) == 1 for c in cs), cs
    assert all(not c.get("inputs") and c["loose"]["rate"] == "Val!C5" and c["loose"]["valuation_date"] == "Val!C4"
               and abs(c["loose"]["rate_value"] - rate) < 1e-9 for c in cs), [c.get("loose") for c in cs]
    assert abs(sum(c["pv"] for c in cs) - (pv["early"] + pv["late"])) < 1e-6
    move, read = ov.discount_date_cells(path, [{"cell": "Val!C18"}], ["Val"])
    assert move == read == ["Val!C4"], (move, read)
    # with the factors as a row dcf.compute reads, the mask multiplies the cash flows in
    early = next(c for c in cs if c["cell"] == "Val!C16")
    got = dcf.compute(db, cashflow=early["cashflow"], rate="Val!C5", valuation_date="Val!C4", mask=early["mask"],
                      timing="end", day_count="actual/365", fix=False)
    want = sum(100.0 + 4 * k for k in range(5))  # the early window's cash flows only
    assert abs(got["undiscounted"] - want) < 1e-9, (got["undiscounted"], want)
    # rows with periods of their own: period ends with a stub at each end, year numbers; and an amount isn't one
    ends = [serial(date(2016, 10, 31))] + [serial(date(y, 6, 30)) for y in range(2017, 2067)] + [serial(date(2066, 10, 31))]
    a, b = ov.own_rule(ends), ov.own_rule([float(y) for y in range(2012, 2068)])
    assert a and a["kind"] == "own_date" and a["plen"] == 12 and "the first and the last a stub" in a["text"], a
    assert b and b["kind"] == "own_year", b
    assert ov.own_rule([100.0 + 3 * k for k in range(20)]) is None and ov.own_rule([float(serial(date(2020, 1, 15))), 44000.0, 44100.0]) is None
    print("masked: ok (a SUMPRODUCT with flags is two windows' discountings, the flags a mask; on a convention the app "
          "doesn't recompute, the rate and the date are the cells the factors read; rows of periods of their own told "
          "from amounts)")


def horizon_check() -> None:
    """Rolled a year, this year's model forecasting half a year further than last year's (half-yearly periods): the
    overlay's own periods move a year, so its last ones fall past this year's forecast, where nothing stands in (not
    last year's figures, not this year's last period read again from a column past last year's timeline); and the
    overlay's own copy of the forecast's end (a typed date its terminal value is placed at) goes to this year's."""
    import overlay as ov
    import result
    import xlcompile
    from xlruntime import serial
    out = Path(tempfile.mkdtemp(prefix="trace_horizon_"))
    vd, new, rate = date(2025, 6, 30), date(2026, 6, 30), 0.09
    half = lambda y, k: date(y + (k + 1) // 2, 12 if (k % 2 == 0) else 6, 31 if k % 2 == 0 else 30)
    ly = [half(2025, k) for k in range(10)]  # 31 Dec 2025 to 30 Jun 2030
    ty = [half(2026, k - 1) for k in range(1, 11)]  # 30 Jun 2026 to 31 Dec 2030
    flows = {"prior": [10.0 + k for k in range(10)], "current": [12.0 + k for k in range(10)]}
    paths = {}
    for which, ends in (("prior", ly), ("current", ty)):
        wb = xlsxwriter.Workbook(out / f"{which}.xlsx")
        dt = wb.add_format({"num_format": "dd-mmm-yy"})
        cl = wb.add_worksheet("Client")
        cl.write(5, 1, "Period ending")
        cl.write(9, 1, "Distributions")
        for k, e in enumerate(ends):
            cl.write_datetime(5, 3 + k, e, dt)
            cl.write_number(9, 3 + k, flows[which][k])
        if which == "prior":
            va = wb.add_worksheet("Val")
            for r, label, v in ((3, "Valuation date", vd), (4, "Model end date", ly[-1])):
                va.write(r, 1, label)
                va.write_datetime(r, 2, v, dt)
            va.write(5, 1, "Discount rate")
            va.write_number(5, 2, rate)
            va.write(6, 1, "Terminal value")
            va.write_number(6, 2, 50.0)
            for r, label in ((7, "Period ending"), (9, "Distributions"), (10, "Terminal value at the model end"),
                             (11, "Cash flow"), (12, "Discount factor"), (13, "Present value"), (15, "Equity value")):
                va.write(r, 1, label)
            total = 0.0
            ends_ov = ly + [half(2030, 0), half(2030, 1)]  # two periods past last year's forecast
            for k, e in enumerate(ends_ov):
                c = COL(3 + k)
                f = 1 / (1 + rate) ** ((e - vd).days / 365)
                cf = (flows["prior"][k] if k < 10 else 0.0) + (50.0 if e == ly[-1] else 0.0)
                total += cf * f
                va.write_datetime(7, 3 + k, e, dt)
                va.write_formula(f"{c}10", f"=Client!{c}10", None, flows["prior"][k] if k < 10 else 0.0)
                va.write_formula(f"{c}11", f"=IF({c}8=$C$5,$C$7,0)", None, 50.0 if e == ly[-1] else 0.0)
                va.write_formula(f"{c}12", f"={c}10+{c}11", None, cf)
                va.write_formula(f"{c}13", f"=1/(1+$C$6)^(({c}8-$C$4)/365)", None, f)
                va.write_formula(f"{c}14", f"={c}12*{c}13", None, cf * f)
            va.write_formula("C16", f"=SUM(D14:{COL(3 + len(ends_ov) - 1)}14)", None, total)
        wb.close()
        paths[which] = build_map.main(str(out / f"{which}.xlsx"), str(out / f"db_{which}"))["db"]
    db = paths["prior"]
    src, _ = xlcompile.compile_overlay(db, ["Val"])
    (out / "overlay.py").write_text(src)
    sess = ov.Session(str(out / "overlay.py"), db, ["Val"], None, paths["current"], None, ["Client"])
    sess.horizon_set = "rolling"
    roll = ov.plan_roll(sess, None, {"sheets": ["Val"]}, True, vd.isoformat(), None, new.isoformat())
    roll.update(ov.date_cells(db, [{"cell": "Val!C16"}], ["Val"], {"cell": "Val!C4"}))
    sess.cutoffs = [(*ov.parse_a1(c), serial(date.fromisoformat(d)), *[ov.parse_a1(x) for x in at]) for c, d, *at in roll["cutoff"]]
    summary = {"wiring": {"overlay": {"db_path": db}}, "sheets": ["Val"], "roll": roll, "held_values": {}}
    where = {"low": "Val!C16", "high": "Val!C16"}
    hz = ov.deep(result.this_year_horizon, sess, summary, where)
    assert hz and hz["was"] == ly[-1].isoformat() and hz["now"] == ty[-1].isoformat() and hz["cells"] == ["Val!C5"], hz
    d, _, m = ov._feed(summary, "current", None, None)
    sess.configure("current", d, m)
    v = sess.values([("Val", 16, 3)])[0]
    beyond = dict(sess.beyond)
    sess.configure("workbook")
    want = sum(x / (1 + rate) ** ((e - new).days / 365) for x, e in zip(flows["current"], ty) if e > new) + \
        50.0 / (1 + rate) ** ((ty[-1] - new).days / 365)
    assert abs(v - want) < 1e-6, (v, want)
    assert beyond and all(x for x in beyond.values()), beyond
    print(f"horizon: ok (this year's model half a year further on, rolled a year: nothing past its forecast, the "
          f"terminal value at its end, {v:,.2f})")


def window_check() -> None:
    """An overlay whose periods count from the valuation date, on a client model with a fixed calendar (the same
    quarters both years), this year's forecasting a year further: rolled a year, the valuation date moves on the cell
    it's typed in, not on the copy the discounting reads through a name (=Val_Date), so the cash flow dates counted
    from it and the discount periods counted from the copy stay one column; the cut-off reads each period's date as
    this year's feed works it out (the columns are this year's quarters: nothing to cut); and the terminal value date
    a sum up to a date reads (SUMIF(dates, "<"&end+1, present values)) goes to this year's last cash flow."""
    import overlay as ov
    import result
    import xlcompile
    from xlruntime import serial
    out = Path(tempfile.mkdtemp(prefix="trace_window_"))
    vd, new, rate, mult = date(2025, 6, 30), date(2026, 6, 30), 0.1, 10.0
    eom = lambda d, n: date(d.year + (d.month + n) // 12, (d.month + n) % 12 + 1, 1) - timedelta(days=1)  # EOMONTH(d, n)
    cal =[eom(vd, 3 * (k + 1)) for k in range(12)]  # the client's quarters, 30 September 2025 to 30 June 2028
    flows = {"prior": [20.0 + k if k < 8 else 0.0 for k in range(12)], "current": [21.0 + k for k in range(12)]}
    tv_date = cal[7]  # last year's last cash flow, 30 June 2027
    paths = {}
    for which in ("prior", "current"):
        wb = xlsxwriter.Workbook(out / f"{which}.xlsx")
        dt = wb.add_format({"num_format": "dd-mmm-yy"})
        cl = wb.add_worksheet("Client")
        cl.write(5, 1, "Period ending")
        cl.write(9, 1, "Distributions")
        for k, e in enumerate(cal):
            cl.write_datetime(5, 3 + k, e, dt)
            cl.write_number(9, 3 + k, flows[which][k])
        if which == "prior":  # the overlay, inside last year's client model
            inp = wb.add_worksheet("Inputs")
            inp.write(3, 1, "Valuation date")
            inp.write_datetime(3, 2, vd, dt)
            inp.write(4, 1, "Terminal value date")
            inp.write_datetime(4, 2, tv_date, dt)
            wb.define_name("Val_Date", "=Inputs!$C$4")
            va = wb.add_worksheet("Val")
            va.write(3, 1, "Valuation")
            va.write_formula("C4", "=Val_Date", dt, (vd - date(1899, 12, 30)).days)
            va.write(4, 1, "Discount rate")
            va.write_number(4, 2, rate)
            va.write(5, 1, "Exit multiple")
            va.write_number(5, 2, mult)
            va.write(8, 1, "Horizon")
            va.write_formula("C9", "=Inputs!C5", dt, (tv_date - date(1899, 12, 30)).days)
            for r, label in ((6, "Cash flow date"), (7, "Period ending"), (9, "Distributions"), (10, "Terminal value"),
                             (12, "Discount factor"), (13, "Present value"), (15, "Equity value")):
                va.write(r, 1, label)
            ser = lambda d: (d - date(1899, 12, 30)).days
            total = 0.0
            for k in range(8):
                c, p = COL(3 + k), COL(2 + k)
                e = cal[k]
                f = 1 / (1 + rate) ** ((e - vd).days / 365)
                tv = flows["prior"][k] * mult if e == tv_date else 0.0
                total += flows["prior"][k] * f + tv * f
                va.write_formula(f"{c}7", "=EOMONTH(Inputs!$C$4,3)" if k == 0 else f"=EOMONTH({p}7,3)", dt, ser(e))
                va.write_formula(f"{c}8", "=EOMONTH($C$4,3)" if k == 0 else f"=EOMONTH({p}8,3)", dt, ser(e))
                va.write_formula(f"{c}10", f"=SUMIFS(Client!$D$10:$O$10,Client!$D$6:$O$6,{c}7)", None, flows["prior"][k])
                va.write_formula(f"{c}11", f"=IF({c}7=$C$9,{c}10*$C$6,0)", None, tv)
                va.write_formula(f"{c}13", f"=1/(1+$C$5)^(({c}8-$C$4)/365)", None, f)
                va.write_formula(f"{c}14", f"={c}10*{c}13", None, flows["prior"][k] * f)
            va.write_formula("C16", '=SUMIF(D7:K7,"<"&$C$9+1,D14:K14)+SUMPRODUCT(D11:K11,D13:K13)', None, total)
        wb.close()
        paths[which] = build_map.main(str(out / f"{which}.xlsx"), str(out / f"db_{which}"))["db"]
    db = paths["prior"]
    sheets = ["Inputs", "Val"]
    src, _ = xlcompile.compile_overlay(db, sheets)
    (out / "overlay.py").write_text(src)
    sess = ov.Session(str(out / "overlay.py"), db, sheets, None, paths["current"], None, ["Client"])
    roll = ov.plan_roll(sess, None, {"sheets": sheets}, True, vd.isoformat(), None, new.isoformat())
    roll.update(ov.date_cells(db, [{"cell": "Val!C16"}], sheets, None))
    assert roll["months"] == 12 and roll["fixed_horizon"], roll
    assert roll["valuation_date_cells"] == ["Inputs!C4"], roll["valuation_date_cells"]  # through the name to its cell
    assert roll["cutoff"] and all(len(x) == 3 and x[2].startswith("Val!") for x in roll["cutoff"]), roll["cutoff"][:2]
    sess.cutoffs = [(*ov.parse_a1(c), serial(date.fromisoformat(d)), *[ov.parse_a1(x) for x in at]) for c, d, *at in roll["cutoff"]]
    summary = {"wiring": {"overlay": {"db_path": db}}, "sheets": sheets, "roll": roll, "held_values": {}}
    where = {"low": "Val!C16", "high": "Val!C16"}
    hz = ov.deep(result.this_year_horizon, sess, summary, where)
    assert hz and hz["was"] == tv_date.isoformat() and hz["now"] == cal[-1].isoformat() and hz["cells"] == ["Inputs!C5"], hz
    d, _, m = ov._feed(summary, "current", None, None)
    sess.configure("current", d, m)
    v = sess.values([("Val", 16, 3)])[0]
    n_cut = len(sess.cut)
    sess.configure("workbook")
    fac = lambda e: 1 / (1 + rate) ** ((e - new).days / 365)
    want = sum(x * fac(e) for x, e in zip(flows["current"], cal) if e > new) + flows["current"][-1] * mult * fac(cal[-1])
    assert n_cut == 0, n_cut
    assert abs(v - want) < 1e-6, (v, want)
    print(f"window: ok (periods counted from the valuation date, read through a name: moved on its cell, nothing cut "
          f"off this year's quarters, the terminal value date to this year's last cash flow, {v:,.2f})")


def sum_rows_check() -> None:
    """Two blocks of cash flows discounted at one factor row, their present-value rows added up column by column, and
    a terminal value discounted on its own (one period, its factor row nil elsewhere): each a discounting of the
    blocks' cash flows together, its rate and valuation date the cells the parts' factors read, not a cell labelled
    as a valuation date that they don't (one factor fits any date). The figure above them, through an IF and a
    units name (=(forecast + IF(switch, terminal, )) / thousand), is recomputed from them; the units constant isn't
    an input held at last year's, the switch is."""
    import dcf
    import held
    import valuation
    vd, rate, mult = date(2025, 6, 30), 0.1, 10.0
    ends, y, m = [], 2025, 9
    for _ in range(8):
        ends.append(date(y + (m == 12), m % 12 + 1, 1) - timedelta(days=1))
        y, m = (y + 1, 3) if m == 12 else (y, m + 3)
    a, b = [10.0 + k for k in range(8)], [5.0 + 2 * k for k in range(8)]
    out = Path(tempfile.mkdtemp(prefix="trace_sumrows_"))
    wb = xlsxwriter.Workbook(out / "sum.xlsx")
    dt = wb.add_format({"num_format": "dd-mmm-yy"})
    ser = lambda d: (d - date(1899, 12, 30)).days
    inp = wb.add_worksheet("Inputs")
    inp.write(3, 1, "Discounting from")
    inp.write_datetime(3, 2, vd, dt)
    inp.write(4, 1, "Discount rate")
    inp.write_number(4, 2, rate)
    cl = wb.add_worksheet("Client")
    cl.write(2, 1, "Valuation date (prior)")  # a year earlier: one factor fits it too, with another rate
    cl.write_datetime(2, 2, date(2024, 6, 30), dt)
    cl.write(3, 1, "Valuation date")  # the same date, labelled so, read by nothing the factors read
    cl.write_datetime(3, 2, vd, dt)
    va = wb.add_worksheet("Val")
    labels = ((2, "Period ending"), (5, "Rate"), (6, "Discount period"), (7, "Period"), (9, "Distributions A"),
              (10, "Distributions B"), (12, "Discount factor"), (14, "PV A"), (15, "PV B"), (16, "Present value"),
              (18, "Terminal flag"), (20, "Terminal date"), (21, "Terminal value A"), (22, "Terminal value B"),
              (23, "Terminal discount factor"), (25, "Terminal PV A"), (26, "Terminal PV B"),
              (27, "Terminal value present value"), (29, "PV forecast"), (30, "PV terminal"), (32, "Equity value"),
              (34, "Terminal value switch"), (35, "Units"))
    for r, label in labels:
        va.write(r, 1, label)
    va.write_formula("C6", "=Inputs!C5", None, rate)
    va.write_datetime(20, 2, ends[-1], dt)
    pv_f = pv_t = 0.0
    for k, e in enumerate(ends):
        c = COL(3 + k)
        t = (e - vd).days / 365
        f = 1 / (1 + rate) ** t
        last = k == len(ends) - 1
        va.write_datetime(2, 3 + k, e, dt)
        va.write_formula(f"{c}7", f"=({c}3-Inputs!$C$4)/365", None, t)
        va.write_formula(f"{c}8", f"={c}7", None, t)
        va.write_number(9, 3 + k, a[k])
        va.write_number(10, 3 + k, b[k])
        va.write_formula(f"{c}13", f"=1/(1+$C$6)^{c}8", None, f)
        va.write_formula(f"{c}15", f"={c}10*{c}13", None, a[k] * f)
        va.write_formula(f"{c}16", f"={c}11*{c}13", None, b[k] * f)
        va.write_formula(f"{c}17", f"=SUM({c}15,{c}16)", None, (a[k] + b[k]) * f)
        va.write_formula(f"{c}19", f"=IF({c}3=$C$21,1,0)", None, 1 if last else 0)
        ta, tb = (a[k] * mult, b[k] * mult) if last else (0.0, 0.0)
        va.write_formula(f"{c}22", f"={c}10*{mult}*{c}19", None, ta)
        va.write_formula(f"{c}23", f"={c}11*{mult}*{c}19", None, tb)
        va.write_formula(f"{c}24", f"=IF({c}19,1/(1+$C$6)^{c}8,0)", None, f if last else 0.0)
        va.write_formula(f"{c}26", f"={c}22*{c}24", None, ta * (f if last else 0.0))
        va.write_formula(f"{c}27", f"={c}23*{c}24", None, tb * (f if last else 0.0))
        va.write_formula(f"{c}28", f"=SUM({c}26,{c}27)", None, (ta + tb) * (f if last else 0.0))
        pv_f += (a[k] + b[k]) * f
        pv_t += (ta + tb) * (f if last else 0.0)
    va.write_formula("C30", "=SUM(D17:K17)", None, pv_f)
    va.write_formula("C31", "=SUM(D28:K28)", None, pv_t)
    va.write_number(34, 2, 1)
    va.write_number(35, 2, 1000)
    wb.define_name("thousand", "=Val!$C$36")
    va.write_formula("C33", "=(C30+IF(C35,C31,))/thousand", None, (pv_f + pv_t) / 1000)
    wb.close()
    db = sqlite3.connect(build_map.main(str(out / "sum.xlsx"), str(out / "db"))["db"])
    t = dcftrace.trace(db, "Val!C33")
    cs = {c["cell"]: c for c in dcftrace.cores(t) if c.get("inputs")}
    assert set(cs) == {"Val!C30", "Val!C31"}, sorted(cs)
    assert cs["Val!C30"]["inputs"]["cashflow"] == ["Val!D10:K10", "Val!D11:K11"], cs["Val!C30"]["inputs"]["cashflow"]
    assert cs["Val!C31"]["inputs"]["cashflow"] == ["Val!D22:K22", "Val!D23:K23"] and cs["Val!C31"]["periods"] == 1
    for c in cs.values():
        i = c["inputs"]
        assert (i["rate"], i["valuation_date"]) == ("Inputs!C5", "Inputs!C4"), (c["cell"], i["rate"], i["valuation_date"])
    # one factor, its formulas read for the rate alone: the labelled dates are all there is, and only the one whose
    # rate a cell they read holds is the date (the earlier date fits the factor too, at a rate no cell holds)
    one = valuation.read_factors(db, None, [11], theirs={11: 1 / (1 + rate) ** ((ends[-1] - vd).days / 365)},
                                 sheet="Val", starts={"expr": "Inputs!C5", "here": "Val"})
    assert one and (one["valuation_date"], one["rate"]) == ("Client!C4", "Inputs!C5"), one
    totals = {k: dcf.compute(db, **{**c["inputs"], "compare_to": None}, fix=False)["total"] for k, c in cs.items()}
    assert abs(totals["Val!C30"] - pv_f) < 1e-9 and abs(totals["Val!C31"] - pv_t) < 1e-9, (totals, pv_f, pv_t)
    again = dcftrace.recompute(db, t, {"Val!C30": 2 * pv_f, "Val!C31": pv_t})
    assert again is not None and abs(again - (2 * pv_f + pv_t) / 1000) < 1e-12, again
    found = {x["cell"] for x in held.find(db, [("low", "Val!C33")], ["Inputs", "Val"])}
    assert "Val!C35" in found and "Val!C36" not in found, found
    print(f"sum rows: ok (two blocks' present values added up, and a terminal value discounted on its own: each one "
          f"discounting at Inputs!C5 from Inputs!C4, not the labelled date the factors don't read; the figure recomputed "
          f"through IF and /thousand; the units constant not held)")


def sumif_check() -> None:
    """A value made, at each end, of a terminal value discounted by a factor row (SUMPRODUCT), whose factor on the
    valuation date's own period is 1 (IF(date < valuation date, 0, 1/(1+r)^YEARFRAC(valuation date, date, 3))), and
    the present values summed up to the terminal value date (SUMIF(dates, "<"&end+1, present values) at the low,
    "<="&end at the high), the cash flows going on past it. Both are read exactly: the factor of 1 has no cash flow
    against it (where it has, the fit doesn't take it), and the sum's cut-off is the end date's cell, read on every
    feed. So the formulas above them recompute, and the bridge splits into time value, cash flows and forecast."""
    import dcf
    import result
    import rodb
    import valuation
    vd, end, rates, n = date(2025, 6, 30), date(2028, 6, 30), {"low": 0.105, "high": 0.095}, 17
    ends, y, m = [vd], 2025, 9
    for _ in range(n - 1):
        ends.append(date(y + (m == 12), m % 12 + 1, 1) - timedelta(days=1))
        y, m = (y + 1, 3) if m == 12 else (y, m + 3)
    flows = [0.0] + [100.0 + 5 * k for k in range(1, n)]
    tv = [8 * x if e == end else 0.0 for x, e in zip(flows, ends)]
    ser = lambda d: (d - date(1899, 12, 30)).days
    out = Path(tempfile.mkdtemp(prefix="trace_sumif_"))
    wb = xlsxwriter.Workbook(out / "sumif.xlsx")
    dt = wb.add_format({"num_format": "dd-mmm-yy"})
    inp = wb.add_worksheet("Inputs")
    for r, label, v in ((3, "Valuation date", vd), (4, "Discount rate (low value)", rates["low"]),
                        (5, "Discount rate (high value)", rates["high"]), (6, "Terminal value date", end)):
        inp.write(r, 1, label)
        if isinstance(v, date):
            inp.write_datetime(r, 2, v, dt)
        else:
            inp.write_number(r, 2, v)
    va = wb.add_worksheet("Val")
    va.write(2, 1, "Period ending")
    va.write(4, 1, "Cash flow")
    va.write(6, 1, "Terminal value date")
    va.write_formula("C7", "=Inputs!$C$7", dt, ser(end))
    last, figure = COL(3 + n - 1), {}
    for k, e in enumerate(ends):
        va.write_datetime(2, 3 + k, e, dt)
        va.write_number(4, 3 + k, flows[k])
    for (end_, crit), top in ((("low", '"<"&$C$7+1'), 8), (("high", '"<="&$C$7'), 14)):
        rate, rc = rates[end_], "$C$5" if end_ == "low" else "$C$6"
        f = [0.0 if e < vd else 1 / (1 + rate) ** ((e - vd).days / 365) for e in ends]
        for r, label in ((top, "Discount factor"), (top + 1, "Terminal value"), (top + 2, "Equity DCF"),
                         (top + 3, "Total NPV"), (top + 4, "Equity value")):
            va.write(r - 1, 1, f"{label} ({end_})")
        for k in range(n):
            c = COL(3 + k)
            va.write_formula(f"{c}{top}", f"=IF({c}$3<Inputs!$C$4,0,1/(1+Inputs!{rc})^YEARFRAC(Inputs!$C$4,{c}$3,3))", None, f[k])
            va.write_number(top, 3 + k, tv[k])
            va.write_formula(f"{c}{top + 2}", f"={c}{top}*{c}5", None, f[k] * flows[k])
        sp = sum(a * b for a, b in zip(f, tv))
        si = sum(f[k] * flows[k] for k in range(n) if ends[k] <= end)
        va.write_formula(f"C{top + 1}", f"=SUMPRODUCT(D{top}:{last}{top},D{top + 1}:{last}{top + 1})", None, sp)
        va.write_formula(f"C{top + 2}", f"=SUMIF($D$3:${last}$3,{crit},$D${top + 2}:${last}${top + 2})", None, si)
        va.write_formula(f"C{top + 3}", f"=SUM(C{top + 1}:C{top + 2})", None, sp + si)
        va.write_formula(f"C{top + 4}", f"=C{top + 3}/1000", None, (sp + si) / 1000)
        figure[end_] = (f"Val!C{top + 4}", (sp + si) / 1000, f)
    wb.close()
    path = build_map.main(str(out / "sumif.xlsx"), str(out / "db"))["db"]
    db = sqlite3.connect(path)
    later = date(2029, 3, 31)  # the end date moved, as on this year's feed
    moved = rodb.patched(path, {("Val", 7, 3): later.isoformat()})
    for end_, (cell, v0, f) in figure.items():
        t = dcftrace.trace(db, cell)
        every = dcftrace.cores(t)
        cs = {c["kind"]: c for c in every if c.get("inputs")}
        assert len(every) == 2 and set(cs) == {"sumproduct", "pv row"}, [(c["cell"], c["kind"], c.get("inputs")) for c in every]
        for c in cs.values():
            i = c["inputs"]
            assert (i["rate"], i["valuation_date"], i["timing"], i["day_count"]) == (
                f"Inputs!C{5 if end_ == 'low' else 6}", "Inputs!C4", "end", "actual/365"), (end_, c["cell"], i)
            r = dcf.compute(db, **i, fix=False)
            assert dcf._close(r["total"], r["compare_to"]), (end_, c["cell"], r["total"], r["compare_to"])
        assert cs["sumproduct"]["inputs"]["terminal_date"] is None
        assert cs["pv row"]["inputs"]["terminal_date"] == "Val!C7" and cs["pv row"]["last_period"] == end.isoformat()
        got = dcf.compute(moved, **{**cs["pv row"]["inputs"], "compare_to": None}, fix=False)["total"]
        want = sum(f[k] * flows[k] for k in range(n) if ends[k] <= later)
        assert abs(got - want) < 1e-9, (end_, got, want)
        totals = {c["cell"]: dcf.compute(db, **{**c["inputs"], "compare_to": None}, fix=False)["total"] for c in every}
        assert abs(dcftrace.recompute(db, t, totals) - v0) < 1e-9
        steps, note = result._steps(db, (t, list(cs.values())), v0, 1.05 * v0, "2026-06-30")
        assert note is None and [s["key"] for s in steps] == ["time", "cash", "forecast"], (steps, note)
        paid = -sum(x * (1 + rates[end_]) ** ((date(2026, 6, 30) - e).days / 365) for x, e in zip(flows, ends)
                    if vd < e <= date(2026, 6, 30)) / 1000
        assert abs(steps[1]["value"] - paid) < 1e-9, (steps[1], paid)
    # a factor of 1 on the valuation date's own period with a cash flow against it: dcf leaves that out, so no fit
    cols = list(range(4, 4 + n))
    assert valuation.read_factors(db, ("Val", 8), cols, flows={c: 0.0 for c in cols})
    assert valuation.read_factors(db, ("Val", 8), cols, flows={4: 5.0}) is None
    assert valuation.read_factors(db, ("Val", 8), cols) is None
    print(f"sumif: ok (a terminal value discounted by factors of 1 on the valuation date, and present values summed to "
          f"{end:%d %b %Y} by SUMIF (\"<\"&end+1 and \"<=\"&end): both read exactly, the cut-off read from Val!C7 "
          f"(moved to {later:%d %b %Y}, it moves), the bridge split into time value, cash flows and forecast)")


def multiple_check() -> None:
    """Where the report's terminal value is an exit multiple: the multiple is the cell the terminal value's formula
    reads in M x X, the terminal value recomputed from it, the metric X what the report says it's a multiple of, the
    report's multiple tied; typed into the formula, it isn't sourced; a metric that isn't the report's is flagged."""
    import sourced
    vd, rate, mult, years = date(2025, 6, 30), 0.08, 12.0, list(range(2026, 2031))
    got = {}
    for how in ("cell", "typed"):
        out = Path(tempfile.mkdtemp(prefix=f"trace_mult_{how}_"))
        wb = xlsxwriter.Workbook(out / f"{how}.xlsx")
        dt = wb.add_format({"num_format": "dd-mmm-yy"})
        inp = wb.add_worksheet("Inputs")
        inp.write(3, 0, "Valuation date")
        inp.write_datetime(3, 2, vd, dt)
        inp.write(4, 0, "Discount rate")
        inp.write_number(4, 2, rate)
        inp.write(6, 0, "Exit multiple (EV/EBITDA)")
        inp.write_number(6, 2, mult)
        fl = wb.add_worksheet("Flows")
        for r, label in ((2, "Period ending"), (6, "EBITDA"), (7, "Free cash flow"), (8, "Terminal value"),
                         (9, "Valuation cash flow"), (12, "Discount factor"), (13, "Present value"), (15, "Equity value")):
            fl.write(r, 1, label)
        total, last = 0.0, COL(3 + len(years) - 1)
        for k, y in enumerate(years):
            c, end = COL(3 + k), date(y, 6, 30)
            f, ebitda, cf = 1 / (1 + rate) ** ((end - vd).days / 365), 150.0 + 6 * k, 100.0 + 5 * k
            tv = mult * ebitda if k == len(years) - 1 else 0.0
            fl.write_datetime(2, 3 + k, end, dt)
            fl.write_number(6, 3 + k, ebitda)
            fl.write_number(7, 3 + k, cf)
            if tv:
                fl.write_formula(f"{c}9", f"=Inputs!$C$7*{c}7" if how == "cell" else f"={c}7*12", None, tv)
            else:
                fl.write_number(8, 3 + k, 0.0)
            fl.write_formula(f"{c}10", f"={c}8+{c}9", None, cf + tv)
            fl.write_formula(f"{c}13", f"=1/(1+Inputs!$C$5)^(({c}3-Inputs!$C$4)/365)", None, f)
            fl.write_formula(f"{c}14", f"={c}10*{c}13", None, (cf + tv) * f)
            total += (cf + tv) * f
        fl.write_formula("C16", f"=SUM(D14:{last}14)", None, total)
        wb.close()
        db = sqlite3.connect(build_map.main(str(out / f"{how}.xlsx"), str(out / "db"))["db"])
        cs = [c for c in dcftrace.cores(dcftrace.trace(db, "Flows!C16")) if c.get("inputs")]
        facts = [{"key": "terminal_multiple", "value_text": "12.0x", "basis": "EV/EBITDA", "status": "approved"}]
        traced = {"low": cs, "high": cs}
        got[how] = sourced.multiple(db, traced, facts, {"kind": "exit_ebitda", "label": "An exit multiple of EBITDA"})["ends"]["low"]
        if how == "cell":  # the report says EV/RAB, the model multiplies EBITDA
            got["rab"] = sourced.multiple(db, traced, facts, {"kind": "exit_rab", "label": "An exit multiple of the RAB"})["ends"]["low"]
    cell, typed, rab = got["cell"], got["typed"], got["rab"]
    assert cell["sourced"] and cell["cell"] == "Inputs!C7" and cell["value"] == mult and cell["ok"] and cell["ties"], cell
    assert cell["metric"]["label"] == "EBITDA" and [c["ok"] for c in cell["checks"]] == [True, True, True, True], cell
    assert not typed["sourced"] and typed["cell"] is None and typed["value"] == mult and not typed["ok"], typed
    assert not rab["ok"] and "The report says a multiple of the RAB; the metric here is EBITDA" in \
        [c["text"] for c in rab["checks"]], rab
    print("multiple: ok (the exit multiple is the cell the terminal value's formula reads in M × X, recomputed; the "
          "metric is what the report says; typed into the formula, it isn't sourced)")


if __name__ == "__main__":
    main()
    mid_year_check()
    sourcing_check()
    growth_check()
    roll_dates_check()
    cutoff_check()
    masked_check()
    horizon_check()
    window_check()
    sum_rows_check()
    sumif_check()
    multiple_check()
