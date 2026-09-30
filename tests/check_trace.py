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


if __name__ == "__main__":
    main()
    mid_year_check()
    sourcing_check()
    growth_check()
