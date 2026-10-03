"""Reading the models: the shapes a workbook can come in, and what the app makes of each.

Every case keeps to one rule: the app takes the right cells and value, or it says what's wrong (a need, the value held,
the file refused with a reason); never a different figure in silence, never a crash.

  references   a number in scientific notation (1E9, 2.5E-3) isn't a cell, nor a function's name (LOG10() however
               like a cell it looks; a figure typed into a formula is read with its exponent
  choose       a CHOOSE whose selector is typed in (a scenario's) reads the case it picks, not the others
  beyond       a client figure right of the timeline (the model's own terminal value) read from as far right of this
               year's timeline only where that column is headed as last year's was; another column there holds
  saved state  what the sheet XML says that the values reader can't: formulas saved with no result (a workbook saved
               without being calculated; not a formula whose result is empty text), Excel's own errors, the last
               calculation's settings (one that didn't finish), and every cell however small the size the sheet
               declares (a tool other than Excel can write it too small)
  containers   an encrypted workbook (a compound file, not a zip), a file that isn't a workbook, a document renamed
               and a workbook saved as Strict Open XML are refused with what to do
  the pack     the synthetic pack run end to end with one file changed, as a model can arrive: protected sheets, a
               macro-enabled workbook, units that start with # ("#/Day"), Excel's errors where the value doesn't read
               them, manual calculation, a data table: the same value. Formulas this year's or last year's model saved
               no result for, the overlay's own, a calculation that didn't finish, Excel's errors under the value, an
               iterative loop under the value, an encrypted file: each held or refused, saying so

The pack is changed at the XML level, as Excel or another tool saves a file: openpyxl load + save would drop every
formula's saved result.
    uv run python tests/make_pack.py && uv run python tests/check_ingest.py
"""
import json
import re
import shutil
import sys
import tempfile
import time
import zipfile
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "engine"))
sys.path.insert(0, str(ROOT / "tests"))
PACK = ROOT / "tests" / "pack"
REPORT, PRIOR, OVERLAY, CURRENT = ("AssetA_valuation_report_FY25.pdf", "AssetA_BP25_client_model_Jun25.xlsx",
                                   "Alpha_valuation_overlay_FY25.xlsx", "AssetA_BP26_client_model_Jun26.xlsx")
SLIDES, WITH_OVERLAY = "AssetA_valuation_report_FY25.pptx", "AssetA_BP25_with_overlay.xlsx"  # the overlay inside a copy

import xlsxwriter  # noqa: E402

import build_map  # noqa: E402
import dcf  # noqa: E402
import sourced  # noqa: E402


# ---- a workbook changed as a tool saves it ---------------------------------------------------------------------------

def parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def save(path: Path, ps: dict[str, bytes]) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in ps.items():
            z.writestr(n, b)
    tmp.replace(path)


def sheet_part(ps: dict, sheet: str) -> str:
    wbx, rels = ps["xl/workbook.xml"].decode(), ps["xl/_rels/workbook.xml.rels"].decode()
    rid = re.search(r'<sheet [^>]*name="%s"[^>]*r:id="(rId\d+)"' % re.escape(sheet), wbx)[1]
    target = (re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels)
              or re.search(r'Target="([^"]+)"[^>]*Id="%s"' % rid, rels))[1].lstrip("/")
    return target if target.startswith("xl/") else "xl/" + target


def cell_xml(path: Path, sheet: str, fn) -> None:
    ps = parts(path)
    sp = sheet_part(ps, sheet)
    ps[sp] = fn(ps[sp].decode()).encode()
    save(path, ps)


def strip_cached(path: Path, sheet: str, cells: list[str]) -> None:
    """Formula cells saved with no result (a workbook saved without being calculated)."""
    def fn(x):
        for a in cells:
            x, n = re.subn(r'(<c r="%s"[^>]*>\s*<f>[^<]*</f>)\s*<v>[^<]*</v>' % a, r"\1", x)
            assert n == 1, (sheet, a)
        return x
    cell_xml(path, sheet, fn)


def set_error(path: Path, sheet: str, cells: list[str], err: str) -> None:
    """Formula cells whose saved result is one of Excel's errors."""
    def fn(x):
        for a in cells:
            x, n = re.subn(r'<c r="%s"([^>]*)>(\s*<f>[^<]*</f>)\s*<v>[^<]*</v>' % a,
                           lambda m: f'<c r="{a}"{m[1]} t="e">{m[2]}<v>{err}</v>', x)
            assert n == 1, (sheet, a)
        return x
    cell_xml(path, sheet, fn)


def set_text(path: Path, sheet: str, cells: dict[str, str]) -> None:
    """Cells replaced by text (inline strings)."""
    def fn(x):
        for a, t in cells.items():
            x, n = re.subn(r'<c r="%s"[^>]*?(/>|>.*?</c>)' % a, f'<c r="{a}" t="inlineStr"><is><t>{t}</t></is></c>', x,
                           flags=re.S)
            assert n == 1, (sheet, a)
        return x
    cell_xml(path, sheet, fn)


def set_calc(path: Path, **attrs) -> None:
    """The workbook's calculation settings (<calcPr>)."""
    ps = parts(path)
    x = ps["xl/workbook.xml"].decode()
    tag = re.search(r"<calcPr[^>]*/>", x)
    new = "<calcPr " + " ".join(f'{k}="{v}"' for k, v in attrs.items()) + "/>"
    ps["xl/workbook.xml"] = (x.replace(tag[0], new) if tag else x.replace("</workbook>", new + "</workbook>")).encode()
    save(path, ps)


OLE = bytes.fromhex("D0CF11E0A1B11AE1")  # a compound file's signature: an encrypted package, or the old .xls format


# ---- the checks on their own ------------------------------------------------------------------------------------------

def references_check() -> None:
    cases = {"IF(C9>1E9,0.001*C9,0)": ["C9", "C9"], "A1*2.5E3+E9": ["A1", "E9"], "SUM(D5:W5)*1E-3": ["D5:W5"],
             "Sheet1!E9+1e9": ["Sheet1!E9"], "LOG10(1E+9)": [], "'Val (2)'!$A$1+[1]CashFlow!D9": ["'Val (2)'!$A$1",
                                                                                                    "CashFlow!D9"],
             "1/(1+Val_Inputs!$C$5)^YEARFRAC(Val_Inputs!$C$4,D$3,1)": ["Val_Inputs!$C$5", "Val_Inputs!$C$4", "D$3"],
             "FY2025*2": ["FY2025"]}
    for f, want in cases.items():
        got = [m[0] for m in dcf._FREF.finditer(f)]
        assert got == want, (f, got, want)
    typed = lambda f: [float(m[1]) for m in sourced._LIT.finditer(dcf._FREF.sub(" ", f))]
    assert typed("=C5*1E-2") == [0.01] and typed("=B2*2.5e+3") == [2500.0] and typed("=1/(1+0.0775)^D3") == [1, 1, 0.0775]
    print("references: ok (1E9 and 2.5E-3 are numbers, not cells; LOG10( is a function; a typed 1E-2 is 0.01)")


def choose_check() -> None:
    """A CHOOSE whose selector is typed in (a scenario's) reads the case it picks: the tracer and the walk the
    inputs are sourced by follow that one, not the others (a downside case's rate holding the same figure)."""
    import sqlite3
    import dcftrace
    import valuation
    out = Path(tempfile.mkdtemp(prefix="ingest_"))
    p = out / "AssetA_cases.xlsx"
    wbk = xlsxwriter.Workbook(p)
    ws = wbk.add_worksheet("Inputs")
    for r, (label, a, b, c) in enumerate((("Discount rate", 0.0775, 0.0775, 0.07), ("Growth", 0.025, 0.02, 0.03)), 3):
        ws.write(r, 0, label)
        for k, v in enumerate((a, b, c)):
            ws.write_number(r, 2 + k, v)
    ws.write(1, 0, "Case (1 base, 2 down, 3 up)")
    ws.write_number(1, 2, 1)
    ws.write(7, 0, "Discount rate (live)")
    ws.write_formula(7, 2, "=CHOOSE($C$2,C4,D4,E4)", None, 0.0775)
    ws.write(8, 0, "Case worked out")
    ws.write_formula(8, 2, "=CHOOSE(1+0*C4,C5,D5,E5)", None, 0.025)
    wbk.close()
    db = sqlite3.connect(build_map.main(str(p), str(out / "db"))["db"])
    assert dcftrace._chosen(db, "=CHOOSE($C$2,C4,D4,E4)*2", "Inputs") == "=(C4)*2"
    assert dcftrace._chosen(db, "=CHOOSE(1+0*C4,C5,D5,E5)", "Inputs") == "=CHOOSE(1+0*C4,C5,D5,E5)"  # worked out: whole
    got = valuation.reads(db, cells=[("Inputs", 8, 3)])
    assert "Inputs!C4" in got and "Inputs!D4" not in got and "Inputs!E4" not in got, sorted(got)
    print("choose: ok (a typed selector's CHOOSE read as the case it picks; one worked out by a formula left whole)")


def beyond_check() -> None:
    """A client figure right of last year's timeline (the model's own terminal value, two columns past its last period),
    read by the overlay: this year it's read as far right of this year's timeline where that column is headed as last
    year's was, and where this year's model has another column there (a check inserted before it), last year's figure
    stands in and the value holds, never another column's figure in silence."""
    import overlay as ov
    import result
    import xlcompile
    from xlsxwriter.utility import xl_col_to_name as col_
    vd, new, rate = date(2025, 6, 30), date(2026, 6, 30), 0.08

    def run(inserted: bool):
        out = Path(tempfile.mkdtemp(prefix="beyond_"))
        paths = {}
        for which in ("prior", "current"):
            wbk = xlsxwriter.Workbook(out / f"{which}.xlsx")
            dt = wbk.add_format({"num_format": "dd-mmm-yy"})
            cl = wbk.add_worksheet("Client")
            cl.write(2, 1, "Period ending")
            cl.write(9, 1, "Cash flow")
            y0 = 2026 if which == "prior" else 2027
            for k in range(6):
                cl.write_datetime(2, 3 + k, date(y0 + k, 6, 30), dt)
                cl.write_number(9, 3 + k, 100.0 + 10 * k + (5 if which == "current" else 0))
            tv_col = 10 + (1 if inserted and which == "current" else 0)  # K, or L past a check column in K
            if inserted and which == "current":
                cl.write(2, 10, "Check")
                cl.write_number(9, 10, 1.0)
            cl.write(2, tv_col, "Terminal value")
            cl.write_number(9, tv_col, 1000.0 if which == "prior" else 1100.0)
            if which == "prior":
                va = wbk.add_worksheet("Val")
                va.write(3, 1, "Valuation date")
                va.write_datetime(3, 2, vd, dt)
                va.write(4, 1, "Discount rate")
                va.write_number(4, 2, rate)
                va.write(5, 1, "Period ending")
                va.write(9, 1, "Cash flow")
                va.write(12, 1, "Discount factor")
                va.write(13, 1, "Present value")
                total = 0.0
                for k in range(6):
                    c, e = col_(3 + k), date(2026 + k, 6, 30)
                    f = 1 / (1 + rate) ** ((e - vd).days / 365)
                    va.write_formula(f"{c}6", f"=Client!{c}3", dt, (e - date(1899, 12, 30)).days)
                    va.write_formula(f"{c}10", f"=Client!{c}10", None, 100.0 + 10 * k)
                    va.write_formula(f"{c}13", f"=1/(1+$C$5)^(({c}6-$C$4)/365)", None, f)
                    va.write_formula(f"{c}14", f"={c}10*{c}13", None, (100.0 + 10 * k) * f)
                    total += (100.0 + 10 * k) * f
                last_f = 1 / (1 + rate) ** ((date(2031, 6, 30) - vd).days / 365)
                va.write(17, 1, "Equity value")
                va.write_formula("C18", "=SUM(D14:I14)+Client!$K$10*I13", None, total + 1000.0 * last_f)
            wbk.close()
            paths[which] = build_map.main(str(out / f"{which}.xlsx"), str(out / f"db_{which}"))["db"]
        db = paths["prior"]
        src, _ = xlcompile.compile_overlay(db, ["Val"])
        (out / "overlay.py").write_text(src)
        sess = ov.Session(str(out / "overlay.py"), db, ["Val"], None, paths["current"], None, ["Client"])
        roll = ov.plan_roll(sess, None, {"sheets": ["Val"]}, True, vd.isoformat(), None, new.isoformat())
        roll.update(ov.date_cells(db, [{"cell": "Val!C18"}], ["Val"], None))
        summary = {"wiring": {"overlay": {"db_path": db}, "current": {"db_path": paths["current"]}}, "sheets": ["Val"],
                   "roll": roll, "held_values": {}, "outputs": [{"cell": "Val!C18", "label": "Equity value"}],
                   "balance_decisions": {}}
        figs = ov.deep(result.figures, sess, summary, ["Val!C18"])
        gate = result._gate_holds(summary, {"texts": {}}, {"low": "Val!C18", "high": "Val!C18", "scale": 1, "sign": 1},
                                  figs, lambda x: x)
        return figs, gate
    figs, gate = run(False)
    assert not figs["feed"]["beyond_stood"] and "beyond-standin" not in [h["id"] for h in gate], figs["feed"]
    moved = figs["this_year"]["Val!C18"]
    figs2, gate2 = run(True)
    assert [x["cell"] for x in figs2["feed"]["beyond_stood"]] == ["Client!K10"], figs2["feed"]
    assert ("beyond-standin", "block") in [(h["id"], h["severity"]) for h in gate2], gate2
    assert abs(figs2["this_year"]["Val!C18"] - moved) > 1.0  # (last year's 1,000 stands in for this year's 1,100)
    print("beyond: ok (a figure right of the timeline read where this year's column is headed as last year's; another "
          "column there holds the value, last year's figure standing in, said so)")


def saved_state_check() -> None:
    out = Path(tempfile.mkdtemp(prefix="ingest_"))
    p = out / "AssetA_small.xlsx"
    wbk = xlsxwriter.Workbook(p)
    dt = wbk.add_format({"num_format": "dd-mmm-yy"})
    ws = wbk.add_worksheet("Calc")
    ws.write(2, 0, "Period ending")
    for k in range(4):
        ws.write_datetime(2, 2 + k, date(2026 + k, 6, 30), dt)
    for r, label in ((4, "Revenue"), (5, "Costs"), (6, "Margin"), (7, "Check")):
        ws.write(r, 0, label)
        ws.write(r, 1, "#" if label == "Check" else "A$m")  # a unit that starts with #, as a count's does
    for k in range(4):
        c = chr(ord("C") + k)
        ws.write_number(4, 2 + k, 100.0 + k)
        ws.write_number(5, 2 + k, -40.0)
        ws.write_formula(6, 2 + k, f"={c}5+{c}6", None, 60.0 + k)
        ws.write_formula(7, 2 + k, f'=IF({c}7>0,"",1)', None, "")  # a formula whose result is empty text
    wbk.close()
    # empty text saved as Excel saves it (t="str"; xlsxwriter leaves the type off, as openpyxl does for a formula it
    # never calculated, which is what an untyped empty result is taken for)
    cell_xml(p, "Calc", lambda x: re.sub(r'<c r="([C-F]8)"', r'<c r="\1" t="str"', x))
    strip_cached(p, "Calc", ["D7", "E7"])  # two never calculated
    set_error(p, "Calc", ["F7"], "#N/A")
    set_calc(p, calcId="191029", calcMode="manual", calcCompleted="0")
    # a cell beyond the size the sheet declares (a tool other than Excel wrote <dimension> too small)
    cell_xml(p, "Calc", lambda x: re.sub(r'(<row r="7"[^>]*>.*?)(</row>)',
                                         lambda m: m[1] + '<c r="J7"><v>7</v></c>' + m[2], x, count=1, flags=re.S))
    db_path = build_map.main(str(p), str(out / "db"))["db"]
    import sqlite3
    db = sqlite3.connect(db_path)
    assert sorted(db.execute("SELECT sheet, row, col FROM unsaved")) == [("Calc", 7, 4), ("Calc", 7, 5)]
    meta = dict(db.execute("SELECT key, value FROM meta"))
    assert meta.get("calc.calcCompleted") == "0" and meta.get("calc.calcMode") == "manual", meta
    v = dict(((r, c), x) for r, c, x in db.execute("SELECT row, col, value FROM cells WHERE sheet='Calc'"))
    assert v[(7, 6)] == "#N/A" and v[(7, 10)] == 7 and v[(8, 2)] == "#", v
    assert v.get((8, 3)) is None and db.execute("SELECT formula FROM cells WHERE sheet='Calc' AND row=8 AND col=3")\
        .fetchone()[0] == '=IF(C7>0,"",1)'  # empty text is a result, not "unsaved"
    print("saved state: ok (formulas saved with no result recorded, not empty text; Excel's errors; a calculation that "
          "didn't finish; a cell beyond the declared size read; a unit '#' kept as text)")


def containers_check() -> None:
    out = Path(tempfile.mkdtemp(prefix="ingest_"))
    def zipped(files: dict) -> bytes:
        import io
        b = io.BytesIO()
        with zipfile.ZipFile(b, "w") as z:
            for n, x in files.items():
                z.writestr(n, x)
        return b.getvalue()
    strict = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://purl.oclc.org/ooxml/'
              'spreadsheetml/main"><sheets/></workbook>')
    for name, body, said in (("AssetA_locked.xlsx", OLE + b"\0" * 504, "password"),
                             ("AssetA_not_a_workbook.xlsx", b"just some text", "isn't an Excel workbook"),
                             ("AssetA_a_document.xlsx", zipped({"word/document.xml": "<w/>"}), "no workbook inside"),
                             ("AssetA_strict.xlsx", zipped({"xl/workbook.xml": strict}), "Strict Open XML")):
        (out / name).write_bytes(body)
        try:
            build_map.main(str(out / name), str(out / name.replace(".xlsx", "")))
            raise AssertionError(f"{name} was read")
        except ValueError as e:
            assert said in str(e) and name in str(e), e
    print("containers: ok (an encrypted workbook, a file that isn't one, a document renamed and a Strict Open XML workbook "
          "are refused, saying what to do)")


# ---- the pack, one thing changed --------------------------------------------------------------------------------------

def _xlsm(d: Path) -> Path:
    """The overlay as a macro-enabled workbook: a VBA project part, the macro-enabled content type."""
    ps = parts(d / OVERLAY)
    ps["xl/vbaProject.bin"] = OLE + b"\0" * 1016
    ct = ps["[Content_Types].xml"].decode().replace(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
    ps["[Content_Types].xml"] = ct.replace(
        "</Types>", '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>').encode()
    ps["xl/_rels/workbook.xml.rels"] = ps["xl/_rels/workbook.xml.rels"].decode().replace(
        "</Relationships>", '<Relationship Id="rId99" Type="http://schemas.microsoft.com/office/2006/relationships/'
                            'vbaProject" Target="vbaProject.bin"/></Relationships>').encode()
    new = d / OVERLAY.replace(".xlsx", ".xlsm")
    save(new, ps)
    (d / OVERLAY).unlink()
    return new


def _protect(d: Path) -> None:
    ps = parts(d / OVERLAY)
    for n in [n for n in ps if n.startswith("xl/worksheets/sheet")]:
        ps[n] = re.sub(r"(</sheetData>)", r'\1<sheetProtection sheet="1" objects="1" scenarios="1"/>',
                       ps[n].decode(), count=1).encode()
    ps["xl/workbook.xml"] = re.sub(r"(<bookViews>)", r'<workbookProtection lockStructure="1"/>\1',
                                   ps["xl/workbook.xml"].decode(), count=1).encode()
    save(d / OVERLAY, ps)


def _loop(d: Path) -> None:
    """A loop under the low end (a fee on the equity value, nil at these figures), iterative calculation on."""
    def fn(x):
        x, n = re.subn(r'(<c r="C9"[^>]*>\s*<f>)C7\+C8(</f>)', r"\1C7+C8-G9\2", x)
        assert n == 1
        return re.sub(r'(<c r="E9"[^>]*>.*?</c>)', r'\1<c r="G9"><f>IF(C9&gt;1E9,0.001*C9,0)</f><v>0</v></c>', x,
                      count=1, flags=re.S)
    cell_xml(d / OVERLAY, "Summary", fn)
    set_calc(d / OVERLAY, calcId="191029", iterate="1", iterateCount="100", iterateDelta="0.001")


def _data_table(d: Path) -> None:
    """A one-way data table beside the inputs, its results as Excel last saved them, not recalculated (autoNoTable)."""
    def fn(x):
        cells = ('<c r="H4"><f t="dataTable" ref="H4:H6" dt2D="0" dtr="0" r1="C6"/><v>2400</v></c>',
                 '<c r="H5"><v>2500</v></c>', '<c r="H6"><v>2600</v></c>')
        for r, c in zip((4, 5, 6), cells):
            x = re.sub(r'(<row r="%d"[^>]*>.*?)(</row>)' % r, lambda m: m[1] + c + m[2], x, count=1, flags=re.S)
        return x
    cell_xml(d / OVERLAY, "Val_Inputs", fn)
    set_calc(d / OVERLAY, calcId="191029", calcMode="autoNoTable")


def cached(x: str, a: str) -> float:
    """A cell's saved value in a sheet's XML."""
    return float(re.search(r'<c r="%s"[^>]*>(?:\s*<f>[^<]*</f>)?\s*<v>([^<]*)</v>' % a, x)[1])


def set_cell(x: str, a: str, formula: str, value: float) -> str:
    """A formula cell's formula and saved value, in place."""
    x, n = re.subn(r'(<c r="%s"[^>]*>)\s*<f>[^<]*</f>\s*<v>[^<]*</v>' % a,
                   lambda m: f"{m[1]}<f>{formula}</f><v>{value!r}</v>", x)
    assert n == 1, a
    return x


def add_rows(x: str, rows: dict[int, list[tuple]]) -> str:
    """New rows at the end of a sheet: {row: [(addr, text) or (addr, formula, value)]}."""
    def cell(t):
        if len(t) == 2:
            return f'<c r="{t[0]}" t="inlineStr"><is><t>{t[1]}</t></is></c>'
        return f'<c r="{t[0]}"><f>{t[1]}</f><v>{t[2]!r}</v></c>'
    new = "".join(f'<row r="{r}">' + "".join(cell(t) for t in cs) + "</row>" for r, cs in sorted(rows.items()))
    return x.replace("</sheetData>", new + "</sheetData>", 1)


def _tv_after(d: Path, own_cell: bool) -> None:
    """The terminal value added after the discounting, as many overlays have it: the last cash flow without it, the
    enterprise value the present values' sum plus the terminal value times the last factor (in its formula, or in a
    cell of its own beside the present values). The same value."""
    ps = parts(d / OVERLAY)
    sp = sheet_part(ps, "DCF")
    x = ps[sp].decode()
    pv = {}
    for tv_r, vcf_r, df_r, pv_r in ((10, 11, 12, 13), (17, 18, 19, 20)):
        for c in "DEFGHIJKLMNOPQRSTUVW":  # the cash flows without the terminal value, every period
            x = set_cell(x, f"{c}{vcf_r}", f"{c}5", cached(x, f"{c}5"))
        f, w = cached(x, "W5"), cached(x, f"W{df_r}")
        x = set_cell(x, f"W{pv_r}", f"W{vcf_r}*W{df_r}", f * w)
        pv[pv_r] = cached(x, f"W{tv_r}") * w
    if own_cell:  # its present value in the column right of the present values
        for tv_r, df_r, pv_r in ((10, 12, 13), (17, 19, 20)):
            x = re.sub(r'(<row r="%d"[^>]*>.*?)(</row>)' % pv_r,
                       lambda m: m[1] + f'<c r="Y{pv_r}"><f>W{tv_r}*W{df_r}</f><v>{pv[pv_r]!r}</v></c>' + m[2], x,
                       count=1, flags=re.S)
    ps[sp] = x.encode()
    sm = sheet_part(ps, "Summary")
    y = ps[sm].decode()
    for c, tv_r, df_r, pv_r in (("C", 10, 12, 13), ("E", 17, 19, 20)):
        tail = f"DCF!Y{pv_r}" if own_cell else f"DCF!W{tv_r}*DCF!W{df_r}"
        y = set_cell(y, f"{c}4", f"SUM(DCF!D{pv_r}:W{pv_r})+{tail}", cached(y, f"{c}4"))
    ps[sm] = y.encode()
    save(d / OVERLAY, ps)


def _equity_twice(d: Path, copy: bool) -> None:
    """A second row holding the report's low and high: a report table carrying the summary's cells (=C9), or a row
    working them out again (=C7+C8)."""
    def fn(x):
        lo, hi = cached(x, "C9"), cached(x, "E9")
        f = ("C9", "E9") if copy else ("C7+C8", "E7+E8")
        return add_rows(x, {13: [("A13", "Equity value (ex-distribution)"), ("B13", "A$m"), ("C13", f[0], lo),
                                 ("E13", f[1], hi)]})
    cell_xml(d / OVERLAY, "Summary", fn)


COLS = "DEFGHIJKLMNOPQRSTUVW"  # the pack's periods


def _pasted(d: Path, off: bool) -> None:
    """This year's tax and free cash flow typed in as last year's figures (the formulas pasted over as values): in
    place, each period last year's for that period (the last, which last year's model hasn't, left as it was), or
    one period off, each column last year's same column."""
    last = parts(d / PRIOR)[sheet_part(parts(d / PRIOR), "CashFlow")].decode()

    def fn(x):
        for r in (8, 9):
            for i, c in enumerate(COLS):
                src = c if off else (COLS[i + 1] if i + 1 < len(COLS) else None)
                if src is None:
                    continue
                v = cached(last, f"{src}{r}")
                x, n = re.subn(r'<c r="%s%d"([^>]*)>\s*<f>[^<]*</f>\s*<v>[^<]*</v>\s*</c>' % (c, r),
                               lambda m: f'<c r="{c}{r}"{m[1]}><v>{v!r}</v></c>', x)
                assert n == 1, (c, r)
        return x
    cell_xml(d / CURRENT, "CashFlow", fn)


def _client_cpi_growth(d: Path) -> list[Path]:
    """The overlay inside a copy of the client model, its terminal growth rate the client's CPI (Inputs!B7: 2.5%
    last year, 3.0% this year). The copy's Inputs sheet is the client's, so this year's model's CPI is read."""
    cell_xml(d / WITH_OVERLAY, "Val_Inputs", lambda x: re.sub(
        r'<c r="C6"([^>]*)>\s*(?:<f>[^<]*</f>)?\s*<v>([^<]*)</v>\s*</c>', r'<c r="C6"\1><f>Inputs!$B$7</f><v>\2</v></c>', x,
        count=1))
    return [d / SLIDES, d / WITH_OVERLAY, d / CURRENT]


def _ref_under_value(d: Path) -> None:
    """A cell under the value whose reference was deleted (=#REF!, saved as the error; what it fed saved before)."""
    cell_xml(d / OVERLAY, "DCF", lambda x: re.sub(r'<c r="H5"([^>]*)>\s*<f>[^<]*</f>\s*<v>[^<]*</v>\s*</c>',
                                                  r'<c r="H5"\1 t="e"><f>#REF!</f><v>#REF!</v></c>', x, count=1))


def _timeline_times(d: Path) -> None:
    """This year's period dates saved with a time of day (30 June 23:59), as a model built from timestamps has them."""
    def fn(x):
        for c in COLS:
            v = cached(x, f"{c}3")
            x = re.sub(r'(<c r="%s3"[^>]*>(?:\s*<f>[^<]*</f>)?\s*<v>)[^<]*(</v>)' % c, lambda m: f"{m[1]}{v + 0.999}{m[2]}", x,
                       count=1)
        return x
    for sh in ("Operations", "CashFlow"):
        cell_xml(d / CURRENT, sh, fn)


def _decoy_reads_real(d: Path) -> None:
    """The equity value's row labelled otherwise (a fair market value), and a row labelled as the equity value that
    reads it (with an adjustment of nothing much): the row it reads is the one the model works out."""
    def fn(x):
        x = re.sub(r'<c r="A9"[^>]*?(?:/>|>.*?</c>)', '<c r="A9" t="inlineStr"><is><t>Fair market value (ex-distribution)'
                   '</t></is></c>', x, count=1, flags=re.S)
        lo, hi = cached(x, "C9"), cached(x, "E9")
        return add_rows(x, {13: [("A13", "Equity value"), ("B13", "A$m"), ("C13", "C9+0.04", lo + 0.04),
                                 ("E13", "E9+0.04", hi + 0.04)]})
    cell_xml(d / OVERLAY, "Summary", fn)


def _no_high(d: Path) -> None:
    """The equity value's high end gone from the overlay (its cell a note): the low and the mid only."""
    cell_xml(d / OVERLAY, "Summary", lambda x: re.sub(r'<c r="E9"[^>]*?(?:/>|>.*?</c>)',
                                                      '<c r="E9" t="inlineStr"><is><t>see the note</t></is></c>', x,
                                                      count=1, flags=re.S))


def _summary_formulas(d: Path, cells: dict[str, str], array: bool = False) -> None:
    """Summary cells written another way, their saved values as they were (an array formula: t="array")."""
    def fn(x):
        for a, f in cells.items():
            v = cached(x, a)
            tag = f'<f t="array" ref="{a}">' if array else "<f>"
            x, n = re.subn(r'(<c r="%s"[^>]*>)\s*<f>[^<]*</f>\s*<v>[^<]*</v>' % a,
                           lambda m: f"{m[1]}{tag}{f}</f><v>{v!r}</v>", x)
            assert n == 1, a
        return x
    cell_xml(d / OVERLAY, "Summary", fn)


def _first_days(d: Path) -> None:
    """This year's model dating each period by its first day (1 July 2026 for the year to 30 June 2027)."""
    from datetime import timedelta
    epoch = date(1899, 12, 30)

    def fn(x):
        for c in COLS:
            end = epoch + timedelta(days=int(cached(x, f"{c}3")))
            start = (date(end.year - 1, end.month, 1) + timedelta(days=32)).replace(day=1) if end.month == 12 else \
                date(end.year - 1, end.month + 1, 1)
            x = re.sub(r'(<c r="%s3"[^>]*>(?:\s*<f>[^<]*</f>)?\s*<v>)[^<]*(</v>)' % c,
                       lambda m: f"{m[1]}{float((start - epoch).days)!r}{m[2]}", x, count=1)
        return x
    for sh in ("Operations", "CashFlow"):
        cell_xml(d / CURRENT, sh, fn)


SAME = {}  # the same value and cells, and nothing new to look at
CASES = {
    "protected sheets": (lambda d: _protect(d), SAME),
    "a macro-enabled overlay": (lambda d: _xlsm(d), SAME),
    "units starting with #": (lambda d: (set_text(d / CURRENT, "Operations", {"B6": "#/Day", "B7": "#", "B8": "# 000s"}),
                                         set_text(d / PRIOR, "Operations", {"B6": "#/Day", "B7": "#", "B8": "# 000s"})), SAME),
    "errors the value doesn't read": (lambda d: (set_error(d / CURRENT, "Operations", ["D6", "E6", "F6"], "#N/A"),
                                                 set_error(d / CURRENT, "Operations", ["G7"], "#REF!")), SAME),
    "manual calculation": (lambda d: set_calc(d / CURRENT, calcId="191029", calcMode="manual", calcOnSave="0"), SAME),
    "a data table": (lambda d: _data_table(d), SAME),
    "this year's cash flows never calculated": (lambda d: strip_cached(d / CURRENT, "CashFlow", [f"{c}9" for c in "DEFGH"]),
                                                {"unsaved-current": "block"}),
    "this year's tax never calculated": (lambda d: strip_cached(d / CURRENT, "CashFlow", [f"{c}8" for c in "DEFGH"]),
                                         {"unsaved-current": "block"}),
    "last year's cash flows never calculated": (lambda d: strip_cached(d / PRIOR, "CashFlow", [f"{c}9" for c in "DEFGH"]),
                                                {"unsaved-prior": "check", "rebuild-low": "block"}),
    "the overlay's own never calculated": (lambda d: strip_cached(d / OVERLAY, "DCF", ["E13", "F13"]),
                                           {"unsaved-overlay": "check"}),
    "a calculation that didn't finish": (lambda d: set_calc(d / CURRENT, calcId="191029", calcCompleted="0"),
                                         {"calc-incomplete": "check", "": "same"}),
    "Excel's errors under the value": (lambda d: set_error(d / CURRENT, "CashFlow", ["E9", "F9", "G9"], "#N/A"),
                                       {"errors-current": "block"}),
    "an iterative loop under the value": (lambda d: _loop(d), {"circular": "block"}),
    "an encrypted client model": (lambda d: (d / CURRENT).write_bytes(OLE + b"\0" * 4088), {"files": "failed"}),
    "a report table carrying the equity cells": (lambda d: _equity_twice(d, True), SAME),
    "a second row working out the equity value": (lambda d: _equity_twice(d, False), {"equity": "block", "cells": False}),
    "the terminal value added after the discounting": (lambda d: _tv_after(d, False), SAME),
    "the terminal value's present value in a cell of its own": (lambda d: _tv_after(d, True), SAME),
    "last year's cash flows pasted in place": (lambda d: _pasted(d, False), {"pasted": "block"}),
    "last year's cash flows pasted one period off": (lambda d: _pasted(d, True), {"pasted": "block"}),
    "the client's CPI as the growth rate, the overlay inside": (_client_cpi_growth, {
        "mid": 3601.2849538954897, "growth": "Inputs!B7", "assumption-moved": "check"}),
    "a reference deleted under the value": (_ref_under_value, {"rebuild-error-low": "block", "rebuild-error-high": "block"}),
    "period dates with a time of day": (_timeline_times, SAME),
    "a decoy row reading the equity value's": (_decoy_reads_real, SAME),
    "the equity value's high end missing": (_no_high, {"equity": "block", "cells": False}),
    "the discountings as array SUMs": (lambda d: _summary_formulas(d, {"C4": "SUM(DCF!D11:W11*DCF!D12:W12)",
                                                                       "E4": "SUM(DCF!D18:W18*DCF!D19:W19)"}, True), SAME),
    "the discountings over a fixed OFFSET": (lambda d: _summary_formulas(d, {"C4": "SUM(OFFSET(DCF!D13,0,0,1,20))",
                                                                             "E4": "SUM(OFFSET(DCF!D20,0,0,1,20))"}), SAME),
    "net debt by a lookup on the inputs' labels": (lambda d: _summary_formulas(d, {
        a: '-_xlfn.XLOOKUP("Net debt at valuation date",Val_Inputs!$A$1:$A$10,Val_Inputs!$C$1:$C$10)' for a in ("C6", "E6")}),
        SAME),
    "this year's periods dated by their first day": (_first_days, SAME),
}
ORACLE = {"where": {"low": "Summary!C9", "high": "Summary!E9"}, "rate": {"low": ["Val_Inputs!C5"], "high": ["Val_Inputs!E5"]},
          "growth": {"low": "Val_Inputs!C6", "high": "Val_Inputs!C6"},
          "franking": {"low": "Val_Inputs!C9", "high": "Val_Inputs!C9"}}
BASE = {"held-0", "held-1", "review-0"}  # the clean pack's: the two inputs held at last year's, a review note


def run(files: list[Path], name: str, timeout: float = 600) -> dict:
    """The whole run on these files with the models stubbed (check_workbench's), this year's insurance confirmed
    ahead (the pack's one new term): its stages, needs, value and the cells it took."""
    import check_workbench as cw
    import library
    wb, orc = cw.wb, cw.orc
    eid = wb.create(name)["id"]
    cw.confirm_insurance(eid)
    for f in files:
        tmp = Path(tempfile.mkdtemp()) / f.name
        shutil.copy(f, tmp)
        wb.add_upload(eid, tmp, f.name, library.sha256_file(tmp))
    t0, quiet = time.time(), 0
    while time.time() - t0 < timeout and quiet < 3:
        v = orc.view(eid)
        st = {s["stage"]: s["status"] for s in v["stages"]}
        settled = st.get("review") in (*orc.SETTLED, "blocked", "failed") or any(x in ("blocked", "failed") for x in st.values())
        quiet = quiet + 1 if settled and not v["busy"] else 0
        time.sleep(1)
    v, e = orc.view(eid), wb.get(eid)
    res = e.get("result") or {}
    ins = res.get("inputs") or {}
    ends = lambda k: {x: ((ins.get(k) or {}).get("ends") or {}).get(x, {}).get("cells")
                      or ((ins.get(k) or {}).get("ends") or {}).get(x, {}).get("cell") for x in ("low", "high")}
    return {"stages": {s["stage"]: s["status"] for s in v["stages"]}, "needs": {n["id"]: n["severity"] for n in v["needs"]},
            "mid": ((res.get("values") or {}).get("this_year") or {}).get("mid"),
            "where": {k: (res.get("where") or {}).get(k) for k in ("low", "high")},
            "rate": ends("rate"), "growth": ends("growth"), "franking": ends("franking"),
            "errors": {w["filename"]: w.get("error") for w in e.get("workbooks") or [] if w["status"] == "error"}}


def pack_check(only: list[str] | None = None) -> None:
    import check_workbench as cw
    import library
    if not PACK.exists():
        sys.exit("run tests/make_pack.py first")
    cw.sandbox()
    cw.stub_models()
    import llm
    stubbed = llm.create

    def create(client, model, input, text=None, **kw):  # the equity cells in doubt: gpt-sol can't tell either
        if text and text["format"]["name"] == "equity_cells":
            return cw.Reply({"choice": "escalate", "reason": "two rows hold the report's figures", "question":
                             "Which row is last year's equity value?", "low_cell": "", "high_cell": ""})
        return stubbed(client, model, input, text=text, **kw)
    llm.create = create
    library.start()
    cw.orc.start()
    clean = run([PACK / f for f in (REPORT, PRIOR, OVERLAY, CURRENT)], "clean")
    assert clean["mid"] and {k: clean[k] for k in ORACLE} == ORACLE and set(clean["needs"]) == BASE, clean
    for name, (change, expect) in CASES.items():
        if only and name not in only:
            continue
        d = Path(tempfile.mkdtemp(prefix="pack_"))
        for f in (REPORT, PRIOR, OVERLAY, CURRENT, SLIDES, WITH_OVERLAY):
            shutil.copy(PACK / f, d / f)
        own = change(d)  # a list: the files this case uploads (the overlay inside a copy); else the pack A four
        files = own if isinstance(own, list) else [d / REPORT, d / PRIOR, d / OVERLAY if (d / OVERLAY).exists()
                                                   else d / OVERLAY.replace(".xlsx", ".xlsm"), d / CURRENT]
        t0 = time.time()
        got = run(files, name)
        print(f"  {name}: {time.time() - t0:.0f}s", flush=True)
        new = {k: s for k, s in got["needs"].items() if k not in BASE}
        if "mid" in expect:  # another value by design: that value, the cell named, and the needs expected
            assert got["mid"] is not None and abs(got["mid"] - expect["mid"]) < 1e-6, (name, got["mid"], new)
            assert got["growth"]["low"] == expect["growth"], (name, got["growth"])
            assert all(new.get(k) == s for k, s in expect.items() if k not in ("mid", "growth")), (name, expect, new)
        elif expect is SAME or expect.get("") == "same":  # the same value and cells; only the need expected, if any
            assert got["mid"] is not None and abs(got["mid"] - clean["mid"]) < 1e-6, (name, got["mid"], new)
            assert {k: got[k] for k in ORACLE} == ORACLE, (name, {k: got[k] for k in ORACLE})
            assert set(new) == {k for k in expect if k}, (name, new)
        elif "files" in expect:  # refused: the file's reason says what to do, nothing goes further
            assert got["stages"]["files"] == expect["files"] and got["mid"] is None, (name, got["stages"])
            assert any("password" in (x or "") for x in got["errors"].values()), (name, got["errors"])
        else:  # held, saying why: each expected need at its severity, the value not given
            assert got["mid"] is None, (name, got["mid"], new)
            assert all(new.get(k) == s for k, s in expect.items() if k != "cells"), (name, expect, new)
            if expect.get("cells", True):
                assert {k: got[k] for k in ("where", "rate", "growth", "franking")} == ORACLE, (name, got)
    print(f"pack: ok ({len(CASES)} ways a model arrives: {sum(1 for _n, (_c, x) in CASES.items() if x is SAME)} give the "
          f"same value and cells; the rest held or refused, each saying why)")


if __name__ == "__main__":
    if sys.argv[1:]:  # some of the pack's cases, by name
        pack_check(sys.argv[1:])
        sys.exit()
    references_check()
    choose_check()
    beyond_check()
    saved_state_check()
    containers_check()
    pack_check()
