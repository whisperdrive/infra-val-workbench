"""Reading the models: the shapes a workbook can come in, and what the app makes of each.

Every case keeps to one rule: the app takes the right cells and value, or it says what's wrong (a need, the value held,
the file refused with a reason); never a different figure in silence, never a crash.

  references   a number in scientific notation (1E9, 2.5E-3) isn't a cell, nor a function's name (LOG10() however
               like a cell it looks; a figure typed into a formula is read with its exponent
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


def pack_check() -> None:
    import check_workbench as cw
    import library
    if not PACK.exists():
        sys.exit("run tests/make_pack.py first")
    cw.sandbox()
    cw.stub_models()
    library.start()
    cw.orc.start()
    clean = run([PACK / f for f in (REPORT, PRIOR, OVERLAY, CURRENT)], "clean")
    assert clean["mid"] and {k: clean[k] for k in ORACLE} == ORACLE and set(clean["needs"]) == BASE, clean
    for name, (change, expect) in CASES.items():
        d = Path(tempfile.mkdtemp(prefix="pack_"))
        for f in (REPORT, PRIOR, OVERLAY, CURRENT):
            shutil.copy(PACK / f, d / f)
        change(d)
        files = sorted((p for p in d.iterdir() if p.suffix in (".pdf", ".xlsx", ".xlsm")), key=lambda p: p.name)
        got = run(files, name)
        new = {k: s for k, s in got["needs"].items() if k not in BASE}
        if expect is SAME or expect.get("") == "same":  # the same value and cells; only the need expected, if any
            assert got["mid"] is not None and abs(got["mid"] - clean["mid"]) < 1e-6, (name, got["mid"], new)
            assert {k: got[k] for k in ORACLE} == ORACLE, (name, {k: got[k] for k in ORACLE})
            assert set(new) == {k for k in expect if k}, (name, new)
        elif "files" in expect:  # refused: the file's reason says what to do, nothing goes further
            assert got["stages"]["files"] == expect["files"] and got["mid"] is None, (name, got["stages"])
            assert any("password" in (x or "") for x in got["errors"].values()), (name, got["errors"])
        else:  # held, saying why: each expected need at its severity, the value not given
            assert got["mid"] is None, (name, got["mid"], new)
            assert all(new.get(k) == s for k, s in expect.items()), (name, expect, new)
            assert {k: got[k] for k in ("where", "rate", "growth", "franking")} == ORACLE, (name, got)
    print(f"pack: ok ({len(CASES)} ways a model arrives: {sum(1 for _n, (_c, x) in CASES.items() if x is SAME)} give the "
          f"same value and cells; the rest held or refused, each saying why)")


if __name__ == "__main__":
    references_check()
    saved_state_check()
    containers_check()
    pack_check()
