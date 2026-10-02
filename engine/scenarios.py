"""The client models' settings: the scenario or case each was saved on, and when it was saved.

The overlay reads a client model's saved values, so the scenario the model was saved on is the one the value takes:
saved on another case, this year's model gives another value, and nothing in the roll shows it. A selector is a
typed cell on a row labelled as one (a scenario, a case, a sensitivity, a switch, the active one), holding a small
whole number or a word or two (an option), that the model's formulas read. This year's are shown next to last year's
(the prior model's, and the overlay's own copy where it has one), paired by rowfind, so a person confirms the scenario
the valuation should use; one that differs, or can't be paired, is a point to check. Each model's save time is its
own (a cell labelled as when it was saved), else the file's (its properties).
    find(wb, sheets) -> [{"sheet", "row", "col", "cell", "label", "value"}]
    saved(wb, path) -> {"date", "from" ("cell" or "file"), "cell", "label"} | None
    settings(sess, summary) -> {"selectors": [...], "saved": {"this_year", "last_year"}} | None
"""
import json
import re
import zipfile
from xml.etree import ElementTree as ET

import overlay as ov

SELECTOR = re.compile(r"scenario|sensitivit|\bswitch|\btoggle|\bselect|\b(active|current|chosen|live|run)\s+case\b|"
                      r"^\s*(case|active)\b", re.I)
AMOUNT = re.compile(r"[$€£%]|\bm\b|\bbn\b|'000|\bk\b", re.I)  # the units of an amount, not of an option
SAVED = re.compile(r"\bsaved?\b|time ?stamp|last (updated|modified)|date modified", re.I)
OPTIONS = 20  # a selector's number is a whole number from 0 to this
WORDS = 4     # ... or a word or two (an option), at most this many
SAVED_SPAN = (36526, 73051)  # a save time is a date from 2000 to 2100 (serials)


def _number(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def find(wb, sheets=None) -> list[dict]:
    """The selectors in a workbook (on these sheets; all of them by default), in the workbook's order. A row with
    more than two numbers is a series or a table, not a selector; a row the formulas don't read is a note (where the
    model has its formulas' edges at all)."""
    db = wb.db
    lay = {s: json.loads(x or "{}") for s, x in db.execute("SELECT sheet, layout FROM sheets")}
    try:
        read = {(s, r) for s, r in db.execute("SELECT DISTINCT dst_sheet, dst_row FROM edges "
                                               "WHERE src_sheet <> dst_sheet OR src_row <> dst_row")}
    except Exception:
        read = set()
    out = []
    for s, r, label, units in db.execute("SELECT sheet, row, label, units FROM rows ORDER BY rowid").fetchall():
        if (sheets is not None and s not in sheets) or not SELECTOR.search(label or "") or AMOUNT.search(units or "") \
                or (read and (s, r) not in read):
            continue
        label_col, units_col = lay.get(s, {}).get("label_col") or 0, lay.get(s, {}).get("units_col")
        typed = [(c, ov.from_db(v)) for c, v in db.execute("SELECT col, value FROM cells WHERE sheet=? AND row=? AND "
                                                           "formula IS NULL AND value IS NOT NULL ORDER BY col", (s, r))]
        nums = [(c, v) for c, v in typed if _number(v)]
        whole = [(c, v) for c, v in nums if float(v).is_integer() and 0 <= v <= OPTIONS]
        texts = [(c, v.strip()) for c, v in typed if isinstance(v, str) and c > label_col and c != units_col
                 and re.search(r"[A-Za-z]", v) and len(v.split()) <= WORDS and len(v.strip()) <= 30]
        if len(nums) > 2:
            continue
        got = whole[0] if whole else texts[0] if len(texts) == 1 and not nums else None
        if got:
            out.append({"sheet": s, "row": r, "col": got[0], "cell": ov._a1(s, r, got[0]), "label": label,
                        "value": got[1]})
    return out


def file_saved(path: str | None) -> str | None:
    """When the file was last saved, by its properties (docProps/core.xml): YYYY-MM-DD, or None."""
    if not path:
        return None
    try:
        with zipfile.ZipFile(path) as z:
            root = ET.fromstring(z.read("docProps/core.xml"))
    except (OSError, KeyError, zipfile.BadZipFile, ET.ParseError):
        return None
    got = next((el.text.strip() for el in root.iter() if el.tag.endswith("}modified") and (el.text or "").strip()), "")
    return got[:10] if re.match(r"^\d{4}-\d{2}-\d{2}", got) else None


def saved(wb, path: str | None = None) -> dict | None:
    """A model's save time: a cell labelled as when it was saved, holding a date (the model's own), else the file's
    properties. -> {"date", "from", "cell", "label"} or None."""
    for s, r, label in wb.db.execute("SELECT sheet, row, label FROM rows ORDER BY rowid").fetchall():
        if not SAVED.search(label or ""):
            continue
        for c, v in wb.db.execute("SELECT col, value FROM cells WHERE sheet=? AND row=? ORDER BY col", (s, r)):
            v = ov.from_db(v)
            if _number(v) and SAVED_SPAN[0] <= v < SAVED_SPAN[1]:
                return {"date": ov.to_date(v).isoformat(), "from": "cell", "cell": ov._a1(s, r, c), "label": label}
    d = file_saved(path)
    return {"date": d, "from": "file", "cell": None, "label": None} if d else None


def _same(a, b) -> bool:
    if _number(a) and _number(b):
        return abs(a - b) < 1e-9
    return isinstance(a, str) and isinstance(b, str) and a.strip().lower() == b.strip().lower()


def settings(sess, summary: dict) -> dict | None:
    """This year's selectors next to last year's, and when each model was saved. Last year's are on last year's
    client model (where the overlay is inside it, on the client's sheets and on any sheet this year's model has too:
    the roles can count a client's inputs sheet as the overlay's); the overlay's copy is the same cell on its own copy
    of the client's sheets, or what its link to the client model saved. Each selector:
    {"cell", "label", "this_year", "last_cell", "last_year", "overlay", "status"}, the status one of "same",
    "differs", "unmatched" (none like it last year), "gone" (last year's, not found this year). None without this
    year's model."""
    if not sess.current or not sess.rowmap:
        return None
    w = summary.get("wiring") or {}
    last = sess.prior or sess.ov
    mine = find(sess.current)
    theirs = find(last, None if sess.prior is not None else
                  set(sess.client_sheets) | {s for s in sess.sheets if sess.rowmap.sheet_for(s)})

    def copy(x):
        if sess.prior is not None and x["sheet"] in sess.client_sheets:
            return sess.ov.value(x["sheet"], x["row"], x["col"])
        if sess.client_link is not None:
            return sess.ext_cached.get((sess.client_link, x["sheet"], x["row"], x["col"]))
        return None

    by_row = {(y["sheet"], y["row"]): y for y in mine}
    paired, out = {}, []
    for x in theirs:
        k = sess.rowmap.locate(x["sheet"], x["row"])
        if k in by_row and k not in paired:
            paired[k] = x
    for k, y in by_row.items():
        x = paired.get(k)
        c = copy(x) if x else None
        status = "unmatched" if not x else "same" if _same(y["value"], x["value"]) and (c is None or _same(y["value"], c)) \
            else "differs"
        out.append({"cell": y["cell"], "label": y["label"], "this_year": y["value"], "last_cell": x and x["cell"],
                    "last_year": x and x["value"], "overlay": c, "status": status})
    for x in theirs:
        if x not in paired.values():
            out.append({"cell": None, "label": x["label"], "this_year": None, "last_cell": x["cell"],
                        "last_year": x["value"], "overlay": copy(x), "status": "gone"})
    last_path = ((w.get("prior") or {}) if sess.prior is not None else (w.get("overlay") or {})).get("source_path")
    return {"selectors": out, "saved": {"this_year": saved(sess.current, (w.get("current") or {}).get("source_path")),
                                        "last_year": saved(last, last_path)}}
