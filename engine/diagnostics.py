"""An engagement's run described for diagnosis without a word of the client's: counts, yes / no, ratios, dates and
the app's own words (stage names, statuses, kinds). Made to be exported from a machine with real files and pasted
into a session that can't see them. Every string in it is checked against the app's own vocabulary and dates before
it leaves; anything else is redacted and counted.
    export(eid) -> dict

What it says, per engagement:
  models      for each client model and the overlay: sheets, line items, formula cells, and the timelines by
              frequency (monthly / quarterly / semi-annual / annual: how many sheets, how many periods, first and last
              period): the model's shape (a finite-life quarterly model, an annual one with half-yearly distributions)
  alike       how alike last year's and this year's client models are: the share of line-item labels and of sheet
              names they have in common (a new version shares most; a rebuilt model few)
  profile     the financial-year end and the horizon (fixed or rolling)
  run         each stage's status, and the needs by kind and severity
  facts       how many, by status; failing checks by kind; image checks by status; which keys were found
  result      located, tied, held; the ratio of this year's value to last year's; the bridge's steps; the terminal
              value's method; each input checked or not; the discountings traced; the roll; the reliability gate's
              counts and ratios; the rows to find, each described by its shape (has a formula, adds up, how many rows it
              reads, how many overlay cells read it, a discounted cash flow or not, how many candidates)
"""
import json
import re
import statistics
import subprocess
import time
from pathlib import Path

import workbench as wb

DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
VERSION = re.compile(r"^[0-9a-f]{7,12}$")
REDACTED = "<redacted>"


def _vocabulary() -> set[str]:
    """Every word the export may carry: the app's own (stages, statuses, kinds, keys), never the files'."""
    import keyfacts
    import orchestrator
    words = set(orchestrator.STAGES) | set(orchestrator.LABEL) | {k for _p, k, _w in orchestrator.KINDS} | set(keyfacts.KNOWN)
    words |= {k for k, _l, _r in keyfacts.TV_KINDS} | {"unknown"}
    import methods
    words |= set(methods.LABEL)
    words |= {"waiting", "queued", "running", "done", "attention", "blocked", "failed", "error", "pending", "approved",
              "rejected", "agreed", "escalated", "withdrawn", "open", "ready", "processing", "unread", "verified",
              "flagged", "resolved", "edited", "figure", "confirmed", "confirmed by the reviewer", "corrected", "disputed",
              "native", "you", "orchestrator", "agents", "agent", "code", "block", "check", "info", "note",
              "monthly", "quarterly", "semi-annual", "annual", "irregular", "none", "fixed", "rolling",
              "dates", "date_check", "client_dates", "timelines", "assumed", "this_year_date",
              "quote", "table", "range", "ex", "cum", "low", "mid", "high", "rate", "growth", "multiple", "franking",
              "checked", "figure only", "not found", "no row this year", "no date this year", "nil",
              "report", "rounding", "prior_feed", "rebuilt", "roll", "time", "cash", "forecast", "held", "this_year",
              "SUMPRODUCT", "NPV", "XNPV", "SUM", "other", "dcf_missing", "blank_rows", "weak_rows", "timing_open",
              "prior_report", "prior_model", "prior_overlay", "current_model", "document", "workbook", "pdf", "pptx",
              "prior_model_vs_current_model", "overlay_vs_prior_model", "found", "picked", "stand_in", "m", "k", "bn",
              "end", "start", "rebuilt_rows", "label", "history", "words", "neighbours", "banner", "agents", "your"}
    return words


def _freq(serials: list[float]) -> str:
    """A timeline's frequency, from the median gap between its periods."""
    if len(serials) < 2:
        return "irregular"
    gap = statistics.median(b - a for a, b in zip(serials, serials[1:]))
    for name, lo, hi in (("monthly", 26, 35), ("quarterly", 85, 95), ("semi-annual", 175, 190), ("annual", 355, 375)):
        if lo <= gap <= hi:
            return name
    return "irregular"


def _iso(serial: float) -> str:
    from xlruntime import to_date
    return to_date(serial).isoformat()


def model_shape(db_path: str) -> dict:
    """A workbook's shape: its sheets, line items, formula cells, and its timelines by frequency."""
    t = wb._Timelines(db_path)
    try:
        db = t.db
        sheets = [s for (s,) in db.execute("SELECT sheet FROM sheets")]
        items = db.execute("SELECT COUNT(*) FROM rows WHERE label IS NOT NULL AND label <> ''").fetchone()[0]
        formulas = db.execute("SELECT COUNT(*) FROM cells WHERE formula IS NOT NULL").fetchone()[0]
        by = {}
        for s in sheets:
            tl = sorted(v for v in t.timeline(s).values())
            if len(tl) < 2:
                continue
            f = _freq(tl)
            x = by.setdefault(f, {"sheets": 0, "periods_max": 0, "first": None, "last": None})
            x["sheets"] += 1
            x["periods_max"] = max(x["periods_max"], len(tl))
            x["first"] = min(filter(None, (x["first"], _iso(tl[0]))))
            x["last"] = max(filter(None, (x["last"], _iso(tl[-1]))))
        return {"sheets": len(sheets), "line_items": items, "formula_cells": formulas, "timelines": by,
                "sheets_without_timeline": len(sheets) - sum(x["sheets"] for x in by.values())}
    finally:
        t.close()


def _labels(db_path: str) -> tuple[set, set]:
    import rodb
    with rodb.connect(db_path) as db:
        labels = {re.sub(r"\W+", " ", (l or "").lower()).strip() for (l,) in db.execute("SELECT label FROM rows")}
        names = {re.sub(r"\W+", " ", s.lower()).strip() for (s,) in db.execute("SELECT sheet FROM sheets")}
    return labels - {""}, names


def alike(a_path: str, b_path: str) -> dict:
    """How alike two models are: the share of line-item labels, and of sheet names, they have in common."""
    (la, sa), (lb, sb) = _labels(a_path), _labels(b_path)
    j = lambda x, y: round(len(x & y) / len(x | y), 3) if x | y else None
    return {"labels": j(la, lb), "sheet_names": j(sa, sb)}


def _call_kind(call: str | None) -> str:
    m = re.match(r"\s*=?\s*([A-Z]+)\s*\(", call or "")
    return m[1] if m and m[1] in ("SUMPRODUCT", "NPV", "XNPV", "SUM") else "other"


def _roll_basis(text: str | None) -> str | None:
    t = (text or "").lower()
    return None if not t else "date_check" if t.startswith("check:") else "assumed" if "assumed" in t else \
        "this_year_date" if "set for the engagement" in t else "client_dates" if "two client models" in t else \
        "timelines" if "timelines" in t else "dates"


def _count(xs) -> dict:
    out = {}
    for x in xs:
        out[x] = out.get(x, 0) + 1
    return out


def _version() -> str | None:
    try:
        return subprocess.run(["git", "rev-parse", "--short", "HEAD"], cwd=Path(__file__).resolve().parent.parent,
                              capture_output=True, text=True, timeout=5).stdout.strip() or None
    except (OSError, subprocess.SubprocessError):
        return None


def export(eid: int) -> dict | None:
    g = wb.get(eid)
    if not g:
        return None
    res, rl = g.get("result") or {}, g.get("roles") or {}
    wbs = {w["id"]: w for w in wb.workbooks(eid)}
    path = lambda role: (wbs.get((rl.get(role) or {}).get("id")) or {}).get("db_path") \
        if (rl.get(role) or {}).get("kind") == "workbook" else None
    out = {"app": _version(), "generated": time.strftime("%Y-%m-%d"),
           "files": {"reports": len(g.get("documents") or []), "workbooks": len(wbs)},
           "roles": {r: {"placed": bool(rl.get(r)), "by": (rl.get(r) or {}).get("by"),
                         "confirmed": bool((rl.get(r) or {}).get("confirmed"))}
                     for r in ("prior_report", "prior_overlay", "prior_model", "current_model")}}
    out["roles"]["prior_model"]["same_file_as_overlay"] = bool(path("prior_model")) and path("prior_model") == path("prior_overlay")
    models = {}
    for role in ("prior_overlay", "prior_model", "current_model"):
        p = path(role)
        if p and not (role == "prior_model" and out["roles"]["prior_model"]["same_file_as_overlay"]):
            try:
                models[role] = model_shape(p)
            except Exception as ex:  # a shape it can't read says so, by its kind only
                models[role] = {"error": type(ex).__name__}
    out["models"] = models
    out["alike"] = {}
    if path("prior_model") and path("current_model"):
        out["alike"]["prior_model_vs_current_model"] = alike(path("prior_model"), path("current_model"))
    if path("prior_overlay") and path("prior_model") and path("prior_overlay") != path("prior_model"):
        out["alike"]["overlay_vs_prior_model"] = alike(path("prior_overlay"), path("prior_model"))
    try:
        prof = wb.profile_view(eid).get("fields") or {}
        units = (prof.get("units") or {}).get("value") or ""
        conv = str((prof.get("discounting") or {}).get("value") or "").lower()
        out["profile"] = {"units_from_facts": bool(units), "fy_end_month": (prof.get("fy_end_month") or {}).get("value"),
                          "horizon": (prof.get("horizon") or {}).get("value"),
                          "frequency": (prof.get("frequency") or {}).get("value"),
                          "units_scale": "bn" if re.search(r"bn|billion", units, re.I) else "m" if re.search(r"m\b|million", units, re.I)
                          else "k" if re.search(r"k\b|'000|thousand", units, re.I) else None,
                          "timing": "mid" if "mid" in conv else "end" if "end" in conv else "start" if "start" in conv else None}
    except Exception as ex:
        out["profile"] = {"error": type(ex).__name__}
    out["dates"] = [{"date": d.get("date"), "ok": d.get("ok")} for d in g.get("dates") or []]
    run = g.get("run") or {}
    out["run"] = {"stages": {s["stage"]: s["status"] for s in run.get("stages") or []},
                  "needs": _count(f"{n.get('kind')}:{n.get('severity')}" for n in run.get("needs") or [])}
    facts = g.get("facts") or []
    out["facts"] = {"n": len(facts), "by_status": _count(f.get("status") for f in facts),
                    "failing_checks": _count(i.get("kind") or "figure" for f in facts
                                             for i in (f.get("check") or {}).get("items") or [] if not i.get("ok")),
                    "image": _count((f.get("visual") or {}).get("status") for f in facts if f.get("visual")),
                    "keys": sorted({f.get("key") for f in facts if f.get("key")})}
    out["result"] = _result(eid, res)
    return redact(out)


def _result(eid: int, res: dict) -> dict:
    if not res:
        return {"worked_out": False}
    if res.get("stop"):
        return {"worked_out": False, "stopped": True}
    v, tie = res.get("values") or {}, res.get("tie") or {}
    r = {"worked_out": True, "basis": (res.get("head") or {}).get("basis"),
         "located": bool(res.get("where")), "tied": {e: (tie.get(e) or {}).get("ok") for e in ("low", "high")},
         "rebuilt_tied": {e: (tie.get(e) or {}).get("rebuilt_ok") for e in ("low", "high")},
         "held": bool((res.get("bridges") or {}).get("held")),
         "this_to_last": round(v["this_year"]["mid"] / v["report"]["mid"], 4)
         if (v.get("this_year") or {}).get("mid") and (v.get("report") or {}).get("mid") else None,
         "bridge_steps": [s["key"] for s in ((res.get("bridges") or {}).get("mid") or {}).get("steps") or []],
         "bridge_notes": len((res.get("bridges") or {}).get("notes") or []),
         "terminal": (res.get("terminal") or {}).get("kind"),
         "inputs": {k: {"ok": x.get("ok"), "not_applicable": bool(x.get("na")),
                        "sourced": {e: (x.get("ends") or {}).get(e, {}).get("sourced") for e in ("low", "high")}}
                    for k, x in (res.get("inputs") or {}).items()},
         "held_inputs": {"n": len(res.get("held") or []), "still_held": sum(1 for h in res.get("held") or [] if h["held"]),
                         "suggestions": _count((h.get("suggestion") or {}).get("status") for h in res.get("held") or [])},
         "rate_this_year": {"set": bool(((res.get("inputs") or {}).get("rate") or {}).get("this_year")),
                            "applied": bool((((res.get("inputs") or {}).get("rate") or {}).get("this_year") or {}).get("applied"))},
         "cut_off": (res.get("figures") or {}).get("cut_off"),
         "methods": _methods(res.get("methods") or {}),
         "chart": {"years": len((res.get("chart") or {}).get("years") or []), "found": bool((res.get("chart") or {}).get("years"))}}
    # the discountings traced under each end: how many, and how many can be read here
    try:
        import dcftrace
        import rodb
        sess, summary = wb.overlay_session(eid)
        with rodb.connect(summary["wiring"]["overlay"]["db_path"]) as db:
            disc = {}
            for e in ("low", "high"):
                cell = (res.get("where") or {}).get(e)
                try:
                    every = dcftrace.cores(dcftrace.trace(db, cell)) if cell else []
                except ValueError:
                    every = []
                disc[e] = {"found": len(every), "readable": sum(1 for c in every if c.get("inputs")),
                           "calls": _count(_call_kind(c.get("call")) for c in every)}
        r["discountings"] = disc
        roll = summary.get("roll") or {}
        r["roll"] = {"months": roll.get("months"), "basis": _roll_basis(roll.get("months_basis")),
                     "date_check": bool(roll.get("date_check")), "fixed_horizon": bool(roll.get("fixed_horizon")),
                     "date_cells_moved": len(roll.get("valuation_date_cells") or []),
                     "date_cells_read": len(roll.get("valuation_date_reads") or [])}
    except Exception as ex:
        sess, r["discountings"] = None, {"error": type(ex).__name__}
    g = (res.get("figures") or {}).get("gaps") or {}
    if g:
        r["gate"] = {"reliable": g.get("reliable"), "no_reads": g.get("no_reads"), "reads": g.get("reads"),
                     "unmatched": g.get("values"), "found_share": g.get("found_share"), "dcf_rows": g.get("dcf_rows"),
                     **{k: len(g.get(k) or []) for k in ("dcf_missing", "blank_rows", "weak_rows", "timing_open")},
                     "zero_roll": {c: x["zero_roll"].get("ratio") for c, x in (g.get("by_cell") or {}).items()},
                     "date_cells_off": len((g.get("date_cells") or {}).get("off") or []),
                     "family": g.get("family"), "rebuilt": g.get("rebuilt"), "rebuilt_rows": len(g.get("rebuilt_rows") or []),
                     "date_by_label": bool((g.get("date_cells") or {}).get("by_label"))}
        # the cells' names aren't the client's words, but keep them out anyway: the ends, in order
        r["gate"]["zero_roll"] = list(r["gate"]["zero_roll"].values())
        tc = (res.get("figures") or {}).get("time") or {}
        if tc:
            r["gate"]["time"] = {"measured": tc.get("measured"), "ok": tc.get("ok"), "hold": tc.get("hold"),
                                 "off_bp": [round(1e4 * x["off"]) for x in tc.get("discountings") or []],
                                 "flows_matched": [x["periods"] for x in tc.get("discountings") or []]}
        if sess is not None and sess.rowmap:
            origins = set(g.get("dcf_origins") or [])
            rows = [(k, x) for k in ("dcf_missing", "blank_rows", "weak_rows", "timing_open", "rebuilt_rows")
                    for x in g.get(k) or []][:15]
            r["family"] = sess.rowmap.family()
            shapes = []
            for kind, x in rows:
                try:
                    import overlay as ovmod
                    info = ovmod.deep(wb.row_found, sess, *wb._row_ref(x["row"]), origins)
                except Exception as ex:
                    shapes.append({"kind": kind, "error": type(ex).__name__})
                    continue
                c, f = info.get("context") or {}, (info.get("context") or {}).get("formula") or {}
                how = str(info.get("how") or "").lower()
                shapes.append({"kind": kind, "found": bool(info.get("found")), "confidence": info.get("confidence"),
                               "how": next((w for w in ("label", "history", "words", "neighbours", "banner", "agents", "your")
                                            if w in how), "none" if not how else "other"),
                               "copies_passed_over": len(info.get("copies") or []),
                               "has_formula": bool(f), "adds_up": bool(f.get("adds_up")), "reads_rows": f.get("n", 0),
                               "read_by_cells": sum(x["cells"] for x in c.get("read_by") or []),
                               "feeds_dcf": bool(c.get("feeds_dcf")), "values": len(c.get("values") or []),
                               "candidates": len(info.get("candidates") or []),
                               "candidates_with_figures": sum(1 for k in info.get("candidates") or []
                                                              if any(v is not None for v in k.get("figures") or []))})
            r["rows"] = shapes
    return r


def _methods(inv: dict) -> dict:
    """The methods inventory as ratios: each method's mid against the default's, whether it was worked out, the
    preferred one, whether the recompute ties, how many forecast flags of the overlay's own were found."""
    base = next((m for m in inv.get("methods") or [] if m["key"] == inv.get("default")), {}).get("mid")
    return {"preferred": inv.get("preferred"), "ties": inv.get("ties"), "flags": len(inv.get("flags") or []),
            "error": bool(inv.get("error")),
            "each": [{"key": m["key"], "ok": m["ok"], "vs_default": round(m["mid"] / base - 1, 5)
                      if m.get("mid") is not None and base else None} for m in inv.get("methods") or []]}


def redact(obj, words: set | None = None, n: list | None = None):
    """The export with every string not the app's own word, a date or the app's version replaced (keys too)."""
    words = words if words is not None else _vocabulary()
    top = n is None
    n = [0] if n is None else n

    def ok(s: str) -> bool:
        return s in words or bool(DATE.match(s)) or bool(VERSION.match(s)) or bool(re.fullmatch(r"[a-z_-]+:(block|check|info)", s)) \
            and s.split(":")[0] in words

    def walk(x):
        if isinstance(x, dict):
            return {(k if ok(k) or k in _KEYS else _red(n)): walk(v) for k, v in x.items()}
        if isinstance(x, list):
            return [walk(v) for v in x]
        if isinstance(x, str):
            return x if ok(x) else _red(n)
        return x

    out = walk(obj)
    if top:
        out["redacted"] = n[0]
    return out


def _red(n: list) -> str:
    n[0] += 1
    return REDACTED


# the export's own field names: never the client's
_KEYS = {"app", "generated", "files", "reports", "workbooks", "roles", "placed", "by", "confirmed", "same_file_as_overlay",
         "models", "sheets", "line_items", "formula_cells", "timelines", "periods_max", "first", "last",
         "sheets_without_timeline", "error", "alike", "labels", "sheet_names", "profile", "fy_end_month", "horizon", "dates",
         "date", "ok", "run", "stages", "needs", "facts", "n", "by_status", "failing_checks", "image", "keys", "result",
         "worked_out", "stopped", "basis", "located", "tied", "rebuilt_tied", "held", "this_to_last", "bridge_steps",
         "bridge_notes", "terminal", "inputs", "not_applicable", "sourced", "held_inputs", "still_held", "suggestions",
         "chart", "years", "found", "discountings", "readable", "calls", "roll", "months", "date_check", "fixed_horizon",
         "date_cells_moved", "date_cells_read", "gate", "reliable", "no_reads", "reads", "unmatched", "found_share",
         "dcf_rows", "dcf_missing", "blank_rows", "weak_rows", "timing_open", "zero_roll", "date_cells_off",
         "date_by_label", "family", "rows", "time", "measured", "hold", "off_bp", "flows_matched", "frequency", "units_scale", "timing", "units_from_facts", "kind", "confidence", "has_formula", "adds_up", "reads_rows", "read_by_cells",
         "feeds_dcf", "values", "candidates", "candidates_with_figures", "redacted", "rebuilt", "rebuilt_rows", "how",
         "copies_passed_over", "low", "high", "mid", "rate_this_year", "set", "applied", "cut_off",
         "methods", "preferred", "ties", "flags", "each", "key", "vs_default",
         "monthly", "quarterly", "semi-annual", "annual", "irregular"}


def as_text(eid: int) -> str:
    return json.dumps(export(eid), indent=1, default=str)
