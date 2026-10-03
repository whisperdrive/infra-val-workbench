"""What drives the value: every input under the equity value, traced from the value down to the figures typed in the
overlay and read from the client model, grouped by what each does, with how far a 1% move in each moves the value.

  walk     from the equity value's cells down through every formula the Python rebuild reads (cell by cell, on last
           year's feed), to the leaves: figures typed in the overlay, and figures read from the client model
  group    the leaves by what they do, from what the app has already sourced and checked: the discount rate (the
           cells the factors read), the terminal value (its growth rate or multiple, and the figure it grows from),
           franking (the utilisation), the valuation date and the periods, the inputs held at last year's (net debt,
           a distribution), the cash flows (the client rows the discountings' cash flows come from), and the rest
  rows     leaves collapsed to rows: a cash-flow row is one input, not one per period
  effect   each row (or single figure) moved 1% (all its periods together), the value worked out again: the move in
           the mid, in the report's units; at most NUDGES rows, the ones the groups name first
  client   for a client row, the client model's own typed inputs it's worked out from (its rows' lineage: volumes,
           tariffs, CPI), named, not moved (the client model's formulas aren't rebuilt here)
"""
from collections import deque

import overlay as ov

WALK = 60000   # cells visited at most on the way down from the value
NUDGES = 40    # rows moved to measure their effect, at most
NUDGE = 0.01   # each moved by this much, relative
GROUPS = ("Discount rate", "Terminal value", "Franking credits", "Dates and periods", "Held at last year's",
          "Cash flows (client model)", "Other overlay inputs")


def _num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def walk(sess, cells: list[tuple]) -> tuple[set, set]:
    """The leaves under these cells on the session's current feed: (overlay cells typed in: {(sheet, row, col)},
    client cells read: {(link or "", sheet, row, col)}). The cells are worked out first, so their reads are known."""
    B = sess.B
    sess.values(cells)
    seen, q = set(), deque(("", s, r, c) for s, r, c in cells)
    typed, client = set(), set()
    while q and len(seen) < WALK:
        src, s, r, c = q.popleft()
        if (src, s, r, c) in seen:
            continue
        seen.add((src, s, r, c))
        if src != "" or s in sess.client_sheets:
            client.add((src, s, r, c))
            continue
        if not B.is_formula(s, r, c):
            typed.add((s, r, c))
            continue
        reads, _ = B.reads(s, r, c)
        q.extend(x for x in reads if x not in seen)
    return typed, client


def _known(summary: dict, inputs: dict, held_inputs: list, figs: dict) -> dict:
    """{(sheet, row): group} for the overlay rows the app has already sourced: their group."""
    out = {}

    def put(cell, group):
        try:
            s, r, _c = ov.parse_a1(cell)
        except (TypeError, ValueError, AttributeError):
            return
        out.setdefault((s, r), group)
    for e in ("low", "high"):
        x = ((inputs.get("rate") or {}).get("ends") or {}).get(e) or {}
        for c in (x.get("cells") or []) + ([x["cell"]] if x.get("cell") else []):
            put(c, "Discount rate")
        for key in ("growth", "multiple"):
            y = ((inputs.get(key) or {}).get("ends") or {}).get(e) or {}
            if y.get("cell"):
                put(y["cell"], "Terminal value")
            if (y.get("base") or {}).get("cell"):
                put(y["base"]["cell"], "Terminal value")
        z = ((inputs.get("franking") or {}).get("ends") or {}).get(e) or {}
        if z.get("cell"):
            put(z["cell"], "Franking credits")
    for h in held_inputs or []:
        put(h.get("cell"), "Held at last year's")
    roll = summary.get("roll") or {}
    for c in (roll.get("valuation_date_cells") or []) + (roll.get("valuation_date_reads") or []):
        put(c, "Dates and periods")
    return out


def value_inputs(sess, summary: dict, where: dict, figs: dict, inputs: dict, held_inputs: list, unit) -> dict:
    """{"groups": [{"group", "inputs": [{"where" ("overlay" / "client"), "row", "label", "cells", "value", "effect",
    "made_from"}]}], "nudged", "capped", "leaves"}. On last year's feed (the rebuild the report ties to). Runs in
    overlay.deep."""
    import time
    import structure
    t0 = time.time()
    feed = figs["feeds"]["rebuilt"]
    cells = [ov.parse_a1(c) for c in dict.fromkeys(x for x in (where.get("low"), where.get("high")) if x)]
    try:
        sess.configure(feed)
        typed, client = walk(sess, cells)
        base = dict(zip(cells, sess.values(cells)))
    finally:
        sess.configure("workbook")
    known = _known(summary, inputs, held_inputs, figs)
    olab = sess.ov.labels()
    # the overlay's typed figures, by row (text and blank cells aren't inputs)
    rows = {}
    for s, r, c in typed:
        v = _num(sess.ov.value(s, r, c))
        if v is None:
            continue
        x = rows.setdefault(("", s, r), {"where": "overlay", "row": f"{s}!r{r}", "label": olab.get((s, r), ""),
                                         "cells": set(), "group": known.get((s, r), "Other overlay inputs")})
        x["cells"].add(c)
        x["value"] = v
    # the client model's figures, by row
    pri = sess.prior or sess.ov
    plab = pri.labels()
    for src, s, r, c in client:
        x = rows.setdefault((src, s, r), {"where": "client", "row": f"{s}!r{r}", "label": plab.get((s, r), ""),
                                          "cells": set(), "group": "Cash flows (client model)", "link": src})
        x["cells"].add(c)
    # what each client row is worked out from in the client model: its typed inputs, named
    try:
        cst = structure.load(pri.path)
        reads, _ = structure.edges(pri.path)
        for key, x in rows.items():
            if x["where"] != "client":
                continue
            up = structure.closure((key[1], key[2]), reads)
            ins = [k for k in up if k != (key[1], key[2]) and not reads.get(k)
                   and cst["info"].get(k, {}).get("kind") in ("inputs", "figure", "index", "share")]
            x["made_from"] = sorted({cst["info"][k]["label"] for k in ins if cst["info"][k]["label"]})[:8]
    except Exception:
        pass
    # each row moved 1%, all its periods together: the move in the mid, in the report's units
    order = sorted(rows.items(), key=lambda kv: (GROUPS.index(kv[1]["group"]) if kv[1]["group"] in GROUPS else 9,
                                                 -len(kv[1]["cells"])))
    nudged = 0
    mid0 = _mid(base, cells, unit)
    for key, x in order[:NUDGES]:
        moved = _nudged(sess, feed, cells, key, x)
        mid1 = _mid(moved, cells, unit) if moved else None
        x["effect"] = (mid1 - mid0) if mid1 is not None and mid0 is not None else None
        nudged += 1
    groups = []
    for g in GROUPS:
        xs = [x for x in rows.values() if x["group"] == g]
        if xs:
            xs.sort(key=lambda x: -abs(x.get("effect") or 0.0))
            groups.append({"group": g, "inputs": [{**{k: v for k, v in x.items() if k != "cells"}, "cells": len(x["cells"]),
                                                   "value": x.get("value") if len(x["cells"]) == 1 else None}
                                                  for x in xs]})
    return {"groups": groups, "nudged": nudged, "capped": len(rows) > NUDGES, "leaves": {"overlay": len(typed),
            "client": len(client)}, "rows": len(rows), "per": NUDGE, "secs": round(time.time() - t0, 2)}


def _mid(vals: dict, cells: list, unit) -> float | None:
    xs = [unit(_num(vals.get(c))) for c in cells if _num(vals.get(c)) is not None]
    return sum(xs) / len(xs) if xs else None


def _nudged(sess, feed: str, cells: list, key: tuple, x: dict) -> dict | None:
    """The value's cells with one input row moved NUDGE (relative), on last year's feed. The session is left on the
    workbook feed."""
    src, s, r = key
    cols = x["cells"]
    bump = lambda v: v * (1 + NUDGE) if _num(v) is not None else v
    try:
        sess.configure(feed)
        if x["where"] == "overlay":
            for c in cols:
                v = sess.ov.value(s, r, c)
                if _num(v) is not None:
                    sess.B.overrides[(s, r, c)] = bump(v)
            sess.B.reset()
        else:
            f0, e0 = sess.B.feed, sess.B.ext
            hit = lambda s_, r_, c_: (s_, r_) == (s, r) and c_ in cols
            sess.B.feed = lambda s_, r_, c_: bump(f0(s_, r_, c_)) if hit(s_, r_, c_) else f0(s_, r_, c_)
            sess.B.ext = lambda i, s_, r_, c_: bump(e0(i, s_, r_, c_)) if hit(s_, r_, c_) else e0(i, s_, r_, c_)
            sess.B.reset()
        return dict(zip(cells, sess.values(cells)))
    except Exception:
        return None
    finally:
        sess.configure("workbook")
