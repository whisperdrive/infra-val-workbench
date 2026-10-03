"""A workbook's structure, read from its shapes rather than its words: what each row is, how rows group into blocks,
which blocks are copies of each other, and where a row's figures come from and go. One of the tools the row finder
and the agents combine (labels, numbers, positions, lineage, structure): cheap ones first, the next when they
disagree or can't settle it.

  kinds    each line item's kind, from its values along the timeline and its formulas: dates, flags (0 / 1),
           discount factors (between 0 and 1, falling), a share (one figure between 0 and 1), an index (each period
           the one before times a near-constant 1 + x), a subtotal (adds up the rows above it), a ratio, a
           calculated series, typed inputs, a single figure
  blocks   consecutive line items under one heading, each with a signature: the sequence of its rows' kinds and
           formula shapes. Two blocks with the same signature are copies (a base and a downside case, a 100% and a
           share, nominal and real): a label found in one is found in the other too, so a label alone can't say
           which is meant
  lineage  a row's upstream (what it's worked out from, down to the inputs) and downstream (what reads it), across
           the overlay's link to the client model, and whether it reaches the equity value
  map      the structure as a few lines a person or a model can read: sheets, blocks and their rows, by kind

Reads a model.db (build_map): rows (section, label, formula patterns), cells, edges (row -> the rows it reads),
and an overlay's extrefs (its rows reading the client model's).
"""
import difflib
import json
import re
import sqlite3
import statistics
from collections import defaultdict, deque

MIN_SERIES = 3     # periods a row needs before its shape says what it is
COPY_SIMILAR = 0.85  # blocks this alike (their rows' kinds and shapes, in order) are copies of each other
COPY_ROWS = 3      # ... with at least this many rows
GAP = 1            # blank rows a block may have inside it
INDEX_STEP = (0.995, 1.2)  # an index's period-on-period ratio: about constant, within this


def _ro(path: str) -> sqlite3.Connection:
    return sqlite3.connect(f"file:{path}?mode=ro", uri=True)


def _num(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.replace(",", ""))
        except ValueError:
            return None
    return None


_ISO = re.compile(r"^\d{4}-\d{2}-\d{2}")
_ABS_ROW = re.compile(r"R(\d+)")
_SHEETREF = re.compile(r"'[^']+'!|[A-Za-z_][\w.]*!")


def shape(pattern: str | None) -> str:
    """A row's dominant formula, as a shape: its operators and functions, references as R[] (relative) or R#
    (absolute) without the distance (a block copied elsewhere reads the rows it shares from further away), sheet
    names left out."""
    first = _first(pattern)
    first = re.sub(r"\[-?\d+\]", "[]", _ABS_ROW.sub("R#", _SHEETREF.sub("!", first)))
    return re.sub(r"C\d+", "C#", first)


def _first(pattern: str | None) -> str:
    """The row's most common formula, in R1C1, as build_map wrote it."""
    if not pattern:
        return ""
    return re.sub(r"\s+x\d+.*$", "", pattern.split(";")[0]).strip()


def kind_of(vals: list, n_formula: int, n_const: int, pattern: str | None) -> str:
    """What a row is, from its values (in column order) and formulas."""
    texts = [v for v in vals if isinstance(v, str) and not _ISO.match(v)]
    dates = [v for v in vals if isinstance(v, str) and _ISO.match(v)]
    nums = [x for x in (_num(v) for v in vals if not (isinstance(v, str) and _ISO.match(v))) if x is not None]
    if len(dates) >= MIN_SERIES and len(dates) >= len(nums):
        return "dates"
    if not nums:
        return "text" if texts else "empty"
    if len(nums) == 1:
        return "figure"
    raw = _first(pattern)
    if re.search(r"SUM\(R\[-\d+\]C(?:\[\d+\])?:R\[-1\]C", raw) or re.search(r"=\s*R\[-\d+\]C(\s*[+-]\s*R\[-\d+\]C){2,}", raw):
        return "subtotal"
    if len(nums) >= MIN_SERIES:
        nz = [x for x in nums if x]
        if set(nums) <= {0.0, 1.0} and 1.0 in nums:
            return "flags"
        if len(set(round(x, 12) for x in nums)) == 1 and 0 < nums[0] < 1:
            return "share"
        if all(0 < x <= 1 for x in nz) and len(nz) >= MIN_SERIES and nz[0] > 0.5 and \
                all(b <= a + 1e-12 for a, b in zip(nz, nz[1:])) and nz[-1] < nz[0]:
            return "factors"  # starting near 1 and falling (a row of rates falling across the ends isn't)
        if len(nz) >= 4 and 0.5 <= nz[0] <= 2 and all(a > 0 for a in nz):
            steps = [b / a for a, b in zip(nz, nz[1:])]
            if all(INDEX_STEP[0] <= s <= INDEX_STEP[1] for s in steps) and statistics.pstdev(steps) < 0.01:
                return "index"
        if all(abs(x) < 1 for x in nz) and nz:
            return "ratio"
    if n_formula >= max(1, n_const):
        return "series"
    return "inputs"


SCAN = 40  # columns read to the right of the labels on a sheet with no timeline


def rows(db_path: str) -> dict:
    """{(sheet, row): {"label", "section", "kind", "shape", "n_formula", "n_const", "units"}} for every line item.
    Each sheet's cells are read along its timeline only (or the first SCAN columns right of its labels), not the
    whole workbook at once."""
    out = {}
    with _ro(db_path) as db:
        layouts = {s: json.loads(l or "{}") for s, l in db.execute("SELECT sheet, layout FROM sheets")}
        items = defaultdict(list)
        for s, r, section, label, units, nf, nc, pats in db.execute(
                "SELECT sheet, row, section, label, units, n_formula, n_const, patterns FROM rows ORDER BY sheet, row"):
            items[s].append((r, section, label, units, nf, nc, pats))
        for s, its in items.items():
            lay = layouts.get(s) or {}
            a, b = lay.get("tl_first"), lay.get("tl_last")
            if a is None or b is None:
                a = (lay.get("label_col") or 1) + 1
                b = a + SCAN
            skip = {lay.get("label_col"), lay.get("units_col")}
            vals = defaultdict(dict)
            for r, c, v in db.execute("SELECT row, col, value FROM cells WHERE sheet=? AND col BETWEEN ? AND ?", (s, a, b)):
                if c not in skip:
                    vals[r][c] = v
            for r, section, label, units, nf, nc, pats in its:
                row_vals = vals.get(r, {})
                ordered = [row_vals[c] for c in sorted(row_vals)]
                out[(s, r)] = {"label": label or "", "section": section or "", "units": units or "",
                               "n_formula": nf or 0, "n_const": nc or 0, "shape": shape(pats),
                               "kind": kind_of(ordered, nf or 0, nc or 0, pats)}
    return out


def blocks(info: dict) -> list[dict]:
    """Consecutive line items under one heading on one sheet: {"sheet", "heading", "first", "last", "rows", "signature"}.
    The signature: each row's kind and formula shape, in order."""
    out, cur = [], None
    for (s, r) in sorted(info):
        x = info[(s, r)]
        if x["kind"] in ("empty", "text"):
            continue
        if cur and cur["sheet"] == s and cur["heading"] == x["section"] and r - cur["last"] <= GAP + 1:
            cur["rows"].append(r)
            cur["last"] = r
            continue
        cur = {"sheet": s, "heading": x["section"], "first": r, "last": r, "rows": [r]}
        out.append(cur)
    for b in out:
        b["signature"] = [f"{info[(b['sheet'], r)]['kind']}|{info[(b['sheet'], r)]['shape']}" for r in b["rows"]]
        b["labels"] = [re.sub(r"[^a-z]+", " ", info[(b["sheet"], r)]["label"].lower()).strip() for r in b["rows"]]
    return out


def _cover(a: list, b: list) -> tuple[float, int]:
    """How much of the shorter sequence the other contains, in order: (share of the shorter, rows matched)."""
    m = sum(x.size for x in difflib.SequenceMatcher(None, a, b, autojunk=False).get_matching_blocks())
    return m / max(1, min(len(a), len(b))), m


def copies(bl: list[dict]) -> list[dict]:
    """Pairs of blocks where one holds a copy of the other: the same kinds of row in the same order, under the same
    labels or with the same formulas' shapes, over most of the shorter block (a block running on into other rows
    without a heading between still holds its copy). {"a", "b", "similar"}, each block as "Sheet!rA:rB"."""
    out = []
    big = [b for b in bl if len(b["rows"]) >= COPY_ROWS]
    for i, a in enumerate(big):
        for b in big[i + 1:]:
            by_label = _cover([f"{s.split('|')[0]}|{l}" for s, l in zip(a["signature"], a["labels"])],
                              [f"{s.split('|')[0]}|{l}" for s, l in zip(b["signature"], b["labels"])])
            by_shape = _cover(a["signature"], b["signature"])
            sim, n = max(by_label, by_shape)
            if sim >= COPY_SIMILAR and n >= COPY_ROWS:
                out.append({"a": f"{a['sheet']}!r{a['first']}:r{a['last']}", "b": f"{b['sheet']}!r{b['first']}:r{b['last']}",
                            "a_heading": a["heading"], "b_heading": b["heading"], "similar": round(sim, 3), "rows": n})
    return out


def block_of(bl: list[dict], sheet: str, row: int) -> dict | None:
    return next((b for b in bl if b["sheet"] == sheet and row in b["rows"]), None)


def _words(text: str) -> set:
    return {w for w in re.findall(r"[a-z0-9]+", (text or "").lower()) if len(w) > 2}


def which_copy(prior_bl: list[dict], cur_bl: list[dict], cur_copies: list[dict], row: tuple, cands: list[tuple]) -> dict:
    """Where last year's row has a label this year's model has in more than one copy of a block (a base and a
    downside case inserted above it, say), which copy is meant: the one whose heading is last year's block's, and
    the row at the same place in it. {"pick" (sheet, row) or None, "why", "ambiguous" (the copies)}."""
    mine = block_of(prior_bl, *row)
    blocks = [(k, block_of(cur_bl, *k)) for k in cands]
    blocks = [(k, b) for k, b in blocks if b]
    names = {f"{b['sheet']}!r{b['first']}:r{b['last']}" for _, b in blocks}
    copied = [c for c in cur_copies if c["a"] in names and c["b"] in names]
    if len(blocks) < 2 or not copied:
        return {"pick": None, "why": "no copies among the candidates", "ambiguous": []}
    want = _words(mine["heading"]) if mine else set()
    scored = sorted(((len(want & _words(b["heading"])) / max(1, len(want | _words(b["heading"]))), k, b) for k, b in blocks),
                    key=lambda x: -x[0])
    best = scored[0]
    if best[0] > 0 and (len(scored) == 1 or best[0] > scored[1][0]):
        return {"pick": best[1], "why": f"the copy headed '{best[2]['heading']}', as last year's block was "
                                        f"('{mine['heading'] if mine else ''}')", "ambiguous": []}
    return {"pick": None, "why": "copies of the block with headings that don't say which is last year's",
            "ambiguous": [f"{b['sheet']}!r{b['first']}:r{b['last']} '{b['heading']}'" for _, k, b in scored]}


VERSION = 1  # the structure's rules: tables written by older rules are worked out again (in memory)
_CACHE: dict = {}


def build(db_path: str) -> dict:
    """The structure worked out and written into the model.db (struct_rows, struct_blocks, struct_copies,
    struct_meta), at the end of the workbook's build, while nothing else reads it. -> counts."""
    info = rows(db_path)
    bl = blocks(info)
    cp = copies(bl)
    with sqlite3.connect(db_path) as db:
        db.executescript("""
            DROP TABLE IF EXISTS struct_rows; DROP TABLE IF EXISTS struct_blocks; DROP TABLE IF EXISTS struct_copies;
            DROP TABLE IF EXISTS struct_meta;
            CREATE TABLE struct_rows(sheet TEXT, row INT, kind TEXT, shape TEXT, block INT);
            CREATE TABLE struct_blocks(id INT, sheet TEXT, heading TEXT, first INT, last INT, rows TEXT, signature TEXT,
                                       labels TEXT);
            CREATE TABLE struct_copies(a TEXT, b TEXT, a_heading TEXT, b_heading TEXT, similar REAL, n INT);
            CREATE TABLE struct_meta(key TEXT, value TEXT);
        """)
        where = {(b["sheet"], r): i for i, b in enumerate(bl) for r in b["rows"]}
        db.executemany("INSERT INTO struct_rows VALUES (?,?,?,?,?)",
                       [(s, r, x["kind"], x["shape"], where.get((s, r))) for (s, r), x in info.items()])
        db.executemany("INSERT INTO struct_blocks VALUES (?,?,?,?,?,?,?,?)",
                       [(i, b["sheet"], b["heading"], b["first"], b["last"], json.dumps(b["rows"]),
                         json.dumps(b["signature"]), json.dumps(b["labels"])) for i, b in enumerate(bl)])
        db.executemany("INSERT INTO struct_copies VALUES (?,?,?,?,?,?)",
                       [(c["a"], c["b"], c["a_heading"], c["b_heading"], c["similar"], c["rows"]) for c in cp])
        db.execute("INSERT INTO struct_meta VALUES ('version', ?)", (str(VERSION),))
    return {"rows": len(info), "blocks": len(bl), "copies": len(cp)}


def load(db_path: str) -> dict:
    """{"info", "blocks", "copies"} for a model.db: from its tables where the build wrote them by these rules, else
    worked out in memory (an older database isn't written to while others read it). Cached by path and time."""
    import os
    key = (db_path, os.path.getmtime(db_path))
    if key in _CACHE:
        return _CACHE[key]
    got = None
    try:
        with _ro(db_path) as db:
            have = {r[0] for r in db.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            if "struct_meta" in have and (db.execute("SELECT value FROM struct_meta WHERE key='version'").fetchone()
                                          or [None])[0] == str(VERSION):
                labels = {(s, r): (sec or "", lab or "", u or "", nf or 0, nc or 0) for s, r, sec, lab, u, nf, nc in
                          db.execute("SELECT sheet, row, section, label, units, n_formula, n_const FROM rows")}
                info = {}
                for s, r, kind, sh, _b in db.execute("SELECT sheet, row, kind, shape, block FROM struct_rows"):
                    sec, lab, u, nf, nc = labels.get((s, r), ("", "", "", 0, 0))
                    info[(s, r)] = {"label": lab, "section": sec, "units": u, "n_formula": nf, "n_const": nc,
                                    "shape": sh, "kind": kind}
                bl = [{"sheet": s, "heading": h, "first": f, "last": l, "rows": json.loads(rs), "signature": json.loads(sg),
                       "labels": json.loads(lb)} for _i, s, h, f, l, rs, sg, lb in
                      db.execute("SELECT * FROM struct_blocks ORDER BY id")]
                cp = [{"a": a, "b": b, "a_heading": ah, "b_heading": bh, "similar": sim, "rows": n}
                      for a, b, ah, bh, sim, n in db.execute("SELECT * FROM struct_copies")]
                got = {"info": info, "blocks": bl, "copies": cp}
    except sqlite3.Error:
        got = None
    if got is None:
        info = rows(db_path)
        bl = blocks(info)
        got = {"info": info, "blocks": bl, "copies": copies(bl)}
    if len(_CACHE) > 32:
        _CACHE.clear()
    _CACHE[key] = got
    return got


def edges(db_path: str) -> tuple[dict, dict]:
    """(reads, read_by): {(sheet, row): {(sheet, row), ...}} within the workbook; inactive lookup options left out."""
    reads, read_by = defaultdict(set), defaultdict(set)
    with _ro(db_path) as db:
        for ss, sr, ds, dr, kind in db.execute("SELECT src_sheet, src_row, dst_sheet, dst_row, kind FROM edges"):
            if kind == "inactive" or (ss, sr) == (ds, dr):
                continue
            reads[(ss, sr)].add((ds, dr))
            read_by[(ds, dr)].add((ss, sr))
    return reads, read_by


def ext_reads(overlay_db: str, link: int | None = None) -> dict:
    """{client (sheet, row): {overlay (sheet, row), ...}}: the overlay's rows reading each client row (a link)."""
    out = defaultdict(set)
    with _ro(overlay_db) as db:
        have = {r[0] for r in db.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        if "extrefs" not in have:
            return out
        for s, r, idx, es, er in db.execute("SELECT sheet, row, idx, ext_sheet, ext_row FROM extrefs"):
            if link is None or idx == link:
                out[(es, er)].add((s, r))
    return out


def closure(start, graph: dict, limit: int = 2000) -> dict:
    """Every row reached from start through graph, with the step it came from: {row: previous row}."""
    seen, q = {start: None}, deque([start])
    while q and len(seen) < limit:
        x = q.popleft()
        for y in graph.get(x, ()):
            if y not in seen:
                seen[y] = x
                q.append(y)
    return seen


def path(seen: dict, end) -> list:
    out = []
    while end is not None:
        out.append(end)
        end = seen.get(end)
    return out[::-1]


def lineage(client_db: str, row: tuple, overlay_db: str | None = None, link: int | None = None,
            targets: set | None = None, info: dict | None = None) -> dict:
    """A client row's lineage: upstream (what it's worked out from, its inputs at the ends), downstream within the
    client model, the overlay rows that read it (or any row downstream of it), and, given the rows of the equity
    value in the overlay (targets), whether its figures reach the value and by which path."""
    info = info or rows(client_db)
    reads, read_by = edges(client_db)
    up = closure(row, reads)
    down = closure(row, read_by)
    inputs = [k for k in up if k != row and not reads.get(k)]
    out = {"row": row, "label": info.get(row, {}).get("label"), "kind": info.get(row, {}).get("kind"),
           "upstream": len(up) - 1, "inputs": [(k, info.get(k, {}).get("label"), info.get(k, {}).get("kind")) for k in inputs[:12]],
           "downstream": len(down) - 1, "reaches_value": None, "via": None}
    if overlay_db and overlay_db == client_db:  # the overlay inside the client model's workbook: one graph, no link
        if targets:
            end = next((t_ for t_ in targets if t_ in down), None)
            out["reaches_value"], out["via"] = bool(end), (path(down, end) if end else None)
        out["overlay_reads"] = sorted(read_by.get(row, ()))[:8]
        return out
    if overlay_db:
        ext = ext_reads(overlay_db, link)
        o_reads, o_read_by = edges(overlay_db)
        hits = [(k, o) for k in down for o in ext.get(k, ())]
        out["overlay_reads"] = sorted({o for _, o in hits})
        if targets:
            for k, o in hits:
                seen = closure(o, o_read_by)
                end = next((t for t in targets if t in seen), None)
                if end:
                    out["reaches_value"], out["via"] = True, path(down, k) + ["(the overlay)"] + path(seen, end)
                    break
            else:
                out["reaches_value"] = False
    return out


def upstream_tree(db_path: str, row: tuple, info: dict | None = None, depth: int = 3) -> list[str]:
    """A row's upstream as indented lines (label, kind), a few levels: what it's worked out from."""
    info = info or rows(db_path)
    reads, _ = edges(db_path)
    lines, seen = [], set()

    def walk(k, d):
        x = info.get(k, {})
        lines.append(f"{'  ' * d}{k[0]}!r{k[1]} {x.get('label') or '?'} [{x.get('kind', '?')}]")
        if d >= depth or k in seen:
            return
        seen.add(k)
        for y in sorted(reads.get(k, ()))[:8]:
            walk(y, d + 1)
    walk(row, 0)
    return lines


def text_map(info: dict, bl: list[dict], cp: list[dict], max_rows: int = 400) -> str:
    """The structure in a few lines: per sheet, each block (its heading and rows by kind), and the copies."""
    lines, n = [], 0
    by_sheet = defaultdict(list)
    for b in bl:
        by_sheet[b["sheet"]].append(b)
    for s, bs in by_sheet.items():
        lines.append(f"SHEET {s}: {len(bs)} block(s)")
        for b in bs:
            kinds = defaultdict(int)
            for r in b["rows"]:
                kinds[info[(s, r)]["kind"]] += 1
            lines.append(f"  block r{b['first']}-r{b['last']} '{b['heading'] or '(no heading)'}': "
                         + ", ".join(f"{v} {k}" for k, v in sorted(kinds.items(), key=lambda kv: -kv[1])))
            for r in b["rows"]:
                if n >= max_rows:
                    break
                x = info[(s, r)]
                lines.append(f"    r{r} {x['label'] or '?'} [{x['kind']}]" + (f" {x['units']}" if x["units"] else ""))
                n += 1
    if cp:
        lines.append("COPIES (blocks alike enough that a label in one is in the other too):")
        lines += [f"  {c['a']} '{c['a_heading']}' ~ {c['b']} '{c['b_heading']}' ({c['similar']:.0%})" for c in cp]
    return "\n".join(lines)
