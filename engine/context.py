"""What the report says around a figure: context for every model that reads a table or judges a figure.

A figure means what the report around it says: the entity and the interest valued, the basis (ex- or cum-distribution,
nominal or real, pre- or post-tax), the valuation date, the units, what a column heading stands for. So the models that
read a key table, review a fact, or check one on its image get, besides the figure itself:
  the page        the text on the page the figure sits on: its section heading, the table's caption, the notes and
                  sources beneath it (a table's own rows left out when it's the table being read)
  the letter      the report's transmittal letter: its first pages addressed to the client ("Dear ...", "Yours
                  faithfully"); without one, the executive summary
  the scope       the sections on the engagement's scope: purpose, interest valued, basis and standard of value,
                  valuation date, sources, reliance and limitations
  definitions     a glossary or defined terms, where the report has one
  the facts       what's established so far: target, client, valuation date, currency and units, the equity basis
Found in code from the report's own text (its headings and words), no model calls; each part is capped so the prompts
stay small. The models are told it's context, not a source of figures: they read figures from the image or the cited
text.
"""
import re

LETTER = re.compile(r"\bdear\b|yours (?:faithfully|sincerely)|private (?:and|&) confidential|\bthe directors\b", re.I)
LETTER_END = re.compile(r"yours (?:faithfully|sincerely)|kind regards", re.I)
SCOPE = re.compile(r"scope|terms of reference|purpose of (?:the|this|our) (?:report|valuation)|basis of (?:value|valuation)|"
                   r"interest valued|instructions|standard of value|premise of value|sources? of information|reliance|"
                   r"limitations|valuation (?:date|approach and scope)", re.I)
DEFINITIONS = re.compile(r"glossary|definitions|abbreviations|defined terms", re.I)
SUMMARY = re.compile(r"executive summary|summary", re.I)
CAPS = {"letter": 3000, "scope": 3000, "definitions": 1500, "page": 2500}
BASIS = re.compile(r"[^.]*\b(?:ex|cum)[- ]?(?:distribution|dividend|div)\b[^.]*\.", re.I)


def pages(markdown: str) -> dict[int, str]:
    parts = re.split(r"<!-- page (\d+) -->\n", markdown or "")
    return {int(parts[i]): parts[i + 1] for i in range(1, len(parts) - 1, 2)}


def _clean(text: str) -> str:
    text = re.sub(r"<!--.*?-->", " ", text or "")
    text = re.sub(r"^## (?:Page|Slide) \d+\s*$", "", text, flags=re.M)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def sections(page_md: str) -> list[tuple[str, str]]:
    """A page's sections: (heading, text under it until the next heading); text before the first heading under ""."""
    out, head, buf = [], "", []
    for line in page_md.splitlines():
        m = re.match(r"^#{2,4}\s+(.*)$", line)
        if m and not re.match(r"^(?:Page|Slide) \d+$", m[1].strip()):
            if buf or head:
                out.append((head, "\n".join(buf).strip()))
            head, buf = m[1].strip(), []
        else:
            buf.append(line)
    out.append((head, "\n".join(buf).strip()))
    return [(h, t) for h, t in out if h or t]


def _cap(text: str, n: int) -> str:
    return text if len(text) <= n else text[:n].rsplit(" ", 1)[0] + " …"


def find(markdown: str) -> dict:
    """{"letter"|"summary", "scope", "definitions": {"pages": [n], "text"}}, each only where the report has it."""
    pg = pages(markdown)
    out = {}
    first = sorted(pg)[:8]
    start = next((n for n in first if len(set(m.lower() for m in LETTER.findall(pg[n]))) >= 1
                  and re.search(r"\bdear\b", pg[n], re.I)), None)
    if start is not None:
        keep = [start]
        while not LETTER_END.search(pg[keep[-1]]) and keep[-1] + 1 in pg and len(keep) < 3:
            keep.append(keep[-1] + 1)
        out["letter"] = {"pages": keep, "text": _cap("\n\n".join(_clean(pg[n]) for n in keep), CAPS["letter"])}
    for key, rx in (("scope", SCOPE), ("definitions", DEFINITIONS)):
        found, texts = [], []
        for n in sorted(pg):
            for h, t in sections(pg[n]):
                if h and rx.search(h) and (key != "scope" or not DEFINITIONS.search(h)):
                    found.append(n)
                    texts.append(f"{h}\n{_clean(t)}")
        if texts:
            out[key] = {"pages": sorted(set(found)), "text": _cap("\n\n".join(texts), CAPS[key])}
    if "letter" not in out:  # no letter: the executive summary says much the same
        for n in sorted(pg):
            hit = next((f"{h}\n{_clean(t)}" for h, t in sections(pg[n]) if h and SUMMARY.search(h)), None)
            if hit:
                out["summary"] = {"pages": [n], "text": _cap(hit, CAPS["letter"])}
                break
    return out


def pages_for(ctx: dict) -> list[int]:
    return sorted({n for part in ctx.values() for n in part.get("pages", [])})


def page_text(markdown: str, page: int | None, table_id: str | None = None) -> str:
    """The text on a page, a table's own block left out when it's the table being read or judged."""
    text = pages(markdown).get(page or 0, "")
    if table_id:
        text = re.sub(rf"<!-- table {re.escape(table_id)} \(\w+\) -->\n(?:\*\*.*?\*\*\n\n)?.*?(?=\n\n(?!\|)|\Z)", "[the table]",
                      text, flags=re.S)
        text = re.sub(rf"\*\[Table {re.escape(table_id)}:[^\]]*\]\*", "[the table]", text)
    return _cap(_clean(text), CAPS["page"])


def established(facts: list[dict] | None) -> str:
    """What's established so far, in a few lines: who, when, units, the equity value's basis."""
    if not facts:
        return ""
    want = {"target_name": "Target", "client": "Client", "valuation_date": "Valuation date", "currency_units": "Units"}
    lines = [f"{label}: {f.get('value_text')}" for key, label in want.items()
             for f in facts[:200] if f.get("key") == key and f.get("value_text") and f.get("status") != "rejected"][:4]
    eq = next((f for f in facts if f.get("key") == "equity_value" and f.get("basis")), None)
    if eq:
        lines.append(f"Equity value basis: {eq['basis']}")
    return "\n".join(lines)


def masked(text: str) -> str:
    """The text with its dates and figures masked ("[a date]", "#"): what things are, without the answer. For a
    read that must be blind to the text layer's figures."""
    import keyfacts
    for rx, _kind in keyfacts._DATES:
        text = rx.sub("[a date]", text)
    return re.sub(r"\d[\d,]*(?:\.\d+)?", "#", text)


def block(markdown: str, ctx: dict | None = None, page: int | None = None, table_id: str | None = None,
          facts: list[dict] | None = None, figures_from: str = "the image", blind: bool = False) -> str:
    """The context as one block for a prompt. blind: its figures and dates masked (masked()), for the read of an
    image that is checked against the text layer."""
    ctx = ctx if ctx is not None else find(markdown)
    hide = masked if blind else (lambda t: t)
    parts = []
    est = hide(established(facts))
    if est:
        parts.append(f"Established so far:\n{est}")
    if page:
        parts.append(f"The text on page {page}" + (" (the table itself left out)" if table_id else "")
                     + (" (its figures and dates masked)" if blind else "") + ":\n"
                     + (hide(page_text(markdown, page, table_id)) or "(nothing but the table)"))
    names = {"letter": "The report's transmittal letter", "summary": "The report's executive summary",
             "scope": "The scope of the engagement", "definitions": "Definitions"}
    for key in ("letter", "summary", "scope", "definitions"):
        if key in ctx and not (page and ctx[key]["pages"] == [page]):
            said = "\n".join(parts)
            if hide(ctx[key]["text"])[:160] in said:  # a scope inside the letter: said once
                continue
            parts.append(f"{names[key]} (page {', '.join(map(str, ctx[key]['pages']))}):\n{hide(ctx[key]['text'])}")
    if not parts:
        return ""
    basis = sorted({m.strip() for m in BASIS.findall(markdown or "")})[:3]
    if basis:
        parts.append("Where the report states a basis:\n" + "\n".join(f"- {hide(_cap(b, 240))}" for b in basis))
    return ("Context from the report, to judge what the figures mean (the entity, the interest valued, the basis, the "
            f"date, the units, what a heading stands for). It is context only: take figures from {figures_from}, not "
            "from here.\n\n" + "\n\n".join(parts))
