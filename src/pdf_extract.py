"""
PDF extraction for document comparison.

First-principles design:
- PDFs have no paragraph structure; visual lines must be re-assembled into
  paragraphs using geometry (vertical gaps, line width, indentation) so that
  re-wrapped text does not produce spurious diffs.
- Repeated headers/footers are detected by normalised-text frequency across
  pages (position-restricted) and removed, rather than by hard-coded patterns.
- Ruled tables are extracted as structured blocks when detection looks sane;
  otherwise the region falls back to plain text flow (graceful degradation).
"""

import re
import statistics
import html as html_mod


def _esc(s):
    return html_mod.escape(s or "")


# ── Table sanity ──────────────────────────────────────────────

def _clean_table(rows):
    """None→'', strip cells, drop fully-empty rows/columns."""
    rows = [[(c or "").strip() for c in r] for r in rows if r]
    rows = [r for r in rows if any(r)]
    if not rows:
        return []
    ncols = max(len(r) for r in rows)
    rows = [r + [""] * (ncols - len(r)) for r in rows]
    keep = [j for j in range(ncols) if any(r[j] for r in rows)]
    return [[r[j] for j in keep] for r in rows]


def _is_real_table(rows):
    """Reject chart/figure misdetections: need a minimally filled 2-D grid."""
    if not rows or len(rows) < 2:
        return False
    ncols = max(len(r) for r in rows)
    if ncols < 2:
        return False
    cells = [c for r in rows for c in r]
    nonempty = sum(1 for c in cells if c and str(c).strip())
    return nonempty >= 6 and nonempty / max(1, len(cells)) >= 0.25


# ── Header / footer detection ─────────────────────────────────

def _line_signature(text):
    """Normalise a line so 'Page 15' and 'Page 16' share a signature."""
    s = re.sub(r"\d+", "#", text)
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def _find_repeating_bands(pages):
    """Signatures of top/bottom-band lines that repeat across many pages."""
    from collections import Counter
    counts = Counter()
    for height, lines in pages:
        seen = set()
        for ln in lines:
            if ln["top"] < 0.08 * height or ln["bottom"] > 0.92 * height:
                sig = _line_signature(ln["text"])
                if sig and sig not in seen:
                    seen.add(sig)
                    counts[sig] += 1
    npages = max(1, len(pages))
    threshold = max(3, int(0.3 * npages))
    return {sig for sig, c in counts.items() if c >= threshold}


# ── Line classification helpers ───────────────────────────────

_BULLET_RE = re.compile(r"^\s*(?:[•▪◦‣·○*–—-]|\(?[a-zivx]{1,4}\)|\d{1,2}[.)])\s+")
_HEADING_NUM_RE = re.compile(r"^\d+(?:\.\d+)*\.?\s+\S")
_TERMINAL = (".", "!", "?", ":", ";")


def _line_size(ln):
    sizes = [c.get("size", 0) for c in ln.get("chars", []) if c.get("size")]
    return statistics.median(sizes) if sizes else 0.0


def _is_heading(text, size, body_size):
    if len(text) > 120 or text.endswith(_TERMINAL[:2]):
        return False
    if body_size and size >= body_size * 1.15:
        return True
    if _HEADING_NUM_RE.match(text) and len(text) < 90 and not text.rstrip().endswith("."):
        return True
    return False


def _heading_level(text, size, body_size):
    m = re.match(r"^(\d+(?:\.\d+)*)", text)
    if m:
        return min(m.group(1).count(".") + 1, 3)
    if body_size and size >= body_size * 1.35:
        return 1
    return 2


# ── Main extraction ───────────────────────────────────────────

def extract_pdf_blocks(filepath, block_cls):
    """Return list of block_cls(kind, text, html, level, page, rows)."""
    import pdfplumber

    pages = []          # (height, [line dict])
    page_tables = []    # [ (top, rows) ] per page

    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            tables = []
            try:
                for t in page.find_tables():
                    rows = _clean_table(t.extract())
                    if _is_real_table(rows):
                        tables.append((t.bbox, rows))
            except Exception:
                tables = []
            try:
                lines = page.extract_text_lines()
            except Exception:
                lines = []
            kept = []
            for ln in lines:
                txt = (ln.get("text") or "").strip()
                if len(txt) <= 2 or not re.search(r"[A-Za-z0-9]", txt):
                    continue  # chart debris / decorative fragments
                cx = (ln["x0"] + ln["x1"]) / 2
                cy = (ln["top"] + ln["bottom"]) / 2
                if any(bx0 <= cx <= bx1 and by0 <= cy <= by1
                       for (bx0, by0, bx1, by1), _ in tables):
                    continue  # text belongs to an extracted table
                ln["text"] = txt
                kept.append(ln)
            kept.sort(key=lambda l: (round(l["top"], 1), l["x0"]))
            pages.append((page.height, kept))
            page_tables.append([(bbox[1], rows) for bbox, rows in tables])

    repeat_sigs = _find_repeating_bands(pages)

    # Document-wide body font size
    all_sizes = [_line_size(ln) for _, lines in pages for ln in lines]
    all_sizes = [s for s in all_sizes if s]
    body_size = statistics.median(all_sizes) if all_sizes else 0.0

    blocks = []

    for page_num, ((height, lines), tables) in enumerate(
            zip(pages, page_tables), start=1):

        lines = [ln for ln in lines
                 if not ((ln["top"] < 0.08 * height or
                          ln["bottom"] > 0.92 * height) and
                         _line_signature(ln["text"]) in repeat_sigs)]

        # Typical text edges on this page (15th/85th percentile of x extents)
        x1s = sorted(ln["x1"] for ln in lines)
        right_edge = x1s[int(0.85 * (len(x1s) - 1))] if x1s else 0
        x0s = sorted(ln["x0"] for ln in lines)
        left_edge = x0s[int(0.15 * (len(x0s) - 1))] if x0s else 0
        body_width = max(1.0, right_edge - left_edge)

        items = []  # (top, payload) for ordering text & tables on the page

        cur = None          # accumulating paragraph dict
        prev_ln = None

        def flush():
            nonlocal cur
            if cur:
                items.append((cur["top"], ("para", cur["text"])))
                cur = None

        for ln in lines:
            text, size = ln["text"], _line_size(ln)
            line_h = max(1.0, ln["bottom"] - ln["top"])

            if _is_heading(text, size, body_size):
                flush()
                lvl = _heading_level(text, size, body_size)
                items.append((ln["top"], ("heading", text, lvl)))
                prev_ln = ln
                continue

            starts_new = False
            if _BULLET_RE.match(text):
                starts_new = True
            elif cur is None:
                starts_new = True
            elif prev_ln is not None:
                gap = ln["top"] - prev_ln["bottom"]
                prev_short = prev_ln["x1"] < right_edge - 0.12 * body_width
                indented = ln["x0"] > left_edge + 0.04 * body_width and \
                    cur["x0"] <= left_edge + 0.02 * body_width
                if gap > 0.65 * line_h or prev_short or indented:
                    starts_new = True

            if starts_new:
                flush()
                cur = {"top": ln["top"], "text": text, "x0": ln["x0"]}
            else:
                if cur["text"].endswith("-") and text[:1].islower():
                    cur["text"] = cur["text"][:-1] + text
                else:
                    cur["text"] += " " + text
            prev_ln = ln

        flush()

        for top, rows in tables:
            items.append((top, ("table", rows)))

        items.sort(key=lambda it: it[0])

        for _top, item in items:
            kind = item[0]
            if kind == "heading":
                _, text, lvl = item
                h = min(lvl + 1, 4)
                blocks.append(block_cls(
                    kind="heading", text=text,
                    html=f"<h{h}>{_esc(text)}</h{h}>",
                    level=lvl, page=page_num))
            elif kind == "para":
                text = item[1]
                k = "list" if _BULLET_RE.match(text) else "para"
                tag = "li" if k == "list" else "p"
                blocks.append(block_cls(
                    kind=k, text=text,
                    html=f"<{tag}>{_esc(text)}</{tag}>", page=page_num))
            else:
                rows = item[1]
                html_rows = []
                for i, r in enumerate(rows):
                    tn = "th" if i == 0 else "td"
                    html_rows.append("<tr>" + "".join(
                        f"<{tn}>{_esc(c)}</{tn}>" for c in r) + "</tr>")
                full_text = " | ".join(
                    " ".join(c for c in r if c) for r in rows)
                blocks.append(block_cls(
                    kind="table", text=full_text,
                    html="<table>" + "".join(html_rows) + "</table>",
                    rows=rows, page=page_num))

    # Merge paragraphs split across page breaks (conservative)
    merged = []
    for b in blocks:
        if (merged and b.kind == "para" and merged[-1].kind == "para"
                and b.page == merged[-1].page + 1
                and not merged[-1].text.rstrip().endswith(_TERMINAL)
                and b.text[:1].islower()):
            prev = merged[-1]
            if prev.text.endswith("-"):
                prev.text = prev.text[:-1] + b.text
            else:
                prev.text += " " + b.text
            prev.html = f"<p>{_esc(prev.text)}</p>"
        else:
            merged.append(b)
    return merged
