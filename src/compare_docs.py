#!/usr/bin/env python3
"""
Document Comparator v3 — Dense side-by-side, word-level diff
Matches reference UI: status badges, match confidence, page metadata, section dividers.
Supports: .docx, .txt, .pdf
Usage: python compare_docs.py <file1> <file2> [--output report.html]
"""

import sys, os, re, difflib, argparse, html as html_mod
from datetime import datetime
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional


# ══════════════════════════════════════════════════════════════
#  DATA STRUCTURES
# ══════════════════════════════════════════════════════════════

@dataclass
class Block:
    kind:    str          # 'para' | 'table' | 'heading' | 'list' | 'spacer'
    text:    str          # plain text (for diffing)
    html:    str          # rendered HTML (structure preserved)
    level:   int = 0      # heading level
    page:    int = 0      # approximate source page (1-based, 0 = unknown)
    rows:    list = field(default_factory=list)   # table rows: [[cell, ...], ...]


# ══════════════════════════════════════════════════════════════
#  TEXT EXTRACTION  →  List[Block]
# ══════════════════════════════════════════════════════════════

def _esc(s: str) -> str:
    return html_mod.escape(s or "")


def extract_blocks_docx(filepath: str) -> list[Block]:
    try:
        from docx import Document
        from docx.text.paragraph import Paragraph as DocxPara
        from docx.table import Table as DocxTable
        from docx.oxml.ns import qn
    except ImportError:
        sys.exit("ERROR: python-docx not installed. Run: pip install python-docx")

    doc = Document(filepath)
    blocks: list[Block] = []
    body = doc.element.body

    # Rough page estimation: ~40 paragraphs per page
    para_count = 0
    PARAS_PER_PAGE = 40

    for child in body:
        tag = child.tag.split("}")[1] if "}" in child.tag else child.tag

        if tag == "p":
            para = DocxPara(child, doc)
            text = para.text.strip()
            para_count += 1
            page = max(1, (para_count // PARAS_PER_PAGE) + 1)

            if not text:
                blocks.append(Block(kind="spacer", text="", html="", page=page))
                continue

            style = (para.style.name or "").lower()
            if "heading" in style:
                try:    lvl = int(style.split()[-1])
                except: lvl = 1
                h = min(lvl + 1, 4)
                blocks.append(Block(kind="heading", text=text,
                                    html=f"<h{h}>{_esc(text)}</h{h}>",
                                    level=lvl, page=page))
            elif "list" in style or para._element.find(qn("w:numPr")) is not None:
                blocks.append(Block(kind="list", text=text,
                                    html=f"<li>{_esc(text)}</li>", page=page))
            else:
                blocks.append(Block(kind="para", text=text,
                                    html=f"<p>{_esc(text)}</p>", page=page))

        elif tag == "tbl":
            from docx.table import Table as DocxTable
            tbl = DocxTable(child, doc)
            para_count += len(tbl.rows) * 2
            page = max(1, (para_count // PARAS_PER_PAGE) + 1)
            rows_data, html_rows = [], []
            for i, row in enumerate(tbl.rows):
                cells = [c.text.strip() for c in row.cells]
                rows_data.append(cells)
                tn = "th" if i == 0 else "td"
                html_rows.append("<tr>" + "".join(
                    f"<{tn}>{_esc(c)}</{tn}>" for c in cells) + "</tr>")
            table_html = "<table>" + "".join(html_rows) + "</table>"
            full_text = " | ".join(" ".join(c for c in r if c) for r in rows_data)
            blocks.append(Block(kind="table", text=full_text, html=table_html,
                                rows=rows_data, page=page))
    return blocks


def extract_blocks_txt(filepath: str) -> list[Block]:
    with open(filepath, "r", encoding="utf-8", errors="replace") as f:
        raw = f.read()
    blocks: list[Block] = []
    line_num = 0
    prev_spacer = False
    LINES_PER_PAGE = 50
    for line in raw.splitlines():
        line_num += 1
        page = max(1, (line_num // LINES_PER_PAGE) + 1)
        stripped = line.strip()
        if not stripped:
            if not prev_spacer:
                blocks.append(Block(kind="spacer", text="", html="", page=page))
            prev_spacer = True
        else:
            blocks.append(Block(kind="para", text=stripped,
                                html=f"<p>{_esc(stripped)}</p>", page=page))
            prev_spacer = False
    return blocks


def extract_blocks_pdf(filepath: str) -> list[Block]:
    """PDF extraction: reconstructs paragraphs from line geometry, strips
    repeated headers/footers, extracts ruled tables. See pdf_extract.py."""
    try:
        import pdfplumber  # noqa: F401
    except ImportError:
        sys.exit("ERROR: pdfplumber not installed. Run: pip install pdfplumber")
    from pdf_extract import extract_pdf_blocks
    return extract_pdf_blocks(filepath, Block)


def extract_blocks(filepath: str) -> list[Block]:
    ext = Path(filepath).suffix.lower()
    if ext == ".docx": return extract_blocks_docx(filepath)
    if ext == ".txt":  return extract_blocks_txt(filepath)
    if ext == ".pdf":  return extract_blocks_pdf(filepath)
    sys.exit(f"ERROR: Unsupported file type '{ext}'. Supported: .txt, .docx, .pdf")


# ══════════════════════════════════════════════════════════════
#  WORD-LEVEL DIFF
# ══════════════════════════════════════════════════════════════

def _tokenize(text: str) -> list[str]:
    return re.findall(r'\S+|\s+', text)


def word_diff(text_a: str, text_b: str) -> tuple[str, str]:
    toks_a = _tokenize(text_a)
    toks_b = _tokenize(text_b)
    sm = difflib.SequenceMatcher(None, toks_a, toks_b, autojunk=False)
    left_parts, right_parts = [], []

    for op, i1, i2, j1, j2 in sm.get_opcodes():
        seg_a = "".join(toks_a[i1:i2])
        seg_b = "".join(toks_b[j1:j2])
        ea, eb = _esc(seg_a), _esc(seg_b)

        if op == "equal":
            left_parts.append(ea)
            right_parts.append(eb)
        elif op == "replace":
            if seg_a.strip(): left_parts.append(f"<mark class='w-del'>{ea}</mark>")
            else:              left_parts.append(ea)
            if seg_b.strip(): right_parts.append(f"<mark class='w-ins'>{eb}</mark>")
            else:              right_parts.append(eb)
        elif op == "delete":
            if seg_a.strip(): left_parts.append(f"<mark class='w-del'>{ea}</mark>")
            else:              left_parts.append(ea)
        elif op == "insert":
            if seg_b.strip(): right_parts.append(f"<mark class='w-ins'>{eb}</mark>")
            else:              right_parts.append(eb)

    return "".join(left_parts), "".join(right_parts)


def _inject_word_diff(ba: Block, bb: Block) -> tuple[str, str]:
    """Inject word-diff into the HTML wrappers of two matched blocks."""
    if ba.kind == "table" and bb.kind == "table":
        rows_a, rows_b = ba.rows, bb.rows
        ha_rows, hb_rows = [], []

        def _pair_rows(ca_list, cb_list, tn):
            max_c = max(len(ca_list), len(cb_list))
            row_a, row_b = [], []
            for j in range(max_c):
                ca = ca_list[j] if j < len(ca_list) else ""
                cb = cb_list[j] if j < len(cb_list) else ""
                da, db = word_diff(ca, cb)
                row_a.append(f"<{tn}>{da}</{tn}>")
                row_b.append(f"<{tn}>{db}</{tn}>")
            ha_rows.append("<tr>" + "".join(row_a) + "</tr>")
            hb_rows.append("<tr>" + "".join(row_b) + "</tr>")

        def _solo_row(cells, tn, side):
            cls = "w-del" if side == "a" else "w-ins"
            h = "<tr>" + "".join(
                f"<{tn}><mark class='{cls}'>{_esc(c)}</mark></{tn}>"
                for c in cells) + "</tr>"
            (ha_rows if side == "a" else hb_rows).append(h)

        # Align rows by content so added/removed rows don't shift the rest
        key_a = [" ".join(r) for r in rows_a]
        key_b = [" ".join(r) for r in rows_b]
        rsm = difflib.SequenceMatcher(None, key_a, key_b, autojunk=False)
        for op, i1, i2, j1, j2 in rsm.get_opcodes():
            if op == "equal":
                for i, j in zip(range(i1, i2), range(j1, j2)):
                    _pair_rows(rows_a[i], rows_b[j],
                               "th" if (i == 0 or j == 0) else "td")
            elif op == "replace":
                n = min(i2 - i1, j2 - j1)
                for k in range(n):
                    _pair_rows(rows_a[i1 + k], rows_b[j1 + k],
                               "th" if (i1 + k == 0 or j1 + k == 0) else "td")
                for i in range(i1 + n, i2):
                    _solo_row(rows_a[i], "th" if i == 0 else "td", "a")
                for j in range(j1 + n, j2):
                    _solo_row(rows_b[j], "th" if j == 0 else "td", "b")
            elif op == "delete":
                for i in range(i1, i2):
                    _solo_row(rows_a[i], "th" if i == 0 else "td", "a")
            elif op == "insert":
                for j in range(j1, j2):
                    _solo_row(rows_b[j], "th" if j == 0 else "td", "b")

        return ("<table>" + "".join(ha_rows) + "</table>",
                "<table>" + "".join(hb_rows) + "</table>")

    da, db = word_diff(ba.text, bb.text)
    if ba.kind == "heading":
        h = min(ba.level + 1, 4)
        return f"<h{h}>{da}</h{h}>", f"<h{h}>{db}</h{h}>"
    if ba.kind == "list":
        return f"<li>{da}</li>", f"<li>{db}</li>"
    return f"<p>{da}</p>", f"<p>{db}</p>"


# ══════════════════════════════════════════════════════════════
#  BLOCK ALIGNMENT  →  list of DiffRow
# ══════════════════════════════════════════════════════════════

SIMILARITY_THRESHOLD = 0.35


@dataclass
class DiffRow:
    status:     str            # 'equal'|'changed'|'deleted'|'inserted'|'spacer'
    left_html:  str
    right_html: str
    page_a:     int = 0
    page_b:     int = 0
    confidence: int = 100      # match confidence 0-100 (for 'changed' rows)


def align_blocks(blocks_a: list[Block], blocks_b: list[Block]) -> list[DiffRow]:
    rows: list[DiffRow] = []
    texts_a = [b.text for b in blocks_a]
    texts_b = [b.text for b in blocks_b]
    sm = difflib.SequenceMatcher(None, texts_a, texts_b, autojunk=False)

    for op, i1, i2, j1, j2 in sm.get_opcodes():

        if op == "equal":
            for ai, bi in zip(range(i1, i2), range(j1, j2)):
                ba, bb = blocks_a[ai], blocks_b[bi]
                if ba.kind == "spacer":
                    rows.append(DiffRow("spacer", "", "", ba.page, bb.page))
                else:
                    rows.append(DiffRow("equal", ba.html, bb.html, ba.page, bb.page))

        elif op == "replace":
            seg_a = blocks_a[i1:i2]
            seg_b = blocks_b[j1:j2]

            # Build best pairings by similarity.
            # One SequenceMatcher per b-block (seq2 preprocessing is cached);
            # quick_ratio gates skip most full ratio() calls — keeps large
            # replace segments fast.
            matchers = []
            for bb in seg_b:
                m = difflib.SequenceMatcher(None, "", bb.text, autojunk=False)
                matchers.append(m)

            used_b: set[int] = set()
            pairs: list[tuple[int, int, float]] = []
            for ia, ba in enumerate(seg_a):
                best_r, best_ib = 0.0, -1
                for ib, bb in enumerate(seg_b):
                    if ib in used_b: continue
                    if ba.kind != bb.kind and not (
                        ba.kind in ("para","list","heading") and
                        bb.kind in ("para","list","heading")): continue
                    m = matchers[ib]
                    m.set_seq1(ba.text)
                    floor = max(best_r, SIMILARITY_THRESHOLD)
                    if m.real_quick_ratio() <= floor or m.quick_ratio() <= floor:
                        continue
                    r = m.ratio()
                    if r > best_r:
                        best_r, best_ib = r, ib
                if best_r >= SIMILARITY_THRESHOLD and best_ib >= 0:
                    pairs.append((ia, best_ib, best_r))
                    used_b.add(best_ib)

            paired_a = {p[0] for p in pairs}
            paired_b = {p[1] for p in pairs}
            pairs.sort(key=lambda x: x[0])

            ia_cursor = ib_cursor = pair_idx = 0

            while ia_cursor < len(seg_a) or ib_cursor < len(seg_b):
                if pair_idx < len(pairs) and pairs[pair_idx][0] == ia_cursor:
                    ia, ib, ratio = pairs[pair_idx]
                    ba, bb = seg_a[ia], seg_b[ib]

                    # Emit unmatched b blocks before this ib
                    while ib_cursor < ib:
                        if ib_cursor not in paired_b:
                            bb2 = seg_b[ib_cursor]
                            if bb2.kind != "spacer":
                                rows.append(DiffRow("inserted", "", bb2.html,
                                                    0, bb2.page))
                        ib_cursor += 1
                    ib_cursor = ib + 1

                    if ba.kind == "spacer":
                        rows.append(DiffRow("spacer", "", "", ba.page, bb.page))
                    elif ratio >= 0.999:
                        rows.append(DiffRow("equal", ba.html, bb.html,
                                            ba.page, bb.page))
                    else:
                        lh, rh = _inject_word_diff(ba, bb)
                        conf = round(ratio * 100)
                        rows.append(DiffRow("changed", lh, rh,
                                            ba.page, bb.page, conf))
                    ia_cursor += 1
                    pair_idx += 1

                elif ia_cursor < len(seg_a) and ia_cursor not in paired_a:
                    ba = seg_a[ia_cursor]
                    if ba.kind != "spacer":
                        # For deleted: show struck-through full text on left
                        rows.append(DiffRow("deleted", ba.html, "",
                                            ba.page, 0))
                    ia_cursor += 1
                elif ia_cursor < len(seg_a):
                    ia_cursor += 1
                else:
                    if ib_cursor < len(seg_b) and ib_cursor not in paired_b:
                        bb = seg_b[ib_cursor]
                        if bb.kind != "spacer":
                            rows.append(DiffRow("inserted", "", bb.html,
                                                0, bb.page))
                    ib_cursor += 1

            while ib_cursor < len(seg_b):
                if ib_cursor not in paired_b:
                    bb = seg_b[ib_cursor]
                    if bb.kind != "spacer":
                        rows.append(DiffRow("inserted", "", bb.html,
                                            0, bb.page))
                ib_cursor += 1

        elif op == "delete":
            for ba in blocks_a[i1:i2]:
                if ba.kind != "spacer":
                    rows.append(DiffRow("deleted", ba.html, "", ba.page, 0))

        elif op == "insert":
            for bb in blocks_b[j1:j2]:
                if bb.kind != "spacer":
                    rows.append(DiffRow("inserted", "", bb.html, 0, bb.page))

    return rows


def summarise(rows: list[DiffRow]) -> dict:
    return {
        "changed":  sum(1 for r in rows if r.status == "changed"),
        "deleted":  sum(1 for r in rows if r.status == "deleted"),
        "inserted": sum(1 for r in rows if r.status == "inserted"),
        "equal":    sum(1 for r in rows if r.status == "equal"),
    }


# ══════════════════════════════════════════════════════════════
#  HTML GENERATION
# ══════════════════════════════════════════════════════════════

CSS = """
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
  --bg:          #f7f7f5;
  --surface:     #ffffff;
  --border:      #e2e0da;
  --border-dark: #c8c4bc;
  --text:        #1a1814;
  --text-dim:    #6b6760;
  --text-faint:  #a8a4a0;

  --del-text:    #b91c1c;
  --ins-text:    #15803d;
  --chg-text:    #92400e;

  --meta-bg:     #f0efe9;
  --meta-border: #d8d4cc;

  --header-bg:   #1c1a17;

  --font-body:   'Source Serif 4', Georgia, serif;
  --font-ui:     'Inter', system-ui, sans-serif;
}

body {
  font-family: var(--font-body);
  font-size: 14px;
  line-height: 1.65;
  background:
    radial-gradient(1100px 600px at 12% -10%, rgba(99,102,241,0.28), transparent 60%),
    radial-gradient(900px 600px at 95% 5%, rgba(45,212,191,0.18), transparent 55%),
    radial-gradient(800px 700px at 50% 110%, rgba(168,85,247,0.16), transparent 60%),
    linear-gradient(160deg, #0c1222 0%, #131a2e 55%, #0e1626 100%);
  background-attachment: fixed;
  color: var(--text);
  min-height: 100vh;
}

/* ── Top header ── */
.top-header {
  background: rgba(13, 18, 33, 0.62);
  backdrop-filter: blur(18px) saturate(150%);
  -webkit-backdrop-filter: blur(18px) saturate(150%);
  border-bottom: 1px solid rgba(255,255,255,0.10);
  padding: 14px 24px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  flex-wrap: wrap;
  gap: 10px;
  position: sticky;
  top: 0;
  z-index: 200;
}

.brand {
  display: flex;
  align-items: center;
  gap: 8px;
  font-family: var(--font-ui);
  color: #f5f4f0;
  font-size: 15px;
  font-weight: 700;
}
.brand-icon {
  background: #c9a84c;
  border-radius: 5px;
  width: 28px; height: 28px;
  display: flex; align-items: center; justify-content: center;
  font-size: 14px; font-weight: 800; color: #1c1a17;
}

.header-files {
  display: flex;
  align-items: center;
  gap: 8px;
  font-family: var(--font-ui);
  font-size: 12px;
  color: #a09c98;
}
.file-chip {
  background: rgba(255,255,255,0.08);
  border: 1px solid rgba(255,255,255,0.16);
  border-radius: 8px;
  padding: 3px 10px;
  color: #d8dce6;
  font-size: 11px;
  backdrop-filter: blur(8px);
}

.home-link {
  font-family: var(--font-ui);
  font-size: 11px;
  font-weight: 600;
  color: #cbd5e1;
  text-decoration: none;
  border: 1px solid rgba(255,255,255,0.18);
  background: rgba(255,255,255,0.07);
  border-radius: 8px;
  padding: 5px 12px;
  transition: background .15s;
}
.home-link:hover { background: rgba(255,255,255,0.15); }
.file-chip b { color: #94a3b8; margin-right: 5px; font-size: 10px; letter-spacing: 0.8px; }
.file-chip.base b { color: #f87171; }
.file-chip.rev  b { color: #4ade80; }

/* ── Stats bar ── */
.stats-bar {
  background: rgba(255,255,255,0.06);
  backdrop-filter: blur(14px);
  -webkit-backdrop-filter: blur(14px);
  border-bottom: 1px solid rgba(255,255,255,0.10);
  padding: 10px 24px;
  display: flex;
  gap: 20px;
  flex-wrap: wrap;
  align-items: center;
  font-family: var(--font-ui);
  font-size: 12px;
}

.stat-item {
  display: flex;
  align-items: center;
  gap: 5px;
  color: #aeb6c6;
}
.stat-num { font-weight: 700; font-size: 14px; }
.stat-num.del { color: #f87171; }
.stat-num.ins { color: #4ade80; }
.stat-num.chg { color: #fbbf24; }
.stat-num.eq  { color: #94a3b8; }

.stats-ts {
  margin-left: auto;
  color: #7c8497;
  font-size: 11px;
}

/* ── Column header strip ── */
.col-strip {
  display: grid;
  grid-template-columns: 200px 1fr 1fr;
  background: rgba(20, 27, 45, 0.72);
  backdrop-filter: blur(16px);
  -webkit-backdrop-filter: blur(16px);
  border-bottom: 1px solid rgba(255,255,255,0.12);
  position: sticky;
  top: 57px;
  z-index: 100;
}
.col-strip-meta { border-right: 1px solid rgba(255,255,255,0.08); }
.col-strip-label {
  padding: 7px 16px;
  font-family: var(--font-ui);
  font-size: 10px;
  font-weight: 700;
  letter-spacing: 1px;
  text-transform: uppercase;
  color: #aeb6c6;
  border-right: 1px solid rgba(255,255,255,0.08);
}
.col-strip-label:last-child { border-right: none; }
.col-strip-label.base { color: #f87171; }
.col-strip-label.rev  { color: #4ade80; }

/* ── Section dividers ── */
.section-divider {
  background: rgba(255,255,255,0.07);
  backdrop-filter: blur(10px);
  -webkit-backdrop-filter: blur(10px);
  border-top: 1px solid rgba(255,255,255,0.10);
  border-bottom: 1px solid rgba(255,255,255,0.10);
  padding: 5px 16px;
  font-family: var(--font-ui);
  font-size: 11px;
  font-weight: 600;
  color: #c3cad8;
  letter-spacing: 0.3px;
}

/* ── Diff rows ── */
.diff-row {
  display: grid;
  grid-template-columns: 200px 1fr 1fr;
  border-bottom: 1px solid var(--border);
  background: var(--surface);
}

.diff-row:hover { background: #faf9f6; }

/* Meta column */
.meta-col {
  padding: 10px 14px;
  border-right: 1px solid var(--border);
  font-family: var(--font-ui);
  font-size: 11px;
  line-height: 1.5;
  color: var(--text-dim);
  display: flex;
  flex-direction: column;
  gap: 5px;
  background: var(--meta-bg);
}

.meta-pages { color: var(--text-dim); }
.meta-pages span { color: var(--text-faint); }

/* Status badge */
.badge {
  display: inline-block;
  padding: 2px 8px;
  border-radius: 3px;
  font-size: 10px;
  font-weight: 700;
  letter-spacing: 0.5px;
  text-transform: uppercase;
  width: fit-content;
  font-family: var(--font-ui);
}
.badge-deleted  { background: #fca5a5; color: #7f1d1d; }
.badge-inserted { background: #86efac; color: #14532d; }
.badge-modified { background: #fde68a; color: #78350f; }
.badge-equal    { background: #e2e8f0; color: #475569; }

.confidence {
  color: var(--text-faint);
  font-size: 10px;
  font-family: var(--font-ui);
}
.confidence span { color: var(--text-dim); }

/* Content columns */
.content-col {
  padding: 10px 16px;
  border-right: 1px solid var(--border);
  overflow-wrap: break-word;
  word-break: break-word;
  min-height: 38px;
  vertical-align: top;
}
.content-col:last-child { border-right: none; }

/* Status-specific content styles */
.diff-row.row-deleted .content-col.left {
  color: var(--del-text);
}
.diff-row.row-deleted .content-col.left p,
.diff-row.row-deleted .content-col.left li,
.diff-row.row-deleted .content-col.left h2,
.diff-row.row-deleted .content-col.left h3,
.diff-row.row-deleted .content-col.left h4 {
  text-decoration: line-through;
  text-decoration-color: #f87171;
}
.diff-row.row-deleted .content-col.right {
  background: #fafaf8;
}

.diff-row.row-inserted .content-col.right {
  color: var(--ins-text);
}
.diff-row.row-inserted .content-col.left {
  background: #fafaf8;
}

/* Word-level marks — bold+underline style matching reference */
mark.w-del {
  background: none;
  color: var(--del-text);
  font-weight: 700;
  text-decoration: underline;
  text-decoration-style: solid;
  text-decoration-color: #f87171;
  text-underline-offset: 2px;
  padding: 0;
}

mark.w-ins {
  background: none;
  color: var(--ins-text);
  font-weight: 700;
  text-decoration: underline;
  text-decoration-style: solid;
  text-decoration-color: #4ade80;
  text-underline-offset: 2px;
  padding: 0;
}

/* Content typography */
.content-col p   { margin-bottom: 0; line-height: 1.65; }
.content-col li  { margin-left: 16px; line-height: 1.65; }
.content-col h2  { font-family: var(--font-ui); font-size: 15px; font-weight: 700; line-height: 1.3; }
.content-col h3  { font-family: var(--font-ui); font-size: 13px; font-weight: 700; line-height: 1.3; }
.content-col h4  { font-family: var(--font-ui); font-size: 12px; font-weight: 700; line-height: 1.3; }

.content-col table {
  width: 100%; border-collapse: collapse;
  font-size: 12px; margin: 2px 0;
}
.content-col table th,
.content-col table td {
  border: 1px solid var(--border-dark);
  padding: 4px 8px;
  text-align: left; vertical-align: top;
}
.content-col table th {
  background: #f0ece4;
  font-family: var(--font-ui); font-size: 11px; font-weight: 700;
}

/* Spacer rows */
.diff-row.row-spacer {
  min-height: 10px;
  background: var(--bg);
}
.diff-row.row-spacer .meta-col,
.diff-row.row-spacer .content-col { padding: 0; border: none; background: transparent; }

/* Equal rows — dimmed */
.diff-row.row-equal .content-col { color: #555250; }

/* ── Diff body container ── */
.diff-body {
  max-width: 1640px;
  margin: 16px auto 48px;
  border: 1px solid rgba(255,255,255,0.14);
  border-radius: 14px;
  box-shadow: 0 24px 60px rgba(0,0,0,0.45);
}
.diff-body > :first-child { border-top-left-radius: 13px; border-top-right-radius: 13px; }
.diff-body > :last-child  { border-bottom-left-radius: 13px; border-bottom-right-radius: 13px; overflow: hidden; }
.section-divider:first-child { border-top: none; }

/* ── No-diff ── */
.no-diff {
  text-align: center;
  padding: 80px;
  font-family: var(--font-ui);
  color: #c3cad8;
  background: rgba(255,255,255,0.05);
  backdrop-filter: blur(10px);
  border-radius: 13px;
}
"""

HTML_SHELL = """\
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DocDiff: {name_a} ↔ {name_b}</title>
<style>{css}</style>
</head>
<body>

<div class="top-header">
  <div class="brand"><div class="brand-icon">⇄</div>DocDiff</div>
  <div class="header-files">
    <div class="file-chip base"><b>BASE</b>{name_a}</div>
    <span style="color:#4a4640">vs</span>
    <div class="file-chip rev"><b>REVISED</b>{name_b}</div>
    {home_link}
  </div>
</div>

<div class="stats-bar">
  <div class="stat-item"><span class="stat-num chg">{changed}</span> modified</div>
  <div class="stat-item"><span class="stat-num del">{deleted}</span> deleted</div>
  <div class="stat-item"><span class="stat-num ins">{inserted}</span> inserted</div>
  <div class="stat-item"><span class="stat-num eq">{equal}</span> unchanged</div>
  <div class="stats-ts">{timestamp}</div>
</div>

<div class="col-strip">
  <div class="col-strip-meta col-strip-label"></div>
  <div class="col-strip-label base">Base document</div>
  <div class="col-strip-label rev">Revised document</div>
</div>

<div class="diff-body">
{body_html}
</div>

</body>
</html>
"""


def _page_label(page: int) -> str:
    return f"Page {page}" if page > 0 else "—"


def _section_label(rows_slice: list[DiffRow]) -> str:
    """Build 'Around Page X' or 'Around Pages X–Y' label from a group of rows."""
    pages_a = [r.page_a for r in rows_slice if r.page_a > 0]
    pages_b = [r.page_b for r in rows_slice if r.page_b > 0]
    all_pages = sorted(set(pages_a + pages_b))
    if not all_pages:
        return ""
    if len(all_pages) == 1:
        return f"Around Page {all_pages[0]}"
    return f"Around Pages {all_pages[0]}–{all_pages[-1]}"


def rows_to_html(rows: list[DiffRow], name_a: str, name_b: str) -> str:
    """
    Emit rows grouped into sections. Each section of consecutive non-equal
    rows (plus 2 context rows either side) gets a section divider header.
    """
    if not rows:
        return '<div class="no-diff"><p>✓ No differences found. Documents are identical in content.</p></div>'

    # Build groups: each change cluster + surrounding context
    CONTEXT = 2
    change_idx = {i for i, r in enumerate(rows) if r.status not in ("equal","spacer")}
    show_idx: set[int] = set()
    for ci in change_idx:
        for j in range(max(0, ci - CONTEXT), min(len(rows), ci + CONTEXT + 1)):
            show_idx.add(j)

    parts: list[str] = []
    in_section = False
    section_rows: list[DiffRow] = []

    def flush_section_divider(slice_rows: list[DiffRow]):
        label = _section_label(slice_rows)
        if label:
            parts.append(f'<div class="section-divider">{_esc(label)}</div>')

    prev_shown = False

    for i, row in enumerate(rows):
        shown = i in show_idx

        if shown and not prev_shown:
            # Start of a new visible group → section divider
            # Collect the upcoming change cluster for page label
            cluster = [rows[j] for j in range(i, min(len(rows), i + 20))
                       if j in show_idx]
            flush_section_divider(cluster)
            in_section = True

        if not shown:
            if prev_shown and in_section:
                # Gap between sections — emit a small gap marker
                parts.append('<div style="height:6px;background:var(--bg);border-bottom:1px solid var(--border)"></div>')
                in_section = False
            prev_shown = False
            continue

        prev_shown = True

        status = row.status
        css_cls = f"diff-row row-{status}"

        # Meta column content
        if status == "spacer":
            parts.append(f'<div class="{css_cls}"><div class="meta-col"></div>'
                         f'<div class="content-col left"></div>'
                         f'<div class="content-col right"></div></div>')
            continue

        # Page info
        pa = _page_label(row.page_a)
        pb = _page_label(row.page_b)
        if row.page_a == row.page_b and row.page_a > 0:
            page_html = f'<div class="meta-pages">Doc 1: {pa}<br>Doc 2: {pb}</div>'
        elif row.page_a > 0 and row.page_b > 0:
            page_html = f'<div class="meta-pages">Doc 1: {pa}<br>Doc 2: {pb}</div>'
        elif row.page_a > 0:
            page_html = f'<div class="meta-pages">Doc 1: {pa}<br><span>Doc 2: —</span></div>'
        else:
            page_html = f'<div class="meta-pages"><span>Doc 1: —</span><br>Doc 2: {pb}</div>'

        # Badge
        if status == "deleted":
            badge = '<span class="badge badge-deleted">Deleted</span>'
        elif status == "inserted":
            badge = '<span class="badge badge-inserted">Inserted</span>'
        elif status == "changed":
            badge = '<span class="badge badge-modified">Modified</span>'
        else:
            badge = '<span class="badge badge-equal">Unchanged</span>'

        # Confidence (only for changed)
        conf_html = ""
        if status == "changed":
            conf_html = f'<div class="confidence">Match confidence: <span>{row.confidence}%</span></div>'

        meta = f'<div class="meta-col">{page_html}{badge}{conf_html}</div>'

        # Content
        left_content  = row.left_html  or ""
        right_content = row.right_html or ""

        parts.append(
            f'<div class="{css_cls}">'
            f'{meta}'
            f'<div class="content-col left">{left_content}</div>'
            f'<div class="content-col right">{right_content}</div>'
            f'</div>'
        )

    return "\n".join(parts)


def build_html(rows: list[DiffRow], summary: dict, file_a: str, file_b: str,
               home_link: bool = False) -> str:
    name_a = html_mod.escape(Path(file_a).name)
    name_b = html_mod.escape(Path(file_b).name)
    body_html = rows_to_html(rows, name_a, name_b)
    link = '<a class="home-link" href="/">+ New comparison</a>' if home_link else ""
    return HTML_SHELL.format(
        css=CSS,
        name_a=name_a,
        name_b=name_b,
        home_link=link,
        timestamp=datetime.now().strftime("%d %b %Y, %H:%M"),
        body_html=body_html,
        **summary,
    )


def main():
    parser = argparse.ArgumentParser(
        description="Side-by-side document diff with word-level highlighting.")
    parser.add_argument("file_a", help="Base document (.txt / .docx / .pdf)")
    parser.add_argument("file_b", help="New document (.txt / .docx / .pdf)")
    parser.add_argument("--output", "-o", default="",
                        help="Output HTML path (default: auto-named)")
    parser.add_argument("--no-open", action="store_true",
                        help="Do not open the report in a browser")
    args = parser.parse_args()

    for f in (args.file_a, args.file_b):
        if not Path(f).exists():
            sys.exit(f"ERROR: File not found: {f}")

    if not args.output:
        out_path = f"diff_{Path(args.file_a).stem}_vs_{Path(args.file_b).stem}.html"
    else:
        out_path = args.output

    print(f"\n  Extracting: {args.file_a}")
    blocks_a = extract_blocks(args.file_a)
    print(f"  Extracting: {args.file_b}")
    blocks_b = extract_blocks(args.file_b)

    print(f"  Aligning {len(blocks_a)} <-> {len(blocks_b)} blocks ...")
    rows    = align_blocks(blocks_a, blocks_b)
    summary = summarise(rows)

    print(f"  Building report ...")
    html_out = build_html(rows, summary, args.file_a, args.file_b)

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html_out)

    abs_path = os.path.abspath(out_path)
    print(f"  OK {abs_path}")
    print(f"  Modified {summary['changed']} | Deleted {summary['deleted']} | "
          f"Inserted {summary['inserted']} | Unchanged {summary['equal']}")

    if not args.no_open:
        import webbrowser
        webbrowser.open(f"file:///{abs_path.replace(chr(92), '/')}")


if __name__ == "__main__":
    main()
