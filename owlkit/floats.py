"""Put figures and tables where the LaTeX PDF put them.

LaTeX *floats*: `\\begin{figure}` says "this figure belongs about here", and
the output routine then moves it to the top or bottom of whatever page it fits
on -- often a different page than the source position. Word has no float
algorithm at all; pandoc drops each figure inline exactly where the source
declared it.

So the same document has its figures in genuinely different places in the two
outputs. That is why naive page matching fails: a Word page inherits a figure
its LaTeX counterpart never had, overflows, and every later boundary
cascades.

This module reads each float's real placement out of the compiled PDF (which
page, top or bottom), detaches the corresponding block from the Word document,
and re-attaches it at the matching spot.
"""

import re
import subprocess
import xml.etree.ElementTree as ET

from docx.oxml.ns import qn


CAPTION_RE = re.compile(r"^\s*(Figure|Table|Scheme|Chart)\s+([A-Za-z]?\d+)\s*[:.]")


# ---------------------------------------------------------------------------
# where the PDF put them
# ---------------------------------------------------------------------------

def _pages_with_words(pdf_path):
    """[(page_height, [(y, x, text), ...]), ...] in reading order."""
    out = subprocess.run(["pdftotext", "-q", "-bbox", pdf_path, "-"],
                         capture_output=True, text=True)
    if out.returncode != 0:
        raise RuntimeError("pdftotext -bbox failed")
    root = ET.fromstring(out.stdout)
    pages = []
    for page in root.iter():
        if not page.tag.endswith("page"):
            continue
        height = float(page.get("height", 792))
        words = [(float(w.get("yMin")), float(w.get("xMin")), (w.text or ""))
                 for w in page.iter() if w.tag.endswith("word")]
        words.sort(key=lambda t: (round(t[0], 1), t[1]))
        pages.append((height, words))
    return pages


def _pages_with_lines(pdf_path):
    """[(page_height, [(y, text), ...]), ...] using poppler's word boxes."""
    out = subprocess.run(["pdftotext", "-q", "-bbox", pdf_path, "-"],
                         capture_output=True, text=True)
    if out.returncode != 0:
        raise RuntimeError("pdftotext -bbox failed")
    root = ET.fromstring(out.stdout)
    ns = {"x": "http://www.w3.org/1999/xhtml"}
    pages = []
    for page in root.iter():
        if not page.tag.endswith("page"):
            continue
        height = float(page.get("height", 792))
        words = []
        for w in page.iter():
            if not w.tag.endswith("word"):
                continue
            words.append((float(w.get("yMin")), float(w.get("xMin")),
                          (w.text or "")))
        words.sort(key=lambda t: (round(t[0], 1), t[1]))
        lines, cur_y, cur = [], None, []
        for y, x, text in words:
            if cur_y is None or abs(y - cur_y) <= 2.0:
                cur.append(text)
                cur_y = y if cur_y is None else cur_y
            else:
                lines.append((cur_y, " ".join(cur)))
                cur, cur_y = [text], y
        if cur:
            lines.append((cur_y, " ".join(cur)))
        pages.append((height, lines))
    return pages


def placements(pdf_path):
    """{(kind, number): (page_index, 'top'|'bottom')} for every float caption.

    Placement is decided by the caption's actual y position on the page, not by
    its index in the text stream -- figure-internal labels inflate line counts
    and make index-based guessing unreliable.
    """
    found = {}
    for page_idx, (height, words) in enumerate(_pages_with_words(pdf_path)):
        for i, (y, x, text) in enumerate(words):
            if text not in ("Figure", "Table", "Scheme", "Chart"):
                continue
            if i + 1 >= len(words):
                continue
            nxt = words[i + 1][2]
            m = re.match(r"^([A-Za-z]?\d+)[:.]?$", nxt)
            if not m:
                continue
            # a figure caption sits *below* its graphic, a table caption above,
            # so the caption's own y approximates where the float body lives
            where = "top" if y < height * 0.45 else "bottom"
            found.setdefault((text, m.group(1)), (page_idx, where))
    return found


# ---------------------------------------------------------------------------
# finding the blocks in the Word document
# ---------------------------------------------------------------------------

def _para_text(p):
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def _has_drawing(el):
    return el.find(".//" + qn("w:drawing")) is not None or \
        el.find(".//" + qn("w:pict")) is not None


def find_blocks(doc):
    """Locate float blocks: the graphic paragraph(s) plus their caption.

    Returns [(kind, number, [elements])] in document order.
    """
    body = doc.element.body
    children = list(body)
    blocks = []
    used = set()

    for i, el in enumerate(children):
        if id(el) in used:
            continue
        if el.tag != qn("w:p"):
            continue
        text = _para_text(el)
        m = CAPTION_RE.match(text)
        if not m:
            continue
        kind, num = m.group(1), m.group(2)

        group = [el]
        if kind == "Table":
            # caption sits on top of the table
            j = i + 1
            while j < len(children) and children[j].tag not in (qn("w:tbl"),):
                if children[j].tag == qn("w:p") and _para_text(children[j]).strip():
                    break
                j += 1
            if j < len(children) and children[j].tag == qn("w:tbl"):
                group.append(children[j])
        else:
            # the graphic paragraph(s) sit immediately above the caption
            j = i - 1
            while j >= 0 and children[j].tag == qn("w:p"):
                if _has_drawing(children[j]):
                    group.insert(0, children[j])
                    j -= 1
                elif not _para_text(children[j]).strip():
                    group.insert(0, children[j])
                    j -= 1
                else:
                    break
            # trim leading empties that carry no image
            while group and group[0].tag == qn("w:p") and \
                    not _has_drawing(group[0]) and not _para_text(group[0]).strip():
                group.pop(0)

        if len(group) > 1 or _has_drawing(group[0]):
            for g in group:
                used.add(id(g))
            blocks.append((kind, num, group))
    return blocks


def detach(blocks):
    """Remove every block from the document, keeping the elements alive."""
    for _, _, group in blocks:
        for el in group:
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)


# ---------------------------------------------------------------------------
# what each caption actually says, straight from the source
# ---------------------------------------------------------------------------

def source_captions(tex_src):
    """[(kind, caption_words), ...] for every float, in document order.

    Guessing where a caption ends from PDF geometry alone is unreliable -- the
    gap below a caption at the top of a page is an ordinary paragraph skip.
    The source already knows the exact wording, so the caption block can be cut
    out by matching it word for word instead.
    """
    from .counters import strip_comments
    src = strip_comments(tex_src)
    body = src.split("\\begin{document}", 1)[-1]
    out = []
    for m in re.finditer(r"\\begin\{(figure|table)\}\*?", body):
        kind = "Figure" if m.group(1) == "figure" else "Table"
        end = body.find("\\end{" + m.group(1) + "}", m.end())
        chunk = body[m.end():end if end != -1 else len(body)]
        cm = re.search(r"\\caption\s*(?:\[[^\]]*\])?\s*\{", chunk)
        if not cm:
            out.append((kind, []))
            continue

        depth, j = 1, cm.end()
        while j < len(chunk) and depth:
            if chunk[j] == "\\":
                j += 2
                continue
            if chunk[j] == "{":
                depth += 1
            elif chunk[j] == "}":
                depth -= 1
            j += 1
        # a table's *body* is prose-font text in the PDF but a detached float
        # block in the Word file, so it has to come out of the comparison too
        extra = []
        if kind == "Table":
            tb = re.search(r"\\begin\{tabular[xy*]?\}", chunk)
            if tb:
                te = chunk.find("\\end{tabular", tb.end())
                cells = chunk[tb.end():te if te != -1 else len(chunk)]
                cells = re.sub(r"\\\\|&", " ", cells)
                cells = re.sub(r"\\(?:textcolor|colorbox)\s*\{[^}]*\}", " ", cells)
                cells = re.sub(r"\\[a-zA-Z]+\s*(?:\[[^\]]*\])?", " ", cells)
                cells = re.sub(r"[{}$~]", " ", cells)
                extra = re.findall(r"[0-9A-Za-z][0-9A-Za-z'.-]*", cells)

        text = chunk[cm.end():j - 1]
        # drop the *first* argument of two-argument markup so its value does
        # not leak into the caption words (\textcolor{red}{...} -> ...)
        text = re.sub(r"\\(?:textcolor|colorbox)\s*\{[^}]*\}", " ", text)
        text = re.sub(r"\\label\s*\{[^}]*\}", " ", text)
        text = re.sub(r"\\[a-zA-Z]+\s*(?:\[[^\]]*\])?", " ", text)
        text = re.sub(r"[{}$~\\\\]", " ", text)
        words = re.findall(r"[0-9A-Za-z][0-9A-Za-z'-]*", text)
        out.append((kind, words + extra))
    return out
