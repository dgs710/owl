"""Extract the *prose* of a compiled PDF, page by page.

Naive `pdftotext` output cannot be used to find page boundaries, because it
also returns every label drawn *inside* a figure -- axis ticks, legends,
annotations -- interleaved with the body text.  Those words do not exist in
the Word document (the graphic is a single picture there), so they inflate the
per-page word counts and every computed boundary lands in the wrong place.

Two things separate prose from everything else, and both are in the PDF:

* **font** — the body is set in the document's dominant serif at its dominant
  size; figure internals are whatever the plotting tool used, almost always a
  different family at a much smaller size;
* **position** — captions are known from the float scan, so the caption block
  can be cut out by geometry.

What comes back is the same word stream the Word document has, in the same
order, which is what makes the alignment trustworthy.
"""

import re
from collections import Counter

import pymupdf

CAPTION_RE = re.compile(r"^(Figure|Table|Scheme|Chart)$")


def _family(font_name):
    """'NimbusRomNo9L-Regu' -> 'nimbusromno9l'; ignore the weight suffix."""
    return re.split(r"[-,]", font_name or "")[0].lower()


def body_style(doc):
    """The (size, family) that most of the document is set in."""
    tally = Counter()
    for page in doc:
        for block in page.get_text("dict")["blocks"]:
            for line in block.get("lines", []):
                for span in line["spans"]:
                    n = len(span["text"].split())
                    if n:
                        tally[(round(span["size"], 1),
                               _family(span["font"]))] += n
    return tally.most_common(1)[0][0] if tally else (12.0, "")


def page_words(page, size, family, tol=0.8):
    """Body-font words on one page, in reading order, with their geometry."""
    words = []
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            for span in line["spans"]:
                if abs(span["size"] - size) > tol:
                    continue
                if family and _family(span["font"]) != family:
                    continue
                y = round(span["bbox"][1], 1)
                for m in re.finditer(r"\S+", span["text"]):
                    words.append({
                        "text": m.group(0),
                        "y": y,
                        "x": span["bbox"][0],
                    })
    words.sort(key=lambda w: (w["y"], w["x"]))
    return words


def _leading(words):
    """The dominant baseline pitch, i.e. the body line spacing."""
    ys = sorted({w["y"] for w in words})
    gaps = Counter(round(b - a, 1) for a, b in zip(ys, ys[1:]) if 0 < b - a < 40)
    return gaps.most_common(1)[0][0] if gaps else 14.5


def _caption_span(words, start):
    """Where the caption beginning at index `start` ends.

    A caption runs on until the vertical gap to the next line exceeds the
    normal line pitch -- that gap is the float's separation from the body.
    """
    lead = _leading(words)
    i = start
    while i + 1 < len(words):
        gap = words[i + 1]["y"] - words[i]["y"]
        if gap > lead * 1.6:
            return i + 1
        i += 1
    return len(words)


def _norm(word):
    return re.sub(r"[^0-9a-z]", "", word.lower())


def _strip_captions(words, captions_by_number=None):
    """Remove caption blocks.

    Where the caption text is known from the LaTeX source, the block is cut by
    matching those words -- exact, and immune to the fact that the gap under a
    caption at the top of a page is just an ordinary paragraph skip.  Only when
    the text is unknown does this fall back to the vertical-gap heuristic.
    """
    captions_by_number = captions_by_number or {}
    starts = []
    for i, w in enumerate(words):
        if not CAPTION_RE.match(w["text"].strip()):
            continue
        nxt = words[i + 1]["text"] if i + 1 < len(words) else ""
        m = re.match(r"^([A-Za-z]?\d+)[:.]?$", nxt.strip())
        if m:
            starts.append((i, w["text"].strip(), m.group(1)))
    if not starts:
        return words

    drop = set()
    for s, kind, num in starts:
        expected = captions_by_number.get((kind, num))
        if expected:
            # "Figure" + "N:" then the caption words themselves
            j, k = s + 2, 0
            while j < len(words) and k < len(expected):
                if _norm(words[j]["text"]) == _norm(expected[k]):
                    k += 1
                    j += 1
                elif not _norm(words[j]["text"]):
                    j += 1
                else:
                    # a stray token (a maths fragment, a stripped macro):
                    # skip it, but do not skip past the end of the caption
                    j += 1
                    if j - s > len(expected) * 2 + 12:
                        break
            end = j
        else:
            end = _caption_span(words, s)
        drop.update(range(s, min(end, len(words))))
    return [w for i, w in enumerate(words) if i not in drop]


def prose_pages(pdf_path, placements=None, captions=None):
    """[[word, ...], ...] -- the body prose of each page, captions and figure
    internals removed.

    `placements` is the mapping from `floats.placements()`; it tells us which
    pages carry a float and whether it sits at the top or the bottom.
    """
    doc = pymupdf.open(pdf_path)
    size, family = body_style(doc)
    pages = []
    for page in doc:
        words = page_words(page, size, family)
        words = _strip_captions(words, captions)
        pages.append([w["text"] for w in words])
    doc.close()
    return pages


def strip_running_heads(pages):
    """Drop the repeating header / footer words from each page."""
    if len(pages) < 3:
        return pages

    def shape(seq):
        return " ".join("#" if w.isdigit() else w for w in seq)

    firsts = Counter(shape(p[:6]) for p in pages if p)
    lasts = Counter(shape(p[-6:]) for p in pages if p)
    threshold = max(3, len(pages) // 4)

    # the header is the longest repeating prefix; find its length by testing
    # progressively shorter prefixes
    def repeating_len(counter_source, take_front):
        for n in range(6, 0, -1):
            c = Counter(shape(p[:n] if take_front else p[-n:])
                        for p in pages if len(p) >= n)
            if any(v >= threshold for v in c.values()):
                common = {k for k, v in c.items() if v >= threshold}
                return n, common
        return 0, set()

    nf, common_f = repeating_len(firsts, True)
    nl, common_l = repeating_len(lasts, False)

    out = []
    for p in pages:
        q = list(p)
        if nf and len(q) >= nf and shape(q[:nf]) in common_f:
            q = q[nf:]
        if nl and len(q) >= nl and shape(q[-nl:]) in common_l:
            q = q[:-nl]
        out.append(q)
    return out
