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


def graphic_regions(page, min_height=6.0, min_width=40.0):
    """Rectangles on the page occupied by artwork.

    Filtering text by font size alone keeps headings (good) but also keeps the
    large annotations people put inside figures -- "waiting for Brian's raw
    data" is set at body size and is not part of the prose at all.  Those words
    live inside the graphic, so the reliable way to drop them is geometric:
    find where the artwork is and ignore any text inside it.
    """
    rects = []
    for d in page.get_drawings():
        r = d["rect"]
        w, h = r[2] - r[0], r[3] - r[1]
        if h < min_height or w < min_width:
            continue                      # rules, underlines, tick marks
        if h > page.rect.height * 0.95:
            continue                      # a full-page background
        rects.append([r[0], r[1], r[2], r[3]])
    try:
        for info in page.get_images(full=True):
            for r in page.get_image_rects(info[0]):
                rects.append([r[0], r[1], r[2], r[3]])
    except Exception:
        pass
    if not rects:
        return []

    # merge overlapping rectangles into whole figures
    merged = []
    for r in sorted(rects, key=lambda t: t[1]):
        for m in merged:
            if not (r[3] < m[1] - 2 or r[1] > m[3] + 2):
                m[0], m[1] = min(m[0], r[0]), min(m[1], r[1])
                m[2], m[3] = max(m[2], r[2]), max(m[3], r[3])
                break
        else:
            merged.append(list(r))
    # only regions big enough to be a figure
    return [m for m in merged
            if (m[3] - m[1]) > 30 and (m[2] - m[0]) > 80]


def page_words(page, size, family, tol=0.8, band=3.0, min_ratio=0.95):
    """Body-font words on one page, in true reading order.

    Words are grouped into lines by a vertical band rather than by exact top
    coordinate, then ordered left to right within each line.  This matters
    because a hyperlinked cross-reference renders with a marginally different
    bounding box than the text around it: sorting on the raw coordinate pulls
    every link out of its sentence and clusters them elsewhere on the page,
    which reads as though the Word file had put the cross-references in the
    wrong place when they are exactly where they belong.
    """
    zones = graphic_regions(page)

    def inside_graphic(x0, y0, x1, y1):
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        for gx0, gy0, gx1, gy1 in zones:
            if gx0 - 2 <= cx <= gx1 + 2 and gy0 - 2 <= cy <= gy1 + 2:
                return True
        return False

    raw = []
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            for span in line["spans"]:
                # Filter on SIZE ALONE, keeping body text and anything larger.
                #
                # Matching the font family as well looks safer but is not: the
                # regular and bold cuts of one typeface report different base
                # names in a Word export ("TimesNewRomanPSMT" vs
                # "TimesNewRomanPS-BoldMT") and identical ones in a LibreOffice
                # export, so a family test silently drops every heading from
                # one side only.  A heading sitting on a page boundary then
                # lands on whichever page the aligner guesses.
                #
                # Size alone is symmetric: headings are larger than the body,
                # and the labels drawn inside figures are far smaller.
                if span["size"] < size * min_ratio:
                    continue
                x0, y0, x1, y1s = span["bbox"]
                if inside_graphic(x0, y0, x1, y1s):
                    continue
                text = span["text"]
                n = max(1, len(text))
                for m in re.finditer(r"\S+", text):
                    # spread the span's width across its characters so words
                    # inside one span still order correctly
                    frac = m.start() / n
                    raw.append((y0, x0 + (x1 - x0) * frac, m.group(0)))
    if not raw:
        return []

    raw.sort(key=lambda t: t[0])
    lines, cur, cur_y = [], [], raw[0][0]
    for y, x, text in raw:
        if abs(y - cur_y) <= band:
            cur.append((y, x, text))
        else:
            lines.append((cur_y, cur))
            cur, cur_y = [(y, x, text)], y
    lines.append((cur_y, cur))

    words = []
    for line_y, items in lines:
        for y, x, text in sorted(items, key=lambda t: t[1]):
            words.append({"text": text, "y": line_y, "x": x})
    return _join_line_hyphens(words)


def _join_line_hyphens(words):
    """Rejoin words LaTeX hyphenated at a line break.

    LaTeX hyphenates freely, so the PDF's text layer contains "build-" at the
    end of one line and "ing" at the start of the next.  Word never breaks
    those words, so left alone they read as content the Word file is missing --
    they accounted for most of the apparent 10% shortfall.  Only a hyphen at
    the very end of a line is a break; a genuine compound such as
    "sequence-independent" arrives as a single token and is left alone.
    """
    out = []
    i = 0
    while i < len(words):
        w = words[i]
        text = w["text"]
        if (text.endswith("-") and len(text) > 2 and i + 1 < len(words)
                and words[i + 1]["y"] != w["y"]):
            nxt = words[i + 1]["text"]
            if nxt[:1].islower():
                out.append({**w, "text": text[:-1] + nxt})
                i += 2
                continue
        out.append(w)
        i += 1
    return out


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


def _find_run(words, expected, lookahead=6):
    """Best starting index in `words` for the sequence `expected`.

    Returns None when nothing matches well enough.
    """
    if not expected:
        return None
    probe = [_norm(x) for x in expected[:lookahead] if _norm(x)]
    if not probe:
        return None
    keys = [_norm(w["text"]) for w in words]
    best, best_score = None, 0
    for i in range(len(keys)):
        score = 0
        k = i
        for want in probe:
            # allow a couple of stray tokens between matches
            for skip in range(3):
                if k + skip < len(keys) and keys[k + skip] == want:
                    score += 1
                    k = k + skip + 1
                    break
            else:
                break
        if score > best_score:
            best, best_score = i, score
    # a one-word caption (a placeholder such as "Caption.") can only ever
    # score 1, so do not demand two matches when there is only one to make
    need = 1 if len(probe) == 1 else max(2, len(probe) // 2)
    if best_score < need:
        return None
    return best


def _label_start(words, kind, num):
    """Fallback: find the "Figure 3:" label itself.

    The label can be split across font spans -- a caption whose label is set in
    a different weight can arrive as a lone "T" followed by "able" -- so match
    on a prefix and allow the number to be one or two tokens away.
    """
    prefix = kind[:3].lower()
    want = _norm(num)
    for i, w in enumerate(words):
        t = _norm(w["text"])
        if not t or not (t.startswith(prefix) or prefix.startswith(t)):
            continue
        for k in range(1, 4):
            if i + k < len(words) and _norm(words[i + k]["text"]) == want:
                return i
    return None


def _strip_float_text(words, expected, kind=None, num=None, body=None):
    """Remove one float's caption (and, for a table, its cells) from a page.

    Locating the block by its *text* rather than by a "Figure N:" label makes
    this immune to the label being split across font spans -- a real case in
    the reference document, where the caption's "Table" arrives as a lone "T"
    because the label is set in a different weight, so a token-equality test
    never fires and the whole caption and table leak into the prose.
    """
    start = _find_run(words, expected)
    if start is None and kind and num:
        start = _label_start(words, kind, num)
        if start is not None:
            end = _caption_span(words, start)
            return [w for i, w in enumerate(words) if not (start <= i < end)]
    if start is None:
        return words
    from collections import Counter
    budget = Counter(_norm(x) for x in expected if _norm(x))
    # Hard ceiling: a float's text cannot be longer than what the source says
    # it is, plus a little slack.  Without it the matcher keeps going on
    # ambiguous tokens -- a table full of numbers will happily "match" the
    # "2.2" of the next heading and swallow it.
    limit = min(len(words), start + len(expected) + 6)
    j, misses, end = start, 0, start
    while j < limit and misses <= 3:
        key = _norm(words[j]["text"])
        if not key:
            j += 1
            continue
        if budget.get(key):
            budget[key] -= 1
            misses = 0
            j += 1
            end = j
        else:
            misses += 1
            j += 1
    # take the label tokens sitting immediately in front of the caption too
    lo = start
    while lo > 0 and start - lo < 3 and len(words[lo - 1]["text"]) <= 6:
        lo -= 1

    # a table's cells sit on the other side of its caption; match them as
    # their own run so they cannot spill into the surrounding prose
    if body:
        bud = Counter(_norm(x) for x in body if _norm(x))
        k, misses2 = lo - 1, 0
        floor = max(0, lo - len(body) - 6)
        while k >= floor and misses2 <= 3:
            key = _norm(words[k]["text"])
            if not key:
                k -= 1
                continue
            if bud.get(key):
                bud[key] -= 1
                misses2 = 0
                lo = k
            else:
                misses2 += 1
            k -= 1

    return [w for i, w in enumerate(words) if not (lo <= i < end)]


_PAGENO = re.compile(r"^[ivxlcdmIVXLCDM]+$|^\d{1,4}$")


def prose_pages(pdf_path, placements=None, captions=None):
    """[[word, ...], ...] -- the body prose of each page.

    Figure internals go by font; each float's caption (and a table's cells) go
    by matching the text the LaTeX source says they contain, on the page the
    PDF says the float is on.
    """
    doc = pymupdf.open(pdf_path)
    size, family = body_style(doc)
    captions = captions or {}

    on_page = {}
    for key, (page_idx, _where) in (placements or {}).items():
        on_page.setdefault(page_idx, []).append(key)

    pages = []
    for idx, page in enumerate(doc):
        words = page_words(page, size, family)
        for key in on_page.get(idx, []):
            entry = captions.get(key)
            if isinstance(entry, tuple):
                expected, body = entry
            else:
                expected, body = (entry or []), None
            words = _strip_float_text(words, expected or [], key[0], key[1],
                                      body)
        pages.append([w["text"] for w in words])
    doc.close()
    return _dehyphenate(pages)


def _dehyphenate(pages):
    """Rejoin a word LaTeX split across a page break.

    LaTeX hyphenates, so a page can end "...the measure-" and the next begin
    "ment that...".  Word never splits a word across pages, so comparing the
    two as-is reports a mismatch on a boundary that is in fact exactly right.
    The fragment is folded back into the page where the word starts.
    """
    for i in range(len(pages) - 1):
        if not pages[i] or not pages[i + 1]:
            continue
        last = pages[i][-1]
        if not last.endswith("-") or len(last) < 3:
            continue
        nxt = pages[i + 1][0]
        if not nxt or not nxt[:1].islower():
            continue
        pages[i][-1] = last[:-1] + nxt
        pages[i + 1] = pages[i + 1][1:]
    return pages


def strip_running_heads(pages):
    """Drop the repeating header / footer words from each page."""
    if len(pages) < 3:
        return pages

    roman = re.compile(r"^[ivxlcdmIVXLCDM]+$")

    def shape(seq):
        # the page number in a running head may be arabic OR roman -- an
        # appendix commonly renumbers to i, ii, iii... and leaving those as
        # literal words means no two headers share a shape, so none of them
        # ever reaches the repetition threshold and the numbers leak into the
        # body word stream as if they were prose.
        out = []
        for w in seq:
            if w.isdigit() or roman.match(w):
                out.append("#")
            else:
                out.append(w)
        return " ".join(out)

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

    # A page number can survive on its own when the rest of the running head
    # is set smaller and has already been filtered out by size -- and a bare
    # "5" at the head of a page reads as the first word of the page.
    lead_no = sum(1 for p in pages if p and _PAGENO.match(p[0]))
    tail_no = sum(1 for p in pages if p and _PAGENO.match(p[-1]))

    out = []
    for p in pages:
        q = list(p)
        if nf and len(q) >= nf and shape(q[:nf]) in common_f:
            q = q[nf:]
        if nl and len(q) >= nl and shape(q[-nl:]) in common_l:
            q = q[:-nl]
        if lead_no >= threshold and q and _PAGENO.match(q[0]):
            q = q[1:]
        if tail_no >= threshold and q and _PAGENO.match(q[-1]):
            q = q[:-1]
        out.append(q)
    return out
