"""Make the Word file break pages exactly where the LaTeX PDF breaks them.

Word will never *naturally* break where LaTeX does: LaTeX uses Knuth-Plass
total-fit line breaking, Word uses greedy first-fit, the Times metrics differ,
and Word has no \\microtype.  Chasing natural agreement is hopeless.

So we do not chase it.  The compiled PDF is taken as ground truth, its page
boundaries are read off, and *hard page breaks* are inserted into the .docx at
exactly those word positions.  The boundaries then hold by construction, in
any Word version.

The one failure mode left is overflow: if Word's text runs longer than
LaTeX's, a page's content spills past the frame before it reaches the hard
break, and Word emits a stray page.  So the tuning only ever runs one way --
make the Word text slightly *denser* than the LaTeX text, never looser -- by
widening the text block (smaller margins) and, if needed, nudging line spacing
and figure scale.  Density is uniform across the document.

Convergence test: after inserting N-1 hard breaks, a render must produce
exactly N pages.  More than N means something overflowed.
"""

import copy
import difflib
import os
import re
import shutil
import subprocess
import tempfile

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, Twips

from . import floats, pdfprose


# ---------------------------------------------------------------------------
# reading the ground truth
# ---------------------------------------------------------------------------

_WORD = re.compile(r"[0-9A-Za-zÀ-ɏ]+")


def normalise(word):
    """Compare words the way a human would: case- and accent-insensitively,
    ignoring punctuation and the ligatures pdftotext emits."""
    w = (word.lower()
         .replace("ﬀ", "ff").replace("ﬁ", "fi").replace("ﬂ", "fl")
         .replace("ﬃ", "ffi").replace("ﬄ", "ffl")
         .replace("’", "'").replace("–", "-").replace("—", "-"))
    m = _WORD.findall(w)
    return "".join(m)


def pdf_pages_text(pdf_path):
    """Per-page text of a PDF, using poppler."""
    out = subprocess.run(["pdftotext", "-q", pdf_path, "-"],
                         capture_output=True, text=True)
    if out.returncode != 0:
        raise RuntimeError(f"pdftotext failed on {pdf_path}")
    return out.stdout.split("\f")[:-1] or [out.stdout]


def strip_running_heads(pages):
    """Drop the repeating running header / footer lines.

    They carry the page number and the short title, appear on nearly every
    page, and would otherwise pollute the word stream that we align against
    the Word document (whose header lives outside the body text).
    """
    if len(pages) < 3:
        return pages

    def shape(line):
        # a header differs page to page only by its number; the page number can
        # sit on either side, so odd and even pages produce two variants
        return re.sub(r"\d+", "#", line.strip())

    first_shapes, last_shapes = {}, {}
    for p in pages:
        lines = [l for l in p.split("\n") if l.strip()]
        for l in lines[:2]:
            first_shapes[shape(l)] = first_shapes.get(shape(l), 0) + 1
        for l in lines[-2:]:
            last_shapes[shape(l)] = last_shapes.get(shape(l), 0) + 1

    # a quarter of the pages is enough: odd/even header variants each hit ~50%,
    # and a bare page number hits ~100%
    threshold = max(3, len(pages) // 4)
    common_first = {k for k, v in first_shapes.items() if v >= threshold and k}
    common_last = {k for k, v in last_shapes.items() if v >= threshold and k}

    cleaned = []
    for p in pages:
        lines = [l for l in p.split("\n") if l.strip()]
        while lines and shape(lines[0]) in common_first:
            lines = lines[1:]
        while lines and shape(lines[-1]) in common_last:
            lines = lines[:-1]
        cleaned.append("\n".join(lines))
    return cleaned


def page_word_counts(pdf_path, placements=None, captions=None):
    """(words, counts): the PDF's *prose* word stream and how many words fall
    on each page.

    Uses the font-filtered extractor, not raw pdftotext: labels drawn inside
    figures do not exist in the Word document, and counting them puts every
    boundary in the wrong place.
    """
    pages = pdfprose.strip_running_heads(
        pdfprose.prose_pages(pdf_path, placements, captions))
    words, counts = [], []
    for p in pages:
        ws = [w for w in (normalise(x) for x in p) if w]
        words.extend(ws)
        counts.append(len(ws))
    return words, counts


# ---------------------------------------------------------------------------
# the Word side: a word stream we can point back into the XML with
# ---------------------------------------------------------------------------

def _iter_body_paragraphs(doc):
    """Body paragraphs in document order, including those inside tables."""
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    def walk(parent, element):
        for child in element.iterchildren():
            if child.tag == qn("w:p"):
                yield Paragraph(child, parent)
            elif child.tag == qn("w:tbl"):
                table = Table(child, parent)
                for row in table.rows:
                    for cell in row.cells:
                        for el in cell._tc.iterchildren():
                            if el.tag == qn("w:p"):
                                yield Paragraph(el, cell)
    return list(walk(doc, doc.element.body))


def docx_word_stream(doc):
    """Every word in the document body, each tagged with where it lives.

    Returns (words, sites) where sites[i] = (paragraph, run, char_offset) for
    the i-th normalised word.
    """
    words, sites = [], []
    for para in _iter_body_paragraphs(doc):
        # Walk the paragraph in document order, including runs wrapped in a
        # <w:hyperlink> and the text inside <m:oMath>.
        #
        # python-docx's .runs returns only direct <w:r> children, which drops
        # every cross-reference; and maths lives in <m:t>, not <w:t>.  Both are
        # present in the PDF's text, so leaving either out of this stream makes
        # the two sides disagree on how many words a page holds and drags every
        # later boundary out of place.  Maths is included for alignment but
        # marked unsplittable -- a page must never break inside an equation.
        for node in para._p.iter():
            if node.tag == qn("w:r"):
                if _in_math(node):
                    continue
                text = "".join(t.text or "" for t in node.findall(qn("w:t")))
                if not text:
                    continue
                for m in re.finditer(r"\S+", text):
                    w = normalise(m.group(0))
                    if w:
                        words.append(w)
                        sites.append((para, node, m.start()))
            elif node.tag == qn("m:oMath"):
                text = "".join(t.text or "" for t in node.iter(qn("m:t")))
                for m in re.finditer(r"\S+", text):
                    w = normalise(m.group(0))
                    if w:
                        words.append(w)
                        sites.append((para, None, None))     # unsplittable
    return words, sites


def _in_math(el):
    parent = el.getparent()
    while parent is not None:
        if parent.tag in (qn("m:oMath"), qn("m:oMathPara")):
            return True
        parent = parent.getparent()
    return False


# ---------------------------------------------------------------------------
# alignment
# ---------------------------------------------------------------------------

def align_boundaries(pdf_words, page_counts, docx_words, anchor=10):
    """Map each PDF page boundary onto an index in the Word word stream.

    Each page is located by its own **opening words**, not by counting words
    from the start of the document.  Cumulative counting compounds every small
    disagreement between the two streams -- a symbol the PDF renders as text
    and Word keeps as an equation, a ligature, a hyphenated word -- so by the
    middle of a long document the boundary can be twenty words out even though
    every page individually is fine.  Searching for the page's first words
    makes each boundary independent of every other one.

    A global difflib alignment still runs first, to give each search a
    neighbourhood to look in and to catch pages whose opening words repeat
    elsewhere.
    """
    matcher = difflib.SequenceMatcher(None, pdf_words, docx_words, autojunk=False)
    blocks = [b for b in matcher.get_matching_blocks() if b.size]

    boundaries, cum = [], 0
    for count in page_counts[:-1]:          # no break after the final page
        cum += count
        boundaries.append(cum)

    def rough(b):
        """Where difflib thinks this position lands, interpolating gaps."""
        prev_end_a = prev_end_b = 0
        for a, bx, size in blocks:
            if a <= b <= a + size:
                return bx + (b - a)
            if a > b:                       # b sits in the gap before this block
                span_a = a - prev_end_a
                if span_a <= 0:
                    return prev_end_b
                frac = (b - prev_end_a) / span_a
                return int(prev_end_b + frac * (bx - prev_end_b))
            prev_end_a, prev_end_b = a + size, bx + size
        return prev_end_b

    mapped = []
    for b in boundaries:
        guess = rough(b)
        probe = [w for w in pdf_words[b:b + anchor] if w]
        best = guess
        if probe:
            window = 300
            lo = max(0, guess - window)
            hi = min(len(docx_words), guess + window)
            best_score, best_pos = 0, None
            for k in range(lo, hi):
                score = 0
                for t in range(len(probe)):
                    if k + t < len(docx_words) and docx_words[k + t] == probe[t]:
                        score += 1
                    else:
                        break
                if score > best_score:
                    best_score, best_pos = score, k
                    if score == len(probe):
                        break
            if best_pos is not None and best_score >= min(3, len(probe)):
                best = best_pos
        mapped.append(best)

    # boundaries must never run backwards
    for k in range(1, len(mapped)):
        if mapped[k] is not None and mapped[k - 1] is not None:
            mapped[k] = max(mapped[k], mapped[k - 1])
    return mapped


# ---------------------------------------------------------------------------
# inserting the breaks
# ---------------------------------------------------------------------------

def _split_run_at(run, offset):
    """Split a run so that everything from `offset` on lives in a new run that
    directly follows it.  Returns the new run element."""
    text = run.text
    head, tail = text[:offset], text[offset:]
    run.text = head
    new_r = copy.deepcopy(run._r)
    # clear the copied text nodes and set the tail
    for t in new_r.findall(qn("w:t")):
        new_r.remove(t)
    t = new_r.makeelement(qn("w:t"), {})
    t.text = tail
    t.set(qn("xml:space"), "preserve")
    new_r.append(t)
    run._r.addnext(new_r)
    return new_r


def _pPr(p_el):
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        pPr = p_el.makeelement(qn("w:pPr"), {})
        p_el.insert(0, pPr)
    return pPr


def _set_page_break_before(p_el, on=True):
    pPr = _pPr(p_el)
    tag = qn("w:pageBreakBefore")
    existing = pPr.find(tag)
    if on and existing is None:
        el = pPr.makeelement(tag, {})
        pPr.insert(0, el)
    elif not on and existing is not None:
        pPr.remove(existing)


def split_paragraph_at(para, run, offset):
    """Split `para` so everything from `offset` onward lives in a new
    paragraph immediately after it.  Returns the new paragraph element.

    The split walks the paragraph's *inline children* -- runs, hyperlinks,
    maths, bookmarks -- not just its direct <w:r> children.  Splitting on runs
    alone leaves every <w:hyperlink> in both halves, which duplicates each
    cross-reference and drops it at the wrong end of the page.

    The visual seam is hidden: the first half gets last-line justification so
    its final line still reaches the right margin, and the second half loses
    its first-line indent and space-before so it reads as a continuation.
    """
    p_el = para._p if hasattr(para, "_p") else para
    r_el = run if not hasattr(run, "_r") else run._r

    def inline_children(el):
        return [c for c in el if c.tag not in (qn("w:pPr"),)]

    # which inline child holds this run?
    children = inline_children(p_el)
    holder_idx = None
    for k, ch in enumerate(children):
        if ch is r_el or r_el in list(ch.iter()):
            holder_idx = k
            break
    if holder_idx is None:
        holder_idx = 0

    tail_el = copy.deepcopy(p_el)
    tail_children = inline_children(tail_el)

    # split the text of the run itself
    if offset:
        ts = r_el.findall(qn("w:t"))
        text = "".join(t.text or "" for t in ts)
        head_text, tail_text = text[:offset], text[offset:]
        for t in ts[1:]:
            r_el.remove(t)
        if ts:
            ts[0].text = head_text
            ts[0].set(qn("xml:space"), "preserve")
        # the same run inside the copied paragraph keeps the remainder
        twin = None
        for cand in tail_children[holder_idx].iter(qn("w:r")):
            twin = cand
            if cand.tag == qn("w:r"):
                # match by position among runs in that child
                pass
        holder = tail_children[holder_idx]
        runs_in_holder = list(holder.iter(qn("w:r")))
        orig_runs = list(children[holder_idx].iter(qn("w:r")))
        try:
            pos = orig_runs.index(r_el)
        except ValueError:
            pos = 0
        if pos < len(runs_in_holder):
            twin = runs_in_holder[pos]
            tts = twin.findall(qn("w:t"))
            for t in tts[1:]:
                twin.remove(t)
            if tts:
                tts[0].text = tail_text
                tts[0].set(qn("xml:space"), "preserve")
            # anything before this run inside the same child goes to the head
            for earlier in runs_in_holder[:pos]:
                parent = earlier.getparent()
                if parent is not None:
                    parent.remove(earlier)
        # the tail KEEPS the child that was split -- it holds the remainder of
        # the text -- while the head drops everything after it
        keep_from = holder_idx
        drop_head_from = holder_idx + 1
    else:
        keep_from = holder_idx
        drop_head_from = holder_idx

    for ch in children[drop_head_from:]:
        p_el.remove(ch)
    for ch in tail_children[:keep_from]:
        tail_el.remove(ch)

    # Justify the head's final line so the seam is invisible -- but only when
    # the head actually keeps some text.  Stretching an empty or one-word
    # remnant across the measure is worse than the seam it hides.
    head_text = "".join(t.text or "" for t in p_el.iter(qn("w:t")))
    if len(head_text.split()) >= 6:
        head_pPr = _pPr(p_el)
        jc = head_pPr.find(qn("w:jc"))
        if jc is None:
            jc = head_pPr.makeelement(qn("w:jc"), {})
            head_pPr.append(jc)
        if jc.get(qn("w:val")) in (None, "both", "justify"):
            jc.set(qn("w:val"), "distribute")

    tail_pPr = _pPr(tail_el)
    ind = tail_pPr.find(qn("w:ind"))
    if ind is not None:
        ind.set(qn("w:firstLine"), "0")
    spacing = tail_pPr.find(qn("w:spacing"))
    if spacing is not None:
        spacing.set(qn("w:before"), "0")

    p_el.addnext(tail_el)
    return tail_el


def strip_existing_page_breaks(doc):
    """Remove page breaks that came from \\clearpage etc. -- the ground-truth
    map already accounts for every break in the PDF."""
    removed = 0
    for br in list(doc.element.body.iter(qn("w:br"))):
        if br.get(qn("w:type")) == "page":
            parent = br.getparent()
            parent.remove(br)
            removed += 1
    for pPr in list(doc.element.body.iter(qn("w:pPr"))):
        for pbb in pPr.findall(qn("w:pageBreakBefore")):
            pPr.remove(pbb)
            removed += 1
    return removed


def section_break_paragraphs(doc):
    """Paragraphs carrying an inline <w:sectPr>.  A section break of type
    nextPage is itself a page break."""
    out = []
    for p in doc.element.body.findall(qn("w:p")):
        pPr = p.find(qn("w:pPr"))
        if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
            out.append(p)
    return out


def drop_redundant_breaks(doc):
    """Where a section break already starts a page, remove our own break."""
    dropped = 0
    body = doc.element.body
    children = list(body)
    for p in section_break_paragraphs(doc):
        try:
            idx = children.index(p)
        except ValueError:
            continue
        for nxt in children[idx + 1:]:
            if nxt.tag != qn("w:p"):
                break
            pPr = nxt.find(qn("w:pPr"))
            if pPr is not None and pPr.find(qn("w:pageBreakBefore")) is not None:
                pPr.remove(pPr.find(qn("w:pageBreakBefore")))
                dropped += 1
            break
    return dropped


def apply_boundaries(doc, sites, indices, prose_pages=None):
    """Turn each mapped word index into a real page start.

    `prose_pages` marks which pages actually carry body text.  A page holding
    nothing but a full-width figure has no prose at all, so two consecutive
    boundaries land on the same word; splitting there anyway manufactures an
    empty paragraph, and the figure then arrives on a page of its own -- two
    Word pages where LaTeX has one.  Those pages are skipped here and the
    float is made the page start instead.

    Returns {page_number: element_that_starts_it}.
    """
    starts = {}
    ordered = []
    for page_no, idx in enumerate(indices, start=2):
        if idx is None or idx >= len(sites):
            continue
        if prose_pages is not None and not prose_pages.get(page_no, True):
            continue
        # never break inside an equation: walk on to the next real run
        while idx < len(sites) and sites[idx][1] is None:
            idx += 1
        if idx >= len(sites):
            continue
        ordered.append((page_no, idx))

    # back to front: splitting a later paragraph cannot disturb earlier sites
    for page_no, idx in sorted(ordered, key=lambda t: t[1], reverse=True):
        para, run, offset = sites[idx]

        # A heading is atomic.  Splitting one leaves "2.2 Experiment" at the
        # foot of a page and "1: chaotropic softening..." at the head of the
        # next -- and the justification that hides an ordinary mid-sentence
        # split stretches the orphaned half across the full measure.  LaTeX
        # never breaks there either: it pushes the whole heading over.
        if _is_heading(para):
            first = _first_run(para)
            if first is not None:
                run, offset = first, 0

        tail = split_paragraph_at(para, run, offset)
        _set_page_break_before(tail, True)
        starts[page_no] = tail
    return starts


def _is_heading(para):
    try:
        name = (para.style.name if para.style else "") or ""
    except Exception:
        return False
    name = name.strip().lower()
    return name.startswith("heading") or name in ("title", "subtitle")


def _first_run(para):
    p_el = para._p if hasattr(para, "_p") else para
    for r in p_el.iter(qn("w:r")):
        if "".join(t.text or "" for t in r.findall(qn("w:t"))):
            return r
    return None


def build_candidate(docx_path, pdf_words, counts, placements, density):
    """One attempt: apply a density, re-place the floats where the PDF has
    them, and split the prose at the ground-truth page boundaries."""
    doc = Document(docx_path)
    strip_existing_page_breaks(doc)
    apply_density(doc, density)

    # take the floats out first -- they must not join the prose alignment, and
    # they are re-attached at page granularity anyway
    blocks = floats.find_blocks(doc)
    floats.detach(blocks)

    words, sites = docx_word_stream(doc)
    mapped = align_boundaries(pdf_words, counts, words)

    # which pages carry prose at all (page 1 always does -- it is the cover)
    prose_pages = {n: counts[n - 1] > 0 for n in range(1, len(counts) + 1)}
    starts = apply_boundaries(doc, sites, mapped, prose_pages)

    body = doc.element.body

    def anchor_after(page_no):
        """The element that begins the first prose page at or after page_no."""
        for n in range(page_no, len(counts) + 2):
            if n in starts:
                return starts[n]
        return None

    # place the floats last-page-first so earlier insertions stay valid
    by_page = []
    for kind, num, group in blocks:
        target = placements.get((kind, num))
        if target is None:
            by_page.append((10 ** 6, kind, num, group, None))
        else:
            by_page.append((target[0] + 1, kind, num, group, target[1]))

    placed, orphaned = 0, []
    for page_no, kind, num, group, where in sorted(by_page, reverse=True):
        if where is None:
            orphaned.append(f"{kind} {num}")
            continue
        # a float at the top of page P goes before P's own start; one at the
        # bottom goes before the start of P+1
        anchor = anchor_after(page_no if where == "top" else page_no + 1)
        if anchor is None:
            body.append(group[0])
            for el in group[1:]:
                group[0].addnext(el)
            placed += 1
            continue
        for el in group:
            anchor.addprevious(el)
        if where == "top" or not prose_pages.get(page_no, True):
            # the float leads its page, so it carries the break, not the prose
            if where == "top" and page_no in starts:
                _set_page_break_before(starts[page_no], False)
            _set_page_break_before(group[0], True)
            starts[page_no] = group[0]
        placed += 1

    drop_redundant_breaks(doc)
    return doc, len(starts), placed, orphaned


def fit(docx_path, target_pdf, tex_src, out_path, log=print, max_rounds=10,
        min_margin_in=0.45, start_margin_in=1.0):
    """Match the LaTeX page boundaries exactly.

    The text block is widened uniformly (one setting for the whole document,
    as asked) until no page overflows, then every page is verified against the
    LaTeX PDF -- first word, last word and prose word count -- and the loop
    repeats while anything still disagrees.
    """
    placements = floats.placements(target_pdf)
    captions = caption_index(tex_src)
    pdf_words, counts = page_word_counts(target_pdf, placements, captions)
    target_pages = len(counts)
    n_images_before = image_count(docx_path)
    log(f"ground truth: {target_pages} pages, {len(pdf_words)} prose words, "
        f"{len(placements)} floats, {n_images_before} images")

    margin, spacing, fig_scale = start_margin_in, 1.0, 1.0
    workdir = tempfile.mkdtemp(prefix="owlfit_")
    best = None

    for round_no in range(1, max_rounds + 1):
        d = Density(margin_in=margin, line_spacing=spacing,
                    figure_scale=fig_scale)
        doc, n_starts, n_floats, orphaned = build_candidate(
            docx_path, pdf_words, counts, placements, d)

        candidate = os.path.join(workdir, f"try{round_no}.docx")
        doc.save(candidate)
        rendered = render_pdf(candidate, workdir)
        pages = page_count(rendered)
        checks = verify(target_pdf, rendered, placements, captions)
        good = sum(1 for c in checks if c.ok)
        imgs = image_count(candidate)

        log(f"  round {round_no}: {d} -> {pages} pages (want {target_pages}), "
            f"{good}/{len(checks)} pages verified, {n_floats} floats, "
            f"{imgs} images")

        score = (good, -abs(pages - target_pages))
        if best is None or score > best[0]:
            best = (score, candidate, rendered, d, pages, checks, orphaned, imgs)

        if pages == target_pages and good == len(checks) and imgs == n_images_before:
            break

        # still overflowing somewhere: widen the text block uniformly.  Margins
        # first -- that was the explicit preference -- then leading, then the
        # figures, which are the last thing worth shrinking.
        if margin > min_margin_in:
            margin = round(max(min_margin_in, margin - 0.05), 3)
        elif spacing > 0.92:
            spacing = round(spacing - 0.02, 3)
        elif fig_scale > 0.85:
            fig_scale = round(fig_scale - 0.05, 3)
        else:
            log("  no room left to tighten further")
            break

    _, candidate, rendered, d, pages, checks, orphaned, imgs = best
    shutil.copy(candidate, out_path)
    good = sum(1 for c in checks if c.ok)
    log(f"settled: {d}")
    log(f"  {pages}/{target_pages} pages, {good}/{len(checks)} verified, "
        f"{imgs}/{n_images_before} images kept")
    return {"density": d, "pages": pages, "target_pages": target_pages,
            "checks": checks, "verified": good, "render": rendered,
            "orphaned": orphaned, "images": imgs,
            "images_expected": n_images_before}


# ---------------------------------------------------------------------------
# uniform density controls
# ---------------------------------------------------------------------------

class Density:
    """One uniform setting for the whole document."""

    def __init__(self, margin_in=1.0, line_spacing=1.0, font_delta=0.0,
                 figure_scale=1.0, image_width_mm=None, compact_tables=True):
        self.margin_in = margin_in
        self.line_spacing = line_spacing
        self.font_delta = font_delta        # points added to the base size
        self.figure_scale = figure_scale
        # A fixed image width decouples figure height from the text block.
        # Sizing images as a fraction of text width is self-defeating here:
        # widening the block to win vertical room makes every picture taller
        # and takes more room away than it gained.
        self.image_width_mm = image_width_mm
        self.compact_tables = compact_tables

    def __repr__(self):
        img = (f"images={self.image_width_mm:.0f}mm" if self.image_width_mm
               else f"figures={self.figure_scale:.2f}")
        return (f"margins={self.margin_in:.2f}in spacing={self.line_spacing:.3f} "
                f"font{self.font_delta:+.1f}pt {img}")


def apply_density(doc, d):
    """Apply a Density uniformly."""
    twips = int(d.margin_in * 1440)
    for section in doc.sections:
        section.left_margin = Twips(twips)
        section.right_margin = Twips(twips)
        section.top_margin = Twips(twips)
        section.bottom_margin = Twips(twips)

    if d.line_spacing != 1.0 or d.font_delta:
        _retune_styles(doc, d)

    # NB: a fixed image width is applied later, per float, in pagebuild --
    # applying it here would also hit inline images that are not figures at
    # all, such as a crest on the title page sized in absolute units.
    if d.figure_scale != 1.0:
        for i in range(len(doc.inline_shapes)):
            shape = doc.inline_shapes[i]
            shape.width = int(shape.width * d.figure_scale)
            shape.height = int(shape.height * d.figure_scale)

    if d.compact_tables:
        compact_tables(doc)


def _retune_styles(doc, d):
    """Apply the size and leading change to *every* style, not just Normal.

    pandoc's output uses its own paragraph styles (BodyText, FirstParagraph,
    Compact, ...) and those carry their own explicit sizes.  Changing only
    Normal therefore changes nothing on screen -- which is exactly why earlier
    tightening rounds produced identical renders at different settings.
    """
    from docx.oxml import OxmlElement

    def set_size(rPr, pt):
        half = str(int(round(pt * 2)))
        for tag in ("w:sz", "w:szCs"):
            for old in rPr.findall(qn(tag)):
                rPr.remove(old)
            el = OxmlElement(tag)
            el.set(qn("w:val"), half)
            rPr.append(el)

    def set_spacing(pPr, mult):
        for old in pPr.findall(qn("w:spacing")):
            line = old.get(qn("w:line"))
            if line:
                old.set(qn("w:line"), str(max(120, int(int(line) * mult))))
                old.set(qn("w:lineRule"), "auto")
                return
            pPr.remove(old)
        el = OxmlElement("w:spacing")
        el.set(qn("w:line"), str(int(240 * mult)))
        el.set(qn("w:lineRule"), "auto")
        pPr.append(el)

    styles_el = doc.styles.element
    for defaults in styles_el.findall(qn("w:docDefaults")):
        for rPrDef in defaults.findall(qn("w:rPrDefault")):
            rPr = rPrDef.find(qn("w:rPr"))
            if rPr is None:
                rPr = OxmlElement("w:rPr")
                rPrDef.append(rPr)
            cur = rPr.find(qn("w:sz"))
            base = int(cur.get(qn("w:val"))) / 2 if cur is not None else 12.0
            if d.font_delta:
                set_size(rPr, base + d.font_delta)
        for pPrDef in defaults.findall(qn("w:pPrDefault")):
            pPr = pPrDef.find(qn("w:pPr"))
            if pPr is None:
                pPr = OxmlElement("w:pPr")
                pPrDef.append(pPr)
            if d.line_spacing != 1.0:
                set_spacing(pPr, d.line_spacing)

    for style in styles_el.findall(qn("w:style")):
        rPr = style.find(qn("w:rPr"))
        if rPr is not None and d.font_delta:
            cur = rPr.find(qn("w:sz"))
            if cur is not None:
                set_size(rPr, int(cur.get(qn("w:val"))) / 2 + d.font_delta)
        pPr = style.find(qn("w:pPr"))
        if d.line_spacing != 1.0:
            if pPr is None:
                pPr = OxmlElement("w:pPr")
                style.insert(0, pPr)
            set_spacing(pPr, d.line_spacing)

    if d.font_delta:
        for rPr in doc.element.body.iter(qn("w:rPr")):
            cur = rPr.find(qn("w:sz"))
            if cur is not None:
                set_size(rPr, int(cur.get(qn("w:val"))) / 2 + d.font_delta)


def compact_tables(doc, text_width_emu=None):
    """Give table columns the width their contents actually need.

    pandoc derives fixed column widths from the LaTeX column spec and they come
    out consistently too narrow: a cell like "ACC TAC AC" that is one line in
    the original wraps onto two, so an eight-row table renders sixteen rows
    tall.  Font size and cell padding are left exactly as they are -- this only
    stops Word wrapping text the original never wrapped.
    """
    from docx.oxml import OxmlElement

    if text_width_emu is None:
        sec = doc.sections[0]
        text_width_emu = ((sec.page_width or 7772400)
                          - (sec.left_margin or 914400)
                          - (sec.right_margin or 914400))
    total_twips = int(text_width_emu / 914400 * 1440)

    for table in doc.tables:
        rows = table.rows
        if not rows:
            continue
        n_cols = len(rows[0].cells)
        need = [1] * n_cols
        for row in rows:
            cells = row.cells
            for k in range(min(n_cols, len(cells))):
                text = cells[k].text.strip()
                if not text:
                    continue
                longest = max((len(w) for w in text.split()), default=1)
                need[k] = max(need[k], longest, min(len(text), longest + 6))
        span = sum(need) or 1
        widths = [max(600, int(total_twips * n / span)) for n in need]
        over = sum(widths) - total_twips
        if over > 0:
            widths = [max(600, w - int(over * w / sum(widths))) for w in widths]

        tblPr = table._tbl.tblPr
        for tag in ("w:tblLayout", "w:tblW"):
            for old in tblPr.findall(qn(tag)):
                tblPr.remove(old)
        layout = OxmlElement("w:tblLayout")
        layout.set(qn("w:type"), "fixed")
        tblPr.append(layout)
        tw = OxmlElement("w:tblW")
        tw.set(qn("w:w"), str(sum(widths)))
        tw.set(qn("w:type"), "dxa")
        tblPr.append(tw)

        grid = table._tbl.find(qn("w:tblGrid"))
        if grid is not None:
            for gc, wdt in zip(grid.findall(qn("w:gridCol")), widths):
                gc.set(qn("w:w"), str(wdt))
        for row in rows:
            cells = row.cells
            for k in range(min(n_cols, len(cells))):
                tcPr = cells[k]._tc.get_or_add_tcPr()
                for old in tcPr.findall(qn("w:tcW")):
                    tcPr.remove(old)
                el = OxmlElement("w:tcW")
                el.set(qn("w:w"), str(widths[k]))
                el.set(qn("w:type"), "dxa")
                tcPr.append(el)


def set_float_image_width(blocks, mm, max_height_mm=None):
    """Give every *figure* the same fixed width, in millimetres.

    Applied only to images belonging to a float block, so an inline crest or
    logo keeps the absolute size the LaTeX asked for.

    `max_height_mm` caps a tall figure: at a fixed width, a portrait graphic
    can be taller than the text frame once its caption is added, and then no
    amount of margin tuning will keep it on one page -- it simply spills and
    drags a blank page after it.
    """
    target = int(mm * 36000)                    # 1 mm = 36000 EMU
    cap = int(max_height_mm * 36000) if max_height_mm else None
    changed = 0
    for _, _, group in blocks:
        for el in group:
            for ext in el.iter(qn("wp:extent")):
                cx, cy = int(ext.get("cx")), int(ext.get("cy"))
                if not cx:
                    continue
                new_cx, new_cy = target, int(target * cy / cx)
                if cap and new_cy > cap:
                    new_cx = int(cap * cx / cy)
                    new_cy = cap
                ext.set("cx", str(new_cx))
                ext.set("cy", str(new_cy))
                for a_ext in el.iter(qn("a:ext")):
                    a_ext.set("cx", str(new_cx))
                    a_ext.set("cy", str(new_cy))
                changed += 1
    return changed


# ---------------------------------------------------------------------------
# measuring
# ---------------------------------------------------------------------------

def render_pdf(docx_path, outdir=None):
    """Render a .docx to PDF with LibreOffice, for measurement only."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("LibreOffice is required to measure page fit")
    outdir = outdir or os.path.dirname(os.path.abspath(docx_path))
    subprocess.run([soffice, "--headless", "--convert-to", "pdf",
                    "--outdir", outdir, docx_path],
                   capture_output=True, timeout=900)
    out = os.path.join(outdir,
                       os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    if not os.path.exists(out):
        raise RuntimeError("LibreOffice produced no PDF")
    return out


def page_count(pdf_path):
    out = subprocess.run(["pdfinfo", pdf_path], capture_output=True, text=True)
    m = re.search(r"^Pages:\s+(\d+)", out.stdout, re.M)
    return int(m.group(1)) if m else 0


class PageCheck:
    """The verification for one page: does it start and end on the same word
    as the LaTeX PDF, and does it hold the same amount of prose?"""

    def __init__(self, page, first_ok, last_ok, target_n, got_n,
                 target_first, got_first, target_last, got_last):
        self.page = page
        self.first_ok = first_ok
        self.last_ok = last_ok
        self.target_n = target_n
        self.got_n = got_n
        self.target_first = target_first
        self.got_first = got_first
        self.target_last = target_last
        self.got_last = got_last

    @property
    def ok(self):
        return self.first_ok and self.last_ok

    def describe(self):
        if self.ok:
            return f"page {self.page}: ok ({self.got_n} words)"
        bits = []
        if not self.first_ok:
            bits.append(f"starts {self.got_first!r}, expected {self.target_first!r}")
        if not self.last_ok:
            bits.append(f"ends {self.got_last!r}, expected {self.target_last!r}")
        return f"page {self.page}: " + "; ".join(bits)


def verify(target_pdf, produced_pdf, placements=None, captions=None, sample=4):
    """Compare the produced PDF against the LaTeX PDF page by page.

    Both sides go through the same prose extractor, so figure internals and
    captions are excluded from both and the comparison is like for like.
    """
    tgt = pdfprose.strip_running_heads(
        pdfprose.prose_pages(target_pdf, placements, captions))
    got = pdfprose.strip_running_heads(
        pdfprose.prose_pages(produced_pdf, placements, captions))

    checks = []
    for i in range(max(len(tgt), len(got))):
        t = [w for w in (normalise(x) for x in (tgt[i] if i < len(tgt) else [])) if w]
        g = [w for w in (normalise(x) for x in (got[i] if i < len(got) else [])) if w]
        tf, gf = " ".join(t[:sample]), " ".join(g[:sample])
        tl, gl = " ".join(t[-sample:]), " ".join(g[-sample:])
        checks.append(PageCheck(i + 1, tf == gf, tl == gl, len(t), len(g),
                                tf, gf, tl, gl))
    return checks


def image_count(docx_path):
    from docx import Document as _Doc
    return len(_Doc(docx_path).inline_shapes)


def caption_index(tex_src):
    """{(kind, printed_number): caption_words} for every float."""
    from .counters import scan_preamble, float_numbers
    model = scan_preamble(tex_src)
    fig_nums, tab_nums = float_numbers(tex_src, model)
    out, fi, ti = {}, 0, 0
    for kind, words, body in floats.source_captions(tex_src):
        if kind == "Figure":
            if fi < len(fig_nums):
                out[("Figure", fig_nums[fi])] = (words, body)
            fi += 1
        else:
            if ti < len(tab_nums):
                out[("Table", tab_nums[ti])] = (words, body)
            ti += 1
    return out
