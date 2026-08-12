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
        for run in para.runs:
            text = run.text
            if not text:
                continue
            for m in re.finditer(r"\S+", text):
                w = normalise(m.group(0))
                if w:
                    words.append(w)
                    sites.append((para, run, m.start()))
    return words, sites


# ---------------------------------------------------------------------------
# alignment
# ---------------------------------------------------------------------------

def align_boundaries(pdf_words, page_counts, docx_words):
    """Map each PDF page boundary onto an index in the Word word stream.

    The two streams are never identical -- equations, captions and table cells
    serialise differently -- so this uses a longest-matching-block alignment
    and, for each boundary, takes the nearest reliable anchor.
    """
    matcher = difflib.SequenceMatcher(None, pdf_words, docx_words, autojunk=False)
    blocks = matcher.get_matching_blocks()

    boundaries = []
    cum = 0
    for count in page_counts[:-1]:          # no break after the final page
        cum += count
        boundaries.append(cum)

    mapped = []
    for b in boundaries:
        best = None
        for a, bx, size in blocks:
            if size == 0:
                continue
            if a <= b <= a + size:                       # inside a matched run
                best = bx + (b - a)
                break
            # otherwise remember the closest block that ends before b
            if a + size <= b:
                cand = bx + size
                if best is None or cand > best:
                    best = cand
        mapped.append(best)
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
    """Split `para` so that everything from `offset` in `run` onwards lives in
    a new paragraph immediately after it.  Returns the new paragraph element.

    The visual seam is hidden: the first half gets last-line justification so
    its final line still reaches the right margin like a mid-paragraph line
    would, and the second half loses its first-line indent and space-before so
    it reads as a continuation rather than a new paragraph.
    """
    p_el = para._p
    tail_el = copy.deepcopy(p_el)

    runs = list(p_el.findall(qn("w:r")))
    tail_runs = list(tail_el.findall(qn("w:r")))
    try:
        pos = runs.index(run._r)
    except ValueError:
        pos = 0

    if offset:
        text = run.text
        run.text = text[:offset]
        for t in tail_runs[pos].findall(qn("w:t")):
            tail_runs[pos].remove(t)
        t = tail_runs[pos].makeelement(qn("w:t"), {})
        t.text = text[offset:]
        t.set(qn("xml:space"), "preserve")
        tail_runs[pos].append(t)
        keep_from = pos
    else:
        keep_from = pos

    # head keeps runs [0, keep_from) plus the truncated one
    for r in runs[keep_from + (1 if offset else 0):]:
        p_el.remove(r)
    # tail keeps runs [keep_from, end)
    for r in tail_runs[:keep_from]:
        tail_el.remove(r)

    # hide the seam
    head_pPr = _pPr(p_el)
    jc = head_pPr.find(qn("w:jc"))
    if jc is None:
        jc = head_pPr.makeelement(qn("w:jc"), {})
        head_pPr.append(jc)
    if jc.get(qn("w:val")) in (None, "both", "justify"):
        jc.set(qn("w:val"), "distribute")     # justify the final line too

    tail_pPr = _pPr(tail_el)
    for tag in ("w:ind", "w:spacing"):
        el = tail_pPr.find(qn(tag))
        if el is not None:
            el.set(qn("w:firstLine"), "0") if tag == "w:ind" else None
            if tag == "w:spacing":
                el.set(qn("w:before"), "0")

    p_el.addnext(tail_el)
    return tail_el


# ---------------------------------------------------------------------------
# uniform density controls
# ---------------------------------------------------------------------------

class Density:
    """One uniform setting for the whole document."""

    def __init__(self, margin_in=1.0, line_spacing=1.0, font_delta=0.0,
                 figure_scale=1.0):
        self.margin_in = margin_in
        self.line_spacing = line_spacing
        self.font_delta = font_delta        # points added to the base size
        self.figure_scale = figure_scale

    def __repr__(self):
        return (f"margins={self.margin_in:.2f}in spacing={self.line_spacing:.3f} "
                f"font{self.font_delta:+.1f}pt figures={self.figure_scale:.2f}")


def apply_density(doc, d):
    """Apply a Density uniformly: page margins, line spacing, font size and
    figure scale."""
    twips = int(d.margin_in * 1440)
    for section in doc.sections:
        section.left_margin = Twips(twips)
        section.right_margin = Twips(twips)
        section.top_margin = Twips(twips)
        section.bottom_margin = Twips(twips)

    if d.line_spacing != 1.0 or d.font_delta:
        style = doc.styles["Normal"]
        if d.font_delta and style.font.size is not None:
            style.font.size = Pt(style.font.size.pt + d.font_delta)
        elif d.font_delta:
            style.font.size = Pt(12 + d.font_delta)
        pf = style.paragraph_format
        pf.line_spacing = d.line_spacing

    if d.figure_scale != 1.0:
        for shape in doc.inline_shapes:
            shape.width = int(shape.width * d.figure_scale)
            shape.height = int(shape.height * d.figure_scale)


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
                   capture_output=True, timeout=600)
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
    from docx import Document
    return len(Document(docx_path).inline_shapes)


# ---------------------------------------------------------------------------
# the fitting loop
# ---------------------------------------------------------------------------

def caption_index(tex_src):
    """{(kind, printed_number): caption_words} for every float."""
    from .counters import scan_preamble, float_numbers
    model = scan_preamble(tex_src)
    fig_nums, tab_nums = float_numbers(tex_src, model)
    out, fi, ti = {}, 0, 0
    for kind, words in floats.source_captions(tex_src):
        if kind == "Figure":
            if fi < len(fig_nums):
                out[("Figure", fig_nums[fi])] = words
            fi += 1
        else:
            if ti < len(tab_nums):
                out[("Table", tab_nums[ti])] = words
            ti += 1
    return out


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
    """Paragraphs carrying an inline <w:sectPr>.

    The postprocessor creates sections so the appendix can renumber its pages.
    A section break is itself a page break, so if a ground-truth boundary lands
    on the same spot and we also set pageBreakBefore, Word emits *two* breaks
    and leaves a blank page -- which then cascades through every later
    boundary.
    """
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
        ordered.append((page_no, idx))

    # back to front: splitting a later paragraph cannot disturb earlier sites
    for page_no, idx in sorted(ordered, key=lambda t: t[1], reverse=True):
        para, run, offset = sites[idx]
        tail = split_paragraph_at(para, run, offset)
        _set_page_break_before(tail, True)
        starts[page_no] = tail
    return starts


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
