"""The one entry point everything else calls.

    result = convert("main.tex")

`app.py` and the CLI both go through here, so the UI carries no conversion
logic and the two can never drift apart.
"""

import os
import shutil
import subprocess
import sys

from . import preflight
from .counters import (scan_preamble, build_label_map, float_numbers,
                       labels_from_aux, section_numbers)
from .engine import (find_bib, scan_floats, collect_referenced, preprocess,
                     postprocess_docx, ensure_csl, make_resolver, log)
from .constants import REFERENCE_DOCX


class ConversionError(Exception):
    """Raised when the document cannot be converted.  `issues` carries the
    preflight findings so the UI can show line numbers rather than a stack
    trace."""

    def __init__(self, message, issues=None, detail=None):
        super().__init__(message)
        self.issues = issues or []
        self.detail = detail


class Result:
    match = None            # page-matching report, when a PDF was supplied

    def __init__(self, docx_path, warnings, bib, labels, log_lines):
        self.docx_path = docx_path
        self.warnings = warnings
        self.bib = bib
        self.labels = labels
        self.log_lines = log_lines


def convert(tex_path, output=None, keep_drafts=False, strict=True,
            aux_path=None, on_log=None, match_pdf=None, on_progress=None):
    """LaTeX -> editable .docx.

    strict=True refuses to produce a file when preflight finds errors, which
    is the point: a document that silently loses its citations, or that pandoc
    mangles from a stray brace, is worse than a clear refusal.

    `match_pdf` is the compiled PDF.  Given one, the pages of the .docx are
    rebuilt to break exactly where that PDF's pages break.  It is required for
    that: nothing in the .tex says where the page breaks fall -- they are the
    output of a line-breaking algorithm run against the fonts and margins, and
    only exist once the document has been compiled.

    `on_progress(fraction, message)` is called throughout, for a progress bar.
    """
    lines = []

    def emit(msg):
        lines.append(msg)
        if on_log:
            on_log(msg)

    def progress(frac, msg):
        if on_progress:
            on_progress(max(0.0, min(1.0, frac)), msg)

    if not shutil.which("pandoc"):
        raise ConversionError(
            "pandoc is not installed on this machine.",
            detail="Install it with `brew install pandoc` (macOS) or "
                   "`apt-get install pandoc` (Linux). On Streamlit Cloud it "
                   "comes from packages.txt.")

    tex_path = os.path.abspath(tex_path)
    if not os.path.exists(tex_path):
        raise ConversionError(f"no such file: {tex_path}")
    tex_dir = os.path.dirname(tex_path)
    out_path = os.path.abspath(output) if output else \
        os.path.splitext(tex_path)[0] + ".docx"

    with open(tex_path, encoding="utf-8", errors="replace") as f:
        src = f.read()

    bib = find_bib(src, tex_dir)
    emit(f"bibliography: {os.path.basename(bib)}" if bib else
         "no .bib found")

    # ---- the label map -------------------------------------------------
    model = scan_preamble(src)
    labels = build_label_map(src, model)
    if aux_path and os.path.exists(aux_path):
        # LaTeX's own numbers beat any reconstruction of them
        with open(aux_path, encoding="utf-8", errors="replace") as f:
            labels.update(labels_from_aux(f.read(), model))
        emit(f"label numbers taken from {os.path.basename(aux_path)}")

    # ---- preflight -----------------------------------------------------
    resolver = make_resolver(src, tex_dir)
    errors, warnings = preflight.run(src, bib_path=bib,
                                     figure_resolver=resolver, labels=labels)
    if errors and strict:
        raise ConversionError(
            f"the LaTeX source has {len(errors)} problem(s) that would break "
            f"or silently damage the conversion.", issues=errors)
    for w in warnings:
        emit("warning: " + w.as_text().replace("\n    \u2192 ", " — "))

    # ---- convert -------------------------------------------------------
    import re
    m = (re.search(r"\\newcommand\{\\shorttitle\}\{(.+?)\}", src)
         or re.search(r"\\newcommand\{\\reporttitle\}\{(.+?)\}", src))
    header_title = m.group(1) if m else os.path.splitext(os.path.basename(tex_path))[0]
    header_title = (header_title.replace("---", "\u2014").replace("--", "\u2013")
                    .replace("\\&", "&").strip())

    fig_nums, tab_nums = float_numbers(src, scan_preamble(src))
    _, fig_labels, tab_labels = scan_floats(src)
    referenced = collect_referenced(src)
    # img_widths comes back from preprocess so it lines up with the images
    # pandoc actually emits, not with every \includegraphics in the source
    pre, img_widths = preprocess(src, tex_dir, labels, keep_drafts, model)
    floats_info = (img_widths, fig_labels, tab_labels)

    pre_path = os.path.join(tex_dir, ".owl.pre.tex")
    with open(pre_path, "w", encoding="utf-8") as f:
        f.write(pre)

    cmd = ["pandoc", pre_path, "-o", out_path,
           "--from", "latex+raw_tex",
           "--resource-path", tex_dir]
    if os.path.exists(REFERENCE_DOCX):
        cmd += ["--reference-doc", REFERENCE_DOCX]
    if bib:
        csl = ensure_csl()
        cmd += ["--citeproc", "--bibliography", bib,
                "--metadata", "reference-section-title=References",
                "--metadata", "link-citations=true"]
        if csl:
            cmd += ["--csl", csl]

    progress(0.05, "checking the LaTeX source")
    emit("running pandoc…")
    progress(0.10, "converting with pandoc")
    proc = subprocess.run(cmd, capture_output=True, text=True)
    try:
        os.remove(pre_path)
    except OSError:
        pass
    if proc.returncode != 0:
        raise ConversionError("pandoc could not read the LaTeX source.",
                              detail=(proc.stderr or proc.stdout or "")[-3000:])

    progress(0.20, "rebuilding figures, captions and cross-references")
    postprocess_docx(out_path, header_title, floats_info, referenced,
                     numbering=(fig_nums, tab_nums))

    # restore heading numbers and carry \small / \footnotesize declarations
    # from inside float environments across -- pandoc drops both
    try:
        from docx import Document as _Doc
        from . import floats as _floats
        _d = _Doc(out_path)
        _blocks = _floats.find_blocks(_d)
        _sizes = _floats.source_float_sizes(src)
        _changed = bool(_floats.apply_float_sizes(_d, _blocks, _sizes))
        _changed |= bool(number_headings(_d, src))
        if _changed:
            _d.save(out_path)
    except Exception as exc:                       # never fail the conversion
        emit(f"note: could not apply float font sizes ({exc})")
    match = None
    if match_pdf:
        if not os.path.exists(match_pdf):
            raise ConversionError(f"no such PDF: {match_pdf}")
        from .pagebuild import match_pages
        emit("matching pages to the compiled PDF…")

        def relay(frac, msg):
            emit(msg)
            progress(0.25 + 0.72 * frac, msg)

        match = match_pages(out_path, match_pdf, src, out_path,
                            on_progress=relay)
        emit(f"pages: {match['pages']} (PDF has {match['target_pages']}); "
             f"{match['exact_starts']}/{len(match['checks'])} start on the "
             f"same word")

    progress(1.0, "done")
    emit(f"done: {os.path.basename(out_path)}")
    result = Result(out_path, warnings, bib, labels, lines)
    result.match = match
    return result


def number_headings(doc, tex_src):
    """Put the section numbers back on the headings.

    LaTeX prints "2.2 Experiment 1: ..."; pandoc maps the heading to a Word
    Heading style, which carries no number of its own, so the number is simply
    lost.  Beyond looking wrong, it makes a heading on a page boundary read as
    different text on the two sides.
    """
    import re as _re
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    numbers = section_numbers(tex_src)
    heads = []
    for p in doc.paragraphs:
        name = (p.style.name if p.style else "") or ""
        m = _re.match(r"Heading\s*([1-6])$", name.strip())
        if m:
            heads.append((int(m.group(1)), p))

    applied = 0
    for (level, num), (h_level, para) in zip(numbers, heads):
        if not num or level != h_level:
            continue
        if para.text.strip().startswith(num + " "):
            continue
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = f"{num} "
        t.set(qn("xml:space"), "preserve")
        r.append(t)
        pPr = para._p.find(qn("w:pPr"))
        (pPr.addnext(r) if pPr is not None else para._p.insert(0, r))
        applied += 1
    return applied
