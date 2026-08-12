# 🦉 OWL — Overleaf · Word · LaTeX

Turns your **Overleaf** sources (LaTeX + `references.bib` + figures) into an
**editable Word `.docx`** — Times New Roman, justified, title page, running
header, real Word equations, numbered figure and table captions, clickable
cross-references, and a static ACS reference list with live DOI links.

No LaTeX install and no GPU: it uses pandoc, not a LaTeX compile.

---

## Layout

```
app.py                  Streamlit UI — upload, convert, download. No logic.
tex2docx.py             CLI front end (same pipeline as the app)
owlkit/
  convert.py            the one entry point: preflight → pandoc → postprocess
  preflight.py          validates the .tex and reports source line numbers
  counters.py           LaTeX counter machine: what number each \label prints
  engine.py             preprocess, figure resolution, docx postprocessing
  floats.py             where the compiled PDF actually placed each float
  pagefit.py            page-boundary matching (experimental — see below)
  constants.py          shared markers and asset paths
assets/
  reference.docx        the Word style template
  american-chemical-society.csl
tests/test_owl.py       regression tests for every bug listed below
```

Everything goes through `owlkit.convert()`, so the app and the CLI cannot
drift apart.

---

## Run it locally

```bash
pip install -r requirements.txt
# system tools, once:
#   macOS:   brew install pandoc poppler
#   Ubuntu:  sudo apt-get install -y pandoc poppler-utils
streamlit run app.py          # or: python tex2docx.py main.tex
python tests/test_owl.py      # regression tests
```

Set `OWL_PASSWORD` in Streamlit **Secrets**. There is deliberately no default
password in the code — if the secret is missing the app refuses to unlock
rather than falling back to a value that is public in this repository.

---

## Preflight

OWL checks the source *before* pandoc sees it, and reports a **line number**
and a fix for each problem. This exists because a single stray `{` used to
surface as

```
Error at "....pre.tex" (line 686, column 1): unexpected \end
```

— pointing at `\end{document}`, hundreds of lines from the actual typo.

It checks for: unbalanced braces (per paragraph), unbalanced environments,
citations with no `.bib`, figures that do not resolve, `\textcolor` inside
`$…$`, duplicate labels, and cross-references to labels that do not exist.

**A missing `.bib` is an error, not a warning.** Pandoc without a bibliography
converts *successfully* and silently discards every citation and the whole
reference list — a worse outcome than a clear refusal. Untick "Stop if the
LaTeX has problems" to convert anyway.

---

## Fixed in the rebuild

- **`\includegraphics[width=…]{…}` was never processed.** Figure conversion
  matched the literal token `\includegraphics{`, so every call with an
  optional argument — i.e. essentially all of them — skipped PDF→PNG
  conversion, and raw PDFs got embedded into the `.docx`, which Word cannot
  display.
- **`\graphicspath` and extension-less names were ignored.** Only literal
  `.pdf`/`.eps` names in the main directory resolved.
- **Cross-references to user-defined sectioning macros were dead.** A macro
  such as `\appsection` that steps its own counter fell through the label map,
  so every `\cref{app:…}` printed the raw label text. `counters.py` now reads
  the preamble the way LaTeX does — `\newcounter`, `\the…`, `\crefname`,
  `\refstepcounter` inside macro bodies — and simulates the counters.
- **Equation cross-references** (`\cref{eq:…}`) were unresolved for the same
  reason.
- **Comments were parsed as code**, so an `\appsection` mentioned in a `%`
  comment shifted every appendix letter by one.
- **Supplementary figures were misnumbered.** Captions were numbered by
  position, ignoring `\setcounter{figure}{0}` plus
  `\renewcommand{\thefigure}{S\arabic{figure}}` — so "Figure S1" came out as
  "Figure 6" while the cross-reference pointing at it said something else.
- **The password had a hard-coded fallback** in the source of this repository.

If an `.aux` file from a real LaTeX run is available, pass `--aux main.aux`
and its numbers are used instead of OWL's reconstruction — LaTeX is always
right about its own numbering.

---

## Fixed in this pass

- **Figures were the wrong size.** Widths were collected by scanning every
  `\includegraphics` in the raw source, including one inside a `%` comment in
  the preamble and one whose file was missing. Both shift the positional
  mapping, so every picture got another picture's width — the first real
  figure fell back to its natural size, about 38% of the text width instead of
  the 85% the LaTeX asked for. Widths are now collected *during* the rewrite,
  only for figures that actually resolve, so the list cannot drift out of step
  with the images pandoc emits. Absolute units (`width=2in`, `3cm`, `100pt`)
  are honoured too.
- **`\textcolor` was dropped entirely.** pandoc has no colour model for the
  LaTeX reader, so a document using red for open questions and green for
  settled ones arrived uniformly black. `colors.py` reads `\definecolor` —
  respecting the case-sensitive difference between xcolor's `{rgb}` fractions
  and `{RGB}` integers — and re-applies real `w:color` runs, including across
  runs that pandoc split for emphasis, maths or links.
- **Comments are stripped before preprocessing**, so nothing commented out is
  ever treated as content.

## Page matching

Every Word page starts and ends on the same word as the compiled LaTeX PDF.
This is on by default in the app and requires the PDF.

**The PDF is not optional and cannot be worked around.** Nothing in the `.tex`
says where the page breaks fall — they are the output of LaTeX's line-breaking
algorithm run against your fonts and margins, and only exist once the document
has been compiled.

```python
from owlkit import convert
convert("main.tex", match_pdf="main.pdf", on_progress=print)
```

Measured on the reference document (36-page PDF, 14 floats, 52 references):

| | |
|---|---|
| pages | **36 / 36, confirmed in Word** |
| page starts matching the PDF | 31–35 of 36, depending on how strictly you count |
| total boundary error | **4 words** |
| source prose present | **99.2%** |
| wall clock | ~2.5 minutes |

### How it works

Each page is treated as its own object — `[float at top] + prose + [float at
bottom]` — rendered **on its own**, and asked whether it fits **with an inch of
vertical room to spare**. The settings are tightened down a ladder (margins
first, then figure width, then leading) until every page clears that bar; the
result is one uniform layout for the whole document.

The inch matters. Measurement happens in LibreOffice; the document is opened in
Word; the two lay text out differently. Sizing each page to *just* fit gave 38
pages in Word, then 37. An inch of headroom gave 36.

### What it needs

**LibreOffice**, to render the page probes — hence `libreoffice-writer` in
`packages.txt`. It makes the build heavier and the conversion slower, which is
the price of the feature.

## Deploy free on Streamlit Community Cloud

1. Push this folder to a GitHub repo.
2. share.streamlit.io → **Create app** → branch `main`, main file `app.py`.
3. **Advanced → Secrets:**
   ```toml
   OWL_PASSWORD = "your-password-here"
   ```
4. Deploy. `packages.txt` (apt) and `requirements.txt` (pip) install
   automatically. Every push redeploys.

`packages.txt` deliberately does **not** install `libreoffice-writer` any
more: nothing in the conversion path used it, and it added hundreds of
megabytes to every build.

---

© 2026 David G. Schauer · All rights reserved. OWL — the app, workflow and
code — is an original work of the author. It builds on the open-source
[pandoc](https://pandoc.org) and
[python-docx](https://python-docx.readthedocs.io) projects.
