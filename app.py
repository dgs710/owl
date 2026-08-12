"""OWL — Overleaf · Word · LaTeX.

A thin Streamlit front end.  All conversion logic lives in `owlkit`, so the UI
cannot drift away from the CLI, and both refuse the same bad input for the
same reasons.
"""

import io
import os
import shutil
import tempfile
import zipfile

import streamlit as st

from owlkit import convert, ConversionError

st.set_page_config(page_title="OWL · Overleaf → Word", page_icon="🦉",
                   layout="centered")

# ── styling to echo the website (dark + emerald) ──────────────────────────
st.markdown("""
<style>
  :root { --acc:#34D399; }
  .stApp { background:#06090B; color:#E9EFEC; }
  h1, h2, h3, h4 { font-family: 'Space Grotesk','Inter',sans-serif; }
  .owl-title { font-size:2.4rem; font-weight:700; letter-spacing:.06em;
    background:linear-gradient(100deg,#6EE7B7,#34D399);
    -webkit-background-clip:text; background-clip:text;
    -webkit-text-fill-color:transparent; margin:0; }
  .owl-sub { font-family:'JetBrains Mono',monospace; font-size:.72rem;
    letter-spacing:.24em; text-transform:uppercase; color:#7E8B85; margin-top:.2rem; }
  .stButton>button, .stDownloadButton>button {
    background:#34D399; color:#04110B; border:none; border-radius:99px;
    font-weight:600; padding:.5rem 1.4rem; }
  .stButton>button:hover, .stDownloadButton>button:hover { background:#6EE7B7; }
  .owl-note { font-size:.82rem; color:#B7C2BC; line-height:1.7; }
  .owl-issue { font-family:'JetBrains Mono',monospace; font-size:.8rem;
    border-left:3px solid #F87171; padding:.4rem .8rem; margin:.35rem 0;
    background:rgba(248,113,113,.07); }
  .owl-warn { border-left-color:#FBBF24; background:rgba(251,191,36,.07); }
  code { color:#6EE7B7; }
</style>
""", unsafe_allow_html=True)

st.markdown('<p class="owl-title">🦉 OWL</p>', unsafe_allow_html=True)
st.markdown('<p class="owl-sub">Overleaf · Word · LaTeX — one upload, a Word bundle back</p>',
            unsafe_allow_html=True)
st.write("")


# ── password gate ──────────────────────────────────────────────────────────
def get_password():
    try:
        return st.secrets["OWL_PASSWORD"]
    except Exception:
        return None


if "owl_ok" not in st.session_state:
    st.session_state.owl_ok = False

if not st.session_state.owl_ok:
    st.markdown("#### 🔒 Access")
    expected = get_password()
    if not expected:
        st.error("No access password is configured. Set `OWL_PASSWORD` in the "
                 "app's Secrets.")
        st.stop()
    pw = st.text_input("Access password", type="password",
                       label_visibility="collapsed",
                       placeholder="Enter access password")
    if st.button("Unlock"):
        if pw == expected:
            st.session_state.owl_ok = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.caption("© 2026 David G. Schauer · All rights reserved. "
               "OWL is an original tool by the author.")
    st.stop()


# ── helpers ────────────────────────────────────────────────────────────────
def find_main_tex(root):
    """The .tex holding \\documentclass is the main file."""
    texs = []
    for dirpath, _, files in os.walk(root):
        if "__MACOSX" in dirpath:
            continue
        for f in files:
            if f.lower().endswith(".tex"):
                texs.append(os.path.join(dirpath, f))
    for probe in ("\\documentclass", "\\begin{document}"):
        for t in texs:
            try:
                with open(t, encoding="utf-8", errors="ignore") as fh:
                    if probe in fh.read():
                        return t
            except OSError:
                pass
    return texs[0] if texs else None


def show_issues(issues, kind="error"):
    css = "owl-issue" if kind == "error" else "owl-issue owl-warn"
    for i in issues:
        where = f"Line {i.line} — " if i.line else ""
        hint = f"<br><span style='opacity:.75'>→ {i.hint}</span>" if i.hint else ""
        st.markdown(f'<div class="{css}"><b>{where}</b>{i.message}{hint}</div>',
                    unsafe_allow_html=True)


# ── upload ─────────────────────────────────────────────────────────────────
st.markdown("### 1 · Upload your Overleaf sources")
st.markdown(
    '<p class="owl-note">In Overleaf use <b>Menu → Download → Source</b> and '
    'upload that <code>.zip</code> — it already contains your <code>.tex</code>, '
    '<code>references.bib</code> and figures. Without the <code>.bib</code> '
    'every citation and the whole reference list are lost, so OWL will stop '
    'and tell you rather than hand you a document that quietly dropped them.</p>',
    unsafe_allow_html=True)

up_zip = st.file_uploader("Overleaf source .zip", type=["zip"])
with st.expander("…or upload individual files instead"):
    up_tex = st.file_uploader("main.tex", type=["tex"])
    up_bib = st.file_uploader("references.bib", type=["bib"])
    up_figs = st.file_uploader("figures (png / jpg / pdf / eps)",
                               accept_multiple_files=True)

st.markdown("### 2 · Upload the compiled PDF  (required)")
st.markdown(
    '<p class="owl-note">Overleaf → <b>Menu → Download → PDF</b>. This is not '
    'optional and cannot be worked around: <b>nothing in the .tex says where '
    'the page breaks fall</b>. They are the result of LaTeX\'s line-breaking '
    'run against your fonts and margins, so they only exist once the document '
    'has been compiled. The PDF is what OWL matches the Word pages to.</p>',
    unsafe_allow_html=True)
up_pdf = st.file_uploader("Compiled PDF", type=["pdf"])

st.markdown("### 3 · Convert")
st.markdown(
    '<p class="owl-note">Produces an <b>editable Word</b> document — Times New '
    'Roman, justified, title page, running header, real Word equations, '
    'numbered headings, numbered figure and table captions, clickable '
    'cross-references, and a static ACS reference list built from your '
    '<code>.bib</code> with live DOI links — and its pages broken where your '
    'PDF\'s pages break. Expect <b>two to four minutes</b>: matching the '
    'pages means rendering every page on its own and checking it fits with '
    'room to spare.</p>', unsafe_allow_html=True)

strict = st.checkbox("Stop if the LaTeX has problems (recommended)", value=True,
                     help="Uncheck to convert anyway. Citations may be lost "
                          "and cross-references may print as raw labels.")

if st.button("Convert to Word →", type="primary"):
    workdir = tempfile.mkdtemp(prefix="owl_")
    try:
        if up_zip is not None:
            with zipfile.ZipFile(io.BytesIO(up_zip.read())) as z:
                z.extractall(workdir)
        elif up_tex is not None:
            with open(os.path.join(workdir, "main.tex"), "wb") as fh:
                fh.write(up_tex.getbuffer())
            if up_bib is not None:
                with open(os.path.join(workdir, up_bib.name), "wb") as fh:
                    fh.write(up_bib.getbuffer())
            for f in (up_figs or []):
                with open(os.path.join(workdir, f.name), "wb") as fh:
                    fh.write(f.getbuffer())
        else:
            st.warning("Please upload your Overleaf .zip (or at least a .tex).")
            st.stop()

        if up_pdf is None:
            st.error("The compiled PDF is required — it is the only thing that "
                     "says where your pages break. Overleaf → Menu → Download "
                     "→ PDF.")
            st.stop()

        main_tex = find_main_tex(workdir)
        if not main_tex:
            st.error("No .tex file found in the upload.")
            st.stop()

        srcdir = os.path.dirname(main_tex)
        paper = os.path.splitext(os.path.basename(main_tex))[0]
        with open(os.path.join(srcdir, paper + ".pdf"), "wb") as fh:
            fh.write(up_pdf.getbuffer())

        log_lines = []
        bar = st.progress(0.0)
        status = st.empty()

        def on_progress(frac, msg):
            bar.progress(frac)
            status.markdown(f"**{frac * 100:.0f}%** — {msg}")

        try:
            result = convert(main_tex, strict=strict,
                             on_log=log_lines.append,
                             match_pdf=os.path.join(srcdir, paper + ".pdf"),
                             on_progress=on_progress)
        except ConversionError as e:
            bar.empty()
            status.empty()
            st.error(str(e))
            if e.issues:
                show_issues(e.issues, "error")
            if e.detail:
                with st.expander("pandoc output"):
                    st.code(e.detail)
            st.stop()
        bar.empty()
        status.empty()

        m = getattr(result, "match", None)
        if m:
            same = m["pages"] == m["target_pages"]
            st.markdown(
                f"**Pages:** {m['pages']} — your PDF has {m['target_pages']}"
                + ("  ✅" if same else "  ⚠️ not the same") + "  \n"
                f"**Page starts matching the PDF exactly:** "
                f"{m['exact_starts']} of {len(m['checks'])}  \n"
                f"**Layout used:** {m['density']}")
            if not same:
                st.warning("The page count does not match. Send the .tex and "
                           "PDF on so the difference can be tracked down.")

        if result.warnings:
            with st.expander(f"⚠ {len(result.warnings)} warning(s) — "
                             f"the file was still produced"):
                show_issues(result.warnings, "warning")

        # ── bundle ──────────────────────────────────────────────────────
        out = result.docx_path
        IMG_EXT = (".png", ".jpg", ".jpeg", ".gif", ".eps", ".pdf",
                   ".svg", ".tif", ".tiff")
        pdf_path = os.path.join(srcdir, paper + ".pdf")
        bundle, included = io.BytesIO(), []
        with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as z:
            z.write(out, os.path.basename(out))
            included.append(os.path.basename(out))
            if os.path.exists(pdf_path):
                z.write(pdf_path, paper + ".pdf")
                included.append(paper + ".pdf")
            for f in sorted(os.listdir(srcdir)):
                if f.lower().endswith(".bib"):
                    z.write(os.path.join(srcdir, f), f)
                    included.append(f)
            seen = set()
            for dirpath, _, files in os.walk(srcdir):
                if "__MACOSX" in dirpath:
                    continue
                for f in sorted(files):
                    low = f.lower()
                    if not low.endswith(IMG_EXT) or low.endswith("_conv.png"):
                        continue
                    full = os.path.join(dirpath, f)
                    if os.path.abspath(full) == os.path.abspath(pdf_path):
                        continue
                    if f in seen:
                        continue
                    seen.add(f)
                    z.write(full, os.path.join("figures", f))
                    included.append("figures/" + f)
        bundle.seek(0)

        n_fig = sum(1 for x in included if x.startswith("figures/"))
        st.success(f"Done — {len(included)} files bundled "
                   f"({n_fig} figure{'s' if n_fig != 1 else ''}).")
        st.download_button("⬇ Download bundle (.zip)", bundle.getvalue(),
                           file_name=f"{paper}_bundle.zip",
                           mime="application/zip", type="primary")
        with open(out, "rb") as fh:
            st.download_button("…or just the Word file", fh.read(),
                               file_name=os.path.basename(out),
                               mime="application/vnd.openxmlformats-officedocument."
                                    "wordprocessingml.document")
        with st.expander("what's in the zip"):
            st.write("\n".join("• " + x for x in included))
        with st.expander("conversion log"):
            st.code("\n".join(log_lines))
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


# ── footer ─────────────────────────────────────────────────────────────────
st.markdown("---")
with st.expander("ℹ️  Editable is close, not pixel-identical"):
    st.markdown("""
Word re-wraps text with its own line-breaking engine (greedy first-fit) while
LaTeX uses Knuth–Plass total-fit, the "Times" font files differ slightly, and
Word has no LaTeX hyphenation or microtype — so **line and page breaks drift**
even with identical fonts and margins. That is a property of the two systems,
not a setting. If a reader needs it to *look* identical, hand them the compiled
Overleaf **PDF**; this tool is for the **editable** hand-off.
""")
st.caption("OWL · built on pandoc + python-docx · no LaTeX install required.")
st.caption("© 2026 David G. Schauer · All rights reserved. OWL — the app, workflow "
           "and code — is an original work of the author. It gratefully builds on the "
           "open-source [pandoc](https://pandoc.org) and "
           "[python-docx](https://python-docx.readthedocs.io) projects — "
           "thanks to their developers. No reuse or redistribution without "
           "written permission.")
