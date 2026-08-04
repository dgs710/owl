"""
OWL — Overleaf · Word · LaTeX
A Streamlit front-end for tex2docx.py: upload your Overleaf sources, pick a
mode, download an editable Word (.docx). Deployable free on Streamlit
Community Cloud or Hugging Face Spaces (CPU). No GPU, no LaTeX install needed.
"""

import os
import io
import re
import sys
import shutil
import zipfile
import tempfile
import subprocess

import streamlit as st

HERE = os.path.dirname(os.path.abspath(__file__))
TEX2DOCX = os.path.join(HERE, "tex2docx.py")

# ── shared access password ────────────────────────────────────────────────
# Set OWL_PASSWORD in Streamlit "Secrets" (recommended). Falls back to the
# default below if no secret is configured.
def get_password():
    try:
        return st.secrets["OWL_PASSWORD"]
    except Exception:
        return "xX2357Xx"

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
  .owl-card { border:1px solid rgba(52,211,153,.28); border-radius:14px;
    background:rgba(255,255,255,.03); padding:1rem 1.2rem; margin:.4rem 0; }
  .owl-note { font-size:.82rem; color:#B7C2BC; line-height:1.7; }
  code { color:#6EE7B7; }
</style>
""", unsafe_allow_html=True)

st.markdown('<p class="owl-title">🦉 OWL</p>', unsafe_allow_html=True)
st.markdown('<p class="owl-sub">Overleaf · Word · LaTeX — one upload, a Word bundle back</p>',
            unsafe_allow_html=True)
st.write("")

# ── password gate ──────────────────────────────────────────────────────────
if "owl_ok" not in st.session_state:
    st.session_state.owl_ok = False

if not st.session_state.owl_ok:
    st.markdown("#### 🔒 Access")
    st.caption("This converter is password protected.")
    pw = st.text_input("Access password", type="password",
                       label_visibility="collapsed", placeholder="Enter access password")
    if st.button("Unlock"):
        if pw == get_password():
            st.session_state.owl_ok = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.caption("© 2026 David G. Schauer · All rights reserved. "
               "OWL is an original tool by the author.")
    st.stop()

# ── helpers ────────────────────────────────────────────────────────────────
def find_main_tex(root):
    """Pick the .tex that contains \\documentclass (the main file); fall back
    to \\begin{document}, then to the only/first .tex."""
    texs = []
    for dp, _, files in os.walk(root):
        for f in files:
            if f.lower().endswith(".tex"):
                texs.append(os.path.join(dp, f))
    if not texs:
        return None
    for t in texs:
        try:
            s = open(t, encoding="utf-8", errors="ignore").read()
            if "\\documentclass" in s:
                return t
        except Exception:
            pass
    for t in texs:
        try:
            if "\\begin{document}" in open(t, encoding="utf-8", errors="ignore").read():
                return t
        except Exception:
            pass
    return texs[0]

def run_convert(workdir, main_tex, zotero):
    cmd = [sys.executable, TEX2DOCX, main_tex]
    if zotero:
        cmd.append("--zotero")
    proc = subprocess.run(cmd, capture_output=True, text=True, cwd=workdir)
    return proc

# ── main UI ────────────────────────────────────────────────────────────────
st.markdown("### 1 · Upload your Overleaf sources")
st.markdown(
    '<p class="owl-note">Easiest: in Overleaf use <b>Menu → Download → Source</b> '
    'and upload that <code>.zip</code> here (it already contains your '
    '<code>.tex</code>, <code>references.bib</code>, and figures). '
    'Or upload the individual files below.</p>', unsafe_allow_html=True)

up_zip = st.file_uploader("Overleaf source .zip", type=["zip"])
with st.expander("…or upload individual files instead"):
    up_tex = st.file_uploader("main.tex", type=["tex"])
    up_bib = st.file_uploader("references.bib", type=["bib"])
    up_figs = st.file_uploader("figures (png / jpg / pdf / eps)",
                               accept_multiple_files=True)

st.markdown(
    '<p class="owl-note">Optional: also drop in your <b>compiled PDF</b> '
    '(Overleaf → Download → PDF) and it gets added to the download bundle. '
    'The Overleaf <em>source</em> zip doesn\'t contain the PDF, so upload it '
    'here if you want it included.</p>', unsafe_allow_html=True)
up_pdf = st.file_uploader("Compiled PDF (optional)", type=["pdf"])

st.markdown("### 2 · Convert")
st.markdown(
    '<p class="owl-note">Produces an <b>editable Word</b> document — Times New '
    'Roman, justified, title page, running header, real Word equations, and '
    'clickable ACS citations built from your <code>.bib</code>. You get back a '
    '<b>zip</b> with the Word file, your compiled PDF (if provided), the '
    '<code>.bib</code>, and all figures.</p>', unsafe_allow_html=True)
zotero = False

go = st.button("Convert to Word →", type="primary")

if go:
    workdir = tempfile.mkdtemp(prefix="owl_")
    try:
        # materialise the uploads into one flat working directory
        if up_zip is not None:
            with zipfile.ZipFile(io.BytesIO(up_zip.read())) as z:
                z.extractall(workdir)
        elif up_tex is not None:
            open(os.path.join(workdir, "main.tex"), "wb").write(up_tex.getbuffer())
            if up_bib is not None:
                open(os.path.join(workdir, up_bib.name), "wb").write(up_bib.getbuffer())
            for f in (up_figs or []):
                open(os.path.join(workdir, f.name), "wb").write(f.getbuffer())
        else:
            st.warning("Please upload your Overleaf .zip (or at least a .tex file).")
            st.stop()

        main_tex = find_main_tex(workdir)
        if not main_tex:
            st.error("No .tex file found in the upload.")
            st.stop()

        # optional compiled PDF -> save next to the sources for bundling
        if up_pdf is not None:
            pdf_name = os.path.splitext(os.path.basename(main_tex))[0] + ".pdf"
            open(os.path.join(os.path.dirname(main_tex), pdf_name), "wb").write(up_pdf.getbuffer())

        with st.spinner("Converting with pandoc…"):
            proc = run_convert(os.path.dirname(main_tex), main_tex, zotero)

        # locate the produced .docx
        base = os.path.splitext(main_tex)[0]
        out = base + ".docx"
        if proc.returncode != 0 or not os.path.exists(out):
            st.error("Conversion failed. Details below.")
            st.code((proc.stderr or proc.stdout or "no output")[-3000:])
            st.stop()

        # ── assemble the download bundle: Word + PDF + .bib + figures ──
        srcdir = os.path.dirname(main_tex)
        paper = os.path.splitext(os.path.basename(main_tex))[0]
        IMG_EXT = (".png", ".jpg", ".jpeg", ".gif", ".eps", ".pdf",
                   ".svg", ".tif", ".tiff")
        bundle = io.BytesIO()
        included = []
        with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as z:
            # the converted Word file
            z.write(out, os.path.basename(out))
            included.append(os.path.basename(out))
            # the compiled PDF, if the user supplied one
            pdf_path = os.path.join(srcdir, paper + ".pdf")
            if up_pdf is not None and os.path.exists(pdf_path):
                z.write(pdf_path, paper + ".pdf")
                included.append(paper + ".pdf")
            # the bibliography
            for f in os.listdir(srcdir):
                if f.lower().endswith(".bib"):
                    z.write(os.path.join(srcdir, f), f)
                    included.append(f)
            # every figure / image, under figures/
            for dp, _, files in os.walk(srcdir):
                for f in files:
                    low = f.lower()
                    if not low.endswith(IMG_EXT):
                        continue
                    if low.endswith("_conv.png"):        # skip our auto-conversions
                        continue
                    full = os.path.join(dp, f)
                    if os.path.abspath(full) == os.path.abspath(pdf_path):
                        continue                          # already added as the paper PDF
                    rel = os.path.relpath(full, srcdir)
                    z.write(full, os.path.join("figures", os.path.basename(rel)))
                    included.append("figures/" + os.path.basename(rel))
        bundle.seek(0)

        n_fig = sum(1 for x in included if x.startswith("figures/"))
        st.success(f"Done — bundled {len(included)} files "
                   f"({n_fig} figure{'s' if n_fig != 1 else ''}).")
        st.download_button("⬇ Download bundle (.zip)", bundle.getvalue(),
                           file_name=f"{paper}_bundle.zip", mime="application/zip",
                           type="primary")
        with st.expander("what's in the zip"):
            st.write("\n".join("• " + x for x in included))
            if up_pdf is None:
                st.info("No compiled PDF included — upload one above if you want it "
                        "in the bundle.")
        # also offer the Word file on its own
        st.download_button("…or just the Word file", open(out, "rb").read(),
                           file_name=os.path.basename(out),
                           mime="application/vnd.openxmlformats-officedocument."
                                "wordprocessingml.document")
        if proc.stdout.strip():
            with st.expander("conversion log"):
                st.code(proc.stdout)
    finally:
        shutil.rmtree(workdir, ignore_errors=True)

# ── footer notes ───────────────────────────────────────────────────────────
st.markdown("---")
with st.expander("ℹ️  Good to know — editable is close, not pixel-identical"):
    st.markdown("""
Word re-wraps text with its own line-breaking engine (greedy first-fit) while
LaTeX uses Knuth–Plass total-fit, the “Times” font files differ slightly, and
Word lacks LaTeX hyphenation/microtype — so **line and page breaks will drift**
even with identical fonts and margins. That's a property of the two systems, not
a setting. If a reader needs it to *look* pixel-for-pixel identical, just hand
them the compiled Overleaf **PDF** — this tool is for the **editable** hand-off.
""")
st.caption("OWL · built on pandoc + python-docx · no LaTeX install required.")
st.caption("© 2026 David G. Schauer · All rights reserved. OWL — the app, workflow "
           "and code — is an original work of the author. It gratefully builds on the "
           "open-source [pandoc](https://pandoc.org) and "
           "[python-docx](https://python-docx.readthedocs.io) projects — "
           "thanks to their developers. No reuse or redistribution without "
           "written permission.")
