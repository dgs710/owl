# 🦉 OWL — Overleaf · Word · LaTeX

A tiny web app that turns your **Overleaf** sources (LaTeX + `references.bib` +
figures) into an **editable Word `.docx`** — Times New Roman, justified, title
page, running header, real Word equations, and clickable ACS citations built
from your `.bib`. Optional **Zotero** mode emits live-citation markers.

It runs the proven `tex2docx.py` engine behind a Streamlit upload → convert →
download page. **No GPU and no LaTeX install** — it uses pandoc, not a LaTeX
compile.

---

## Run it locally

```bash
pip install -r requirements.txt
# system tools (once): pandoc + poppler (+ libreoffice optional)
#   macOS:   brew install pandoc poppler
#   Ubuntu:  sudo apt-get install -y pandoc poppler-utils libreoffice-writer
streamlit run app.py
```

Open the URL it prints, enter the password (default `xX2357Xx`), upload your
Overleaf **source `.zip`**, pick a mode, download the `.docx`.

---

## Deploy free on Streamlit Community Cloud  (recommended)

1. Put this whole folder in a **GitHub repo** (e.g. `owl`). Push it (GitHub
   Desktop is fine — same as your website).
2. Go to **share.streamlit.io** → sign in with GitHub → **Create app** →
   pick your repo, branch `main`, main file `app.py`.
3. Under **Advanced → Secrets**, add your password so it isn't in the code:
   ```toml
   OWL_PASSWORD = "your-password-here"
   ```
4. Click **Deploy**. First build takes a few minutes (it installs the tools in
   `packages.txt`). You get a URL like `https://owl-schauer.streamlit.app`.
5. Every time you push to GitHub, it redeploys automatically.

`packages.txt` (apt) and `requirements.txt` (pip) are read automatically — you
don't install anything by hand on the server.

---

## Deploy free on Hugging Face Spaces  (alternative, CPU tier)

1. Create a **Space** → SDK **Streamlit** → free CPU hardware.
2. Upload these files (or connect the GitHub repo).
3. Add `OWL_PASSWORD` under the Space's **Settings → Secrets**.
4. `packages.txt` + `requirements.txt` are honored the same way. It builds and
   serves at `https://huggingface.co/spaces/<you>/owl`.

---

## Wiring it into the website

The website's **OWL** tab (password-gated) has a **Launch OWL** button. Once the
app is deployed, set that button's link to your `…streamlit.app` (or HF Spaces)
URL and it opens the converter. The app has its own password gate too, so the
URL isn't wide open.

---

## Files

| file | purpose |
|------|---------|
| `app.py` | the Streamlit web app (gate, upload, mode, convert, download) |
| `tex2docx.py` | the conversion engine (editable + `--zotero`) |
| `reference.docx` | Word style template (Times New Roman, 12 pt, single, 1 in) |
| `american-chemical-society.csl` | ACS citation style |
| `requirements.txt` | Python deps (pip) |
| `packages.txt` | system deps (apt) — pandoc, poppler, libreoffice |
| `.streamlit/config.toml` | dark-emerald theme + upload size |

## Note on fidelity

Editable output is **close, not pixel-identical** — Word and LaTeX break lines
differently, so page breaks drift. For a look-identical copy, hand over the
compiled Overleaf **PDF**. This tool is the *editable* hand-off.
