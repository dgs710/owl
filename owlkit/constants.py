"""Shared markers and asset paths for OWL.

The PAGEBREAK / APPENDIX / REFERENCES marks are plain-text sentinels that
survive the pandoc round-trip untouched; the docx postprocessor turns them
back into real Word section furniture.

The XREF_* delimiters are Unicode private-use characters.  They can never
occur in real prose, and pandoc passes them through verbatim, so they are a
safe way to smuggle "this run is a cross-reference" through the conversion.
"""

import os

PAGEBREAK_MARK = "TEX2DOCXPAGEBREAKMARK"
APPENDIX_MARK = "TEX2DOCXAPPENDIXMARK"
REFERENCES_MARK = "TEX2DOCXREFERENCESMARK"

# <A> anchor <S> display <E>
XREF_A = "\ue000"
XREF_S = "\ue001"
XREF_E = "\ue002"

_PKG = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.normpath(os.path.join(_PKG, os.pardir, "assets"))

REFERENCE_DOCX = os.path.join(ASSETS, "reference.docx")
CSL_PATH = os.path.join(ASSETS, "american-chemical-society.csl")
CSL_URL = ("https://raw.githubusercontent.com/citation-style-language/"
           "styles/master/american-chemical-society.csl")
