"""OWL — Overleaf · Word · LaTeX.

Public surface:

    from owlkit import convert, preflight
    result = convert("main.tex")

The pieces underneath, in the order the pipeline uses them:

    counters   the LaTeX counter machine: what number does each \\label print
    preflight  validate the source and report line numbers, before pandoc runs
    engine     preprocess -> pandoc -> docx postprocess (captions, cross-refs)
    floats     where the compiled PDF actually placed each figure and table
    pagefit    match the Word page boundaries to the compiled PDF
"""

from .convert import convert, ConversionError, Result   # noqa: F401
from . import preflight                                  # noqa: F401

__all__ = ["convert", "preflight", "ConversionError", "Result"]
