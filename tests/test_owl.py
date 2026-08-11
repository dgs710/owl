"""Regression tests for the bugs that actually bit.

Run with:  python -m pytest tests/ -q     (or just: python tests/test_owl.py)
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from owlkit import preflight                                    # noqa: E402
from owlkit.counters import (scan_preamble, build_label_map,     # noqa: E402
                             float_numbers, strip_comments)
from owlkit.engine import make_resolver                          # noqa: E402


PREAMBLE = r"""
\documentclass{article}
\usepackage{graphicx}
\graphicspath{{Figures/}{./}}
\newcounter{appx}
\renewcommand{\theappx}{\Alph{appx}}
\newcommand{\appsection}[1]{\refstepcounter{appx}\section*{Appendix \theappx: #1}}
\crefname{appx}{Appendix}{Appendices}
\begin{document}
"""


def test_appendix_cross_references_resolve():
    """A user macro that steps its own counter used to fall through the label
    map, so \\cref{app:x} printed the raw label text in the Word file."""
    src = PREAMBLE + r"""
\section{Intro}\label{sec:intro}
\appsection{First}\label{app:one}
\appsection{Second}\label{app:two}
\end{document}"""
    labels = build_label_map(src)
    assert labels["sec:intro"] == ("Section", "1")
    assert labels["app:one"] == ("Appendix", "A")
    assert labels["app:two"] == ("Appendix", "B")


def test_commented_out_macros_are_ignored():
    """A \\appsection mentioned in a comment must not step the counter."""
    src = PREAMBLE + r"""
% every appendix uses \appsection{...} like this
\appsection{Real}\label{app:real}
\end{document}"""
    assert build_label_map(src)["app:real"] == ("Appendix", "A")


def test_supplementary_figure_numbering():
    """\\setcounter + \\renewcommand{\\thefigure} must give S1, S2 — not 3, 4."""
    src = PREAMBLE + r"""
\begin{figure}\label{fig:a}\end{figure}
\begin{figure}\label{fig:b}\end{figure}
\setcounter{figure}{0}
\renewcommand{\thefigure}{S\arabic{figure}}
\begin{figure}\label{fig:si}\end{figure}
\end{document}"""
    figs, _ = float_numbers(src)
    assert figs == ["1", "2", "S1"]
    assert build_label_map(src)["fig:si"] == ("Figure", "S1")


def test_equation_labels():
    src = PREAMBLE + r"""
\begin{equation}\label{eq:one}E=mc^2\end{equation}
\end{document}"""
    assert build_label_map(src)["eq:one"] == ("Equation", "1")


def test_stray_brace_is_reported_with_a_line_number():
    """The failure that started all this: one unmatched { produced a pandoc
    error pointing at \\end{document}, hundreds of lines away."""
    src = PREAMBLE + "\nGood paragraph here.\n\nA broken {paragraph here.\n\n" \
                     "\\end{document}"
    errors, _ = preflight.run(src)
    brace = [e for e in errors if "brace" in e.message]
    assert brace, "stray brace not detected"
    assert brace[0].line == len(PREAMBLE.split("\n")) + 3


def test_missing_bib_is_an_error_not_a_silent_drop():
    src = PREAMBLE + r"""
Some text \cite{smith2020}.
\end{document}"""
    errors, _ = preflight.run(src, bib_path=None)
    assert any("citation" in e.message for e in errors)
    errors, _ = preflight.run(src, bib_path="references.bib")
    assert not any("citation" in e.message for e in errors)


def test_unbalanced_environment_is_reported():
    src = PREAMBLE + "\n\\begin{figure}\nno end here\n\\end{document}"
    errors, _ = preflight.run(src)
    assert any("figure" in e.message for e in errors)


def test_strip_comments_keeps_escaped_percent():
    assert strip_comments(r"50\% yield % a note").strip() == r"50\% yield"


def test_figure_resolver_handles_graphicspath_and_extensions(tmpdir=None):
    """\\includegraphics{Fig1} in a document with \\graphicspath{{Figures/}}
    must find Figures/Fig1.pdf — the old code only matched literal '.pdf'
    names in the main directory, so it silently converted nothing."""
    import tempfile
    root = tempfile.mkdtemp()
    os.makedirs(os.path.join(root, "Figures"))
    target = os.path.join(root, "Figures", "Fig1.pdf")
    open(target, "wb").write(b"%PDF-1.4\n")
    resolve = make_resolver(PREAMBLE, root)
    assert resolve("Fig1") == target
    assert resolve("Fig1.pdf") == target
    assert resolve("Nope") is None


if __name__ == "__main__":
    failed = 0
    for name, fn in sorted(globals().items()):
        if not name.startswith("test_") or not callable(fn):
            continue
        try:
            fn()
            print(f"  PASS  {name}")
        except AssertionError as exc:
            failed += 1
            print(f"  FAIL  {name}: {exc}")
    print(f"\n{'all tests passed' if not failed else f'{failed} test(s) failed'}")
    sys.exit(1 if failed else 0)
