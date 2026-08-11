#!/usr/bin/env python3
"""Command-line front end for OWL.

    python tex2docx.py main.tex
    python tex2docx.py main.tex -o Schauer_QE.docx
    python tex2docx.py main.tex --aux main.aux --force

Kept at the repository root under its original name so existing habits and
scripts keep working; the conversion itself lives in `owlkit`.
"""

import argparse
import sys

from owlkit import convert, ConversionError


def main():
    ap = argparse.ArgumentParser(description="Overleaf LaTeX -> editable Word")
    ap.add_argument("tex", help="the .tex file (e.g. main.tex)")
    ap.add_argument("-o", "--output", help="output .docx (default: same name)")
    ap.add_argument("--keep-drafts", action="store_true",
                    help="keep \\draftnote{...} notes instead of removing them")
    ap.add_argument("--aux", help="main.aux from a LaTeX run; its label numbers "
                                  "override OWL's own reconstruction")
    ap.add_argument("--force", action="store_true",
                    help="convert even if preflight finds problems")
    args = ap.parse_args()

    try:
        result = convert(args.tex, output=args.output,
                         keep_drafts=args.keep_drafts,
                         strict=not args.force, aux_path=args.aux,
                         on_log=lambda m: print("  " + m))
    except ConversionError as e:
        print(f"\nERROR: {e}", file=sys.stderr)
        for issue in e.issues:
            print("  " + issue.as_text(), file=sys.stderr)
        if e.detail:
            print("\n" + e.detail, file=sys.stderr)
        sys.exit(1)
    print(f"OK  ->  {result.docx_path}")


if __name__ == "__main__":
    main()
