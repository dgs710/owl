"""Check the LaTeX source *before* handing it to pandoc.

The failure that motivated this module: a single stray `{` on line 337 of a
600-line document.  Pandoc's reader stayed inside the group to end of file and
died with

    Error at "....pre.tex" (line 686, column 1): unexpected \\end

which the app relayed verbatim.  Nothing in that message points at line 337,
or at a brace, or at anything the author could act on.

Everything here reports a *source line number* and a fix.  Errors block the
conversion; warnings let it through but are shown to the user.
"""

import os
import re

from .counters import strip_comments


class Issue:
    """One preflight finding."""

    def __init__(self, level, line, message, hint=None):
        self.level = level              # "error" | "warning"
        self.line = line                # 1-based source line, or None
        self.message = message
        self.hint = hint

    def __repr__(self):
        where = f"line {self.line}: " if self.line else ""
        return f"[{self.level}] {where}{self.message}"

    def as_text(self):
        where = f"Line {self.line} — " if self.line else ""
        out = f"{where}{self.message}"
        if self.hint:
            out += f"\n    → {self.hint}"
        return out


def _line_of(src, index):
    return src.count("\n", 0, index) + 1


# ---------------------------------------------------------------------------
# individual checks
# ---------------------------------------------------------------------------

def check_braces(src):
    """Find unbalanced { } — the single most common hard failure.

    Reports the *paragraph* where the imbalance opens, which is what the author
    needs, rather than the end of the file, which is where pandoc notices.
    """
    clean = strip_comments(src)
    issues = []
    depth = 0
    para_start_line = 1
    open_positions = []
    line = 1
    i, n = 0, len(clean)
    para_blank = True
    while i < n:
        ch = clean[i]
        if ch == "\n":
            line += 1
            # a blank line ends a paragraph
            if clean[i:i + 2] == "\n\n":
                if depth > 0 and open_positions:
                    issues.append(Issue(
                        "error", _line_of(clean, open_positions[-1]),
                        f"unmatched opening brace {{ — {depth} brace(s) left open "
                        f"at the end of this paragraph",
                        "Usually a half-deleted command such as "
                        "\\textcolor{red}{...}. Delete the stray { or restore "
                        "the command."))
                    depth = 0
                    open_positions = []
                para_blank = True
                para_start_line = line
        elif ch == "\\":
            i += 2                      # \{ and \} are literal characters
            continue
        elif ch == "{":
            depth += 1
            open_positions.append(i)
        elif ch == "}":
            depth -= 1
            if open_positions:
                open_positions.pop()
            if depth < 0:
                issues.append(Issue(
                    "error", line,
                    "unmatched closing brace }",
                    "There is one more } than { at this point."))
                depth = 0
        i += 1
    if depth > 0 and open_positions:
        issues.append(Issue(
            "error", _line_of(clean, open_positions[-1]),
            f"unmatched opening brace {{ — {depth} brace(s) never closed",
            "Delete the stray { or close it."))
    return issues


def check_environments(src):
    """\\begin{x} without a matching \\end{x}, and vice versa."""
    clean = strip_comments(src)
    issues = []
    stack = []
    for m in re.finditer(r"\\(begin|end)\s*\{([^}]+)\}", clean):
        kind, env = m.group(1), m.group(2)
        if kind == "begin":
            stack.append((env, _line_of(clean, m.start())))
        else:
            if not stack:
                issues.append(Issue("error", _line_of(clean, m.start()),
                                    f"\\end{{{env}}} with no matching \\begin"))
            elif stack[-1][0] != env:
                open_env, open_line = stack[-1]
                issues.append(Issue(
                    "error", _line_of(clean, m.start()),
                    f"\\end{{{env}}} closes \\begin{{{open_env}}} "
                    f"(opened on line {open_line})"))
                stack.pop()
            else:
                stack.pop()
    for env, ln in stack:
        if env == "document":
            continue
        issues.append(Issue("error", ln, f"\\begin{{{env}}} is never closed"))
    return issues


def check_bibliography(src, bib_path):
    """A .tex that cites but has no .bib converts *successfully* and silently
    loses every citation and the whole reference list.  That is worse than
    failing, so it is an error unless the document genuinely has no citations.
    """
    clean = strip_comments(src)
    cites = re.findall(r"\\(?:cite|parencite|autocite|citep|citet|textcite)"
                       r"\s*(?:\[[^\]]*\])?\s*\{([^}]*)\}", clean)
    n_cites = sum(len([k for k in c.split(",") if k.strip()]) for c in cites)
    wants_bib = bool(re.search(r"\\(?:addbibresource|bibliography)\s*\{", clean))
    if (n_cites or wants_bib) and not bib_path:
        return [Issue(
            "error", None,
            f"the document has {n_cites} citation(s) but no .bib file was found",
            "Upload references.bib (Overleaf → Menu → Download → Source "
            "includes it). Without it every citation and the whole References "
            "section are dropped silently.")]
    return []


def check_figures(src, resolver):
    """Every \\includegraphics must resolve to a file on disk."""
    clean = strip_comments(src)
    issues = []
    for m in re.finditer(r"\\includegraphics(?:\s*\[[^\]]*\])?\s*\{([^}]*)\}",
                         clean):
        name = m.group(1).strip()
        if not resolver(name):
            issues.append(Issue(
                "warning", _line_of(clean, m.start()),
                f"figure '{name}' not found",
                "It will appear as a placeholder box. Check the filename or "
                "include the figure in the upload."))
    return issues


def check_math_markup(src):
    """\\textcolor inside $...$ is not valid math; pandoc falls back to
    printing the raw TeX, which looks like a bug in the Word file."""
    clean = strip_comments(src)
    issues = []
    for m in re.finditer(r"\$([^$]{1,400})\$", clean):
        inner = m.group(1)
        for cmd in ("textcolor", "textbf", "textit", "emph"):
            if "\\" + cmd in inner:
                issues.append(Issue(
                    "warning", _line_of(clean, m.start()),
                    f"\\{cmd} inside math ($...$) — pandoc cannot convert it",
                    f"Move the \\{cmd}{{...}} outside the $ $ delimiters."))
                break
    return issues


def check_labels(src, labels):
    """Cross-references OWL could not resolve would print as raw label text
    ('app:protocols') in the Word file."""
    clean = strip_comments(src)
    issues = []
    seen = set()
    for m in re.finditer(r"\\(?:c|C)?ref\s*\{([^}]*)\}|\\autoref\s*\{([^}]*)\}",
                         clean):
        keys = (m.group(1) or m.group(2) or "")
        for key in [k.strip() for k in keys.split(",") if k.strip()]:
            if key in labels or key in seen:
                continue
            seen.add(key)
            issues.append(Issue(
                "warning", _line_of(clean, m.start()),
                f"cross-reference to unknown label '{key}'",
                "It will print as raw text. Check the \\label spelling."))
    return issues


def check_duplicate_labels(src):
    clean = strip_comments(src)
    issues = []
    first = {}
    for m in re.finditer(r"\\label\s*\{([^}]+)\}", clean):
        key, ln = m.group(1), _line_of(clean, m.start())
        if key in first:
            issues.append(Issue(
                "warning", ln,
                f"duplicate \\label{{{key}}} (also on line {first[key]})",
                "Cross-references to it will point at the first one."))
        else:
            first[key] = ln
    return issues


# ---------------------------------------------------------------------------
# runner
# ---------------------------------------------------------------------------

def run(src, bib_path=None, figure_resolver=None, labels=None):
    """Run every check.  Returns (errors, warnings)."""
    issues = []
    issues += check_braces(src)
    issues += check_environments(src)
    issues += check_bibliography(src, bib_path)
    if figure_resolver is not None:
        issues += check_figures(src, figure_resolver)
    issues += check_math_markup(src)
    issues += check_duplicate_labels(src)
    if labels is not None:
        issues += check_labels(src, labels)

    errors = [i for i in issues if i.level == "error"]
    warnings = [i for i in issues if i.level == "warning"]
    errors.sort(key=lambda i: (i.line is None, i.line or 0))
    warnings.sort(key=lambda i: (i.line is None, i.line or 0))
    return errors, warnings
