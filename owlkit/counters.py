"""A miniature LaTeX counter machine.

The old label map only understood \\section, \\subsection, figure and table.
Anything else -- \\begin{equation}, or a user macro like

    \\newcounter{appx}
    \\renewcommand{\\theappx}{\\Alph{appx}}
    \\newcommand{\\appsection}[1]{\\refstepcounter{appx}\\section*{Appendix \\theappx: #1}}
    \\crefname{appx}{Appendix}{Appendices}

-- fell straight through, so every \\cref{app:...} came out as raw label text
in the Word file.

Rather than special-casing macros, this module *reads the preamble* the way
LaTeX does: it learns which counters exist, how each one is formatted
(\\arabic / \\Alph / \\alph / \\roman / \\Roman), which user macros step which
counter (\\refstepcounter inside a \\newcommand body), and what cleveref
should call each counter (\\crefname).  Then it walks the body once and
assigns a printed number to every \\label.

If a label still cannot be resolved, it is reported rather than silently
emitted as raw text -- preflight turns that into a visible warning.
"""

import re

# ---------------------------------------------------------------------------
# number formatting
# ---------------------------------------------------------------------------

def _roman(n):
    if n <= 0:
        return str(n)
    vals = ((1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"),
            (90, "xc"), (50, "l"), (40, "xl"), (10, "x"), (9, "ix"),
            (5, "v"), (4, "iv"), (1, "i"))
    out = []
    for v, s in vals:
        while n >= v:
            out.append(s)
            n -= v
    return "".join(out)


def _alph(n):
    # LaTeX's \alph only defines 1..26
    return chr(ord("a") + n - 1) if 1 <= n <= 26 else str(n)


FORMATTERS = {
    "arabic": str,
    "Alph": lambda n: _alph(n).upper(),
    "alph": _alph,
    "roman": _roman,
    "Roman": lambda n: _roman(n).upper(),
}

# cleveref's own default names -- note that several are ABBREVIATED, which is
# easy to miss: \cref{fig:x} prints "fig. 3", not "figure 3", and an equation
# reference prints "eq. (1)" with the number in parentheses.  Emitting
# "Figure 3" where the LaTeX says "Fig. 3" is a real difference in the text.
DEFAULT_CREFNAMES = {
    "section": ("section", "sections"),
    "subsection": ("section", "sections"),
    "subsubsection": ("section", "sections"),
    "figure": ("fig.", "figs."),
    "table": ("table", "tables"),
    "equation": ("eq.", "eqs."),
}

# counters whose reference number is printed in parentheses
PARENTHESISED = {"equation"}


class CounterModel:
    """What the preamble told us about counters, macros and cref names."""

    def __init__(self):
        self.values = {}          # counter -> current value
        self.formats = {}         # counter -> list of (counter, formatter)
        self.crefnames = {}       # counter -> display name
        self.macro_steps = {}     # macro name -> counter it \refstepcounter's
        self.parents = {}         # counter -> parent counter (for "2.1" style)
        self.crefnames_plural = {}
        # \usepackage[capitalise]{cleveref} makes \cref capitalise like \Cref
        self.capitalise = False
        self.label_counter = {}   # label -> the counter it belongs to

    # -- formatting ------------------------------------------------------
    def render(self, counter):
        """Render a counter the way \\thecounter would."""
        spec = self.formats.get(counter)
        if not spec:
            return str(self.values.get(counter, 0))
        return "".join(
            FORMATTERS.get(fmt, str)(self.values.get(c, 0)) if fmt else fmt_lit
            for c, fmt, fmt_lit in spec)

    def step(self, counter):
        self.values[counter] = self.values.get(counter, 0) + 1
        # stepping a counter resets everything that lists it as a parent
        for child, parent in self.parents.items():
            if parent == counter:
                self.values[child] = 0

    def name_for(self, counter, plural=False):
        if plural:
            name = self.crefnames_plural.get(counter)
            if name:
                return name
            default = DEFAULT_CREFNAMES.get(counter)
            return default[1] if default else None
        name = self.crefnames.get(counter)
        if name:
            return name
        default = DEFAULT_CREFNAMES.get(counter)
        return default[0] if default else None


# ---------------------------------------------------------------------------
# preamble scanning
# ---------------------------------------------------------------------------

# \the<counter> definitions look like  \renewcommand{\theappx}{\Alph{appx}}
# or  \renewcommand\thesubsection{\thesection.\arabic{subsection}}
_THE_DEF = re.compile(
    r"\\(?:re)?newcommand\s*\*?\s*\{?\\the([a-zA-Z@]+)\}?\s*\{(.+?)\}\s*(?:%|$)",
    re.M)

_FMT_CALL = re.compile(r"\\(arabic|Alph|alph|roman|Roman)\s*\{([a-zA-Z@]+)\}"
                       r"|\\the([a-zA-Z@]+)")


def _parse_the_body(body, model, seen=None):
    """Turn the body of a \\the<counter> definition into a render spec:
    a list of (counter, formatter_name, literal) triples."""
    seen = seen or set()
    spec = []
    pos = 0
    for m in _FMT_CALL.finditer(body):
        lit = body[pos:m.start()]
        if lit.strip():
            spec.append((None, None, lit))
        if m.group(1):
            spec.append((m.group(2), m.group(1), None))
        else:                                   # nested \thesection
            inner = m.group(3)
            if inner in model.formats and inner not in seen:
                spec.extend(model.formats[inner])
            else:
                spec.append((inner, "arabic", None))
        pos = m.end()
    tail = body[pos:]
    if tail.strip():
        spec.append((None, None, tail))
    return spec


def scan_preamble(src):
    """Read counter declarations, \\the definitions, \\crefname and any user
    macro whose body steps a counter."""
    model = CounterModel()
    preamble = src.split("\\begin{document}")[0] if "\\begin{document}" in src else src

    # LaTeX's own counters
    for c in ("section", "subsection", "subsubsection", "figure", "table",
              "equation"):
        model.values[c] = 0
    model.parents["subsection"] = "section"
    model.parents["subsubsection"] = "subsection"
    model.formats["section"] = [("section", "arabic", None)]
    model.formats["subsection"] = [("section", "arabic", None), (None, None, "."),
                                   ("subsection", "arabic", None)]
    model.formats["subsubsection"] = list(model.formats["subsection"]) + \
        [(None, None, "."), ("subsubsection", "arabic", None)]
    model.formats["figure"] = [("figure", "arabic", None)]
    model.formats["table"] = [("table", "arabic", None)]
    model.formats["equation"] = [("equation", "arabic", None)]

    # \newcounter{appx}[parent]
    for m in re.finditer(r"\\newcounter\s*\{([a-zA-Z@]+)\}\s*(?:\[([a-zA-Z@]+)\])?",
                         preamble):
        model.values[m.group(1)] = 0
        model.formats.setdefault(m.group(1), [(m.group(1), "arabic", None)])
        if m.group(2):
            model.parents[m.group(1)] = m.group(2)

    # \renewcommand{\theappx}{\Alph{appx}}
    for m in _THE_DEF.finditer(preamble):
        model.formats[m.group(1)] = _parse_the_body(m.group(2), model)

    # \crefname{appx}{Appendix}{Appendices}
    for m in re.finditer(r"\\[Cc]refname\s*\{([a-zA-Z@]+)\}\s*\{([^}]*)\}"
                         r"\s*(?:\{([^}]*)\})?", preamble):
        name = m.group(2).strip()
        model.crefnames.setdefault(m.group(1), name)
        if name[:1].isupper():
            model.crefnames[m.group(1)] = name
        if m.group(3):
            plural = m.group(3).strip()
            model.crefnames_plural.setdefault(m.group(1), plural)
            if plural[:1].isupper():
                model.crefnames_plural[m.group(1)] = plural

    # cleveref's "capitalise" option makes \cref behave like \Cref
    for m in re.finditer(r"\\usepackage\s*\[([^\]]*)\]\s*\{cleveref\}",
                         preamble):
        if re.search(r"\bcapitali[sz]e\b", m.group(1)):
            model.capitalise = True

    # \newcommand{\appsection}[1]{ ... \refstepcounter{appx} ... }
    for m in re.finditer(r"\\(?:re|provide)?newcommand\s*\*?\s*\{?\\([a-zA-Z@]+)\}?"
                         r"\s*(?:\[\d+\])?\s*(?:\[[^\]]*\])?\s*\{", preamble):
        name = m.group(1)
        body, ok = _brace_body(preamble, m.end() - 1)
        if not ok:
            continue
        step = re.search(r"\\refstepcounter\s*\{([a-zA-Z@]+)\}", body)
        if step:
            model.macro_steps[name] = step.group(1)

    return model


def _brace_body(src, open_idx):
    """Return the brace-matched body starting at src[open_idx] == '{'."""
    if open_idx >= len(src) or src[open_idx] != "{":
        return "", False
    depth, j = 1, open_idx + 1
    while j < len(src) and depth:
        if src[j] == "{":
            depth += 1
        elif src[j] == "}":
            depth -= 1
        j += 1
    return src[open_idx + 1:j - 1], depth == 0


# ---------------------------------------------------------------------------
# body walk
# ---------------------------------------------------------------------------

def strip_comments(src):
    """Remove LaTeX comments, keeping escaped \\% and preserving line count so
    reported line numbers stay meaningful."""
    out = []
    for line in src.split("\n"):
        i, n = 0, len(line)
        cut = None
        while i < n:
            if line[i] == "\\":
                i += 2
                continue
            if line[i] == "%":
                cut = i
                break
            i += 1
        out.append(line if cut is None else line[:cut])
    return "\n".join(out)


def build_label_map(src, model=None):
    """Walk the body in document order and give every \\label a printed name.

    Returns {label: (display_name_or_None, number_or_None)} -- matching the
    shape the rest of OWL already expects, e.g. {"fig:one": ("Figure", "3")}.
    """
    model = model or scan_preamble(src)
    body = strip_comments(src).split("\\begin{document}", 1)[-1]

    numbered_envs = ("figure", "table", "equation", "align", "figure*", "table*")
    user_macros = "|".join(re.escape(k) for k in sorted(model.macro_steps,
                                                        key=len, reverse=True))
    parts = [r"\\(section|subsection|subsubsection)(\*?)\s*\{",
             r"\\begin\{(" + "|".join(re.escape(e) for e in numbered_envs) + r")\}",
             r"\\label\s*\{([^}]+)\}",
             # mid-document counter surgery: the SI often does
             #   \setcounter{figure}{0}\renewcommand{\thefigure}{S\arabic{figure}}
             r"\\(setcounter|addtocounter|stepcounter)\s*\{([a-zA-Z@]+)\}"
             r"(?:\s*\{(-?\d+)\})?",
             r"\\(?:re)?newcommand\s*\*?\s*\{?\\the([a-zA-Z@]+)\}?\s*\{"]
    if user_macros:
        parts.append(r"\\(" + user_macros + r")\b")
    pattern = re.compile("|".join(parts))

    labels = {}
    pending = None
    for m in pattern.finditer(body):
        if m.group(1):                                   # sectioning
            counter = m.group(1)
            if m.group(2) == "*":                        # starred: unnumbered
                pending = (model.name_for(counter), None, counter)
            else:
                model.step(counter)
                pending = (model.name_for(counter), model.render(counter), counter)
        elif m.group(3):                                 # numbered environment
            env = m.group(3).rstrip("*")
            counter = "equation" if env in ("align", "equation") else env
            model.step(counter)
            pending = (model.name_for(counter), model.render(counter), counter)
        elif m.group(4) is not None:                     # \label{...}
            if pending is not None:
                labels[m.group(4)] = pending[:2]
                model.label_counter[m.group(4)] = pending[2]
                pending = None
        elif m.group(5):                                 # \setcounter etc.
            op, counter, val = m.group(5), m.group(6), m.group(7)
            if op == "stepcounter":
                model.step(counter)
            elif val is not None:
                cur = model.values.get(counter, 0)
                model.values[counter] = (int(val) if op == "setcounter"
                                         else cur + int(val))
        elif m.group(8):                                 # \renewcommand{\thefig}
            counter = m.group(8)
            body_txt, ok = _brace_body(body, m.end() - 1)
            if ok:
                model.formats[counter] = _parse_the_body(body_txt, model)
        elif m.lastindex and m.group(m.lastindex):       # user macro
            counter = model.macro_steps[m.group(m.lastindex)]
            model.step(counter)
            pending = (model.name_for(counter) or counter.title(),
                       model.render(counter), counter)
    return labels


def macro_bodies(src):
    """{macro name: (arg_count, body)} for every \\newcommand in the preamble."""
    preamble = strip_comments(src).split("\\begin{document}")[0]
    out = {}
    for m in re.finditer(r"\\(?:re|provide)?newcommand\s*\*?\s*\{?\\([a-zA-Z@]+)\}?"
                         r"\s*(?:\[(\d+)\])?\s*(?:\[[^\]]*\])?\s*\{", preamble):
        body, ok = _brace_body(preamble, m.end() - 1)
        if ok:
            out[m.group(1)] = (int(m.group(2) or 0), body)
    return out


def expand_counter_macros(src, model=None):
    """Expand user macros that step a counter, substituting the counter value.

    pandoc does not run TeX, so a macro like

        \\newcommand{\\appsection}[1]{\\refstepcounter{appx}%
                                    \\section*{Appendix \\theappx: #1}}

    reaches the Word file with `\\theappx` evaluating to nothing -- the heading
    comes out as "Appendix : Title", missing its letter.  Expanding the macro
    here, with the counter simulated, restores it.
    """
    model = model or scan_preamble(src)
    if not model.macro_steps:
        return src
    bodies = macro_bodies(src)
    head, sep, body = src.partition("\\begin{document}")
    if not sep:
        head, sep, body = "", "", src

    names = sorted(model.macro_steps, key=len, reverse=True)
    pattern = re.compile(r"\\(" + "|".join(re.escape(n) for n in names) + r")\b")

    def in_comment(text, index):
        """True if `index` sits after an unescaped % on its own line."""
        start = text.rfind("\n", 0, index) + 1
        i = start
        while i < index:
            if text[i] == "\\":
                i += 2
                continue
            if text[i] == "%":
                return True
            i += 1
        return False

    out, pos = [], 0
    while True:
        m = pattern.search(body, pos)
        if not m:
            out.append(body[pos:])
            break
        if in_comment(body, m.start()):
            out.append(body[pos:m.end()])
            pos = m.end()
            continue
        name = m.group(1)
        nargs, macro_body = bodies.get(name, (0, None))
        if macro_body is None:
            out.append(body[pos:m.end()])
            pos = m.end()
            continue

        j = m.end()
        args = []
        for _ in range(nargs):
            while j < len(body) and body[j] in " \t\n":
                j += 1
            if j >= len(body) or body[j] != "{":
                break
            arg, ok = _brace_body(body, j)
            if not ok:
                break
            args.append(arg)
            j += len(arg) + 2
        if len(args) != nargs:
            out.append(body[pos:m.end()])
            pos = m.end()
            continue

        counter = model.macro_steps[name]
        model.step(counter)
        expanded = macro_body
        expanded = re.sub(r"\\refstepcounter\s*\{[a-zA-Z@]+\}", "", expanded)
        expanded = re.sub(r"\\stepcounter\s*\{[a-zA-Z@]+\}", "", expanded)
        expanded = re.sub(r"\\phantomsection\b", "", expanded)
        expanded = re.sub(r"\\addcontentsline\s*\{[^}]*\}\s*\{[^}]*\}\s*\{",
                          lambda mm: "\\@gobblearg{", expanded)
        expanded = expanded.replace("\\the" + counter, model.render(counter))
        for k, arg in enumerate(args, start=1):
            expanded = expanded.replace(f"#{k}", arg)
        # drop the \addcontentsline argument we neutralised above
        while "\\@gobblearg{" in expanded:
            i = expanded.index("\\@gobblearg{")
            inner, ok = _brace_body(expanded, i + len("\\@gobblearg"))
            if not ok:
                break
            expanded = expanded[:i] + expanded[i + len("\\@gobblearg") + len(inner) + 2:]

        out.append(body[pos:m.start()])
        out.append(expanded)
        pos = j

    return head + sep + "".join(out)


def float_numbers(src, model=None):
    """The printed number of every figure and table, in document order.

    Captions used to be numbered by position ("the 6th figure is Figure 6"),
    which is wrong the moment a document restarts or reformats the counter --
    a supplementary section doing

        \\setcounter{figure}{0}
        \\renewcommand{\\thefigure}{S\\arabic{figure}}

    makes its first figure "Figure S1", not "Figure 6".  Running the same
    counter machine the label map uses keeps captions and cross-references
    telling the same story.

    Returns (figure_numbers, table_numbers) as lists of strings.
    """
    model = model or scan_preamble(src)
    body = strip_comments(src).split("\\begin{document}", 1)[-1]
    pattern = re.compile(
        r"\\begin\{(figure|table)\}\*?|"
        r"\\(setcounter|addtocounter|stepcounter)\s*\{([a-zA-Z@]+)\}"
        r"(?:\s*\{(-?\d+)\})?|"
        r"\\(?:re)?newcommand\s*\*?\s*\{?\\the([a-zA-Z@]+)\}?\s*\{")
    figs, tabs = [], []
    for m in pattern.finditer(body):
        if m.group(1):
            counter = m.group(1)
            model.step(counter)
            (figs if counter == "figure" else tabs).append(model.render(counter))
        elif m.group(2):
            op, counter, val = m.group(2), m.group(3), m.group(4)
            if op == "stepcounter":
                model.step(counter)
            elif val is not None:
                cur = model.values.get(counter, 0)
                model.values[counter] = (int(val) if op == "setcounter"
                                         else cur + int(val))
        elif m.group(5):
            counter = m.group(5)
            body_txt, ok = _brace_body(body, m.end() - 1)
            if ok:
                model.formats[counter] = _parse_the_body(body_txt, model)
    return figs, tabs


# ---------------------------------------------------------------------------
# authoritative override: LaTeX's own .aux
# ---------------------------------------------------------------------------

_AUX_CREF = re.compile(
    r"\\newlabel\{(?P<label>[^}]+)@cref\}\{\{\[(?P<type>[^\]]*)\]"
    r"\[[^\]]*\]\[\](?P<num>[^}]*)\}")


def labels_from_aux(aux_text, model=None):
    """LaTeX already solved this problem.  If an .aux file is available it is
    authoritative: cleveref writes both the reference type and the printed
    number for every label.  Handles every package and macro, exactly.

    `model` supplies the document's own \\crefname declarations, so a counter
    called `appx` still displays as "Appendix" rather than "Appx".
    """
    out = {}
    for m in _AUX_CREF.finditer(aux_text):
        ctype = m.group("type")
        name = (model.name_for(ctype) if model else None) \
            or DEFAULT_CREFNAMES.get(ctype) \
            or (ctype.title() if ctype else None)
        num = m.group("num").strip()
        out[m.group("label")] = (name, num or None)
    return out


def section_numbers(src, model=None):
    """[(level, number_or_None), ...] for every sectioning command, in order.

    LaTeX prints "2.2 Experiment 1: ...".  pandoc maps the heading to a Word
    Heading style, which carries no number, so the Word file loses it -- and a
    numbered heading sitting on a page boundary then reads as different text
    on the two sides.  Starred sections are unnumbered and come back as None.
    """
    model = model or scan_preamble(src)
    body = strip_comments(src).split("\\begin{document}", 1)[-1]
    levels = {"section": 1, "subsection": 2, "subsubsection": 3}
    pattern = re.compile(
        r"\\(section|subsection|subsubsection)(\*?)\s*\{|"
        r"\\(setcounter|addtocounter|stepcounter)\s*\{([a-zA-Z@]+)\}"
        r"(?:\s*\{(-?\d+)\})?")
    out = []
    for m in pattern.finditer(body):
        if m.group(1):
            counter = m.group(1)
            if m.group(2) == "*":
                out.append((levels[counter], None))
            else:
                model.step(counter)
                out.append((levels[counter], model.render(counter)))
        elif m.group(3):
            op, c, val = m.group(3), m.group(4), m.group(5)
            if op == "stepcounter":
                model.step(c)
            elif val is not None:
                cur = model.values.get(c, 0)
                model.values[c] = (int(val) if op == "setcounter"
                                   else cur + int(val))
    return out
