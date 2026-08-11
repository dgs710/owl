"""Keep `\\textcolor` colours in the Word file.

pandoc's LaTeX reader drops colour entirely, so a document that uses
`\\textcolor{red}{...}` for open questions and `\\textcolor{fixgreen}{...}` for
resolved ones arrives in Word as undifferentiated black text -- the colour was
carrying meaning, and the meaning is gone.

Same trick as the cross-references: wrap the coloured span in private-use
sentinels that pandoc passes through untouched, then turn them into real
`w:color` runs afterwards.  The span's *contents* stay ordinary LaTeX, so
pandoc still converts whatever is inside -- maths, citations, emphasis.
"""

import re

from docx.oxml.ns import qn
from docx.oxml import OxmlElement

COLOR_A = "\ue003"        # <A> RRGGBB <S> ...text... <E>
COLOR_S = "\ue004"
COLOR_E = "\ue005"

# the xcolor names that ship by default
NAMED = {
    "red": "FF0000", "green": "00FF00", "blue": "0000FF", "cyan": "00FFFF",
    "magenta": "FF00FF", "yellow": "FFFF00", "black": "000000",
    "white": "FFFFFF", "gray": "808080", "grey": "808080",
    "darkgray": "404040", "lightgray": "BFBFBF", "brown": "BF8040",
    "lime": "BFFF00", "olive": "808000", "orange": "FF8000",
    "pink": "FFBFBF", "purple": "BF0040", "teal": "008080",
    "violet": "800080",
}


def palette(src):
    """Read \\definecolor declarations and merge them over the defaults."""
    out = dict(NAMED)
    for m in re.finditer(r"\\definecolor\s*\{([^}]+)\}\s*\{([^}]+)\}\s*\{([^}]+)\}",
                         src):
        # xcolor is case-sensitive here: {rgb} takes fractions 0-1,
        # {RGB} takes integers 0-255.  Lower-casing the model turns
        # {RGB}{0,125,60} into a fraction triple and produces nonsense.
        name, model, value = (m.group(1).strip(), m.group(2).strip(),
                              m.group(3).strip())
        try:
            if model == "rgb":
                r, g, b = (float(x) for x in value.split(","))
                out[name] = "%02X%02X%02X" % (round(r * 255), round(g * 255),
                                              round(b * 255))
            elif model == "RGB":
                r, g, b = (max(0, min(255, int(float(x))))
                           for x in value.split(","))
                out[name] = "%02X%02X%02X" % (r, g, b)
            elif model.lower() == "html":
                out[name] = value.upper().lstrip("#")[:6]
            elif model.lower() == "gray":
                v = round(float(value) * 255)
                out[name] = "%02X%02X%02X" % (v, v, v)
            elif model.lower() == "cmyk":
                c, m_, y, k = (float(x) for x in value.split(","))
                out[name] = "%02X%02X%02X" % (
                    round(255 * (1 - c) * (1 - k)),
                    round(255 * (1 - m_) * (1 - k)),
                    round(255 * (1 - y) * (1 - k)))
        except (ValueError, ZeroDivisionError):
            continue
    return out


def _match_brace(src, open_idx):
    depth, j = 1, open_idx + 1
    while j < len(src) and depth:
        if src[j] == "\\":
            j += 2
            continue
        if src[j] == "{":
            depth += 1
        elif src[j] == "}":
            depth -= 1
        j += 1
    return (src[open_idx + 1:j - 1], j) if depth == 0 else (None, open_idx)


def mark(src, colors=None):
    """Rewrite \\textcolor{name}{body} -> sentinel-wrapped body."""
    colors = colors or palette(src)
    out, pos = [], 0
    pattern = re.compile(r"\\textcolor\s*(?:\[[^\]]*\])?\s*\{")
    while True:
        m = pattern.search(src, pos)
        if not m:
            out.append(src[pos:])
            break
        name, after = _match_brace(src, m.end() - 1)
        if name is None:
            out.append(src[pos:m.end()])
            pos = m.end()
            continue
        while after < len(src) and src[after] in " \t\n":
            after += 1
        if after >= len(src) or src[after] != "{":
            out.append(src[pos:m.end()])
            pos = m.end()
            continue
        body, end = _match_brace(src, after)
        if body is None:
            out.append(src[pos:m.end()])
            pos = m.end()
            continue
        hexv = colors.get(name.strip())
        out.append(src[pos:m.start()])
        if hexv:
            out.append(COLOR_A + hexv + COLOR_S + body + COLOR_E)
        else:
            out.append(body)
        pos = end
    return "".join(out)


# ---------------------------------------------------------------------------
# docx side
# ---------------------------------------------------------------------------

def _set_color(r_el, hexv):
    rPr = r_el.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        r_el.insert(0, rPr)
    for old in rPr.findall(qn("w:color")):
        rPr.remove(old)
    c = OxmlElement("w:color")
    c.set(qn("w:val"), hexv)
    rPr.append(c)


def _split_run(r_el, offset):
    """Split a run's text at `offset`; returns the new trailing run."""
    import copy
    t = r_el.find(qn("w:t"))
    if t is None:
        return None
    text = t.text or ""
    new_r = copy.deepcopy(r_el)
    t.text = text[:offset]
    t.set(qn("xml:space"), "preserve")
    nt = new_r.find(qn("w:t"))
    nt.text = text[offset:]
    nt.set(qn("xml:space"), "preserve")
    r_el.addnext(new_r)
    return new_r


def apply(doc):
    """Turn the sentinels into real coloured runs.

    A coloured span can cross runs -- pandoc splits on emphasis, maths and
    links -- so this walks the document's runs in order and colours everything
    between an opening and a closing sentinel, splitting only the two runs the
    sentinels actually sit in.
    """
    body = doc.element.body
    runs = [r for r in body.iter(qn("w:r")) if r.find(qn("w:t")) is not None]
    i = 0
    coloured = 0
    while i < len(runs):
        r = runs[i]
        t = r.find(qn("w:t"))
        text = t.text or ""
        m = re.search(re.escape(COLOR_A) + r"([0-9A-Fa-f]{6})" + re.escape(COLOR_S),
                      text)
        if not m:
            i += 1
            continue
        hexv = m.group(1).upper()

        # drop the opening sentinel, keeping anything before it uncoloured
        head, tail = text[:m.start()], text[m.end():]
        if head:
            t.text = head
            nr = _split_run(r, len(head))
            if nr is not None:
                r = nr
            t = r.find(qn("w:t"))
            runs.insert(i + 1, r)
            i += 1
        t.text = tail

        # colour forward until the closing sentinel
        j = i
        while j < len(runs):
            rj = runs[j]
            tj = rj.find(qn("w:t"))
            txt = tj.text or ""
            k = txt.find(COLOR_E)
            if k == -1:
                _set_color(rj, hexv)
                j += 1
                continue
            tj.text = txt[:k]
            rest = txt[k + 1:]
            _set_color(rj, hexv)
            if rest:
                nr = _split_run(rj, len(tj.text))
                if nr is not None:
                    nt = nr.find(qn("w:t"))
                    nt.text = rest
                    rPr = nr.find(qn("w:rPr"))
                    if rPr is not None:
                        for c in rPr.findall(qn("w:color")):
                            rPr.remove(c)
                    runs.insert(j + 1, nr)
            coloured += 1
            break
        i = j + 1

    # belt and braces: no sentinel may survive into the visible text.
    dead = re.compile(re.escape(COLOR_A) + r"[0-9A-Fa-f]{6}" + re.escape(COLOR_S)
                      + "|" + re.escape(COLOR_E))
    for t in body.iter(qn("w:t")):
        if t.text and (COLOR_A in t.text or COLOR_S in t.text
                       or COLOR_E in t.text):
            t.text = dead.sub("", t.text)

    # Word maths is a special case: pandoc emits one <m:t> per character, so a
    # marker is split across several elements and no per-element regex can see
    # it.  (\textcolor inside $...$ is not valid LaTeX anyway -- preflight warns
    # about it -- but the reader must still get the formula, not private-use
    # gibberish.)  Treat the maths text as one stream and cut the spans out.
    cells = [t for t in body.iter(qn("m:t")) if t.text]
    if cells:
        stream = "".join(t.text for t in cells)
        if COLOR_A in stream or COLOR_S in stream or COLOR_E in stream:
            drop = set()
            for m in dead.finditer(stream):
                drop.update(range(m.start(), m.end()))
            for ch in (COLOR_A, COLOR_S, COLOR_E):
                drop.update(i for i, c in enumerate(stream) if c == ch)
            pos = 0
            for t in cells:
                n = len(t.text)
                kept = "".join(c for k, c in enumerate(t.text)
                               if (pos + k) not in drop)
                if kept != t.text:
                    t.text = kept
                pos += n
    return coloured
