#!/usr/bin/env python3
"""Style-numbered paragraphs and the level they take: a direct ``w:numPr`` holding only
``w:ilvl``, a level past the definition's, and a style whose numbering level is tied
(``w:lvl/w:pStyle``) to another style.

docx-agent's lists (round 23) found three places docx2svg drew a list item otherwise than
Word.  Every case is a page: ``Case N``, then its paragraphs; each case has numbering
definitions and styles of its own (so no count runs from one case into the next):

* ``T`` -- a three-level definition, level 0 ``%1.`` tied to the style ``LN`` (indent
  284/284), level 1 ``%2)`` tied to ``LNb`` (568/284), level 2 ``%3.`` tied to none
  (852/284); ``LN``'s ``w:numPr`` names the definition's ``w:num``, ``LNb``'s level 1 too;
* ``S`` -- a single-level definition (``singleLevel``), its level tied to ``LS``;
* ``M`` -- a definition of one level that says ``multilevel``, tied to ``LM``.

Cases: ``ilvl only`` -- ``LN`` items, one with a direct ``w:numPr`` of ``w:ilvl`` 1 and
one of ``w:ilvl`` 2 alone; ``ilvl 0 on LNb``; ``single, ilvl only`` and ``single, direct``
(``w:numId`` and ``w:ilvl`` 1, and ``w:ilvl`` 3 on a plain paragraph); ``one level,
multilevel``; ``past three levels`` (``w:ilvl`` 4 and 5); ``own numPr`` -- a style ``Own``
whose own ``w:numPr`` names ``T``'s num (level 0, tied to ``LN``), alone and with a direct
``w:numPr``; ``child`` -- a style based on ``LN`` (inheriting its ``w:numPr``); ``own,
untied level`` (``Own2``: level 2); ``own, level 1`` (``Own3``: level 1, tied to ``LNb``);
``untied definition`` (a style numbered by a definition no level of which is tied);
``child of own``.

One document per compatibility setting (none, 15).  Calibri 11, no paragraph spacing.
Reader: ``read_list_level_probe.py``.
"""

from __future__ import annotations

import probe_docx
import wml

SETTINGS = {"none": None, "15": 15}
FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
SPACING = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}
NUMBERING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
NUMBERING_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"


def _numpr(num: int | None = None, ilvl: int | None = None) -> str:
    return ("<w:numPr>" + (f'<w:ilvl w:val="{ilvl}"/>' if ilvl is not None else "")
            + (f'<w:numId w:val="{num}"/>' if num is not None else "") + "</w:numPr>")


def _ind(left: int, hanging: int) -> str:
    return f'<w:ind w:left="{left}" w:hanging="{hanging}"/>'


def _lvl(ilvl: int, fmt: str, text: str, left: int, hanging: int, style: str | None = None) -> str:
    return (f'<w:lvl w:ilvl="{ilvl}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
            + (f'<w:pStyle w:val="{style}"/>' if style else "")
            + f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/><w:pPr>{_ind(left, hanging)}</w:pPr></w:lvl>')


def _abstract(aid: int, kind: str, levels: list[str]) -> str:
    return (f'<w:abstractNum w:abstractNumId="{aid}"><w:nsid w:val="{0x2C000000 + aid:08X}"/>'
            f'<w:multiLevelType w:val="{kind}"/>{"".join(levels)}</w:abstractNum>')


def _style(sid: str, ppr: str = "", based: str = "Normal") -> str:
    return (f'<w:style w:type="paragraph" w:styleId="{sid}"><w:name w:val="{sid}"/>'
            f'<w:basedOn w:val="{based}"/><w:pPr>{ppr}</w:pPr></w:style>')


def _p(text: str, style: str | None = None, ppr: str = "") -> str:
    props = (f'<w:pStyle w:val="{style}"/>' if style else "") + ppr
    return f'<w:p>{f"<w:pPr>{props}</w:pPr>" if props else ""}{wml.run(text)}</w:p>'


class Case:
    """One case's definitions, styles and paragraphs; ids from its number."""

    def __init__(self, number: int):
        self.n = number
        self.abstracts: list[str] = []
        self.nums: list[str] = []
        self.styles: list[str] = []

    def sid(self, name: str) -> str:
        return f"{name}{self.n}"

    def define(self, k: int, kind: str, levels: list[str]) -> int:
        """Definition ``k`` of the case and its ``w:num``: the num's id."""
        ident = 10 * self.n + k
        self.abstracts.append(_abstract(ident, kind, levels))
        self.nums.append(f'<w:num w:numId="{ident}"><w:abstractNumId w:val="{ident}"/></w:num>')
        return ident

    def template(self) -> int:
        num = self.define(1, "multilevel", [_lvl(0, "decimal", "%1.", 284, 284, self.sid("LN")),
                                            _lvl(1, "lowerLetter", "%2)", 568, 284, self.sid("LNb")),
                                            _lvl(2, "lowerRoman", "%3.", 852, 284)])
        self.styles += [_style(self.sid("LN"), _numpr(num) + _ind(284, 284)),
                        _style(self.sid("LNb"), _numpr(num, 1) + _ind(568, 284))]
        return num


def _cases() -> list[tuple[str, object]]:
    out: list[tuple[str, object]] = []

    def case(name):
        def add(build):
            out.append((name, build))
            return build
        return add

    @case("ilvl only")
    def _(c: Case) -> str:
        c.template()
        ln = c.sid("LN")
        return (_p("Aa", ln) + _p("Ab", ln) + _p("Ac", ln, _numpr(ilvl=1)) + _p("Ad", ln) + _p("Ae", c.sid("LNb"))
                + _p("Af", ln, _numpr(ilvl=2)) + _p("Ag", ln))

    @case("ilvl 0 on LNb")
    def _(c: Case) -> str:
        c.template()
        return _p("Ba", c.sid("LN")) + _p("Bb", c.sid("LNb")) + _p("Bc", c.sid("LNb"), _numpr(ilvl=0)) + _p(
            "Bd", c.sid("LNb"))

    @case("single, ilvl only")
    def _(c: Case) -> str:
        num = c.define(2, "singleLevel", [_lvl(0, "decimal", "%1.", 360, 360, c.sid("LS"))])
        c.styles.append(_style(c.sid("LS"), _numpr(num)))
        ls = c.sid("LS")
        return _p("Ca", ls) + _p("Cb", ls) + _p("Cc", ls, _numpr(ilvl=1)) + _p("Cd", ls)

    @case("single, direct")
    def _(c: Case) -> str:
        num = c.define(2, "singleLevel", [_lvl(0, "decimal", "%1.", 360, 360, c.sid("LS"))])
        c.styles.append(_style(c.sid("LS"), _numpr(num)))
        ls = c.sid("LS")
        return _p("Da", ls) + _p("Db", ls, _numpr(num, 1)) + _p("Dc", ls) + _p("Dd", None, _numpr(num, 3))

    @case("one level, multilevel")
    def _(c: Case) -> str:
        num = c.define(3, "multilevel", [_lvl(0, "decimal", "%1.", 360, 360, c.sid("LM"))])
        c.styles.append(_style(c.sid("LM"), _numpr(num)))
        lm = c.sid("LM")
        return _p("Ea", lm) + _p("Eb", lm, _numpr(ilvl=1)) + _p("Ec", lm)

    @case("past three levels")
    def _(c: Case) -> str:
        num = c.template()
        ln = c.sid("LN")
        return _p("Fa", ln) + _p("Fb", ln, _numpr(ilvl=4)) + _p("Fc", ln) + _p("Fd", None, _numpr(num, 5)) + _p(
            "Fe", ln)

    @case("own numPr")
    def _(c: Case) -> str:
        num = c.template()
        c.styles.append(_style(c.sid("Own"), _numpr(num)))
        ln, own = c.sid("LN"), c.sid("Own")
        return (_p("Ga", ln) + _p("Gb", own) + _p("Gc", ln) + _p("Gd", own, _numpr(num, 0)) + _p("Ge", ln))

    @case("child")
    def _(c: Case) -> str:
        c.template()
        c.styles.append(_style(c.sid("Child"), "", based=c.sid("LN")))
        return _p("Ha", c.sid("LN")) + _p("Hb", c.sid("Child")) + _p("Hc", c.sid("LN"))

    @case("own, untied level")
    def _(c: Case) -> str:
        num = c.template()
        c.styles.append(_style(c.sid("Own"), _numpr(num, 2)))
        return _p("Ia", c.sid("LN")) + _p("Ib", c.sid("Own")) + _p("Ic", c.sid("Own")) + _p("Id", c.sid("LN"))

    @case("own, level 1")
    def _(c: Case) -> str:
        num = c.template()
        c.styles.append(_style(c.sid("Own"), _numpr(num, 1)))
        return _p("Ja", c.sid("LN")) + _p("Jb", c.sid("Own")) + _p("Jc", c.sid("LNb")) + _p("Jd", c.sid("LN"))

    @case("untied definition")
    def _(c: Case) -> str:
        num = c.define(4, "multilevel", [_lvl(0, "decimal", "%1.", 360, 360), _lvl(1, "lowerLetter", "%2.", 720, 360)])
        c.styles += [_style(c.sid("U"), _numpr(num)), _style(c.sid("Ub"), _numpr(num, 1))]
        return _p("Ka", c.sid("U")) + _p("Kb", c.sid("Ub")) + _p("Kc", c.sid("U"))

    @case("child of own")
    def _(c: Case) -> str:
        num = c.template()
        c.styles += [_style(c.sid("Own"), _numpr(num)), _style(c.sid("Kid"), "", based=c.sid("Own"))]
        return _p("La", c.sid("LN")) + _p("Lb", c.sid("Kid")) + _p("Lc", c.sid("LN"))

    @case("own numPr, based on LN")
    def _(c: Case) -> str:
        num = c.template()
        c.styles.append(_style(c.sid("Mix"), _numpr(num, 0), based=c.sid("LN")))
        return _p("Ma", c.sid("LN")) + _p("Mb", c.sid("Mix")) + _p("Mc", c.sid("LN"))

    @case("past the levels, indents")
    def _(c: Case) -> str:
        tied = c.define(5, "multilevel", [_lvl(0, "decimal", "%1.", 720, 360, c.sid("Q")),
                                          _lvl(1, "lowerLetter", "%2.", 1440, 360)])
        untied = c.define(6, "multilevel", [_lvl(0, "decimal", "%1.", 720, 360),
                                            _lvl(1, "lowerLetter", "%2.", 1440, 360)])
        c.styles += [_style(c.sid("Q"), _numpr(tied) + _ind(284, 284)), _style(c.sid("Y"), _numpr(untied))]
        q, y = c.sid("Q"), c.sid("Y")
        return (_p("Na", q) + _p("Nb", q, _numpr(ilvl=4)) + _p("Nc", q) + _p("Nd", y) + _p("Ne", y, _numpr(ilvl=4))
                + _p("Nf", y) + _p("Ng", None, _numpr(untied, 5) + _ind(500, 300)) + _p("Nh", None, _numpr(untied, 5))
                + _p("Ni", y))

    @case("tied level, style without ilvl")
    def _(c: Case) -> str:
        num = c.template()
        c.styles[-1] = _style(c.sid("LNb"), _numpr(num) + _ind(568, 284))
        return _p("Oa", c.sid("LN")) + _p("Ob", c.sid("LNb")) + _p("Oc", c.sid("LNb")) + _p("Od", c.sid("LN"))

    return out


CASES = _cases()


def parts() -> tuple[str, str, str]:
    """The body, the numbering part and the styles part's case styles."""
    body = ""
    abstracts: list[str] = []
    nums: list[str] = []
    styles: list[str] = []
    for number, (_, build) in enumerate(CASES):
        case = Case(number)
        content = build(case)
        body += wml.paragraph(wml.run(f"Case {number}"), pageBreakBefore=bool(number)) + content
        abstracts += case.abstracts
        nums += case.nums
        styles += case.styles
    numbering = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                 f'<w:numbering xmlns:w="{probe_docx.W_NS}">' + "".join(abstracts) + "".join(nums) + "</w:numbering>")
    return body, numbering, "".join(styles)


def styles(extra: str) -> str:
    part = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": SPACING},
    )
    return part.replace("</w:styles>", extra + "</w:styles>")


def build(setting: str) -> bytes:
    mode = SETTINGS[setting]
    body, numbering, extra_styles = parts()
    extra = [("word/numbering.xml", NUMBERING_CONTENT_TYPE, NUMBERING_REL, numbering)]
    if mode is not None:
        extra.append(wml.settings_part({"val": "en-GB"}, compatibility_mode=mode))
    return probe_docx.package(body, styles=styles(extra_styles), extra_parts=tuple(extra))
