#!/usr/bin/env python3
"""What moves a page break: widow and orphan control, ``keepLines``, ``keepNext`` and its
chains, ``pageBreakBefore``, and manual page breaks.

Every case starts a page: a plain anchor paragraph with ``w:pageBreakBefore``, then a
filler paragraph of ``f`` lines, then the paragraphs under test.  Every line is Calibri
11 pt, single spaced, no paragraph spacing -- 55,000 layout units -- and the page's text
area (``HEIGHT`` less two 1,440 margins: 13,830 twips) holds 51 of them with half a line
over, so no case sits near the fit threshold (``make_page_fit_probe.py`` measured that).
``room`` is how many lines the page has left for the paragraphs under test; the filler
takes the rest.  A test paragraph's lines are separated by ``w:br`` (a ``natural``
family wraps real words instead, to show that a manual line break is a line like any
other here).

Which page each line of each test paragraph lands on is read from Word's export, as an
offset from its anchor's page.

Families (``CASES``):

* ``widow`` -- one paragraph of ``n`` = 2..7 lines with ``room`` = 1..n: widow control
  as stated on the paragraph (``on``), turned off there (``off``), turned off by its
  paragraph style (``style-off``), and not stated anywhere (``default``).
* ``natural`` -- the same, ``default``, with lines wrapped from words (n = 3..5).
* ``keeplines`` -- ``keepLines`` on paragraphs of 3 and 5 lines, widow control on and
  off; and one of 60 lines (taller than a page) starting with 10 lines of room.
* ``keepnext`` -- ``A`` (``keepNext``) then ``B``: ``A`` of one line or of four, ``B`` of
  one line or three, each with widow control on and off, over every room from 1 to
  their sum; and ``keepLines`` with it.
* ``chain`` -- two ``keepNext`` paragraphs then a plain one of three lines, rooms 1..5;
  and a chain of 60 one-line ``keepNext`` paragraphs (taller than a page) starting with
  10 lines of room, with and without a last paragraph that ends it.
* ``pagebreak`` -- ``pageBreakBefore`` on a paragraph that would start the next page
  anyway (room 0), and after a ``keepNext`` paragraph.
* ``manual`` -- ``w:br w:type="page"`` in the middle of a paragraph, at its end, at its
  start, alone, with widow control on the lines after it; and a column break in a
  one-column section.
* ``manual segment`` -- ``n`` = 2..5 lines, a page break, two lines; rooms 1..n, widow
  control on and off: is the line before a page break a paragraph's last line to
  widow control?

One document per compatibility setting (no ``settings.xml``; modes 12, 14 and 15), and
one with no styles part at all (widow ``default`` cases only).
"""

from __future__ import annotations

import functools
from dataclasses import dataclass, field

import probe_docx
import wml

WIDTH, HEIGHT, MARGIN = 11900, 16710, 1440
PER_PAGE = 51

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
RUN = {"rFonts": FACE, "sz": 22, "szCs": 22}
SPACING = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}

#: A few words per line for the ``natural`` family: Calibri 11 pt, 9,020 twips of column.
WORDS = ("alpha", "bravo", "charlie", "delta", "echo", "foxtrot", "golf", "hotel", "india",
         "juliet", "kilo", "lima", "mike", "november", "oscar", "papa", "quebec", "romeo")


@dataclass(frozen=True)
class P:
    """One paragraph under test: ``lines`` lines (``w:br`` between them), its
    properties, and how its lines are made."""

    lines: int
    props: tuple = ()
    #: ``br`` (lines split by ``w:br``), ``words`` (wrapped), or a manual-break layout:
    #: ``mid`` (two lines, a page break, two lines), ``end`` (text then a page break),
    #: ``start`` (a page break then text), ``alone`` (only a page break), ``widow`` (one
    #: line, a page break, three lines), ``column`` (a column break mid-paragraph).
    make: str = "br"
    style: str | None = None


@dataclass(frozen=True)
class Case:
    family: str
    variant: str
    room: int
    paragraphs: tuple[P, ...]
    #: Only in these documents (``None``: every one).
    documents: tuple[str, ...] | None = None

    @property
    def key(self) -> str:
        shape = "+".join(f"{p.lines}{p.make if p.make != 'br' else ''}" for p in self.paragraphs)
        return f"{self.family}/{self.variant} {shape} room {self.room}"


def _widow(variant: str) -> tuple[tuple, str | None]:
    if variant == "on":
        return (("widowControl", True),), None
    if variant == "off":
        return (("widowControl", False),), None
    if variant == "style-off":
        return (), "NoWidow"
    return (), None


def _cases() -> tuple[Case, ...]:
    out: list[Case] = []
    for variant in ("default", "on", "off", "style-off"):
        props, style = _widow(variant)
        for n in range(2, 8):
            for room in range(1, n + 1):
                out.append(Case("widow", variant, room, (P(n, props, style=style),)))
    for n in range(3, 6):
        for room in range(1, n + 1):
            out.append(Case("natural", "default", room, (P(n, make="words"),)))
    for widow in (True, False):
        for n in (3, 5):
            for room in range(1, n + 1):
                out.append(Case("keeplines", f"widow {'on' if widow else 'off'}", room,
                                (P(n, (("keepLines", True), ("widowControl", widow))),)))
    out.append(Case("keeplines", "tall", 10, (P(60, (("keepLines", True),)),)))
    out.append(Case("keeplines", "tall, at top", 50, (P(60, (("keepLines", True),)),)))
    for a_lines, b_lines in ((1, 1), (1, 3), (4, 1), (4, 3)):
        for widow in (True, False):
            for room in range(1, a_lines + b_lines + 1):
                out.append(Case("keepnext", f"A{a_lines} B{b_lines} widow {'on' if widow else 'off'}", room, (
                    P(a_lines, (("keepNext", True), ("widowControl", widow))),
                    P(b_lines, (("widowControl", widow),)))))
    for room in range(1, 5):
        out.append(Case("keepnext", "A3 keepLines B1", room, (
            P(3, (("keepNext", True), ("keepLines", True))), P(1))))
    for room in range(1, 6):
        out.append(Case("chain", "H1 H2 P3", room, (
            P(1, (("keepNext", True),)), P(1, (("keepNext", True),)), P(3))))
    out.append(Case("chain", "60 tall", 10, tuple(P(1, (("keepNext", True),)) for _ in range(60)) + (P(1),)))
    out.append(Case("chain", "60 tall, open end", 10, tuple(P(1, (("keepNext", True),)) for _ in range(60))))
    out.append(Case("chain", "60 tall, at top", 50, tuple(P(1, (("keepNext", True),)) for _ in range(60)) + (P(1),)))
    out.append(Case("pagebreak", "at a natural top", 0, (P(1, (("pageBreakBefore", True),)),)))
    out.append(Case("pagebreak", "after keepNext", 20, (
        P(1, (("keepNext", True),)), P(1, (("pageBreakBefore", True),)))))
    for make, lines in (("mid", 4), ("end", 1), ("start", 1), ("alone", 0), ("widow", 4), ("column", 4)):
        for room in (20, 1):
            out.append(Case("manual", make, room, (P(lines, make=make), P(1))))
    # The second round, after the pagination probe: widow control over the lines before
    # a manual page break -- n lines, the break, two more -- with widow control on and off.
    for widow in (True, False):
        for n in (2, 3, 4, 5):
            for room in range(1, n + 1):
                out.append(Case("manual segment", f"widow {'on' if widow else 'off'}", room, (
                    P(n + 2, (("widowControl", widow),), make=f"segment{n}"), P(1))))
    return tuple(out)


CASES = _cases()

#: name -> (compatibilityMode, a styles part)
DOCUMENTS = {"none": (None, True), "12": (12, True), "14": (14, True), "15": (15, True),
             "nostyles": (None, False)}


def _in(case: Case, document: str) -> bool:
    if document == "nostyles":
        return case.family == "widow" and case.variant == "default"
    return True


def _br_runs(texts: list[str]) -> str:
    return f"<w:r>{wml.rpr(**RUN)}" + "<w:br/>".join(
        f'<w:t xml:space="preserve">{t}</w:t>' for t in texts) + "</w:r>"


def _text(texts: list[str]) -> str:
    return f"<w:r>{wml.rpr(**RUN)}" + "".join(
        f'<w:t xml:space="preserve">{t}</w:t>' if not t.startswith("<") else t for t in texts) + "</w:r>"


def _paragraph(number: int, index: int, p: P) -> str:
    props = dict(p.props)
    if p.style:
        props["pStyle"] = p.style
    tag = f"Case {number} p{index}"
    if p.make == "br":
        runs = _br_runs([f"{tag} l{k}" for k in range(p.lines)])
    elif p.make == "words":
        # About nine words to a line: enough for ``lines`` lines and a short last one.
        count = 9 * p.lines - 4
        words = [f"{tag}"] + [f"{WORDS[k % len(WORDS)]}{k}" for k in range(count)]
        runs = _text([" ".join(words)])
    else:
        page = '<w:br w:type="page"/>'
        layouts = {
            "mid": [f"{tag} l0", "<w:br/>", f"{tag} l1", page, f"{tag} l2", "<w:br/>", f"{tag} l3"],
            "end": [f"{tag} l0", page],
            "start": [page, f"{tag} l0"],
            "alone": [page],
            "widow": [f"{tag} l0", page, f"{tag} l1", "<w:br/>", f"{tag} l2", "<w:br/>", f"{tag} l3"],
            "column": [f"{tag} l0", "<w:br/>", f"{tag} l1", '<w:br w:type="column"/>', f"{tag} l2", "<w:br/>",
                       f"{tag} l3"],
        }
        if p.make.startswith("segment"):
            n = int(p.make[len("segment"):])
            parts = []
            for k in range(p.lines):
                if k:
                    parts.append(page if k == n else "<w:br/>")
                parts.append(f"{tag} l{k}")
            runs = _text(parts)
        else:
            runs = _text(layouts[p.make])
    props.setdefault("spacing", SPACING)
    return wml.paragraph(runs, mark=RUN, **props)


def case_blocks(number: int, case: Case) -> list[tuple[str, int, str]]:
    """``(role, paragraph index under test or -1, w:p)`` for one case."""
    out = [("anchor", -1, wml.paragraph(_text([f"Case {number} anchor"]), mark=RUN, pageBreakBefore=True,
                                        spacing=SPACING))]
    filler = PER_PAGE - 1 - case.room
    if filler:
        out.append(("filler", -1, wml.paragraph(_br_runs(["f"] * filler), mark=RUN, spacing=SPACING)))
    for index, p in enumerate(case.paragraphs):
        out.append(("test", index, _paragraph(number, index, p)))
    return out


@functools.lru_cache(maxsize=None)
def blocks(document: str) -> list[tuple[int, str, int, str]]:
    """``(case number, role, paragraph index, w:p)`` for every paragraph of ``document``."""
    out = []
    for number, case in enumerate(CASES):
        if _in(case, document):
            out += [(number, role, index, xml) for role, index, xml in case_blocks(number, case)]
    return out


def build(document: str) -> bytes:
    mode, with_styles = DOCUMENTS[document]
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True),
         wml.style("paragraph", "NoWidow", based_on="Normal", ppr_={"widowControl": False})],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": SPACING},
    ) if with_styles else ""
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    section = (
        "<w:sectPr>"
        f'<w:pgSz w:w="{WIDTH}" w:h="{HEIGHT}"/>'
        f'<w:pgMar w:top="{MARGIN}" w:right="{MARGIN}" w:bottom="{MARGIN}" w:left="{MARGIN}"'
        ' w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>'
    )
    return probe_docx.package("".join(xml for *_, xml in blocks(document)), final_section=section,
                              styles=styles, extra_parts=extra)
