#!/usr/bin/env python3
"""What Word draws for ``PAGEREF``, ``REF`` and ``SEQ``: which page, which text, which
number, and in which run's format.

Word computes these three again when it exports (``make_field_probe.py``, ``word``), so
the cached result in the file says nothing; every field here caches a result Word could
not have computed (``999``, ``cached``).  The runs of a field differ in format so the
export shows whose format the result is drawn in: the ``begin`` / ``separate`` / ``end``
runs in the document's Calibri 11, the instruction in Arial 9, the cached result in
Georgia 12 italic, and a bookmark's text in Times New Roman 14 bold.

The bookmarks (``w:bookmarkStart`` / ``w:bookmarkEnd``):

* ``alpha`` -- a run's text on the first page;
* ``beta`` -- a paragraph's text on the third page;
* ``gamma`` -- in the second section (``lowerRoman`` from 10), on its second page;
* ``multi`` -- across two paragraphs;
* ``caption`` -- ``Figure``, a ``SEQ Figure`` and the caption's text;
* ``empty`` -- a start and an end with nothing between, on the fourth page;
* ``cell`` -- a table cell's text on the second page;
* ``late`` -- starting in the last line of a paragraph that begins on the fourth page and
  ends on the fifth;
* ``styled`` -- a plain run in a paragraph of a style that is Cambria 16 (``Heading1``);
* ``charstyled`` -- a run of a character style that is Courier New 16 (``Big``);
* ``missing`` -- referred to, never defined.

Every case is a paragraph of its own, labelled (``R1[`` ... ``]``).  ``SEQ`` cases count
``Figure`` (three, then ``\\r 5``, ``\\c``, ``\\h``, ``\\* roman``, ``\\* alphabetic``, in a
table cell, a lower-case ``figure``), ``Table`` apart, and ``\\s 1``.  A table of contents
holds entries whose page numbers are ``PAGEREF`` fields cached ``77``.

Two documents: no ``settings.xml`` and mode 15.
"""

from __future__ import annotations

import functools

import story_docx
import wml
from story_docx import field, simple_field

SETTINGS = {"none": None, "15": 15}

INSTR = {"rFonts": {"ascii": "Arial", "hAnsi": "Arial", "eastAsia": "Arial", "cs": "Arial"}, "sz": 18, "szCs": 18}
RESULT = {"rFonts": {"ascii": "Georgia", "hAnsi": "Georgia", "eastAsia": "Georgia", "cs": "Georgia"},
          "i": True, "sz": 24, "szCs": 24}
MARKED = {"rFonts": {"ascii": "Times New Roman", "hAnsi": "Times New Roman", "eastAsia": "Times New Roman",
                     "cs": "Times New Roman"}, "b": True, "sz": 28, "szCs": 28}

_IDS = {name: number for number, name in enumerate(
    ("alpha", "beta", "gamma", "multi", "caption", "empty", "cell", "late", "styled", "charstyled"), start=1)}


def start(name: str) -> str:
    return f'<w:bookmarkStart w:id="{_IDS[name]}" w:name="{name}"/>'


def end(name: str) -> str:
    return f'<w:bookmarkEnd w:id="{_IDS[name]}"/>'


def case(label: str, instruction: str, cached: str = "999", *, simple: bool = False) -> str:
    """A labelled paragraph holding one field, its runs formatted apart."""
    if simple:
        inner = simple_field(instruction, cached, rpr=RESULT)
    else:
        inner = field(instruction, cached, rpr={}, instr_rpr=INSTR, result_rpr=RESULT)
    return wml.paragraph(wml.run(f"{label}[") + inner + wml.run("]"), mark={})


#: ``(label, instruction)``: the REF cases.
REFS = (
    ("R1", "REF alpha \\h"), ("R2", "REF alpha \\* MERGEFORMAT"), ("R3", "REF alpha \\* CHARFORMAT"),
    ("R4", "REF alpha \\* Upper"), ("R6", "REF caption \\h"), ("R7", "REF empty"),
    ("R8", "REF cell \\h"), ("R9", "REF missing \\h"), ("R10", "REF beta"), ("R11", "REF late"),
    ("R13", "REF styled \\h"), ("R14", "REF charstyled \\h"), ("R15", "REF alpha \\* FirstCap"),
)
#: The PAGEREF cases.
PAGEREFS = (
    ("P1", "PAGEREF beta \\h"), ("P2", "PAGEREF gamma \\h"), ("P3", "PAGEREF gamma \\* Arabic"),
    ("P4", "PAGEREF late \\h"), ("P5", "PAGEREF cell \\h"), ("P6", "PAGEREF beta \\* MERGEFORMAT"),
    ("P7", "PAGEREF missing \\h"), ("P8", "PAGEREF beta \\p"), ("P9", "PAGEREF alpha"),
    ("P10", "PAGEREF empty \\* roman"), ("P11", 'PAGEREF beta \\# "00"'),
)
#: The SEQ cases, in document order.
SEQS = (
    ("S1", "SEQ Figure \\* ARABIC"), ("S2", "SEQ Figure \\* ARABIC"), ("S3", "SEQ Table \\* ARABIC"),
    ("S4", "SEQ Figure"), ("S5", "SEQ Figure \\r 5"), ("S6", "SEQ Figure \\c"), ("S7", "SEQ Figure \\h"),
    ("S8", "SEQ Figure \\* roman"), ("S9", "SEQ Figure \\* alphabetic"), ("S10", "SEQ figure"),
    ("S11", "SEQ Table \\* MERGEFORMAT"), ("S12", "SEQ Figure \\n"), ("S13", "SEQ Equation \\s 1"),
)


def _filler(lines: int, tag: str) -> str:
    return "".join(wml.paragraph(wml.run(f"{tag} filler line {k + 1}"), mark={}) for k in range(lines))


def _cell_table() -> str:
    cells = [[wml.paragraph(wml.run("Cell one"), mark={}),
              wml.paragraph(start("cell") + wml.run("Cell target", **MARKED) + end("cell"), mark={})],
             [wml.paragraph(wml.run("S20[") + field("SEQ Figure", "999", rpr={}, instr_rpr=INSTR, result_rpr=RESULT)
                            + wml.run("]"), mark={}),
              wml.paragraph(wml.run("Cell four"), mark={})]]
    return wml.table(cells, width=3600)


@functools.lru_cache(maxsize=None)
def parts() -> tuple[str, str]:
    body = ""
    # Page 1: alpha, the caption, the multi-paragraph bookmark, the REF cases.
    body += wml.paragraph(wml.run("Alpha: ") + start("alpha") + wml.run("Alpha target", **MARKED) + end("alpha")
                          + wml.run(" after."), mark={})
    body += wml.paragraph(start("caption") + wml.run("Figure ") + simple_field("SEQ Figure \\* ARABIC", "999")
                          + wml.run(" caption text") + end("caption"), mark={})
    body += wml.paragraph(start("styled") + wml.run("Styled target") + end("styled"), mark={}, pStyle="Heading1")
    body += wml.paragraph(wml.run("Run style: ") + start("charstyled") + wml.run("Charstyled target", rStyle="Big")
                          + end("charstyled"), mark={})
    body += wml.paragraph(start("multi") + wml.run("First paragraph of multi"), mark={})
    body += wml.paragraph(wml.run("Second paragraph of multi") + end("multi"), mark={})
    for label, instruction in REFS:
        body += case(label, instruction)
    body += case("R12", "REF alpha", simple=True)
    # Last on its page: Word draws both paragraphs of ``multi``, which moves what follows.
    body += case("R5", "REF multi \\h")
    # Page 2: the table with a bookmarked cell, and the SEQ cases.
    body += wml.paragraph(wml.run("Second page"), mark={}, pageBreakBefore=True)
    body += _cell_table()
    for label, instruction in SEQS:
        body += case(label, instruction)
    body += case("S14", "SEQ Figure", simple=True)
    # Page 3: beta, the PAGEREF cases, a table of contents.
    body += wml.paragraph(start("beta") + wml.run("Beta target", **MARKED) + end("beta"), mark={},
                          pageBreakBefore=True)
    for label, instruction in PAGEREFS:
        body += case(label, instruction)
    toc_begin = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                 '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>'
                 '<w:r><w:fldChar w:fldCharType="separate"/></w:r>')
    body += wml.paragraph(wml.run("T1[") + toc_begin + wml.run("Entry beta") + wml.run("\t")
                          + field("PAGEREF beta \\h", "77", rpr={}, instr_rpr=INSTR, result_rpr=RESULT), mark={})
    body += wml.paragraph(wml.run("Entry gamma") + wml.run("\t")
                          + field("PAGEREF gamma \\h", "77", rpr={}, instr_rpr=INSTR, result_rpr=RESULT)
                          + '<w:r><w:fldChar w:fldCharType="end"/></w:r>' + wml.run("]"), mark={})
    # Page 4: the empty bookmark; then a paragraph from page 4 to page 5 whose last line
    # holds ``late``.
    body += wml.paragraph(wml.run("Fourth page ") + start("empty") + end("empty") + wml.run("empty mark"),
                          mark={}, pageBreakBefore=True)
    body += _filler(24, "Page four")
    words = " ".join(f"word{k}" for k in range(150))
    body += wml.paragraph(wml.run(words + " ") + start("late") + wml.run("Late target", **MARKED) + end("late")
                          + wml.run(" end."), mark={})
    first = story_docx.section()
    body += wml.paragraph(wml.run("Last of section one"), mark={}, sect=first)
    # Section 2: lowerRoman from 10; gamma on its second page.
    body += wml.paragraph(wml.run("Section two, first page"), mark={})
    body += wml.paragraph(start("gamma") + wml.run("Gamma target", **MARKED) + end("gamma"), mark={},
                          pageBreakBefore=True)
    final = story_docx.section(page_numbers={"fmt": "lowerRoman", "start": 10})
    return body, final


def build(setting: str) -> bytes:
    body, final = parts()
    return story_docx.package(body, final, story_docx.Parts(), compatibility_mode=SETTINGS[setting],
                              styles_xml=STYLES)


CAMBRIA = {"ascii": "Cambria", "hAnsi": "Cambria", "eastAsia": "Cambria", "cs": "Cambria"}
COURIER = {"ascii": "Courier New", "hAnsi": "Courier New", "eastAsia": "Courier New", "cs": "Courier New"}
STYLES = story_docx.styles([
    wml.style("paragraph", "Heading1", name="heading 1", based_on="Normal",
              rpr_={"rFonts": CAMBRIA, "sz": 32, "szCs": 32}),
    wml.style("character", "Big", name="Big", rpr_={"rFonts": COURIER, "sz": 32, "szCs": 32}),
])
