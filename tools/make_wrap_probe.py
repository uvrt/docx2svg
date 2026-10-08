#!/usr/bin/env python3
"""Real text, wrapped: does the model break every line after the same word as Word?

Phase 3's done-condition asks for "a wrapping probe of at least 200 lines across five
column widths and five font sizes".  This is it, and it is wider than that on purpose:
Calibri and Aptos alone can hide a rule (ROADMAP.md), so the prose below is set in ten
faces -- Calibri, Arial, Times New Roman, Cambria, Georgia, Verdana, Aptos, Courier New,
Helvetica Neue (a 1000-upm face, whose advances are not whole layout units) and Century
Gothic -- at five sizes (8.5, 10, 11, 13.5 and 17 pt), in five columns (2160, 3517,
5040, 6803 and 9026 twips wide, set by ``w:ind/@w:right`` on an A4 page with 1-inch
margins), the four styles rotating.  Every combination is one paragraph: 250
paragraphs.

The prose is this project's own, written to hold what real text holds: hyphenated
compounds, em and en dashes, figures with separators, parentheses, quotation marks,
slashes, a URL-like run, and long and short words.

Two small families are appended and scored apart: the same text **justified**
(``w:jc="both"``) and **centred** and **right-aligned**, in three faces and five
columns.  Justification is out of Phase 3's scope; whether it changes where Word breaks
is recorded, not modelled.

Measured by ``read_wrap_probe.py`` through ``tools/breaks.py``.
"""

from __future__ import annotations

from dataclasses import dataclass

import probe_docx as w
import wml

FACES = ("Calibri", "Arial", "Times New Roman", "Cambria", "Georgia", "Verdana", "Aptos",
         "Courier New", "Helvetica Neue", "Century Gothic")
SIZES = (17, 20, 22, 27, 34)
COLUMNS = (2160, 3517, 5040, 6803, 9026)
STYLES = ((False, False), (True, False), (False, True), (True, True))
FULL_COLUMN = 9026

PROSE = (
    "A word processor keeps no boxes for its text: only a column, a sequence of runs, and "
    "a rule for where each line must end. The rule is simple to state and hard to "
    "reproduce, because it depends on the width of every glyph to a fraction of a point.",
    "Well-known faces such as Times New Roman, Arial and Calibri differ in more than their "
    "shapes -- their advance widths differ too, so the same sentence (set at 11 pt in a "
    "5.5-inch column) ends its lines after different words in each of them.",
    "In 2024 the committee met on 3 March, 14 June and 29 November; attendance rose from "
    "1,250 to 3,875 people, an increase of 210 per cent over the year-on-year figure "
    "reported by the secretariat in its first-quarter review.",
    "“Measure first,” she said, “and only then decide.” The advice sounds "
    "obvious — almost too obvious to repeat — yet most layout engines guess at "
    "the widths they cannot read, and a guess is wrong on every long line.",
    "The archive lists entries such as report/2023/final-draft, data-set-A/B and "
    "north–south corridors; each path is one unbroken token unless a hyphen or a "
    "slash gives the line breaker somewhere else to go.",
    "Short words: a, an, to, of, in, on, at, by, up, it, is, be, we, so, no, or, if, as. "
    "Long words: characteristically, internationalization, counterrevolutionaries, "
    "incomprehensibilities, and uncharacteristically.",
    "Nothing in the file says where a line breaks. The breaks are computed each time the "
    "document is opened, from the fonts on the machine that opens it, which is why one "
    "document can reflow differently on two computers with different faces installed.",
    "Prices were quoted as $12.50, €9.99 and £7.25 (all excluding tax), while the "
    "discounts ran from 5% to 35% depending on volume, region and the terms agreed in the "
    "master contract signed last spring.",
    "Sometimes the last word of a line is followed by several spaces, and sometimes a line "
    "ends exactly at the margin; both cases must come out the same way in the copy as "
    "they did in the original, or every line below them moves.",
    "Hyphenated compounds — state-of-the-art, mother-in-law, twenty-five, "
    "well-to-do, one-to-one — are where breakers disagree most often, since some "
    "break after each hyphen and others treat the whole compound as one word.",
    "Quality, quantity and quickly queued queries quietly quadrupled; zealous zebras "
    "zigzagged; jovial jugglers juggled jam jars; wavy waves washed westward while vivid "
    "violets vanished very visibly.",
    "The final paragraph is plain: it has no figures, no dashes and no quotation marks, "
    "only ordinary words of ordinary length, separated by single spaces and ending with "
    "a full stop, the way most text in most documents does.",
)
ALIGNED_FACES = ("Calibri", "Times New Roman", "Georgia")


@dataclass(frozen=True)
class Case:
    face: str
    half_points: int
    column: int  # twips
    bold: bool
    italic: bool
    text: str
    jc: str | None = None


def cases() -> list[Case]:
    out = []
    index = 0
    for face in FACES:
        for hp in SIZES:
            for column in COLUMNS:
                bold, italic = STYLES[index % len(STYLES)]
                out.append(Case(face, hp, column, bold, italic, PROSE[index % len(PROSE)]))
                index += 1
    for jc in ("both", "center", "right"):
        for face in ALIGNED_FACES:
            for column in COLUMNS:
                out.append(Case(face, 22, column, False, False, PROSE[index % len(PROSE)], jc))
                index += 1
    return out


def _paragraph(case: Case) -> str:
    fonts = {"ascii": case.face, "hAnsi": case.face, "eastAsia": case.face, "cs": case.face}
    props = {"rFonts": fonts, "sz": case.half_points, "szCs": case.half_points}
    if case.bold:
        props["b"] = True
    if case.italic:
        props["i"] = True
    run = wml.run(case.text, **props)
    ppr = {"spacing": {"before": 0, "after": 120, "line": 240, "lineRule": "auto"},
           "ind": {"left": 0, "right": FULL_COLUMN - case.column}}
    if case.jc:
        ppr["jc"] = case.jc
    return wml.paragraph(run, mark=props, **ppr)


def build() -> bytes:
    return w.package("".join(_paragraph(case) for case in cases()))


if __name__ == "__main__":
    print(len(cases()), "paragraphs,", len(build()), "bytes")
