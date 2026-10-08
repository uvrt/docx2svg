#!/usr/bin/env python3
"""Where Word wraps a short word, kerned and not: which kern pairs it charges at the edge.

pptx-agent's ``tools/wrap_boundary_probe.py`` found that PowerPoint lays a line out with a
face's legacy ``kern`` table and never with its ``GPOS`` ``kern`` feature: Aptos's ``ss``
(-33/2048 em), which only ``GPOS`` holds, is not charged, so "Pass" at 18 pt breaks
"Pas / s" in a box the ``GPOS``-kerned word fits.  This probe asks Word the same question
with the same words, faces, sizes and boundary widths, so the two applications' rules
(``ooxml_common.drawingml.rules``' ``kerning``) rest on the same evidence:

* ``Pass``, ``Fail``, ``Total``, ``AVAWAY`` and ``Review``, in Aptos, Calibri and Arial, at
  12, 18 and 24 pt -- every word a paragraph of its own;
* each paragraph's text width set by its right indent, in whole twips, to the word's width
  **plain** (``U``), kerned with the **legacy** table (``L``) and kerned with **GPOS**
  (``G``), each less 0.3 pt, less 0.05 pt, plus 0.05 pt and plus 0.3 pt (rounded to the
  twip; duplicates dropped) -- a word too wide for its line is split, so one line or two
  is the verdict;
* every case twice: with ``w:kern w:val="2"`` (kerning on from 1 pt) and without
  ``w:kern``, which is Word's default and, measured in ROADMAP.md 2.3, no kerning at all.

Then, outside the wrap question, one line each of ``AVAWAY`` and ``Pass`` in two
**variable** faces, Noto Sans JP and STIX Two Text, kerned: PowerPoint kerns a variable
face with its ``GPOS`` feature (pptx2svg's ``tools/read_kern_source_probe.py``), and these
say whether Word does.

``tools/read_wrap_kern_probe.py`` reads Word's export.  The widths are computed with
fontTools from the faces Word lays out with, read where they are installed.
"""

from __future__ import annotations

import functools
from dataclasses import dataclass
from pathlib import Path

import probe_docx as w

WORDS = ("Pass", "Fail", "Total", "AVAWAY", "Review")
FACES = ("Aptos", "Calibri", "Arial")
SIZES = (12, 18, 24)
OFFSETS = (-0.3, -0.05, 0.05, 0.3)
#: ``(kern half points or None)``: on from 1 pt, and Word's default.
KERNING = (2, None)
VARIABLE_FACES = ("Noto Sans JP", "STIX Two Text")
VARIABLE_WORDS = ("AVAWAY", "Pass")

PAGE_WIDTH, MARGIN = 11906, 1440
TEXT_WIDTH = PAGE_WIDTH - 2 * MARGIN

_WORD_FONTS = Path("/Applications/Microsoft Word.app/Contents/Resources/DFonts")
#: The regular face Word lays each family out with: the macOS copy first, then Word's
#: bundle (ROADMAP.md 2.4) -- Arial's two copies agree on every advance and pair here.
FILES = {
    "Aptos": _WORD_FONTS / "Aptos.ttf",
    "Calibri": _WORD_FONTS / "Calibri.ttf",
    "Arial": Path("/System/Library/Fonts/Supplemental/Arial.ttf"),
    "Noto Sans JP": Path.home() / "Library/Fonts/NotoSansJP[wght].ttf",
    "STIX Two Text": Path("/System/Library/Fonts/Supplemental/STIXTwoText.ttf"),
}


@dataclass(frozen=True)
class Case:
    face: str
    size: int  # points
    word: str
    kern: int | None
    tag: str  # "L-0.05pt", ...; "width" for a line read for its width alone
    twips: int  # the text width the right indent leaves


@functools.lru_cache(maxsize=None)
def face(family: str):
    """``(cmap, advances, units per em, legacy pairs, gpos)`` of ``family``'s regular face:
    a variable face instanced at 400, the instance a regular run is drawn in."""
    from fontTools.ttLib import TTFont

    font = TTFont(str(FILES[family]))
    if "fvar" in font:
        from fontTools.varLib import instancer

        font = instancer.instantiateVariableFont(font, {"wght": 400})
    legacy: dict = {}
    if "kern" in font:
        for subtable in font["kern"].kernTables:
            for pair, value in getattr(subtable, "kernTable", {}).items():
                legacy.setdefault(pair, value)
    lookups = []
    if "GPOS" in font and font["GPOS"].table.FeatureList is not None:
        table = font["GPOS"].table
        wanted = sorted({i for record in table.FeatureList.FeatureRecord if record.FeatureTag == "kern"
                         for i in record.Feature.LookupListIndex})
        for index in wanted:
            lookup = table.LookupList.Lookup[index]
            group = []
            for subtable in lookup.SubTable:
                kind = lookup.LookupType
                if kind == 9:
                    subtable, kind = subtable.ExtSubTable, subtable.ExtSubTable.LookupType
                if kind == 2:
                    group.append(subtable)
            lookups.append(group)

    def gpos(first: str, second: str) -> int:
        total = 0
        for group in lookups:
            for subtable in group:
                value = _pair(subtable, first, second)
                if value is not None:
                    total += value
                    break
        return total

    return font.getBestCmap(), font["hmtx"].metrics, font["head"].unitsPerEm, legacy, gpos


def _pair(subtable, first: str, second: str) -> int | None:
    coverage = subtable.Coverage.glyphs
    if first not in coverage:
        return None
    if subtable.Format == 1:
        for record in subtable.PairSet[coverage.index(first)].PairValueRecord:
            if record.SecondGlyph == second:
                return getattr(record.Value1, "XAdvance", 0) or 0
        return None
    one = subtable.ClassDef1.classDefs.get(first, 0)
    two = subtable.ClassDef2.classDefs.get(second, 0)
    if one >= subtable.Class1Count or two >= subtable.Class2Count:
        return None
    return getattr(subtable.Class1Record[one].Class2Record[two].Value1, "XAdvance", 0) or 0


def units(family: str, word: str) -> dict:
    """The word's advances and the pairs at its joins, in font units: what the reader and
    the test predict from."""
    cmap, advances, upem, legacy, gpos = face(family)
    glyphs = [cmap[ord(char)] for char in word]
    return {
        "units_per_em": upem,
        "advances": [advances[glyph][0] for glyph in glyphs],
        "legacy": [legacy.get(pair, 0) for pair in zip(glyphs, glyphs[1:])],
        "gpos": [gpos(*pair) for pair in zip(glyphs, glyphs[1:])],
    }


def widths(family: str, word: str, size: float) -> dict[str, float]:
    """The word's width in points: ``U`` plain, ``L`` legacy-kerned, ``G`` GPOS-kerned."""
    found = units(family, word)
    scale = size / found["units_per_em"]
    plain = sum(found["advances"])
    return {"U": plain * scale, "L": (plain + sum(found["legacy"])) * scale,
            "G": (plain + sum(found["gpos"])) * scale}


def cases() -> list[Case]:
    out: list[Case] = []
    for kern in KERNING:
        for family in FACES:
            for size in SIZES:
                for word in WORDS:
                    seen: set[int] = set()
                    for model, width in widths(family, word, size).items():
                        for offset in OFFSETS:
                            twips = round((width + offset) * 20)
                            if twips not in seen:
                                seen.add(twips)
                                out.append(Case(family, size, word, kern, f"{model}{offset:+g}pt", twips))
    for family in VARIABLE_FACES:
        for word in VARIABLE_WORDS:
            out.append(Case(family, 24, word, 2, "width", TEXT_WIDTH))
    return out


def _paragraph(case: Case) -> str:
    half_points = case.size * 2
    run = w.run(case.word, case.face, half_points, kern_half_points=case.kern)
    mark = w.rpr(case.face, half_points, kern_half_points=case.kern)
    return (
        '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
        f'<w:ind w:left="0" w:right="{TEXT_WIDTH - case.twips}"/>{mark}</w:pPr>{run}</w:p>'
    )


def build() -> bytes:
    return w.package("".join(_paragraph(case) for case in cases()),
                     final_section=w.section(PAGE_WIDTH, 16838, MARGIN))
