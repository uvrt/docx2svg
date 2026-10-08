#!/usr/bin/env python3
"""A section's last paragraph at the bottom of a page: does Word move it to the next?

Found on local documents (ROADMAP F.4): a section whose last paragraph -- an empty one,
holding the ``w:sectPr`` -- does not fit on its page is laid out by the model on a page
of its own, and Word makes no such page.  This probe measures where Word puts such a
paragraph, and what decides it.

Every case starts a section on a new page: an anchor line (``Case N anchor``), filler
composed as ``make_page_fit_probe.py`` composes it (exact to the layout unit), and a
**candidate** paragraph whose fit extent (its text's height: an ``auto`` multiple's extra
is not counted, 4.1) ends ``slack`` units above the bottom margin, negative below it.
Slacks near ``-extent`` put the candidate's top at the margin; ``below`` slacks put its
top under the margin, pushed there by the space after of the filler's last paragraph.
What follows the candidate is the case's **kind**:

* ``next-page`` -- the candidate is empty and holds the section's ``w:sectPr``; the next
  section (``Case N next``) starts on a new page.  Its page says whether the candidate
  made a page of its own (two after the anchor's) or not (the one after).
* ``continuous`` -- the same, the next section continuous: its line's page and baseline
  say where the candidate went (the top of the next page if it stayed; a candidate's
  height lower if it moved).
* ``mid`` -- the candidate is empty and holds no ``w:sectPr``; ``Case N next`` follows it
  in its section (the control: an empty paragraph that does not end a section).
* ``text`` -- the candidate holds text (``Case N cand``) and the ``w:sectPr``: its own
  page shows.

The candidate's **family** varies what it is (``FAMILIES``): an 11 pt mark, a 16 pt mark
under ``auto`` 276 (the local documents'), a 48 pt mark, space after, space before, a
table before it, and a pair of empty paragraphs (the first swept, the second holding the
``w:sectPr``).

The document's own last paragraph (the body's ``w:sectPr``) is measured by separate
documents, one case each (``END_CASES``): whether Word makes a second page.

One document per compatibility setting: no ``settings.xml``, mode 14 and mode 15.
"""

from __future__ import annotations

import functools
from dataclasses import dataclass

import probe_docx
import wml
from make_page_fit_probe import (AVAILABLE, CALIBRI, HEIGHT, LINE, MARGIN, WIDTH, _face, _spacing,
                                 compose, twips_units)

SETTINGS = {"none": None, "14": 14, "15": 15}

#: The table family's table: two rows of an exact height, no borders.
TABLE_ROW = 400
TABLE_UNITS = 2 * twips_units(TABLE_ROW)
#: Where a ``below`` slack's filler ends above the margin before its space after.
GAP = 20


@dataclass(frozen=True)
class Family:
    name: str
    half_points: int = 22
    line: int = 240
    before: int = 0
    after: int = 0
    table: bool = False
    pair: bool = False

    @property
    def extent(self) -> int:
        """The candidate's fit extent: its text's height (units)."""
        return CALIBRI * self.half_points

    @property
    def above(self) -> int:
        return twips_units(self.before)


FAMILIES = (
    Family("plain"),
    Family("mark16", half_points=32, line=276),
    Family("mark48", half_points=96),
    Family("after", after=240),
    Family("before", before=240),
    Family("table", table=True),
    Family("pair", pair=True),
)

#: name -> slack as a function of the extent (``None``: below the margin, see BELOW).
SLACKS = {
    "fits": lambda e: 41,
    "over-1": lambda e: -1,
    "over-41": lambda e: -41,
    "over-quarter": lambda e: -(e // 4),
    "top-above": lambda e: -(e - 41),
    "top-at": lambda e: -e,
    "below-12pt": None,
    "below-60pt": None,
}
#: The space after that puts the candidate's top this many twips under the margin.
BELOW = {"below-12pt": 240, "below-60pt": 1200}

KINDS = ("next-page", "continuous", "mid", "text")


@dataclass(frozen=True)
class Case:
    family: Family
    slack: str
    kind: str

    @property
    def key(self) -> str:
        return f"{self.family.name} {self.slack} {self.kind}"

    @property
    def valid(self) -> bool:
        if self.slack in BELOW and self.family.table:
            return False  # a table has no space after to push the candidate down with
        if self.family.pair and self.kind != "continuous":
            return False  # only the next line's baseline tells one of two from both
        return True


CASES = tuple(c for c in (Case(f, s, k) for f in FAMILIES for s in SLACKS for k in KINDS) if c.valid)

#: The document's own last paragraph: one document per case.
END_CASES = tuple(Case(f, s, "end") for f in FAMILIES[:2] + (FAMILIES[5],)
                  for s in ("fits", "over-41", "top-above", "top-at", "below-12pt")
                  if Case(f, s, "end").valid)

RUN = {"rFonts": _face("Calibri"), "sz": 22, "szCs": 22}


def _p(text: str, sect: str = "", **props) -> str:
    props.setdefault("spacing", _spacing())
    return wml.paragraph(wml.run(text, **RUN) if text else "", mark=RUN, sect=sect, **props)


def _section(kind: str | None) -> str:
    return ("<w:sectPr>" + (f'<w:type w:val="{kind}"/>' if kind else "")
            + f'<w:pgSz w:w="{WIDTH}" w:h="{HEIGHT}"/>'
            f'<w:pgMar w:top="{MARGIN}" w:right="{MARGIN}" w:bottom="{MARGIN}" w:left="{MARGIN}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>')


def _table(number: int) -> str:
    cell = _p(f"t{number}")
    rows = "".join(
        f'<w:tr><w:trPr><w:trHeight w:val="{TABLE_ROW}" w:hRule="exact"/></w:trPr>'
        f'<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{cell}</w:tc></w:tr>'
        for _ in range(2))
    return ('<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
            '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
            '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
            '<w:tblLook w:val="0000"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
            f"{rows}</w:tbl>")


def _candidate(case: Case, number: int, sect: str) -> str:
    family = case.family
    mark = {"rFonts": _face("Calibri"), "sz": family.half_points, "szCs": family.half_points}
    runs = wml.run(f"Case {number} cand", **mark) if case.kind == "text" else ""
    return wml.paragraph(runs, mark=mark, sect=sect,
                         spacing=_spacing(before=family.before, after=family.after, line=family.line))


def case_blocks(number: int, case: Case, section_end: str, next_section: str) -> list[tuple[str, str]]:
    """``(role, xml)`` for one case: ``section_end`` is the ``w:sectPr`` that ends the
    case's first section (the anchor's), ``next_section`` the one that ends the ``next``
    line's (``""``: the body's own)."""
    family = case.family
    before_candidate = family.above + (TABLE_UNITS if family.table else 0)
    push = 0
    if SLACKS[case.slack] is None:
        # The filler ends GAP twips above the margin and its space after reaches BELOW past it.
        total = AVAILABLE - LINE - twips_units(GAP)
        push = GAP + BELOW[case.slack]
    else:
        total = AVAILABLE - LINE - before_candidate - family.extent - SLACKS[case.slack](family.extent)
    n, a, b, t = compose(total)
    out = [("anchor", _p(f"Case {number} anchor"))]
    if n:
        lines = "<w:br/>".join('<w:t xml:space="preserve">f</w:t>' for _ in range(n))
        out.append(("filler", wml.paragraph(f"<w:r>{wml.rpr(**RUN)}{lines}</w:r>", mark=RUN,
                                            spacing=_spacing())))
    face = {"rFonts": _face("Times New Roman"), "sz": a, "szCs": a}
    out.append(("filler", wml.paragraph(wml.run(f"f{number}.a", **face), mark=face, spacing=_spacing())))
    big = {"rFonts": _face("Calibri"), "sz": b, "szCs": b}
    out.append(("filler", wml.paragraph(wml.run(f"f{number}.b", **big), mark=big, spacing=_spacing())))
    out.append(("filler", _p(f"f{number}.t", spacing=_spacing(line=t, rule="exact", after=push))))
    if family.table:
        out.append(("table", _table(number)))
    ends_here = case.kind in ("next-page", "continuous", "text")
    if family.pair:
        out.append(("candidate", _candidate(case, number, "")))
        out.append(("second", _p("", sect=section_end if ends_here else "")))
    else:
        out.append(("candidate", _candidate(case, number, section_end if ends_here else "")))
    if case.kind == "end":
        return out
    out.append(("next", _p(f"Case {number} next", sect=next_section if ends_here else section_end)))
    return out


@functools.lru_cache(maxsize=1)
def blocks() -> list[tuple[str, str]]:
    out = []
    for number, case in enumerate(CASES):
        last = number == len(CASES) - 1
        # The anchor's section starts on a new page (the one before it ends with nextPage
        # or its next line's section does); the next line's section is continuous for the
        # continuous kind, a new page otherwise.
        first_kind = None if number == 0 else "nextPage"
        next_kind = "continuous" if case.kind == "continuous" else "nextPage"
        if case.kind == "mid":
            out += case_blocks(number, case, "" if last else _section(first_kind), "")
        else:
            out += case_blocks(number, case, _section(first_kind), "" if last else _section(next_kind))
    return out


def final_section() -> str:
    case = CASES[-1]
    if case.kind in ("mid",):
        return _section(None if len(CASES) == 1 else "nextPage")
    return _section("continuous" if case.kind == "continuous" else "nextPage")


def _styles() -> str:
    return wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults=RUN, paragraph_defaults={"spacing": _spacing()},
    )


def _settings(setting: str) -> tuple:
    mode = SETTINGS[setting]
    return () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)


def build(setting: str) -> bytes:
    body = "".join(xml for _, xml in blocks())
    return probe_docx.package(body, final_section=final_section(), styles=_styles(),
                              extra_parts=_settings(setting))


def end_blocks(case: Case) -> list[tuple[str, str]]:
    return case_blocks(0, case, "", "")


def build_end(case: Case, setting: str) -> bytes:
    body = "".join(xml for _, xml in end_blocks(case))
    return probe_docx.package(body, final_section=_section(None), styles=_styles(),
                              extra_parts=_settings(setting))


if __name__ == "__main__":
    print(len(CASES), "cases,", len(END_CASES), "end documents per setting")
