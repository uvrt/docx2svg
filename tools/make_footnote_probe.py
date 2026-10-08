#!/usr/bin/env python3
"""What a footnote reference does to the page break: the room its note takes.

Measured before footnotes are implemented at all (ROADMAP.md, Phase 4: Word reserves the
footnote area *before* choosing the break).  The layout is ``make_page_fit_probe.py``'s:
every case is a page -- a plain anchor with ``w:pageBreakBefore``, filler composed to the
layout unit, and a candidate line whose pitch ends ``slack`` units above the bottom
margin -- but the candidate (or the anchor) carries footnote references, and the notes
are at the foot of the page.  Whether the candidate stays on its page, against the slack,
gives the room the notes take, to the unit.

The notes' formatting is stated in full, so their height is known: every note paragraph
is ``FootnoteText`` -- Calibri 10 pt (50,000 units a line), no spacing, single -- and the
separator and continuation separator are one paragraph each, Calibri 11 pt, no spacing.
The reference is ``FootnoteReference`` (superscript), at the end of its line.

Families (``FAMILIES``), each swept over ``SLACKS`` (coarse) or about a threshold by
``DELTAS`` (fine):

* ``one`` -- one note of one line, referenced from the candidate.
* ``two-lines`` -- one note of two lines (``w:br``).
* ``two-notes`` -- two one-line notes, both referenced from the candidate.
* ``earlier`` -- one one-line note referenced from the anchor; the candidate is plain.
* ``five-lines`` -- one note of five lines: must all of it fit, or its first line?
* ``twelve`` -- one note of one line at 12 pt.
* ``spaced`` -- one one-line note whose paragraph has 120 twips before and after.
* then fine sweeps of notes with 120 before and after, 120 before and 240 after, and
  240 before and 120 after.

Then, by single units about what the coarse sweep pointed to (``DELTAS``): each of the
above, notes with space only before or only after, a five-line note with its widow
control off, and ``carry`` -- a note split over two pages followed by 60 lines, to see how
many the next page holds under the rest of the note.

One document per compatibility setting (no ``settings.xml``; modes 12, 14 and 15), and
one in mode 15 whose separators are 20 pt.
"""

from __future__ import annotations

import functools
from dataclasses import dataclass

import probe_docx
import wml
from make_page_fit_probe import AVAILABLE, HEIGHT, LINE, MARGIN, WIDTH, compose, twips_units

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
RUN = {"rFonts": FACE, "sz": 22, "szCs": 22}
SPACING = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}
NOTE_LINE = 2500 * 20  # Calibri 10 pt

#: The coarse sweep: slack from below the margin to four notes' worth above it.
SLACKS = tuple(range(-10000, 250001, 10000))
DELTAS = (-1024, -205, -41, -9, -3, -2, -1, 0, 1, 2, 3, 9, 41, 205, 1024)


@dataclass(frozen=True)
class Family:
    name: str
    #: Per note: its lines, its paragraph's space before and after, its size, and
    #: optionally its widow control.
    notes: tuple[tuple, ...]
    #: Where the references are: in the ``candidate`` line or the ``anchor``.
    where: str = "candidate"
    #: Fine thresholds (slack values) to sweep with ``DELTAS``; empty: the coarse sweep.
    thresholds: tuple[int, ...] = ()
    #: Lines of a plain paragraph (widow control off) after the candidate.
    follow: int = 0


SEPARATOR = LINE  # the separator paragraph: Calibri 11 pt, no spacing
ONE = SEPARATOR + NOTE_LINE
S120 = twips_units(120)

FAMILIES = (
    # The coarse sweep: where the threshold is, to 10,000 units.
    Family("one", ((1, 0, 0, 20),)),
    Family("two-lines", ((2, 0, 0, 20),)),
    Family("two-notes", ((1, 0, 0, 20), (1, 0, 0, 20))),
    Family("earlier", ((1, 0, 0, 20),), where="anchor"),
    Family("five-lines", ((5, 0, 0, 20),)),
    Family("twelve", ((1, 0, 0, 24),)),
    Family("spaced", ((1, 120, 120, 20),)),
    # The fine sweep, by single units about what the coarse one pointed to: the
    # separator's line and every note line, whole (two lines of a longer note).
    Family("one fine", ((1, 0, 0, 20),), thresholds=(ONE,)),
    Family("two-lines fine", ((2, 0, 0, 20),), thresholds=(ONE + NOTE_LINE,)),
    Family("two-notes fine", ((1, 0, 0, 20), (1, 0, 0, 20)), thresholds=(ONE + NOTE_LINE,)),
    Family("earlier fine", ((1, 0, 0, 20),), where="anchor", thresholds=(ONE,)),
    Family("five-lines fine", ((5, 0, 0, 20),), thresholds=(ONE + NOTE_LINE,)),
    Family("twelve fine", ((1, 0, 0, 24),), thresholds=(SEPARATOR + 2500 * 24,)),
    Family("before", ((1, 120, 0, 20),), thresholds=(ONE, ONE + twips_units(120))),
    Family("after", ((1, 0, 120, 20),), thresholds=(ONE, ONE + twips_units(120))),
    Family("five-lines no widow", ((5, 0, 0, 20, False),), thresholds=(ONE, ONE + NOTE_LINE)),
    # A third round: space both before and after, in three proportions -- the second
    # showed each alone counting in mode 15, and the coarse sweep both counting once.
    Family("spaced fine", ((1, 120, 120, 20),), thresholds=(ONE, ONE + S120, ONE + 2 * S120)),
    Family("before 120 after 240", ((1, 120, 240, 20),), thresholds=(ONE + S120, ONE + 2 * S120, ONE + 3 * S120)),
    Family("before 240 after 120", ((1, 240, 120, 20),), thresholds=(ONE + S120, ONE + 2 * S120, ONE + 3 * S120)),
    # A note split over two pages, then 60 lines: how many the next page holds under the
    # rest of the note and the continuation separator.
    Family("carry", ((5, 0, 0, 20),), thresholds=(ONE + NOTE_LINE + 5000,), follow=60),
    Family("carry none", (), thresholds=(5000,), follow=60),
)

#: name -> (compatibility mode, the separators' size in half points).
SETTINGS = {"none": (None, 22), "12": (12, 22), "14": (14, 22), "15": (15, 22), "15-sep20": (15, 40)}


@dataclass(frozen=True)
class Case:
    family: Family
    slack: int

    @property
    def key(self) -> str:
        return f"{self.family.name} {self.slack:+d}"


def _cases() -> tuple[Case, ...]:
    out = []
    for family in FAMILIES:
        if family.follow:
            out += [Case(family, t) for t in family.thresholds]
        elif family.thresholds:
            out += [Case(family, t + d) for t in family.thresholds for d in DELTAS]
        else:
            out += [Case(family, s) for s in SLACKS]
    return tuple(out)


CASES = _cases()


def _text(text: str, props: dict = RUN) -> str:
    return f'<w:r>{wml.rpr(**props)}<w:t xml:space="preserve">{text}</w:t></w:r>'


def _reference(note: int) -> str:
    return (f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
            f'<w:footnoteReference w:id="{note}"/></w:r>')


@functools.lru_cache(maxsize=1)
def layout() -> tuple[list[tuple[int, str, str]], list[str]]:
    """``(case number, role, w:p)`` for every body paragraph, and every note's XML."""
    body: list[tuple[int, str, str]] = []
    notes: list[str] = []
    for number, case in enumerate(CASES):
        refs = ""
        for lines, before, after, size, *widow in case.family.notes:
            note_id = len(notes) + 1
            run = {"rFonts": FACE, "sz": size, "szCs": size}
            texts = "<w:br/>".join(f'<w:t xml:space="preserve">n{number}.{note_id}.{k}</w:t>' for k in range(lines))
            widow_xml = '<w:widowControl w:val="0"/>' if widow and not widow[0] else ""
            paragraph = (
                f'<w:p><w:pPr><w:pStyle w:val="FootnoteText"/>{widow_xml}'
                f'<w:spacing w:before="{before}" w:after="{after}" w:line="240" w:lineRule="auto"/>'
                f"{wml.rpr(**run)}</w:pPr>"
                f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
                f"<w:r>{wml.rpr(**run)}<w:t xml:space=\"preserve\"> </w:t>{texts}</w:r></w:p>"
            )
            notes.append(f'<w:footnote w:id="{note_id}">{paragraph}</w:footnote>')
            refs += _reference(note_id)
        anchor_refs = refs if case.family.where == "anchor" else ""
        candidate_refs = refs if case.family.where == "candidate" else ""
        n, a, b, t = compose(AVAILABLE - LINE - LINE - case.slack)
        body.append((number, "anchor", wml.paragraph(_text(f"Case {number} anchor") + anchor_refs, mark=RUN,
                                                     pageBreakBefore=True, spacing=SPACING)))
        if n:
            lines = "<w:br/>".join('<w:t xml:space="preserve">f</w:t>' for _ in range(n))
            body.append((number, "filler", wml.paragraph(f"<w:r>{wml.rpr(**RUN)}{lines}</w:r>", mark=RUN,
                                                         spacing=SPACING)))
        for face, size, text in (("Times New Roman", a, "a"), ("Calibri", b, "b")):
            run = {"rFonts": {"ascii": face, "hAnsi": face, "eastAsia": face, "cs": face}, "sz": size, "szCs": size}
            body.append((number, "filler", wml.paragraph(_text(text, run), mark=run, spacing=SPACING)))
        body.append((number, "filler", wml.paragraph(
            _text("t"), mark=RUN, spacing={"before": 0, "after": 0, "line": t, "lineRule": "exact"})))
        body.append((number, "candidate", wml.paragraph(_text(f"Case {number} cand") + candidate_refs, mark=RUN,
                                                        spacing=SPACING)))
        if case.family.follow:
            lines = "<w:br/>".join(f'<w:t xml:space="preserve">Case {number} follow {k}</w:t>'
                                   for k in range(case.family.follow))
            body.append((number, "follow", wml.paragraph(f"<w:r>{wml.rpr(**RUN)}{lines}</w:r>", mark=RUN,
                                                         spacing=SPACING, widowControl=False)))
    return body, notes


FOOTNOTES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
FOOTNOTES_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"


def _separator(kind: str, note_id: int, size: int = 22) -> str:
    run = {"rFonts": FACE, "sz": size, "szCs": size}
    return (f'<w:footnote w:type="{kind}" w:id="{note_id}"><w:p><w:pPr>'
            f'<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>{wml.rpr(**run)}</w:pPr>'
            f"<w:r>{wml.rpr(**run)}<w:{kind}/></w:r></w:p></w:footnote>")


def build(setting: str) -> bytes:
    body, notes = layout()
    mode, separator = SETTINGS[setting]
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True),
         wml.style("paragraph", "FootnoteText", based_on="Normal",
                   ppr_={"spacing": SPACING}, rpr_={"sz": 20, "szCs": 20}),
         wml.style("character", "FootnoteReference", rpr_={"vertAlign": "superscript"})],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": SPACING},
    )
    footnotes = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:footnotes xmlns:w="{probe_docx.W_NS}">'
        + _separator("separator", -1, separator) + _separator("continuationSeparator", 0, separator)
        + "".join(notes) + "</w:footnotes>"
    )
    extra = [("word/footnotes.xml", FOOTNOTES_CONTENT_TYPE, FOOTNOTES_REL, footnotes)]
    if mode is not None:
        settings = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:settings xmlns:w="{probe_docx.W_NS}">'
            '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word"'
            f' w:val="{mode}"/></w:compat><w:themeFontLang w:val="en-GB"/></w:settings>'
        )
        extra.append(("word/settings.xml", wml.SETTINGS_CONTENT_TYPE, wml.SETTINGS_REL, settings))
    section = (
        "<w:sectPr>"
        f'<w:pgSz w:w="{WIDTH}" w:h="{HEIGHT}"/>'
        f'<w:pgMar w:top="{MARGIN}" w:right="{MARGIN}" w:bottom="{MARGIN}" w:left="{MARGIN}"'
        ' w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>'
    )
    return probe_docx.package("".join(xml for *_, xml in body), final_section=section, styles=styles,
                              extra_parts=tuple(extra))
