#!/usr/bin/env python3
"""Does a paragraph mark larger than its line's text make the line taller?

The model put the paragraph mark on its paragraph's lines like any other item
(``vertical.line_extent``): its ascent and descent took part.  Every earlier probe gave
the mark the text's face and size, and in an empty paragraph the mark is all there is,
so neither could tell whether a mark *larger* than the text of a non-empty line counts.
A table-heavy template exposed that case -- a title line of 24 pt text whose mark is
26 pt, drawn where the 24 pt text alone puts it (ROADMAP.md, "A paragraph mark larger
than its text") -- and this probe isolates it.

One case per page: a plain 11 pt anchor line with ``w:pageBreakBefore``, the case
paragraph, then three plain 11 pt lines, so both the case paragraph's baselines and the
pitch below it are scored.  Single spacing, no space before or after, on everything but
the case paragraph.

**The first two documents** (``mark-none``, ``mark-15``), Calibri throughout, single
spacing:

* ``same``: text and mark 24 pt (the control);
* ``mark-N``: text 24 pt, mark N half points, N over 26 / 28 / 30 pt;
* ``mark-N-space``: the same with the text ending in a space run, as the template's did;
* ``mark-N-bold``: text and mark bold;
* ``small-mark``: text 24 pt, mark 11 pt;
* ``empty-N``: an empty paragraph whose mark is N (the mark is the line: the control
  that it counts there).

**The sweep** (``mark-sweep-<setting>``) settles what those two left open.  Every case
under five line rules on the case paragraph -- ``auto`` 240 and 276, ``exact`` 720
(36 pt, taller than anything here), ``atLeast`` 640 (32 pt: between the natural height
of 24 pt text and of a 30 pt mark, in every face used) and ``atLeast`` 300 (below
both) -- in four settings: no ``settings.xml``, compatibility modes 12, 14 and 15.

* **size**, in six text faces (Calibri, Aptos, Times New Roman, Arial, Cambria, Courier
  New), text 24 pt: the mark at 24 (control), 30 and 16 pt; a paragraph of one space
  run whose mark is 30 pt; an empty paragraph whose mark is 30 pt; and paragraphs of
  about three lines whose mark is 24 or 30 pt.
* **face**, the mark at the *same* size as the text or smaller, in a face that reaches
  further: Arial Black (ascent 1.101 em) under Courier New (0.833) and Calibri (0.952)
  text, Palatino Linotype under Times New Roman; deeper: Lucida Calligraphy (descent
  0.325) under Comic Sans MS (0.292), Bradley Hand (0.399, ascent 0.850) under Arial
  (0.212, 0.938); Arial Black at 20 pt under Courier New 24 pt (a smaller size, a
  taller ascent), Bradley Hand at 20 pt under Arial 24 pt (smaller, deeper); Courier
  New under Arial Black (the mark shorter both ways); a three-line Courier New
  paragraph with an Arial Black mark; an empty Arial Black paragraph.
* **space**, in Calibri, Times New Roman and Courier New, what a line of whitespace is:
  a paragraph of one 30 pt space under a 24 pt mark; of one 26 pt space under a 16 pt
  mark (26, not 30: ``baselines.merge_raised`` would fold the 11 pt line below into a
  30 pt space's); of one 16 pt space under a 24 pt mark; of one 30 pt tab under a 24 pt
  mark; and 24 pt text with a 30 pt space after it, or a 30 pt space or tab inside it;
  a 24 pt space in Arial Black (0.15-0.27 em taller than the text's face), alone under
  a 24 pt mark and between two words; a 30 pt no-break space (U+00A0), alone under a
  24 pt mark and between two words.
* **label**, numbered paragraphs (``%n.``) of 24 pt Calibri and Courier New text: the
  label takes the mark's properties, so a 30 pt mark makes a 30 pt label; a level
  ``w:sz`` of 24 pt keeps the label at the text's size under a 30 pt mark; a level
  ``w:sz`` of 30 pt makes a large label under a 24 pt mark; each on one line and on
  three.

Measured by ``read_mark_probe.py`` through ``baselines.predict``.
"""

from __future__ import annotations

from dataclasses import dataclass

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
TEXT = 48  # half points
MARKS = (52, 56, 60)
SETTINGS = {"none": None, "15": 15}
SWEEP_SETTINGS = {"none": None, "12": 12, "14": 14, "15": 15}
RULES = (("auto", 240), ("auto", 276), ("exact", 720), ("atLeast", 640), ("atLeast", 300))
SIZE_FACES = ("Calibri", "Aptos", "Times New Roman", "Arial", "Cambria", "Courier New")
LABEL_FACES = ("Calibri", "Courier New")
SPACE_FACES = ("Calibri", "Times New Roman", "Courier New")

#: About three lines of 24 pt text on an A4 page with 1 inch margins, in every face.
LONG = ("Hxample paragraph long enough to wrap onto a second line and then onto a third "
        "line of text here")


@dataclass(frozen=True)
class Case:
    name: str
    text: int | None  # half points of the text; None for an empty paragraph
    mark: int
    space: bool = False
    bold: bool = False
    text_face: str = "Calibri"
    mark_face: str | None = None  # None: the text's
    long: bool = False
    only_space: bool = False  # the paragraph is one space run
    level: int | None = None  # numbered, at this level (0: the label takes the mark's size)
    rule: tuple[str, int] = ("auto", 240)
    #: Explicit runs, ``(text, half points)`` in the text's face or ``(text, half points,
    #: face)``, in place of the text.
    runs: tuple[tuple, ...] = ()


CASES = (
    Case("same", TEXT, TEXT),
    Case("small-mark", TEXT, 22),
    *(Case(f"mark-{m}", TEXT, m) for m in MARKS),
    *(Case(f"mark-{m}-space", TEXT, m, space=True) for m in MARKS),
    *(Case(f"mark-{m}-bold", TEXT, m, bold=True) for m in MARKS[:2]),
    *(Case(f"empty-{m}", None, m) for m in MARKS[:2]),
)

#: Label levels: the label's ``w:sz`` (None: the mark's).
LEVELS = (None, 48, 60)


def _short(face: str) -> str:
    return face.replace(" ", "")


def _sweep_cases() -> tuple[Case, ...]:
    out = []
    for rule in RULES:
        tag = f"{rule[0]}{rule[1]}"
        for face in SIZE_FACES:
            f = _short(face)
            out += [
                Case(f"size-same/{f}/{tag}", TEXT, TEXT, text_face=face, rule=rule),
                Case(f"size-big/{f}/{tag}", TEXT, 60, text_face=face, rule=rule),
                Case(f"size-small/{f}/{tag}", TEXT, 32, text_face=face, rule=rule),
                Case(f"size-onlyspace/{f}/{tag}", TEXT, 60, text_face=face, only_space=True, rule=rule),
                Case(f"size-empty/{f}/{tag}", None, 60, text_face=face, rule=rule),
                Case(f"size-multisame/{f}/{tag}", TEXT, TEXT, text_face=face, long=True, rule=rule),
                Case(f"size-multibig/{f}/{tag}", TEXT, 60, text_face=face, long=True, rule=rule),
            ]
        out += [
            Case(f"face-tall/CourierNew-ArialBlack/{tag}", TEXT, TEXT, text_face="Courier New",
                 mark_face="Arial Black", rule=rule),
            Case(f"face-tall/Calibri-ArialBlack/{tag}", TEXT, TEXT, text_face="Calibri",
                 mark_face="Arial Black", rule=rule),
            Case(f"face-tall/TimesNewRoman-PalatinoLinotype/{tag}", TEXT, TEXT,
                 text_face="Times New Roman", mark_face="Palatino Linotype", rule=rule),
            Case(f"face-deep/ComicSansMS-LucidaCalligraphy/{tag}", TEXT, TEXT,
                 text_face="Comic Sans MS", mark_face="Lucida Calligraphy", rule=rule),
            Case(f"face-deep/Arial-BradleyHand/{tag}", TEXT, TEXT, text_face="Arial",
                 mark_face="Bradley Hand", rule=rule),
            Case(f"face-smallertall/CourierNew-ArialBlack/{tag}", TEXT, 40, text_face="Courier New",
                 mark_face="Arial Black", rule=rule),
            Case(f"face-smallerdeep/Arial-BradleyHand/{tag}", TEXT, 40, text_face="Arial",
                 mark_face="Bradley Hand", rule=rule),
            Case(f"face-short/ArialBlack-CourierNew/{tag}", TEXT, TEXT, text_face="Arial Black",
                 mark_face="Courier New", rule=rule),
            Case(f"face-multitall/CourierNew-ArialBlack/{tag}", TEXT, TEXT, text_face="Courier New",
                 mark_face="Arial Black", long=True, rule=rule),
            Case(f"face-empty/ArialBlack/{tag}", None, TEXT, text_face="Courier New",
                 mark_face="Arial Black", rule=rule),
        ]
        for face in SPACE_FACES:
            f = _short(face)
            out += [
                Case(f"space-onlybig/{f}/{tag}", TEXT, TEXT, text_face=face, runs=((" ", 60),), rule=rule),
                Case(f"space-onlybig-smallmark/{f}/{tag}", TEXT, 32, text_face=face, runs=((" ", 52),),
                     rule=rule),
                Case(f"space-onlysmall/{f}/{tag}", TEXT, TEXT, text_face=face, runs=((" ", 32),), rule=rule),
                Case(f"space-trailingbig/{f}/{tag}", TEXT, TEXT, text_face=face,
                     runs=(("Hxample title", TEXT), (" ", 60)), rule=rule),
                Case(f"space-interiorbig/{f}/{tag}", TEXT, TEXT, text_face=face,
                     runs=(("Hxample", TEXT), (" ", 60), ("title", TEXT)), rule=rule),
                Case(f"space-onlytab/{f}/{tag}", TEXT, TEXT, text_face=face, runs=(("\t", 60),), rule=rule),
                Case(f"space-interiortab/{f}/{tag}", TEXT, TEXT, text_face=face,
                     runs=(("Hxample", TEXT), ("\t", 60), ("title", TEXT)), rule=rule),
                Case(f"space-onlytall/{f}/{tag}", TEXT, TEXT, text_face=face,
                     runs=((" ", TEXT, "Arial Black"),), rule=rule),
                Case(f"space-interiortall/{f}/{tag}", TEXT, TEXT, text_face=face,
                     runs=(("Hxample", TEXT), (" ", TEXT, "Arial Black"), ("title", TEXT)), rule=rule),
                Case(f"space-onlynbsp/{f}/{tag}", TEXT, TEXT, text_face=face,
                     runs=(("\u00a0", 60),), rule=rule),
                Case(f"space-interiornbsp/{f}/{tag}", TEXT, TEXT, text_face=face,
                     runs=(("Hxample", TEXT), ("\u00a0", 60), ("title", TEXT)), rule=rule),
            ]
        for face in LABEL_FACES:
            f = _short(face)
            out += [
                Case(f"label-same/{f}/{tag}", TEXT, TEXT, text_face=face, level=0, rule=rule),
                Case(f"label-bigmark/{f}/{tag}", TEXT, 60, text_face=face, level=0, rule=rule),
                Case(f"label-bigmark-label24/{f}/{tag}", TEXT, 60, text_face=face, level=1, rule=rule),
                Case(f"label-label30/{f}/{tag}", TEXT, TEXT, text_face=face, level=2, rule=rule),
                Case(f"label-multibigmark/{f}/{tag}", TEXT, 60, text_face=face, level=0, long=True,
                     rule=rule),
                Case(f"label-multibigmark-label24/{f}/{tag}", TEXT, 60, text_face=face, level=1,
                     long=True, rule=rule),
                Case(f"label-multilabel30/{f}/{tag}", TEXT, TEXT, text_face=face, level=2, long=True,
                     rule=rule),
            ]
    return tuple(out)


SWEEP_CASES = _sweep_cases()


@dataclass(frozen=True)
class Probe:
    setting: str
    sweep: bool = False

    @property
    def name(self) -> str:
        return f"mark-sweep-{self.setting}" if self.sweep else f"mark-{self.setting}"

    @property
    def cases(self) -> tuple[Case, ...]:
        return SWEEP_CASES if self.sweep else CASES

    @property
    def mode(self) -> int | None:
        return (SWEEP_SETTINGS if self.sweep else SETTINGS)[self.setting]


PROBES = (*(Probe(setting) for setting in SETTINGS),
          *(Probe(setting, sweep=True) for setting in SWEEP_SETTINGS))
SPACING = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}


def _fonts(face: str) -> dict:
    return {"ascii": face, "hAnsi": face, "eastAsia": face, "cs": face}


def _plain(text: str, **props) -> str:
    return wml.paragraph(wml.run(text, sz=22), mark={"sz": 22}, spacing=SPACING, **props)


def _case(case: Case) -> str:
    """The first two documents' case paragraph (unchanged, so their recording holds)."""
    bold = {"b": True} if case.bold else {}
    runs = ""
    if case.text is not None:
        runs = wml.run("Hxample title", sz=case.text, **bold)
        if case.space:
            runs += wml.run(" ", sz=case.text, **bold)
    return wml.paragraph(runs, mark={**bold, "sz": case.mark}, spacing=SPACING)


def _sweep_case(case: Case) -> str:
    text_rpr = {"rFonts": _fonts(case.text_face), "sz": case.text, "szCs": case.text}
    mark_face = case.mark_face or case.text_face
    mark_rpr = {"rFonts": _fonts(mark_face), "sz": case.mark, "szCs": case.mark}
    runs = ""
    if case.runs:
        runs = "".join(
            wml.run(text, **{**text_rpr, "sz": size, "szCs": size,
                             **({"rFonts": _fonts(face[0])} if face else {})})
            for text, size, *face in case.runs)
    elif case.only_space:
        runs = wml.run(" ", **text_rpr)
    elif case.text is not None:
        runs = wml.run(LONG if case.long else "Hxample title", **text_rpr)
    props = {"spacing": {"before": 0, "after": 0, "line": case.rule[1], "lineRule": case.rule[0]}}
    if case.level is not None:
        props["numPr"] = f'<w:ilvl w:val="{case.level}"/><w:numId w:val="1"/>'
    return wml.paragraph(runs, mark=mark_rpr, **props)


def kinds(probe: Probe | None = None) -> list[str]:
    """The case each block belongs to, in ``baselines.blocks`` order."""
    cases = CASES if probe is None else probe.cases
    return [case.name for case in cases for _ in range(5)]


def build(probe: Probe) -> bytes:
    body = []
    for index, case in enumerate(probe.cases):
        body.append(_plain("Anchor", pageBreakBefore=index > 0))
        body.append(_sweep_case(case) if probe.sweep else _case(case))
        body.extend(_plain(f"Hx plain {n}") for n in range(3))
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": SPACING},
    )
    extra = []
    if probe.sweep:
        extra.append(wml.numbering_part([
            ("decimal", f"%{level + 1}.", {"ind": {"left": 720, "hanging": 720}},
             {} if size is None else {"sz": size, "szCs": size})
            for level, size in enumerate(LEVELS)
        ]))
    if probe.mode is not None:
        extra.append(wml.settings_part({"val": "en-GB"}, compatibility_mode=probe.mode))
    return probe_docx.package("".join(body), styles=styles, extra_parts=tuple(extra))
