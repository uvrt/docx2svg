#!/usr/bin/env python3
"""Where Word draws a header's and a footer's glyphs: the geometry of the two stories.

Phase 4.4 measured how much room a header or footer takes from the body.  Drawing them
needs more: where the first line of a header sits below ``w:pgMar/@w:header``, where a
footer's lines sit above ``@w:footer`` (a footer grows upward from there), whether a
story's lines round in their line boxes as the body's do, and what else moves them --
the story's paragraph spacing, line rules, faces, alignment and tab stops, borders, a
table, a picture, the gutter, indents into the margin, and the section's vertical
alignment.  Each case is a section of its own (``nextPage``) with a header and a footer
part of its own and a short body, so every page shows one case.

The cases (every story paragraph in Word's ``Header`` / ``Footer`` style unless said):

* ``plain`` -- one line each at Word's default distances;
* ``distance h/f`` -- ``w:header`` / ``w:footer`` 0, 360, 725, 731, 1000, 1417 twips (725
  and 731 are not whole device px);
* ``face`` -- Times New Roman 14.5, Cambria 9.5, Arial 20, Georgia 7.5 pt;
* ``rule`` -- ``auto`` 276 and 360, ``exact`` 300, ``atLeast`` 400, and space before 120 /
  after 240, in both stories;
* ``lines`` -- two paragraphs; a paragraph that wraps to three lines; a ``w:br``;
* ``tabs`` -- left / centre / right text on Word's stops; centred and right-aligned;
* ``border`` -- a bottom border under the header, a top border over the footer;
* ``table`` -- a two-column table in each story;
* ``picture`` -- an inline picture in each story;
* ``tall`` -- a five-line header and footer, past both margins (the body moves: 4.4);
* ``gutter``, ``indent`` -- a 720-twip gutter; indents of -720 into both margins;
* ``A4``, ``A5`` -- other page heights (a footer is placed from the page's foot);
* ``valign`` -- ``w:vAlign`` center and bottom (does the section's alignment move a story?);
* ``normal style`` -- the stories in ``Normal`` (160 after at 259 auto).

One document per compatibility setting: no ``settings.xml``, and modes 12, 14 and 15.
"""

from __future__ import annotations

import functools
from dataclasses import dataclass, field

import story_docx
import wml

SETTINGS = {"none": None, "12": 12, "14": 14, "15": 15}


def run(text: str, **rpr) -> str:
    return wml.run(text, **rpr)


def face(name: str, half_points: int) -> dict:
    return {"rFonts": {"ascii": name, "hAnsi": name, "eastAsia": name, "cs": name}, "sz": half_points,
            "szCs": half_points}


def p(runs: str, style: str, mark: dict | None = None, **props) -> str:
    return wml.paragraph(runs, mark=mark or {}, pStyle=style, **props)


@dataclass(frozen=True)
class Case:
    name: str
    #: The header's and the footer's blocks, as functions of (case number, style) -> XML.
    header: object = None
    footer: object = None
    section: dict = field(default_factory=dict)
    pictures: bool = False


def one_line(tag: str, **rpr):
    return lambda n, style: p(run(f"{tag}{n} Story line", **rpr), style, mark=rpr or None)


def with_props(tag: str, **props):
    return lambda n, style: p(run(f"{tag}{n} Spaced story"), style, **props)


def two_paragraphs(tag: str):
    return lambda n, style: (p(run(f"{tag}{n} First paragraph"), style)
                             + p(run(f"{tag}{n} Second paragraph"), style))


def wrapping(tag: str):
    words = " ".join(["wrapping"] * 40)
    return lambda n, style: p(run(f"{tag}{n} {words}"), style)


def broken(tag: str):
    return lambda n, style: p(f'<w:r><w:t>{tag}{n} Before</w:t><w:br/><w:t>After break</w:t></w:r>', style)


def tabs(tag: str):
    return lambda n, style: p(run(f"{tag}{n} Left") + run("\t") + run("Centre") + run("\t") + run("Right"), style)


def aligned(tag: str, jc: str):
    return lambda n, style: p(run(f"{tag}{n} Aligned {jc}"), style, jc=jc)


def bordered(tag: str, side: str):
    border = f'<w:{side} w:val="single" w:sz="6" w:space="1" w:color="auto"/>'
    return lambda n, style: p(run(f"{tag}{n} Bordered"), style, pBdr=border)


def table(tag: str):
    def make(n, style):
        cell = lambda text: p(run(text), style)  # noqa: E731
        return wml.table([[cell(f"{tag}{n} Cell one"), cell("Cell two")]], width=4680) + p("", style)
    return make


def picture(tag: str):
    return lambda n, style: p(run(f"{tag}{n} Logo ") + story_docx.picture(n * 2 + (tag == "F"), 914400, 457200),
                              style)


def tall(tag: str):
    return lambda n, style: p("".join(
        f'<w:r><w:t>{tag}{n} Tall line {k}</w:t>{"<w:br/>" if k < 4 else ""}</w:r>' for k in range(5)), style)


def indented(tag: str):
    return lambda n, style: p(run(f"{tag}{n} Into the margins " + "wide " * 20), style,
                              ind={"left": -720, "right": -720})


def _cases() -> tuple[Case, ...]:
    out = [Case("plain", one_line("H"), one_line("F"))]
    for d in (0, 360, 725, 731, 1000, 1417):
        out.append(Case(f"distance {d}", one_line("H"), one_line("F"), {"header": d, "footer": d}))
    for name, hp in (("Times New Roman", 29), ("Cambria", 19), ("Arial", 40), ("Georgia", 15)):
        out.append(Case(f"face {name} {hp}", one_line("H", **face(name, hp)), one_line("F", **face(name, hp))))
    for rule, line in (("auto", 276), ("auto", 360), ("exact", 300), ("atLeast", 400)):
        spacing = {"after": 0, "line": line, "lineRule": rule}
        out.append(Case(f"rule {rule} {line}", with_props("H", spacing=spacing), with_props("F", spacing=spacing)))
    spacing = {"before": 120, "after": 240, "line": 240, "lineRule": "auto"}
    out.append(Case("space 120/240", with_props("H", spacing=spacing), with_props("F", spacing=spacing)))
    out.append(Case("two paragraphs", two_paragraphs("H"), two_paragraphs("F")))
    out.append(Case("wrapping", wrapping("H"), wrapping("F")))
    out.append(Case("break", broken("H"), broken("F")))
    out.append(Case("tabs", tabs("H"), tabs("F")))
    out.append(Case("center", aligned("H", "center"), aligned("F", "center")))
    out.append(Case("right", aligned("H", "right"), aligned("F", "right")))
    out.append(Case("border", bordered("H", "bottom"), bordered("F", "top")))
    out.append(Case("table", table("H"), table("F")))
    out.append(Case("picture", picture("H"), picture("F"), pictures=True))
    out.append(Case("tall", tall("H"), tall("F")))
    out.append(Case("gutter", one_line("H"), one_line("F"), {"gutter": 720}))
    out.append(Case("indent", indented("H"), indented("F")))
    out.append(Case("A4", one_line("H"), one_line("F"), {"width": 11906, "height": 16838}))
    out.append(Case("A5", one_line("H"), one_line("F"), {"width": 8391, "height": 11906, "left": 1134,
                                                          "right": 1134}))
    out.append(Case("valign center", one_line("H"), one_line("F"), {"v_align": "center"}))
    out.append(Case("valign bottom", one_line("H"), one_line("F"), {"v_align": "bottom"}))
    out.append(Case("normal style", one_line("H"), one_line("F"), {"style": "Normal"}))
    out.append(Case("distance 1800 / 1600", one_line("H"), one_line("F"), {"header": 1800, "footer": 1600}))
    return tuple(out)


CASES = _cases()


@functools.lru_cache(maxsize=1)
def parts() -> tuple[str, str, story_docx.Parts]:
    stories = story_docx.Parts()
    body = ""
    final = ""
    for number, case in enumerate(CASES):
        options = dict(case.section)
        style = options.pop("style", None)
        references = ""
        for kind, make in (("header", case.header), ("footer", case.footer)):
            blocks = make(number, style or ("Header" if kind == "header" else "Footer"))
            relationship = stories.add(kind, blocks, pictures=case.pictures)
            references += stories.reference(kind, "default", relationship)
        sect = story_docx.section(references=references, **options)
        text = wml.run(f"Case {number} body text, which is the body of the case {case.name}.")
        if number == len(CASES) - 1:
            body += wml.paragraph(text, mark={})
            final = sect
        else:
            body += wml.paragraph(text, mark={}, sect=sect)
    return body, final, stories


def build(setting: str) -> bytes:
    body, final, stories = parts()
    return story_docx.package(body, final, stories, compatibility_mode=SETTINGS[setting])


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        Path(sys.argv[1] if len(sys.argv) > 1 else ".").joinpath(f"story-{setting}.docx").write_bytes(build(setting))
