#!/usr/bin/env python3
"""Where Word draws line numbers (``w:lnNumType``), and which lines it counts.

Sections (``nextPage``), each two pages or one, with:

* ``plain`` -- ``w:countBy`` 1, no distance (auto), restart at each page (unstated);
* ``every5`` -- ``w:countBy`` 5, ``w:start`` 3, ``w:restart`` ``continuous``, ``w:distance``
  720;
* ``section`` -- ``w:countBy`` 2, ``w:restart`` ``newSection``, ``w:distance`` 200;
* ``start`` -- ``w:start`` 4, ``w:restart`` ``newSection``;
* ``none`` -- no ``w:lnNumType``.

Their text: wrapped paragraphs, short ones, an empty paragraph, a paragraph of 20 pt, a
centred and a right-aligned one, one indented 720 twips, one with
``w:suppressLineNumbers``, a two-row table, a header and a footer.  Calibri 11 in a Word
16 ``Normal``.  No ``settings.xml`` and mode 15.
"""

from __future__ import annotations

import functools

import story_docx
import wml

SETTINGS = {"none": None, "15": 15}

WORDS = ("Numbered lines are counted by Word in the margin beside the text of the body, and this "
         "paragraph is long enough to wrap onto a second and a third line of the page so that each of "
         "them gets a number of its own.")


def _numbering(kind: str) -> str:
    if kind == "plain":
        return '<w:lnNumType w:countBy="1"/>'
    if kind == "every5":
        return '<w:lnNumType w:countBy="5" w:start="3" w:distance="720" w:restart="continuous"/>'
    if kind == "section":
        return '<w:lnNumType w:countBy="2" w:distance="200" w:restart="newSection"/>'
    if kind == "start":
        return '<w:lnNumType w:countBy="1" w:start="4" w:restart="newSection"/>'
    return ""


def _section(kind: str, stories: story_docx.Parts) -> str:
    refs = (stories.reference("header", "default", stories.add("header", wml.paragraph(
                wml.run(f"Header {kind}"), mark={}, pStyle="Header")))
            + stories.reference("footer", "default", stories.add("footer", wml.paragraph(
                wml.run(f"Footer {kind}"), mark={}, pStyle="Footer"))))
    sect = story_docx.section(references=refs)
    number = _numbering(kind)
    # CT_SectPr: lnNumType after pgMar (and pgNumType), before cols.
    return sect.replace('<w:cols ', number + '<w:cols ', 1) if number else sect


def _content(kind: str) -> str:
    out = wml.paragraph(wml.run(f"Section {kind}: first line"), mark={})
    out += wml.paragraph(wml.run(WORDS), mark={})
    out += wml.paragraph("", mark={})
    out += wml.paragraph(wml.run("A short line"), mark={})
    out += wml.paragraph(wml.run("Twenty point line", sz=40, szCs=40), mark={})
    out += wml.paragraph(wml.run("Centred line"), mark={}, jc="center")
    out += wml.paragraph(wml.run("Right-aligned line"), mark={}, jc="right")
    out += wml.paragraph(wml.run("Indented line"), mark={}, ind={"left": 720})
    out += wml.paragraph(wml.run("Suppressed line"), mark={}, suppressLineNumbers=True)
    out += wml.table([[wml.paragraph(wml.run("Cell one"), mark={}), wml.paragraph(wml.run("Cell two"), mark={})],
                      [wml.paragraph(wml.run("Cell three"), mark={}),
                       wml.paragraph(wml.run("Cell four"), mark={})]], width=3000)
    out += wml.paragraph(wml.run("After the table"), mark={})
    for k in range(3):
        out += wml.paragraph(wml.run(f"{kind} filler {k}: " + WORDS), mark={})
    return out


KINDS = ("plain", "every5", "section", "none", "start", "plain")


@functools.lru_cache(maxsize=None)
def parts() -> tuple[str, str, story_docx.Parts]:
    stories = story_docx.Parts()
    body = ""
    final = ""
    for number, kind in enumerate(KINDS):
        sect = _section(kind, stories)
        content = _content(kind)
        if kind == "plain":
            # Long enough to run on to a second page: numbers restart there.
            content += "".join(wml.paragraph(wml.run(f"Page filler {k}"), mark={}) for k in range(30))
        if number == len(KINDS) - 1:
            body += content
            final = sect
        else:
            body += content + wml.paragraph(wml.run(f"End of {kind}"), mark={}, sect=sect)
    return body, final, stories


def build(setting: str) -> bytes:
    body, final, stories = parts()
    return story_docx.package(body, final, stories, compatibility_mode=SETTINGS[setting])
