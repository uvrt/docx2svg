#!/usr/bin/env python3
"""Which header and footer Word draws on which page.

ECMA-376 17.10 says a section's ``w:headerReference`` / ``w:footerReference`` name a
``default``, ``first`` and ``even`` story; that ``first`` applies under ``w:titlePg``,
``even`` under ``w:evenAndOddHeaders`` (``settings.xml``); and that a section stating no
reference of a type takes the previous section's.  Each of those is a claim to measure,
and several questions the text leaves open are here too: is a page even by its *number*
(which ``w:pgNumType/@w:start`` restarts) or by its place in the document; what does a
blank page Word inserts for an ``oddPage`` / ``evenPage`` break show, and what number
does it take; what does a page show whose ``first`` or ``even`` story is not defined
anywhere; which section's stories does a page show when a continuous section starts on it.

Every story is one paragraph naming itself (``Hdr A``), and every footer also carries its
page's number (``PAGE``), the section's page count (``SECTIONPAGES``) and the document's
(``NUMPAGES``), so a page shows which stories it took and what it is numbered.  Every
body page is a paragraph ``Dd Ss Pk`` (``pageBreakBefore`` after the first of a section).
"""

from __future__ import annotations

import functools
from dataclasses import dataclass, field

import story_docx
import wml


@dataclass(frozen=True)
class Sect:
    #: Header references: type -> story name (a story named twice is one part).
    headers: dict = field(default_factory=dict)
    footers: dict = field(default_factory=dict)
    pages: int = 2
    title_page: bool = False
    start: str | None = None
    page_numbers: dict | None = None


@dataclass(frozen=True)
class Doc:
    name: str
    even_and_odd: bool
    sections: tuple


def _docs() -> tuple[Doc, ...]:
    return (
        Doc("inherit", False, (
            Sect({"default": "A", "first": "B"}, {"default": "a", "first": "b"}, 3, title_page=True),
            Sect(pages=2),
            Sect(pages=2, title_page=True),
            Sect({"default": "C"}, {"default": "c"}, 2, title_page=True),
            Sect({"first": "D", "even": "E"}, {"first": "d", "even": "e"}, 2),
            Sect({"default": "F"}, {"default": "f"}, 3, title_page=True, page_numbers={"start": 5}),
        )),
        Doc("even", True, (
            Sect({"default": "A", "even": "E", "first": "B"}, {"default": "a", "even": "e", "first": "b"}, 4,
                 title_page=True),
            Sect(pages=3),
            Sect(pages=3, page_numbers={"start": 1}),
            Sect(pages=2, start="oddPage"),
            Sect({"default": "G"}, {"default": "g"}, 2, start="evenPage"),
            Sect({"even": "H"}, {"even": "h"}, 3, title_page=True),
        )),
        Doc("undefined", True, (
            Sect({"default": "A"}, {"default": "a"}, 3, title_page=True),
            Sect(pages=3),
        )),
        Doc("blank", False, (
            Sect({"default": "A"}, {"default": "a"}, 1),
            Sect({"default": "B"}, {"default": "b"}, 1, start="oddPage"),
            Sect({"default": "C", "first": "D"}, {"default": "c", "first": "d"}, 2, title_page=True,
                 start="evenPage"),
            Sect({"default": "E"}, {"default": "e"}, 1, start="oddPage", page_numbers={"start": 2}),
            Sect({"default": "F"}, {"default": "f"}, 1, start="continuous"),
            Sect({"default": "G", "first": "K"}, {"default": "g", "first": "k"}, 2, start="continuous",
                 title_page=True),
            Sect({"default": "J"}, {"default": "j"}, 1, start="evenPage", page_numbers={"start": 7}),
        )),
    ) + tuple(Doc(f"restart {'even' if even else 'odd'}", even, (
        Sect({"default": "A", "even": "E"}, {"default": "a", "even": "e"}, 1),
        Sect(pages=1, page_numbers={"start": 2}),
        Sect(pages=1, page_numbers={"start": 5}),
        Sect(pages=2, page_numbers={"start": 3}),
        Sect(pages=1, start="oddPage", page_numbers={"start": 4}),
        Sect(pages=1, start="evenPage"),
        Sect(pages=1, start="oddPage", page_numbers={"start": 9}),
    )) for even in (True, False))


DOCS = _docs()


def footer_story(name: str) -> str:
    runs = (wml.run(f"Ftr {name} page ") + story_docx.field("PAGE", "0") + wml.run(" of ")
            + story_docx.field("NUMPAGES", "0") + wml.run(" section ") + story_docx.field("SECTIONPAGES", "0"))
    return wml.paragraph(runs, mark={}, pStyle="Footer")


def header_story(name: str) -> str:
    return wml.paragraph(wml.run(f"Hdr {name}"), mark={}, pStyle="Header")


@functools.lru_cache(maxsize=None)
def parts(index: int) -> tuple[str, str, story_docx.Parts]:
    doc = DOCS[index]
    stories = story_docx.Parts()
    made: dict[tuple[str, str], str] = {}
    body = ""
    final = ""
    for number, sect in enumerate(doc.sections):
        references = ""
        for kind, table, make in (("header", sect.headers, header_story), ("footer", sect.footers, footer_story)):
            for which, name in table.items():
                if (kind, name) not in made:
                    made[(kind, name)] = stories.add(kind, make(name))
                references += stories.reference(kind, which, made[(kind, name)])
        xml = story_docx.section(references=references, start=sect.start, title_page=sect.title_page,
                                 page_numbers=sect.page_numbers)
        for page in range(sect.pages):
            text = wml.run(f"D{index} S{number} P{page}")
            last = page == sect.pages - 1
            props = {"pageBreakBefore": True} if page else {}
            if last and number == len(doc.sections) - 1:
                body += wml.paragraph(text, mark={}, **props)
                final = xml
            else:
                body += wml.paragraph(text, mark={}, sect=xml if last else "", **props)
    return body, final, stories


def build(index: int) -> bytes:
    body, final, stories = parts(index)
    return story_docx.package(body, final, stories, compatibility_mode=15, even_and_odd=DOCS[index].even_and_odd)
