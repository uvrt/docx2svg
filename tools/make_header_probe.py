#!/usr/bin/env python3
"""Whether a header or footer taller than its margin takes room from the body.

A header sits ``w:pgMar/@w:header`` below the page's top edge; the body starts at the top
margin.  When the header reaches below the top margin, does the body start lower -- and
by how much, counting what of the header?  The same for a footer and the bottom margin.
This decides page breaks, so it is measured here although headers and footers are
otherwise out of Phase 4's scope.

One section per case (``nextPage``), each with its own header and footer parts and one
body paragraph of 60 Calibri 11 pt lines (``w:br``, widow control off), so the page shows
both where the body starts (the first line's baseline) and how many lines it holds (where
it ends).  Header and footer contents are an ``exact`` line of ``t`` twips, unless said:

* ``header t`` / ``footer t`` -- ``t`` = 300, 700, 720, 721, 740, 900, 1500 at the default
  distance 720 and margins 1440: from well inside the margin to 780 twips past it;
* ``header spaced`` -- two natural lines with 240 twips after (does the space count?);
  ``header before`` -- one natural line with 480 twips before;
* ``header distance`` -- the header 200 twips from the edge, 1300 tall (1500 past it);
* ``header, negative top`` -- ``w:top="-1440"``: a top margin Word keeps however tall
  the header ("exactly"), with a 1500 header;
* ``plain`` -- no header or footer.

One document per compatibility setting: no ``settings.xml``, and modes 12, 14 and 15.
"""

from __future__ import annotations

import functools
from dataclasses import dataclass

import probe_docx
import wml

WIDTH, HEIGHT = 11900, 16840
FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
RUN = {"rFonts": FACE, "sz": 22, "szCs": 22}
SPACING = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}
LINES = 60
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
HEADER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
FOOTER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
HEADER_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
FOOTER_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"


@dataclass(frozen=True)
class Case:
    name: str
    #: The header's and the footer's paragraphs, as ``(lines, exact twips or None,
    #: before, after)``; empty for none.
    header: tuple = ()
    footer: tuple = ()
    top: int = 1440
    bottom: int = 1440
    header_distance: int = 720
    footer_distance: int = 720


def _cases() -> tuple[Case, ...]:
    out = [Case("plain")]
    for t in (300, 700, 720, 721, 740, 900, 1500):
        out.append(Case(f"header {t}", header=((1, t, 0, 0),)))
    for t in (300, 700, 720, 721, 740, 900, 1500):
        out.append(Case(f"footer {t}", footer=((1, t, 0, 0),)))
    out.append(Case("header spaced", header=((2, None, 0, 240),)))
    out.append(Case("header before", header=((1, None, 480, 0),)))
    out.append(Case("header distance", header=((1, 1300, 0, 0),), header_distance=200))
    out.append(Case("header, negative top", header=((1, 1500, 0, 0),), top=-1440))
    out.append(Case("header and footer 1500", header=((1, 1500, 0, 0),), footer=((1, 1500, 0, 0),)))
    return tuple(out)


CASES = _cases()
SETTINGS = {"none": None, "12": 12, "14": 14, "15": 15}


def _story(tag: str, number: int, paragraphs: tuple) -> str:
    """A header's or footer's paragraphs."""
    out = ""
    for index, (lines, exact, before, after) in enumerate(paragraphs):
        spacing = {"before": before, "after": after, "line": exact or 240, "lineRule": "exact" if exact else "auto"}
        texts = [f"{tag}{number}.{index}.{k}" for k in range(lines)]
        runs = f"<w:r>{wml.rpr(**RUN)}" + "<w:br/>".join(
            f'<w:t xml:space="preserve">{t}</w:t>' for t in texts) + "</w:r>"
        out += wml.paragraph(runs, mark=RUN, spacing=spacing)
    return out


@functools.lru_cache(maxsize=1)
def parts() -> tuple[str, tuple]:
    """The body, and the header and footer parts ``(name, content type, rel, xml)``."""
    body = ""
    extra = []
    for number, case in enumerate(CASES):
        refs = ""
        for kind, paragraphs, content_type, rel in (("header", case.header, HEADER_CONTENT_TYPE, HEADER_REL),
                                                     ("footer", case.footer, FOOTER_CONTENT_TYPE, FOOTER_REL)):
            # An empty part where the case has none, so no section inherits the one before's.
            root = "hdr" if kind == "header" else "ftr"
            xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                   f'<w:{root} xmlns:w="{probe_docx.W_NS}">'
                   + (_story("h" if kind == "header" else "g", number, paragraphs)
                      or wml.paragraph("", mark=RUN, spacing={"before": 0, "after": 0, "line": 20,
                                                              "lineRule": "exact"}))
                   + f"</w:{root}>")
            extra.append((f"word/{kind}{number}.xml", content_type, rel, xml))
            refs += f'<w:{kind}Reference w:type="default" r:id="rId{len(extra) + 1}"/>'
        section = (
            f'<w:sectPr xmlns:r="{R_NS}">{refs}'
            f'<w:pgSz w:w="{WIDTH}" w:h="{HEIGHT}"/>'
            f'<w:pgMar w:top="{case.top}" w:right="1440" w:bottom="{case.bottom}" w:left="1440"'
            f' w:header="{case.header_distance}" w:footer="{case.footer_distance}" w:gutter="0"/>'
            '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>'
        )
        runs = f"<w:r>{wml.rpr(**RUN)}" + "<w:br/>".join(
            f'<w:t xml:space="preserve">Case {number} l{k}</w:t>' for k in range(LINES)) + "</w:r>"
        last = number == len(CASES) - 1
        body += wml.paragraph(runs, mark=RUN, spacing=SPACING, widowControl=False,
                              sect="" if last else section)
        if last:
            final = section
    return body, final, tuple(extra)


def build(setting: str) -> bytes:
    body, final, extra = parts()
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": SPACING},
    )
    mode = SETTINGS[setting]
    settings = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    # extra_parts are related as rId2, rId3, ... in order: headers and footers first.
    return probe_docx.package(body, final_section=final, styles=styles, extra_parts=extra + settings)
