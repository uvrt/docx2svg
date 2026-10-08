#!/usr/bin/env python3
"""Floating drawings in headers and footers: the logo case.

``make_anchor_probe.py`` measured where an anchor goes in the body.  A header or footer
is laid out as a document of its own (ROADMAP.md, "Headers, footers and fields --
measured"), so its anchors meet frames of their own: the story's paragraph and line, and
the page's margins seen from a story.  And they stack with the body's.  Every case is a
section of one page (``nextPage``) with its own header and footer, and a body of a few
lines of text; the pictures are ``story_docx``'s grey PNG, the shapes solid fills.
Cases (``CASES``):

* a header's picture against its paragraph, its line and character, the page, the margin,
  the top margin; in front of the body's text it overlaps, and behind it;
* a footer's picture against its paragraph (a footer grows upward from its distance),
  the page's foot and the bottom margin;
* a header's text box (a filled shape with a paragraph);
* a body's picture behind and in front of the header's text, over the header's area.

Two documents: no ``settings.xml`` and mode 15.  Reader: ``read_anchor_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass

import make_drawing_probe as drawing_probe
import story_docx
import wml
from make_anchor_probe import Anchor, pic_graphic

SETTINGS = {"none": None, "15": 15}
WORDS = ("Body text the header's drawings float over or under, laid out as if they were not "
         "there, line after line down the page.")


def picture(number: int, cx: int, cy: int) -> str:
    return pic_graphic(number, cx, cy).replace('r:embed="rId2"', f'r:embed="{story_docx.IMAGE_ID}"')


@dataclass(frozen=True)
class Case:
    note: str
    header: tuple = ()
    footer: tuple = ()
    body: tuple = ()


def _cases() -> tuple[Case, ...]:
    size = (1300000, 500000)

    def pic(h, v, k=0, **kw):
        return Anchor(h, v, size[0] + k * 3701, size[1] + k * 37003, **kw)

    out = [
        Case("header paragraph, in front", header=(pic(("column", "offset", 300000), ("paragraph", "offset", 400000)),)),
        Case("header paragraph, behind", header=(pic(("column", "offset", 300000), ("paragraph", "offset", 400000), 1,
                                                     behind=True),)),
        Case("header page", header=(pic(("page", "offset", 1000003), ("page", "offset", 200003), 2),)),
        Case("header margin", header=(pic(("margin", "align", "right"), ("margin", "align", "top"), 3),)),
        Case("header top margin", header=(pic(("rightMargin", "align", "right"), ("topMargin", "align", "center"), 4),)),
        Case("header line and character", header=(pic(("character", "offset", 100000), ("line", "offset", 50000), 5),)),
        Case("footer paragraph", footer=(pic(("column", "offset", 200000), ("paragraph", "offset", -300000), 6),)),
        Case("footer page foot", footer=(pic(("margin", "align", "center"), ("page", "align", "bottom"), 7),)),
        Case("footer bottom margin", footer=(pic(("column", "offset", 0), ("bottomMargin", "offset", 100000), 8),)),
        Case("header text box", header=(Anchor(("margin", "offset", 500000), ("paragraph", "offset", 100000), 2000000,
                                               600000, graphic=drawing_probe.graphic(drawing_probe.shape(
                                                   drawing_probe.preset("rect"),
                                                   drawing_probe.solid(drawing_probe.srgb("DEEBF7")),
                                                   drawing_probe.NO_LINE, cx=2000000, cy=600000,
                                                   text=drawing_probe.text_paragraph("Header text box.")),
                                                   "wps")),)),
        Case("body behind, over the header", body=(pic(("column", "offset", 100000), ("page", "offset", 400000), 9,
                                                       behind=True),)),
        Case("body in front, over the header", body=(pic(("column", "offset", 100000), ("page", "offset", 400000), 10),)),
    ]
    return tuple(out)


CASES = _cases()


def _anchors(anchors, serial: int) -> str:
    """The anchors' runs, a picture of each anchor's extent where it names no graphic."""
    import dataclasses

    out = ""
    for k, anchor in enumerate(anchors):
        if anchor.graphic is None:
            anchor = dataclasses.replace(anchor, graphic=picture(serial + k, anchor.cx, anchor.cy))
        out += anchor.xml(serial + k)
    return out


def build(setting: str) -> bytes:
    stories = story_docx.Parts()
    body = ""
    final = ""
    serial = 1
    for number, case in enumerate(CASES):
        header = wml.paragraph(wml.run(f"Header {number} ") + _anchors(case.header, serial)
                               + wml.run("text of the header."), mark={}, pStyle="Header")
        serial += len(case.header)
        footer = wml.paragraph(wml.run(f"Footer {number} ") + _anchors(case.footer, serial)
                               + wml.run("text of the footer."), mark={}, pStyle="Footer")
        serial += len(case.footer)
        refs = (stories.reference("header", "default", stories.add("header", header, pictures=True))
                + stories.reference("footer", "default", stories.add("footer", footer, pictures=True)))
        sect = story_docx.section(references=refs)
        content = wml.paragraph(wml.run(f"Case {number} {case.note}. ") + _anchors(case.body, serial)
                                + wml.run(WORDS), mark={})
        serial += len(case.body)
        content += wml.paragraph(wml.run(f"Case {number} after."), mark={})
        if number == len(CASES) - 1:
            body += content
            final = sect
        else:
            body += content + wml.paragraph(wml.run(f"Case {number} ends."), mark={}, sect=sect)
    return story_docx.package(body, final, stories, compatibility_mode=SETTINGS[setting], body_pictures=True)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"story-anchor-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)
