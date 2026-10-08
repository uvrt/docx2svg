#!/usr/bin/env python3
"""Text in the faces Office for Mac keeps in its cloud-font cache, laid out as Word lays it.

Word 365 makes every new document with Aptos Display headings, and Office downloads that
face -- and others a document asks for, Segoe UI, Lato... -- into
``~/Library/Group Containers/UBF8T346G9.Office/FontCache/4/CloudFonts/``, one folder per
family, where neither macOS nor Word's bundle has them.  docx2svg reads that cache in
place (:data:`docx2svg.fonts.OFFICE_CLOUD_FONTS`); this probe checks the faces it finds
there are the ones Word lays out and draws with.

One page per family the cache holds on the machine that built the probe
(:data:`FAMILIES`): a heading-sized line in each of the four styles, a wrapped paragraph at
11 pt and two at 20 and 28 pt, with no ``settings.xml`` and in mode 15.  Every face is
named in all four ``w:rFonts`` slots.
"""

from __future__ import annotations

import functools

import story_docx
import wml

SETTINGS = {"none": None, "15": 15}

#: The families of the cloud-font cache the probe was built for (the folders' names); a
#: machine without one of them lays its page out in a substitute.
FAMILIES = ("Aptos Display", "Segoe UI", "Lato", "Raleway", "Ubuntu", "Ubuntu Mono", "Anton",
            "Merriweather Sans Light")

TEXT = ("The quick brown fox jumps over the lazy dog while seventeen wizards quietly box "
        "jumbled vexing quartz figures; Pack my box with five dozen liquor jugs, then "
        "measure every line again.")


def _fonts(family: str) -> dict:
    return {"ascii": family, "hAnsi": family, "eastAsia": family, "cs": family}


@functools.lru_cache(maxsize=None)
def parts() -> tuple[str, str]:
    body = ""
    for number, family in enumerate(FAMILIES):
        faces = _fonts(family)
        content = wml.paragraph(wml.run(f"{family} heading 20", rFonts=faces, sz=40, szCs=40), mark={},
                                pageBreakBefore=number > 0)
        for bold, italic, tag in ((True, False, "bold"), (False, True, "italic"), (True, True, "bold italic")):
            content += wml.paragraph(wml.run(f"{family} {tag} 16", rFonts=faces, b=bold, i=italic, sz=32, szCs=32),
                                     mark={})
        for size in (22, 40, 56):
            content += wml.paragraph(wml.run(f"{size // 2} pt: " + TEXT, rFonts=faces, sz=size, szCs=size),
                                     mark={})
        body += content
    return body, story_docx.section()


def build(setting: str) -> bytes:
    body, final = parts()
    return story_docx.package(body, final, story_docx.Parts(), compatibility_mode=SETTINGS[setting])
