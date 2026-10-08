#!/usr/bin/env python3
"""Where Word starts a line when the left margin is not a whole device pixel.

Found by ``make_story_probe.py``: in its A5 case (left margin 1134 twips, 236.25 px) Word
draws the body's, the header's and the footer's first glyph at 236.0 px.  Here each
section (``nextPage``) has another left margin -- 1440 to 1449 twips, one twip apart, so
the margin takes every fraction of a pixel a twip reaches (5/24 px) -- or a gutter, or a
first-line or left indent on top of such a margin, and a header, a footer and a short
body line, left-aligned, centred and right-aligned (the right margin is stepped too).

One document per setting: no ``settings.xml`` and mode 15.
"""

from __future__ import annotations

import functools

import story_docx
import wml

SETTINGS = {"none": None, "15": 15}


def _cases() -> tuple[dict, ...]:
    out = []
    for k in range(10):
        out.append({"left": 1440 + k, "right": 1440 + k})
    for k in (1, 3, 7):
        out.append({"left": 1440, "right": 1440, "gutter": k})
        out.append({"left": 1440 + k, "right": 1440, "indent": 5})
    out.append({"left": 1134, "right": 1134, "width": 8391, "height": 11906})
    # Margins of whole pixels whose layout unit is past a half (120 px is 117,964.8 units,
    # 240 px 235,929.6, 270 px 265,420.8): which unit the column starts on.
    for left in (576, 1152, 1296):
        out.append({"left": left, "right": 1440})
    return tuple(out)


CASES = _cases()


@functools.lru_cache(maxsize=None)
def parts() -> tuple[str, str, story_docx.Parts]:
    stories = story_docx.Parts()
    body = ""
    final = ""
    for number, case in enumerate(CASES):
        case = dict(case)
        indent = case.pop("indent", 0)
        refs = (stories.reference("header", "default", stories.add("header", wml.paragraph(
                    wml.run(f"H{number} margin"), mark={}, pStyle="Header")))
                + stories.reference("footer", "default", stories.add("footer", wml.paragraph(
                    wml.run(f"F{number} margin"), mark={}, pStyle="Footer", jc="right"))))
        sect = story_docx.section(references=refs, **case)
        props = {"ind": {"left": indent}} if indent else {}
        content = (wml.paragraph(wml.run(f"Case {number} left"), mark={}, **props)
                   + wml.paragraph(wml.run(f"Case {number} centre"), mark={}, jc="center"))
        closing = wml.run(f"Case {number} right")
        if number == len(CASES) - 1:
            body += content + wml.paragraph(closing, mark={}, jc="right")
            final = sect
        else:
            body += content + wml.paragraph(closing, mark={}, jc="right", sect=sect)
    return body, final, stories


def build(setting: str) -> bytes:
    body, final, stories = parts()
    return story_docx.package(body, final, stories, compatibility_mode=SETTINGS[setting])
