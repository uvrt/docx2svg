#!/usr/bin/env python3
"""Whether a line's space after has to fit above the page's footnotes.

``make_page_fit_probe.py`` measured that a page's last line need not fit its paragraph's
space after above the bottom margin, and ``make_footnote_probe.py`` that a line stays
when the separator and every note line fit below it -- with paragraphs of no space after.
docx-agent's chart fixture has a paragraph of 160 twips after just above a footnote, which
Word sends to the next page where docx2svg kept it (ROADMAP.md, F.24).  Every case is a
page: a line carrying a footnote reference, a spacer paragraph of an exact height, and
the candidate -- one line with 240 twips after -- then the next case on a new page.  The
spacer's height steps by 20 twips over the range where the candidate leaves the page, so
where it goes says, to the twip, what of it had to fit.  The same with a candidate of no
space after (``after0``).  Notes are ``FootnoteText`` (Calibri 10 pt, no spacing).

Three documents each: no ``settings.xml``, mode 14 and mode 15.  Reader:
``read_note_after_probe.py``.
"""

from __future__ import annotations

import make_footnote_draw_probe as notes
import probe_docx
import wml

SETTINGS = {"none": None, "14": 14, "15": 15}
KINDS = {"after240": 240, "after0": 0}
#: The spacer's exact heights, twips: about where the candidate leaves the page.
HEIGHTS = tuple(range(12200, 13000, 20))
FOOTNOTES_CONTENT_TYPE = notes.FOOTNOTES_CONTENT_TYPE
FOOTNOTES_REL = notes.FOOTNOTES_REL


def body(after: int) -> str:
    out = ""
    for k, height in enumerate(HEIGHTS):
        out += wml.paragraph(notes.text(f"Case {k} spacer {height}") + notes.reference(k + 1),
                             spacing=notes.SPACING, pageBreakBefore=bool(k))
        out += wml.paragraph(notes.text("s"), spacing={"before": 0, "after": 0, "line": height, "lineRule": "exact"})
        out += wml.paragraph(notes.text(f"Candidate {k}"), spacing={"before": 0, "after": after, "line": 240,
                                                                     "lineRule": "auto"})
    return out


def build(name: str) -> bytes:
    """``name``: ``noteafter-<kind>-<setting>``."""
    _, kind, setting = name.split("-")
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True),
         wml.style("paragraph", "FootnoteText", based_on="Normal", ppr_={"spacing": notes.SPACING},
                   rpr_={"sz": 20, "szCs": 20}),
         wml.style("character", "FootnoteReference", rpr_={"vertAlign": "superscript"})],
        run_defaults={"rFonts": notes.FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": notes.SPACING},
    )
    footnotes = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                 f'<w:footnotes xmlns:w="{probe_docx.W_NS}">{notes.separators()}'
                 + "".join(notes.Note((f"note {k}",)).xml(k + 1) for k in range(len(HEIGHTS))) + "</w:footnotes>")
    extra = [("word/footnotes.xml", FOOTNOTES_CONTENT_TYPE, FOOTNOTES_REL, footnotes)]
    mode = SETTINGS[setting]
    if mode is not None:
        extra.append(wml.settings_part({"val": "en-GB"}, compatibility_mode=mode))
    return probe_docx.package(body(KINDS[kind]), final_section=notes.section(), styles=styles,
                              extra_parts=tuple(extra))


DOCUMENTS = tuple(f"noteafter-{kind}-{setting}" for kind in KINDS for setting in SETTINGS)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for name in DOCUMENTS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"{name}.docx"
        path.write_bytes(build(name))
        print(path)
