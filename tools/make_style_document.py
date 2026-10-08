#!/usr/bin/env python3
"""A document whose runs state nothing: every face, size and spacing is inherited.

The closing test of the style-inheritance work.  Phase 2 reproduced every baseline of
probes whose runs name their face and size directly; a real document names them through
``w:docDefaults``, a theme and a ``w:basedOn`` chain.  This one does it the way Word's
own templates do:

* **theme fonts** -- ``w:docDefaults`` names the minor theme font in all four slots, the
  headings and the title name the major one; the theme is the Office 2013-2022 pair
  (Calibri Light / Calibri);
* **a style chain** -- Heading 2 is based on Heading 1, which is based on Normal; Body
  Text and Quote are based on Normal; the list and code styles likewise;
* **character styles** -- Strong, Emphasis (inside Quote, which is itself italic: the
  toggle XOR makes that run upright), and one that enlarges a run (a mixed-size line);
* **every line rule** -- ``auto`` single (Normal), ``atLeast`` (Body Text), ``exact``
  (Code), and spacing before/after chosen so that "add" and "take the larger" disagree;
* **a list** whose bullet comes from a numbering level with its own ``w:rPr``;
* **a manual page break** followed by a heading with space before.

No run and no paragraph carries a face, a size or a spacing of its own -- only
``w:pStyle``, ``w:rStyle`` and the numbering reference.  ``settings.xml`` holds only
``w:themeFontLang``, so ``w:compat`` is what every other probe holds it at.

Regenerate with ``python tools/make_style_document.py``; measure with
``python tools/read_baselines.py tests/fixtures/style-document.docx``.
"""

from __future__ import annotations

from pathlib import Path

import probe_docx
import wml

OUT = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "style-document.docx"

THEMED = {"asciiTheme": "minorHAnsi", "hAnsiTheme": "minorHAnsi",
          "eastAsiaTheme": "minorEastAsia", "cstheme": "minorBidi"}
MAJOR = {"asciiTheme": "majorHAnsi", "hAnsiTheme": "majorHAnsi",
         "eastAsiaTheme": "majorEastAsia", "cstheme": "majorBidi"}
SYMBOL = {"ascii": "Symbol", "hAnsi": "Symbol", "hint": "default"}

STYLES = [
    wml.style("paragraph", "Normal", default=True),
    wml.style("character", "DefaultParagraphFont", name="Default Paragraph Font", default=True),
    wml.style("table", "TableNormal", name="Normal Table", default=True, tbl=True),
    wml.style("paragraph", "Title", based_on="Normal", next_="Normal",
              ppr_={"spacing": {"after": 0, "line": 240, "lineRule": "auto"}, "contextualSpacing": True},
              rpr_={"rFonts": MAJOR, "kern": 28, "sz": 56, "szCs": 56}),
    wml.style("paragraph", "Heading1", name="heading 1", based_on="Normal", next_="Normal",
              link="Heading1Char",
              ppr_={"keepNext": True, "keepLines": True, "spacing": {"before": 240, "after": 0},
                    "outlineLvl": 0},
              rpr_={"rFonts": MAJOR, "color": "2F5496", "sz": 32, "szCs": 32}),
    wml.style("paragraph", "Heading2", name="heading 2", based_on="Heading1", next_="Normal",
              link="Heading2Char",
              ppr_={"spacing": {"before": 40}, "outlineLvl": 1},
              rpr_={"sz": 26, "szCs": 26}),
    wml.style("character", "Heading1Char", name="Heading 1 Char", based_on="DefaultParagraphFont",
              link="Heading1", rpr_={"rFonts": MAJOR, "color": "2F5496", "sz": 32, "szCs": 32}),
    wml.style("paragraph", "BodyText", name="Body Text", based_on="Normal",
              ppr_={"spacing": {"after": 120, "line": 300, "lineRule": "atLeast"}}),
    wml.style("paragraph", "Quote", based_on="Normal", next_="Normal",
              ppr_={"spacing": {"before": 200, "after": 160}, "ind": {"left": 864, "right": 864},
                    "jc": "center"},
              rpr_={"i": True, "iCs": True, "color": "404040"}),
    wml.style("paragraph", "ListBullet", name="List Bullet", based_on="Normal",
              ppr_={"numPr": '<w:numId w:val="1"/>', "contextualSpacing": True}),
    wml.style("paragraph", "NoSpacing", name="No Spacing",
              ppr_={"spacing": {"after": 0, "line": 240, "lineRule": "auto"}}),
    wml.style("paragraph", "Code", based_on="Normal",
              ppr_={"spacing": {"after": 0, "line": 280, "lineRule": "exact"}, "contextualSpacing": True},
              rpr_={"rFonts": {"ascii": "Courier New", "hAnsi": "Courier New"}, "sz": 20}),
    wml.style("character", "Strong", based_on="DefaultParagraphFont", rpr_={"b": True, "bCs": True}),
    wml.style("character", "Emphasis", based_on="DefaultParagraphFont", rpr_={"i": True, "iCs": True}),
    wml.style("character", "Large", based_on="DefaultParagraphFont", rpr_={"sz": 30, "szCs": 30}),
]

LOREM = ("Style inheritance decides every face and size on this page: nothing here names a "
         "font, a point size or a spacing directly, and each line still has to land on the "
         "device pixel Word put it on.")


def p(style: str | None, *runs: tuple[str, str | None], **extra) -> str:
    body = "".join(wml.run(text, **({"rStyle": rstyle} if rstyle else {})) for text, rstyle in runs)
    props = dict(extra)
    if style:
        props["pStyle"] = style
    return wml.paragraph(body, **props)


def body() -> str:
    parts = [
        p("Title", ("Inherited styles", None)),
        p("Heading1", ("Where the faces come from", None)),
        p("BodyText", ("The body text style is based on Normal and sets only its spacing.", None)),
        p("BodyText", ("This run is ", None), ("strong", "Strong"), (", this one is ", None),
          ("emphasised", "Emphasis"), (".", None)),
        p("BodyText", (LOREM, None)),
        p("Heading2", ("A heading based on a heading", None)),
        p(None, ("A Normal paragraph: docDefaults and nothing else.", None)),
        p(None, ("Another, so that two Normal paragraphs meet.", None)),
        p("Quote", ("A quote is italic, so its ", None), ("emphasis", "Emphasis"),
          (" comes out upright.", None)),
        p("ListBullet", ("The bullet comes from a numbering level", None)),
        p("ListBullet", ("whose own run properties name Symbol,", None)),
        p("ListBullet", ("and list paragraphs share contextual spacing.", None)),
        p(None, ("A line with a ", None), ("larger run", "Large"), (" in it.", None)),
        p("Code", ("code = resolve(style)", None)),
        p("Code", ("exact = 14 pt", None)),
        p("Code", ("contextual = True", None)),
        p("NoSpacing", ("No Spacing is based on nothing at all.", None)),
        p("NoSpacing", ("It still inherits the document defaults.", None)),
        p("BodyText", (LOREM, None)),
        p("Heading2", ("Second-level heading after body text", None)),
        p("BodyText", ("Body text after a heading with no space after.", None)),
        "<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>",
        p("Heading1", ("A heading after a page break", None)),
        p("BodyText", (LOREM, None)),
        p("Quote", ("Quote after body text: before 200 meets after 120.", None)),
        p(None, ("Normal after a quote: after 160 meets before 0.", None)),
        p("Heading1", ("Heading after Normal", None)),
        p("ListBullet", ("A list straight after a heading", None)),
        p("ListBullet", ("and a second item.", None)),
        p("BodyText", (LOREM, None)),
    ]
    return "".join(parts)


def build(compatibility_mode: int | None = None) -> bytes:
    styles = wml.styles_part(
        STYLES,
        run_defaults={"rFonts": THEMED, "sz": 22, "szCs": 22,
                      "lang": {"val": "en-GB", "eastAsia": "en-US", "bidi": "ar-SA"}},
        paragraph_defaults={"spacing": {"after": 160, "line": 240, "lineRule": "auto"}},
    )
    theme = wml.theme_part(
        major={"latin": "Calibri Light", "ea": "", "cs": ""},
        minor={"latin": "Calibri", "ea": "", "cs": ""},
        major_scripts={"Hebr": "Times New Roman", "Arab": "Times New Roman"},
        minor_scripts={"Hebr": "Arial", "Arab": "Arial"},
    )
    numbering = wml.numbering_part([
        ("bullet", "", {"ind": {"left": 720, "hanging": 360}}, {"rFonts": SYMBOL}),
    ])
    settings = wml.settings_part({"val": "en-GB", "eastAsia": "en-US", "bidi": "ar-SA"},
                                 compatibility_mode=compatibility_mode)
    return probe_docx.package(
        body(), styles=styles, extra_parts=(theme, numbering, settings),
        final_section=probe_docx.section(),
    )


if __name__ == "__main__":
    OUT.write_bytes(build())
    print(f"wrote {OUT} ({OUT.stat().st_size} bytes)")
