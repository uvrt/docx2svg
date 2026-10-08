#!/usr/bin/env python3
"""Finding 6: the size (and alignment) Word gives a table cell's text by style.

``filesamples/sample1`` (mode 12) draws its ``TableGrid`` cells at 46 px (11 pt) where
the cascade resolves Normal's 12 pt; with mode 15 set, at 50 px.  ECMA-376 puts the
table style below the paragraph style, and the style probe measured that (p06-p08:
Normal's 11 pt beat a table style's 14 pt) -- in a document of no stated mode, with
Normal equal to ``w:docDefaults``.  Word's ``overrideTableStyleFontSizeAndJustification``
compatibility option is named for exactly this.  So this probe crosses:

* **Normal** 12 pt (``w:sz`` 24) over ``w:docDefaults`` 11 pt, and in two documents
  centred (``w:jc``) as well;
* **the table style**: ``Plain`` (states only table properties), ``Sized`` (``w:sz`` 18,
  9 pt), ``Right`` (``w:jc`` right), ``Spaced`` (a ``w:spacing`` only, as Word's
  ``TableGrid`` states), ``Coloured`` (a ``w:color`` only), and none at all;
* **the paragraph's style**: none (Normal), and ``Body`` (based on Normal, stating
  nothing), and ``Big`` (``w:sz`` 32);
* **the setting**: no ``settings.xml``, modes 12, 14 and 15, and 14 and 15 with the
  ``overrideTableStyleFontSizeAndJustification`` compatibility setting on -- which every
  document Word 2013 and later writes carries beside mode 15;
* **the sizes** (``sizes-<docDefaults>-<Normal>``, no mode stated): ``w:docDefaults``
  from 8 to 14 pt and Normal from 10 to 14 pt, eleven pairings -- the table style's size
  wins exactly where Normal is 12 pt, and a silent table style's ``w:docDefaults`` size
  except where that is 10 pt.

Each case is a one-row, two-cell table (a left-aligned and a right-aligned word, so the
alignment shows too) after a body paragraph of the same styles for comparison.  Reader:
``read_table_probes.py``.
"""

from __future__ import annotations

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
#: Document -> (mode, the override setting on, Normal centred, w:docDefaults' size, Normal's).
SETTINGS = {"none": (None, False, False, 22, 24), "12": (12, False, False, 22, 24), "14": (14, False, False, 22, 24),
            "15": (15, False, False, 22, 24), "14-override": (14, True, False, 22, 24),
            "15-override": (15, True, False, 22, 24), "14-centred": (14, False, True, 22, 24),
            "15-centred": (15, True, True, 22, 24),
            # Which Normal sizes a table style's size beats, with no mode stated.
            "sizes-24-22": (None, False, False, 24, 22), "sizes-22-22": (None, False, False, 22, 22),
            "sizes-24-28": (None, False, False, 24, 28), "sizes-20-24": (None, False, False, 20, 24),
            "sizes-24-20": (None, False, False, 24, 20), "sizes-16-24": (None, False, False, 16, 24),
            "sizes-21-24": (None, False, False, 21, 24), "sizes-26-24": (None, False, False, 26, 24),
            "sizes-22-23": (None, False, False, 22, 23), "sizes-22-25": (None, False, False, 22, 25),
            "sizes-20-22": (None, False, False, 20, 22)}
TABLE_STYLES = (None, "Plain", "Sized", "Right", "Spaced", "Coloured")
PARAGRAPH_STYLES = (None, "Body", "Big")
PER_PAGE = 10


def _table(number: int, table_style: str | None, paragraph_style: str | None) -> str:
    pr = (f'<w:tblStyle w:val="{table_style}"/>' if table_style else "")
    pr += ('<w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
           '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
           '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
           '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0"'
           ' w:noHBand="1" w:noVBand="1"/>')
    style = {"pStyle": paragraph_style} if paragraph_style else {}
    cells = ""
    for index, text in enumerate((f"Cell{number}a", f"Cell{number}b")):
        props = dict(style)
        if index == 1:
            props["jc"] = "right"
        cells += (f'<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>'
                  f"{wml.paragraph(wml.run(text), **props)}</w:tc>")
    return (f'<w:tbl><w:tblPr>{pr}</w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/>'
            f"</w:tblGrid><w:tr>{cells}</w:tr></w:tbl>")


def cases() -> list[tuple[str | None, str | None]]:
    return [(t, p) for t in TABLE_STYLES for p in PARAGRAPH_STYLES]


def body() -> str:
    out = ""
    for number, (table_style, paragraph_style) in enumerate(cases()):
        props = {"pageBreakBefore": True} if number and number % PER_PAGE == 0 else {}
        if paragraph_style:
            props["pStyle"] = paragraph_style
        out += wml.paragraph(wml.run(f"Body{number} {table_style} {paragraph_style}"), **props)
        out += _table(number, table_style, paragraph_style)
    return out + wml.paragraph(wml.run("End"))


def styles(centred: bool = False, defaults: int = 22, normal: int = 24) -> str:
    table_styles = []
    for style_id, ppr, rpr in (("Plain", "", ""), ("Sized", "", '<w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'),
                                ("Right", '<w:pPr><w:jc w:val="right"/></w:pPr>', ""),
                                ("Spaced", '<w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>', ""),
                                ("Coloured", "", '<w:rPr><w:color w:val="C00000"/></w:rPr>')):
        table_styles.append(
            f'<w:style w:type="table" w:styleId="{style_id}"><w:name w:val="{style_id}"/>{ppr}{rpr}'
            '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/>'
            '<w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>'
            "</w:tblCellMar></w:tblPr></w:style>")
    return wml.styles_part(
        [wml.style("paragraph", "Normal", default=True, rpr_={"sz": normal, "szCs": normal},
                   ppr_={"jc": "center"} if centred else None),
         wml.style("paragraph", "Body", based_on="Normal"),
         wml.style("paragraph", "Big", based_on="Normal", rpr_={"sz": 32, "szCs": 32}),
         *table_styles],
        run_defaults={"rFonts": FACE, "sz": defaults, "szCs": defaults},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )


def build(setting: str) -> bytes:
    mode, override, centred, defaults, normal = SETTINGS[setting]
    extra = ()
    if mode is not None:
        compat = ('<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification"'
                  ' w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>') if override else ""
        extra = (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode, compat_elements=compat),)
    return probe_docx.package(body(), styles=styles(centred, defaults, normal), extra_parts=extra)
