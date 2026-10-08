#!/usr/bin/env python3
"""Style-inheritance probes: what Word draws for a run that states nothing itself.

The cascade is "solved in the field" in the sense that ECMA-376 describes it.  It is also
the classic place where Word, LibreOffice and the spec have disagreed, so each claim the
resolver in ``docx2svg.resolve`` relies on is asked of Word here, one variable per case.
Everything asked is observable in the PDF export without a font file: the face and
weight are the ``/BaseFont`` of the text object the glyphs are drawn in, the size is its
``Tm`` scale, ``w:caps`` shows as uppercase text, ``w:vanish`` as absent text, and an
indent as the pen x of the first glyph (``quartz_pdf.py``).

Documents -- one per thing that has to be global:

* **toggles** -- ECMA-376 17.7.3 says ``w:b``, ``w:i``, ``w:caps``, ``w:vanish``... are
  *toggle properties*: inside the style hierarchy each occurrence flips the value, and
  only direct formatting sets it absolutely.  Cases pair paragraph style, character
  style, ``basedOn`` parents, table style, ``w:val="0"`` in a style, and direct
  formatting, and one case per other toggle (italic, caps, vanish).
* **defaults-bold** / **normal-bold** / **charstyle-bold** -- the same question for the
  three places a toggle can come from without being named by the run: ``w:docDefaults``,
  the default paragraph style, and the default character style.
* **fonts** -- theme fonts (``w:asciiTheme`` and friends through ``theme1.xml``), the
  precedence of a theme attribute over the explicit one on the same element and across
  levels, the four ``w:rFonts`` slots chosen per character, ``w:hint``, and the complex
  script size and weight (``w:szCs``, ``w:bCs``).
* **no-styles** / **no-theme** / **theme-no-rfonts** -- what Word uses when nothing names
  a face or a size.  Phase 0's ``hello.docx`` came back with Aptos in it although it
  named only Calibri.
* **precedence** -- numbering and table styles against paragraph styles, for indent and
  size.

Throwaway: exported through ``tools/oracle.py``, never committed.  Measure with
``python tools/read_style_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import probe_docx
import wml

CALIBRI = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
PLAIN_DEFAULTS = {"rFonts": CALIBRI, "sz": 22, "szCs": 22, "lang": {"val": "en-GB"}}
SINGLE = {"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}}


@dataclass(frozen=True)
class Case:
    """One paragraph: a label run, then the run whose resolution is the question.

    ``label`` is written into the label run (``"t01 "``) and is how the reader finds the
    line again.  ``sample`` is the tested run's text.  The label run carries
    ``label_props`` only, so it shows what the *paragraph* level resolves to.
    """

    label: str
    pstyle: str | None = None
    rstyle: str | None = None
    direct: dict = field(default_factory=dict)
    sample: str = "sample"
    label_props: dict = field(default_factory=dict)
    paragraph: dict = field(default_factory=dict)
    table_style: str | None = None
    #: What the case asks, for the reader's report.
    note: str = ""

    def xml(self) -> str:
        props = dict(self.direct)
        if self.rstyle:
            props["rStyle"] = self.rstyle
        runs = wml.run(f"{self.label} ", **self.label_props) + wml.run(self.sample, **props)
        para_props = dict(self.paragraph)
        if self.pstyle:
            para_props["pStyle"] = self.pstyle
        return wml.paragraph(runs, **para_props)


@dataclass(frozen=True)
class Probe:
    name: str
    cases: tuple[Case, ...]
    styles: str
    extra_parts: tuple = ()

    def build(self) -> bytes:
        body = ""
        for case in self.cases:
            if case.table_style is not None:
                body += wml.table([[case.xml()]], style=case.table_style or None)
                body += wml.paragraph("")  # a table may not be followed by another directly
            else:
                body += case.xml()
        return probe_docx.package(body, styles=self.styles, extra_parts=self.extra_parts)


def _basic_styles(extra: list[str], run_defaults=PLAIN_DEFAULTS, normal_rpr=None) -> list[str]:
    return [
        wml.style("paragraph", "Normal", default=True, rpr_=normal_rpr),
        wml.style("character", "DefaultParagraphFont", name="Default Paragraph Font", default=True),
        wml.style("table", "TableNormal", name="Normal Table", default=True, tbl=True),
        *extra,
    ]


# -- toggles -----------------------------------------------------------------

TOGGLE_STYLES = [
    wml.style("paragraph", "PB", based_on="Normal", rpr_={"b": True}),
    wml.style("paragraph", "PB2", based_on="PB", rpr_={"b": True}),
    wml.style("paragraph", "PB0", based_on="PB", rpr_={"b": False}),
    wml.style("paragraph", "PBkid", based_on="PB"),
    wml.style("paragraph", "PI", based_on="Normal", rpr_={"i": True}),
    wml.style("paragraph", "PBI", based_on="Normal", rpr_={"b": True, "i": True}),
    wml.style("paragraph", "PCaps", based_on="Normal", rpr_={"caps": True}),
    wml.style("paragraph", "PV", based_on="Normal", rpr_={"vanish": True}),
    wml.style("paragraph", "PNB", based_on="Normal"),
    wml.style("paragraph", "PLink", based_on="Normal", link="PLinkChar", rpr_={"i": True}),
    wml.style("character", "CB", based_on="DefaultParagraphFont", rpr_={"b": True}),
    wml.style("character", "CB2", based_on="CB", rpr_={"b": True}),
    wml.style("character", "CB0", based_on="DefaultParagraphFont", rpr_={"b": False}),
    wml.style("character", "CI", based_on="DefaultParagraphFont", rpr_={"i": True}),
    wml.style("character", "CCaps", based_on="DefaultParagraphFont", rpr_={"caps": True}),
    wml.style("character", "CV", based_on="DefaultParagraphFont", rpr_={"vanish": True}),
    wml.style("character", "PLinkChar", based_on="DefaultParagraphFont", link="PLink",
              rpr_={"b": True}),
    wml.style("table", "TB", based_on="TableNormal", rpr_={"b": True}, tbl=True),
]

VISIBLE = {"vanish": False}

TOGGLE_CASES = (
    Case("t01", pstyle="PB", note="paragraph style bold"),
    Case("t02", rstyle="CB", note="character style bold"),
    Case("t03", pstyle="PB", rstyle="CB", note="paragraph bold + character bold"),
    Case("t04", pstyle="PB", rstyle="CB", direct={"b": True}, note="... + direct b"),
    Case("t05", pstyle="PB", rstyle="CB", direct={"b": False}, note="... + direct b=0"),
    Case("t06", pstyle="PB", direct={"b": True}, note="paragraph bold + direct b"),
    Case("t07", pstyle="PB", direct={"b": False}, note="paragraph bold + direct b=0"),
    Case("t08", pstyle="PB2", note="bold style basedOn bold style"),
    Case("t09", pstyle="PB0", note="b=0 style basedOn bold style"),
    Case("t10", pstyle="PBkid", note="empty style basedOn bold style"),
    Case("t11", pstyle="PB", rstyle="CB0", note="paragraph bold + character b=0"),
    Case("t12", rstyle="CB2", note="character bold basedOn character bold"),
    Case("t13", pstyle="PB2", rstyle="CB", note="PB2 (bold on bold) + character bold"),
    Case("t14", pstyle="PB", rstyle="CB2", note="paragraph bold + CB2 (bold on bold)"),
    Case("t15", pstyle="PI", rstyle="CI", note="paragraph italic + character italic"),
    Case("t16", pstyle="PBI", rstyle="CB", note="paragraph bold+italic + character bold"),
    Case("t17", pstyle="PCaps", rstyle="CCaps", note="paragraph caps + character caps"),
    Case("t18", pstyle="PCaps", note="paragraph caps"),
    Case("t19", rstyle="CV", note="character vanish"),
    Case("t20", pstyle="PV", rstyle="CV", label_props=VISIBLE,
         note="paragraph vanish + character vanish"),
    Case("t21", pstyle="PV", label_props=VISIBLE, note="paragraph vanish"),
    Case("t22", rstyle="PLink", note="rStyle naming a linked *paragraph* style"),
    Case("t23", pstyle="Nope", note="pStyle naming no style"),
    Case("t24", table_style="TB", note="table style bold"),
    Case("t25", table_style="TB", pstyle="PB", note="table bold + paragraph bold"),
    Case("t26", table_style="TB", rstyle="CB", note="table bold + character bold"),
    Case("t27", direct={"b": True}, note="direct b"),
    Case("t28", pstyle="PNB", rstyle="CB", note="plain style + character bold"),
    Case("t29", table_style="TB", pstyle="PB", rstyle="CB",
         note="table + paragraph + character bold"),
    Case("t30", rstyle="CB", direct={"b": False}, note="character bold + direct b=0"),
)

# -- globally bold: docDefaults, Normal, default character style -----------------

GLOBAL_CASES = (
    Case("g01", note="nothing named"),
    Case("g02", pstyle="PB", note="+ paragraph style bold"),
    Case("g03", rstyle="CB", note="+ character style bold"),
    Case("g04", direct={"b": False}, note="+ direct b=0"),
    Case("g05", pstyle="PNB", note="+ empty paragraph style basedOn Normal"),
    Case("g06", pstyle="PFree", note="+ paragraph style with no basedOn"),
    Case("g07", pstyle="PB", rstyle="CB", note="+ paragraph bold + character bold"),
    Case("g08", table_style="TB", note="+ table style bold"),
    Case("g09", rstyle="CNB", note="+ empty character style basedOn the default one"),
    Case("g10", pstyle="P0", note="+ paragraph style b=0 basedOn Normal"),
    Case("g11", rstyle="C0", note="+ character style b=0"),
)
GLOBAL_STYLES = [
    wml.style("paragraph", "PB", based_on="Normal", rpr_={"b": True}),
    wml.style("paragraph", "PNB", based_on="Normal"),
    wml.style("paragraph", "PFree"),
    wml.style("paragraph", "P0", based_on="Normal", rpr_={"b": False}),
    wml.style("character", "CNB", based_on="DefaultParagraphFont"),
    wml.style("character", "C0", based_on="DefaultParagraphFont", rpr_={"b": False}),
    wml.style("character", "CB", based_on="DefaultParagraphFont", rpr_={"b": True}),
    wml.style("table", "TB", based_on="TableNormal", rpr_={"b": True}, tbl=True),
]


def _global_probe(name: str, where: str) -> Probe:
    run_defaults = dict(PLAIN_DEFAULTS)
    if where == "defaults":
        run_defaults["b"] = True
    styles = [
        wml.style("paragraph", "Normal", default=True,
                  rpr_={"b": True} if where == "normal" else None),
        wml.style("character", "DefaultParagraphFont", name="Default Paragraph Font",
                  default=True, rpr_={"b": True} if where == "charstyle" else None),
        wml.style("table", "TableNormal", name="Normal Table", default=True, tbl=True),
        *GLOBAL_STYLES,
    ]
    return Probe(name, GLOBAL_CASES,
                 wml.styles_part(styles, run_defaults=run_defaults, paragraph_defaults=SINGLE))


# -- fonts ---------------------------------------------------------------------

THEME = wml.theme_part(
    major={"latin": "Trebuchet MS", "ea": "", "cs": "Verdana"},
    minor={"latin": "Georgia", "ea": "", "cs": "Courier New"},
    major_scripts={"Hebr": "Courier New"},
    minor_scripts={"Hebr": "Arial", "Hans": "DengXian"},
)
THEMED = {"asciiTheme": "minorHAnsi", "hAnsiTheme": "minorHAnsi",
          "eastAsiaTheme": "minorEastAsia", "cstheme": "minorBidi"}
FONT_STYLES = [
    wml.style("paragraph", "PVerdana", based_on="Normal",
              rpr_={"rFonts": {"ascii": "Verdana", "hAnsi": "Verdana"}}),
    wml.style("paragraph", "PMajor", based_on="Normal",
              rpr_={"rFonts": {"asciiTheme": "majorHAnsi", "hAnsiTheme": "majorHAnsi"}}),
    wml.style("character", "CArial", based_on="DefaultParagraphFont",
              rpr_={"rFonts": {"ascii": "Arial", "hAnsi": "Arial"}}),
    wml.style("character", "CMajor", based_on="DefaultParagraphFont",
              rpr_={"rFonts": {"asciiTheme": "majorHAnsi", "hAnsiTheme": "majorHAnsi"}}),
]
HEBREW = "שלום"
ARABIC = "سلام"
FONT_CASES = (
    Case("f01", note="docDefaults minorHAnsi"),
    Case("f02", direct={"rFonts": {"asciiTheme": "majorHAnsi", "hAnsiTheme": "majorHAnsi"}},
         note="direct majorHAnsi"),
    Case("f03", direct={"rFonts": {"ascii": "Arial", "hAnsi": "Arial"}},
         note="direct explicit face over inherited theme attribute"),
    Case("f04", direct={"rFonts": {"ascii": "Arial", "hAnsi": "Arial",
                                   "asciiTheme": "majorHAnsi", "hAnsiTheme": "majorHAnsi"}},
         note="explicit and theme on the same element"),
    Case("f05", pstyle="PVerdana", note="paragraph style explicit face over docDefaults theme"),
    Case("f06", pstyle="PVerdana",
         direct={"rFonts": {"asciiTheme": "majorHAnsi", "hAnsiTheme": "majorHAnsi"}},
         note="direct theme over paragraph style explicit"),
    Case("f07", direct={"rFonts": {"ascii": "Arial"}}, sample="aéa",
         note="ascii slot only: a in Arial, e-acute from hAnsi"),
    Case("f08", direct={"rFonts": {"asciiTheme": "majorAscii", "hAnsiTheme": "majorAscii"}},
         note="majorAscii"),
    Case("f09", sample=HEBREW, direct={"szCs": 40}, note="Hebrew: cs slot (minorBidi), szCs"),
    Case("f10", sample=HEBREW, direct={"rFonts": {"cs": "Arial"}, "szCs": 40},
         note="Hebrew: direct cs face over inherited cstheme"),
    Case("f11", sample="“x” é § —",
         direct={"rFonts": {"ascii": "Verdana", "hAnsi": "Arial", "eastAsia": "Courier New",
                            "hint": "eastAsia"}},
         note="hint eastAsia on shared-range characters"),
    Case("f12", sample="“x” é § —",
         direct={"rFonts": {"ascii": "Verdana", "hAnsi": "Arial", "eastAsia": "Courier New"}},
         note="same, no hint"),
    Case("f13", sample=HEBREW, direct={"b": True, "szCs": 40}, note="Hebrew with b only"),
    Case("f14", sample=HEBREW, direct={"bCs": True, "szCs": 40}, note="Hebrew with bCs only"),
    Case("f15", sample="abc", direct={"cs": True, "szCs": 40},
         note="w:cs on Latin text: complex-script slot and size?"),
    Case("f16", rstyle="CArial", pstyle="PMajor", note="character explicit over paragraph theme"),
    Case("f17", rstyle="CMajor", pstyle="PVerdana", note="character theme over paragraph explicit"),
    Case("f18", sample="中文", note="CJK with an empty theme ea typeface"),
    Case("f19", sample="abc", direct={"b": True, "sz": 30, "bCs": False},
         note="Latin: b and sz apply, bCs irrelevant"),
    Case("f20", sample=HEBREW, direct={"rtl": True, "szCs": 40}, note="Hebrew with w:rtl"),
    Case("f21", sample=HEBREW, direct={"cs": True, "szCs": 40}, note="Hebrew with w:cs"),
    Case("f22", sample="abc", direct={"rFonts": {"cs": "Arial"}, "cs": True, "szCs": 40},
         note="w:cs Latin, explicit cs face"),
    Case("f23", sample="abc", direct={"rFonts": {"cstheme": "majorBidi"}, "cs": True, "szCs": 40},
         note="w:cs Latin, cstheme majorBidi"),
    Case("f24", sample=ARABIC, direct={"szCs": 40}, note="Arabic, no markers"),
    Case("f25", sample="abc", direct={"cs": True, "szCs": 40, "bCs": True, "b": False},
         note="w:cs Latin: bCs applies, b does not"),
    Case("f26", sample=HEBREW, direct={"rtl": True, "rFonts": {"cs": "Arial"}, "szCs": 40},
         note="Hebrew with w:rtl and explicit cs face"),
    Case("f27", sample=ARABIC, direct={"rtl": True, "szCs": 40}, note="Arabic with w:rtl"),
    Case("f28", sample=HEBREW, direct={"rtl": True, "szCs": 40, "lang": {"bidi": "ar-SA"}},
         note="Hebrew with w:rtl, bidi language Arabic (no Arab script font)"),
)


# -- where nothing names a face or a size ---------------------------------------

BARE_CASES = (
    Case("n01", note="no rPr anywhere"),
    Case("n02", direct={"rFonts": {"asciiTheme": "minorHAnsi", "hAnsiTheme": "minorHAnsi"}},
         note="minorHAnsi"),
    Case("n03", direct={"rFonts": {"asciiTheme": "majorHAnsi", "hAnsiTheme": "majorHAnsi"}},
         note="majorHAnsi"),
    Case("n04", direct={"rFonts": {"ascii": "Calibri"}}, note="ascii only, as in hello.docx"),
)


# -- numbering and table style against paragraph style ---------------------------

NUMBERING = wml.numbering_part([
    ("decimal", "%1.", {"ind": {"left": 720, "hanging": 360}}, {}),
])
PRECEDENCE_STYLES = [
    wml.style("paragraph", "PInd", based_on="Normal", ppr_={"ind": {"left": 2160}}),
    wml.style("paragraph", "PNum", based_on="Normal",
              ppr_={"numPr": '<w:ilvl w:val="0"/><w:numId w:val="1"/>',
                    "ind": {"left": 2160}}),
    wml.style("paragraph", "PSz", based_on="Normal", rpr_={"sz": 18}),
    wml.style("paragraph", "PNoSz", based_on="Normal", rpr_={"i": True}),
    wml.style("table", "TSz", based_on="TableNormal", rpr_={"sz": 28},
              ppr_={"ind": {"left": 1080}}, tbl=True),
]
NUM = {"numPr": '<w:ilvl w:val="0"/><w:numId w:val="1"/>'}
PRECEDENCE_CASES = (
    Case("p01", paragraph=NUM, note="numbering indent 720"),
    Case("p02", pstyle="PInd", paragraph=NUM, note="direct numbering + style indent 2160"),
    Case("p03", pstyle="PNum", note="style carrying numbering + its own indent 2160"),
    Case("p04", paragraph={**NUM, "ind": {"left": 2880}}, note="direct numbering + direct ind"),
    Case("p05", pstyle="PInd", note="style indent 2160 (control)"),
    Case("p06", table_style="TSz", note="table style sz 28 vs Normal sz 22"),
    Case("p07", table_style="TSz", pstyle="PSz", note="table sz 28 vs paragraph style sz 18"),
    Case("p08", table_style="TSz", pstyle="PNoSz", note="table sz 28 vs style without sz"),
    Case("p09", note="Normal sz 22 (control)"),
)


def probes() -> list[Probe]:
    toggles = Probe("style-toggles", TOGGLE_CASES, wml.styles_part(
        _basic_styles(TOGGLE_STYLES), run_defaults=PLAIN_DEFAULTS, paragraph_defaults=SINGLE))
    fonts = Probe("style-fonts", FONT_CASES, wml.styles_part(
        _basic_styles(FONT_STYLES),
        run_defaults={"rFonts": THEMED, "sz": 22, "szCs": 22,
                      "lang": {"val": "en-GB", "eastAsia": "zh-CN", "bidi": "he-IL"}},
        paragraph_defaults=SINGLE), extra_parts=(THEME,))
    fonts_langs = Probe("style-fonts-langs", FONT_CASES, fonts.styles, extra_parts=(
        THEME, wml.settings_part({"val": "en-GB", "eastAsia": "zh-CN", "bidi": "he-IL"})))
    fonts_nolang = Probe("style-fonts-unlisted-langs", FONT_CASES, fonts.styles, extra_parts=(
        THEME, wml.settings_part({"val": "en-GB", "eastAsia": "ja-JP", "bidi": "ar-SA"})))
    bare_styles = wml.styles_part(_basic_styles([]), paragraph_defaults=SINGLE)
    precedence = Probe("style-precedence", PRECEDENCE_CASES, wml.styles_part(
        _basic_styles(PRECEDENCE_STYLES, normal_rpr={"sz": 22}),
        run_defaults={"rFonts": CALIBRI, "sz": 24, "szCs": 24},
        paragraph_defaults=SINGLE), extra_parts=(NUMBERING,))
    return [
        toggles,
        _global_probe("style-defaults-bold", "defaults"),
        _global_probe("style-normal-bold", "normal"),
        _global_probe("style-charstyle-bold", "charstyle"),
        fonts,
        fonts_langs,
        fonts_nolang,
        Probe("style-no-styles", BARE_CASES, ""),
        Probe("style-no-theme", BARE_CASES, bare_styles),
        Probe("style-theme-no-rfonts", BARE_CASES, bare_styles, extra_parts=(THEME,)),
        Probe("style-empty-styles", BARE_CASES, wml.styles_part(_basic_styles([]))),
        precedence,
    ]
