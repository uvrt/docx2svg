"""Schema-ordered WordprocessingML fragments for probes that need real styles.

``probe_docx.py`` writes documents whose runs state everything directly, which is what
the Phase 2 measurements needed.  The style probes need the opposite -- a styles part
with ``w:basedOn`` chains, a theme, numbering, a table style -- so this module writes
those, keeping ``probe_docx``'s rule that **child order is the schema's**: callers pass
properties by element name and the order below decides where each lands.  Word repairs a
document whose children are out of order, and a repaired probe measures the repair.

Values: ``True`` writes ``<w:x/>``, ``False`` writes ``<w:x w:val="0"/>``, a string or
integer writes ``w:val``, and a dict writes those attributes (``rFonts``, ``spacing``,
``ind``...).
"""

from __future__ import annotations

from probe_docx import W_NS, escape

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

#: ECMA-376 Part 1, 17.3.2.28 (CT_RPr / EG_RPrBase), in order.
RPR_ORDER = (
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
    "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
    "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
    "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
    "specVanish",
)
#: 17.3.1.26 (CT_PPrBase), in order; ``rPr`` and ``sectPr`` follow it in a ``w:pPr``.
PPR_ORDER = (
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
    "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku",
    "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi",
    "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap",
    "outlineLvl",
)


def _element(name: str, value) -> str:
    if value is True:
        return f"<w:{name}/>"
    if value is False:
        return f'<w:{name} w:val="0"/>'
    if isinstance(value, dict):
        attrs = "".join(f' w:{k}="{v}"' for k, v in value.items())
        return f"<w:{name}{attrs}/>"
    if isinstance(value, str) and value.startswith("<"):
        return f"<w:{name}>{value}</w:{name}>"
    return f'<w:{name} w:val="{value}"/>'


def _ordered(order: tuple[str, ...], props: dict) -> str:
    unknown = set(props) - set(order)
    if unknown:
        raise ValueError(f"no schema position recorded for {sorted(unknown)}")
    return "".join(_element(name, props[name]) for name in order if name in props)


def rpr(**props) -> str:
    return f"<w:rPr>{_ordered(RPR_ORDER, props)}</w:rPr>" if props else ""


def ppr(*, rpr_: str = "", sect: str = "", **props) -> str:
    body = _ordered(PPR_ORDER, props) + rpr_ + sect
    return f"<w:pPr>{body}</w:pPr>" if body else ""


def run(text: str, **props) -> str:
    content = "<w:tab/>" if text == "\t" else f'<w:t xml:space="preserve">{escape(text)}</w:t>'
    return f"<w:r>{rpr(**props)}{content}</w:r>"


def paragraph(runs: str, *, mark: dict | None = None, sect: str = "", **props) -> str:
    return f"<w:p>{ppr(rpr_=rpr(**(mark or {})), sect=sect, **props)}{runs}</w:p>"


def table(rows: list[list[str]], *, style: str | None = None, width: int = 4000) -> str:
    """A fixed-width table; each cell is a string of ``w:p`` elements."""
    columns = len(rows[0])
    tbl_pr = (
        (f'<w:tblStyle w:val="{style}"/>' if style else "")
        + f'<w:tblW w:w="{width * columns}" w:type="dxa"/>'
        + '<w:tblLayout w:type="fixed"/>'
        + '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0"'
        ' w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>'
    )
    grid = "".join(f'<w:gridCol w:w="{width}"/>' for _ in range(columns))
    body = "".join(
        "<w:tr>"
        + "".join(f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>{cell}</w:tc>' for cell in row)
        + "</w:tr>"
        for row in rows
    )
    return f"<w:tbl><w:tblPr>{tbl_pr}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{body}</w:tbl>"


def style(kind: str, style_id: str, *, name: str | None = None, based_on: str | None = None,
          next_: str | None = None, link: str | None = None, default: bool = False,
          ppr_: dict | None = None, rpr_: dict | None = None, tbl: bool = False) -> str:
    """One ``w:style``, children in CT_Style order."""
    attrs = f' w:type="{kind}"' + (' w:default="1"' if default else "") + f' w:styleId="{style_id}"'
    body = f'<w:name w:val="{name or style_id}"/>'
    if based_on:
        body += f'<w:basedOn w:val="{based_on}"/>'
    if next_:
        body += f'<w:next w:val="{next_}"/>'
    if link:
        body += f'<w:link w:val="{link}"/>'
    body += "<w:qFormat/>"
    if ppr_:
        body += f"<w:pPr>{_ordered(PPR_ORDER, ppr_)}</w:pPr>"
    if rpr_:
        body += rpr(**rpr_)
    if tbl:
        body += (
            '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar>'
            '<w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
            '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>'
            "</w:tblCellMar></w:tblPr>"
        )
    return f"<w:style{attrs}>{body}</w:style>"


def styles_part(styles: list[str], *, run_defaults: dict | None = None,
                paragraph_defaults: dict | None = None) -> str:
    defaults = ""
    if run_defaults is not None or paragraph_defaults is not None:
        defaults = "<w:docDefaults>"
        if run_defaults is not None:
            defaults += f"<w:rPrDefault>{rpr(**run_defaults) or '<w:rPr/>'}</w:rPrDefault>"
        if paragraph_defaults is not None:
            body = _ordered(PPR_ORDER, paragraph_defaults)
            defaults += f"<w:pPrDefault><w:pPr>{body}</w:pPr></w:pPrDefault>"
        defaults += "</w:docDefaults>"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:styles xmlns:w="{W_NS}">{defaults}{"".join(styles)}</w:styles>'
    )


THEME_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.theme+xml"
THEME_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
NUMBERING_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
)
NUMBERING_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"


def theme_part(major: dict[str, str], minor: dict[str, str],
               major_scripts: dict[str, str] | None = None,
               minor_scripts: dict[str, str] | None = None) -> tuple[str, str, str, str]:
    """A complete ``theme1.xml``: colour scheme, font scheme and format scheme.

    ``major``/``minor`` map ``latin``/``ea``/``cs`` to a typeface (``""`` is legal and
    means "none"); ``*_scripts`` map a script tag (``Hebr``, ``Hans``) to a typeface.  The colour and format schemes are required by the schema and are
    written in full, minimally, so Word has nothing to repair.
    """
    def fonts(tag: str, faces: dict[str, str], scripts: dict[str, str] | None) -> str:
        return (
            f"<a:{tag}>"
            + "".join(f'<a:{slot} typeface="{faces.get(slot, "")}"/>' for slot in ("latin", "ea", "cs"))
            + "".join(f'<a:font script="{k}" typeface="{v}"/>' for k, v in (scripts or {}).items())
            + f"</a:{tag}>"
        )

    colours = "".join(
        f'<a:{name}><a:srgbClr val="{value}"/></a:{name}>'
        for name, value in (
            ("dk1", "000000"), ("lt1", "FFFFFF"), ("dk2", "44546A"), ("lt2", "E7E6E6"),
            ("accent1", "4472C4"), ("accent2", "ED7D31"), ("accent3", "A5A5A5"),
            ("accent4", "FFC000"), ("accent5", "5B9BD5"), ("accent6", "70AD47"),
            ("hlink", "0563C1"), ("folHlink", "954F72"),
        )
    )
    solid = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    line = f'<a:ln w="6350">{solid}</a:ln>'
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<a:theme xmlns:a="{A_NS}" name="Probe"><a:themeElements>'
        f'<a:clrScheme name="Probe">{colours}</a:clrScheme>'
        f'<a:fontScheme name="Probe">{fonts("majorFont", major, major_scripts)}{fonts("minorFont", minor, minor_scripts)}</a:fontScheme>'
        '<a:fmtScheme name="Probe">'
        f"<a:fillStyleLst>{solid * 3}</a:fillStyleLst>"
        f"<a:lnStyleLst>{line * 3}</a:lnStyleLst>"
        f"<a:effectStyleLst>{'<a:effectStyle><a:effectLst/></a:effectStyle>' * 3}</a:effectStyleLst>"
        f"<a:bgFillStyleLst>{solid * 3}</a:bgFillStyleLst>"
        "</a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>"
    )
    return ("word/theme/theme1.xml", THEME_CONTENT_TYPE, THEME_REL, xml)


def numbering_part(levels: list[tuple[str, str, dict, dict]]) -> tuple[str, str, str, str]:
    """One abstract numbering, ``w:num`` 1 -> it.  ``levels``: (numFmt, lvlText, pPr, rPr)."""
    body = ""
    for index, (fmt, text, level_ppr, level_rpr) in enumerate(levels):
        body += (
            f'<w:lvl w:ilvl="{index}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
            f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
            f"<w:pPr>{_ordered(PPR_ORDER, level_ppr)}</w:pPr>{rpr(**level_rpr)}</w:lvl>"
        )
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:numbering xmlns:w="{W_NS}">'
        f'<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>{body}</w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
        "</w:numbering>"
    )
    return ("word/numbering.xml", NUMBERING_CONTENT_TYPE, NUMBERING_REL, xml)


SETTINGS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
)
SETTINGS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"


def settings_part(theme_font_lang: dict[str, str], *,
                  compatibility_mode: int | None = None, compat_elements: str = "",
                  empty_compat: bool = False) -> tuple[str, str, str, str]:
    """A ``settings.xml`` holding ``w:themeFontLang`` and, if given, the compatibility mode.

    ``compatibility_mode`` (``w:compatSetting compatibilityMode``) is the one setting
    here that moves layout; without it Word treats the document as it treats every
    Phase 2 probe, which carries no ``settings.xml`` at all.  ``compat_elements`` are
    legacy ``w:compat`` children, written before the ``w:compatSetting`` as the schema
    orders them (``<w:suppressSpBfAfterPgBrk/>``); ``empty_compat`` writes ``<w:compat/>``
    when nothing else goes in it, which is what Word 12 writes.
    """
    attrs = "".join(f' w:{k}="{v}"' for k, v in theme_font_lang.items())
    inner = compat_elements
    if compatibility_mode is not None:
        inner += (
            '<w:compatSetting w:name="compatibilityMode"'
            ' w:uri="http://schemas.microsoft.com/office/word"'
            f' w:val="{compatibility_mode}"/>'
        )
    compat = f"<w:compat>{inner}</w:compat>" if inner else ("<w:compat/>" if empty_compat else "")
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:settings xmlns:w="{W_NS}">{compat}<w:themeFontLang{attrs}/></w:settings>'
    )
    return ("word/settings.xml", SETTINGS_CONTENT_TYPE, SETTINGS_REL, xml)
