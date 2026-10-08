"""The smallest WordprocessingML writer the Phase 2 probes need.

``make_layout_sweep.py`` carries its own writer because it was the first fixture and its
docstring is its specification.  The Phase 2 probes vary *one* property at a time over
many paragraphs -- face, size, line rule, indent, tab stop, colour, kerning, tracking --
so they share this module rather than each re-growing the same helpers.

Two rules from Phase 0 are enforced here rather than remembered at each call site:

* **Child order is the schema's.**  Word *repairs* a document with ``w:pPr`` or ``w:rPr``
  children out of order, and a repaired document exports under another name or not at
  all.  Callers pass properties by name; this module emits them in schema order.
* **The face is named in all four ``w:rFonts`` slots and in ``w:docDefaults``.**  A slot
  left unset falls back to the theme's minor font (Aptos on Word 16.106), and then a
  measurement is a measurement of substitution.

No ``settings.xml`` is written, matching the Phase 0 fixture: ``w:compat`` is held fixed
at "whatever Word assumes when the part is absent" across every probe, which is a
variable held fixed and is recorded as such in ROADMAP.md.
"""

from __future__ import annotations

import io
import zipfile

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

#: ECMA-376 Part 1, 17.3.1.26 (CT_PPrBase) -- the subset the probes use, in order.
_PPR_ORDER = (
    "keepNext", "keepLines", "pageBreakBefore", "widowControl", "tabs", "spacing", "ind", "jc",
)
#: 17.3.2.28 (CT_RPr) -- the subset the probes use, in order.
_RPR_ORDER = ("rFonts", "b", "color", "spacing", "w", "kern", "sz", "szCs", "lang")


def escape(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _ordered(order: tuple[str, ...], parts: dict[str, str]) -> str:
    unknown = set(parts) - set(order)
    if unknown:
        raise ValueError(f"no schema position recorded for {sorted(unknown)}")
    return "".join(parts[name] for name in order if name in parts)


def rpr(face: str, half_points: int, *, bold: bool = False, color: str | None = None,
        spacing_twips: int | None = None, kern_half_points: int | None = None) -> str:
    parts = {
        "rFonts": f'<w:rFonts w:ascii="{face}" w:hAnsi="{face}" w:cs="{face}" w:eastAsia="{face}"/>',
        "sz": f'<w:sz w:val="{half_points}"/>',
        "szCs": f'<w:szCs w:val="{half_points}"/>',
    }
    if bold:
        parts["b"] = "<w:b/>"
    if color:
        parts["color"] = f'<w:color w:val="{color}"/>'
    if spacing_twips is not None:
        parts["spacing"] = f'<w:spacing w:val="{spacing_twips}"/>'
    if kern_half_points is not None:
        parts["kern"] = f'<w:kern w:val="{kern_half_points}"/>'
    return "<w:rPr>" + _ordered(_RPR_ORDER, parts) + "</w:rPr>"


def run(text: str, face: str, half_points: int, **props) -> str:
    if text == "\t":
        return f"<w:r>{rpr(face, half_points, **props)}<w:tab/></w:r>"
    return (
        f"<w:r>{rpr(face, half_points, **props)}"
        f'<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'
    )


def paragraph(runs: str, *, line: int = 240, line_rule: str = "auto", indent_left: int = 0,
              first_line: int | None = None, tabs: tuple[int, ...] = (),
              jc: str | None = None, page_break_before: bool = False,
              mark_rpr: str = "", section: str = "") -> str:
    """One ``w:p``.  ``mark_rpr`` is the paragraph mark's own ``w:rPr``.

    The paragraph mark has a size too, and it takes part in the line height: an empty
    mark rPr means the docDefaults size, which would make a line of 24 pt text also
    carry an 11 pt mark.  Probes that vary the size pass the same rPr for the mark.
    """
    parts: dict[str, str] = {}
    if page_break_before:
        parts["pageBreakBefore"] = "<w:pageBreakBefore/>"
    if tabs:
        stops = "".join(f'<w:tab w:val="left" w:pos="{pos}"/>' for pos in tabs)
        parts["tabs"] = f"<w:tabs>{stops}</w:tabs>"
    parts["spacing"] = (
        f'<w:spacing w:before="0" w:after="0" w:line="{line}" w:lineRule="{line_rule}"/>'
    )
    if indent_left or first_line is not None:
        first = f' w:firstLine="{first_line}"' if first_line is not None else ""
        parts["ind"] = f'<w:ind w:left="{indent_left}"{first}/>'
    if jc:
        parts["jc"] = f'<w:jc w:val="{jc}"/>'
    # sectPr and the mark's rPr come after every CT_PPrBase child, in that order.
    ppr = _ordered(_PPR_ORDER, parts) + section + mark_rpr
    return f"<w:p><w:pPr>{ppr}</w:pPr>{runs}</w:p>"


def section(width: int = 11906, height: int = 16838, margin: int = 1440, *,
            left: int | None = None) -> str:
    left = margin if left is None else left
    return (
        "<w:sectPr>"
        f'<w:pgSz w:w="{width}" w:h="{height}"/>'
        f'<w:pgMar w:top="{margin}" w:right="{margin}" w:bottom="{margin}" w:left="{left}"'
        ' w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="708"/>'
        '<w:docGrid w:linePitch="360"/>'
        "</w:sectPr>"
    )


def package(body: str, *, final_section: str | None = None, face: str = "Calibri",
            half_points: int = 22, styles: str | None = None,
            extra_parts: tuple[tuple[str, str, str, str], ...] = ()) -> bytes:
    """A complete ``.docx``.

    ``styles`` replaces the default styles part (a docDefaults naming ``face``); pass
    ``""`` for no styles part at all.  ``extra_parts`` are ``(part name, content type,
    relationship type, xml)`` related from ``word/document.xml`` -- a theme, numbering.
    """
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{W_NS}"><w:body>'
        + body
        + (final_section if final_section is not None else section())
        + "</w:body></w:document>"
    )
    default_styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:styles xmlns:w="{W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>'
        f'<w:rFonts w:ascii="{face}" w:hAnsi="{face}" w:cs="{face}" w:eastAsia="{face}"/>'
        f'<w:sz w:val="{half_points}"/><w:szCs w:val="{half_points}"/>'
        '<w:lang w:val="en-GB"/>'
        "</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>"
        '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
        "</w:pPr></w:pPrDefault></w:docDefaults></w:styles>"
    )
    if styles is None:
        styles = default_styles
    overrides = ""
    rels = ""
    if styles:
        overrides += (
            '<Override PartName="/word/styles.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        )
        rels += (
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/'
            '2006/relationships/styles" Target="styles.xml"/>'
        )
    for index, (name, content_type, rel_type, _) in enumerate(extra_parts, start=2):
        overrides += f'<Override PartName="/{name}" ContentType="{content_type}"/>'
        target = name.split("/", 1)[1]
        rels += f'<Relationship Id="rId{index}" Type="{rel_type}" Target="{target}"/>'
    parts = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels"'
            ' ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            + overrides
            + "</Types>"
        ),
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/'
            '2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
        ),
        "word/_rels/document.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + rels
            + "</Relationships>"
        ),
        "word/document.xml": document,
    }
    if styles:
        parts["word/styles.xml"] = styles
    for name, _, _, xml in extra_parts:
        parts[name] = xml
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, text in parts.items():
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            archive.writestr(info, text.encode("utf-8"))
    return buffer.getvalue()
