#!/usr/bin/env python3
"""Build ``tests/fixtures/layout-sweep.docx`` -- the first fixture, and it is ours.

Why this document exists
------------------------

DOCX's hard problem is *layout you did not author*: where a line breaks and where a page
breaks are computed by Word, not written in the file.  A fixture for this project is
therefore not a document that looks interesting -- it is a document whose layout answers
are **differences that cancel everything we cannot measure**.

That is the whole design of this file.  Every block below is a set of paragraphs that
differ in exactly one authored quantity, so that subtracting one measurement from another
removes the glyph's left side bearing, the font's ascent, Word's own rounding of the page
box, and any face substitution -- none of which this project may ship a font file to
model (see ROADMAP.md, "No Microsoft font file enters this repository").  What survives
the subtraction is a number that is authored in twips in this file and must come back out
of Word's PDF in points.

Blocks, and what each one measures
----------------------------------

**A -- the indent ladder.**  Five paragraphs, every one beginning with the same glyph, at
``w:ind/@w:left`` of 0, 360, 720, 1080 and 1440 twips.  The left side bearing of that
glyph is identical in all five, so the *differences* between their first glyphs' ink-box
left edges are exactly 18, 36, 54 and 72 pt with no font knowledge whatsoever.  This is
the loop-closing measurement: it is pure layout, it is exact, and it needs nothing this
repository cannot hold.

**B -- the tab ladder.**  One paragraph with explicit ``w:tabs`` at 1440, 2880 and 4320
twips, and the same glyph after each tab.  Differences again cancel the bearing, and this
time they also test that a tab stop is measured from the *text column* origin rather than
from the page.

**C -- alignment.**  The same string set left, centred and right.  Centred and right
positions are functions of the measured string width, which we do not know a priori --
so what this block measures is the *column*: left-ink of the left-set line and right-ink
of the right-set line bracket the text column, and the centred line's midpoint must be
their midpoint.

**D -- one wrapping paragraph.**  Long enough to take several lines in the A4 column.
Nothing here is predicted yet; what it gives is the *observation* -- how many lines Word
took and which word each line ends on -- against which a future line breaker is scored.

**E -- the pagination block.**  Seventy short numbered paragraphs, each one line, which
overflow the A4 text column.  Where Word put the break is emergent: nothing in this file
says "page 2".  This is the observation the whole of Phase 3 exists to reproduce.

**F -- a second section, with a page size chosen to falsify a law.**  Word's exported
A4 page box is 595.2 x 841.92 pt, not the 595.3 x 841.9 the ``w:pgSz`` asks for.  Both
exported numbers are whole 1/300-inch device pixels (2480 and 3508, which are exactly A4
at 300 dpi).  Section 2 asks for 10000 x 13000 twips, which is 500 x 650 pt and lands on
2083.33 x 2708.33 device pixels -- so if the law holds, the exported box reads
**499.92 x 649.92 pt**, and if the page box is simply the authored size it reads
500 x 650.  One page distinguishes them.  See ROADMAP.md for the verdict.

Font policy
-----------

Every run names Calibri explicitly *and* ``w:docDefaults`` names it, because a run that
names no face inherits the theme's minor font and Word 16.106 resolves that to Aptos --
which is visible in the exported ``/BaseFont`` entries and would make this document's
measurements a measurement of font substitution.  ``tools/read_layout_sweep.py`` reports
the ``/BaseFont`` set so that a substitution cannot creep in unnoticed.  No font file is
read, shipped or committed: the measurements above are differences precisely so that none
is needed.

Acceptance test
---------------

**Word must open this document and export it.**  Hand-written OOXML that Word silently
repairs, or refuses, is worthless as a fixture no matter how well-formed it looks -- the
sibling project lost days to exactly that.  ``tests/test_oracle.py`` runs the export and
fails if the PDF does not appear, which is the only check that means anything here.

Regenerate with::

    python tools/make_layout_sweep.py

"""

from __future__ import annotations

import sys
import zipfile
from pathlib import Path

# ---------------------------------------------------------------------------
# Authored constants.  Everything the reader and the tests check is derived from
# these names rather than re-typed, so the fixture and its expectations cannot
# drift apart.
# ---------------------------------------------------------------------------

#: Section 1 page: ISO A4 as Word writes it, in twips.
A4_WIDTH_TWIPS = 11906
A4_HEIGHT_TWIPS = 16838

#: Section 2 page: deliberately *not* a standard size, and deliberately one whose
#: point size (500 x 650) is exact while its 300 dpi device size (2083.33 x 2708.33)
#: is not.  See "block F" above.
ODD_WIDTH_TWIPS = 10000
ODD_HEIGHT_TWIPS = 13000

MARGIN_TWIPS = 1440
ODD_MARGIN_TWIPS = 720

#: Block A.  The differences between consecutive entries are what gets measured.
INDENT_LADDER_TWIPS = (0, 360, 720, 1080, 1440)

#: Block B.
TAB_STOPS_TWIPS = (1440, 2880, 4320)

#: Block E.  Seventy lines overflows an A4 column of 9.69 inches at 11 pt single
#: spacing (about 50 lines), so the break lands inside this block rather than at
#: its edge -- a break at a block boundary would be consistent with several
#: different wrong models.
PAGINATION_LINES = 70

#: The glyph every position-measuring block starts with.  Any glyph works; what
#: matters is that it is the *same* one, so its left side bearing subtracts out.
PROBE_GLYPH = "H"

FACE = "Calibri"
BODY_HALF_POINTS = 22  # 11 pt
HEADING_HALF_POINTS = 28  # 14 pt

WRAPPING_TEXT = (
    "Line breaking is the first thing a word processor does that a slide renderer "
    "never has to do, and it is the reason this project exists as a sibling rather "
    "than a feature. A shape positions its text box and then fits text inside it; a "
    "document has no boxes at all, only a column and a sequence of runs, and every "
    "line is a decision about where the previous one ended."
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _escape(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _run(text: str, *, half_points: int = BODY_HALF_POINTS, bold: bool = False) -> str:
    """One ``w:r``.

    ``w:rFonts`` names the face in all four slots.  ``w:ascii`` alone is not enough:
    a character outside the ASCII range takes ``w:hAnsi``, and Word falls back to the
    theme font -- Aptos here -- for whichever slot is unset.  The probe text is ASCII,
    but naming all four means a later edit cannot reintroduce a substitution silently.
    """
    weight = "<w:b/>" if bold else ""
    return (
        "<w:r><w:rPr>"
        f'<w:rFonts w:ascii="{FACE}" w:hAnsi="{FACE}" w:cs="{FACE}" w:eastAsia="{FACE}"/>'
        f"{weight}"
        f'<w:sz w:val="{half_points}"/><w:szCs w:val="{half_points}"/>'
        "</w:rPr>"
        f'<w:t xml:space="preserve">{_escape(text)}</w:t></w:r>'
    )


def _paragraph(
    text: str,
    *,
    indent_twips: int = 0,
    alignment: str | None = None,
    tab_stops: tuple[int, ...] = (),
    half_points: int = BODY_HALF_POINTS,
    bold: bool = False,
    section_properties: str = "",
) -> str:
    """One ``w:p``.

    Child order inside ``w:pPr`` is enforced by the schema, not advisory -- Word
    repairs a document that gets it wrong, and a repaired document exports under a
    different name or not at all.  The order used here is the schema's:
    ``pStyle, keepNext, ..., tabs, ..., spacing, ind, jc, ..., sectPr, rPr``.
    """
    properties = ""
    if tab_stops:
        stops = "".join(
            f'<w:tab w:val="left" w:pos="{position}"/>' for position in tab_stops
        )
        properties += f"<w:tabs>{stops}</w:tabs>"
    # Pin the line rule rather than inherit it.  ``auto`` at 240 twentieths is
    # single spacing; leaving it unset means the measurement depends on a default
    # this file did not choose.
    properties += '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
    if indent_twips:
        properties += f'<w:ind w:left="{indent_twips}"/>'
    if alignment:
        properties += f'<w:jc w:val="{alignment}"/>'
    properties += section_properties

    runs = ""
    for index, piece in enumerate(text.split("\t")):
        if index:
            runs += "<w:r><w:tab/></w:r>"
        if piece:
            runs += _run(piece, half_points=half_points, bold=bold)

    return f"<w:p><w:pPr>{properties}</w:pPr>{runs}</w:p>"


def _section_properties(width_twips: int, height_twips: int, margin_twips: int) -> str:
    return (
        "<w:sectPr>"
        f'<w:pgSz w:w="{width_twips}" w:h="{height_twips}"/>'
        f'<w:pgMar w:top="{margin_twips}" w:right="{margin_twips}"'
        f' w:bottom="{margin_twips}" w:left="{margin_twips}"'
        ' w:header="720" w:footer="720" w:gutter="0"/>'
        "<w:cols w:space=\"708\"/>"
        '<w:docGrid w:linePitch="360"/>'
        "</w:sectPr>"
    )


def build_document_xml() -> str:
    body: list[str] = []

    body.append(_paragraph("A. Indent ladder", half_points=HEADING_HALF_POINTS, bold=True))
    for indent in INDENT_LADDER_TWIPS:
        body.append(_paragraph(f"{PROBE_GLYPH} indent {indent}", indent_twips=indent))

    body.append(_paragraph("B. Tab ladder", half_points=HEADING_HALF_POINTS, bold=True))
    body.append(
        _paragraph(
            "\t".join([PROBE_GLYPH] * (len(TAB_STOPS_TWIPS) + 1)),
            tab_stops=TAB_STOPS_TWIPS,
        )
    )

    body.append(_paragraph("C. Alignment", half_points=HEADING_HALF_POINTS, bold=True))
    for alignment in ("left", "center", "right"):
        body.append(_paragraph(f"{PROBE_GLYPH}{alignment}{PROBE_GLYPH}", alignment=alignment))

    body.append(_paragraph("D. Wrapping", half_points=HEADING_HALF_POINTS, bold=True))
    body.append(_paragraph(WRAPPING_TEXT))

    body.append(_paragraph("E. Pagination", half_points=HEADING_HALF_POINTS, bold=True))
    for number in range(1, PAGINATION_LINES + 1):
        body.append(_paragraph(f"L{number:02d}"))

    # Section 1 ends here.  A non-final section carries its ``sectPr`` inside the
    # ``w:pPr`` of its own last paragraph; only the *final* section's is a direct
    # child of ``w:body``.  Getting this backwards is one of the ways a hand-written
    # document gets repaired on open.
    body.append(
        _paragraph(
            "",
            section_properties=_section_properties(
                A4_WIDTH_TWIPS, A4_HEIGHT_TWIPS, MARGIN_TWIPS
            ),
        )
    )

    body.append(_paragraph("F. Second section", half_points=HEADING_HALF_POINTS, bold=True))
    body.append(
        _paragraph(f"{PROBE_GLYPH} page box probe", indent_twips=0)
    )
    body.append(_paragraph(f"{PROBE_GLYPH} page box probe", indent_twips=INDENT_LADDER_TWIPS[-1]))

    final_section = _section_properties(
        ODD_WIDTH_TWIPS, ODD_HEIGHT_TWIPS, ODD_MARGIN_TWIPS
    )

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{W_NS}"><w:body>'
        + "".join(body)
        + final_section
        + "</w:body></w:document>"
    )


def build_styles_xml() -> str:
    """``w:docDefaults`` only.

    No named styles: this document inherits nothing, which is the point -- every
    quantity a measurement depends on is written on the paragraph that uses it.
    Style inheritance is a Phase 2 subject and gets a fixture of its own.
    """
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:styles xmlns:w="{W_NS}">'
        "<w:docDefaults>"
        "<w:rPrDefault><w:rPr>"
        f'<w:rFonts w:ascii="{FACE}" w:hAnsi="{FACE}" w:cs="{FACE}" w:eastAsia="{FACE}"/>'
        f'<w:sz w:val="{BODY_HALF_POINTS}"/><w:szCs w:val="{BODY_HALF_POINTS}"/>'
        '<w:lang w:val="en-GB"/>'
        "</w:rPr></w:rPrDefault>"
        "<w:pPrDefault><w:pPr>"
        '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
        "</w:pPr></w:pPrDefault>"
        "</w:docDefaults>"
        "</w:styles>"
    )


CONTENT_TYPES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels"'
    ' ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml"'
    ' ContentType="application/vnd.openxmlformats-officedocument'
    '.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml"'
    ' ContentType="application/vnd.openxmlformats-officedocument'
    '.wordprocessingml.styles+xml"/>'
    "</Types>"
)

PACKAGE_RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1"'
    ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    '/officeDocument" Target="word/document.xml"/>'
    "</Relationships>"
)

DOCUMENT_RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1"'
    ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    '/styles" Target="styles.xml"/>'
    "</Relationships>"
)


def build_package() -> bytes:
    import io

    parts = {
        "[Content_Types].xml": CONTENT_TYPES_XML,
        "_rels/.rels": PACKAGE_RELS_XML,
        "word/_rels/document.xml.rels": DOCUMENT_RELS_XML,
        "word/document.xml": build_document_xml(),
        "word/styles.xml": build_styles_xml(),
    }

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, text in parts.items():
            # A fixed timestamp so regenerating an unchanged document produces an
            # unchanged file and `git status` stays quiet.  The compressed bytes are
            # *not* claimed to be stable across zlib versions; the test that checks
            # this fixture compares the parts, not the archive -- see tests/.
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            archive.writestr(info, text.encode("utf-8"))
    return buffer.getvalue()


def main(argv: list[str]) -> int:
    destination = (
        Path(argv[1])
        if len(argv) > 1
        else Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "layout-sweep.docx"
    )
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_bytes(build_package())
    print(f"wrote {destination} ({destination.stat().st_size} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
