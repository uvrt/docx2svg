"""The layout carries on past what it cannot lay out (round 16): a placeholder of the item's
estimated height, or an approximate rule, each reported -- never a document that stops at
its first unsupported item.  Also: an absent face's default substitute, a slot that names
no face, a character a face lacks, VML (a watermark's text path, an inline picture), a
picture's crop, and a cell's footnote on a page shared with a section of several columns.

Every document here is made up, built by :func:`_docx`, set in Calibri: laid out with
Calibri where it is installed, Carlito (CI installs it) where not.
"""

from __future__ import annotations

import io
import struct
import sys
import zipfile
import zlib
from pathlib import Path
from xml.etree import ElementTree

import pytest

REPO = Path(__file__).resolve().parent.parent
FIXTURES = REPO / "tests" / "fixtures"
sys.path.insert(0, str(REPO / "tools"))

import docx2svg  # noqa: E402
from docx2svg import ConvertOptions, convert_docx_to_layout, convert_docx_to_svg, fonts, paginate  # noqa: E402
from docx2svg.parse import parse_package  # noqa: E402
from docx2svg.resolve import character_format, resolve_run  # noqa: E402

SVG = "{http://www.w3.org/2000/svg}"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = (f'xmlns:w="{W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
      'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" '
      'xmlns:w10="urn:schemas-microsoft-com:office:word" '
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"')
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
CT = "application/vnd.openxmlformats-officedocument.wordprocessingml."
TEXT = ("A made-up paragraph long enough to wrap across more than one line of the page, about nothing "
        "in particular, written for this test. ")


def options(**kwargs) -> ConvertOptions:
    """The installed faces: Calibri, or Carlito in its place (CI installs it)."""
    return ConvertOptions(**kwargs)


def _p(text: str = "", ppr: str = "", rpr: str = "") -> str:
    run = (f'<w:r>{"<w:rPr>" + rpr + "</w:rPr>" if rpr else ""}<w:t xml:space="preserve">{text}</w:t></w:r>'
           if text else "")
    return f'<w:p>{"<w:pPr>" + ppr + "</w:pPr>" if ppr else ""}{run}</w:p>'


def _text(n: int) -> str:
    return "".join(_p(f"Body {k}. " + TEXT * 3) for k in range(n))


def _tbl(rows: list[list[str]], *, width: int = 9000, float_pr: dict | None = None, layout: str = "fixed") -> str:
    floating = ("<w:tblpPr " + " ".join(f'w:{k}="{v}"' for k, v in float_pr.items()) + "/>") if float_pr else ""
    borders = "<w:tblBorders>" + "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                                         for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
    cols = len(rows[0])
    out = (f'<w:tbl><w:tblPr>{floating}<w:tblW w:w="{width}" w:type="dxa"/>{borders}</w:tblBorders>'
           f'<w:tblLayout w:type="{layout}"/></w:tblPr><w:tblGrid>'
           + "".join(f'<w:gridCol w:w="{width // cols}"/>' for _ in range(cols)) + "</w:tblGrid>")
    for row in rows:
        out += "<w:tr>" + "".join(f'<w:tc><w:tcPr><w:tcW w:w="{width // cols}" w:type="dxa"/></w:tcPr>{c}</w:tc>'
                                  for c in row) + "</w:tr>"
    return out + "</w:tbl>"


def _sect(header: str | None = None, cols: int = 1, kind: str | None = None) -> str:
    reference = f'<w:headerReference w:type="default" r:id="{header}"/>' if header else ""
    start = f'<w:type w:val="{kind}"/>' if kind else ""
    return (f'<w:sectPr>{reference}{start}<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" '
            f'w:gutter="0"/><w:cols w:num="{cols}" w:space="708"/></w:sectPr>')


def _png() -> bytes:
    """A 40 x 20 made-up picture: a red quarter, a green one, a blue half."""
    rows = b"".join(b"\0" + b"".join(b"\xff\0\0" if x < 10 else b"\0\x80\0" if x < 20 else b"\0\0\xff"
                                     for x in range(40)) for _ in range(20))

    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 40, 20, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b""))


def _docx(body: str, *, final: str | None = None, headers: dict | None = None, notes: dict | None = None,
          face: str = "Calibri", theme: str | None = None, font_table: str | None = None,
          picture: bool = False) -> bytes:
    """A made-up ``.docx``: ``body`` in a document whose default face is ``face`` (or the
    theme's minor font, given ``theme``), with headers (``{rid: xml}``), footnotes and a
    picture (``rIdm1``) if asked."""
    fonts_xml = (f'<w:rFonts w:ascii="{face}" w:hAnsi="{face}" w:eastAsia="{face}" w:cs="{face}"/>' if theme is None
                 else '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>')
    parts = {
        "word/document.xml": f'<w:document {NS}><w:body>{body}{final or _sect()}</w:body></w:document>',
        "word/styles.xml": (f'<w:styles xmlns:w="{W}"><w:docDefaults><w:rPrDefault><w:rPr>{fonts_xml}'
                            '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>'
                            '<w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
                            '</w:docDefaults><w:style w:type="character" w:styleId="FootnoteReference"><w:name '
                            'w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
                            '</w:style></w:styles>'),
    }
    rels = [("rIdS", "styles", "styles.xml")]
    types = [("/word/styles.xml", CT + "styles+xml")]
    if theme is not None:
        parts["word/theme/theme1.xml"] = (
            '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements>'
            '<a:clrScheme name="T"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>'
            '</a:clrScheme><a:fontScheme name="T">'
            f'<a:majorFont><a:latin typeface="{theme}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
            f'<a:minorFont><a:latin typeface="{theme}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
            '</a:fontScheme></a:themeElements></a:theme>')
        rels.append(("rIdT", "theme", "theme/theme1.xml"))
        types.append(("/word/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml"))
    if font_table is not None:
        parts["word/fontTable.xml"] = f'<w:fonts xmlns:w="{W}">{font_table}</w:fonts>'
        rels.append(("rIdN", "fontTable", "fontTable.xml"))
        types.append(("/word/fontTable.xml", CT + "fontTable+xml"))
    if notes:
        separators = ('<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
                      '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/>'
                      '</w:r></w:p></w:footnote>')
        parts["word/footnotes.xml"] = f'<w:footnotes {NS}>{separators}' + "".join(
            f'<w:footnote w:id="{k}"><w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/>'
            f'</w:r><w:r><w:t xml:space="preserve"> {text}</w:t></w:r></w:p></w:footnote>'
            for k, text in notes.items()) + "</w:footnotes>"
        rels.append(("rIdF", "footnotes", "footnotes.xml"))
        types.append(("/word/footnotes.xml", CT + "footnotes+xml"))
    for rid, xml in (headers or {}).items():
        name = f"header{rid[-1]}.xml"
        parts[f"word/{name}"] = xml
        rels.append((rid, "header", name))
        types.append((f"/word/{name}", CT + "header+xml"))
    if picture:
        parts["word/media/p.png"] = _png()
        rels.append(("rIdm1", "image", "media/p.png"))
    parts["word/_rels/document.xml.rels"] = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(f'<Relationship Id="{rid}" Type="{REL}{kind}" Target="{target}"/>' for rid, kind, target in rels)
        + "</Relationships>")
    parts["[Content_Types].xml"] = (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
        f'<Override PartName="/word/document.xml" ContentType="{CT}document.main+xml"/>'
        + "".join(f'<Override PartName="{name}" ContentType="{kind}"/>' for name, kind in types) + "</Types>")
    parts["_rels/.rels"] = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{REL}officeDocument" Target="word/document.xml"/></Relationships>')
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data if isinstance(data, bytes) else data.encode("utf-8"))
    return buffer.getvalue()


# -- the constructs ----------------------------------------------------------------------------

NESTED_FLOATING = _tbl([[_p("Cover title") + _tbl([[_p("Inner A"), _p("Inner B")]], width=3000, float_pr={
    "vertAnchor": "page", "horzAnchor": "page", "tblpX": 2000, "tblpY": 3000}) + _p("Below the inner table")]])
#: A VML object that states no size: what no rule lays out.
UNSIZED_VML = '<w:p><w:r><w:pict><v:shape style="position:absolute"><v:fill/></v:shape></w:pict></w:r></w:p>'
WATERMARK = (
    '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Watermarks"/></w:docPartObj></w:sdtPr><w:sdtContent>'
    '<w:p><w:r><w:pict><v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136"><v:textpath on="t" '
    'fitshape="t"/></v:shapetype><v:shape id="WaterMark1" type="#_x0000_t136" style="position:absolute;'
    'margin-left:0;margin-top:0;width:400pt;height:200pt;rotation:315;z-index:-251657216;'
    'mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;'
    'mso-position-vertical-relative:margin" fillcolor="silver" stroked="f"><v:fill opacity=".5"/>'
    '<v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="DRAFT"/></v:shape></w:pict>'
    '</w:r></w:p></w:sdtContent></w:sdt>')


def _cell_note(note: int = 1) -> str:
    return (f'<w:p><w:r><w:t xml:space="preserve">Reviewed</w:t></w:r><w:r><w:rPr><w:rStyle '
            f'w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="{note}"/></w:r></w:p>')


def _picture(crop: str, *, anchored: bool) -> str:
    graphic = ('<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
               '<pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill>'
               f'<a:blip r:embed="rIdm1"/><a:srcRect {crop}/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
               '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="1500000"/></a:xfrm><a:prstGeom '
               'prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>')
    if anchored:
        return ('<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
                'relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" '
                'y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>'
                '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
                '<wp:extent cx="3000000" cy="1500000"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
                f'<wp:wrapTopAndBottom/><wp:docPr id="1" name="p"/><wp:cNvGraphicFramePr/>{graphic}</wp:anchor>'
                '</w:drawing></w:r></w:p>')
    return ('<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="3000000" '
            'cy="1500000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="2" name="p"/>'
            f'<wp:cNvGraphicFramePr/>{graphic}</wp:inline></w:drawing></w:r></w:p>')


def _last_text(layout) -> str:
    line = layout.pages[-1].lines[-1] if layout.pages[-1].lines else None
    return "".join("".join(span.chars) for span in line.spans) if line else ""


# -- carrying on -------------------------------------------------------------------------------


def test_a_document_with_one_of_each_unsupported_construct_renders_every_page():
    """A nested floating table, a VML object with no size (body and cell), a frame, a cell
    whose face has no metrics: every page is laid out, the last paragraph drawn, and each
    reported -- placeholders by element path and page."""
    body = (_p("Before") + NESTED_FLOATING + _text(6) + UNSIZED_VML + _text(6)
            + _tbl([[UNSIZED_VML, _p("Beside it")]]) + _p("Framed", '<w:framePr w:w="2000" w:x="3000" w:y="200"/>')
            + _text(6) + _p("The very last paragraph."))
    opts = options()
    layout = convert_docx_to_layout(_docx(body), opts)
    codes = [w.code for w in opts.warnings]
    assert not [code for code in codes if code.startswith(("layout-stopped", "story-stopped"))]
    assert len(layout.pages) >= 3 and _last_text(layout) == "The very last paragraph."
    assert codes.count("layout-placeholder:drawing") == 2
    placeholders = {entry.path for entry in layout.coverage.placeholders}
    assert placeholders == {"w:body/w:p[8]", "w:body/w:tbl[2]/w:tr[1]/w:tc[1]/w:p[1]"}
    assert "layout-approximate:frame" in codes and "layout-approximate:table" in codes
    coverage = layout.coverage
    assert coverage.stop is None and not coverage.complete and coverage.status == "partial"
    assert coverage.estimated_pages == len(layout.pages) and coverage.blocks_skipped == 1
    assert "placeholder" in coverage.summary()


def test_a_placeholder_is_a_marked_box_of_its_estimated_height():
    documents = convert_docx_to_svg(_docx(_text(2) + UNSIZED_VML + _text(2)), options())
    root = ElementTree.fromstring(documents[0])
    boxes = [g for g in root.iter(f"{SVG}g") if g.get("data-docx-placeholder") == "1"]
    assert len(boxes) == 1 and boxes[0].get("data-docx-path") == "w:body/w:p[3]"
    assert boxes[0].get("data-docx-unsupported") == "drawing"
    rect = boxes[0].find(f"{SVG}rect")
    assert rect.get("stroke-dasharray") == "12 8" and float(rect.get("height")) > 0
    # The paragraphs after it are drawn below it.
    paths = [g.get("data-docx-path") for g in root.iter(f"{SVG}g") if g.get("data-docx-line") is None]
    assert "w:body/w:p[5]" in paths


def test_without_carrying_on_the_layout_stops_as_before(monkeypatch):
    monkeypatch.setattr(paginate, "CARRY_ON", False)
    opts = options()
    layout = convert_docx_to_layout(_docx(_text(2) + UNSIZED_VML + _text(20)), opts)
    assert "layout-stopped:drawing" in [w.code for w in opts.warnings]
    assert len(layout.pages) == 1 and layout.coverage.stop.path == "w:body/w:p[3]"


def test_a_nested_floating_table_is_laid_out_in_its_cell():
    opts = options()
    layout = convert_docx_to_layout(_docx(_p("Before") + NESTED_FLOATING + _text(30)), opts)
    lines = {"".join("".join(s.chars) for s in line.spans): line for page in layout.pages for line in page.lines}
    assert {"Cover title", "Inner A", "Inner B", "Below the inner table"} <= set(lines)
    # Against its cell, as Word draws it: 3,000 twips below the row's top, 2,000 right of the
    # cell's edge, taking no room in the cell's text; the row as tall as it.
    assert lines["Below the inner table"].baseline < lines["Inner A"].baseline
    assert lines["Inner A"].spans[0].xs[0] - lines["Cover title"].spans[0].xs[0] == pytest.approx(2000 / 20 * 300 / 72,
                                                                                                abs=3)
    first_body = next(line for line in layout.pages[0].lines if "".join(line.spans[0].chars).startswith("Body 0"))
    assert first_body.baseline > lines["Inner A"].baseline
    assert len(layout.pages) > 1 and layout.coverage.stop is None and layout.coverage.status == "approximate"
    assert any("positioned against its cell" in w.message for w in opts.warnings
               if w.code == "layout-approximate:table")


def test_a_table_row_taller_than_the_page_is_drawn_whole_and_the_rest_goes_on():
    tall = "".join(_p(f"Line {k}") for k in range(80))
    body = _tbl([[tall]]).replace("<w:tr>", '<w:tr><w:trPr><w:cantSplit/></w:trPr>', 1) + _p("After the table")
    big = _picture("", anchored=False).replace('cx="3000000" cy="1500000"', 'cx="3000000" cy="12000000"')
    opts = options()
    layout = convert_docx_to_layout(_docx(_tbl([[big]]) + _p("After the table"), picture=True), opts)
    assert _last_text(layout) == "After the table" and layout.coverage.stop is None
    assert any("not one line" in w.message for w in opts.warnings if w.code == "layout-approximate:table")
    assert convert_docx_to_layout(_docx(body), options()).coverage.stop is None


# -- faces ---------------------------------------------------------------------------------------


@pytest.fixture
def index(monkeypatch):
    """An installed index of our own: the faces it names point at one real file (any
    face this machine has), so what is found is decided by the names alone."""
    real = fonts.installed_index()
    if not real:
        pytest.skip("no installed face to stand in")
    entry = next(iter(real.values()))
    names = {"calibri", "carlito", "liberation serif", "arial"}

    def fake(dirs=(), *, bundle_first=False):
        return {(name, bold, italic): entry for name in names for bold in (False, True) for italic in (False, True)}

    monkeypatch.setattr(fonts, "installed_index", fake)
    return names


def test_an_absent_face_takes_the_face_powerpoint_rendering_uses_then_its_clone(index):
    found = fonts.InstalledFonts(substitutes={})
    assert found.face("Calibri Light") is not None
    substitution = found.substitutions["Calibri Light"]
    assert (substitution.substitute, substitution.kind, substitution.metric_compatible) == ("Carlito", "default",
                                                                                           False)
    index.discard("carlito")
    found = fonts.InstalledFonts(substitutes={})
    found.face("Calibri Light")
    assert found.substitutions["Calibri Light"].substitute == "Calibri"
    assert "approximation" in found.substitutions["Calibri Light"].message()


def test_an_absent_face_with_no_rule_takes_the_first_face_of_its_class(index):
    table = ('<w:font w:name="Nofont Book"><w:family w:val="roman"/></w:font>'
             '<w:font w:name="Nofont Code"><w:family w:val="modern"/><w:pitch w:val="fixed"/></w:font>')
    package = _docx(_p("x"), font_table=table)
    found = fonts.InstalledFonts(package, substitutes={})
    found.face("Nofont Book")
    found.face("Nofont Sans")
    assert found.substitutions["Nofont Book"].substitute == "Liberation Serif"
    assert found.substitutions["Nofont Book"].face_class == "serif"
    assert found.substitutions["Nofont Sans"].substitute == "Arial"
    assert found.face("Nofont Code") is None and found.missing == {"Nofont Code"}  # no monospace face here
    assert fonts.InstalledFonts(package, substitutes={}, fallbacks=False).face("Nofont Book") is None


def test_named_substitutes_are_followed_to_an_installed_face(index):
    index.discard("calibri")
    found = fonts.InstalledFonts(substitutes={"calibri light": "Calibri", "calibri": "Carlito"})
    found.face("Calibri Light")
    assert found.substitutions["Calibri Light"].substitute == "Carlito"
    assert found.substitutions["Calibri Light"].kind == "named"


def test_a_character_the_face_lacks_is_measured_in_another_and_reported():
    found = fonts.InstalledFonts()
    family = next((name for name in ("Calibri", "Carlito", "Arial", "Liberation Sans") if found.face(name)), None)
    if family is None:
        pytest.skip("no face to measure with")
    char = "\U0001D11E"  # a musical symbol no text face has
    assert found.face(family).advance(char) is None
    assert found.advance(family, False, False, char) is not None
    assert (family, char) in found.glyph_fallbacks


def test_a_slot_that_names_no_face_is_drawn_in_word_s_fallback():
    """A complex-script run whose ``minorBidi`` reference has no ``w:themeFontLang``: Times
    New Roman, as Word drew it (``make_style_probe.py``, f15)."""
    body = '<w:p><w:r><w:rPr><w:rtl/></w:rPr><w:t>abc</w:t></w:r></w:p>'
    document = parse_package(_docx(body, theme="Calibri"))
    paragraph = document.body[0]
    fmt = character_format(resolve_run(document, paragraph, paragraph.runs[0]), "a", document)
    assert (fmt.slot, fmt.face) == ("cs", "Times New Roman")


# -- VML -----------------------------------------------------------------------------------------


def test_a_watermark_is_drawn_behind_the_text_on_every_page_and_its_header_goes_on():
    header = f'<w:hdr {NS}>{WATERMARK}{_p("Header after the watermark")}</w:hdr>'
    data = _docx(_text(30), headers={"rIdh1": header}, final=_sect("rIdh1"))
    opts = options()
    documents = convert_docx_to_svg(data, opts)
    assert len(documents) > 1
    assert not [w for w in opts.warnings if w.code.startswith(("story-stopped", "layout-placeholder"))]
    for document in documents:
        root = ElementTree.fromstring(document)
        floats = [g for g in root.iter(f"{SVG}g") if g.get("data-docx-float") == "behind"]
        assert floats and floats[0].get("data-docx-story") == "header"
        text = floats[0].find(f".//{SVG}text")
        assert text.text == "DRAFT" and text.get("fill") == "#C0C0C0" and text.get("fill-opacity") == "0.5"
        assert "rotate(315)" in floats[0].find(f"{SVG}g").get("transform")
        # It fits its shape (its shape type's fitshape): stretched by its face's outlines
        # where the face is at hand, else by its advances, reported.
        fitted = (text.get("transform") or "").startswith("matrix(")
        assert fitted or text.get("lengthAdjust") == "spacingAndGlyphs"
        assert any(g.get("data-docx-path") == "w:hdr/w:p[1]" for g in root.iter(f"{SVG}g"))
    assert ("layout-approximate:text-path" in [w.code for w in opts.warnings]) is not fitted


def test_the_watermark_s_path_finds_its_shape():
    from docx2svg.paths import resolve_path

    header = f'<w:hdr {NS}>{WATERMARK}</w:hdr>'
    layout = convert_docx_to_layout(_docx(_text(1), headers={"rIdh1": header}, final=_sect("rIdh1")), options())
    placed = layout.pages[0].floats[0]
    element = resolve_path(ElementTree.fromstring(header), placed.path)
    assert element is not None and element.get("id") == "WaterMark1"


#: A watermark as Word writes one, made up: no ``w:sdt`` round it, its shape type (``o:spt``
#: 136, ``fitshape="t"``) beside it in its ``w:pict``, no ``v:fill``.
FITTED_WATERMARK = (
    '<w:p><w:r><w:pict><v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800">'
    '<v:path textpathok="t"/><v:textpath on="t" fitshape="t"/></v:shapetype><v:shape id="Made1" '
    'type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:420pt;height:180pt;'
    'rotation:315;z-index:-251650000;mso-position-horizontal:center;mso-position-horizontal-relative:margin;'
    'mso-position-vertical:center;mso-position-vertical-relative:margin" fillcolor="silver" stroked="f">'
    '<v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="SAMPLE"/></v:shape></w:pict>'
    '</w:r></w:p>')


def _text_path(xml: str):
    from docx2svg.parse.vml import read_vml

    pict = ElementTree.fromstring(f'<w:hdr {NS}>{xml}</w:hdr>').find(f".//{{{W}}}pict")
    return read_vml(pict, "w:hdr/w:p[1]/w:r[1]")[1].graphic


@pytest.mark.parametrize("own, kind, beside, typed, fits", [
    ("", "_x0000_t136", True, ' fitshape="t"', True),    # the shape type beside it says fitshape="t"
    ("", "_x0000_t136", False, ' fitshape="t"', True),   # Word's watermark shape type, not at hand
    (' fitshape="f"', "_x0000_t136", True, ' fitshape="t"', True),  # the shape's own: not read (measured)
    (' fitshape="t"', "_x0000_t136", True, "", False),   # nor here: the type beside it does not fit
    ("", "_x0000_t136", True, ' fitshape="f"', False),
    ("", "_x0000_t999", False, ' fitshape="t"', None),   # another type, not at hand: not known
])
def test_a_text_path_fits_its_shape_as_its_shape_type_says(own, kind, beside, typed, fits):
    xml = FITTED_WATERMARK.replace('"#_x0000_t136"', f'"#{kind}"').replace(
        'string="SAMPLE"', f'string="SAMPLE"{own}').replace('<v:textpath on="t" fitshape="t"/>',
                                                           f'<v:textpath on="t"{typed}/>')
    if not beside:
        xml = xml[:xml.index("<v:shapetype")] + xml[xml.index("</v:shapetype>") + len("</v:shapetype>"):]
    graphic = _text_path(xml)
    assert graphic.kind == "textpath" and graphic.body["fit"] is fits


def _matrix(text) -> list[float]:
    found = text.get("transform")
    assert found.startswith("matrix(")
    return [float(v) for v in found[len("matrix("):-1].split()]


def test_a_fitted_text_path_stretches_its_outlines_to_the_box_less_its_insets():
    """Word's watermark (measured, ROADMAP.md): the outlines' box of the string, scaled on
    each axis apart, fills the shape's box less 0.1 in left and right and 0.05 in top and
    bottom, centred -- here from made-up outlines, so the numbers are exact."""
    from docx2svg import drawing

    body = dict(_text_path(FITTED_WATERMARK).body)
    asked = []

    def ink(face, bold, italic, text, oblique):
        asked.append((face, bold, italic, text, oblique))
        return 1000, (50, -20, 3050, 680)   # 3000 by 700 font units, its foot below the baseline

    x, y, width, height = 100, 200, 1750, 750   # px: 420 pt by 180 pt at 300 dpi
    drawn = drawing.draw(drawing.Graphic("textpath", body=body), x, y, width, height, "p", theme={},
                         color_map={}, ink=ink)
    assert asked == [("Calibri", False, False, "SAMPLE", drawing.TEXT_PATH_OBLIQUE)]
    assert drawn.warnings == []                       # fitted: not an approximation
    (primitive,) = drawn.primitives
    assert primitive.face == ("Calibri", False, False)
    text = ElementTree.fromstring(f'<g xmlns="http://www.w3.org/2000/svg">{primitive.markup}</g>')[0]
    sx, b, c, sy, tx, ty = _matrix(text)
    assert (b, c) == (0, 0) and text.get("font-size") == "1000" and "font-kerning:none" in text.get("style")
    inset_x, inset_y = drawing.TEXT_PATH_INSETS
    left, right = tx + 50 * sx, tx + 3050 * sx
    top, bottom = ty - 680 * sy, ty + 20 * sy   # y up in font units, down on the page
    assert left == pytest.approx(x + inset_x, abs=0.01) and right == pytest.approx(x + width - inset_x, abs=0.5)
    assert top == pytest.approx(y + inset_y, abs=0.01) and bottom == pytest.approx(y + height - inset_y, abs=0.5)
    assert (top + bottom) / 2 == pytest.approx(y + height / 2, abs=0.5)
    assert sy / sx == pytest.approx((720 / 700) / (1690 / 3000), rel=0.01)   # stretched taller than wide
    assert text.get("textLength") is None


def test_a_text_path_without_outlines_is_stretched_by_its_advances_and_reported():
    from docx2svg import drawing

    body = dict(_text_path(FITTED_WATERMARK).body)
    for ink, fit in ((lambda *args: None, True), (None, True), (lambda *args: (1000, (0, 0, 10, 10)), None)):
        drawn = drawing.draw(drawing.Graphic("textpath", body={**body, "fit": fit}), 0, 0, 1750, 750, "p",
                             theme={}, color_map={}, ink=ink)
        text = ElementTree.fromstring(f'<g xmlns="http://www.w3.org/2000/svg">{drawn.primitives[0].markup}</g>')[0]
        assert text.get("lengthAdjust") == "spacingAndGlyphs" and text.get("transform") is None
        assert drawn.warnings == ["drawing-approximate:text path"]


def test_a_text_path_its_shape_type_does_not_fit_draws_nothing():
    """Word draws no text for one (measured at 1 to 200 pt, ROADMAP.md): nor does docx2svg."""
    from docx2svg import drawing

    body = dict(_text_path(FITTED_WATERMARK).body)
    drawn = drawing.draw(drawing.Graphic("textpath", body={**body, "fit": False}), 0, 0, 1750, 750, "p",
                         theme={}, color_map={}, ink=lambda *args: (1000, (0, 0, 10, 10)))
    assert drawn.primitives == [] and drawn.warnings == []


def _fitted_text(body: dict, outlines: tuple):
    from docx2svg import drawing

    drawn = drawing.draw(drawing.Graphic("textpath", body=body), 100, 200, 1750, 750, "p", theme={}, color_map={},
                         ink=lambda *args: outlines)
    (primitive,) = drawn.primitives
    return primitive, ElementTree.fromstring(f'<g xmlns="http://www.w3.org/2000/svg">{primitive.markup}</g>')[0]


def test_a_fitted_text_path_sets_its_letters_at_their_kerned_pens():
    """Word kerns a watermark's letters (measured: "AVATAR", "Tyre", "WAVE" and "LT" are
    up to 25 pt off unkerned): the pens the face's kerning gives, one ``x`` per letter."""
    body = dict(_text_path(FITTED_WATERMARK).body, string="AVA")
    primitive, text = _fitted_text(body, (1000, (0, 0, 1700, 700), [0, 560, 1080], 0.0))
    assert text.get("x") == "0 560 1080" and "font-kerning:none" in text.get("style")
    assert text.get("font-style") is None and primitive.face == ("Calibri", False, False)


def test_an_italic_text_path_in_a_face_without_italic_is_slanted_and_fitted_slanted():
    """No italic face: the upright one slanted by TEXT_PATH_OBLIQUE in font units (measured),
    the slanted outlines' box fitted -- and the rasteriser given the upright face."""
    from docx2svg import drawing

    body = dict(_text_path(FITTED_WATERMARK).body, italic=True)
    primitive, text = _fitted_text(body, (1000, (0, -20, 3238, 680), [0, 600], drawing.TEXT_PATH_OBLIQUE))
    sx, b, c, sy, tx, ty = _matrix(text)
    assert b == 0 and c == pytest.approx(-drawing.TEXT_PATH_OBLIQUE * sx, rel=1e-4)
    assert text.get("font-style") is None and primitive.face == ("Calibri", False, False)
    primitive, text = _fitted_text(body, (1000, (0, -20, 3238, 680), [0, 600], 0.0))  # an italic face
    assert _matrix(text)[2] == 0 and text.get("font-style") == "italic" and primitive.face[2] is True


@pytest.mark.parametrize("font, minor, drawn", [
    ("Times New Roman", "Georgia", "Georgia"), (None, "Georgia", "Georgia"), ("times new roman", "Aptos", "Aptos"),
    ("Arial", "Georgia", "Arial"), ("Calibri", "Georgia", "Calibri"), ("Times New Roman", None, "Times New Roman"),
    (None, None, "Calibri"),
])
def test_a_text_path_in_times_new_roman_or_no_face_is_drawn_in_the_minor_face(font, minor, drawn):
    from docx2svg import drawing

    assert drawing.text_path_face(font, minor) == drawn


def test_a_watermark_in_times_new_roman_is_drawn_in_the_theme_s_minor_face():
    xml = FITTED_WATERMARK.replace("&quot;Calibri&quot;", "&quot;Times New Roman&quot;")
    header = f'<w:hdr {NS}>{xml}{_p("Header text")}</w:hdr>'
    layout = convert_docx_to_layout(_docx(_text(1), headers={"rIdh1": header}, final=_sect("rIdh1"),
                                          theme="Georgia"), options())
    assert layout.pages[0].floats[0].primitives[0].face == ("Georgia", False, False)


def test_a_face_s_outlines_box_spans_its_glyphs_at_their_advances():
    found = fonts.InstalledFonts().drawing_face("Calibri", record=False)
    if found is None or found.ink("H") is None:
        pytest.skip("no Calibri or Carlito with glyf outlines here")
    word = found.ink("HxH")
    h, x = found.ink("H"), found.ink("x")
    assert word[0] == h[0] and word[1] == min(h[1], x[1]) and word[3] == max(h[3], x[3])
    assert word[2] == found.advance("H") + found.advance("x") + h[2]
    assert x[3] < h[3]          # an x-height below a cap height
    assert found.ink("H\u0378") is None       # a character the face lacks


def test_a_face_s_kerned_pens_close_up_a_kerned_pair():
    found = fonts.InstalledFonts().drawing_face("Calibri", record=False)
    if found is None or found.ink("AV") is None or not found.pair_kern("A", "V"):
        pytest.skip("no Calibri or Carlito that kerns A V here")
    pens = found.pens("AVA", kerned=True)
    assert pens[1] == found.advance("A") + found.pair_kern("A", "V")
    assert found.pens("AVA") == [0, found.advance("A"), found.advance("A") + found.advance("V")]
    assert found.ink("AVA", kerned=True)[2] < found.ink("AVA")[2]
    slanted = found.ink("AVA", kerned=True, slant=0.34)
    assert slanted[1] == found.ink("AVA", kerned=True)[1] and slanted[2] > found.ink("AVA", kerned=True)[2]


def test_a_word_made_watermark_is_drawn_fitted_and_its_face_given_to_the_rasteriser():
    header = f'<w:hdr {NS}>{FITTED_WATERMARK}{_p("Header text")}</w:hdr>'
    opts = options()
    layout = convert_docx_to_layout(_docx(_text(1), headers={"rIdh1": header}, final=_sect("rIdh1")), opts)
    placed = layout.pages[0].floats[0]
    assert placed.primitives[0].face == ("Calibri", False, False)
    assert ("Calibri", False, False) in docx2svg.drawn_faces(layout)
    if fonts.InstalledFonts().ink("Calibri", False, False, "SAMPLE") is not None:
        assert "matrix(" in placed.primitives[0].markup
        assert "layout-approximate:text-path" not in [w.code for w in opts.warnings]


def test_an_inline_vml_picture_is_drawn_as_a_picture_of_its_size():
    body = ('<w:p><w:r><w:object><v:shape style="width:150pt;height:75pt"><v:imagedata r:id="rIdm1"/></v:shape>'
            '</w:object></w:r></w:p>' + _p("After"))
    layout = convert_docx_to_layout(_docx(body, picture=True), options())
    picture = layout.pages[0].pictures[0]
    assert picture.relationship == "rIdm1" and picture.width == 625  # 150 pt at 300 dpi


# -- crop and notes ----------------------------------------------------------------------------


@pytest.mark.parametrize("anchored", [True, False])
def test_a_cropped_picture_shows_what_is_kept_in_its_box(anchored):
    opts = options()
    document = convert_docx_to_svg(_docx(_picture('l="12000" r="30000"', anchored=anchored), picture=True),
                                   opts)[0]
    assert not [w for w in opts.warnings if w.code == "drawing-not-drawn"]
    root = ElementTree.fromstring(document)
    frame = next(e for e in root.iter(f"{SVG}svg") if e.get("data-docx-path"))
    image = frame.find(f"{SVG}image")
    width = float(frame.get("width"))
    assert frame.get("overflow") == "hidden"
    assert float(image.get("width")) == pytest.approx(width / 0.58, abs=0.01)
    assert float(image.get("x")) == pytest.approx(-0.12 * width / 0.58, abs=0.01)


def test_a_negative_crop_pads_the_picture():
    document = convert_docx_to_svg(_docx(_picture('l="-10000" t="-10000"', anchored=False), picture=True),
                                   options())[0]
    frame = next(e for e in ElementTree.fromstring(document).iter(f"{SVG}svg") if e.get("data-docx-path"))
    image = frame.find(f"{SVG}image")
    assert float(image.get("x")) > 0 and float(image.get("y")) > 0
    assert float(image.get("width")) < float(frame.get("width"))


@pytest.mark.parametrize("columns_first", [False, True])
def test_a_cell_s_footnote_is_drawn_on_a_page_shared_with_columns(columns_first):
    """A section of three columns holding one empty paragraph, a continuous break, then a
    table whose cell references a footnote: the note is drawn at the page's foot, where it
    is without the first section."""
    first = f'<w:p><w:pPr>{_sect(cols=3)}</w:pPr></w:p>' if columns_first else ""
    opts = options()
    layout = convert_docx_to_layout(_docx(first + _tbl([[_cell_note(), _p("x")]]) + _text(2),
                                          final=_sect(kind="continuous"), notes={1: "The cell's note."}), opts)
    lines = [line for line in layout.pages[0].lines if line.path.startswith("w:footnote[@w:id=1]")]
    assert lines and "".join(lines[0].spans[-1].chars).endswith("note.")
    assert "footnotes-not-drawn" not in [w.code for w in opts.warnings]
    if columns_first:
        plain = convert_docx_to_layout(_docx(_tbl([[_cell_note(), _p("x")]]) + _text(2),
                                             final=_sect(kind="continuous"), notes={1: "The cell's note."}),
                                       options())
        assert lines[0].baseline == [line for line in plain.pages[0].lines
                                     if line.path.startswith("w:footnote[@w:id=1]")][0].baseline

