"""Open substitutes for absent faces, recorded symbol faces, and the layout's coverage.

Where Word's face is neither installed nor embedded, :class:`docx2svg.fonts.InstalledFonts`
lays it out with its metric compatible open substitute (:data:`docx2svg.fonts.SUBSTITUTES`)
and an absent Symbol or Wingdings from the integers recorded from Word's copies
(:mod:`docx2svg.recorded`) -- explicitly: every use is a ``font-substituted`` warning and
is listed in the layout's :class:`docx2svg.coverage.Coverage`.

The tests that lay out with the substitutes need them installed (CI installs Carlito,
Caladea and Liberation on Linux and macOS; ``DOCX2SVG_SUBSTITUTE_FONTS`` names more
directories), and skip where they are not.  They see *only* the substitutes' folders, so
they lay out as a machine without Office does, here too.  The coverage tests need no
font at all.
"""

from __future__ import annotations

import io
import json
import os
import sys
import zipfile
from pathlib import Path

import pytest

FIXTURES = Path(__file__).resolve().parent / "fixtures"
sys.path.insert(0, str(FIXTURES.parent.parent / "tools"))

import render_record  # noqa: E402

import docx2svg  # noqa: E402
from docx2svg import fonts  # noqa: E402
from docx2svg.vertical import FaceMetrics  # noqa: E402

DATA = json.loads((FIXTURES / "render-observations.json").read_text(encoding="utf-8"))


def _substitute_dirs() -> tuple[Path, ...]:
    """The folders holding Carlito and the Liberation faces here (and only those)."""
    extra = tuple(Path(p) for p in os.environ.get("DOCX2SVG_SUBSTITUTE_FONTS", "").split(os.pathsep) if p)
    index = fonts.installed_index(tuple(fonts.default_font_dirs()) + extra)
    found = []
    for family in sorted(set(fonts.SUBSTITUTES.values())):
        entry = index.get((family.lower(), False, False))
        if entry is None:
            return ()
        found.append(Path(entry[0]).parent)
    return tuple(dict.fromkeys(found))


@pytest.fixture
def substitutes_only(monkeypatch) -> tuple[Path, ...]:
    """Lay out seeing only the substitutes' folders: Word's faces are absent."""
    dirs = _substitute_dirs()
    if not dirs:
        pytest.skip("Carlito and Liberation are not installed here (fonts-crosextra-carlito, fonts-liberation)")
    monkeypatch.setattr(fonts, "default_font_dirs", lambda: dirs)
    return dirs


def _glyphs(layout) -> list:
    return [(page.number, line.path, line.baseline,
             tuple((tuple(span.chars), tuple(span.xs), span.y, span.half_points) for span in line.spans))
            for page in layout.pages for line in page.text_lines()]


# -- what is substituted, and when ----------------------------------------------------------------


def test_the_table_holds_only_the_measured_clones():
    assert fonts.SUBSTITUTES == {"calibri": "Carlito", "arial": "Liberation Sans",
                                 "times new roman": "Liberation Serif", "courier new": "Liberation Mono"}


def test_a_face_that_is_present_is_never_substituted(substitutes_only):
    found = fonts.InstalledFonts(dirs=substitutes_only, substitutes={"liberation sans": "Carlito"})
    face = found.face("Liberation Sans")
    assert face is not None and "Carlito" not in face.families()
    assert found.substitutions == {} and found.missing == set()


def test_carlito_lays_calibri_out_as_word_does(substitutes_only):
    """``layout-sweep.docx`` is all Calibri: with Carlito in its place every glyph's pen
    position and baseline is the one the layout gives from the numbers of Word's own
    Calibri (recorded: ``tests/fixtures/render-observations.json``), which
    ``tests/test_render.py`` holds to Word's PDF."""
    data = (FIXTURES / "layout-sweep.docx").read_bytes()
    word = docx2svg.convert_docx_to_layout(data, render_record.options(render_record.RecordedFonts(DATA["faces"])))
    options = docx2svg.ConvertOptions()
    layout = docx2svg.convert_docx_to_layout(data, options)
    assert _glyphs(layout) == _glyphs(word)
    assert len(layout.pages) == 3
    coverage = layout.coverage
    assert coverage.complete and coverage.substituted_fonts == [
        {"family": "Calibri", "substitute": "Carlito", "metric_compatible": True}]
    assert [w.code for w in options.warnings].count("font-substituted") == 1
    assert "Carlito" in next(w.message for w in options.warnings if w.code == "font-substituted")


def test_the_svg_names_the_substitute_after_the_document_s_face(substitutes_only):
    svgs = docx2svg.convert_docx_to_svg((FIXTURES / "layout-sweep.docx").read_bytes(), docx2svg.ConvertOptions(pages=[1]))
    assert 'font-family="Calibri, Carlito"' in svgs[0]


def test_without_substitution_an_absent_face_stops_the_layout(substitutes_only):
    options = docx2svg.ConvertOptions(substitute_fonts=False)
    layout = docx2svg.convert_docx_to_layout((FIXTURES / "layout-sweep.docx").read_bytes(), options)
    assert not layout.coverage.complete
    assert layout.coverage.stop.code == "layout-stopped:unmeasurable"
    assert layout.coverage.missing_fonts == ["Calibri"] and layout.coverage.substituted_fonts == []
    assert "font-substituted" not in [w.code for w in options.warnings]


def test_a_substitute_the_caller_names_is_reported_approximate(substitutes_only):
    """Calibri Light has no clone; Carlito is 1.2 to 3.4% wider.  Named by the caller it is
    used, and said to be approximate."""
    options = docx2svg.ConvertOptions(font_substitutes={"Calibri Light": "Carlito"})
    layout = docx2svg.convert_docx_to_layout((FIXTURES / "style-document.docx").read_bytes(), options)
    approximate = [s for s in layout.coverage.substituted_fonts if not s["metric_compatible"]]
    assert approximate == [{"family": "Calibri Light", "substitute": "Carlito", "metric_compatible": False}]
    assert any("approximate" in w.message for w in options.warnings if w.code == "font-substituted")


# -- symbol faces ------------------------------------------------------------------------------


def test_the_recorded_symbol_faces_are_the_installed_ones():
    """Where Word's Symbol and Wingdings are installed, :mod:`docx2svg.recorded` answers as
    they do (``tools/record_symbol_faces.py`` wrote it from them)."""
    from docx2svg import recorded

    if sys.platform != "darwin" or not fonts.WORD_FONTS.is_dir():
        pytest.skip("Word's faces are not installed here")
    installed = fonts.InstalledFonts(substitutes={}, recorded_faces=False)
    for family in ("Symbol", "Wingdings"):
        face = installed.face(family)
        assert face is not None, family
        stand_in = fonts.RecordedFace(family)
        assert stand_in.metrics == face.metrics
        assert stand_in.decorations == face.decorations
        assert {code: face.advance(chr(code)) for code in recorded.ADVANCES[family.lower()]} == \
            recorded.ADVANCES[family.lower()]


def test_an_absent_symbol_face_is_laid_out_from_its_recorded_metrics():
    found = fonts.InstalledFonts(dirs=(), substitutes={})
    face = found.face("Symbol")
    assert isinstance(face, fonts.RecordedFace)
    assert face.metrics == FaceMetrics(2048, 2059, 450, 0, 1331, 1331, 928, 293)  # SymbolMT: 1.2251 em
    assert found.advance("Symbol", False, False, "") == (face.advance(""), 2048)
    assert found.substitutions["Symbol"].substitute == fonts.RECORDED
    name = found.drawing_name("Symbol")
    assert name.family == "sans-serif" and name.chars[""] == "•"
    assert fonts.InstalledFonts(dirs=(), substitutes={}, recorded_faces=False).face("Symbol") is None


def test_an_absent_symbol_face_draws_every_glyph_from_the_shared_tables():
    """Not only the bullet library: every code of the face whose glyph Unicode encodes
    (ooxml-common's ``symbol_fonts``, shared with pptx2svg)."""
    from docx2svg import recorded
    from ooxml_common.text.symbol_fonts import to_unicode

    wingdings, symbol = fonts.RecordedFace("Wingdings"), fonts.RecordedFace("Symbol")
    assert wingdings.unicode["\uf071"] == "\u2751"  # LOWER RIGHT SHADOWED WHITE SQUARE
    assert wingdings.unicode["\uf0fc"] == "\u2713"  # CHECK MARK, as Word and PowerPoint draw it
    assert wingdings.unicode["\uf038"] == "\U0001f5b0"  # TWO BUTTON MOUSE
    assert symbol.unicode["\uf061"] == symbol.unicode["a"] == "\u03b1"  # GREEK SMALL LETTER ALPHA
    for family, advances in recorded.ADVANCES.items():
        mapped = {code for code in advances if to_unicode(family, chr(code))}
        assert len(mapped) > 0.75 * len(advances), family
        assert {ord(char) for char in fonts.RecordedFace(family).unicode} == mapped


def test_symbol_bullets_lay_out_and_draw_without_the_faces(substitutes_only):
    data = _document(_bullets(), numbering=_numbering())
    options = docx2svg.ConvertOptions()
    conversion = docx2svg.convert_docx(data, options)
    coverage = conversion.layout.coverage
    assert coverage.complete, coverage.summary()
    assert {s["family"]: s["substitute"] for s in coverage.substituted_fonts} == {
        "Calibri": "Carlito", "Symbol": fonts.RECORDED, "Wingdings": fonts.RECORDED}
    # The text after each bullet starts at the hanging indent: 1440 + 720 twips, in px.
    starts = [line.spans[1].xs[0] for line in conversion.layout.pages[0].text_lines()[:2]]
    assert starts == [450, 450]
    # Drawn as the symbols' Unicode equivalents, in a generic face after the document's.
    assert "font-family=\"Symbol, sans-serif\"" in conversion.svgs[0] and "•</text>" in conversion.svgs[0]
    assert "➢</text>" in conversion.svgs[0]


# -- coverage ----------------------------------------------------------------------------------


class Uniform:
    """Every face the same: half an em per character, Calibri's line.  Enough to lay a
    probe out anywhere, without a font."""

    def advance(self, face, bold, italic, char):
        return 1024, 2048

    def kern(self, face, bold, italic, left, right):
        return 0

    def metrics(self, face, bold=False, italic=False):
        return FaceMetrics(2048, 1950, 550, 0, 1331, 1331, 976, 286)

    def decorations(self, face, bold=False, italic=False):
        return fonts.Decorations(2048, 976, 286, 512, 134, -232, 134, 1950, 550)

    def drawing_name(self, face, bold=False, italic=False):
        return None


def _uniform(**kwargs) -> docx2svg.ConvertOptions:
    return render_record.options(Uniform(), **kwargs)


def test_a_complete_layout_says_so():
    layout = docx2svg.convert_docx_to_layout(_document(_paragraphs(3)), _uniform())
    coverage = layout.coverage
    assert coverage.complete and coverage.stop is None and coverage.story_stops == []
    assert (coverage.pages, coverage.estimated_pages, coverage.estimate_source) == (1, 1, "layout")
    assert (coverage.blocks, coverage.blocks_laid_out, coverage.blocks_skipped) == (3, 3, 0)
    assert coverage.summary() == "complete: 1 page(s), 3 block(s)"


def test_a_body_stop_is_counted_and_located():
    body = _paragraphs(2) + f"<w:p>{_VML}<w:r><w:t>a shape</w:t></w:r></w:p>" + _paragraphs(3)
    options = _uniform()
    layout = docx2svg.convert_docx_to_layout(_document(body, pages=5), options)
    coverage = layout.coverage
    assert not coverage.complete
    assert coverage.stop.code == "layout-stopped:drawing" and coverage.stop.page == 1
    assert coverage.stop.path == "w:body/w:p[3]"
    assert (coverage.blocks, coverage.blocks_laid_out, coverage.blocks_skipped) == (6, 2, 4)
    # Past the stop the pages are not known: Word's count as last saved is the estimate.
    assert (coverage.estimated_pages, coverage.estimate_source) == (5, "app.xml")
    assert options.coverage is coverage
    assert coverage.as_dict()["stop"]["reason"] == "drawing"
    assert coverage.summary().startswith("partial: 1 of ~5 page(s), 2 of 6 block(s) laid out; stopped on page 1")


def test_a_header_stop_leaves_the_body_complete_but_not_the_coverage():
    layout = docx2svg.convert_docx_to_layout(_document(_paragraphs(2), header=_VML), _uniform())
    coverage = layout.coverage
    assert coverage.stop is None and coverage.blocks_laid_out == 2
    assert [s.code for s in coverage.story_stops] == ["story-stopped:drawing"]
    assert not coverage.complete


def test_without_an_estimate_none_is_invented():
    body = f"<w:p>{_VML}</w:p>" + _paragraphs(2)
    coverage = docx2svg.convert_docx_to_layout(_document(body), _uniform()).coverage
    assert (coverage.estimated_pages, coverage.estimate_source, coverage.blocks_laid_out) == (None, None, 0)


# -- probes ------------------------------------------------------------------------------------

_W = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:v="urn:schemas-microsoft-com:vml"')
_VML = ('<w:r><w:pict><v:rect style="position:absolute;margin-left:0;margin-top:0;width:100pt;height:20pt;'
        'z-index:-1" fillcolor="#4472c4" stroked="f"/></w:pict></w:r>')
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"


def _paragraphs(count: int) -> str:
    return "".join(f"<w:p><w:r><w:t>Paragraph {k + 1}.</w:t></w:r></w:p>" for k in range(count))


def _bullets() -> str:
    return "".join(
        f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{k + 1}"/></w:numPr><w:ind w:left="720" '
        f'w:hanging="360"/></w:pPr><w:r><w:t>A bullet in {face}.</w:t></w:r></w:p>'
        for k, face in enumerate(("Symbol", "Wingdings")))


def _numbering() -> str:
    levels = [("Symbol", ""), ("Wingdings", "")]
    abstract = "".join(
        f'<w:abstractNum w:abstractNumId="{k}"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        f'<w:lvlText w:val="{char}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
        f'<w:rPr><w:rFonts w:ascii="{face}" w:hAnsi="{face}" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>'
        for k, (face, char) in enumerate(levels))
    nums = "".join(f'<w:num w:numId="{k + 1}"><w:abstractNumId w:val="{k}"/></w:num>' for k in range(len(levels)))
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering {_W}>{abstract}{nums}</w:numbering>'


def _document(body: str, *, header: str | None = None, numbering: str | None = None, pages: int | None = None) -> bytes:
    """A minimal package: Calibri 11 pt, A4 with 1-inch margins."""
    overrides = [("/word/document.xml", f"{_CT}.document.main+xml"), ("/word/styles.xml", f"{_CT}.styles+xml")]
    rels = [("styles", "styles.xml")]
    reference = ""
    parts = {}
    if header is not None:
        overrides.append(("/word/header1.xml", f"{_CT}.header+xml"))
        rels.append(("header", "header1.xml"))
        reference = f'<w:headerReference w:type="default" r:id="rId{len(rels)}"/>'
        parts["word/header1.xml"] = f'<w:hdr {_W}><w:p>{header}<w:r><w:t>Header</w:t></w:r></w:p></w:hdr>'
    if numbering is not None:
        overrides.append(("/word/numbering.xml", f"{_CT}.numbering+xml"))
        rels.append(("numbering", "numbering.xml"))
        parts["word/numbering.xml"] = numbering
    package_rels = [(f"{_REL}/officeDocument", "word/document.xml")]
    if pages is not None:
        overrides.append(("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml"))
        package_rels.append((f"{_REL}/extended-properties", "docProps/app.xml"))
        parts["docProps/app.xml"] = ('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/'
                                     f'extended-properties"><Pages>{pages}</Pages></Properties>')
    styles = (f'<w:styles {_W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" '
              'w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing '
              'w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
              '</w:styles>')
    document = (f'<w:document {_W}><w:body>{body}<w:sectPr>{reference}<w:pgSz w:w="11906" w:h="16838"/>'
                '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" '
                'w:gutter="0"/></w:sectPr></w:body></w:document>')
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("[Content_Types].xml", (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" '
            'ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" '
            'ContentType="application/xml"/>'
            + "".join(f'<Override PartName="{name}" ContentType="{kind}"/>' for name, kind in overrides) + "</Types>"))
        archive.writestr("_rels/.rels", (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(f'<Relationship Id="rId{k + 1}" Type="{kind}" Target="{target}"/>'
                      for k, (kind, target) in enumerate(package_rels)) + "</Relationships>"))
        archive.writestr("word/_rels/document.xml.rels", (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(f'<Relationship Id="rId{k + 1}" Type="{_REL}/{kind}" Target="{target}"/>'
                      for k, (kind, target) in enumerate(rels)) + "</Relationships>"))
        archive.writestr("word/document.xml", document)
        archive.writestr("word/styles.xml", styles)
        for name, text in parts.items():
            archive.writestr(name, text)
    return buffer.getvalue()
