"""``ConvertOptions(text="outlines")``: a page's text as the glyph paths its PNG is drawn
with (:func:`docx2svg.png.outline_svg`, :mod:`ooxml_common.outlines`) -- the files
:func:`docx2svg.convert_docx_to_png` hands the rasteriser, in the passes it draws them in.

The faces are the made-up pair of ``tests/test_drawing_passes.py`` (two open faces filed
under one name, so the page is drawn in two passes); rasterised, the outlined page must
draw the pixels the text draws.
"""

from __future__ import annotations

import re

import pytest

import docx2svg
from docx2svg import ConvertOptions
from test_drawing_passes import BODY, DISPLAY, _document, _needs_resvg, _rows, folder  # noqa: F401


def _page(document: bytes, options: ConvertOptions, device_size: bool = False) -> str:
    layout, data, found = docx2svg.lay_out_docx(document, options)
    return docx2svg.page_svgs(layout, data, options, device_size=device_size, fonts=found)[0]


def test_the_outlined_page_draws_the_pixels_its_text_draws(folder):  # noqa: F811
    _needs_resvg()
    document = _document(DISPLAY, BODY)
    text = docx2svg.convert_docx_to_png(document, ConvertOptions(width=600, font_dirs=[str(folder)]))[0]
    options = ConvertOptions(width=600, font_dirs=[str(folder)], text="outlines")
    outlined = _page(document, options, device_size=True)
    assert "<text" not in outlined and "<use " in outlined
    assert not [w for w in options.warnings if w.code == "text-not-outlined"]
    # Nothing left to draw with a font: the folder is not needed to draw it.
    png = docx2svg.svg_to_png(outlined, width=600, font_dirs=[], skip_system_fonts=True)
    rows = (60, 200)
    worst = max(abs(a - b) for left, right in zip(_rows(text, *rows), _rows(png, *rows))
                for a, b in zip(left, right))
    assert worst <= 64             # an anti-aliased edge pixel at most
    # The display face's line is the display face's, not the body face's that ties with it.
    body = docx2svg.convert_docx_to_png(_document(BODY, BODY), ConvertOptions(width=600, font_dirs=[str(folder)]))[0]
    assert _rows(png, 60, 112) != _rows(body, 60, 112)


def test_the_default_is_text_and_unchanged(folder):  # noqa: F811
    document = _document(BODY)
    options = ConvertOptions(font_dirs=[str(folder)])
    assert _page(document, options) == _page(document, ConvertOptions(font_dirs=[str(folder)], text="text"))
    assert "<text" in _page(document, options)


def test_an_unknown_text_mode_is_refused(folder):  # noqa: F811
    with pytest.raises(ValueError, match="text must be one of text, outlines"):
        _page(_document(BODY), ConvertOptions(font_dirs=[str(folder)], text="glyphs"))


def test_selectable_text_is_kept_transparent_over_the_outlines(folder):  # noqa: F811
    svg = _page(_document(BODY), ConvertOptions(font_dirs=[str(folder)], text="outlines", selectable_text=True))
    layer = re.search(r"<text\b[^>]*>Handgloves</text>", svg)
    assert layer and 'fill-opacity="0"' in layer.group(0) and "data-text-layer" in layer.group(0)
    assert 'aria-hidden="true"' in svg


def test_png_output_is_the_same_whatever_the_text_mode(folder):  # noqa: F811
    _needs_resvg()
    document = _document(DISPLAY, BODY)
    assert docx2svg.convert_docx_to_png(document, ConvertOptions(width=300, font_dirs=[str(folder)], text="outlines")) \
        == docx2svg.convert_docx_to_png(document, ConvertOptions(width=300, font_dirs=[str(folder)]))


# -- Round 32: stable glyph ids, one glyph set per document ----------------------------

def _two_pages(first: str, second: str) -> bytes:
    """Two pages of the body face: ``first``, a page break, ``second``."""
    import io
    import zipfile

    source = zipfile.ZipFile(io.BytesIO(_document(BODY)))
    run = '<w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:t xml:space="preserve">{}</w:t></w:r>'
    body = (f'<w:p>{run.format(first)}</w:p><w:p><w:r><w:br w:type="page"/></w:r>{run.format(second)}</w:p>')
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name in source.namelist():
            data = source.read(name)
            if name == "word/document.xml":
                text = data.decode("utf-8")
                start, end = text.index("<w:body>") + len("<w:body>"), text.index("<w:sectPr>")
                data = (text[:start] + body + text[end:]).encode("utf-8")
            archive.writestr(name, data)
    return buffer.getvalue()


def _pages(document: bytes, options: ConvertOptions) -> list[str]:
    layout, data, found = docx2svg.lay_out_docx(document, options)
    return docx2svg.page_svgs(layout, data, options, fonts=found)


def test_an_edit_on_one_page_leaves_the_others_byte_identical(folder, tmp_path):  # noqa: F811
    before = _pages(_two_pages("Quick brown fox", "Lazy dog, 2026"), ConvertOptions(font_dirs=[str(folder)],
                                                                                    text="outlines"))
    # The first page now draws the second's glyphs first, in another order.
    after = _pages(_two_pages("9 1 0 2 , god yzaL", "Lazy dog, 2026"), ConvertOptions(font_dirs=[str(folder)],
                                                                                      text="outlines"))
    assert len(before) == len(after) == 2
    assert before[0] != after[0] and before[1] == after[1]
    # A glyph's id is its outline's, not its file's: the same faces read from another
    # folder (as an embedded face is, from a new temporary folder each time) draw the
    # same bytes.
    elsewhere = tmp_path / "elsewhere"
    elsewhere.mkdir()
    for face in folder.iterdir():
        (elsewhere / face.name).write_bytes(face.read_bytes())
    moved = _pages(_two_pages("Quick brown fox", "Lazy dog, 2026"), ConvertOptions(font_dirs=[str(elsewhere)],
                                                                                   text="outlines"))
    assert moved == before


def test_one_glyph_set_for_the_document(folder):  # noqa: F811
    document = _two_pages("Quick brown fox", "Lazy dog, 2026")
    alone = _pages(document, ConvertOptions(font_dirs=[str(folder)], text="outlines"))
    glyphs: dict = {}
    shared = _pages(document, ConvertOptions(font_dirs=[str(folder)], text="outlines", glyphs=glyphs))
    assert glyphs and not any('<path id="g' in page for page in shared)
    assert [docx2svg.embed_glyphs(page, glyphs) for page in shared] == alone
    sheet = docx2svg.glyph_sheet(glyphs)
    ids = re.findall(r'\bid="(g[0-9A-Za-z]+)"', sheet + "".join(shared))
    assert sorted(ids) == sorted(glyphs) and len(ids) == len(set(ids))
