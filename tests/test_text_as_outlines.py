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
