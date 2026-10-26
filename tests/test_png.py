"""``docx2svg.png`` on ooxml-common's rasterisers (:mod:`ooxml_common.raster`, review S7):
the names and errors it had, with the mechanism shared with pptx2svg."""

from __future__ import annotations

import sys

import pytest

import docx2svg
from docx2svg import png
from ooxml_common import raster

SVG = '<svg xmlns="http://www.w3.org/2000/svg" width="4" height="2"><rect width="4" height="2"/></svg>'


def test_the_public_names_are_the_shared_ones():
    assert docx2svg.RasterizerNotAvailable is png.RasterizerNotAvailable is raster.RasterizerNotAvailable
    assert docx2svg.available_backends is raster.available_backends


def test_no_backend_says_how_to_install_one(monkeypatch):
    monkeypatch.setitem(sys.modules, "resvg_py", None)
    monkeypatch.setitem(sys.modules, "cairosvg", None)
    with pytest.raises(docx2svg.RasterizerNotAvailable, match=r"pip install docx2svg\[png\]"):
        png.svg_to_png(SVG)
    with pytest.raises(docx2svg.RasterizerNotAvailable, match=r"pip install docx2svg\[png\]"):
        png.svg_to_png(SVG, backend="resvg")


def test_an_unknown_backend_is_a_value_error():
    with pytest.raises(ValueError, match="unknown rasterizer backend"):
        png.svg_to_png(SVG, backend="pdf")  # type: ignore[arg-type]
