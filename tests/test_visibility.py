"""Every glyph the model draws shows as ink in the raster (``tools/visibility.py``).

``tests/test_render.py`` holds every glyph to Word's position; it cannot see a fill
painted over a glyph, and a banded table's shading once was, on every shaded row of
``sample-with-table`` (ROADMAP.md, "Tables -- measured").  This rasterises each page as
emitted, without the model's text, and the text alone, and fails any glyph whose box
shows no ink over what is under it.

Offline, as the VRT is: the documents are laid out from the recorded face numbers
(``render_record.RecordedFonts``), with no Word and no font file; the rasteriser draws
with whatever faces the host has, and a glyph none of them has is counted as not drawable
here rather than failed.  Needs numpy, pillow and resvg-py (the ``dev`` extra); skipped
without them, and skipped when the host has no face at all.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_anchor_probe  # noqa: E402
import read_render  # noqa: E402
import read_story_probe  # noqa: E402
import read_table_probes  # noqa: E402
import read_wrap_side_probe  # noqa: E402
import render_record  # noqa: E402

#: The committed documents, and the probes that draw every kind of fill text sits on or
#: under: run shading and highlight, paragraph shading and borders (``render-none``),
#: cell shading under and beside text (``table-border-none``).
PROBES = {"render-none": read_render, "table-border-none": read_table_probes,
          # Headers and footers: borders, a table and a picture in them, over every case.
          "story-15": read_story_probe, "fields-computed": read_story_probe,
          # Floating drawings behind the text (under a paragraph's shading), in front of it
          # (over it: counted as covered, as Word paints them over it too), shapes, text
          # boxes and groups.
          "anchor-15": read_anchor_probe, "drawing-15": read_anchor_probe,
          # A header's drawings behind and in front of its text, and the body's over it.
          "story-anchor-15": read_anchor_probe,
          # Text beside drawings text wraps around: none of it under them (147 pages: slow).
          "wrap-side-15": read_wrap_side_probe}


def _source(name: str):
    if name in PROBES:
        module = PROBES[name]
        return dict(module.documents())[name], json.loads(module.OBSERVATIONS.read_text(encoding="utf-8"))["faces"]
    faces = json.loads(read_render.OBSERVATIONS.read_text(encoding="utf-8"))["faces"]
    return read_render.FIXTURES / name, faces


@pytest.fixture(scope="module")
def visibility():
    for module in ("numpy", "PIL", "resvg_py"):
        pytest.importorskip(module)
    import visibility

    probe = ('<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" viewBox="0 0 200 100">\n'
             '<text data-docx-path="probe" x="10" y="70" font-family="sans-serif" font-size="60">Ink</text>\n</svg>\n')
    import numpy

    if not (visibility._raster(probe, 200, 100) < 200).any():
        pytest.skip("the rasteriser has no face to draw text with on this host")
    return visibility


#: Probes whose check takes minutes: run with ``--run-slow``.
SLOW = {"wrap-side-15"}


@pytest.mark.parametrize("name", [pytest.param(name, marks=pytest.mark.slow) if name in SLOW else name
                                  for name in list(read_render.COMMITTED) + list(PROBES)])
def test_every_glyph_drawn_is_visible(name, visibility):
    source, faces = _source(name)
    result = visibility.check(name, source, render_record.options(render_record.RecordedFonts(faces)))
    if result.hidden:
        pytest.fail(f"{len(result.hidden)} of {result.glyphs} glyphs draw no ink:\n" + visibility.describe(result.hidden))
    # A host that lacks a face falls back to another; one that draws almost nothing checks nothing.
    assert result.undrawable <= result.glyphs // 100, result.row()


def test_a_glyph_under_a_fill_is_hidden(visibility):
    """The instrument itself: a rectangle painted after a word hides it, and one painted
    before it does not."""
    from docx2svg.layout import Layout, Line, Page, Span

    span = Span("w:p[1]/w:r[1]", list("Hidden"), [100 + 30 * k for k in range(6)], 200, 280, "sans-serif",
                False, False, 22)
    page = Page(0, 600, 300, 144, 72, (0, 600, 0, 300), lines=[Line("w:p[1]", 0, 200, 150, 60, [span])])
    text = ('<text data-docx-path="w:p[1]/w:r[1]" x="100 130 160 190 220 250" y="200" font-family="sans-serif"'
            ' font-size="46" fill="#000000">Hidden</text>')
    rect = '<rect x="90" y="150" width="220" height="70" fill="#D3DFEE"/>'
    header = '<svg xmlns="http://www.w3.org/2000/svg" width="144pt" height="72pt" viewBox="0 0 600 300">'
    over = "\n".join([header, "<g>", text, "</g>", rect, "</svg>"]) + "\n"
    under = "\n".join([header, rect, "<g>", text, "</g>", "</svg>"]) + "\n"
    layout = Layout([page])
    assert len(visibility.check_pages("over", [over], layout).hidden) == 6
    assert visibility.check_pages("under", [under], layout).visible == 6
