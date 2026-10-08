"""The public API and the command line: shapes, warnings, identity and honesty.

Offline: every render here measures with the faces' recorded numbers
(``tools/render_record.RecordedFonts``), so it needs no font file.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from xml.etree import ElementTree

import pytest

REPO = Path(__file__).resolve().parent.parent
FIXTURES = REPO / "tests" / "fixtures"
sys.path.insert(0, str(REPO / "tools"))

import render_record  # noqa: E402

from docx2svg import ConvertOptions, convert_docx_to_layout, convert_docx_to_svg  # noqa: E402
from docx2svg.cli import main, parse_page_selection  # noqa: E402

SVG = "{http://www.w3.org/2000/svg}"
FACES = json.loads((FIXTURES / "render-observations.json").read_text(encoding="utf-8"))["faces"]


def options(**kwargs) -> ConvertOptions:
    return render_record.options(render_record.RecordedFonts(FACES), **kwargs)


def test_one_svg_per_page_in_device_pixels():
    documents = convert_docx_to_svg(FIXTURES / "samplelib" / "sample-long.docx", options())
    layout = convert_docx_to_layout(FIXTURES / "samplelib" / "sample-long.docx", options())
    assert len(documents) == len(layout.pages) > 1
    root = ElementTree.fromstring(documents[0])
    assert root.get("viewBox") == "0 0 2550 3300"  # US Letter at 300 dpi: Word's export grid
    assert root.get("width") == "612pt" and root.get("height") == "792pt"


def test_every_element_drawn_from_the_document_names_where_it_came_from():
    document = convert_docx_to_svg(FIXTURES / "samplelib" / "sample-simple.docx", options())[0]
    root = ElementTree.fromstring(document)
    texts = list(root.iter(f"{SVG}text"))
    assert texts
    for text in texts:
        if text.get("font-family") != "sans-serif":
            assert text.get("data-docx-path", "").startswith(("w:body/w:p[", "w:body/w:tbl["))
    assert any(t.get("data-docx-path") == "w:body/w:p[1]/w:r[1]" for t in texts)
    # One x per character: the rasteriser cannot move a glyph (the stop band's label is
    # ours, not the document's, and is set as one string).
    for text in texts:
        if text.get("font-family") != "sans-serif":
            assert len(text.get("x").split()) == len(text.text or "")


def _floating_table(tmp_path) -> Path:
    """``sample-simple`` with its table made floating (``w:tblpPr``) and aligned
    ``inside``, which the layout does not model (``make_float_table_probe.py`` did not
    measure it): a table it stops at."""
    import zipfile

    source = FIXTURES / "samplelib" / "sample-simple.docx"
    target = tmp_path / "sample-simple-floating.docx"
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for item in original.infolist():
            data = original.read(item.filename)
            if item.filename == "word/document.xml":
                data = data.replace(b"<w:tblPr>", b'<w:tblPr><w:tblpPr w:vertAnchor="text" w:tblpXSpec="inside" w:tblpY="1"/>', 1)
            out.writestr(item, data)
    return target


def test_a_table_stops_the_layout_and_says_so(tmp_path):
    opts = options()
    documents = convert_docx_to_svg(_floating_table(tmp_path), opts)
    codes = [w.code for w in opts.warnings]
    assert "layout-stopped:table" in codes
    stop = [w for w in opts.warnings if w.code == "layout-stopped:table"][0]
    assert stop.page == 1 and "w:body/w:tbl[1]" in stop.message and "floating" in stop.message
    root = ElementTree.fromstring(documents[-1])
    band = [g for g in root.iter(f"{SVG}g") if g.get("data-docx-unsupported") == "table"]
    assert band and band[0].get("data-docx-path") == "w:body/w:tbl[1]"
    # Nothing past the table is drawn: Word's page holds text below it; ours does not.
    assert len(documents) == 1


def test_a_table_is_laid_out_and_the_text_after_it_goes_on():
    opts = options()
    documents = convert_docx_to_svg(FIXTURES / "samplelib" / "sample-simple.docx", opts)
    assert not [w for w in opts.warnings if w.code.startswith("layout-stopped")]
    paths = [t.get("data-docx-path", "") for d in documents for t in ElementTree.fromstring(d).iter(f"{SVG}text")]
    assert any(p.startswith("w:body/w:tbl[1]/w:tr[1]/w:tc[1]/w:p[1]") for p in paths)
    assert len(documents) == 2


def test_a_document_that_draws_everything_warns_of_nothing():
    opts = options()
    convert_docx_to_svg(FIXTURES / "style-document.docx", opts)
    assert opts.warnings == []


def test_pages_can_be_selected():
    assert parse_page_selection("1,3-4") == [1, 3, 4]
    documents = convert_docx_to_svg(FIXTURES / "samplelib" / "sample-long.docx", options(pages=[2]))
    assert len(documents) == 1 and 'data-docx-page="2"' in documents[0]


def test_glyph_size_scales_outlines_and_not_positions():
    device = convert_docx_to_svg(FIXTURES / "style-document.docx", options())[0]
    exact = convert_docx_to_svg(FIXTURES / "style-document.docx", options(glyph_size="exact"))[0]
    strip = lambda s: [line for line in s.splitlines() if "font-size" not in line]  # noqa: E731
    assert strip(device) == strip(exact)
    assert 'font-size="46"' in device and 'font-size="45.8333"' in exact


def test_the_command_line_writes_one_file_per_page(tmp_path, monkeypatch, capsys):
    import docx2svg

    original = docx2svg._fonts
    monkeypatch.setattr(docx2svg, "_fonts", lambda data, opts: render_record.RecordedFonts(FACES))
    try:
        status = main([str(_floating_table(tmp_path)), "-o", str(tmp_path / "out"), "--strict"])
    finally:
        monkeypatch.setattr(docx2svg, "_fonts", original)
    assert status == 2  # --strict: the table is a warning
    assert sorted(p.name for p in (tmp_path / "out").iterdir()) == ["sample-simple-floating-1.svg"]
    assert "layout-stopped:table" in capsys.readouterr().err


def test_png_needs_its_extra_and_draws_at_300_dpi():
    pytest.importorskip("resvg_py")
    from docx2svg import convert_docx_to_png

    png = convert_docx_to_png(FIXTURES / "samplelib" / "sample-blank.docx", options())[0]
    assert png[:8] == b"\x89PNG\r\n\x1a\n"
    import struct

    width, height = struct.unpack(">II", png[16:24])
    assert (width, height) == (2550, 3300)


def test_one_call_lays_out_once_and_returns_the_layout_and_the_svgs(monkeypatch):
    """:func:`docx2svg.convert_docx`: the layout :func:`convert_docx_to_layout` returns and
    the SVGs :func:`convert_docx_to_svg` writes, byte for byte, from one layout; with a
    page selection, the selected pages' SVGs and every page's layout."""
    import docx2svg
    import docx2svg.layout

    source = FIXTURES / "samplelib" / "sample-long.docx"
    calls = []
    lay_out = docx2svg.layout.lay_out
    monkeypatch.setattr(docx2svg.layout, "lay_out", lambda *a, **k: calls.append(1) or lay_out(*a, **k))
    both_options = options()
    both = docx2svg.convert_docx(source, both_options)
    assert len(calls) == 1
    svg_options = options()
    assert both.svgs == convert_docx_to_svg(source, svg_options)
    assert [str(w) for w in both_options.warnings] == [str(w) for w in svg_options.warnings]
    alone = convert_docx_to_layout(source, options())
    assert [[(line.path, line.top, line.baseline) for line in page.lines] for page in both.layout.pages] == \
        [[(line.path, line.top, line.baseline) for line in page.lines] for page in alone.pages]
    assert both.page_numbers == list(range(1, len(alone.pages) + 1))

    selected = docx2svg.convert_docx(source, options(pages=[2, 99]))
    assert selected.page_numbers == [2]
    assert selected.svgs == [both.svgs[1]]
    assert len(selected.layout.pages) == len(alone.pages)


def test_every_page_carries_its_public_page_info():
    """``Page.info`` is a :class:`docx2svg.PageInfo`: the number Word prints, the section,
    whether it starts there, the story kind and whether the page is blank."""
    import read_story_probe

    from docx2svg import PageInfo

    data = dict(read_story_probe.documents())["select-even"]
    layout = convert_docx_to_layout(data)
    infos = [page.info for page in layout.pages]
    assert all(isinstance(info, PageInfo) for info in infos)
    assert [info.blank for info in infos].count(True) >= 1
    assert all(info.story in ("first", "even", "default") for info in infos)
    assert infos[0].section == 0 and infos[0].section_first
    assert len({info.section for info in infos}) > 1
