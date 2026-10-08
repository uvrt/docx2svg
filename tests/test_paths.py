"""Element paths: every drawn object's ``path`` names the element it was drawn from, and
:func:`docx2svg.paths.resolve_path` finds that element again (``src/docx2svg/paths.py``).

Offline: every document is laid out with recorded face numbers.
"""

from __future__ import annotations

import io
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree

import pytest

REPO = Path(__file__).resolve().parent.parent
FIXTURES = REPO / "tests" / "fixtures"
sys.path.insert(0, str(REPO / "tools"))

import render_record  # noqa: E402

from docx2svg import convert_docx, convert_docx_to_layout, resolve_path  # noqa: E402
from docx2svg.opc import Package  # noqa: E402
from docx2svg.paths import locate, path_part, split_path  # noqa: E402
from docx2svg.xmlutil import local_name, parse_xml  # noqa: E402

W14_PARA_ID = "{http://schemas.microsoft.com/office/word/2010/wordml}paraId"
COMMITTED = (
    "layout-sweep.docx", "style-document.docx",
    "samplelib/sample-blank.docx", "samplelib/sample-long.docx", "samplelib/sample-resume.docx",
    "samplelib/sample-simple.docx",
    "wordto/sample-1page.docx", "wordto/sample-5pages.docx", "wordto/sample-10pages.docx",
    "wordto/sample-with-images.docx", "wordto/sample-with-table.docx",
)
#: Generated probes whose paths reach what the committed documents do not: text columns
#: and their notes, headers and footers, notes, text boxes and groups, charts.
PROBES = ("read_columns_probe", "read_story_probe", "read_footnote_draw_probe", "read_endnote_probe",
          "read_anchor_probe", "read_dml_probe", "read_chart_probe", "read_nested_table_probe")


def _faces(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))["faces"]


def _documents(source: str):
    """The committed documents, or one probe reader's documents: ``(name, bytes, faces)``."""
    if source == "committed":
        render = _faces(FIXTURES / "render-observations.json")
        for name in COMMITTED:
            yield name, (FIXTURES / name).read_bytes(), render
        return
    import importlib

    module = importlib.import_module(source)
    faces = _faces(Path(module.OBSERVATIONS))
    for name, data in module.documents():
        yield name, data, faces


def _drawn(layout):
    """``(what, path, the story's part or None, paraId or None)`` of everything drawn."""
    for page in layout.pages:
        for line in page.text_lines():
            yield "line", line.path, line.part, line.paragraph_id
            for span in line.spans:
                yield "span", span.path, line.part, None
        for placed in page.floats:
            yield "float", placed.path, placed.part, None
        if page.stop is not None:
            yield "stop", page.stop.path, None, None


@pytest.mark.parametrize("source", ("committed",) + PROBES)
def test_every_drawn_path_names_the_element_it_was_drawn_from(source):
    """Every line's path names a ``w:p`` (its ``w14:paraId`` the line's, where it has one),
    every span's a ``w:r`` or the element a computed field or label is drawn for, every
    float's a ``wp:anchor``; only Word's own default note separators, which the document
    does not hold, name nothing."""
    checked = 0
    for name, data, faces in _documents(source):
        layout = convert_docx_to_layout(data, render_record.options(render_record.RecordedFonts(faces)))
        package = Package.open(data)
        trees: dict = {}
        for what, path, part, para_id in _drawn(layout):
            where = path_part(package, path, part)
            if where not in trees:
                trees[where] = parse_xml(package.read(where)) if where and package.exists(where) else None
            element = resolve_path(trees[where], path)
            if path.startswith(("w:footnotes/", "w:endnotes/")):
                assert element is None, (name, path)
                continue
            assert element is not None, (name, what, path)
            last = [step for step in split_path(path) if step[0] is not None][-1][1]
            if path.endswith("/label"):
                last = "p"
            assert local_name(element.tag) == last, (name, what, path)
            if para_id is not None:
                assert element.get(W14_PARA_ID) == para_id, (name, path)
            checked += 1
    assert checked


def _with_body(body: str) -> bytes:
    """``sample-simple`` with its body replaced."""
    source = FIXTURES / "samplelib" / "sample-simple.docx"
    out = io.BytesIO()
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target:
        for item in original.infolist():
            data = original.read(item.filename)
            if item.filename == "word/document.xml":
                start, end = data.index(b"<w:body>") + len(b"<w:body>"), data.index(b"<w:sectPr")
                data = data[:start] + body.encode() + data[end:]
            target.writestr(item, data)
    return out.getvalue()


_CONTROLS = (
    '<w:p><w:r><w:t>before</w:t></w:r></w:p>'
    '<w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>in a control</w:t></w:r></w:p></w:sdtContent></w:sdt>'
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>'
    '<w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>'
    '<w:sdt><w:sdtPr/><w:sdtContent><w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:sdtContent></w:sdt></w:tr>'
    '<w:sdt><w:sdtPr/><w:sdtContent><w:tr><w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc>'
    '<w:tc><w:p><w:hyperlink><w:r><w:t>d</w:t></w:r></w:hyperlink></w:p></w:tc></w:tr></w:sdtContent></w:sdt>'
    '</w:tbl>'
)


def _text(element) -> str:
    return "".join(node.text or "" for node in element.iter() if local_name(node.tag) == "t")


@pytest.mark.parametrize("parser", ["etree", "lxml"])
def test_rows_and_cells_are_counted_through_content_controls(parser):
    """A cell in a cell-level control and a row in a row-level one are counted as the
    row's and the table's own; a block-level control's paragraph is under its
    ``w:sdtContent``; a run under a hyperlink, under it.  Resolved the same in the standard
    library's tree and lxml's."""
    data = _with_body(_CONTROLS)
    layout = convert_docx_to_layout(data, render_record.options(render_record.RecordedFonts(
        _faces(FIXTURES / "render-observations.json"))))
    spans = {"".join(span.chars): span.path for page in layout.pages for line in page.text_lines()
             for span in line.spans if span.chars}
    assert spans["in a control"] == "w:body/w:sdt[1]/w:sdtContent[1]/w:p[1]/w:r[1]"
    assert spans["b"] == "w:body/w:tbl[1]/w:tr[1]/w:tc[2]/w:p[1]/w:r[1]"
    assert spans["d"] == "w:body/w:tbl[1]/w:tr[2]/w:tc[2]/w:p[1]/w:hyperlink[1]/w:r[1]"
    xml = zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml")
    if parser == "lxml":
        etree = pytest.importorskip("lxml.etree")
        root = etree.fromstring(xml)
    else:
        root = ElementTree.fromstring(xml)
    for text, path in spans.items():
        assert _text(resolve_path(root, path)) == text
    assert resolve_path(root, "w:body/w:tbl[1]/w:tr[3]") is None
    assert resolve_path(root, "not a path") is None


def test_locate_finds_the_part_and_the_element():
    """:func:`locate` against the package: a body path in the main part, a header's in
    the part its line names, a note's in the notes part."""
    data = (FIXTURES / "samplelib" / "sample-simple.docx").read_bytes()
    part, element = locate(data, "w:body/w:p[1]/w:r[1]")
    assert part == "word/document.xml" and local_name(element.tag) == "r"
    assert locate(data, "w:hdr/w:p[1]", "word/header1.xml")[0] == "word/header1.xml"
    package = Package.open(data)
    assert path_part(package, "w:footnote[@w:id=2]/w:p[1]") in (None, *package.related(
        package.main_document_part, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"))


def test_the_svg_carries_the_layout_s_paths():
    """``data-docx-path`` is the layout object's path: the SVG's paragraph groups name
    the layout's lines, in order."""
    options = render_record.options(render_record.RecordedFonts(_faces(FIXTURES / "render-observations.json")))
    result = convert_docx(FIXTURES / "samplelib" / "sample-simple.docx", options)
    root = ElementTree.fromstring(result.svgs[0])
    groups = [g.get("data-docx-path") for g in root.iter("{http://www.w3.org/2000/svg}g")
              if g.get("data-docx-path") and g.get("data-docx-float") is None
              and g.get("data-docx-unsupported") is None]
    lines = []
    for line in result.layout.pages[0].text_lines():
        if not lines or lines[-1] != line.path:
            lines.append(line.path)
    assert [g for k, g in enumerate(groups) if k == 0 or groups[k - 1] != g] == lines
