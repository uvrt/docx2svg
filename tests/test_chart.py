"""Charts against what Word drew, offline (ROADMAP.md, "Floating drawings -- measured",
F.20): ``tools/make_chart_probe.py``'s 58 charts, recorded by ``tools/read_chart_probe.py``
in ``tests/fixtures/chart-observations.json`` -- every text span, fill and stroke Word drew
-- and drawn here by ``ooxml-common``'s chart layout under its ``WORD`` rules
(``docx2svg.chart``).

What is left is pinned by family: the radar's labels (nine, half a point to 2.4 pt off),
five labels of charts whose text is 6, 14 or 18 pt (0.55 to 0.6 pt), and one scatter marker
Word draws 0.6 pt off its point (two fills and four strokes, counting the model's own).
"""

from __future__ import annotations

import io
import json
import sys
import zipfile
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import make_chart_probe  # noqa: E402
import read_chart_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
#: What the comparison is known to miss, per family (module docstring).
KNOWN = {"text": 14, "fills": 2, "strokes": 4}


@pytest.mark.parametrize("name,data", reader.documents(), ids=[name for name, _ in reader.documents()])
def test_charts_against_word(name, data):
    recorded = DATA["documents"][name]
    fonts = render_record.RecordedFonts(DATA["faces"])
    result = reader.score(data, fonts, recorded["word"])
    assert result["scores"] == recorded["scores"]
    for family, (agree, compared) in result["scores"].items():
        assert compared - agree == KNOWN[family], (family, result["problems"])


def _render(data: bytes):
    from docx2svg import ConvertOptions, convert_docx_to_layout

    options = ConvertOptions()
    layout = convert_docx_to_layout(data, options)
    return layout, [str(warning) for warning in options.warnings]


def _without(data: bytes, drop: str = "", replace: dict | None = None) -> bytes:
    source = zipfile.ZipFile(io.BytesIO(data))
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for info in source.infolist():
            if info.filename == drop:
                continue
            archive.writestr(info, (replace or {}).get(info.filename, source.read(info.filename)))
    return buffer.getvalue()


def _one(case: make_chart_probe.Case) -> bytes:
    make_chart_probe.CASES, saved = (case,), make_chart_probe.CASES
    try:
        return make_chart_probe.build("15")
    finally:
        make_chart_probe.CASES = saved


@pytest.mark.faces
def test_a_chart_is_drawn_inline_and_floating_in_its_frame():
    for where in ("inline", "float"):
        layout, warnings = _render(_one(make_chart_probe.Case("text", "one", make_chart_probe.column_chart(),
                                                              where=where)))
        drawn = [p for placed in layout.pages[0].floats for p in placed.primitives
                 if p.kind == "markup" and p.what == "chart"]
        assert len(drawn) == 1 and "Sales by region" in drawn[0].markup, where
        assert not [w for w in warnings if "chart" in w or "drawing-not-drawn" in w], warnings


@pytest.mark.faces
def test_a_chart_part_that_is_missing_is_a_placeholder_and_a_warning():
    data = _without(_one(make_chart_probe.CASES[0]), drop="word/charts/chart1.xml")
    layout, warnings = _render(data)
    kinds = [p.kind for placed in layout.pages[0].floats for p in placed.primitives]
    assert kinds == ["placeholder"]
    assert any("chart-unreadable" in w for w in warnings) and any("drawing-not-drawn" in w for w in warnings)


@pytest.mark.faces
def test_a_chart_type_not_drawn_is_said():
    space = make_chart_probe.chart_space(
        make_chart_probe.NO_TITLE + "<c:plotArea><c:layout/><c:unknownChart/></c:plotArea>")
    layout, warnings = _render(_one(make_chart_probe.Case("type", "unknown", space)))
    assert any("chart-unsupported-type" in w for w in warnings), warnings


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RELS = "http://schemas.openxmlformats.org/package/2006/relationships"


@pytest.mark.faces
def test_a_chart_in_a_header_is_related_from_the_header():
    """The chart part a header's drawing names is the header's relationship, not the
    document's -- here the document has no relationship of that id at all."""
    drawing = make_chart_probe.inline("rId1", 1, make_chart_probe.CX // 2, make_chart_probe.CY // 2)
    parts = {
        "[Content_Types].xml": (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.'
            'wordprocessingml.document.main+xml"/>'
            '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.'
            'wordprocessingml.header+xml"/>'
            f'<Override PartName="/word/charts/chart1.xml" ContentType="{make_chart_probe.CHART_TYPE}"/></Types>'),
        "_rels/.rels": (f'<Relationships xmlns="{RELS}"><Relationship Id="rId1" Type="{R}/officeDocument" '
                        'Target="word/document.xml"/></Relationships>'),
        "word/_rels/document.xml.rels": (f'<Relationships xmlns="{RELS}"><Relationship Id="rIdH" Type="{R}/header" '
                                         'Target="header1.xml"/></Relationships>'),
        "word/_rels/header1.xml.rels": (f'<Relationships xmlns="{RELS}"><Relationship Id="rId1" '
                                        f'Type="{make_chart_probe.CHART_REL}" Target="charts/chart1.xml"/>'
                                        "</Relationships>"),
        "word/header1.xml": f'<w:hdr xmlns:w="{W}" xmlns:r="{R}"><w:p>{drawing}</w:p></w:hdr>',
        "word/document.xml": (f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p>'
                              '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>'
                              '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="4000" w:right="1440" w:bottom="1440" '
                              'w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body>'
                              "</w:document>"),
        "word/charts/chart1.xml": make_chart_probe.column_chart(),
    }
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, text in parts.items():
            archive.writestr(name, text)
    layout, warnings = _render(buffer.getvalue())
    drawn = [p for placed in layout.pages[0].floats for p in placed.primitives if p.kind == "markup"]
    assert len(drawn) == 1 and drawn[0].what == "chart" and "Sales by region" in drawn[0].markup
    assert not [w for w in warnings if "chart" in w], warnings


@pytest.mark.faces
def test_a_chart_in_a_group_s_graphic_frame_is_drawn_in_its_place():
    import make_drawing_probe as drawing_probe

    frame = (f'<wpg:graphicFrame><wpg:cNvPr id="7" name="Chart"/><wpg:cNvFrPr/>'
             f'<wpg:xfrm><a:off x="1000000" y="0"/><a:ext cx="2000000" cy="1500000"/></wpg:xfrm>'
             f'<a:graphic><a:graphicData uri="{make_chart_probe.C_NS}"><c:chart xmlns:c="{make_chart_probe.C_NS}" '
             f'xmlns:r="{R}" r:id="rId4"/></a:graphicData></a:graphic></wpg:graphicFrame>')
    content = drawing_probe.group(3000000, 1500000, frame)
    anchor = make_chart_probe.anchor_probe.Anchor(("margin", "offset", 0), ("paragraph", "offset", 0), 3000000,
                                                  1500000, graphic=drawing_probe.graphic(content, "wpg")).xml(1)
    case = make_chart_probe.Case("place", "grouped", make_chart_probe.column_chart())
    data = _one(case)
    source = zipfile.ZipFile(io.BytesIO(data))
    document = source.read("word/document.xml").decode()
    start = document.index("<w:drawing")
    end = document.index("</w:drawing>") + len("</w:drawing>")
    run_start = document.rindex("<w:r>", 0, start)
    run_end = document.index("</w:r>", end) + len("</w:r>")
    document = document[:run_start] + anchor + document[run_end:]
    layout, warnings = _render(_without(data, replace={"word/document.xml": document.encode()}))
    drawn = [p for placed in layout.pages[0].floats for p in placed.primitives if p.kind == "markup"]
    assert len(drawn) == 1 and drawn[0].what == "chart", warnings
    # The frame's place in the group: 1,000,000 EMU in from the group's left, at 300 dpi (the
    # group's box is held in whole twips, F.3).
    x = float(drawn[0].transform.split("(")[1].split()[0])
    assert x == pytest.approx(float(layout.pages[0].floats[0].x) + 1000000 * 300 / 914400, abs=0.1)


# -- charts in a footnote and in a text box ------------------------------------------------

def _places() -> bytes:
    """A footnote holding an inline chart, and a text box holding another, related from
    the notes part and the document part each (the same chart part)."""
    import io
    import zipfile

    import make_anchor_probe as anchor_probe
    import probe_docx

    chart = make_chart_probe.column_chart(heading=None, where="b")
    inline = make_chart_probe.inline("rIdC", 1, 2743200, 1600200)
    note = (f'<w:footnotes xmlns:w="{W}" xmlns:r="{make_chart_probe.R_NS}"><w:footnote w:type="separator" '
            'w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" '
            'w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p>'
            '<w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> A chart in a note:</w:t></w:r></w:p>'
            f'<w:p>{inline}</w:p></w:footnote></w:footnotes>')
    box = ('<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
           '<wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2926080" cy="2240280"/></a:xfrm>'
           '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent>'
           f'<w:p>{make_chart_probe.inline("rIdC", 2, 2743200, 1600200)}</w:p></w:txbxContent></wps:txbx>'
           '<wps:bodyPr/></wps:wsp>')
    graphic = (f'<a:graphic xmlns:a="{make_chart_probe.A_NS}"><a:graphicData '
               'uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
               f"{box}</a:graphicData></a:graphic>")
    anchored = anchor_probe.Anchor(("column", "offset", 0), ("paragraph", "offset", 0), 2926080, 2240280,
                                   wrap="<wp:wrapTopAndBottom/>", graphic=graphic).xml(3)
    body = ('<w:p><w:r><w:t>Charts.</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
            f"<w:p><w:r><w:t>A text box.</w:t></w:r>{anchored}</w:p>")
    data = probe_docx.package(body, final_section=anchor_probe.section())
    source = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    rel = (f'<Relationship Id="rIdC" Type="{make_chart_probe.CHART_REL}" Target="charts/chart1.xml"/>')
    with zipfile.ZipFile(out, "w") as archive:
        for info in source.infolist():
            content = source.read(info.filename)
            if info.filename == "word/_rels/document.xml.rels":
                content = content.replace(b"</Relationships>", (
                    rel + '<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                    'relationships/footnotes" Target="footnotes.xml"/></Relationships>').encode())
            elif info.filename == "[Content_Types].xml":
                content = content.replace(b"</Types>", (
                    '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-'
                    'officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/charts/chart1.xml" '
                    f'ContentType="{make_chart_probe.CHART_TYPE}"/></Types>').encode())
            archive.writestr(info, content)
        archive.writestr("word/footnotes.xml", note)
        archive.writestr("word/_rels/footnotes.xml.rels", (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f"{rel}</Relationships>"))
        archive.writestr("word/charts/chart1.xml", chart)
    return out.getvalue()


@pytest.mark.faces
def test_a_chart_in_a_footnote_is_related_from_the_notes_part_and_one_in_a_text_box_is_drawn():
    from docx2svg import ConvertOptions, convert_docx_to_layout

    options = ConvertOptions()
    layout = convert_docx_to_layout(_places(), options)
    drawn = {p.path.split("/")[0]: p for page in layout.pages for placed in page.floats for p in placed.primitives
             if p.kind == "markup" and p.what == "chart"}
    assert set(drawn) == {"w:footnote[@w:id=1]", "w:body"}, drawn
    assert "/wps:txbx/w:txbxContent/w:p[1]/w:r[1]" in drawn["w:body"].path
    assert not [w for w in options.warnings if "chart" in str(w) or "drawing-not-drawn" in str(w)], options.warnings
