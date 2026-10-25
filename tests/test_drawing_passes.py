"""Two faces the rasteriser files under one name, drawn in passes of their own.

resvg files a face under its typographic family (name ID 16) alone: Aptos Display is
"Aptos" at 400, as Aptos is, and of two faces that tie the first loaded draws both -- a
document with headings in one and body text in the other drew every line in one face,
and warned ``face-not-addressable``.  :func:`docx2svg.fonts.drawing_plan` draws each in a
pass without the other's file (as the sibling pptx2svg does since round 10).

The faces are made up: two open faces (the ``pptx2svg-fonts`` bundle's Tinos and Carlito,
or Liberation Serif and Carlito where CI installs them) given names in a temporary folder
-- "Body Face Display" and "Body Face", both filed as "Body Face" -- so nothing else
answers to them and nothing Microsoft's is used.
"""

from __future__ import annotations

import io
import struct
import zipfile
import zlib
from pathlib import Path

import pytest

import docx2svg
from docx2svg import ConvertOptions, fonts
from ooxml_common.fonts import bundle_dir
from ooxml_common.fonts.office import FONT_DIRS_ENV

DISPLAY, BODY = "Body Face Display", "Body Face"
_W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _open_faces() -> tuple[Path, Path] | None:
    """Two open faces with different outlines: a serif and a sans."""
    bundle = bundle_dir()
    if bundle is not None and (bundle / "Tinos-Regular.ttf").is_file() \
            and (bundle / "Carlito-Regular.ttf").is_file():
        return bundle / "Tinos-Regular.ttf", bundle / "Carlito-Regular.ttf"
    index = fonts.installed_index(tuple(fonts.default_font_dirs()))
    found = [index.get((family, False, False)) for family in ("liberation serif", "carlito")]
    if all(item and item[1] == 0 and item[0].lower().endswith(".ttf") for item in found):
        return Path(found[0][0]), Path(found[1][0])
    return None


def _named(source: Path, family: str, typographic: str) -> bytes:
    """``source`` with a made-up name table: name ID 1 ``family``, ID 16 ``typographic``."""
    from ooxml_common.fonts.sfnt import read_tables, write_sfnt

    records = [(1, family), (2, "Regular"), (4, f"{family} Regular"), (6, family.replace(" ", "")),
               (16, typographic), (17, "Regular")]
    strings = entries = b""
    for name_id, text in records:
        raw = text.encode("utf-16-be")
        entries += struct.pack(">HHHHHH", 3, 1, 0x409, name_id, len(raw), len(strings))
        strings += raw
    name = struct.pack(">HHH", 0, len(records), 6 + 12 * len(records)) + entries + strings
    tables = read_tables(source.read_bytes())
    tables[b"name"] = name
    return write_sfnt(list(tables.items()))


@pytest.fixture(autouse=True)
def _no_environment(monkeypatch):
    monkeypatch.delenv(FONT_DIRS_ENV, raising=False)


@pytest.fixture
def folder(tmp_path) -> Path:
    """The application's folder holding both faces; the body face sorts (and loads) first."""
    faces = _open_faces()
    if faces is None:
        pytest.skip("no two open faces here (pptx2svg-fonts, or fonts-liberation and Carlito)")
    target = tmp_path / "app-fonts"
    target.mkdir()
    (target / "a-body.ttf").write_bytes(_named(faces[1], BODY, BODY))
    (target / "b-display.ttf").write_bytes(_named(faces[0], DISPLAY, BODY))
    return target


def _document(*families: str) -> bytes:
    """One page, a paragraph per family, each "Handgloves" at 28 pt."""
    paragraphs = "".join(
        f'<w:p><w:r><w:rPr><w:rFonts w:ascii="{family}" w:hAnsi="{family}"/><w:sz w:val="56"/></w:rPr>'
        '<w:t>Handgloves</w:t></w:r></w:p>' for family in families)
    styles = (f'<w:styles {_W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="{BODY}" '
              f'w:hAnsi="{BODY}" w:cs="{BODY}"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
              '</w:styles>')
    document = (f'<w:document {_W}><w:body>{paragraphs}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" '
                'w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>')
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("[Content_Types].xml", (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" '
            'ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" '
            f'ContentType="application/xml"/><Override PartName="/word/document.xml" '
            f'ContentType="{_CT}.document.main+xml"/><Override PartName="/word/styles.xml" '
            f'ContentType="{_CT}.styles+xml"/></Types>'))
        archive.writestr("_rels/.rels", (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{_REL}/officeDocument" Target="word/document.xml"/></Relationships>'))
        archive.writestr("word/_rels/document.xml.rels", (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{_REL}/styles" Target="styles.xml"/></Relationships>'))
        archive.writestr("word/document.xml", document)
        archive.writestr("word/styles.xml", styles)
    return buffer.getvalue()


def _needs_resvg() -> None:
    if "resvg" not in docx2svg.available_backends():
        pytest.skip("needs resvg-py")


def _rows(png: bytes, top: int, bottom: int) -> list[bytes]:
    """Rows ``top`` to ``bottom`` of a PNG as resvg writes it (8-bit RGBA, not interlaced),
    unfiltered here: no image library needed."""
    width, height, depth, colour = struct.unpack(">IIBB", png[16:26])
    assert (depth, colour) == (8, 6)
    data, at = b"", 8
    while at < len(png):
        length, kind = struct.unpack(">I4s", png[at:at + 8])
        if kind == b"IDAT":
            data += png[at + 8:at + 8 + length]
        at += 12 + length
    raw, stride, rows, previous = zlib.decompress(data), width * 4, [], bytes(width * 4)
    for y in range(height):
        kind, line = raw[y * (stride + 1)], bytearray(raw[y * (stride + 1) + 1:(y + 1) * (stride + 1)])
        for x in range(stride):
            left = line[x - 4] if x >= 4 else 0
            up, corner = previous[x], previous[x - 4] if x >= 4 else 0
            if kind == 1:
                line[x] = (line[x] + left) & 255
            elif kind == 2:
                line[x] = (line[x] + up) & 255
            elif kind == 3:
                line[x] = (line[x] + (left + up) // 2) & 255
            elif kind == 4:
                estimate = left + up - corner
                pa, pb, pc = abs(estimate - left), abs(estimate - up), abs(estimate - corner)
                line[x] = (line[x] + (left if pa <= pb and pa <= pc else up if pb <= pc else corner)) & 255
        rows.append(bytes(line))
        previous = line
    return rows[top:bottom]


def test_a_heading_face_filed_under_the_body_face_s_name_draws_its_own_text(folder):
    """The first line in the display face, as a page of it alone draws it -- not in the
    body face, which the rasteriser loads first and which ties with it -- and no warning."""
    _needs_resvg()
    options = ConvertOptions(width=600, font_dirs=[str(folder)])
    both = docx2svg.convert_docx_to_png(_document(DISPLAY, BODY), options)[0]
    assert not [w for w in options.warnings if w.code == "face-not-addressable"]
    display = docx2svg.convert_docx_to_png(_document(DISPLAY, DISPLAY), ConvertOptions(
        width=600, font_dirs=[str(folder)]))[0]
    body = docx2svg.convert_docx_to_png(_document(BODY, BODY), ConvertOptions(
        width=600, font_dirs=[str(folder)]))[0]
    first_line = (60, 112)
    assert _rows(both, *first_line) == _rows(display, *first_line)
    assert _rows(both, *first_line) != _rows(body, *first_line)


def test_the_plan_leaves_the_other_face_s_file_out_of_each_pass(folder):
    """The display face's text in a pass without the body face's file, the body text in
    a pass of its own; a page that needs no
    pass (one family, loaded first) is drawn as it was."""
    options = ConvertOptions(width=600, font_dirs=[str(folder)])
    layout, data, found = docx2svg.lay_out_docx(_document(DISPLAY, BODY), options)
    svg = docx2svg.page_svgs(layout, data, options, device_size=True, fonts=found)[0]
    files, missing, unaddressable = fonts.rasteriser_files(found, docx2svg.drawn_faces(layout))
    assert not missing and not unaddressable
    body = str(folder / "a-body.ttf")
    plan = fonts.drawing_plan(svg, files, [str(folder)])
    assert plan is not None and plan.passes == 2
    assert plan.withheld == [frozenset({body}), frozenset()]
    assert list(plan.stacks.values()) == [1] and BODY in next(iter(plan.stacks))

    layout, data, found = docx2svg.lay_out_docx(_document(BODY, BODY), options)
    svg = docx2svg.page_svgs(layout, data, options, device_size=True, fonts=found)[0]
    files, _, _ = fonts.rasteriser_files(found, docx2svg.drawn_faces(layout))
    assert fonts.drawing_plan(svg, files, [str(folder)]) is None
