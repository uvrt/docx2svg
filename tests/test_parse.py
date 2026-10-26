"""The reader, against the fixture we generated.

These run everywhere -- no Word, no PDF.  They check that what the reader says the file
asks for is what ``tools/make_layout_sweep.py`` wrote, which is the half of the loop that
does not need an oracle.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))

from docx2svg import parse_package  # noqa: E402
from docx2svg.opc import REL_STYLES, Package, content_type, main_part, read_part  # noqa: E402
from docx2svg.parse import parse_document  # noqa: E402
from ooxml_common.opc import OpcPackage  # noqa: E402

import make_layout_sweep as sweep  # noqa: E402


@pytest.fixture(scope="module")
def document(layout_sweep_path: Path):
    return parse_package(layout_sweep_path)


def test_the_committed_fixture_is_what_the_generator_writes(layout_sweep_path: Path):
    """The fixture is regenerable, and the committed one is current.

    Compared **part by part**, not archive byte by archive byte: a ZIP carries the
    compressor's output, and zlib does not promise the same bytes across versions or
    platforms.  The sibling project learned this from a fixture test that was red on one
    CI leg for a reason that had nothing to do with the fixture.
    """
    import io
    import zipfile

    committed = zipfile.ZipFile(layout_sweep_path)
    regenerated = zipfile.ZipFile(io.BytesIO(sweep.build_package()))
    assert committed.namelist() == regenerated.namelist()
    for name in committed.namelist():
        assert committed.read(name) == regenerated.read(name), name


def test_sections_are_found_inside_paragraph_properties(document):
    """Two sections, and the first one is not where a naive reader looks.

    Only the *last* section's ``w:sectPr`` is a child of ``w:body``.  A reader that
    looks only there reports one section for this document and silently lays the whole
    thing out on the wrong page size from paragraph 1.
    """
    assert len(document.sections) == 2

    first, second = document.sections
    assert (first.page_size.width_twips, first.page_size.height_twips) == (
        sweep.A4_WIDTH_TWIPS,
        sweep.A4_HEIGHT_TWIPS,
    )
    assert (second.page_size.width_twips, second.page_size.height_twips) == (
        sweep.ODD_WIDTH_TWIPS,
        sweep.ODD_HEIGHT_TWIPS,
    )
    assert first.margins.left == sweep.MARGIN_TWIPS
    assert second.margins.left == sweep.ODD_MARGIN_TWIPS
    # The paragraph carrying a sectPr belongs to the section it terminates.
    assert first.last_paragraph == second.first_paragraph


def test_the_indent_ladder_reads_back_as_authored(document):
    indents = [
        (paragraph.properties.indent.left if paragraph.properties.indent else 0) or 0
        for paragraph in document.paragraphs
        if paragraph.text.startswith("H indent ")
    ]
    assert indents == list(sweep.INDENT_LADDER_TWIPS)


def test_tab_stops_read_back_sorted(document):
    stops = [
        paragraph.properties.tab_stops
        for paragraph in document.paragraphs
        if paragraph.properties and paragraph.properties.tab_stops
    ]
    assert len(stops) == 1
    assert [stop.position_twips for stop in stops[0]] == list(sweep.TAB_STOPS_TWIPS)
    assert all(stop.alignment == "left" for stop in stops[0])


def test_a_tab_is_recorded_as_a_mark_not_as_whitespace(document):
    """A tab is a jump to a stop, not a width.

    Folding ``w:tab`` into the run text would make it indistinguishable from a space and
    would silently turn every tabbed line into a wrongly measured one.
    """
    paragraph = next(
        paragraph
        for paragraph in document.paragraphs
        if paragraph.properties and paragraph.properties.tab_stops
    )
    marks = [mark for run in paragraph.runs for mark in run.breaks]
    assert [kind for _, kind in marks] == ["tab"] * len(sweep.TAB_STOPS_TWIPS)
    assert "\t" not in paragraph.text


def test_the_pagination_block_is_all_there(document):
    # Matched by the label's shape, not by its first letter: the wrapping paragraph in
    # block D also begins with an "L", and a prefix match silently counted it.
    labels = [
        paragraph.text
        for paragraph in document.paragraphs
        if re.fullmatch(r"L\d\d", paragraph.text)
    ]
    assert len(labels) == sweep.PAGINATION_LINES
    assert labels[0] == "L01" and labels[-1] == f"L{sweep.PAGINATION_LINES:02d}"


def test_doc_defaults_are_read_from_the_styles_part(document):
    assert document.default_run_properties is not None
    assert document.default_run_properties.ascii_font == sweep.FACE
    assert document.default_run_properties.size_half_points == sweep.BODY_HALF_POINTS


def test_the_reader_names_what_it_dropped(document):
    """This fixture uses nothing the reader skips, so it must warn about nothing.

    The value of the assertion is the day it fails: a fixture that grows a table would
    otherwise gain a silently unread element.
    """
    assert document.warnings == []


def test_an_absent_val_on_a_toggle_means_true():
    """``<w:b/>`` is bold.  Only ``<w:b w:val="0"/>`` is not.

    The opposite of the usual XML convention, and the single most common way a reader
    of this format gets weights backwards.
    """
    namespace = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    xml = (
        f"<w:document {namespace}><w:body>"
        '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>on</w:t></w:r></w:p>'
        '<w:p><w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>off</w:t></w:r></w:p>'
        '<w:p><w:r><w:rPr/><w:t>unset</w:t></w:r></w:p>'
        "</w:body></w:document>"
    ).encode()
    document = parse_document(xml)
    weights = [paragraph.runs[0].properties.bold for paragraph in document.paragraphs]
    assert weights == [True, False, None]


def test_the_main_part_comes_from_the_relationship_not_a_hard_coded_path(layout_sweep_path):
    package = OpcPackage.open(layout_sweep_path)
    assert main_part(package) == "word/document.xml"
    assert package.related_parts_of_type("word/document.xml", "http://schemas.openxmlformats.org"
                                         "/officeDocument/2006/relationships/styles") == ["word/styles.xml"]


def _zip(parts: dict[str, str]) -> bytes:
    import io
    import zipfile

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, text in parts.items():
            archive.writestr(name, text)
    return buffer.getvalue()


RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"


def _renamed_package() -> bytes:
    return _zip({
        "[Content_Types].xml": f'<Types xmlns="{TYPES}"><Default Extension="xml" ContentType="application/xml"/>'
                               '<Override PartName="/doc/main.xml" ContentType="main"/></Types>',
        "_rels/.rels": f'<Relationships xmlns="{RELS}"><Relationship Id="rId1" Type="{OFFICE}/officeDocument" '
                       'Target="doc/main.xml"/></Relationships>',
        "doc/main.xml": "<document/>",
        "doc/_rels/main.xml.rels": f'<Relationships xmlns="{RELS}"><Relationship Id="rIdS" Type="{REL_STYLES}" '
                                   f'Target="../shared/styles.xml"/><Relationship Id="rIdL" Type="{OFFICE}/hyperlink" '
                                   'Target="https://example.com/" TargetMode="External"/></Relationships>',
        "media/picture.png": "not a picture",
    })


def test_word_s_readings_of_a_package():
    """What docx2svg.opc reads where OpcPackage answers otherwise: no root relationship
    means ``word/document.xml``, a missing part raises, and a type the package does not
    state is ``""`` (OpcPackage would guess ``image/png`` from the extension)."""
    package = OpcPackage.open(_renamed_package())
    assert main_part(package) == "doc/main.xml"
    assert main_part(OpcPackage(parts={})) == "word/document.xml"
    assert package.related_parts_of_type("doc/main.xml", REL_STYLES) == ["shared/styles.xml"]
    assert read_part(package, "/doc/main.xml") == b"<document/>"
    with pytest.raises(KeyError, match="shared/styles.xml"):
        read_part(package, "shared/styles.xml")
    assert content_type(package, "doc/main.xml") == "main"
    assert content_type(package, "doc/other.xml") == "application/xml"
    assert content_type(package, "media/picture.png") == ""
    assert package.content_type("media/picture.png") == "image/png"


def test_the_old_reader_s_names_are_kept_for_a_release():
    package = Package.open(_renamed_package())
    assert isinstance(package, OpcPackage)
    assert package.main_document_part == "doc/main.xml"
    assert package.related("doc/main.xml", REL_STYLES) == ["shared/styles.xml"]
    assert package.part_by_id("doc/main.xml", "rIdS") == "shared/styles.xml"
    assert package.part_by_id("doc/main.xml", "rIdL") is None
    assert package.exists("doc/main.xml") and not package.exists("shared/styles.xml")
    assert package.read("doc/main.xml") == b"<document/>"
    with pytest.raises(KeyError):
        package.read("shared/styles.xml")
    assert package.content_type("media/picture.png") == ""
    assert Package.rels_part_for("doc/main.xml") == "doc/_rels/main.xml.rels"
    assert Package.rels_part_for("") == "_rels/.rels"
    relationships = package.relationships("doc/main.xml")
    assert package.resolve("doc/main.xml", relationships["rIdS"]) == "shared/styles.xml"
    assert package.resolve("doc/main.xml", relationships["rIdL"]) == "https://example.com/"


def test_the_compatibility_mode_is_read_and_unstated_stays_unstated():
    import make_page_top_probe as probe

    from docx2svg import parse_package

    modes = {setting: parse_package(probe.build(probe.Probe(setting, 0))).compatibility_mode
             for setting in probe.SETTINGS}
    assert modes == {"none": None, "nocompat": None, "compat-empty": None,
                     "12": 12, "14": 14, "15": 15, "14-suppress": 14}


def test_legacy_compat_options_are_read_by_name():
    import make_autospacing_probe as autospacing
    import make_page_top_probe as page_top

    from docx2svg import parse_package

    assert parse_package(page_top.build(page_top.Probe("14-suppress", 0))).compat_options == {
        "suppressSpBfAfterPgBrk"}
    assert parse_package(autospacing.build(autospacing.Probe("15-noHTML"))).compat_options == {
        "doNotUseHTMLParagraphAutoSpacing"}
    assert parse_package(autospacing.build(autospacing.Probe("15"))).compat_options == frozenset()
