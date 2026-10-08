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
from docx2svg.opc import Package  # noqa: E402
from docx2svg.parse import parse_document  # noqa: E402

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
    package = Package.open(layout_sweep_path)
    assert package.main_document_part == "word/document.xml"
    assert package.related("word/document.xml", "http://schemas.openxmlformats.org"
                           "/officeDocument/2006/relationships/styles") == ["word/styles.xml"]


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
