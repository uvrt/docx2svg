"""Theme colours on text and shading read as Word draws them (production round 21, D2):
``w:themeColor`` / ``w:themeFill`` with their tint or shade, not the ``w:val`` / ``w:fill``
beside them (measured with decoy values; the arithmetic is ``ooxml_common.themetint``)."""

from __future__ import annotations

import io
import zipfile
from xml.etree.ElementTree import fromstring

from docx2svg.parse.document import parse_package, word_theme_colour
from docx2svg.parse.properties import (read_paragraph_properties, read_run_properties, read_table_properties,
                                       theme_colours)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
THEME = {"dk1": "000000", "lt1": "FFFFFF", "dk2": "0E2841", "lt2": "E8E8E8", "accent1": "156082",
         "accent2": "E97132", "hlink": "467886"}


def element(xml: str):
    return fromstring(f'<w:x xmlns:w="{W}">{xml}</w:x>')[0]


def test_text_and_shading_take_the_theme_colour():
    run = element('<w:rPr><w:color w:val="FF00FF" w:themeColor="accent1" w:themeShade="BF"/></w:rPr>')
    paragraph = element('<w:pPr><w:shd w:val="clear" w:color="auto" w:fill="00FF00" w:themeFill="background2"'
                        ' w:themeFillTint="BF"/></w:pPr>')
    cell = element('<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="00FF00" w:themeFill="accent2"'
                   ' w:themeFillShade="80"/></w:tcPr>')
    with theme_colours(word_theme_colour(THEME, {})):
        assert read_run_properties(run)["color"] == "0F4761"
        assert read_paragraph_properties(paragraph)["shd"] == ("clear", "auto", "EDEDED")
        assert read_table_properties(cell)["shd"] == ("clear", "auto", "80340D")
    # Without the package's theme (a part parsed alone): the stated values, as before.
    assert read_run_properties(run)["color"] == "FF00FF"
    assert read_paragraph_properties(paragraph)["shd"][2] == "00FF00"


def test_a_runs_shading_takes_the_theme_colour():
    # Round 24 (measured: a character style's and a direct run's w:themeFill drawn, the
    # decoy w:fill beside it not): accent1 tint 66 -> 83CAEB, accent2 shade BF -> BF4E14.
    tinted = element('<w:rPr><w:shd w:val="clear" w:color="auto" w:fill="00FF00" w:themeFill="accent1"'
                     ' w:themeFillTint="66"/></w:rPr>')
    shaded = element('<w:rPr><w:shd w:val="clear" w:color="auto" w:fill="FF00FF" w:themeFill="accent2"'
                     ' w:themeFillShade="BF"/></w:rPr>')
    with theme_colours(word_theme_colour(THEME, {})):
        assert read_run_properties(tinted)["shd"] == ("clear", "auto", "83CAEB")
        assert read_run_properties(shaded)["shd"] == ("clear", "auto", "BF4E14")
    assert read_run_properties(tinted)["shd"][2] == "00FF00"


def test_the_text_and_background_names_go_through_the_mapping():
    resolve = word_theme_colour(THEME, {"t1": "light1", "bg1": "dark1"})
    assert resolve("text1", None, None) == "FFFFFF"
    assert resolve("background1", None, "80") == "000000"
    assert word_theme_colour(THEME, {})("text1", "BF", None) == "404040"
    assert resolve("hyperlink", None, None) == "467886"
    assert resolve("accent6", None, None) is None             # not in the theme: the stated value stays


def test_a_package_reads_its_theme():
    parts = {
        "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '</Types>',
        "_rels/.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
        'officeDocument" Target="word/document.xml"/></Relationships>',
        "word/_rels/document.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
        'relationships"><Relationship Id="rT" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/theme" Target="theme/theme1.xml"/></Relationships>',
        "word/theme/theme1.xml": '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:themeElements><a:clrScheme name="T"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:accent1>'
        '<a:srgbClr val="156082"/></a:accent1></a:clrScheme></a:themeElements></a:theme>',
        "word/document.xml": f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:rPr><w:color w:val="FF00FF" '
        'w:themeColor="accent1" w:themeTint="99"/></w:rPr><w:t>Tinted</w:t></w:r></w:p></w:body></w:document>',
    }
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, text in parts.items():
            archive.writestr(name, text)
    document = parse_package(buffer.getvalue())
    run = document.body[0].runs[0]
    assert run.properties.declared["color"] == run.properties.color == "45B0E1"
