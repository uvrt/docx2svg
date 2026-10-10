"""The application's own font folders: ``font_dirs``, or ``OOXML_FONT_DIRS``.

Production: an application keeps its licensed faces in a folder the system does not
search.  Named with ``ConvertOptions(font_dirs=...)`` or the environment variable, a face
there is laid out and drawn; not named, the layout reports it missing and incomplete.  The
face is an open one -- Cousine from the ``pptx2svg-fonts`` bundle, or Liberation Mono
where CI installs it -- relabelled in a temporary folder as a family nothing else answers
to, so no system copy can stand in for it.
"""

from __future__ import annotations

import io
import os
import zipfile
from pathlib import Path

import pytest

import docx2svg
from docx2svg import ConvertOptions, fonts
from ooxml_common.fonts import bundle_dir
from ooxml_common.fonts.office import FONT_DIRS_ENV
from ooxml_common.fonts.sfnt import relabel

FAMILY = "Fontdirs Probe Mono"
_W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _open_face() -> Path | None:
    """An open monospaced face to relabel: the bundle's Cousine, else Liberation Mono."""
    bundle = bundle_dir()
    if bundle is not None and (bundle / "Cousine-Regular.ttf").is_file():
        return bundle / "Cousine-Regular.ttf"
    found = fonts.installed_index(tuple(fonts.default_font_dirs())).get(("liberation mono", False, False))
    return Path(found[0]) if found and found[0].endswith(".ttf") and found[1] == 0 else None


@pytest.fixture(autouse=True)
def _no_environment(monkeypatch):
    monkeypatch.delenv(FONT_DIRS_ENV, raising=False)


@pytest.fixture
def folder(tmp_path) -> Path:
    source = _open_face()
    if source is None:
        pytest.skip("no open monospaced face here (pptx2svg-fonts, or fonts-liberation)")
    target = tmp_path / "app-fonts" / "probe"
    target.mkdir(parents=True)
    data = relabel(source.read_bytes(), FAMILY, bold=False, italic=False)
    (target / "FontdirsProbeMono-Regular.ttf").write_bytes(data)
    return tmp_path / "app-fonts"


def _document() -> bytes:
    """One page of text in the probe's family, A4."""
    run = f'<w:r><w:rPr><w:rFonts w:ascii="{FAMILY}" w:hAnsi="{FAMILY}"/></w:rPr><w:t>Licensed face</w:t></w:r>'
    styles = (f'<w:styles {_W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="{FAMILY}" '
              f'w:hAnsi="{FAMILY}" w:cs="{FAMILY}"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
              '</w:styles>')
    document = (f'<w:document {_W}><w:body><w:p>{run}</w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
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


def _coverage(options: ConvertOptions):
    return docx2svg.convert_docx(_document(), options).layout.coverage


def test_the_folder_is_laid_out_from_only_when_named(folder, monkeypatch):
    unnamed = _coverage(ConvertOptions())
    assert not unnamed.complete and unnamed.missing_fonts == [FAMILY]

    named = _coverage(ConvertOptions(font_dirs=[str(folder)]))
    assert named.complete and named.missing_fonts == []

    monkeypatch.setenv(FONT_DIRS_ENV, os.pathsep.join([str(folder / "absent"), str(folder)]))
    from_env = _coverage(ConvertOptions())
    assert from_env.complete and from_env.missing_fonts == []

    # An explicit empty list wins over the variable: none.
    none = _coverage(ConvertOptions(font_dirs=[]))
    assert not none.complete and none.missing_fonts == [FAMILY]


def test_the_folder_is_drawn_with(folder, monkeypatch):
    if "resvg" not in docx2svg.available_backends():
        pytest.skip("needs resvg-py")
    named = docx2svg.convert_docx_to_png(_document(), ConvertOptions(width=400, font_dirs=[str(folder)]))
    monkeypatch.setenv(FONT_DIRS_ENV, str(folder))
    assert docx2svg.convert_docx_to_png(_document(), ConvertOptions(width=400)) == named
