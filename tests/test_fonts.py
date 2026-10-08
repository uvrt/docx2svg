"""``docx2svg.fonts``: the standard-library reader answers exactly what the ``fontTools``
tools recorded.

Every recording under ``tests/fixtures`` holds numbers ``tools/face_metrics.py`` and
``tools/face_advances.py`` read from the installed faces with ``fontTools``: the four
vertical integers and the ``OS/2`` script sizes of every face, and every advance and
kern pair a probe or document asked for.  The renderer reads the same faces with
:mod:`struct` alone (Phase 5 needs them at run time, where no dev extra is installed),
so it must give the same answer for every one of them.

Needs the faces installed where Word finds them; skipped where they are not (Linux CI).
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx2svg.fonts import InstalledFonts, installed_index

FIXTURES = Path(__file__).parent / "fixtures"
SCRATCH_SAMPLE1 = Path(__file__).parent.parent / "scratch" / "filesamples" / "sample1.docx"


def _recordings():
    for path in sorted(FIXTURES.glob("*.json")):
        data = json.loads(path.read_text())
        if not isinstance(data, dict):
            continue
        faces = data.get("faces") or {}
        advances = data.get("advances")
        if advances is None and faces and isinstance(next(iter(faces.values())), dict):
            advances, faces = faces, {}
        yield path.name, faces, advances or {}


def _fonts() -> InstalledFonts:
    if not any(key[0] == "calibri" for key in installed_index()):
        pytest.skip("Word's faces are not installed here")
    return InstalledFonts()


def _key(key: str):
    parts = key.split("|")
    return (parts[0], parts[1] == "1", parts[2] == "1") if len(parts) == 3 else None


def test_every_recorded_metric():
    fonts = _fonts()
    checked = 0
    for name, faces, _ in _recordings():
        for key, integers in faces.items():
            face = _key(key)
            if face is None or not isinstance(integers, list):
                continue
            found = fonts.metrics(*face)
            got = [found.units_per_em, found.ascent, found.descent, found.line_gap]
            if len(integers) > 4:
                got += [found.superscript_size, found.subscript_size]
            assert got == integers, (name, key)
            checked += 1
    assert checked > 50


def test_every_recorded_advance_and_kern_pair():
    fonts = _fonts()
    checked = 0
    for name, _, advances in _recordings():
        for key, entry in advances.items():
            face = _key(key)
            if face is None or not isinstance(entry, dict) or "advances" not in entry:
                continue
            for char, width in entry["advances"].items():
                assert fonts.advance(*face, char) == (width, entry["upm"]), (name, key, char)
                checked += 1
            for pair, value in entry.get("kern", {}).items():
                assert fonts.kern(*face, pair[0], pair[1]) == value, (name, key, pair)
                checked += 1
    assert checked > 10_000


def test_an_embedded_face_where_none_is_installed():
    """``filesamples/sample1`` embeds Ubuntu, which no installed folder holds; Word draws it,
    and ``tools/face_metrics.embedded`` recorded its integers.  Office's cloud-font cache may
    hold Ubuntu too (Office downloads it): where it does, its copy is found first, and its
    metrics and advances are the embedded face's."""
    if not SCRATCH_SAMPLE1.is_file():
        pytest.skip("scratch/filesamples is absent")
    _fonts()
    from docx2svg.fonts import FONT_DIRS, cloud_font_dirs, embedded_faces

    data = SCRATCH_SAMPLE1.read_bytes()
    installed = tuple(d for d in FONT_DIRS if d not in cloud_font_dirs())
    fonts = InstalledFonts(data, dirs=installed)
    metrics = fonts.metrics("Ubuntu")
    assert metrics is not None
    assert fonts.advance("Ubuntu", False, False, "H") is not None
    assert InstalledFonts(dirs=installed).metrics("Ubuntu") is None
    cloud = InstalledFonts(data).face("Ubuntu")
    if cloud is not None and not cloud.source.startswith("embedded:"):
        embedded = embedded_faces(data)[("ubuntu", False, False)]
        assert cloud.metrics == embedded.metrics
        assert all(cloud.advance(char) == embedded.advance(char) for char in "Hamburgefonstiv 0123456789")


# -- naming a face for the rasteriser, without copying it -----------------------------------


class _Fake:
    """A face as :func:`docx2svg.fonts.css_match` sees it."""

    def __init__(self, name: str, weight: int, width: int = 5, style: str = "normal") -> None:
        self.name, self.css = name, (weight, width, style)
        self.rasteriser_families = ("F",)

    def __repr__(self) -> str:
        return self.name


def test_css_matching_is_resvgs():
    """CSS Fonts 3 5.2 as resvg's database applies it: stretch, then style, then weight
    (400-500 looks up to 500 first, lighter below 400, heavier above 500)."""
    from docx2svg.fonts import css_match

    light, regular, medium, semibold, bold = (_Fake(n, w) for n, w in (
        ("light", 300), ("regular", 400), ("medium", 500), ("semibold", 600), ("bold", 700)))
    family = [light, regular, medium, semibold, bold]
    assert css_match(family, 5, "normal", 300) == [light]
    assert css_match([regular, medium, bold], 5, "normal", 300) == [regular]
    assert css_match([light, medium, bold], 5, "normal", 400) == [medium]
    assert css_match([light, bold], 5, "normal", 500) == [light]
    assert css_match([light, regular], 5, "normal", 600) == [regular]
    condensed = _Fake("condensed", 400, 3)
    assert css_match([regular, condensed], 3, "normal", 400) == [condensed]
    assert css_match([regular, condensed], 4, "normal", 400) == [condensed]
    italic = _Fake("italic", 400, style="italic")
    assert css_match([regular, italic], 5, "oblique", 700) == [italic]
    # Two faces that answer the same query: the SVG cannot say which (Avenir Book and Roman).
    assert len(css_match([regular, _Fake("book", 400)], 5, "normal", 400)) == 2


def test_a_superfamily_member_is_named_by_a_fallback_list():
    """Calibri Light's file is filed as "Calibri", weight 300: the SVG asks for
    ``Calibri Light, Calibri`` at 300, and the file is handed over in place."""
    fonts = _fonts()
    if fonts.drawing_face("Calibri Light") is None:
        pytest.skip("no Calibri Light here")
    from docx2svg.fonts import DrawingName, rasteriser_files

    assert fonts.drawing_name("Calibri Light") == DrawingName("Calibri", 300, "normal", "normal")
    assert fonts.drawing_name("Calibri") is None
    assert fonts.drawing_name("Calibri", True, True) is None
    files, missing, unaddressable = rasteriser_files(
        fonts, [("Calibri", False, False), ("Calibri Light", False, False), ("Calibri", True, False)])
    assert not missing and not unaddressable
    assert files == [fonts.drawing_face(*face).source for face in
                     [("Calibri", False, False), ("Calibri", True, False), ("Calibri Light", False, False)]]


def test_the_rasteriser_is_given_no_written_font_file(tmp_path):
    """Installed faces go over in place: nothing is written to the directory offered."""
    fonts = _fonts()
    from docx2svg.fonts import rasteriser_files

    faces = [("Calibri Light", False, False), ("Cambria", True, True), ("Times New Roman", False, True)]
    files, _missing, _ = rasteriser_files(fonts, faces, str(tmp_path))
    assert files and not any(Path(path).parent == tmp_path for path in files)
    assert not list(tmp_path.iterdir())


def test_src_writes_no_relabelled_face():
    """No module under ``src/`` relabels or rewrites a font file (it did until the fallback
    list replaced it)."""
    source = Path(__file__).parent.parent / "src" / "docx2svg"
    for path in source.rglob("*.py"):
        text = path.read_text(encoding="utf-8")
        assert "relabel(" not in text and "write_sfnt" not in text, path
