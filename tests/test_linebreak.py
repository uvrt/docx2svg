"""The line breaker (``docx2svg.linebreak``) against Word, without Word.

Every probe is regenerated from its generator, byte for byte; Word's lines of it are in
``tests/fixtures/*-observations.json`` and the advances every probe is built from in
``tests/fixtures/probe-advances.json`` (``tools/probe_advances.py``); the real
documents' lines are ``baseline-observations.json``'s and their advances
``break-advances.json`` (``tools/read_breaks.py --record``).  Each probe is scored through
``tools/breaks.py``, which runs the model's own breaker and compares its lines with
Word's: a line agrees when it starts and ends where Word's does.  ROADMAP.md, "Phase 3 --
measured", has the tables.
"""

from __future__ import annotations

import json
import sys
from fractions import Fraction
from pathlib import Path

import pytest

FIXTURES = Path(__file__).resolve().parent / "fixtures"
sys.path.insert(0, str(FIXTURES.parent.parent / "tools"))

import baselines  # noqa: E402
import breaks  # noqa: E402
import face_advances  # noqa: E402
import make_break_rules_probe as rules_probe  # noqa: E402
import make_wrap_probe as wrap_probe  # noqa: E402
import read_break_rules_probe  # noqa: E402
import read_wrap_budget_probe  # noqa: E402
import read_wrap_probe  # noqa: E402

from docx2svg import linebreak, parse_package  # noqa: E402
from docx2svg.vertical import FaceMetrics  # noqa: E402


def _totals(results) -> tuple[int, int]:
    scored = [r for r in results if r.status == "scored"]
    return sum(r.agree for r in scored), sum(r.lines for r in scored)


# -- units --------------------------------------------------------------------------------


def test_a_twip_is_204_8_units_rounded_half_up():
    assert linebreak.twips_to_units(5) == 1024
    assert linebreak.twips_to_units(1) == 205  # 204.8
    assert linebreak.twips_to_units(7) == 1434  # 1433.6: tracking's measured value (Phase 2)
    assert linebreak.twips_to_units(-5) == -1024


def test_default_stops_continue_past_the_last_custom_stop():
    geometry = linebreak.Geometry(0, 0, 10**7, tabs=(linebreak.TabStop(linebreak.twips_to_units(3000)),),
                                  default_tab=linebreak.twips_to_units(708))
    assert linebreak.next_tab(100, geometry, True)[:2] == (linebreak.twips_to_units(3000), "left")
    stop, _, default = linebreak.next_tab(linebreak.twips_to_units(3000), geometry, True)
    assert default and stop == 5 * linebreak.twips_to_units(708)


# -- the δ sweep --------------------------------------------------------------------------


def test_the_budget_is_exact_inclusive_and_without_slack():
    """3,120 cases, δ swept through zero in steps of one unit (1/4096 pt), eight faces,
    seven families, four compatibility settings: the line holds its last word exactly
    when it ends at or before the edge.  The strict budget (``x < right``) fails every
    δ = 0 case: 2,919."""
    scored = read_wrap_budget_probe.offline()
    cases = [(r.status == "scored" and r.agree == r.lines) for results in scored.values() for r in results]
    assert (sum(cases), len(cases)) == (3120, 3120)
    assert _totals([r for results in scored.values() for r in results]) == (7101, 7101)
    strict = read_wrap_budget_probe.offline(right_shift=-1)
    cases = [(r.status == "scored" and r.agree == r.lines) for results in strict.values() for r in results]
    assert sum(cases) == 2919


# -- the wrapping probe ---------------------------------------------------------------------


def test_every_line_of_the_wrapping_probe_breaks_after_the_same_word():
    """Ten faces, five sizes, five columns, four styles: 1,778 lines; and 257 more
    justified, centred and right-aligned."""
    table = read_wrap_probe.tabulate(read_wrap_probe.offline())
    assert table[("all", "left")] == [1778, 1778, 1778]
    assert table[("all", "aligned")] == [257, 257, 257]
    assert not [key for key in table if key[0] == "out"]
    for face in wrap_probe.FACES:
        agree, lines, _ = table[("face", face)]
        assert agree == lines, face
    assert {size for (dimension, size) in table if dimension == "size"} == set(wrap_probe.SIZES)
    assert {column for (dimension, column) in table if dimension == "column"} == set(wrap_probe.COLUMNS)


def test_without_the_dashes_twenty_lines_break_elsewhere(monkeypatch):
    """Refuted: a line breaks after a hyphen-minus only (not after an en or em dash)."""
    monkeypatch.setattr(linebreak, "BREAK_AFTER", frozenset("-"))
    table = read_wrap_probe.tabulate(read_wrap_probe.offline())
    assert table[("all", "left")][:2] == [1758, 1778]


# -- the rules probe ------------------------------------------------------------------------

#: Lines per family, every one agreeing.
RULE_FAMILIES = {
    "after": 318, "before": 314, "hang": 108, "shy": 40, "nbh": 12, "tab": 30, "br": 15,
    "kern": 54, "run": 10, "long": 14, "format": 4, "label": 1,
}


def _rules(advances=None) -> dict:
    out: dict = {}
    for name, results in read_break_rules_probe.offline(advances).items():
        for case, r in zip(rules_probe.documents()[name], results):
            cell = out.setdefault(case.family, [0, 0])
            assert r.status == "scored", (case.name, r.status)
            cell[0] += r.agree
            cell[1] += r.lines
    return out


def test_every_rule_case_breaks_where_word_does():
    assert _rules() == {family: [lines, lines] for family, lines in RULE_FAMILIES.items()}


def test_kerning_is_the_legacy_table_not_the_opentype_feature():
    """Aptos kerns ``Bo`` in its OpenType feature and not in its legacy ``kern`` table;
    Word breaks as the legacy table says."""
    import probe_advances

    found = _rules(probe_advances.Advances(gpos=True))
    assert found["kern"] == [52, 54]


def test_a_letter_is_charged_its_pair_with_the_space_after_it(monkeypatch):
    """Refuted: charging a kern pair to its right glyph, so that the last letter of a line
    is never charged its pair with the space after it (Times New Roman ``A `` and ``Y ``)."""
    original = linebreak._kern

    def right_only(piece, following, advances):
        return 0 if following.kind == linebreak.SPACE else original(piece, following, advances)

    monkeypatch.setattr(linebreak, "_kern", right_only)
    assert _rules()["kern"] == [50, 54]


# -- real documents -------------------------------------------------------------------------

DATA = json.loads((FIXTURES / "baseline-observations.json").read_text(encoding="utf-8"))
ADVANCES = json.loads((FIXTURES / "break-advances.json").read_text(encoding="utf-8"))

#: (lines that agree, lines scored) per document; every one.
DOCUMENTS = {
    "layout-sweep.docx": (91, 91),
    "style-document.docx": (34, 34),
    "sample-blank.docx": (0, 0),
    "sample-long.docx": (470, 470),
    "sample-resume.docx": (20, 20),
    "sample-simple.docx": (24, 24),
    "sample-10pages.docx": (76, 76),
    "sample-1page.docx": (7, 7),
    "sample-5pages.docx": (47, 47),
    "sample-with-images.docx": (16, 16),
    "sample-with-table.docx": (9, 9),
}


def _path(name: str) -> Path:
    for directory in (FIXTURES, FIXTURES / "samplelib", FIXTURES / "wordto"):
        if (directory / name).exists():
            return directory / name
    raise FileNotFoundError(name)


def _metrics(face, bold=False, italic=False):
    found = DATA["faces"].get(f"{face}|{int(bold)}|{int(italic)}")
    return None if found is None else FaceMetrics(*found)


@pytest.mark.parametrize("name", sorted(DOCUMENTS))
def test_real_documents_break_where_word_breaks_them(name):
    data = _path(name).read_bytes()
    drawn = [baselines.DrawnLine(*line) for line in DATA["documents"][name]]
    results = breaks.score(parse_package(data), drawn, face_advances.RecordedAdvances(ADVANCES["faces"]),
                           _metrics, package=data)
    assert _totals(results) == DOCUMENTS[name]


def test_the_block_d_paragraph_of_the_fixture():
    """Phase 3's done-condition names it: the layout sweep's wrapping paragraph."""
    data = _path("layout-sweep.docx").read_bytes()
    drawn = [baselines.DrawnLine(*line) for line in DATA["documents"]["layout-sweep.docx"]]
    results = breaks.score(parse_package(data), drawn, face_advances.RecordedAdvances(ADVANCES["faces"]),
                           _metrics, package=data)
    block_d = [r for r in results if r.text.startswith("Line breaking is")]
    assert len(block_d) == 1 and block_d[0].lines == 4 and block_d[0].agree == 4


@pytest.mark.parametrize("name", ["sample-long.docx", "sample-10pages.docx", "style-document.docx"])
def test_baselines_are_unchanged_when_the_model_decides_each_lines_characters(name):
    """``baselines.predict`` with advances takes which characters each line holds from the
    model instead of from Word's lines; every pinned score stays."""
    import test_baselines

    data = _path(name).read_bytes()
    document = parse_package(data)
    drawn = [baselines.DrawnLine(*line) for line in DATA["documents"][name]]
    results = baselines.predict(document, drawn, _metrics, package=data,
                                advances=face_advances.RecordedAdvances(ADVANCES["faces"]))
    assert baselines.summary(results) == test_baselines.EXPECTED[name]
    assert not [r for r in results if r.status in ("exact", "miss") and any("model breaks" in n for n in r.notes)]


def test_the_ooxml_common_tables_agree_where_they_are_trusted():
    """``measure.TableAdvances`` answers from ``ooxml-common``'s tables for six faces; for
    every character a real document here asks, it answers what Word's face does."""
    from docx2svg.measure import TABLES, TableAdvances

    tables = TableAdvances()
    compared = 0
    for key, entry in ADVANCES["faces"].items():
        face, bold, italic = key.split("|")
        if face.lower() not in TABLES or italic == "1":
            continue
        for char, advance in entry["advances"].items():
            found = tables.advance(face, bold == "1", False, char)
            if found is None:
                continue
            assert found == (advance, entry["upm"]), (face, bold, char)
            compared += 1
    assert compared > 100


def test_src_imports_the_standard_library_and_ooxml_common_only():
    """``src/`` is standard-library only at runtime; its one dependency, ooxml-common,
    is too (``pyproject.toml``).  The one exception is ``png.py``, which imports the
    ``png`` extra's rasteriser -- inside the function that uses it, so importing
    ``docx2svg`` never needs it."""
    import ast

    allowed = set(sys.stdlib_module_names) | {"docx2svg", "ooxml_common"}
    extras = {"png.py": {"resvg_py", "cairosvg"}}
    source = FIXTURES.parent.parent / "src" / "docx2svg"
    for path in source.rglob("*.py"):
        permitted = allowed | extras.get(path.name, set())
        for node in ast.walk(ast.parse(path.read_text(encoding="utf-8"))):
            if isinstance(node, ast.Import):
                names = [alias.name for alias in node.names]
            elif isinstance(node, ast.ImportFrom) and node.level == 0:
                names = [node.module]
            else:
                continue
            for name in names:
                assert name.split(".")[0] in permitted, f"{path.name} imports {name}"
