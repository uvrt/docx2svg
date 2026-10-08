"""Tables against what Word drew, offline (ROADMAP.md, "Tables -- measured").

``tests/fixtures/table-observations.json`` holds every text object and filled rectangle
of Word's PDF for each table probe document (``tools/make_table_*_probe.py``), and every
face number the renderer asked for (``tools/read_table_probes.py --record``).  Each
document is laid out and drawn here from the file and those numbers alone, and every
glyph of every cell must be where Word drew it: its pen x (a cell's text start exactly,
a right-aligned word's end with it), its baseline, its face and its size.
"""

from __future__ import annotations

import functools
import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))
sys.path.insert(0, str(REPO / "src"))

import read_render  # noqa: E402
import read_table_probes  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(read_table_probes.OBSERVATIONS.read_text(encoding="utf-8"))


#: Computed on first use, not when the module is collected: every pytest-xdist worker
#: collects every module, and only the one that runs these tests needs them (the module's
#: tests share one worker: tests/conftest.py, ``LAZY_MODULES``).
@functools.cache
def _documents() -> dict[str, bytes]:
    return dict(read_table_probes.documents())


#: What a probe is known to get wrong, and why: glyphs matched but off Word's position.
#: Mode 15's justified cells, known off before Phase 3.8, are exact since.
KNOWN: dict[str, int] = {}


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


@pytest.mark.parametrize("name", sorted(DATA["documents"]))
def test_every_cell_glyph_where_word_drew_it(name, fonts):
    recorded = DATA["documents"][name]
    result, _rects, warnings = read_render.score(name, _documents()[name], fonts, recorded["objects"], recorded["fills"])
    assert read_render.row(result) == recorded["glyphs"]
    assert warnings == recorded["warnings"]
    if name in KNOWN:
        assert result.matched - result.x_agree == KNOWN[name]
        return
    assert result.extra == 0 and result.not_drawn == 0 and result.matched == result.model
    assert result.matched == min(result.x_agree, result.y_agree, result.face_agree, result.size_agree), \
        result.problems[:5]


def test_the_geometry_probe_varies_what_stage_one_measures():
    """The cases the roadmap's stage 1 rules rest on are all in the probe."""
    import make_table_geometry_probe as probe

    names = [name for name, _ in probe.CASES]
    for needed in ("margins none", "border 27", "indent -200", "indent by style", "jc center, border 24",
                   "jc right, indent 333", "widths wider than the column", "tcW 1000/3000/2000 over the grid",
                   "autofit, tcW 1000/3000/2000 over the grid"):
        assert needed in names


def test_a_cell_starts_on_a_whole_pixel_held_in_layout_units():
    """740 px is 727,450 units: Word's PDF prints every cell's text start as 740.0004."""
    from fractions import Fraction

    from docx2svg import parse_package, table
    from docx2svg.linebreak import column_width_twips
    from docx2svg.model import Table
    from docx2svg.vertical import LAYOUT_UNIT_PX

    document = parse_package(_documents()["table-geometry-15"])
    section = document.sections[0]
    first = next(item for item in document.body if isinstance(item, Table))
    resolved = table.resolve_table(document, first)
    lines = table.grid_lines(document, resolved, column_width_twips(section))
    box = table.cell_box(document, resolved, section.margins.left, lines, resolved.rows[0].cells[1])
    assert (box.text_left / LAYOUT_UNIT_PX).denominator == 1
    assert abs(box.text_left - round(box.text_left)) < Fraction(1, 1000)


def test_an_autofit_table_word_resizes_is_laid_out_and_what_is_not_settled_stops():
    """Stage 7 recorded, and stage 7b models (``make_autofit_width_probe.py``), what
    ``make_table_autofit_probe.py`` shows Word resizing: a cell with no width, a table width
    in percent over cells with none and a word wider than its column are laid out now; a
    cell across columns wider than they are still stops the layout, and says why."""
    import make_table_autofit_probe as probe

    from docx2svg import ConvertOptions, _lay_out

    def stop(body: str) -> str:
        data = probe.probe_docx.package(body + probe.wml.paragraph(probe.wml.run("after")), styles=probe.styles())
        layout, _, _ = _lay_out(data, ConvertOptions())
        return " ".join(w[1] for w in layout.warnings if w[0] == "layout-stopped:table")

    for number, (_name, first, options) in enumerate(probe.CASES):
        assert stop(probe._table(number, first, **options)) == ""
    import make_autofit_width_probe as width_probe

    spanned = next(case for case in width_probe.CASES if case.family == "span")
    assert "a cell across columns wider than they are" in stop(spanned.table.xml())
