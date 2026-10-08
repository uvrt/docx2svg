"""What a list label adds to the height of its line, without Word.

``tests/fixtures/label-observations.json`` holds the lines Word drew for every document
of ``tools/make_label_probe.py`` (recorded by ``tools/read_label_probe.py --record``);
the documents are regenerated from the generator.  Every baseline goes through
``baselines.predict``, so this holds ``docx2svg.vertical.line_extent`` -- a label takes
part by its ascent, not its descent -- to Word.  ROADMAP.md, "A list label's descent --
measured", has the table.
"""

from __future__ import annotations

import functools
import json
from pathlib import Path

import pytest

import baselines
import make_label_probe as probe
import read_label_probe as reader

from docx2svg import vertical

DATA = json.loads((Path(__file__).parent / "fixtures" / "label-observations.json").read_text())


#: Computed on first use, not when the module is collected: every pytest-xdist worker
#: collects every module, and only the one that runs these tests needs them (the module's
#: tests share one worker: tests/conftest.py, ``LAZY_MODULES``).
@functools.cache
def _scored():
    return reader.offline(DATA)


def test_the_recording_is_of_these_documents():
    assert set(DATA["documents"]) == {p.name for p in probe.PROBES}
    for p, results in _scored():
        assert all(r.status in ("exact", "miss") for r in results), p.name


@pytest.mark.parametrize("setting", sorted(probe.SETTINGS))
def test_every_baseline_of_every_label_group(setting):
    """Six labels and two inline controls before four text faces, single and ``auto``
    276: 1,408 baselines per setting, all exact."""
    table = reader.tabulate(_scored())
    groups = {group: tuple(v) for (s, group), v in table.items() if s == setting}
    assert len(groups) == len(probe.GROUPS)
    assert all(exact == count == probe.LINES for exact, count in groups.values())


def test_a_label_counting_with_its_descent_is_refuted(monkeypatch):
    """The model before the probe: a label's descent counted like text.  Right only where
    the text's descent is at least the label's -- 783 of the 3,168 label lines -- and
    for the inline controls, which are text."""
    original = vertical.line_extent

    def with_descent(text_items, metrics, label_items=()):
        return original(list(text_items) + list(label_items), metrics)

    monkeypatch.setattr(baselines, "line_extent", with_descent)
    by_extra = reader.by_extra(reader.tabulate(reader.offline(DATA)))
    labels = [by_extra[name] for name in probe.LABELS]
    assert [sum(v[0] for v in labels), sum(v[1] for v in labels)] == [783, 3168]
    assert [by_extra[name] for name in probe.INLINE] == [[528, 528], [528, 528]]


def test_a_label_taller_than_the_text_lends_its_ascent_but_not_its_descent():
    """Courier New 20 pt before Tahoma 11 pt: the label's ascent over Tahoma's descent."""
    tahoma = vertical.FaceMetrics(*DATA["faces"]["Tahoma|0|0"])
    courier = vertical.FaceMetrics(*DATA["faces"]["Courier New|0|0"])

    class Item:
        def __init__(self, face, half_points):
            self.face, self.half_points, self.bold, self.italic = face, half_points, False, False

    faces = {"Tahoma": tahoma, "Courier New": courier}
    face, half_points, mixed = vertical.line_extent([Item("Tahoma", 22)], lambda f, b, i: faces[f],
                                                   [Item("Courier New", 40)])
    assert mixed and half_points == 40
    em = vertical.PX_PER_PT / 2
    assert face.ascent * half_points * em == courier.ascent * 40 * em / courier.units_per_em
    assert face.descent * half_points * em == tahoma.descent * 22 * em / tahoma.units_per_em
