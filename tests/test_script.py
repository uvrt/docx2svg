"""Superscripts, subscripts, ``w:position`` and run borders, without Word.

``tests/fixtures/script-observations.json`` holds what Word drew for every document of
``tools/make_script_probe.py`` (recorded by ``tools/read_script_probe.py --record``): the
baselines, and the offset and size of every script run.  The documents are regenerated
from the generator, and every baseline goes through ``baselines.predict``, so this holds
the model to Word.  ROADMAP.md, "Superscripts, subscripts, position and run borders --
measured" (finding 5), has the tables.
"""

from __future__ import annotations

import functools
import json
import math
from dataclasses import replace
from fractions import Fraction
from pathlib import Path

import pytest

import baselines
import make_script_probe as probe
import probe_documents
import read_script_probe as reader

from docx2svg.resolve import script_half_points

DATA = json.loads((Path(__file__).parent / "fixtures" / "script-observations.json").read_text())


#: Word drew every line of the three compatibility settings alike (the recording keeps
#: modes 12 and 15 as references to the unstated setting, and the model reads the mode
#: only at page tops, where these paragraphs have no space before): one is scored.
#: Computed on first use, not when the module is collected: every pytest-xdist worker
#: collects every module, and only the one that runs these tests needs them (the module's
#: tests share one worker: tests/conftest.py, ``LAZY_MODULES``).
@functools.cache
def _scored():
    return reader.offline(DATA, lambda p: p.setting == "none")


METRICS = probe_documents.recorded_metrics(DATA["faces"])
RECORDED = list(probe.PROBES)


def test_the_recording_is_of_these_documents():
    assert set(DATA["documents"]) == {p.name for p in RECORDED}
    for p, results in _scored():
        assert len(results) == len(reader.paragraph_texts(p))
        assert all(r.status in ("exact", "miss") for r in results), p.name


@pytest.mark.parametrize("kind", ["script", "position", "border"])
def test_every_compatibility_setting_draws_alike(kind):
    """No ``settings.xml``, mode 12 and mode 15: the same baselines and script runs."""
    for setting in ("12", "15"):
        assert DATA["documents"][f"{kind}-{setting}"] == f"{kind}-none"
        assert DATA["scripts"][f"{kind}-{setting}"] == f"{kind}-none"


# -- w:vertAlign: the line --------------------------------------------------------------


def test_every_baseline_of_every_script_group():
    """Superscripts, subscripts and both, in the text's size, twice it and half it, in six
    faces at 11 and 20 pt under four line rules: 4,608 baselines per setting, all exact
    with the run counted **at its own ``w:sz``, on the baseline**."""
    (p, results), = [(p, r) for p, r in _scored() if p.name == "script-none"]
    table = reader.tabulate([(p, results)], lambda kind: "all")
    assert table["all"] == [4608, 4608]


def test_a_script_in_the_texts_size_leaves_the_line_as_it_is():
    """Refutes "a superscript or subscript makes its line ~3 px taller" (``sample1``): in
    the text's own size every such line is drawn where the control line is drawn."""
    compared = 0
    for name in DATA["documents"]:
        if not name.startswith("script-") or name == "script-sizes":
            continue
        pages = reader.recorded(DATA, "documents", name)
        p = next(p for p in RECORDED if p.name == name)
        ys = {group.name: pages[i] for i, group in enumerate(p.groups)}
        for group in p.groups:
            if group.kind in ("sup", "sub", "supsub") and group.extra_hp is None:
                control = replace(group, kind="plain")
                assert ys[group.name] == ys[control.name], group.name
                compared += len(ys[group.name])
    assert compared == 3 * 3 * 6 * 2 * 4 * probe.LINES


def _reduced(items):
    out = []
    for item in items:
        if item.vertical_align:
            face = METRICS(item.face, item.bold, item.italic)
            item = replace(item, half_points=script_half_points(item.half_points, item.vertical_align, face))
        out.append(item)
    return out


def test_a_script_counted_at_its_drawn_size_is_refuted(monkeypatch):
    """The line counting the script at the size it is drawn at: right where the script is
    no taller than the text, wrong for every script twice its size but under ``exact``."""
    line_extent, tallest_natural = baselines.line_extent, baselines.tallest_natural
    monkeypatch.setattr(baselines, "line_extent", lambda t, m, l=(): line_extent(_reduced(t), m, l))
    monkeypatch.setattr(baselines, "tallest_natural", lambda t, m: tallest_natural(_reduced(t), m))
    scored = reader.offline(DATA, lambda p: p.name == "script-none")
    table = reader.tabulate(scored, reader.family)
    assert sum(v[0] for v in table.values()) == 3744  # of 4,608; 11,232 of 13,824 in three settings
    assert table["sup, 2x"] == table["sub, 2x"] == [144, 576]


# -- w:vertAlign: the glyphs ----------------------------------------------------------------


def _sizes():
    """``[(face, w:sz, drawn superscript half points, drawn subscript half points)]``."""
    p = next(p for p in RECORDED if p.name == "script-sizes")
    groups = DATA["scripts"]["script-sizes"]
    out = []
    for face, half_points in p.sizes:
        runs = groups[f"size/{face.replace(' ', '')}{half_points}"]
        sizes = sorted({round(size, 2) for _, size in runs})
        # One size for both runs; the offsets (up and down) differ.
        assert len(sizes) == 1, (face, half_points, runs)
        out.append((face, half_points, sizes[0]))
    return out


#: The seven sizes the rule does not reproduce: (face, w:sz) -> drawn half points.
SIZE_EXCEPTIONS = {
    **{(face, 90): 59 for face in ("Calibri", "Times New Roman", "Arial", "Cambria", "Courier New", "Georgia")},
    ("Baskerville Old Face", 6): 4,
}


def test_a_script_is_drawn_at_the_faces_own_script_size():
    """``w:sz x OS/2.ySuperscriptYSize / unitsPerEm``, half up, at least 2 half points; 3/5
    for a face whose value is absurd.  Every size from 2 to 96 in twelve faces: 1,133 of
    1,140 (both runs are drawn at one size in every line)."""
    right, wrong = 0, {}
    for face, half_points, drawn in _sizes():
        predicted = script_half_points(half_points, "superscript", METRICS(face))
        assert predicted == script_half_points(half_points, "subscript", METRICS(face))
        if predicted == drawn:
            right += 1
        else:
            wrong[(face, half_points)] = drawn
    assert right == 1133
    assert wrong == SIZE_EXCEPTIONS


def test_two_thirds_is_refuted():
    """What ``sample1`` suggested (2/3 of 12 pt is 16 half points, and so is 0.65 of it)."""
    right = sum(max(2, math.floor(Fraction(half_points * 2, 3) + Fraction(1, 2))) == drawn
                for _, half_points, drawn in _sizes())
    assert right == 192


# -- w:vertAlign: the offset (recorded, not modelled) ------------------------------------------


def _offsets():
    """``{(face, w:sz): (superscript, subscript) offset in whole half points}``."""
    p = next(p for p in RECORDED if p.name == "script-sizes")
    out = {}
    for face, half_points in p.sizes:
        runs = DATA["scripts"]["script-sizes"][f"size/{face.replace(' ', '')}{half_points}"]
        up = [-dy for dy, _ in runs if dy < 0] or [0]
        down = [dy for dy, _ in runs if dy > 0] or [0]
        out[(face, half_points)] = tuple(_half_points(px) for px in (up[0], down[0]))
    return out


def _half_points(px: int) -> int | None:
    return next((k for k in range(200) if math.floor(Fraction(k * 25, 12) + Fraction(1, 2)) == px), None)


def test_a_script_moves_by_whole_half_points():
    """Every drawn offset, 2,280 of them, is ``round(k x 25/12)`` px for a whole ``k``:
    Word raises and lowers a script run by whole half points, as ``w:position`` does."""
    assert all(k is not None for pair in _offsets().values() for k in pair)


@pytest.mark.parametrize("face, exact", [
    ("Cambria", 95), ("Georgia", 94), ("Impact", 94), ("Aptos", 91),
    ("Calibri", 2), ("Times New Roman", 6), ("Arial", 5), ("Courier New", 5),
])
def test_the_superscript_offset_follows_os2_only_in_some_faces(face, exact):
    """``round(w:sz x OS/2.ySuperscriptYOffset / unitsPerEm)`` half points: nearly every
    size of 95 in faces whose offset is at most 0.39 em, almost none in faces whose offset
    is 0.42-0.48 em (they are raised by about a third of the size instead).  Not settled
    (ROADMAP.md); nothing in the line height depends on it."""
    offset = _os2_offset(face)
    got = sum(k_up == math.floor(Fraction(half_points) * offset + Fraction(1, 2))
              for (name, half_points), (k_up, _) in _offsets().items() if name == face)
    assert got == exact


def _os2_offset(face: str) -> Fraction:
    return Fraction(*OS2_SUPERSCRIPT_OFFSET[face])


#: ``(OS/2.ySuperscriptYOffset, unitsPerEm)`` of the installed faces, as recorded.
OS2_SUPERSCRIPT_OFFSET = {
    "Cambria": (503, 2048), "Georgia": (542, 2048), "Impact": (800, 2048), "Aptos": (717, 2048),
    "Calibri": (976, 2048), "Times New Roman": (928, 2048), "Arial": (977, 2048), "Courier New": (865, 2048),
}


# -- the size sweep's baselines ----------------------------------------------------------------


def test_the_size_sweep_baselines():
    """The size sweep's one-line paragraphs, every size in twelve faces, paginated by
    Word: exact but for Baskerville Old Face, whose lines Word draws taller than its
    ``hhea`` says (a face-metrics question, not this one's), and the Impact lines after
    it on the same page."""
    (p, results), = [(p, r) for p, r in _scored() if p.name == "script-sizes"]
    names = reader.kinds(p)
    misses = {names[r.block].split("/")[1].rstrip("0123456789") for r in results if r.status == "miss"}
    assert baselines.summary(results)[:2] == (1004, 1140)
    assert misses == {"BaskervilleOldFace", "Impact", "HelveticaNeue"}


# -- w:position and w:bdr ------------------------------------------------------------------


def _position_score():
    return reader.tabulate(reader.offline(DATA, lambda p: p.name == "position-none"), lambda k: "all")["all"]


def test_position_groups():
    """``w:position`` +-1..48 half points in six faces under four rules, and on runs half
    and twice the text's size: a raised run reaches that much higher and that much less
    far down (``vertical.item_extent``); 1,389 of 4,608 before the model moved it."""
    (p, results), = [(p, r) for p, r in _scored() if p.name == "position-none"]
    assert reader.tabulate([(p, results)], lambda kind: "all")["all"] == [4608, 4608]


def test_the_position_in_the_multiples_natural_height_is_refuted(monkeypatch):
    """An ``auto`` multiple's extra is over the item's natural height *without* its raise."""
    from docx2svg import vertical

    def with_position(items, metrics):
        return max(vertical.natural_height_px(face, item.half_points) + abs(vertical.twips_to_px(item.position * 10))
                   for item in items for face in [metrics(item.face, item.bold, item.italic)] if face)

    monkeypatch.setattr(baselines, "tallest_natural", with_position)
    assert _position_score() == [3710, 4608]


def test_a_raised_run_keeping_its_descent_is_refuted(monkeypatch):
    """Raising a run twice the text's size takes its descent up with it."""
    from docx2svg import vertical

    def ascent_only(item, face):
        above, below = vertical._face_extent(face, item.half_points)
        shift = vertical.twips_to_px(item.position * 10)
        return (above + shift, below) if shift > 0 else (above, below - shift)

    monkeypatch.setattr(vertical, "item_extent", ascent_only)
    assert _position_score() == [4468, 4608]


def _border_score(key=lambda kind: "all"):
    return reader.tabulate(reader.offline(DATA, lambda p: p.name == "border-none"), key)


def test_border_groups():
    """``w:bdr`` ``w:sz`` 2..96 and ``w:space`` 0..31 in six faces under four rules, and on
    runs half and twice the text's size: the run reaches its border's space and width
    further up and down (``vertical.item_extent``), and an ``auto`` multiple scales that
    too (``vertical.item_natural_px``); 1,234 of 5,184 before the model had it."""
    (p, results), = [(p, r) for p, r in _scored() if p.name == "border-none"]
    assert reader.tabulate([(p, results)], lambda kind: "all")["all"] == [5184, 5184]


def test_the_border_outside_the_multiples_natural_height_is_refuted(monkeypatch):
    from docx2svg import vertical

    monkeypatch.setattr(baselines, "tallest_natural", lambda items, metrics: max(
        vertical.natural_height_px(face, item.half_points)
        for item in items for face in [metrics(item.face, item.bold, item.italic)] if face))
    assert _border_score(lambda kind: kind.split("/")[2])["auto276"] == [101, 1152]


@pytest.mark.parametrize("width, exact", [
    ("rounded", 5181),  # as a paragraph border's: the three w:sz 27 lines at a tie
    ("half twip", 5013),
])
def test_the_run_border_width_is_truncated(monkeypatch, width, exact):
    from docx2svg import vertical

    rounded = vertical.border_px
    half = lambda sz, space: vertical.twips_to_px(space * 20) + vertical.twips_to_px(Fraction(sz * 20, 8))
    monkeypatch.setattr(vertical, "run_border_px", rounded if width == "rounded" else half)
    assert _border_score()["all"] == [exact, 5184]


def test_a_border_around_the_whole_line_is_refuted(monkeypatch):
    """A border on a run half the text's size grows the line only by what passes the
    text; drawn around the line's whole extent it would grow every such line."""
    from docx2svg import vertical

    line_extent = baselines.line_extent

    def around_the_line(items, metrics, labels=()):
        pads = [vertical.run_border_px(*item.border) for item in items if item.border]
        extent = line_extent([replace(item, border=None) for item in items], metrics, labels)
        if not pads or extent is None:
            return line_extent(items, metrics, labels)
        face, half_points, _ = extent
        em = Fraction(half_points, 2) * vertical.PX_PER_PT
        above = (face.ascent + face.line_gap) * em / face.units_per_em + max(pads)
        below = face.descent * em / face.units_per_em + max(pads)
        return vertical.FaceMetrics(1, above / em, below / em, 0), half_points, True

    monkeypatch.setattr(baselines, "line_extent", around_the_line)
    assert _border_score(reader.family)["bdr, 1/2x"] == [0, 288]
