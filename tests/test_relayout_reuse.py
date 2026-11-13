"""A layout made again from a previous one reuses what did not change (round 57, lane R):
a table that only moved (a block inserted or deleted before it) is not measured again, its
paragraphs' block numbers moved with it (:func:`docx2svg.table.rebased`); a line of an
unchanged paragraph or cell is not drawn again but taken from the layout that drew it,
moved to where its pitch starts now (:meth:`docx2svg.layout._Placer._body_line`,
``_cell_line``); and the pages kept have their paths renumbered without walking every
glyph (:func:`docx2svg.relayout.renumbered`).

Each is held to the relayout property (``tests/test_relayout.py``): the layout is the one a
full layout of the new bytes makes, every page's SVG byte for byte.  Every document is made
up (``tools/report_docs.py``).
"""

from __future__ import annotations

import dataclasses
import pickle
import random
import re
from dataclasses import replace

import pytest

import report_docs
from docx2svg import ConvertOptions, lay_out_docx, page_svgs
from docx2svg import layout as layout_module
from docx2svg import relayout
from docx2svg import table as table_model

_DATA = re.compile(r' data-docx-[a-z-]+="[^"]*"')


def _full_svgs(data: bytes) -> list[str]:
    layout, data, fonts = lay_out_docx(data, ConvertOptions(), keep=False)
    return page_svgs(layout, data, fonts=fonts)


def _again(data: bytes, previous):
    layout, data, fonts = lay_out_docx(data, ConvertOptions(), previous=previous)
    return layout, page_svgs(layout, data, fonts=fonts)


@pytest.fixture(scope="module")
def report16():
    blocks = report_docs.report(16, seed=5)
    data = report_docs.package(blocks)
    layout, _, _ = lay_out_docx(data)
    return blocks, layout


def _first_para(blocks, near: int = 0) -> int:
    return next(k for k in range(near, len(blocks)) if blocks[k].kind == "para")


def test_a_table_that_only_moved_is_not_measured_again(report16, monkeypatch):
    blocks, layout = report16
    edited = list(blocks)
    edited.insert(_first_para(blocks, 2), report_docs.Block("para", "A made-up paragraph, inserted early.",
                                                            para_id="7000BBBB"))
    data = report_docs.package(edited)
    measured: list = []
    flow_table = table_model.flow_table

    def counting(*args, **kwargs):
        measured.append(kwargs.get("block"))
        return flow_table(*args, **kwargs)

    monkeypatch.setattr(table_model, "flow_table", counting)
    again, svgs = _again(data, layout)
    monkeypatch.undo()
    assert again.relaid is None
    assert sum(1 for b in edited if b.kind == "table") > 3
    # Before round 57 every table after the insertion was measured again (its key held its
    # block number); none changed, so none is.
    assert measured == []
    assert svgs == _full_svgs(data)


def test_a_moved_table_is_its_flow_at_the_new_place(report16):
    from docx2svg import _fonts, parse_package
    from docx2svg.paginate import flow

    blocks, _ = report16
    data = report_docs.package(blocks)
    document = parse_package(data)
    fonts = _fonts(data, ConvertOptions())
    items = flow(document, fonts, fonts.metrics)
    tables = [item for item in items if hasattr(item, "flow") and item.flow is not None and item.block > 5]
    assert tables
    item = tables[0]
    moved = table_model.rebased(item.flow, 7)
    # Every paragraph's block number moved, nested and stacked alike; nothing measured changed.
    olds = [p for row in item.flow.rows for cell in row.cells for p in cell.paragraphs]
    news = [p for row in moved.rows for cell in row.cells for p in cell.paragraphs]
    assert [p.block + 7 for p in olds] == [p.block for p in news]
    assert all(a.pieces is b.pieces and a.lines is b.lines for a, b in zip(olds, news))
    for row in moved.rows:
        for cell in row.cells:
            mine = {id(p) for p in cell.paragraphs}
            assert all(id(line.paragraph) in mine for line in cell.stack)
    # The flow it was moved from is as it was.
    assert [p.block for p in olds] == [p.block for row in item.flow.rows for cell in row.cells
                                       for p in cell.paragraphs]
    assert moved.overflowed == set() and moved.rows is not item.flow.rows


def test_lines_that_moved_are_taken_not_drawn_again(report16, monkeypatch):
    blocks, layout = report16
    edited = list(blocks)
    k = _first_para(blocks, 4)
    edited[k] = replace(edited[k], text=edited[k].text + " " + " ".join(["More made-up words here."] * 12))
    data = report_docs.package(edited)
    drawn: list = []
    line = layout_module._Placer._line

    def counting(self, *args, **kwargs):
        drawn.append(1)
        return line(self, *args, **kwargs)

    monkeypatch.setattr(layout_module._Placer, "_line", counting)
    again, svgs = _again(data, layout)
    monkeypatch.undo()
    assert again.relaid is None
    fresh = sorted(again.state.fresh)
    on_fresh = sum(len(again.pages[p].lines) for p in fresh)
    assert len(fresh) > 3, fresh
    # Most of the lines on the pages laid out again are lines of paragraphs and cells that
    # did not change, only moved: taken from the previous layout.
    assert len(drawn) < on_fresh // 2, (len(drawn), on_fresh)
    assert svgs == _full_svgs(data)


def test_lines_are_taken_over_a_chain_of_layouts_and_not_from_a_pickle():
    rng = random.Random(57)
    blocks = report_docs.report(8, seed=57, toc=True)
    layout, _, _ = lay_out_docx(report_docs.package(blocks))
    assert layout.state.lines
    for _ in range(4):
        blocks, what = report_docs.edit(blocks, rng, kinds=("grow", "insert", "cell", "row-insert"))
        data = report_docs.package(blocks)
        layout, svgs = _again(data, layout)
        assert svgs == _full_svgs(data), what
        assert layout.state.lines
    # Kept in its process only: a pickled layout's state has none, and a layout from it draws
    # every line it lays out, the same.
    thawed = pickle.loads(pickle.dumps(layout))
    assert thawed.state.lines is None
    blocks, what = report_docs.edit(blocks, rng, kind="grow")
    data = report_docs.package(blocks)
    again, svgs = _again(data, thawed)
    assert svgs == _full_svgs(data), what


def _renumbered_by_walking(value, moves: dict, seen: dict | None = None):
    # Round 19's walk, every value of every field: what the faster one must equal.
    if not moves:
        return value
    seen = {} if seen is None else seen
    key = id(value)
    if key in seen:
        return seen[key]

    def step(match):
        name, index = match.group(1), int(match.group(2))
        move = moves.get(name)
        if move is None or index < move[0]:
            return match.group(0)
        return f"w:body/w:{name}[{index + move[1]}]"

    if isinstance(value, str):
        out = relayout._STEP.sub(step, value) if "w:body/" in value else value
    elif isinstance(value, list):
        items = [_renumbered_by_walking(item, moves, seen) for item in value]
        out = items if any(a is not b for a, b in zip(items, value)) else value
    elif isinstance(value, tuple):
        members = tuple(_renumbered_by_walking(item, moves, seen) for item in value)
        out = members if any(a is not b for a, b in zip(members, value)) else value
    elif isinstance(value, dict):
        pairs = {_renumbered_by_walking(k, moves, seen): _renumbered_by_walking(v, moves, seen)
                 for k, v in value.items()}
        out = pairs if any(a is not b or pairs[a] is not value[b] for a, b in zip(pairs, value)) else value
    elif dataclasses.is_dataclass(value) and not isinstance(value, type):
        changes = {}
        for item in dataclasses.fields(value):
            current = getattr(value, item.name)
            moved = _renumbered_by_walking(current, moves, seen)
            if moved is not current:
                changes[item.name] = moved
        if changes:
            import copy

            out = copy.copy(value)
            for name, moved in changes.items():
                object.__setattr__(out, name, moved)
        else:
            out = value
    else:
        out = value
    seen[key] = out
    return out


def test_renumbering_a_page_is_the_walk_of_all_it_holds(report16):
    _, layout = report16
    moves = {"p": (3, 2), "tbl": (1, 1)}
    for page in layout.pages:
        fast = relayout.renumbered(page, moves)
        slow = _renumbered_by_walking(page, moves)
        assert fast == slow
        # Copied only where a path changed (the walk copied wherever a path matched).
        assert fast is page or slow is not page
    # Nothing to move: the page itself.
    assert relayout.renumbered(layout.pages[0], {}) is layout.pages[0]


def test_a_body_scanned_from_the_previous_version_is_the_scan_of_the_whole():
    from docx2svg.scope import kept_features, paragraph_features

    rng = random.Random(41)
    blocks = report_docs.report(8, seed=41, toc=True, notes=True, fields=True)
    data = report_docs.package(blocks)
    _, main = relayout.part_digests(data)
    scan = relayout.scan_body(main)
    segments = relayout.segments(main)
    features, by = kept_features(data, scan, [s.digest for s in segments], None)
    assert features == paragraph_features(data)
    kinds = report_docs.EDIT_KINDS + ("caption-insert", "caption-delete", "toc-insert", "toc-delete")
    for _ in range(24):
        blocks, what = report_docs.edit(blocks, rng, kinds=kinds)
        data = report_docs.package(blocks)
        _, main = relayout.part_digests(data)
        again = relayout.scan_body(main, scan)
        assert again is not None and again.spans == relayout.scan_body(main).spans, what
        scan = again
        segments = relayout.segments(main)
        # What the flow leaves out, read from the elements that changed only: all of it.
        features, by = kept_features(data, scan, [s.digest for s in segments], by)
        assert features == paragraph_features(data), what
    # A body the scan cannot read (a comment in it) is not read from the previous one either.
    commented = main.replace(b"<w:body>", b"<w:body><!-- made up -->", 1)
    assert relayout.scan_body(commented, scan) is None and relayout.scan_body(commented) is None
