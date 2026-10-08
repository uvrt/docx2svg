#!/usr/bin/env python3
"""Score ``make_table_foot_probe.py``: every glyph and every filled rectangle against Word's,
where table rows meet a page's foot -- a row holding a cell merged down, a header row, a
paragraph kept whole.

Records, with ``--record``, Word's text objects and fills and every face number the
renderer asked for in ``tests/fixtures/table-foot-observations.json``;
``tests/test_table_foot.py`` holds the renderer to them offline, and to the scores of the
rules the probe refutes (:func:`refuted`).

Usage::

    python tools/read_table_foot_probe.py [-v] [--record]
"""

from __future__ import annotations

import argparse
import contextlib
import inspect
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import read_render  # noqa: E402
import read_story_probe  # noqa: E402
import render_record  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "table-foot-observations.json"


def documents() -> list[tuple[str, bytes]]:
    import make_table_foot_probe

    return [(f"foot-{setting}", make_table_foot_probe.build(setting)) for setting in make_table_foot_probe.SETTINGS]


@contextlib.contextmanager
def _variant(old: str, new: str):
    """``docx2svg.table._place_table`` with one rule taken out (its source edited)."""
    from docx2svg import table

    source = inspect.getsource(table._place_table)
    assert old in source, old
    namespace: dict = {}
    exec(compile(source.replace(old, new), table.__file__, "exec"), table.__dict__, namespace)
    original, table._place_table = table._place_table, namespace["_place_table"]
    try:
        yield
    finally:
        table._place_table = original


def refuted() -> list[tuple[str, object]]:
    """``(what, context manager)`` for each rule the probe refutes."""
    return [
        ("a header row split like any row", lambda: _variant("header = row < header_count and not continued",
                                                             "header = False")),
        ("mode 15: header rows left at the foot with no row after them",
         lambda: _variant("if mode15 and header_count and pieces", "if False and header_count and pieces")),
    ]


def main(argv: list[str]) -> int:
    import oracle

    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv[1:])
    faces: dict = {}
    recorded: dict = {}
    for name, data in documents():
        pdf = oracle.export(data, name=name)
        objects, fills = read_render.word_objects(pdf), read_render.word_fills(pdf)
        fonts = render_record.RecordingFonts(data, faces)
        result, rects, warnings = read_render.score(name, data, fonts, objects, fills)
        print(result.row())
        print(f"{'':34} rectangles: model {rects['all'][0]} px, Word {rects['all'][1]} px, model only "
              f"{rects['all'][2]}, Word only {rects['all'][3]}; warnings {warnings}")
        if args.verbose:
            for problem in result.problems[:40]:
                print("   ", problem.replace("|", ""))
        for what, variant in refuted():
            with variant():
                other, _rects, _warnings = read_render.score(name, data, fonts, objects, [])
            print(f"   refuted, {what}: {other.matched} matched, baseline {other.y_agree}, extra {other.extra}, "
                  f"not drawn {other.not_drawn}")
        recorded[name] = {"objects": objects, "fills": fills, "glyphs": read_story_probe.row(result),
                          "rects": rects["all"], "warnings": warnings}
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for each document of tools/make_table_foot_probe.py -- every text "
                       "object and filled rectangle, as render-observations.json records them -- the renderer's "
                       "glyph and rectangle scores, and every face number it asked for. Measurements only; "
                       "regenerate with tools/read_table_foot_probe.py --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
