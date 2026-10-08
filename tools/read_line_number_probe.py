#!/usr/bin/env python3
"""Score ``make_line_number_probe.py``: every glyph against Word's, line numbers included.

Records, with ``--record``, Word's text objects and every face number the
renderer asked for in ``tests/fixtures/line-number-observations.json``;
``tests/test_line_numbers.py`` holds the renderer to them offline, and to the scores of
the rules the probe refutes (:func:`refuted`).

Usage::

    python tools/read_line_number_probe.py [-v] [--record]
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

OBSERVATIONS = read_render.FIXTURES / "line-number-observations.json"


def documents() -> list[tuple[str, bytes]]:
    import make_line_number_probe

    return [(f"lnnum-{setting}", make_line_number_probe.build(setting)) for setting in make_line_number_probe.SETTINGS]


@contextlib.contextmanager
def _variant(old: str, new: str):
    """``docx2svg.layout._Placer._line_numbers`` with one rule changed (its source edited)."""
    import textwrap

    from docx2svg import layout

    source = textwrap.dedent(inspect.getsource(layout._Placer._line_numbers))
    assert old in source, old
    namespace: dict = {}
    exec(compile(source.replace(old, new), layout.__file__, "exec"), layout.__dict__, namespace)
    original = layout._Placer._line_numbers
    layout._Placer._line_numbers = namespace["_line_numbers"]
    try:
        yield
    finally:
        layout._Placer._line_numbers = original


def refuted() -> list[tuple[str, object]]:
    """``(what, context manager)`` for each rule the probe refutes."""
    return [
        ("before: no line numbers", lambda: _variant("if not any(section.line_numbering for section in sections):",
                                                     "if True:")),
        ("the count starts at w:start", lambda: _variant("count = start\n", "count = start - 1\n")),
        ("w:restart continuous restarts at the section", lambda: _variant(
            'restart == "newSection" and k != last_section', 'restart != "newPage" and k != last_section')),
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
        objects = read_render.word_objects(pdf)
        fonts = render_record.RecordingFonts(data, faces)
        result, _rects, warnings = read_render.score(name, data, fonts, objects, [])
        print(result.row(), warnings)
        if args.verbose:
            for problem in result.problems[:40]:
                print("   ", problem.replace("|", ""))
        for what, variant in refuted():
            with variant():
                other, _rects, _warnings = read_render.score(name, data, fonts, objects, [])
            print(f"   refuted, {what}: {other.matched} matched, baseline {other.y_agree}, extra {other.extra}, "
                  f"not drawn {other.not_drawn}")
        recorded[name] = {"objects": objects, "glyphs": read_story_probe.row(result), "warnings": warnings}
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for each document of tools/make_line_number_probe.py -- every text "
                       "object, as render-observations.json records them -- the renderer's glyph scores, and every "
                       "face number it asked for. Measurements only; "
                       "regenerate with tools/read_line_number_probe.py --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
