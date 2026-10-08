#!/usr/bin/env python3
"""Measure ``make_picture_run_probe.py``: the baseline of the line after each picture, which
says how tall the picture's line was, against the model's.

``--record`` writes Word's baselines to ``tests/fixtures/picture-run-observations.json``;
``tests/test_picture_run.py`` holds the model to them offline.

Usage::

    python tools/read_picture_run_probe.py [--record]
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_picture_run_probe as probe  # noqa: E402
import read_render  # noqa: E402
import render_record  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "picture-run-observations.json"


def model_after(layout) -> list[int]:
    """Per case (a page each), the baseline of its ``after`` line."""
    return [page.lines[-1].baseline for page in layout.pages]


def word_after(pdf: Path) -> list[int]:
    out: dict[int, int] = {}
    for obj in read_render.word_objects(pdf):
        if "after" in obj[4]:
            out[obj[0]] = round(obj[1])
    return [out[k] for k in sorted(out)]


def main(argv: list[str]) -> int:
    import oracle
    from docx2svg import _render

    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--record", action="store_true")
    args = parser.parse_args(argv[1:])
    faces: dict = {}
    recorded: dict = {}
    for setting in probe.SETTINGS:
        data = probe.build(setting)
        word = word_after(oracle.export(data, name=f"picture-run-{setting}"))
        _, layout, _, _ = _render(data, render_record.options(render_record.RecordingFonts(data, faces)))
        ours = model_after(layout)
        print(f"picture-run-{setting}: lines after a picture on Word's baseline "
              f"{sum(a == b for a, b in zip(word, ours))} / {len(word)}")
        recorded[f"picture-run-{setting}"] = word
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("Word 16.106's baseline (device px) of the line after each picture of "
                       "tools/make_picture_run_probe.py, a page a case (tools/read_picture_run_probe.py). "
                       "Measurements only; regenerate with --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
