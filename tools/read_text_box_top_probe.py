#!/usr/bin/env python3
"""Measure ``make_text_box_top_probe.py``: each swept text box's first glyph, Word's
against the model's.

Word's side is its PDF read by PyMuPDF exactly as ``read_text_box_probe.py`` reads it --
each character's origin, its pen position on the baseline, the convention of the model's
``Span.xs`` and ``Span.y`` -- in device px (300 per inch).  For each box the line starting
with its label ``T<n>`` (``V<n>``, its line after a break, and ``U<n>``, a second
paragraph) gives the first glyph's pen x and baseline.  Printed per sweep: how many boxes
agree within half a pixel on each axis, and with ``--all`` each box's ``dx`` / ``dy``
(Word's less the model's); a recorded-only sweep (turned boxes) prints both positions.

``--record`` writes ``tests/fixtures/text-box-top-observations.json``: Word's positions,
the labels the model missed at recording time and the face numbers the model asked for.

Usage::

    python tools/read_text_box_top_probe.py [--record] [--all]
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_text_box_top_probe as probe  # noqa: E402

OBSERVATIONS = HERE.parent / "tests" / "fixtures" / "text-box-top-observations.json"
LABEL = re.compile(r"^([TUV])(\d+)\b")
PX = 300 / 72


def word_positions(data: bytes, name: str) -> dict[str, list[float]]:
    """``T<n>`` / ``U<n>`` -> ``[page, pen x, baseline]`` of the line's first glyph."""
    import oracle
    import pymupdf

    out: dict[str, list[float]] = {}
    document = pymupdf.open(oracle.export(data, name=name))
    for number, page in enumerate(document):
        for block in page.get_text("rawdict")["blocks"]:
            for line in block.get("lines", []):
                chars = [c for span in line["spans"] for c in span["chars"]]
                text = "".join(c["c"] for c in chars).strip()
                match = LABEL.match(text)
                if match and match[0] not in out:
                    first = next(c for c in chars if not c["c"].isspace())
                    out[match[0]] = [number, round(first["origin"][0] * PX, 3), round(first["origin"][1] * PX, 3)]
    return out


def model_positions(data: bytes, fonts=None) -> dict[str, list[float]]:
    """The model's positions, measured with ``fonts`` (``render_record``'s recording or
    recorded faces; the installed faces when ``None``)."""
    import render_record
    from docx2svg import convert_docx_to_layout

    out: dict[str, list[float]] = {}
    layout = convert_docx_to_layout(data, render_record.options(fonts) if fonts is not None else None)
    for page in layout.pages:
        for drawing in page.floats:
            for line in drawing.lines:
                spans = [s for s in line.spans if "".join(s.chars).strip()]
                if not spans:
                    continue
                text = "".join("".join(s.chars) for s in line.spans).strip()
                match = LABEL.match(text)
                if match and match[0] not in out:
                    k = next(i for i, c in enumerate(spans[0].chars) if not c.isspace())
                    out[match[0]] = [page.number, round(float(spans[0].xs[k]), 3), float(spans[0].y)]
    return out


def misses(word: dict, model: dict, axis: int) -> list[str]:
    """The labels the model does not put within half a pixel of Word's on ``axis`` (1:
    pen x, 2: baseline); recorded-only boxes (turned) are not scored."""
    out = []
    for box in probe.BOXES:
        if box.sweep.recorded_only:
            continue
        for key in box.keys:
            w, m = word.get(key), model.get(key)
            if w is None or m is None or w[0] != m[0] or abs(w[axis] - m[axis]) > 0.5:
                out.append(key)
    return out


def report(word: dict, model: dict, *, everything: bool = False) -> list[str]:
    rows = []
    for sweep in probe.SWEEPS:
        members = [b for b in probe.BOXES if b.sweep is sweep]
        x_ok = y_ok = total = 0
        detail = []
        for box in members:
            for key in box.keys:
                w, m = word.get(key), model.get(key)
                total += 1
                if w is None or m is None:
                    detail.append(f"{key}: Word {w} model {m}")
                    continue
                if sweep.recorded_only:
                    detail.append(f"{key}: Word {w} model {m}")
                    continue
                dx, dy = w[1] - m[1], w[2] - m[2]
                x_ok += abs(dx) <= 0.5
                y_ok += abs(dy) <= 0.5
                detail.append(f"{key}:{dx:+.0f}/{dy:+.0f}")
        rows.append(f"  {sweep.note:36} x {x_ok:2}/{total:2}  y {y_ok:2}/{total:2}"
                    + (" (recorded)" if sweep.recorded_only else ""))
        if everything or sweep.recorded_only:
            rows.append("     " + " ".join(detail))
    return rows


def main(argv: list[str]) -> int:
    import oracle
    import render_record

    faces: dict = {}
    documents = {}
    try:
        for setting in probe.SETTINGS:
            data = probe.build(setting)
            word = word_positions(data, f"text-box-top-{setting}")
            model = model_positions(data, render_record.RecordingFonts(data, faces))
            x_miss, y_miss = misses(word, model, 1), misses(word, model, 2)
            documents[setting] = {"word": word, "x_misses": x_miss, "y_misses": y_miss}
            count = sum(len(b.keys) for b in probe.BOXES if not b.sweep.recorded_only)
            print(f"== text-box-top-{setting}: pen x {count - len(x_miss)} / {count}, "
                  f"baseline {count - len(y_miss)} / {count} within half a pixel")
            print("\n".join(report(word, model, everything="--all" in argv)))
    finally:
        oracle.recover()
    if "--record" in argv:
        render_record.dump(OBSERVATIONS, {
            "_about": ("For each document of tools/make_text_box_top_probe.py, the page, pen x and baseline "
                       "(device px, 300 per inch) of the first glyph of each swept text box's labelled lines as "
                       "Word 16.106 drew them, and the labels the model missed on each axis at recording time; "
                       "the numbers of every face the model asked for. Measurements only; regenerate with "
                       "tools/read_text_box_top_probe.py --record."),
            "faces": faces, "documents": documents})
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
