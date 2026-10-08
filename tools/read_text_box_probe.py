#!/usr/bin/env python3
"""Measure ``make_text_box_probe.py``: each text box's first glyphs, Word's against the model's.

Word's side is its PDF read by PyMuPDF: every character's origin -- its pen position on
the baseline, the convention ``docx2svg``'s spans use (``Span.xs`` and ``Span.y``) -- in
device px (300 per inch).  The instrument was held to the body's lines of real documents
first: every one within 0.01 px of the model's.  For each box, the line starting with its
label ``B<n>`` (and ``C<n>``, its second paragraph) gives the first glyph's pen x and
baseline; the model's layout gives its own.  Printed per case: ``dx`` and ``dy``, Word's
less the model's.

``--record`` writes ``tests/fixtures/text-box-observations.json``: Word's positions only.

Usage::

    python tools/read_text_box_probe.py [--record] [--all]
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_text_box_probe as probe  # noqa: E402

OBSERVATIONS = HERE.parent / "tests" / "fixtures" / "text-box-observations.json"
LABEL = re.compile(r"^([BC])(\d+)\b")
PX = 300 / 72


def word_positions(data: bytes, name: str) -> dict[str, list[float]]:
    """``B<n>`` / ``C<n>`` -> ``[page, pen x, baseline]`` of the line's first glyph."""
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


def misses(word: dict, model: dict) -> list[str]:
    """The labels the model does not put within half a pixel of Word's on both axes."""
    out = []
    for number, box in enumerate(probe.CASES):
        if box.recorded_only:
            continue
        for key in (f"B{number}", f"C{number}") if box.second else (f"B{number}",):
            w, m = word.get(key), model.get(key)
            if w is None or m is None or abs(w[1] - m[1]) > 0.5 or abs(w[2] - m[2]) > 0.5:
                out.append(key)
    return out


def compare(word: dict, model: dict, *, everything: bool = False) -> tuple[int, int, list[str]]:
    """``(exact, compared, report lines)``: a label is exact when its glyph is within half
    a pixel of Word's on both axes (the model draws on whole-pixel baselines)."""
    exact = compared = 0
    rows = []
    for number, box in enumerate(probe.CASES):
        for prefix in ("B", "C") if box.second else ("B",):
            key = f"{prefix}{number}"
            w, m = word.get(key), model.get(key)
            if box.recorded_only:
                rows.append(f"  {key:5} {box.family:9} {box.note:34} Word {w} (recorded)")
                continue
            if w is None or m is None:
                rows.append(f"  {key:5} {box.family:9} {box.note:34} Word {w} model {m}")
                compared += 1
                continue
            dx, dy = w[1] - m[1], w[2] - m[2]
            ok = abs(dx) <= 0.5 and abs(dy) <= 0.5
            exact += ok
            compared += 1
            if everything or not ok:
                rows.append(f"  {key:5} {box.family:9} {box.note:34} dx {dx:+7.2f}  dy {dy:+6.2f}")
    return exact, compared, rows


def main(argv: list[str]) -> int:
    import render_record

    faces: dict = {}
    documents = {}
    for setting in probe.SETTINGS:
        data = probe.build(setting)
        word = word_positions(data, f"text-box-{setting}")
        model = model_positions(data, render_record.RecordingFonts(data, faces))
        exact, compared, rows = compare(word, model, everything="--all" in argv)
        documents[setting] = {"word": word, "misses": misses(word, model)}
        print(f"== text-box-{setting}: {exact} / {compared} first glyphs within half a pixel")
        print("\n".join(rows))
    if "--record" in argv:
        render_record.dump(OBSERVATIONS, {
            "_about": ("For each document of tools/make_text_box_probe.py, the page, pen x and baseline (device "
                       "px, 300 per inch) of the first glyph of each text box's labelled lines as Word 16.106 "
                       "drew them, and the labels the model missed at recording time; the numbers of every "
                       "face the model asked for. Measurements only; regenerate with "
                       "tools/read_text_box_probe.py --record."),
            "faces": faces, "documents": documents})
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
