#!/usr/bin/env python3
"""Measure ``make_mixed_line_probe.py``: each group's line pitch against the models.

The pitch of a group is pinned to an interval by its first and last baselines (whole
device px, so ``(last - first +- 1) / (n - 1)``).  Printed against three models of a line
mixing faces or sizes -- the tallest item's height, the largest ascent plus the largest
descent ("combined"), and the model's (``vertical.line_extent`` and
``vertical.mixed_line_pitch``: combined, a list label by its ascent only, with the
multiple's extra over the tallest text item) -- and each is marked inside or outside the
interval.  Then every baseline is predicted with ``baseline_in_box`` on the model pitch.

``--record`` writes ``tests/fixtures/mixed-line-observations.json`` (the baselines, and
the metric integers of the three faces) for ``tests/test_line_box.py``.

Usage::

    python tools/read_mixed_line_probe.py [--record]
"""

from __future__ import annotations

import json
import sys
from fractions import Fraction
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_mixed_line_probe as probe  # noqa: E402

from docx2svg.resolve.fonts import CharacterFormat  # noqa: E402
from docx2svg.vertical import (  # noqa: E402
    PX_PER_PT, FaceMetrics, LineBox, baseline_in_box, line_extent, line_pitch_px, mixed_line_pitch,
    natural_height_px, tallest_natural, twips_to_px,
)

OBSERVATIONS = HERE.parent / "tests" / "fixtures" / "mixed-line-observations.json"
FACES = ("Cambria", "Symbol", "Calibri")


def item(face: str, half_points: int) -> CharacterFormat:
    return CharacterFormat("ascii", face, half_points, False, False, {})


def models(group: probe.Group, faces: dict[str, FaceMetrics]) -> tuple[LineBox, dict[str, Fraction]]:
    """The group's line box on the model pitch, and the three candidate pitches."""
    def metrics(face, bold=False, italic=False):
        return faces.get(face)

    text = [item("Cambria", 22)]
    extra = {"label": [item("Symbol", 22)], "symbol": [item("Symbol", 22)],
             "big": [item("Calibri", group.big)], "plain": []}[group.extra]
    everything = text + extra
    tallest = max(everything, key=lambda i: natural_height_px(metrics(i.face), i.half_points))
    # A label is not text; an inline run is.
    label = group.extra == "label"
    model = line_extent(text, metrics, extra) if label else line_extent(everything, metrics)
    text_only = line_extent(text, metrics) if label else None
    tallest_text = tallest_natural(text if label else everything, metrics)
    pitch = mixed_line_pitch(model, text_only, group.rule, group.line, tallest_text)
    combined = line_extent(everything, metrics)
    face, half_points, _ = model
    em = Fraction(half_points, 2) * PX_PER_PT / face.units_per_em
    box = LineBox(pitch=pitch, text_above=(face.ascent + face.line_gap) * em,
                  text_below=face.descent * em, natural=natural_height_px(face, half_points),
                  rule=group.rule, line=group.line)
    return box, {
        "tallest": line_pitch_px(metrics(tallest.face), tallest.half_points, group.rule, group.line),
        "combined": line_pitch_px(combined[0], combined[1], group.rule, group.line),
        "model": pitch,
    }


def predict(group: probe.Group, faces: dict[str, FaceMetrics], count: int) -> list[int]:
    box, _ = models(group, faces)
    top = twips_to_px(1440)
    return [baseline_in_box(top + k * box.pitch, box) for k in range(count)]


def measure() -> list[list[int]]:
    import oracle
    import quartz_pdf

    pages = quartz_pdf.read(oracle.export(probe.build(), name="mixed-line-probe"))
    if len(pages) != len(probe.GROUPS):
        raise SystemExit(f"{len(pages)} pages for {len(probe.GROUPS)} groups")
    return [sorted({round(run.y) for run in page.runs}) for page in pages]


def main(argv: list[str]) -> int:
    import face_metrics

    faces = {name: face_metrics.metrics(name) for name in FACES}
    observed = measure()
    misses = 0
    exact = total = 0
    print(f"{'group':18} {'measured pitch':>20}   tallest   combined  model")
    for group, ys in zip(probe.GROUPS, observed):
        low = Fraction(ys[-1] - ys[0] - 1, len(ys) - 1)
        high = Fraction(ys[-1] - ys[0] + 1, len(ys) - 1)
        cells = []
        _, candidates = models(group, faces)
        exact += sum(p == y for p, y in zip(predict(group, faces, len(ys)), ys))
        total += len(ys)
        for name, value in candidates.items():
            inside = low <= value <= high
            misses += name == "model" and not inside
            cells.append(f"{float(value):8.3f}{' ' if inside else '*'}")
        print(f"{group.name:18} [{float(low):8.3f},{float(high):8.3f}]  " + "  ".join(cells))
    print(f"\n* = outside the measured interval.  model misses: {misses}")
    print(f"baselines exact: {exact} / {total}")
    if "--record" in argv[1:]:
        payload = {
            "_about": (
                "Baselines, in 1/300-inch device px from the page top, of every group of "
                "tools/make_mixed_line_probe.py as Word 16.106 exported them (one list per group, "
                "in probe.GROUPS order), and the hhea integers (upm, ascent, descent, lineGap) of "
                "each face. Measurements only; regenerate with tools/read_mixed_line_probe.py --record."
            ),
            "faces": {name: [f.units_per_em, f.ascent, f.descent, f.line_gap] for name, f in faces.items()},
            "groups": {group.name: ys for group, ys in zip(probe.GROUPS, observed)},
        }
        OBSERVATIONS.write_text(json.dumps(payload, separators=(",", ":"), sort_keys=True) + "\n")
        print(f"wrote {OBSERVATIONS} ({OBSERVATIONS.stat().st_size} bytes)")
    return 1 if misses else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
