#!/usr/bin/env python3
"""Measure ``make_line_box_probe.py`` against ``docx2svg.vertical``, baseline by baseline.

Every paragraph of the probe is one line, so the lines on a page are its paragraphs in
order.  The reader rebuilds each page's stack from the paragraphs the generator wrote --
the collapsed paragraph gap (``paragraph_gap_px``), the pitch, the bottom border
(``border_px``) -- and predicts every baseline with ``baseline_in_box``.  The score is
the count of baselines that agree **exactly**.

``--record`` writes ``tests/fixtures/line-box-observations.json``: the baselines Word
drew (integers, device px, one list per page) and the four metric integers of each face,
so ``tests/test_line_box.py`` holds the model to them without Word.  Measurements only;
the paragraphs are regenerated from the generator, and the test checks the page split
still matches.

Usage::

    python tools/read_line_box_probe.py [--record] [--no-line-box]

``--no-line-box`` keeps the paragraph spacing and border out of the box the baseline
rounds in (each line rounds in its own pitch, as in Phase 2): the refutation recorded in
ROADMAP.md.
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import defaultdict
from fractions import Fraction
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_line_box_probe as probe  # noqa: E402

from docx2svg.vertical import (  # noqa: E402
    FaceMetrics, baseline_in_box, border_px, face_line_box, paragraph_gap_px, twips_to_px,
)

OBSERVATIONS = HERE.parent / "tests" / "fixtures" / "line-box-observations.json"
#: Every probe page starts at the 1440-twip top margin.
TOP_TWIPS = 1440


def predict(p: probe.Probe, face: FaceMetrics, *, line_box: bool = True) -> list[list[int]]:
    """Every baseline of the probe, page by page, from the paragraphs alone."""
    out = []
    for page in probe.pages(p):
        top = twips_to_px(TOP_TWIPS)
        previous_after = 0
        ys = []
        for index, para in enumerate(page):
            space_before = Fraction(0)
            if index:
                gap, space_before = paragraph_gap_px(previous_after, para.before)
                top += gap
            border = border_px(*para.border) if para.border else Fraction(0)
            below = border + twips_to_px(para.after)
            box = face_line_box(face, p.half_points, para.rule, para.line,
                                space_before=space_before if line_box else Fraction(0),
                                space_after=below if line_box else Fraction(0))
            ys.append(baseline_in_box(top - box.space_before, box))
            top += box.pitch + border
            previous_after = para.after
        out.append(ys)
    return out


def measure(p: probe.Probe) -> list[list[int]]:
    import oracle
    import quartz_pdf

    pages = quartz_pdf.read(oracle.export(probe.build(p), name=p.name))
    expected = probe.pages(p)
    if len(pages) != len(expected):
        raise SystemExit(f"{p.name}: {len(pages)} pages for {len(expected)}")
    out = []
    for page, paras in zip(pages, expected):
        ys = [line.y for line in quartz_pdf.lines([page])]
        if len(ys) != len(paras) or any(abs(y - round(y)) > 1e-3 for y in ys):
            raise SystemExit(f"{p.name}: page {len(out) + 1} has {len(ys)} lines for {len(paras)}")
        out.append([int(round(y)) for y in ys])
    return out


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--no-line-box", action="store_true")
    args = parser.parse_args(argv[1:])

    import face_metrics

    faces: dict[str, FaceMetrics] = {}
    observed: dict[str, list[list[int]]] = {}
    totals: dict[str, list[int]] = defaultdict(lambda: [0, 0])
    for p in probe.PROBES:
        face = faces.setdefault(p.face, face_metrics.metrics(p.face))
        ys = measure(p)
        observed[p.name] = ys
        predicted = predict(p, face, line_box=not args.no_line_box)
        hits = sum(a == b for pa, pb in zip(predicted, ys) for a, b in zip(pa, pb))
        count = sum(len(page) for page in ys)
        totals[p.sweep][0] += hits
        totals[p.sweep][1] += count
        if hits != count:
            print(f"  MISS {p.name}: {hits} / {count}")
    print("sweep            baselines exact")
    for sweep, (hits, count) in totals.items():
        print(f"{sweep:16} {hits:6} / {count}")
    print(f"{'total':16} {sum(t[0] for t in totals.values()):6} / {sum(t[1] for t in totals.values())}")

    if args.record:
        payload = {
            "_about": (
                "Baselines, in 1/300-inch device px from the page top, of every line of every "
                "probe in tools/make_line_box_probe.py as Word 16.106 exported them, one list per "
                "page; and the hhea/typo integers (upm, ascent, descent, lineGap) of each face. "
                "Measurements only; regenerate with tools/read_line_box_probe.py --record."
            ),
            "faces": {name: [f.units_per_em, f.ascent, f.descent, f.line_gap] for name, f in faces.items()},
            "probes": observed,
        }
        OBSERVATIONS.write_text(json.dumps(payload, separators=(",", ":"), sort_keys=True) + "\n")
        print(f"wrote {OBSERVATIONS} ({OBSERVATIONS.stat().st_size} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
