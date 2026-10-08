#!/usr/bin/env python3
"""Measure the line-advance probes against ``docx2svg.vertical``, baseline by baseline.

Exports each probe through ``tools/oracle.py`` (cached by content), reads every baseline
out of the PDF with ``quartz_pdf`` -- integers in device px, because Quartz writes them
that way -- and compares each with the model.  The score is a count of baselines that
agree *exactly*; there is no tolerance, because the quantity is an integer.

``--record`` writes the observations to ``tests/fixtures/line-advance-observations.json``
so that ``tests/test_vertical.py`` can hold the model to them on a machine without Word.
What is recorded is measurements -- baselines in device px, and four metric integers per
face -- never a font file.

Usage::

    python tools/read_line_advance_probe.py [--record]
"""

from __future__ import annotations

import argparse
import json
import sys
from fractions import Fraction
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_line_advance_probe as probe  # noqa: E402
import oracle  # noqa: E402
import quartz_pdf  # noqa: E402

from docx2svg.vertical import FaceMetrics, baseline_px, line_pitch_px, twips_to_px  # noqa: E402

#: hhea (ascender, |descender|, lineGap) per face, as Word *lays out* with them.  Aptos
#: sets USE_TYPO_METRICS and its typo values equal these.  Times New Roman's lineGap is
#: the macOS system copy's (87), not the Office copy's (0) -- see ROADMAP.md.
FACES = {
    "Calibri": FaceMetrics(2048, 1950, 550, 0),
    "Aptos": FaceMetrics(2048, 1923, 577, 0),
    "Arial": FaceMetrics(2048, 1854, 434, 67),
    "Times New Roman": FaceMetrics(2048, 1825, 443, 87),
    "Cambria": FaceMetrics(2048, 1946, 455, 0),
    "Verdana": FaceMetrics(2048, 2059, 430, 0),
}
SWEEPS = {
    "sizes": probe.FACES,
    "rules": probe.RULE_FACES,
    "dense": probe.DENSE_FACES,
}
#: Top margin of every probe page: 1440 twips.
TOP = twips_to_px(1440)
OBSERVATIONS = HERE.parent / "tests" / "fixtures" / "line-advance-observations.json"


def measure(sweep: str, face: str) -> list[tuple[probe.Group, list[int]]]:
    pdf = oracle.export(probe.build(sweep, face), name=f"line-advance-{sweep}-{face.replace(' ', '').lower()}")
    pages = quartz_pdf.read(pdf)
    groups = probe.groups(sweep, face)
    if len(pages) != len(groups):
        raise SystemExit(f"{pdf.name}: {len(pages)} pages for {len(groups)} groups")
    out = []
    for group, page in zip(groups, pages):
        ys = sorted({round(run.y, 3) for run in page.runs if run.text.strip()})
        if len(ys) != group.lines or any(abs(y - round(y)) > 1e-3 for y in ys):
            raise SystemExit(f"{pdf.name}: unexpected baselines for {group}")
        out.append((group, [int(round(y)) for y in ys]))
    return out


def predict(face: str, group: probe.Group, count: int) -> list[int]:
    metrics = FACES[face]
    pitch = line_pitch_px(metrics, group.half_points, group.rule, group.line)
    return [
        baseline_px(TOP + k * pitch, metrics, group.half_points, group.rule, group.line)
        for k in range(count)
    ]


def rule_label(group: probe.Group) -> str:
    if group.rule == "auto" and group.line != 240:
        return "auto>240" if group.line > 240 else "auto<240"
    return group.rule


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    args = parser.parse_args(argv[1:])

    totals: dict[str, list[int]] = {}
    records = []
    for sweep, faces in SWEEPS.items():
        for face in faces:
            for group, ys in measure(sweep, face):
                predicted = predict(face, group, len(ys))
                hits = sum(p == y for p, y in zip(predicted, ys))
                label = rule_label(group)
                t = totals.setdefault(label, [0, 0, 0, 0])
                t[0] += hits
                t[1] += len(ys)
                t[2] += hits == len(ys)
                t[3] += 1
                if hits != len(ys):
                    print(f"  MISS {sweep} {face} {group}: {hits}/{len(ys)}")
                gaps = [b - a for a, b in zip(ys, ys[1:])]
                low = min(gaps) if gaps else 0
                records.append({
                    "face": face,
                    "half_points": group.half_points,
                    "rule": group.rule,
                    "line": group.line,
                    "first": ys[0],
                    "gap": low,
                    # Every gap is `gap` or `gap + 1`; one digit per gap keeps the file small.
                    "extra": "".join(str(g - low) for g in gaps),
                })

    print("rule       baselines exact         groups exact")
    for label, (hits, total, ghits, gtotal) in totals.items():
        print(f"{label:10} {hits:6} / {total:<6}       {ghits:4} / {gtotal}")

    if args.record:
        payload = {
            "_about": (
                "Baselines, in 1/300-inch device px from the page top, of every group in "
                "tools/make_line_advance_probe.py as Word 16.106 exported them. Each group "
                "starts on a new page with a 1440-twip top margin. Measurements only; "
                "regenerate with tools/read_line_advance_probe.py --record."
            ),
            "faces": {name: vars(m) for name, m in FACES.items()},
            "groups": records,
        }
        OBSERVATIONS.write_text(json.dumps(payload, separators=(",", ":"), sort_keys=True) + "\n")
        print(f"wrote {OBSERVATIONS} ({OBSERVATIONS.stat().st_size} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
