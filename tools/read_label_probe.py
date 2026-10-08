#!/usr/bin/env python3
"""Measure ``make_label_probe.py`` with the model: baselines exact, by label and text face.

Every document goes through ``baselines.predict`` (``probe_documents``), so the line
height scored is the model's own (``docx2svg.vertical.line_extent``).  Each group is a
page of 22 lines of one text face with one label (or inline control) under one line
rule; its pitch is pinned to 1/21 px by the first and last baselines, and printed beside
the model's where they disagree.

``--record`` writes ``tests/fixtures/label-observations.json`` for
``tests/test_label.py``.

Usage::

    python tools/read_label_probe.py [--record]
"""

from __future__ import annotations

import json
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_label_probe as probe  # noqa: E402
import probe_documents  # noqa: E402

OBSERVATIONS = probe_documents.FIXTURES / "label-observations.json"


def tabulate(scored) -> dict:
    """``{(setting, group name): [exact, scored]}`` from ``[(probe, results)]``."""
    table: dict = defaultdict(lambda: [0, 0])
    for p, results in scored:
        for group, (exact, count) in probe_documents.tally(results, probe.kinds(p), lambda k: k).items():
            table[(p.setting, group)][0] += exact
            table[(p.setting, group)][1] += count
    return table


def by_extra(table: dict) -> dict:
    """``{label or control: [exact, scored]}`` over every setting, face and rule."""
    out: dict = defaultdict(lambda: [0, 0])
    for (_, group), (exact, count) in table.items():
        extra = group.split("/")[0]
        out[extra][0] += exact
        out[extra][1] += count
    return out


def offline(data: dict | None = None):
    data = data or json.loads(OBSERVATIONS.read_text())
    metrics = probe_documents.recorded_metrics(data["faces"])
    return [(p, probe_documents.score(probe.build(p), probe_documents.expand(data["documents"][p.name]),
                                      metrics))
            for p in probe.PROBES]


def offline_rows():
    for extra, (exact, count) in by_extra(tabulate(offline())).items():
        yield f"label probe: {extra}", exact, count, ""


def report(scored) -> None:
    table = tabulate(scored)
    for extra, (exact, count) in by_extra(table).items():
        print(f"{extra:18} {exact:5} / {count}")
    total = [sum(v[0] for v in table.values()), sum(v[1] for v in table.values())]
    print(f"total {total[0]} / {total[1]}")
    for p, results in scored:
        kinds = probe.kinds(p)
        pages: dict = defaultdict(list)
        for r in results:
            if r.block >= 0 and r.predicted is not None:
                pages[kinds[r.block]].append((r.observed, r.predicted))
        for group, ys in pages.items():
            if all(o == q for o, q in ys):
                continue
            n = len(ys) - 1
            observed = (ys[-1][0] - ys[0][0]) / n
            predicted = (ys[-1][1] - ys[0][1]) / n
            print(f"  {p.setting:4} {group:36} pitch observed {observed:7.3f} (+-{1 / n:.3f})"
                  f"  model {predicted:7.3f}  first {ys[0][0] - ys[0][1]:+d}")


def main(argv: list[str]) -> int:
    recording = probe_documents.RecordingMetrics()
    documents = {}
    scored = []
    for p in probe.PROBES:
        data = probe.build(p)
        drawn = probe_documents.measure(p.name, data)
        documents[p.name] = probe_documents.compact(drawn)
        scored.append((p, probe_documents.score(data, drawn, recording)))
    report(scored)
    if "--record" in argv[1:]:
        probe_documents.write(OBSERVATIONS, (
            "Every line Word 16.106 drew for each document of tools/make_label_probe.py: "
            "[page, baseline in 1/300-inch device px, text]; and the four hhea/typo integers of "
            "every face the model asked for (upm, ascent, descent, lineGap). Measurements only; "
            "regenerate with tools/read_label_probe.py --record."), recording.faces, documents)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
