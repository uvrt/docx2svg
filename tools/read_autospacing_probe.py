#!/usr/bin/env python3
"""Measure ``make_autospacing_probe.py`` with the model: baselines exact, by case.

Every document goes through ``baselines.predict`` (``probe_documents``), so the spacing
scored is what ``docx2svg.resolve`` makes of the autospacing flags.  A page is classed
by its case, and every line on it counts.  Each missed line is listed with its miss in
px, so a spacing that is wrong by a fixed amount shows as one.

``--record`` writes ``tests/fixtures/autospacing-observations.json`` for
``tests/test_autospacing.py``.

Usage::

    python tools/read_autospacing_probe.py [--record]
"""

from __future__ import annotations

import json
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_autospacing_probe as probe  # noqa: E402
import probe_documents  # noqa: E402

OBSERVATIONS = probe_documents.FIXTURES / "autospacing-observations.json"
CASES = ("document start", "stated", "neighbour", "consecutive", "list", "list edges",
         "list stated", "collapse", "contextual", "section start", "manual break", "pageBreakBefore")


def tabulate(scored) -> dict:
    """``{(setting, case): [exact, scored]}`` from ``[(probe, results)]``."""
    table: dict = defaultdict(lambda: [0, 0])
    for p, results in scored:
        for case, (exact, count) in probe_documents.tally(results, probe.kinds(p), lambda k: k).items():
            table[(p.setting, case)][0] += exact
            table[(p.setting, case)][1] += count
    return table


def offline(data: dict | None = None):
    data = data or json.loads(OBSERVATIONS.read_text())
    metrics = probe_documents.recorded_metrics(data["faces"])
    return [(p, probe_documents.score(probe.build(p), probe_documents.expand(data["documents"][p.name]),
                                      metrics))
            for p in probe.PROBES]


def offline_rows():
    table = tabulate(offline())
    for case in CASES:
        exact = sum(v[0] for (s, k), v in table.items() if k == case)
        count = sum(v[1] for (s, k), v in table.items() if k == case)
        yield f"autospacing probe: {case}", exact, count, ""


def report(scored) -> None:
    table = tabulate(scored)
    settings = list(probe.SETTINGS)
    print(f"{'case':16}" + "".join(f"{s:>12}" for s in settings))
    for case in CASES:
        cells = [table.get((s, case)) for s in settings]
        print(f"{case:16}" + "".join(f"{f'{c[0]}/{c[1]}' if c else '-':>12}" for c in cells))
    total = [sum(v[0] for v in table.values()), sum(v[1] for v in table.values())]
    print(f"total {total[0]} / {total[1]}")
    for p, results in scored:
        misses = [(r.text, r.observed - r.predicted) for r in results if r.status == "miss"]
        if misses:
            print(f"  {p.setting}: {len(misses)} lines off, e.g. {misses[:12]}")


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
            "Every line Word 16.106 drew for each document of tools/make_autospacing_probe.py: "
            "[page, baseline in 1/300-inch device px, text]; and the four hhea/typo integers of "
            "every face the model asked for (upm, ascent, descent, lineGap). Measurements only; "
            "regenerate with tools/read_autospacing_probe.py --record."), recording.faces, documents)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
