#!/usr/bin/env python3
"""Measure ``make_border_group_probe.py`` through ``baselines.predict``: per case, every
line's miss in px (observed less predicted).

``--record`` writes ``tests/fixtures/border-group-observations.json``.

Usage::

    python tools/read_border_group_probe.py [--record]
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_border_group_probe as probe  # noqa: E402
import probe_documents  # noqa: E402

OBSERVATIONS = probe_documents.FIXTURES / "border-group-observations.json"


def per_case(results) -> dict[int, list]:
    """Case number -> ``[(line text, observed - predicted)]``."""
    out: dict[int, list] = {}
    for r in results:
        words = r.text.split()
        if len(words) >= 2 and words[0] == "Case" and r.predicted is not None:
            out.setdefault(int(words[1]), []).append((" ".join(words[2:]), r.observed - r.predicted))
    return out


def offline(data: dict | None = None):
    data = data or json.loads(OBSERVATIONS.read_text(encoding="utf-8"))
    metrics = probe_documents.recorded_metrics(data["faces"])
    return {setting: probe_documents.score(probe.build(setting), probe_documents.expand(data["documents"][setting]),
                                           metrics)
            for setting in probe.SETTINGS}


def exact(scored) -> tuple[int, int]:
    lines = [r for results in scored.values() for r in results if r.status in ("exact", "miss")]
    return sum(r.status == "exact" for r in lines), len(lines)


def report(scored) -> None:
    for setting, results in scored.items():
        cases = per_case(results)
        print(f"== {setting}")
        for number, (name, _) in enumerate(probe.CASES):
            offs = cases.get(number, [])
            if any(d for _, d in offs):
                print(f"  {name:18} {offs}")
    print("exact", exact(scored))


def main(argv: list[str]) -> int:
    recording = probe_documents.RecordingMetrics()
    documents, scored = {}, {}
    for setting in probe.SETTINGS:
        data = probe.build(setting)
        drawn = probe_documents.measure(f"border-group-{setting}", data)
        documents[setting] = probe_documents.compact(drawn)
        scored[setting] = probe_documents.score(data, drawn, recording)
    report(scored)
    if "--record" in argv[1:]:
        probe_documents.write(OBSERVATIONS, (
            "Every line Word 16.106 drew for each document of tools/make_border_group_probe.py: [page, "
            "baseline in 1/300-inch device px, text]; and the four hhea/typo integers of every face the model "
            "asked for. Measurements only; regenerate with tools/read_border_group_probe.py --record."),
            recording.faces, documents)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
