#!/usr/bin/env python3
"""Measure ``make_page_top_collapse_probe.py`` through ``baselines.predict``: every line
exact, by kind of page top and setting; and each page top's miss in px against the
space after above the break.

``--record`` writes ``tests/fixtures/page-top-collapse-observations.json``.

Usage::

    python tools/read_page_top_collapse_probe.py [--record]
"""

from __future__ import annotations

import json
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_page_top_collapse_probe as probe  # noqa: E402
import probe_documents  # noqa: E402

OBSERVATIONS = probe_documents.FIXTURES / "page-top-collapse-observations.json"


def tabulate(scored) -> dict:
    table: dict = defaultdict(lambda: [0, 0])
    for setting, results in scored:
        for kind, (exact, count) in probe_documents.tally(results, probe.kinds(), lambda k: k).items():
            table[(setting, kind)][0] += exact
            table[(setting, kind)][1] += count
    return table


def offline(data: dict | None = None):
    data = data or json.loads(OBSERVATIONS.read_text(encoding="utf-8"))
    metrics = probe_documents.recorded_metrics(data["faces"])
    return [(setting, probe_documents.score(probe.build(setting), probe_documents.expand(data["documents"][setting]),
                                            metrics))
            for setting in probe.SETTINGS]


def report(scored) -> None:
    table = tabulate(scored)
    for (setting, kind), (exact, count) in sorted(table.items()):
        print(f"{setting:6} {kind:18} {exact} / {count}")
    for setting, results in scored:
        kinds = probe.kinds()
        seen = set()
        misses = defaultdict(list)
        for r in results:
            if r.page in seen or r.block < 0:
                continue
            seen.add(r.page)
            if r.status == "miss" and kinds[r.block] in probe.KINDS:
                number = int(r.text.split()[1])
                case = probe.CASES[number]
                misses[case.kind].append((case.after, case.before, r.observed - r.predicted))
        for kind, values in misses.items():
            print(f"  {setting} {kind}: {len(values)} page tops off (after, before, px): {values[:12]}")


def main(argv: list[str]) -> int:
    recording = probe_documents.RecordingMetrics()
    documents, scored = {}, []
    for setting in probe.SETTINGS:
        data = probe.build(setting)
        drawn = probe_documents.measure(f"page-top-collapse-{setting}", data)
        documents[setting] = probe_documents.compact(drawn)
        scored.append((setting, probe_documents.score(data, drawn, recording)))
    report(scored)
    if "--record" in argv[1:]:
        probe_documents.write(OBSERVATIONS, (
            "Every line Word 16.106 drew for each document of tools/make_page_top_collapse_probe.py: [page, "
            "baseline in 1/300-inch device px, text]; and the four hhea/typo integers of every face the model "
            "asked for. Measurements only; regenerate with tools/read_page_top_collapse_probe.py --record."),
            recording.faces, documents)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
