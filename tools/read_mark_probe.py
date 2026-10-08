#!/usr/bin/env python3
"""Measure ``make_mark_probe.py`` with the model: does a large paragraph mark count?

Every document goes through ``baselines.predict`` (``probe_documents``), so what is
scored is the model as it stands.  Per case and setting: every line of the case's page
scored, and, for each line of the case paragraph itself, the miss in px (predicted minus
drawn).  ROADMAP.md, "A paragraph mark larger than its text", records what the probe
shows.

``--record`` writes ``tests/fixtures/mark-observations.json`` for ``tests/test_mark.py``.

Usage::

    python tools/read_mark_probe.py [--record]
"""

from __future__ import annotations

import json
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_mark_probe as probe  # noqa: E402
import probe_documents  # noqa: E402

OBSERVATIONS = probe_documents.FIXTURES / "mark-observations.json"


def tabulate(scored) -> dict:
    """``{(setting, case): [exact, scored, [misses in px of the case paragraph's lines]]}``
    for the first two documents; the sweep's cases are keyed ``("sweep-" + setting, case)``."""
    table: dict = defaultdict(lambda: [0, 0, []])
    for p, results in scored:
        kinds = probe.kinds(p)
        setting = f"sweep-{p.setting}" if p.sweep else p.setting
        for r in results:
            if r.status not in ("exact", "miss") or r.block < 0:
                continue
            cell = table[(setting, kinds[r.block])]
            cell[0] += r.status == "exact"
            cell[1] += 1
            if r.block % 5 == 1:
                cell[2].append(r.predicted - r.observed)
    return table


def offline(data: dict | None = None, which=lambda p: True):
    """``[(probe, results)]`` from the recording, without Word."""
    data = data or json.loads(OBSERVATIONS.read_text(encoding="utf-8"))
    metrics = probe_documents.recorded_metrics(data["faces"])
    return [(p, probe_documents.score(probe.build(p), probe_documents.expand(data["documents"][p.name]),
                                      metrics))
            for p in probe.PROBES if p.name in data["documents"] and which(p)]


def rollup(table: dict, key) -> dict:
    """``{key(setting, case): [exact, scored]}``."""
    out: dict = defaultdict(lambda: [0, 0])
    for (setting, case), (exact, count, _) in table.items():
        cell = out[key(setting, case)]
        cell[0] += exact
        cell[1] += count
    return out


def report(scored) -> None:
    table = tabulate(scored)
    for case in probe.CASES:
        cells = [table.get((s, case.name)) for s in probe.SETTINGS]
        print(f"{case.name:18}" + "".join(
            f"{f'{c[0]}/{c[1]} (case line {c[2][0]:+d})' if c else '-':>24}" for c in cells))
    print("sweep, by family and setting:")
    for (family, setting), (exact, count) in sorted(rollup(
            {k: v for k, v in table.items() if k[0].startswith("sweep")},
            lambda s, c: (c.split("/")[0], s)).items()):
        print(f"  {family:28} {setting:10} {exact:5} / {count}")
    misses = defaultdict(list)
    for (setting, case), (exact, count, lines) in table.items():
        if setting.startswith("sweep") and exact < count:
            misses[case].append(f"{setting[6:]}: {exact}/{count} {lines}")
    for case, rows in sorted(misses.items()):
        print(f"  miss {case:52} " + "; ".join(rows))
    total = [sum(v[0] for v in table.values()), sum(v[1] for v in table.values())]
    print(f"total {total[0]} / {total[1]}")


def main(argv: list[str]) -> int:
    recording = probe_documents.RecordingMetrics()
    documents = {}
    scored = []
    try:
        for p in probe.PROBES:
            data = probe.build(p)
            drawn = probe_documents.measure(p.name, data)
            documents[p.name] = probe_documents.compact(drawn)
            scored.append((p, probe_documents.score(data, drawn, recording)))
    finally:
        import oracle

        oracle.recover()
    # Every face the probe names, whether or not the model asked for it: a rule that
    # leaves the mark out does not ask for the mark's face, and the refuted rules the
    # tests score offline need it.
    for case in (*probe.CASES, *probe.SWEEP_CASES):
        for face in {case.text_face, case.mark_face or case.text_face,
                     *(run[2] for run in case.runs if len(run) > 2)}:
            recording(face, case.bold, False)
    report(scored)
    if "--record" in argv[1:]:
        probe_documents.write(OBSERVATIONS, (
            "Every line Word 16.106 drew for each document of tools/make_mark_probe.py: "
            "[page, baseline in 1/300-inch device px, text]; and the metric integers of "
            "every face the model asked for (upm, ascent, descent, lineGap, OS/2 script sizes). "
            "Measurements only; regenerate with tools/read_mark_probe.py --record."),
            recording.faces, documents)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
