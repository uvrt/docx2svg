#!/usr/bin/env python3
"""Measure ``make_page_top_probe.py`` with the model: baselines exact, by kind of page top.

Every document goes through ``baselines.predict`` (``probe_documents``), so what is
scored is the model's own page-top rule.  A page is classed by how its first paragraph
got there, and every line on it counts for that class, because a wrong page top moves the
whole page.  For each page-top line that misses, the miss in px is shown next to the
space before it carried, so "kept in full" (the miss is that space) and "kept in part"
can be told apart.

``--record`` writes ``tests/fixtures/page-top-observations.json`` for
``tests/test_page_top.py``.

Usage::

    python tools/read_page_top_probe.py [--record]
"""

from __future__ import annotations

import json
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_page_top_probe as probe  # noqa: E402
import probe_documents  # noqa: E402

OBSERVATIONS = probe_documents.FIXTURES / "page-top-observations.json"
KINDS = ("document start", "manual break", "pageBreakBefore", "section start", "natural")


def tabulate(scored) -> dict:
    """``{(setting, kind): [exact, scored]}`` from ``[(probe, results)]``."""
    table: dict = defaultdict(lambda: [0, 0])
    for p, results in scored:
        for kind, (exact, count) in probe_documents.tally(results, probe.kinds(p), lambda k: k).items():
            table[(p.setting, kind)][0] += exact
            table[(p.setting, kind)][1] += count
    return table


def offline(data: dict | None = None):
    """``[(probe, results)]`` from the recording, without Word."""
    data = data or json.loads(OBSERVATIONS.read_text())
    metrics = probe_documents.recorded_metrics(data["faces"])
    return [(p, probe_documents.score(probe.build(p), probe_documents.expand(data["documents"][p.name]),
                                      metrics))
            for p in probe.PROBES]


def offline_rows():
    table = tabulate(offline())
    for kind in KINDS:
        exact = sum(v[0] for (s, k), v in table.items() if k == kind)
        count = sum(v[1] for (s, k), v in table.items() if k == kind)
        yield f"page-top probe: {kind}", exact, count, ""


def report(scored) -> None:
    table = tabulate(scored)
    settings = list(probe.SETTINGS)
    print(f"{'page top':16}" + "".join(f"{s:>14}" for s in settings))
    for kind in KINDS + ("plain", "break"):
        cells = [table.get((s, kind)) for s in settings]
        if not any(cells):
            continue
        print(f"{kind:16}" + "".join(f"{f'{c[0]}/{c[1]}' if c else '-':>14}" for c in cells))
    total = [sum(v[0] for v in table.values()), sum(v[1] for v in table.values())]
    print(f"total {total[0]} / {total[1]}")
    # The first line of every page whose top misses: how far, against the space it had.
    misses: dict = defaultdict(list)
    for p, results in scored:
        kinds = probe.kinds(p)
        seen = set()
        for r in results:
            if r.page in seen or r.block < 0:
                continue
            seen.add(r.page)
            if r.status == "miss":
                misses[(p.setting, kinds[r.block])].append((r.text, r.observed - r.predicted))
    for key, values in sorted(misses.items()):
        print(f"  {key}: {len(values)} page tops off; e.g. {values[:6]}")


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
            "Every line Word 16.106 drew for each document of tools/make_page_top_probe.py: "
            "[page, baseline in 1/300-inch device px, text]; and the four hhea/typo integers of "
            "every face the model asked for (upm, ascent, descent, lineGap). Measurements only; "
            "regenerate with tools/read_page_top_probe.py --record."), recording.faces, documents)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
