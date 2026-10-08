#!/usr/bin/env python3
"""Score the model on ``make_wrap_probe.py``: lines that break after the same word as Word.

Every paragraph goes through ``breaks.score`` -- the model's own code -- with the advances
of ``tests/fixtures/probe-advances.json``.  Printed: lines agreeing by face, size,
column and style, and the aligned families apart; ``-v`` prints each paragraph that
disagrees.  ``--record`` writes ``tests/fixtures/wrap-observations.json`` (the drawn
lines, as ``[page, baseline, text]``) for ``tests/test_linebreak.py``.

Usage::

    python tools/read_wrap_probe.py [-v] [--record]
"""

from __future__ import annotations

import json
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import breaks  # noqa: E402
import make_wrap_probe as probe  # noqa: E402
import probe_advances  # noqa: E402
import probe_documents  # noqa: E402

OBSERVATIONS = probe_documents.FIXTURES / "wrap-observations.json"


def score(drawn, advances=None):
    data = probe.build()
    document = probe_documents._parsed(data)
    return breaks.score(document, drawn, advances or probe_advances.Advances(), package=data, sequential=True)


def offline(advances=None):
    data = json.loads(OBSERVATIONS.read_text())
    return score(probe_documents.expand(data["lines"]), advances)


def tabulate(results) -> dict:
    """``{(dimension, value): [agree, lines, conditional]}``."""
    table: dict = defaultdict(lambda: [0, 0, 0])
    for case, r in zip(probe.cases(), results):
        if r.status != "scored":
            table[("out", r.status)][1] += 1
            continue
        style = ("bold " if case.bold else "") + ("italic" if case.italic else "") or "regular"
        keys = [("all", "aligned" if case.jc else "left")]
        if case.jc:
            keys.append(("jc", case.jc))
        else:
            keys += [("face", case.face), ("size", case.half_points), ("column", case.column), ("style", style)]
        for k in keys:
            cell = table[k]
            cell[0] += r.agree
            cell[1] += r.lines
            cell[2] += sum(r.conditional)
    return table


def main(argv: list[str]) -> int:
    verbose = "-v" in argv[1:]
    drawn = probe_documents.measure("wrap-probe", probe.build())
    import oracle

    oracle.recover()
    results = score(drawn)
    table = tabulate(results)
    for key, (agree, lines, conditional) in sorted(table.items(), key=lambda kv: (kv[0][0], str(kv[0][1]))):
        print(f"{key[0]:7} {str(key[1]):28} {agree:5} / {lines:<5} (conditional {conditional})")
    if verbose:
        for case, r in zip(probe.cases(), results):
            if r.status == "scored" and r.agree < r.lines:
                text = "".join(ch for ch in case.text if not ch.isspace())
                print(f"  {case.face} {case.half_points} {case.column} b={case.bold} i={case.italic} jc={case.jc}:"
                      f" {r.agree}/{r.lines}")
                for o, p, c in zip(r.observed, r.predicted + [None] * 50, r.conditional):
                    if o != p:
                        print(f"     word ...{text[max(0, o[1] - 15):o[1]]!r:18} model"
                              f" ...{text[max(0, p[1] - 15):p[1]] if p else ''!r:18} conditional={c}")
    if "--record" in argv[1:]:
        OBSERVATIONS.write_text(json.dumps({
            "_about": ("Every line Word 16.106 drew for tools/make_wrap_probe.py: [page, baseline in "
                       "device px, text]. Measurements only; regenerate with "
                       "tools/read_wrap_probe.py --record."),
            "lines": probe_documents.compact(drawn),
        }, ensure_ascii=False, separators=(",", ":")) + "\n")
        print(f"wrote {OBSERVATIONS}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
