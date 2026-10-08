#!/usr/bin/env python3
"""Score the model on ``make_break_rules_probe.py``, family by family.

Every paragraph goes through ``breaks.score`` with the recorded advances
(``probe_advances``).  Printed per family: lines agreeing; per disagreeing case, Word's
lines against the model's; and for the break classes, what Word did with each
character -- ``after`` (a line may end after it), ``before`` (a line may end before it).
``--gpos`` scores kerning with the OpenType feature's pairs instead of the legacy table.
``--record`` writes ``tests/fixtures/break-rules-observations.json``.

Usage::

    python tools/read_break_rules_probe.py [-v] [--gpos] [--record]
"""

from __future__ import annotations

import json
import re
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import breaks  # noqa: E402
import make_break_rules_probe as probe  # noqa: E402
import probe_advances  # noqa: E402
import probe_documents  # noqa: E402

OBSERVATIONS = probe_documents.FIXTURES / "break-rules-observations.json"


def score(name: str, drawn, advances=None):
    data = probe.build(name)
    document = probe_documents._parsed(data)
    return breaks.score(document, drawn, advances or probe_advances.Advances(), package=data, sequential=True)


def offline(advances=None) -> dict:
    data = json.loads(OBSERVATIONS.read_text())
    return {name: score(name, probe_documents.expand(lines), advances) for name, lines in data["documents"].items()}


def _lines(text: str, spans) -> list[str]:
    return [text[a:b] for a, b in spans]


def classify(case, observed_first: str) -> str:
    """For a break-class case: where Word ended the first line."""
    tail = re.sub(r"\s", "", observed_first)
    context_left = case.name.split()[0]
    left = {"aa": "aaaa", "11": "1234", "a1": "aaaa", "1a": "1234"}[context_left]
    char = case.name.split()[-1] if not case.name.endswith(("U+00A0 \u00a0", "U+202F \u202f")) else ""
    if tail.endswith(left):
        return "before"
    if tail.endswith(left[-1] + char) or (char and tail.endswith(char)):
        return "after"
    return "at the space"


def main(argv: list[str]) -> int:
    verbose = "-v" in argv[1:]
    advances = probe_advances.Advances(gpos="--gpos" in argv[1:])
    measured = {}
    scored = {}
    try:
        for name in probe.documents():
            drawn = probe_documents.measure(name, probe.build(name))
            measured[name] = drawn
            scored[name] = score(name, drawn, advances)
    finally:
        import oracle

        oracle.recover()
    table: dict = defaultdict(lambda: [0, 0])
    classes: dict = defaultdict(dict)
    for name, results in scored.items():
        for case, r in zip(probe.documents()[name], results):
            cell = table[case.family]
            if r.status != "scored":
                print(f"  OUT {case.family} {case.name}: {r.status}")
                continue
            cell[0] += r.agree
            cell[1] += r.lines
            if case.family in ("after", "before"):
                context, char = case.name.split()[0], chr(int(case.name.split()[1][2:], 16))
                left = {"aa": "aaaa", "11": "1234", "a1": "aaaa", "1a": "1234"}[context]
                first = r.texts()[0]
                where = ("after" if first.endswith(left + char) and not char.isspace() else
                         "before" if first.endswith(left) else "at the space")
                classes[(case.face, context, char)][case.family] = where
            if r.agree < r.lines or verbose:
                print(f"  {'!!' if r.agree < r.lines else '  '} {case.family:7} {case.face:16} {case.name}: "
                      f"{r.agree}/{r.lines}")
                print(f"       word:  {r.texts()}")
                if r.agree < r.lines:
                    print(f"       model: {r.texts('predicted')}")
    print("break classes: where Word ended the line (after: X fits, next letter not; before: X does not fit)")
    for (face, context, char), found in sorted(classes.items(), key=lambda kv: (kv[0][1], kv[0][2], kv[0][0])):
        print(f"   {context} U+{ord(char):04X} {char!r:8} {face:16} after-case: {found.get('after', '-'):13}"
              f" before-case: {found.get('before', '-')}")
    for family, (agree, lines) in table.items():
        print(f"{family:8} {agree:5} / {lines}")
    if "--record" in argv[1:]:
        OBSERVATIONS.write_text(json.dumps({
            "_about": ("Every line Word 16.106 drew for each document of tools/make_break_rules_probe.py: "
                       "[page, baseline in device px, text]. Measurements only; regenerate with "
                       "tools/read_break_rules_probe.py --record."),
            "documents": {name: probe_documents.compact(drawn) for name, drawn in measured.items()},
        }, ensure_ascii=False, separators=(",", ":")) + "\n")
        print(f"wrote {OBSERVATIONS}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
