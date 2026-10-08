#!/usr/bin/env python3
"""Read the δ sweep (``make_wrap_budget_probe.py``): where is Word's budget's edge?

Every paragraph goes through ``breaks.score`` -- the model's own breaker -- with the
recorded advances (``probe_advances``), and a case counts when every one of its lines
breaks where Word's does.  Beside it, Word's own edge per family: the largest δ whose
line did not hold its second word, and the smallest that did (δ in 1/4096 pt, the
budget less the line's width).  The refuted strict budget (``x < right``) is scored by
moving every right edge one unit in.  The generator's frozen advances are checked
against the installed faces first.

``--record`` writes ``tests/fixtures/wrap-budget-observations.json`` (the drawn lines,
``[page, baseline, text]``, per document).  Usage::

    python tools/read_wrap_budget_probe.py [--record]
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

import make_wrap_budget_probe as probe  # noqa: E402

OBSERVATIONS = HERE.parent / "tests" / "fixtures" / "wrap-budget-observations.json"


def _norm(text: str) -> str:
    return re.sub(r"\s+", "", text)


def check_widths() -> list[str]:
    try:
        import face_advances
    except ImportError:
        return ["fontTools not installed: frozen advances not checked"]
    problems = []
    for face, table in probe.WIDTHS.items():
        found = face_advances.advances(face)
        if found is None:
            problems.append(f"{face}: not installed")
            continue
        for char, width in table.items():
            if found.advance(char) != width:
                problems.append(f"{face} {char!r}: frozen {width}, installed {found.advance(char)}")
    return problems


def measure() -> dict[str, list]:
    """Word's drawn lines of each document (``baselines.DrawnLine``)."""
    import probe_documents

    out = {}
    try:
        for name in probe.documents():
            out[name] = probe_documents.measure(name, probe.build(name))
    finally:
        import oracle

        oracle.recover()
    return out


def score(name: str, drawn, advances=None, *, right_shift: int = 0):
    """The model on one document: ``breaks.score`` over its paragraphs, in order.

    ``right_shift`` moves every right edge by that many units -- ``-1`` turns the
    inclusive budget into a strict one (``x < right``), which is how the refuted rule is
    scored without a second breaker.
    """
    import breaks
    import probe_advances
    import probe_documents
    from docx2svg import linebreak

    data = probe.build(name)
    document = probe_documents._parsed(data)
    original = linebreak.geometry
    if right_shift:
        def shifted(*args, **kwargs):
            found = original(*args, **kwargs)
            return linebreak.Geometry(found.first_start, found.start, found.right + right_shift, found.tabs,
                                      found.default_tab, found.hanging_stop)
        linebreak.geometry = shifted
    try:
        return breaks.score(document, drawn, advances or probe_advances.Advances(), package=data, sequential=True)
    finally:
        linebreak.geometry = original


def offline(right_shift: int = 0) -> dict:
    import probe_documents

    data = json.loads(OBSERVATIONS.read_text())
    return {name: score(name, probe_documents.expand(lines), right_shift=right_shift)
            for name, lines in data["documents"].items()}


def held(case, result) -> bool:
    """Whether Word's line under test holds all of ``under_test``."""
    index = 1 if case.lead else 0
    texts = result.texts()
    return len(texts) > index and texts[index] == _norm(case.under_test)


def report(scored: dict) -> dict:
    """Per (setting, family): cases the model breaks as Word does, and the edges of δ."""
    table: dict = defaultdict(lambda: [0, 0])
    edges: dict = defaultdict(lambda: [None, None])  # largest δ not held, smallest held
    for name, results in scored.items():
        for case, r in zip(probe.documents()[name], results):
            key = (case.setting, case.family)
            cell = table[key]
            cell[0] += r.status == "scored" and r.agree == r.lines
            cell[1] += 1
            if r.status != "scored":
                continue
            edge = edges[key]
            if held(case, r):
                edge[1] = case.delta if edge[1] is None else min(edge[1], case.delta)
            else:
                edge[0] = case.delta if edge[0] is None else max(edge[0], case.delta)
    total = [sum(v[0] for v in table.values()), sum(v[1] for v in table.values())]
    print(f"cases whose every line the model breaks as Word: {total[0]} / {total[1]}")
    for key, (a, n) in sorted(table.items()):
        no, yes = edges[key]
        print(f"   {key[0]:5} {key[1]:7} {a:4} / {n:<4} Word: not held up to δ = {no}, held from δ = {yes}")
    return table


def main(argv: list[str]) -> int:
    import probe_documents

    for problem in check_widths():
        print("WIDTH CHECK:", problem)
    drawn = measure()
    print("the model (inclusive: x <= right):")
    report({name: score(name, lines) for name, lines in drawn.items()})
    print("refuted, strict (x < right):")
    report({name: score(name, lines, right_shift=-1) for name, lines in drawn.items()})
    if "--record" in argv[1:]:
        OBSERVATIONS.write_text(json.dumps({
            "_about": ("Every line Word 16.106 drew for each document of tools/make_wrap_budget_probe.py: "
                       "[page, baseline in device px, text]. Measurements only; regenerate with "
                       "tools/read_wrap_budget_probe.py --record."),
            "documents": {name: probe_documents.compact(lines) for name, lines in drawn.items()},
        }, ensure_ascii=False, separators=(",", ":"), sort_keys=True) + "\n")
        print(f"wrote {OBSERVATIONS}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
