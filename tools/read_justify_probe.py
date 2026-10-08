#!/usr/bin/env python3
"""Read ``make_justify_probe.py``: did Word hold ``W2`` on the line, δ units past the edge?

For every case, Word's drawn lines (``probe_documents.measure``) are walked paragraph by
paragraph and the first line is compared with the line under test.  Scored against Word:

* ``model`` -- the library's breaker (``linebreak.squeezes``: each space lets a justified
  line in mode 15 run past the edge by a quarter of its width, rounded half up, so long
  as each shrinks by at most half of what each would grow were the word sent down);
* ``no squeeze`` -- the budget of Phase 3.1, no line past the edge;
* ``quarter of the spaces`` -- a quarter of the spaces' total width, unrounded;
* ``per space, floored`` -- each space's quarter truncated.

``--record`` writes Word's answers to ``tests/fixtures/justify-observations.json``;
``tests/test_justify.py`` holds the breaker to them offline.

Usage::

    python tools/read_justify_probe.py [--record]
"""

from __future__ import annotations

import argparse
import math
import re
import sys
from fractions import Fraction
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_justify_probe as probe  # noqa: E402
import make_wrap_budget_probe as budget  # noqa: E402

OBSERVATIONS = HERE.parent / "tests" / "fixtures" / "justify-observations.json"


def _norm(text: str) -> str:
    return re.sub(r"\s+", "", text)


def documents() -> list[tuple[str, bytes, tuple]]:
    """``(name, .docx bytes, cases)`` of every document of both rounds."""
    out = [(f"justify-{name}", probe.build(name), probe.cases()) for name in probe.SETTINGS]
    out += [(f"justify-{name}", probe.build_fine(name), probe.fine_cases()) for name in probe.FINE_SETTINGS]
    out += [(f"justify-{name}", probe.build_third(name), probe.third_cases()) for name in probe.THIRD_SETTINGS]
    return out


def head(case) -> str:
    """The line under test: everything before the tail."""
    return case.text.split(" Hnnn")[0]


def held(lines: list[str], cases) -> list[bool]:
    """Per case, whether the paragraph's first line holds the whole line under test."""
    texts = [_norm(line) for line in lines]
    out, i = [], 0
    for case in cases:
        full = _norm(case.text)
        out.append(texts[i].startswith(_norm(head(case))))
        seen = ""
        while len(seen) < len(full):
            seen += texts[i]
            i += 1
        if seen != full:
            raise ValueError(f"lost the paragraphs at {case}")
    return out


def spaces_of(case) -> list[int]:
    """The width of every space before ``W2``, 1/4096 pt."""
    if hasattr(case, "spaces") and isinstance(case.spaces, tuple):
        sizes = case.spaces
    else:
        sizes = (case.half_points,) * case.spaces
    return [budget.WIDTHS[case.face][" "] * hp for hp in sizes]


def rules(case) -> dict[str, bool]:
    """Whether each rule tried holds ``W2`` (δ is past the edge by that much)."""
    widths = spaces_of(case)
    return {
        "no squeeze": case.delta <= 0,
        "quarter of the spaces": case.delta <= Fraction(sum(widths), 4),
        "per space, floored": case.delta <= sum(w // 4 for w in widths),
        "per space, rounded (the first condition alone)": case.delta <= sum(math.floor(Fraction(w, 4) + Fraction(1, 2))
                                                          for w in widths),
    }


def model_held(layout, cases) -> list[bool]:
    """Per case, whether the laid-out paragraph's first line holds the line under test."""
    firsts: dict[str, str] = {}
    for page in layout.pages:
        for line in page.lines:
            if line.number == 0:
                firsts.setdefault(line.path, "".join("".join(span.chars) for span in line.spans))
    ordered = list(firsts.values())
    return [_norm(text).startswith(_norm(head(case))) for case, text in zip(cases, ordered)]


def main(argv: list[str]) -> int:
    import probe_documents
    import render_record
    from docx2svg import _render

    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--record", action="store_true")
    args = parser.parse_args(argv[1:])
    recorded: dict = {}
    faces: dict = {}
    for name, data, cases in documents():
        lines = [line.text for line in probe_documents.measure(name, data)]
        word = held(lines, cases)
        justified = name.endswith("15-both")
        # The third round's line holds more spaces than ``spaces_of`` knows of; score the
        # rules there only through the model.
        scores = {}
        for case, answer in zip(cases, word):
            for rule, verdict in (rules(case).items() if not name.startswith("justify-third") else ()):
                if not justified and rule != "no squeeze":
                    continue
                scores[rule] = scores.get(rule, 0) + (verdict == answer)
        _, layout, _, _ = _render(data, render_record.options(render_record.RecordingFonts(data, faces)))
        ours = model_held(layout, cases)
        scores["model"] = sum(a == b for a, b in zip(ours, word))
        print(f"{name:22} cases {len(cases)}, Word held {sum(word)}: "
              + ", ".join(f"{rule} {n}" for rule, n in scores.items()))
        recorded[name] = word
    if args.record:
        import json

        OBSERVATIONS.write_text(json.dumps({
            "_about": ("Whether Word 16.106 held the line under test of each case of tools/make_justify_probe.py "
                       "on its paragraph's first line, per document, in case order "
                       "(tools/read_justify_probe.py). Measurements only; regenerate with --record."),
            "documents": recorded,
        }, separators=(",", ":"), sort_keys=True) + "\n")
        print(f"wrote {OBSERVATIONS}")
        render_record.dump(OBSERVATIONS.parent / "justify-faces.json", dict(sorted(faces.items())))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
