#!/usr/bin/env python3
"""Measure ``make_section_end_probe.py``: where Word puts a section's last paragraph that
does not fit on its page.

For every case: the anchor's page, the ``next`` line's page and baseline (and the
candidate's own page when it holds text); from them, whether the candidate **stayed** on
the anchor's page or **moved** to the next.  With the ``next`` line on the page after the
anchor's, it stayed if that line stands where the anchor stands on its page (the page's
first baseline) and moved if it stands a candidate's height lower -- or, for
``next-page``, two pages after the anchor's.  For the document's own last paragraph, the
number of pages Word made.  The model (``docx2svg``, on the file alone) is read the same
way, so the two are compared case by case.

``--record`` writes ``tests/fixtures/section-end-observations.json``.

Usage::

    python tools/read_section_end_probe.py [--record]
"""

from __future__ import annotations

import re
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_section_end_probe as probe  # noqa: E402

OBSERVATIONS = HERE.parent / "tests" / "fixtures" / "section-end-observations.json"
CASE = re.compile(r"Case (\d+) (anchor|next|cand)")


def word_lines(data: bytes, name: str) -> tuple[list[tuple[int, float, str]], int]:
    """Word's ``(page, baseline px, text)`` of every line, and its number of pages.

    Word draws an empty paragraph's mark as a space; such a line is left out *before*
    ``baselines.merge_raised`` folds smaller runs into their neighbours, or the next line
    (11 pt, 98 px below a 48 pt mark: within half the mark's size) is folded into the
    mark's line and read at the mark's baseline.  That fold is what the recording of
    2026-09-27 read as a 48 pt empty paragraph standing 146 px high at a page's top: Word
    stands it 244 px, its line, as the model does."""
    import baselines
    import oracle
    import quartz_pdf

    pages = quartz_pdf.read(oracle.export(data, name=name))
    drawn = [line for line in quartz_pdf.lines(pages) if "".join(run.text for run in line.runs).strip()]
    lines = [(line.page, round(line.y, 3), "".join(run.text for run in line.runs))
             for line in baselines.merge_raised(drawn)]
    return lines, len(pages)


def model_lines(data: bytes, fonts=None) -> tuple[list[tuple[int, float, str]], int]:
    """The model's ``(page, baseline px, text)`` of every body line, and its pages,
    measured with ``fonts`` (``render_record``'s; the installed faces when ``None``)."""
    import render_record
    from docx2svg import convert_docx_to_layout

    layout = convert_docx_to_layout(data, render_record.options(fonts) if fonts is not None else None)
    out = []
    for page in layout.pages:
        for line in page.lines:
            if line.story:
                continue
            text = "".join("".join(span.chars) for span in line.spans)
            out.append((page.number + 1, float(line.baseline), text))
    return out, len(layout.pages)


def outcomes(lines: list[tuple[int, float, str]]) -> dict[int, str]:
    """Case number -> ``stayed`` / ``moved`` (``moved-2`` for both of a pair) / ``?``."""
    found: dict[int, dict[str, tuple[int, float]]] = defaultdict(dict)
    for page, y, text in lines:
        match = CASE.search(text)
        if match:
            found[int(match[1])].setdefault(match[2], (page, y))
    out = {}
    for number, case in enumerate(probe.CASES):
        seen = found.get(number, {})
        if "anchor" not in seen or "next" not in seen:
            out[number] = "?"
            continue
        (anchor_page, anchor_y), (next_page, next_y) = seen["anchor"], seen["next"]
        if case.kind == "text":
            cand_page = seen.get("cand", (None,))[0]
            out[number] = "stayed" if cand_page == anchor_page else "moved"
        elif case.kind == "next-page":
            out[number] = {1: "stayed", 2: "moved"}.get(next_page - anchor_page, "?")
        else:
            if next_page == anchor_page:
                out[number] = "fits"  # the next line itself on the anchor's page
                continue
            drop = next_y - anchor_y
            if next_page != anchor_page + 1:
                out[number] = "?"
            else:
                out[number] = "stayed" if abs(drop) < 1 else f"moved {drop:.0f}"
    return out


def run(record: bool) -> int:
    import render_record

    faces: dict = {}
    documents = {}
    for setting in probe.SETTINGS:
        data = probe.build(setting)
        word, word_count = word_lines(data, f"section-end-{setting}")
        model, model_count = model_lines(data, render_record.RecordingFonts(data, faces))
        ends = {}
        for case in probe.END_CASES:
            end_data = probe.build_end(case, setting)
            _, pages = word_lines(end_data, f"section-end-{setting}-end")
            _, ours = model_lines(end_data, render_record.RecordingFonts(end_data, faces))
            ends[case.key] = [pages, ours]
        documents[setting] = {"word": outcomes(word), "model": outcomes(model), "pages": [word_count, model_count],
                              "ends": ends}
    report(documents)
    if record:
        render_record.dump(OBSERVATIONS, {
            "_about": (
                "For each document of tools/make_section_end_probe.py, where Word 16.106 put each case's "
                "candidate (stayed on its anchor's page, or moved, with how far the next line moved in px), "
                "the pages Word made, and for the document-end cases the pages of each; the cases the model "
                "read otherwise at recording time; the numbers of every face the model asked for. "
                "Measurements only; regenerate with tools/read_section_end_probe.py --record."),
            "faces": faces,
            "documents": {s: {"word": {probe.CASES[n].key: v for n, v in d["word"].items()},
                              "pages": d["pages"][0],
                              "ends": {k: v[0] for k, v in d["ends"].items()},
                              "disagree": sorted(probe.CASES[n].key for n, v in d["word"].items()
                                                 if d["model"][n] != v)}
                          for s, d in documents.items()}})
    return 0


def report(documents: dict) -> None:
    for setting, doc in documents.items():
        print(f"== {setting}: Word {doc['pages'][0]} pages, model {doc['pages'][1]}")
        table: dict = defaultdict(dict)
        agree = 0
        for number, case in enumerate(probe.CASES):
            w, m = doc["word"][number], doc["model"][number]
            agree += w == m
            table[(case.family.name, case.kind)][case.slack] = w if w == m else f"{w}/m:{m}"
        print(f"  model agrees {agree} / {len(probe.CASES)}")
        slacks = list(probe.SLACKS)
        print("  " + f"{'family kind':24}" + "".join(f"{s:>16}" for s in slacks))
        for (family, kind), row in table.items():
            print("  " + f"{family + ' ' + kind:24}" + "".join(f"{row.get(s, '-'):>16}" for s in slacks))
        for key, (word, ours) in doc["ends"].items():
            print(f"  end {key:32} Word {word} page(s), model {ours}")


if __name__ == "__main__":
    raise SystemExit(run("--record" in sys.argv[1:]))
