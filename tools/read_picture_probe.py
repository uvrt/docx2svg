#!/usr/bin/env python3
"""Measure ``make_picture_probe.py``: the baseline of the line after each picture, from the
model's height of the picture's line.

The stack is the probe's own: the top margin, the anchor and the text line before (each
Calibri 11 pt single, the model's pitch), the picture paragraph's first line as
``docx2svg.paginate.flow`` lays it out (``lines.object_line_height``), its space after,
then the line after, whose baseline ``vertical.baseline_px`` rounds.  Every case is a
line to get exact, per setting.

``--record`` writes ``tests/fixtures/picture-observations.json``.

Usage::

    python tools/read_picture_probe.py [--record]
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_picture_probe as probe  # noqa: E402
import probe_documents  # noqa: E402

OBSERVATIONS = probe_documents.FIXTURES / "picture-observations.json"


def observe(drawn) -> dict[int, list]:
    """Case number -> ``[page, baseline]`` of its ``before`` and ``after`` lines."""
    out: dict[int, dict] = {}
    for page, y, text in drawn:
        words = text.split()
        if len(words) >= 3 and words[0] == "Case" and words[2] in ("before", "after"):
            out.setdefault(int(words[1]), {})[words[2]] = [page, y]
    return out


def predict(setting: str, metrics, advances) -> dict[int, int]:
    """Case number -> the model's baseline of its ``after`` line."""
    from docx2svg import paginate, parse_package
    from docx2svg.vertical import baseline_px, twips_to_px

    document = parse_package(probe.build(setting))
    items = paginate.flow(document, advances, metrics)
    kinds = probe.kinds()
    out = {}
    number = -1
    for index, kind in enumerate(kinds):
        if kind == "anchor":
            number += 1
            top = twips_to_px(1440) + items[index].heights[0].pitch + items[index + 1].heights[0].pitch
        elif kind == "picture":
            item = items[index]
            top += item.heights[0].pitch + twips_to_px(item.after)
        elif kind == "after":
            face = metrics("Calibri")
            out[number] = baseline_px(top, face, 22)
    return out


def score(documents: dict, metrics, advances) -> dict[str, tuple[int, int, list]]:
    out = {}
    for setting in probe.SETTINGS:
        seen = observe(documents[setting])
        model = predict(setting, metrics, advances)
        cases = [(n, v) for n, v in seen.items() if "after" in v and "before" in v and v["after"][0] == v["before"][0]]
        wrong = [(probe.CASES[n].key, v["after"][1] - model[n]) for n, v in cases if v["after"][1] != model[n]]
        out[setting] = (len(cases) - len(wrong), len(cases), wrong)
    return out


def offline() -> dict:
    import pages

    data = json.loads(OBSERVATIONS.read_text())
    advances, metrics = pages.recorded(data)
    return score(data["documents"], metrics, advances)


def main(argv: list[str]) -> int:
    import pages

    faces: dict = {}
    advance_faces: dict = {}
    documents = {}
    for setting in probe.SETTINGS:
        data = probe.build(setting)
        drawn, advances, metrics = pages.record_probe(f"picture-{setting}", data, faces, advance_faces)
        documents[setting] = [[line.page, line.y, line.text] for line in drawn if line.text.startswith("Case")]
    # Ask the model once with the recorders, so the recording holds what it needs.
    for setting, (exact, count, wrong) in score(documents, metrics, advances).items():
        print(f"{setting:6} {exact} / {count}  {wrong[:6]}")
    if "--record" in argv[1:]:
        pages.write_recording(OBSERVATIONS, (
            "For each document of tools/make_picture_probe.py, Word 16.106's text lines as [page, baseline in "
            "device px, text]; the face metrics and advance widths the model asked for. Measurements only; "
            "regenerate with tools/read_picture_probe.py --record."), faces, advance_faces, documents)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
