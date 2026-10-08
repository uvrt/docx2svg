#!/usr/bin/env python3
"""Measure ``make_footnote_probe.py``: the slack a candidate line needs when notes are due.

For every case, whether the candidate is drawn on its anchor's page, and where its notes
were drawn (their pages, and the baseline of each note line and of the separator's page).
Each family's threshold is the smallest slack that kept the candidate on its page.

``--record`` writes ``tests/fixtures/footnote-observations.json``.

Usage::

    python tools/read_footnote_probe.py [--record] [-v]
"""

from __future__ import annotations

import json
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_footnote_probe as probe  # noqa: E402
import probe_documents  # noqa: E402

OBSERVATIONS = probe_documents.FIXTURES / "footnote-observations.json"


def observe(lines) -> dict[int, dict]:
    """Case number -> ``{"fits": bool, "notes": [(page offset, baseline, text)]}``."""
    anchor, candidate, notes, follow = {}, {}, defaultdict(list), defaultdict(list)
    for page, y, text in lines:
        words = text.split()
        if len(words) >= 3 and words[0] == "Case":
            number = int(words[1])
            if words[2].startswith("anchor"):
                anchor[number] = page
            elif words[2].startswith("cand"):
                candidate[number] = page
            elif words[2] == "follow":
                follow[number].append(page)
        for word in words:
            core = word.lstrip("0123456789")
            if core.startswith("n") and core.count(".") == 2:
                number = int(core[1:].split(".")[0])
                notes[number].append((page, y, core))
    out = {}
    for number in anchor:
        if number in candidate:
            base = anchor[number]
            out[number] = {"fits": candidate[number] == base,
                           "notes": [(p - base, y, t) for p, y, t in notes[number]],
                           "follow": [sum(1 for p in follow[number] if p - base == k) for k in range(3)]}
    return out


def report(observed: dict[str, dict], verbose: bool = False) -> None:
    for setting, cases in observed.items():
        print(f"== {setting}")
        for family in probe.FAMILIES:
            numbers = [n for n, c in enumerate(probe.CASES) if c.family is family and n in cases]
            if family.follow:
                for n in numbers:
                    print(f"  {family.name:22} slack {probe.CASES[n].slack}: {'fits' if cases[n]['fits'] else 'next'};"
                          f" follow lines per page {cases[n]['follow']}; notes {cases[n]['notes']}")
                continue
            for threshold in family.thresholds or (None,):
                chosen = [n for n in numbers if threshold is None or abs(probe.CASES[n].slack - threshold) <= 1024]
                base = threshold or 0
                held = [probe.CASES[n].slack - base for n in chosen if cases[n]["fits"]]
                lost = [probe.CASES[n].slack - base for n in chosen if not cases[n]["fits"]]
                where = f"about {threshold}" if threshold is not None else "coarse"
                print(f"  {family.name:22} {where:>14}: fits from {min(held) if held else '-':>8},"
                      f" not at {max(lost) if lost else '-':>8}"
                      f"  (non-monotone: {sum(1 for s in lost if held and s > min(held))})")
            if verbose:
                for n in numbers:
                    print(f"      {probe.CASES[n].slack:>8} {'fits' if cases[n]['fits'] else 'next'} {cases[n]['notes']}")


def model(setting: str, advances, metrics, rules=None) -> dict[int, dict]:
    """Case number -> ``{"fits": ..., "follow": [...]}`` by ``docx2svg.paginate``."""
    from docx2svg import paginate, parse_package

    data = probe.build(setting)
    document = parse_package(data)
    items = paginate.flow(document, advances, metrics)
    found = paginate.drawn_pages(items, paginate.paginate(document, items, rules))
    body, _ = probe.layout()
    anchor, out = {}, {}
    for block, (number, role, _) in enumerate(body):
        if role == "anchor":
            anchor[number] = found[block][0]
        elif role == "candidate":
            out[number] = {"fits": found[block][0] == anchor[number], "follow": [0, 0, 0]}
        elif role == "follow":
            out[number]["follow"] = [sum(1 for p in found[block] if p - anchor[number] == k) for k in range(3)]
    return out


def agreement(observed: dict, predicted: dict) -> tuple[int, int]:
    """Cases whose candidate (and following lines) Word and the model put on the same page."""
    same = [observed[n]["fits"] == predicted[n]["fits"] and (
        not probe.CASES[n].family.follow or observed[n]["follow"] == predicted[n]["follow"])
        for n in observed]
    return sum(same), len(same)


def offline(rules=None) -> dict[str, tuple[int, int]]:
    """Per setting, cases the model agrees with Word on, from the recording."""
    import pages

    data = json.loads(OBSERVATIONS.read_text())
    advances, metrics = pages.recorded(data)
    return {setting: agreement(observe(lines), model(setting, advances, metrics, rules))
            for setting, lines in data["documents"].items()}


def main(argv: list[str]) -> int:
    documents = {}
    for setting in probe.SETTINGS:
        data = probe.build(setting)
        drawn = probe_documents.measure(f"footnote-{setting}", data)
        documents[setting] = [[line.page, line.y, line.text] for line in drawn
                              if line.text.startswith("Case") or " n" in f" {line.text}" or line.text[:1].isdigit()]
    observed = {setting: observe(lines) for setting, lines in documents.items()}
    report(observed, "-v" in argv[1:])
    import face_advances
    import pages

    faces: dict = {}
    advance_faces: dict = {}
    metrics = probe_documents.RecordingMetrics()
    metrics.faces = faces
    for setting in probe.SETTINGS:
        advances = face_advances.RecordingAdvances(face_advances.InstalledAdvances(), advance_faces)
        predicted = model(setting, advances, metrics)
        agree, total = agreement(observed[setting], predicted)
        wrong = [probe.CASES[n].key for n in observed[setting]
                 if (observed[setting][n]["fits"], observed[setting][n]["follow"])
                 != (predicted[n]["fits"], predicted[n]["follow"] if probe.CASES[n].family.follow
                     else observed[setting][n]["follow"])]
        print(f"model {setting}: {agree} / {total} cases; wrong: {wrong[:12]}{' ...' if len(wrong) > 12 else ''}")
    if "--record" in argv[1:]:
        payload = {
            "_about": ("Word 16.106's lines of each document of tools/make_footnote_probe.py that hold a case's "
                       "anchor, candidate, following or note text, as [page, baseline in device px, text]; the "
                       "face metrics and advance widths the model asked for. Measurements only; regenerate with "
                       "tools/read_footnote_probe.py --record."),
            "faces": dict(sorted(faces.items())), "advances": dict(sorted(advance_faces.items())),
            "documents": documents}
        OBSERVATIONS.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":"), sort_keys=True) + "\n")
        print(f"wrote {OBSERVATIONS}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
