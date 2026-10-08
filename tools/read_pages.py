#!/usr/bin/env python3
"""Score ``docx2svg.paginate`` on real documents: page tops placed where Word placed them.

Word's pages come from the lines already recorded for the baseline scorer
(``tests/fixtures/baseline-observations.json``; for the uncommitted ``filesamples`` set,
``scratch/filesamples/baseline-observations.json``).  The model runs on the file alone:
the line breaker, the vertical model and the keeps, with advances and metrics from the
installed faces (or the document's embedded ones), recorded so the tests run offline.

Usage::

    python tools/read_pages.py                    # committed and filesamples documents
    python tools/read_pages.py --record           # ... and write the recordings
    python tools/read_pages.py -v                 # every miss

``--record`` writes ``tests/fixtures/page-advances.json`` (the committed documents) and
``scratch/filesamples/page-advances.json``: the advance widths and face metrics the
paginator asked for.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import baselines  # noqa: E402
import pages  # noqa: E402
import read_baselines  # noqa: E402

from docx2svg import parse_package  # noqa: E402

REPO = HERE.parent
ADVANCES = REPO / "tests" / "fixtures" / "page-advances.json"
SCRATCH_ADVANCES = read_baselines.SCRATCH / "page-advances.json"


def drawn_of(recorded: dict, name: str) -> list[baselines.DrawnLine]:
    return [baselines.DrawnLine(page, y, tuple(tuple(run) for run in runs))
            for page, y, runs in recorded["documents"][name]]


def score(path: Path, recorded: dict, advances, metrics, rules=None) -> pages.Score:
    data = path.read_bytes()
    return pages.score(path.name, parse_package(data), data, drawn_of(recorded, path.name), advances, metrics,
                       rules)


def live(paths: list[Path], recorded: dict, faces: dict, advance_faces: dict, verbose: bool) -> dict:
    """Score ``paths`` against the installed faces, recording what the model asked for."""
    import face_advances
    import face_metrics

    out = {}
    for path in paths:
        data = path.read_bytes()
        embedded = face_metrics.embedded(data)

        def metrics(face, bold=False, italic=False):
            found = face_metrics.metrics(face, bold, italic)
            if found is None:
                key = face.lower()
                found = embedded.get((key, bold, italic)) or embedded.get((key, bold, False)) \
                    or embedded.get((key, False, False))
            faces[f"{face}|{int(bold)}|{int(italic)}"] = None if found is None else baselines.face_integers(found)
            return found

        advances = face_advances.RecordingAdvances(face_advances.InstalledAdvances(data), advance_faces)
        result = score(path, recorded, advances, metrics)
        out[path.name] = result
        print(result.row(), f"(line counts differ: {result.line_count_differs})")
        if verbose:
            for miss in result.misses:
                print("    miss (page, Word's top)", miss)
            for miss in result.conditional_misses:
                print("    conditional (page, Word's top, Word's next, the model's next)", miss)
    return out


def offline(recording: Path, observations: Path, paths: list[Path], rules=None) -> dict:
    """The recorded scores, without Word or fonts."""
    import face_advances
    import probe_documents

    recorded = json.loads(recording.read_text(encoding="utf-8"))
    lines = json.loads(observations.read_text(encoding="utf-8"))
    advances = face_advances.RecordedAdvances(recorded["advances"])
    metrics = probe_documents.recorded_metrics(recorded["faces"])
    return {path.name: score(path, lines, advances, metrics, rules) for path in paths if path.name in lines["documents"]}


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv[1:])
    sets = [(list(read_baselines.RECORDED), read_baselines.OBSERVATIONS, ADVANCES)]
    scratch = sorted(read_baselines.SCRATCH.glob("*.docx"))
    if scratch and read_baselines.SCRATCH_OBSERVATIONS.is_file():
        sets.append((scratch, read_baselines.SCRATCH_OBSERVATIONS, SCRATCH_ADVANCES))
    for paths, observations, target in sets:
        faces: dict = {}
        advance_faces: dict = {}
        recorded = json.loads(observations.read_text(encoding="utf-8"))
        results = live([p for p in paths if p.name in recorded["documents"]], recorded, faces, advance_faces,
                       args.verbose)
        if args.record:
            target.write_text(json.dumps({
                "_about": ("Advance widths (and kern pairs) and the four hhea/typo integers of every face the "
                           "paginator asked for while laying out these documents, from the faces Word lays out "
                           "with. Measurements only; regenerate with tools/read_pages.py --record."),
                "faces": dict(sorted(faces.items())), "advances": dict(sorted(advance_faces.items())),
                "scores": {name: [r.tops_matched, r.word_tops, r.paragraphs_matched, r.paragraphs]
                           for name, r in results.items()},
            }, ensure_ascii=False, separators=(",", ":"), sort_keys=True) + "\n")
            print(f"wrote {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
