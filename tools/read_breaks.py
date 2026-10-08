#!/usr/bin/env python3
"""Score ``docx2svg.linebreak`` on whole documents: lines that break where Word breaks them.

Usage::

    python tools/read_breaks.py DOCX [DOCX...] [-v]
    python tools/read_breaks.py --record          # the committed documents
    python tools/read_breaks.py --record-scratch  # scratch/filesamples, recorded there

Each document is exported through ``tools/oracle.py`` (cached by content).  Advances come
from the installed faces (``face_advances``), or the document's embedded ones.  ``-v``
prints every paragraph whose lines do not all agree, Word's lines against the model's.
``--record`` writes the advances every scored paragraph asked for to
``tests/fixtures/break-advances.json`` (the drawn lines are already in
``baseline-observations.json``); ``--record-scratch`` writes them beside the scratch
documents.
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
import breaks  # noqa: E402
import face_advances  # noqa: E402
import face_metrics  # noqa: E402
import read_baselines  # noqa: E402

from docx2svg import parse_package  # noqa: E402

REPO = HERE.parent
ADVANCES = REPO / "tests" / "fixtures" / "break-advances.json"
SCRATCH_ADVANCES = read_baselines.SCRATCH / "break-advances.json"


def report(name: str, document, results, verbose: bool) -> dict:
    plain = breaks.summary(results)
    justified = breaks.summary(results, justified=True)
    print(f"{name:34} {plain['agree']:5} / {plain['lines']:<5} lines agree"
          f" (conditional {plain['conditional']}); justified {justified['agree']} / {justified['lines']};"
          f" {plain['out']} paragraphs out of scope")
    if verbose:
        for r in results:
            if r.status != "scored":
                if not r.status.endswith(("table", "drawing")):
                    print(f"    b{r.block:<4} {r.status}  {r.text!r}")
                continue
            if r.agree == r.lines:
                continue
            paragraph = baselines.blocks(document)[r.block].paragraph
            text = "".join(ch for ch in paragraph.text if not ch.isspace())
            print(f"    b{r.block:<4} p{r.page + 1} {'justified ' if r.justified else ''}{r.agree}/{r.lines}")
            for number in range(max(len(r.observed), len(r.predicted))):
                o = r.observed[number] if number < len(r.observed) else None
                p = r.predicted[number] if number < len(r.predicted) else None
                mark = "  " if o == p else "!!"
                cond = r.conditional[number] if number < len(r.conditional) else None
                end = (lambda s: text[max(0, s[1] - 12):s[1]] if s else "")
                print(f"      {mark} word ...{end(o)!r:16} model ...{end(p)!r:16} conditional={cond}")
    return {"all": plain, "justified": justified}


def score(path: Path, advances_faces: dict, verbose: bool = False, pdf: Path | None = None):
    import oracle
    import quartz_pdf

    data = path.read_bytes()
    if pdf is None and path.name == "layout-sweep.docx":
        pdf = oracle.ORACLE_DIR / "layout-sweep.pdf"
    pdf = pdf or oracle.export(data, name=path.stem)
    document = parse_package(data)
    read_baselines.EMBEDDED.clear()
    read_baselines.EMBEDDED.update(face_metrics.embedded(data))
    drawn = baselines.drawn_lines(quartz_pdf.read(pdf))
    advances = face_advances.RecordingAdvances(face_advances.InstalledAdvances(data), advances_faces)
    results = breaks.score(document, drawn, advances, read_baselines.recording_metrics, package=data)
    return report(path.name, document, results, verbose)


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", nargs="*", type=Path)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--record-scratch", action="store_true")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv[1:])
    paths = (list(read_baselines.RECORDED) if args.record
             else sorted(read_baselines.SCRATCH.glob("*.docx")) if args.record_scratch else args.docx)
    faces: dict = {}
    scores = {}
    for path in paths:
        scores[path.name] = score(path, faces, args.verbose)
    if args.record or args.record_scratch:
        target = ADVANCES if args.record else SCRATCH_ADVANCES
        target.write_text(json.dumps({
            "_about": ("Advance widths (and kern pairs) of every character the line breaker asked for, "
                       "per face, in font units, read from the faces Word lays out with. Measurements "
                       "only; regenerate with tools/read_breaks.py --record."),
            "faces": dict(sorted(faces.items())),
            "scores": scores,
        }, ensure_ascii=False, separators=(",", ":"), sort_keys=True) + "\n")
        print(f"wrote {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
