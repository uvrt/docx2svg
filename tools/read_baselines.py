#!/usr/bin/env python3
"""Score ``baselines.predict`` on whole documents: baselines exact out of baselines scored.

Usage::

    python tools/read_baselines.py DOCX [DOCX...] [--collapse sum|max] [--before-at-top] [-v] [--record]
    python tools/read_baselines.py --record-scratch      # the uncommitted filesamples set

Each document is exported through ``tools/oracle.py`` (cached by content).  Every drawn
line is either scored (``exact`` / ``miss``) or reported out of scope with its reason.
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
import face_metrics  # noqa: E402
import oracle  # noqa: E402
import quartz_pdf  # noqa: E402

from docx2svg import parse_package  # noqa: E402

REPO = HERE.parent
OBSERVATIONS = REPO / "tests" / "fixtures" / "baseline-observations.json"
#: The documents whose baselines are recorded for ``tests/test_baselines.py``.
RECORDED = (
    REPO / "tests" / "fixtures" / "style-document.docx",
    REPO / "tests" / "fixtures" / "layout-sweep.docx",
    *sorted((REPO / "tests" / "fixtures" / "samplelib").glob("*.docx")),
    *sorted((REPO / "tests" / "fixtures" / "wordto").glob("*.docx")),
)
#: Third-party documents with no licence to redistribute them (tests/fixtures/filesamples/
#: PROVENANCE.md).  They live in the gitignored ``scratch/``, and so do their
#: observations: the drawn lines are the documents' text.
SCRATCH = REPO / "scratch" / "filesamples"
SCRATCH_OBSERVATIONS = SCRATCH / "baseline-observations.json"
FACES_USED: dict[str, list[int] | None] = {}


#: Faces embedded in the document being scored (``face_metrics.embedded``).  Used only
#: for a face this machine does not have installed, which is when Word draws them.
EMBEDDED: dict = {}


def recording_metrics(face, bold=False, italic=False):
    """``face_metrics.metrics``, remembering every face asked for (the integers only).

    A face that is not installed is looked up among the document's embedded faces.
    """
    found = face_metrics.metrics(face, bold, italic)
    if found is None:
        key = face.lower()
        found = (EMBEDDED.get((key, bold, italic)) or EMBEDDED.get((key, bold, False))
                 or EMBEDDED.get((key, False, False)))
    FACES_USED[f"{face}|{int(bold)}|{int(italic)}"] = (
        None if found is None else baselines.face_integers(found)
    )
    return found


def score(path: Path, *, collapse: str = "max", before_at_top: bool = False, verbose: bool = False,
          pdf: Path | None = None):
    data = path.read_bytes()
    if pdf is None and path.name == "layout-sweep.docx":
        pdf = oracle.ORACLE_DIR / "layout-sweep.pdf"  # Phase 0's export, reused by name
    pdf = pdf or oracle.export(data, name=path.stem)
    document = parse_package(data)
    EMBEDDED.clear()
    EMBEDDED.update(face_metrics.embedded(data))
    drawn = baselines.drawn_lines(quartz_pdf.read(pdf))
    results = baselines.predict(document, drawn, recording_metrics,
                                collapse=collapse, before_at_page_top=before_at_top,
                                package=data)
    RECORDED_LINES[path.name] = [[line.page, line.y, [list(run) for run in line.runs]] for line in drawn]
    compared, problems = baselines.check_glyphs(document, drawn, recording_metrics)
    GLYPHS[path.name] = [compared, len(problems)]
    exact, scored, excluded = baselines.summary(results)
    print(f"{path.name:32} {exact:4} / {scored:<4} exact   ({excluded} lines out of scope);"
          f"  glyphs {compared - len(problems)} / {compared} in the resolved face and size")
    for problem in problems[:10]:
        print(f"    glyph: {problem}")
    for r in results:
        if verbose or r.status == "miss":
            delta = "" if r.predicted is None else f"{r.predicted - r.observed:+d}"
            print(f"    p{r.page + 1} b{r.block:<3} {r.status:34} obs {r.observed:5} pred {r.predicted!s:5} {delta:3}"
                  f" {r.text[:30]!r} {'; '.join(r.notes)}")
    return results


RECORDED_LINES: dict[str, list] = {}
GLYPHS: dict[str, list[int]] = {}


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", nargs="*", type=Path)
    parser.add_argument("--record", action="store_true",
                        help=f"score the recorded set and write {OBSERVATIONS.name}")
    parser.add_argument("--record-scratch", action="store_true",
                        help=f"score scratch/filesamples/*.docx and write {SCRATCH_OBSERVATIONS.name} there")
    parser.add_argument("--collapse", default="max", choices=("sum", "max"))
    parser.add_argument("--before-at-top", action="store_true",
                        help="keep space before on the first paragraph of a page (refuted)")
    parser.add_argument("--pdf", type=Path)
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv[1:])
    paths = (list(RECORDED) if args.record
             else sorted(SCRATCH.glob("*.docx")) if args.record_scratch else args.docx)
    summary = {}
    for path in paths:
        results = score(path, collapse=args.collapse, before_at_top=args.before_at_top,
                        verbose=args.verbose, pdf=args.pdf)
        summary[path.name] = list(baselines.summary(results))
    if args.record or args.record_scratch:
        target = OBSERVATIONS if args.record else SCRATCH_OBSERVATIONS
        payload = {
            "_about": (
                "Every line Word 16.106 drew for each document: [page, baseline in device px, "
                "[[/BaseFont, drawn size px, text], ...]]; and the four hhea/typo integers of every face the predictor asked for "
                "(upm, ascent, descent, lineGap). Measurements only; regenerate with "
                "tools/read_baselines.py --record."
            ),
            "faces": dict(sorted(FACES_USED.items())),
            "documents": RECORDED_LINES,
            "scores": summary,
            "glyphs": GLYPHS,
        }
        if args.record_scratch:
            import hashlib

            payload["sha256"] = {path.name: hashlib.sha256(path.read_bytes()).hexdigest() for path in paths}
        target.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + "\n")
        print(f"wrote {target} ({target.stat().st_size} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
