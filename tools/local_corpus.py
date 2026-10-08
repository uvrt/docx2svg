#!/usr/bin/env python3
"""Local corpora: documents scored on one machine, of which git holds no trace.

Some documents may be measured here but not redistributed, and not even named: their
terms allow private use only.  Such a corpus lives in a directory under the gitignored
``scratch/``, and everything about it stays in that directory -- the documents, Word's
lines of them, the scores they are pinned at, and any notes.  What is committed is this
mechanism, which knows no corpus by name, and ``tests/test_local_corpora.py``, which runs
whatever it finds and skips where it finds nothing.

A corpus is a directory holding a manifest, ``local-corpus.json``::

    {
      "observations": "local-observations.json",
      "documents": {
        "<file>.docx": {
          "sha256": "<hex>",               # the file these scores belong to
          "baselines": [exact, scored, out_of_scope],
          "glyphs": [compared, problems],  # glyphs in the resolved face and size
          "breaks": [agree, lines],        # lines the line breaker breaks as Word did
          "skip": "<reason>",              # optional: not scored, and why
          "notes": "..."                   # optional, free text, for people
        }
      }
    }

and the observations file it names, in ``read_baselines.py``'s recording format: every
line Word drew for each document and the metric integers of every face the predictor
asked for -- and, under ``advances``, the advance widths the line breaker asked for
(``read_breaks.py``'s format).  The test scores each listed document from those through
the model's own code (``baselines.predict``, ``baselines.check_glyphs`` and
``breaks.score``), offline.

Usage (needs Word)::

    python tools/local_corpus.py record scratch/<name>          # export, record, report
    python tools/local_corpus.py record scratch/<name> --pin    # ... and pin the scores

``record`` scores every ``*.docx`` in the directory, writes the observations, and adds
new documents to the manifest with their scores; ``--pin`` re-pins documents already
listed.  A document for which Word had to substitute a face (one the document names that
is neither installed nor embedded) is listed with a ``skip`` reason instead of being
pinned: its lines measure substitution, not the model.  A ``skip`` written by hand is
kept.  Exports go through ``tools/oracle.py`` under a neutral name, so the oracle cache
(outside the repository) is keyed by content alone.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(REPO / "src"))

import baselines  # noqa: E402

MANIFEST = "local-corpus.json"
OBSERVATIONS = "local-observations.json"
SCRATCH = REPO / "scratch"


def discover(root: Path = SCRATCH) -> list[Path]:
    """Every corpus directory directly under ``root`` (one holding a manifest)."""
    if not root.is_dir():
        return []
    return sorted(path.parent for path in root.glob(f"*/{MANIFEST}"))


def manifest(corpus: Path) -> dict:
    return json.loads((corpus / MANIFEST).read_text(encoding="utf-8"))


def observations(corpus: Path, listed: dict | None = None) -> dict | None:
    listed = listed or manifest(corpus)
    path = corpus / listed.get("observations", OBSERVATIONS)
    return json.loads(path.read_text(encoding="utf-8")) if path.is_file() else None


def recorded_metrics(recorded: dict):
    from docx2svg.vertical import FaceMetrics

    def metrics(face, bold=False, italic=False):
        found = recorded["faces"].get(f"{face}|{int(bold)}|{int(italic)}")
        return None if found is None else FaceMetrics(*found)
    return metrics


def score(data: bytes, lines: list, metrics) -> tuple[tuple[int, int, int], list[int], list[str]]:
    """(exact, scored, out of scope), [glyphs compared, problems], and the problems."""
    from docx2svg import parse_package

    drawn = [baselines.DrawnLine(*line) for line in lines]
    document = parse_package(data)
    results = baselines.predict(document, drawn, metrics, package=data)
    compared, problems = baselines.check_glyphs(document, drawn, metrics)
    return baselines.summary(results), [compared, len(problems)], problems


def score_breaks(data: bytes, lines: list, metrics, advances) -> list[int]:
    """[lines the model breaks where Word did, lines scored] (``breaks.score``)."""
    import breaks

    from docx2svg import parse_package

    drawn = [baselines.DrawnLine(*line) for line in lines]
    results = breaks.score(parse_package(data), drawn, advances, metrics, package=data)
    summary = breaks.summary(results, justified=None)
    return [summary["agree"], summary["lines"]]


def recorded_advances(recorded: dict):
    import face_advances

    return face_advances.RecordedAdvances(recorded.get("advances", {}))


def missing_faces(faces_asked: dict) -> list[str]:
    """Faces the predictor asked for that this machine neither has nor the document
    embeds -- which Word therefore drew with some other face."""
    return sorted({key.split("|")[0] for key, found in faces_asked.items() if found is None})


def drawn_faces(lines: list) -> list[str]:
    return sorted({run[0] for line in lines for run in line[2]})


# -- recording (needs Word) ------------------------------------------------------------


def record(corpus: Path, *, pin: bool = False, verbose: bool = False) -> int:
    import face_advances
    import face_metrics
    import oracle
    import quartz_pdf
    import read_baselines

    from docx2svg import parse_package

    corpus = corpus.resolve()
    listed = manifest(corpus) if (corpus / MANIFEST).is_file() else {}
    listed.setdefault("observations", OBSERVATIONS)
    documents = listed.setdefault("documents", {})
    faces: dict = {}
    advance_faces: dict = {}
    recorded_lines: dict = {}
    per_document: dict = {}
    try:
        for path in sorted(corpus.glob("*.docx")):
            data = path.read_bytes()
            pdf = oracle.export(data, name="local")
            read_baselines.FACES_USED.clear()
            read_baselines.EMBEDDED.clear()
            read_baselines.EMBEDDED.update(face_metrics.embedded(data))
            lines = baselines.drawn_lines(quartz_pdf.read(pdf))
            document = parse_package(data)
            results = baselines.predict(document, lines, read_baselines.recording_metrics, package=data)
            compared, problems = baselines.check_glyphs(document, lines, read_baselines.recording_metrics)
            asked = dict(read_baselines.FACES_USED)
            faces.update(asked)
            as_json = [[line.page, line.y, [list(run) for run in line.runs]] for line in lines]
            recorded_lines[path.name] = as_json
            summary = list(baselines.summary(results))
            advances = face_advances.RecordingAdvances(face_advances.InstalledAdvances(data), advance_faces)
            broken = score_breaks(data, as_json, read_baselines.recording_metrics, advances)
            missing = missing_faces(asked)
            per_document[path.name] = {
                "baselines": summary, "glyphs": [compared, len(problems)],
                "faces_asked": sorted(asked), "faces_missing": missing,
                "faces_drawn": drawn_faces(as_json),
                "embedded": sorted({key[0] for key in read_baselines.EMBEDDED}),
            }
            print(f"{path.name}: {summary[0]} / {summary[1]} exact ({summary[2]} out of scope);"
                  f" glyphs {compared - len(problems)} / {compared}; line breaks {broken[0]} / {broken[1]}")
            print(f"    faces asked {sorted(asked)}; drawn {drawn_faces(as_json)}")
            if missing:
                print(f"    NOT INSTALLED OR EMBEDDED (Word substituted): {missing}")
            for problem in problems[:10]:
                print(f"    glyph: {problem}")
            for r in results:
                if verbose or r.status == "miss":
                    delta = "" if r.predicted is None else f"{r.predicted - r.observed:+d}"
                    print(f"    p{r.page + 1} b{r.block:<3} {r.status:34} obs {r.observed:5}"
                          f" pred {r.predicted!s:5} {delta:3} {r.text[:30]!r} {'; '.join(r.notes)}")

            entry = documents.get(path.name)
            digest = hashlib.sha256(data).hexdigest()
            if entry is None or pin or entry.get("sha256") != digest:
                entry = {key: value for key, value in (entry or {}).items()
                         if key in ("skip", "notes") and entry.get("sha256") == digest}
                entry.update(sha256=digest, baselines=summary, glyphs=[compared, len(problems)], breaks=broken)
                if missing and "skip" not in entry:
                    entry["skip"] = f"Word substituted faces it has neither installed nor embedded: {missing}"
                documents[path.name] = entry
            elif [entry.get("baselines"), entry.get("glyphs")] != [summary, [compared, len(problems)]] or (
                    entry.get("breaks", broken) != broken):
                print(f"    differs from the pinned {entry.get('baselines')} / {entry.get('glyphs')}"
                      f" / {entry.get('breaks')}; rerun with --pin to accept")
            elif "breaks" not in entry:
                # Pinned before the line breaker existed: its first score is a new pin,
                # not a change to an old one.
                entry["breaks"] = broken
    finally:
        oracle.recover()

    payload = {
        "_about": ("Every line Word drew for each document: [page, baseline in device px, "
                   "[[/BaseFont, drawn size px, text], ...]]; the metric integers of every face "
                   "the predictor asked for; the advance widths the line breaker asked for; and "
                   "which faces each document asked for, lacked "
                   "and was drawn in. Written by tools/local_corpus.py record."),
        "faces": dict(sorted(faces.items())),
        "advances": dict(sorted(advance_faces.items())),
        "documents": recorded_lines,
        "checks": per_document,
    }
    (corpus / listed["observations"]).write_text(
        json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + "\n")
    listed["documents"] = dict(sorted(documents.items()))
    (corpus / MANIFEST).write_text(json.dumps(listed, ensure_ascii=False, indent=2) + "\n")
    print(f"wrote {corpus / MANIFEST} and {corpus / listed['observations']}")
    return 0


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)
    rec = sub.add_parser("record", help="export, score and record every .docx in a corpus directory")
    rec.add_argument("corpus", type=Path)
    rec.add_argument("--pin", action="store_true", help="re-pin documents already in the manifest")
    rec.add_argument("-v", "--verbose", action="store_true")
    sub.add_parser("list", help="the corpora found under scratch/")
    args = parser.parse_args(argv[1:])
    if args.command == "list":
        for corpus in discover():
            print(f"{corpus.relative_to(REPO)}: {len(manifest(corpus).get('documents', {}))} documents")
        return 0
    return record(args.corpus, pin=args.pin, verbose=args.verbose)


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
