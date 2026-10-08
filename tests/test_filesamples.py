"""The four ``filesamples.com`` documents -- **not committed**, so these tests skip here.

filesamples.com grants no licence for its files (``tests/fixtures/filesamples/
PROVENANCE.md``), so neither the documents nor Word's lines of them (which are their
text) are in the repository.  Put the four ``.docx`` files in ``scratch/filesamples/``
and run ``tools/read_baselines.py --record-scratch`` (which needs Word) to write
``scratch/filesamples/baseline-observations.json``; from then on these tests run offline,
like ``test_baselines.py``.  Every test skips where either is absent, which is every
machine but one.

They are the evidence for ROADMAP.md, "Nine more real documents": the scores pinned here
are the model's.  Findings 1-4 of that section are now rules of the model, each adopted
on a probe of its own ("Findings 1-4, probed and adopted"), and so is finding 5
("Superscripts, subscripts, position and run borders -- measured"); what is left is
findings 6 and 7.
"""

from __future__ import annotations

import hashlib
import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
SCRATCH = REPO / "scratch" / "filesamples"
sys.path.insert(0, str(REPO / "tools"))

import baselines  # noqa: E402

from docx2svg import parse_package  # noqa: E402
from docx2svg.vertical import FaceMetrics  # noqa: E402

#: The files these scores belong to (downloaded 2026-09-24).
SHA256 = {
    "sample1.docx": "269329fc7ae54b3f289b3ac52efde387edc2e566ef9a48d637e841022c7e0eab",
    "sample2.docx": "40258ea32f3175c3a91cab65dd6a7eddeb94180edd186ad882ca51b8b0dfa7b4",
    "sample3.docx": "58211ac149c4b05110cb95959f0f5f529333af5aea060d5a5baf096f531848b8",
    "sample4.docx": "7c120af503a7b2c72c756b11dd42dc6d666e4756b94b14f0b7b65c181123434e",
}

#: (exact, scored, out of scope) under the model as it stands.
EXPECTED = {
    # Word 12 (no compatibility mode): every Heading 1 has w:pageBreakBefore and 24 pt
    # before, which Word keeps below mode 15 (14 / 66 before the model knew it), and list
    # labels whose descent exceeds Ubuntu's, which Word leaves out (30 / 66 before).  A run
    # with w:bdr grows its line by its border (46 / 66 before; finding 5 -- the "+3 px"
    # ascribed to the super/subscript line was the bordered line's, below it).  What is
    # left is mode 12's: its Ubuntu Mono lines (finding 7) and what they carry, and two
    # lines 1 px off after lines of Ubuntu's other styles.  In mode 14 or 15 it is exact.
    "sample1.docx": (52, 66, 145),
    # Word 15 writing compatibility mode 14: w:beforeAutospacing/afterAutospacing, 14 pt
    # whatever is stated (6 / 10 before the cascade read them).
    "sample2.docx": (10, 10, 1),
    "sample3.docx": (26, 26, 33),
    # 175 pages, every break natural; 104 page tops continue a paragraph.
    "sample4.docx": (5551, 5551, 700),
}

#: Glyphs drawn otherwise than resolved, and why (none is in an in-scope line):
#: sample1 -- 15 where the checker loses its place in the line under the first table (the
#: table's 42 px glyphs set against the body line); its 60 table-cell glyphs, drawn at 46
#: and 42 px against Normal's 50 px, are resolved at Word's size since finding 6 was
#: settled (resolve.cascade.legacy_table_size: modes 12 and 14 give a 12 pt paragraph the
#: table style's size); sample3 -- one glyph the checker loses where two text columns
#: interleave.  (sample1's 12 super/subscript glyphs, drawn at 33 px, were the other 12 of
#: 87 until resolve.script_half_points: Ubuntu's OS/2 script size, 0.65 of 24 half points.)
GLYPH_PROBLEMS = {"sample1.docx": 15, "sample2.docx": 0, "sample3.docx": 1, "sample4.docx": 0}


def _data():
    path = SCRATCH / "baseline-observations.json"
    if not path.is_file():
        pytest.skip(f"no {path}; see this module's docstring")
    return json.loads(path.read_text())


def _document(name: str) -> bytes:
    path = SCRATCH / name
    if not path.is_file():
        pytest.skip(f"no {path}; see this module's docstring")
    data = path.read_bytes()
    assert hashlib.sha256(data).hexdigest() == SHA256[name], f"{name} is not the file these scores belong to"
    return data


def _metrics(recorded):
    def metrics(face, bold=False, italic=False):
        found = recorded["faces"].get(f"{face}|{int(bold)}|{int(italic)}")
        return None if found is None else FaceMetrics(*found)
    return metrics


def _score(name: str):
    data = _document(name)
    recorded = _data()
    if name not in recorded["documents"]:
        pytest.skip(f"{name} was not recorded; rerun tools/read_baselines.py --record-scratch")
    metrics = _metrics(recorded)
    drawn = [baselines.DrawnLine(*line) for line in recorded["documents"][name]]
    return baselines.predict(parse_package(data), drawn, metrics, package=data)


@pytest.mark.parametrize("name", sorted(EXPECTED))
def test_filesamples_baselines(name):
    assert baselines.summary(_score(name)) == EXPECTED[name]


@pytest.mark.parametrize("name", sorted(GLYPH_PROBLEMS))
def test_filesamples_glyphs(name):
    data = _document(name)
    recorded = _data()
    drawn = [baselines.DrawnLine(*line) for line in recorded["documents"][name]]
    compared, problems = baselines.check_glyphs(parse_package(data), drawn, _metrics(recorded))
    assert len(problems) == GLYPH_PROBLEMS[name]
    assert [compared, len(problems)] == recorded["glyphs"][name]


def test_word_draws_the_face_the_document_embeds():
    """sample1 embeds Ubuntu (not installed here) and Word draws it; the four integers
    come from the embedded, obfuscated part (ECMA-376 17.8.1), read in memory, and so do
    the OS/2 script sizes (0.65 em) that sample1's superscripts are drawn at."""
    pytest.importorskip("fontTools")
    import face_metrics

    found = face_metrics.embedded(_document("sample1.docx"))
    assert found[("ubuntu", False, False)] == FaceMetrics(1000, 932, 189, 28, 650, 650)
    assert found[("ubuntu mono", False, False)] == FaceMetrics(1000, 830, 170, 0, 650, 650)


#: Line breaking (ROADMAP.md, "Phase 3 -- measured"): (lines that agree, lines scored).
#: sample1's three are one paragraph's: its Ubuntu Mono run, which Word in compatibility
#: mode 12 advances by Calibri's widths although it draws the embedded Ubuntu Mono (and
#: in modes 14 and 15 by Ubuntu Mono's own) -- finding 7's horizontal half.
BREAKS = {"sample1.docx": (112, 115), "sample2.docx": (10, 10), "sample3.docx": (36, 36),
          "sample4.docx": (6212, 6212)}


@pytest.mark.parametrize("name", sorted(BREAKS))
def test_filesamples_line_breaks(name):
    import breaks
    import face_advances

    data = _document(name)
    recorded = _data()
    path = SCRATCH / "break-advances.json"
    if not path.is_file():
        pytest.skip(f"no {path}; run tools/read_breaks.py --record-scratch")
    advances = face_advances.RecordedAdvances(json.loads(path.read_text())["faces"])
    drawn = [baselines.DrawnLine(*line) for line in recorded["documents"][name]]
    results = breaks.score(parse_package(data), drawn, advances, _metrics(recorded), package=data)
    scored = [r for r in results if r.status == "scored"]
    assert (sum(r.agree for r in scored), sum(r.lines for r in scored)) == BREAKS[name]
