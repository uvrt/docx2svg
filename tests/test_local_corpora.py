"""Local corpora: whatever ``scratch/*/local-corpus.json`` lists, scored offline.

Documents whose terms allow private measurement but not redistribution are kept, with
everything about them, in a directory under the gitignored ``scratch/``, described by a
manifest there (``tools/local_corpus.py`` has the format and writes it).  This module
knows no corpus: it scores every document every manifest lists through the model's own
code (``baselines.predict``, ``baselines.check_glyphs`` and, where pinned, ``breaks.score``)
against the lines Word drew,
recorded beside it, and holds each to the scores its manifest pins.  Where no manifest
exists -- every machine but one -- the per-document test skips, and the mechanism is
still exercised on a corpus made in a temporary directory from a committed fixture.
"""

from __future__ import annotations

import hashlib
import json
import shutil
from pathlib import Path

import pytest

import local_corpus

FIXTURES = Path(__file__).resolve().parent / "fixtures"


def _cases():
    cases = []
    for corpus in local_corpus.discover():
        for name in sorted(local_corpus.manifest(corpus).get("documents", {})):
            cases.append(pytest.param(corpus, name, id=f"{corpus.name}/{name}"))
    return cases or [pytest.param(None, None, id="none",
                                  marks=pytest.mark.skip(reason="no scratch/*/local-corpus.json"))]


def check(corpus: Path, name: str) -> None:
    listed = local_corpus.manifest(corpus)
    entry = listed["documents"][name]
    if entry.get("skip"):
        pytest.skip(entry["skip"])
    path = corpus / name
    if not path.is_file():
        pytest.skip(f"{path} is listed but absent")
    data = path.read_bytes()
    assert hashlib.sha256(data).hexdigest() == entry["sha256"], f"{name} is not the file these scores belong to"
    recorded = local_corpus.observations(corpus, listed)
    if recorded is None or name not in recorded["documents"]:
        pytest.skip(f"{name} has no recorded lines; run tools/local_corpus.py record {corpus}")
    summary, glyphs, _ = local_corpus.score(data, recorded["documents"][name],
                                            local_corpus.recorded_metrics(recorded))
    assert list(summary) == entry["baselines"]
    assert glyphs == entry["glyphs"]
    if "breaks" in entry:
        assert local_corpus.score_breaks(data, recorded["documents"][name], local_corpus.recorded_metrics(recorded),
                                         local_corpus.recorded_advances(recorded)) == entry["breaks"]


@pytest.mark.parametrize(("corpus", "name"), _cases())
def test_local_corpus_document(corpus, name):
    check(corpus, name)


def _make_corpus(root: Path, **entry_overrides) -> Path:
    """A corpus of one committed fixture, with its recorded lines, laid out as a local
    one is."""
    name = "sample-simple.docx"
    data = (FIXTURES / "samplelib" / name).read_bytes()
    recorded = json.loads((FIXTURES / "baseline-observations.json").read_text(encoding="utf-8"))
    corpus = root / "example"
    corpus.mkdir(parents=True)
    shutil.copyfile(FIXTURES / "samplelib" / name, corpus / name)
    (corpus / local_corpus.OBSERVATIONS).write_text(json.dumps(
        {"faces": recorded["faces"], "documents": {name: recorded["documents"][name]}}))
    entry = {"sha256": hashlib.sha256(data).hexdigest(), "baselines": [22, 22, 7],
             "glyphs": recorded["glyphs"][name], **entry_overrides}
    (corpus / local_corpus.MANIFEST).write_text(json.dumps(
        {"observations": local_corpus.OBSERVATIONS, "documents": {name: entry}}))
    return corpus


def test_the_mechanism_on_a_corpus_of_committed_fixtures(tmp_path):
    corpus = _make_corpus(tmp_path)
    assert local_corpus.discover(tmp_path) == [corpus]
    check(corpus, "sample-simple.docx")


def test_a_moved_score_fails(tmp_path):
    corpus = _make_corpus(tmp_path, baselines=[21, 22, 7])
    with pytest.raises(AssertionError):
        check(corpus, "sample-simple.docx")


def test_another_file_fails(tmp_path):
    corpus = _make_corpus(tmp_path, sha256="0" * 64)
    with pytest.raises(AssertionError, match="not the file"):
        check(corpus, "sample-simple.docx")


def test_a_listed_skip_skips(tmp_path):
    corpus = _make_corpus(tmp_path, skip="Word substituted a face")
    with pytest.raises(pytest.skip.Exception, match="substituted"):
        check(corpus, "sample-simple.docx")


def test_no_corpus_is_nothing(tmp_path):
    assert local_corpus.discover(tmp_path / "absent") == []
    assert local_corpus.discover(tmp_path) == []
