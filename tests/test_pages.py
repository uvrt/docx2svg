"""``docx2svg.paginate`` on the real documents, offline: page tops placed where Word placed
them, with nothing taken from Word's export (ROADMAP.md, "Phase 4 -- measured").

Word's pages are the lines recorded for the baseline scorer; the model's advances and
metrics are what ``tools/read_pages.py --record`` recorded from the installed faces.  The
``filesamples`` documents are not committed (no licence): their test skips where
``scratch/filesamples`` is absent.

Every Word page top the model does not place is one it never reached: the paginator stops
at the first table (Phase 6) or floating drawing, rather than guess its height.  (Tables
are laid out since ROADMAP.md "Tables -- measured"; the advances recorded here predate
them and lack the cells' faces, so this paginator still stops at the first table, until
the line matcher can score a table row's cells -- see that section.)
"""

from __future__ import annotations

import hashlib
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_baselines  # noqa: E402
import read_pages  # noqa: E402

#: name -> (page tops placed where Word placed them, Word's page tops, of them never
#: reached, paragraphs on Word's pages, paragraphs scored, where the model stopped)
COMMITTED = {
    "style-document.docx": (1, 1, 0, 30, 30, None),
    # Phase 0's fixture: 70 single-line paragraphs, and Word put the break after L32.
    "layout-sweep.docx": (2, 2, 0, 88, 88, None),
    "sample-blank.docx": (0, 0, 0, 0, 0, None),
    # 17 manual page breaks, 18 inline pictures, 35 page tops.
    "sample-long.docx": (35, 35, 0, 127, 127, None),
    "sample-resume.docx": (0, 0, 0, 19, 19, None),
    "sample-simple.docx": (0, 1, 1, 16, 16, "b16 (table)"),
    # Two natural breaks, then a table at the end.
    "sample-10pages.docx": (2, 2, 0, 36, 36, "b36 (table)"),
    "sample-1page.docx": (0, 0, 0, 3, 3, None),
    "sample-5pages.docx": (0, 1, 1, 14, 14, "b14 (table)"),
    "sample-with-images.docx": (0, 0, 0, 8, 8, None),
    "sample-with-table.docx": (0, 0, 0, 3, 3, "b3 (table)"),
}

FILESAMPLES = {
    "sample1.docx": (2, 7, 5, 17, 17, "b30 (table)"),
    "sample2.docx": (0, 0, 0, 5, 5, None),
    # Text beside two floating drawings (wrapSquare), laid out: its page top reached, every
    # paragraph on Word's page, until a table in its two-column section.
    "sample3.docx": (1, 1, 0, 26, 26, "b26 (table)"),
    # 175 pages: 167 natural page tops (104 inside a paragraph), a manual break, 39
    # paragraphs of inline pictures.
    "sample4.docx": (168, 168, 0, 651, 651, None),
}


def _row(score):
    return (score.tops_matched, score.word_tops, score.tops_unreached, score.paragraphs_matched,
            score.paragraphs, score.stopped)


def test_committed_documents():
    paths = [path for path in read_baselines.RECORDED if path.name in COMMITTED]
    scores = read_pages.offline(read_pages.ADVANCES, read_baselines.OBSERVATIONS, paths)
    assert {name: _row(score) for name, score in scores.items()} == COMMITTED


def test_filesamples():
    paths = [read_baselines.SCRATCH / name for name in FILESAMPLES]
    if not all(path.is_file() for path in paths) or not read_pages.SCRATCH_ADVANCES.is_file():
        pytest.skip("scratch/filesamples is absent; see tests/test_filesamples.py")
    import test_filesamples

    for path in paths:
        assert hashlib.sha256(path.read_bytes()).hexdigest() == test_filesamples.SHA256[path.name]
    scores = read_pages.offline(read_pages.SCRATCH_ADVANCES, read_baselines.SCRATCH_OBSERVATIONS, paths)
    assert {name: _row(score) for name, score in scores.items()} == FILESAMPLES
