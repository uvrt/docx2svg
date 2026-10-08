"""Shared fixtures.

Two things are arranged here.  The source tree goes on ``sys.path`` so the suite runs on
a bare checkout without an install, and the oracle is made *optional*: every test that
needs Microsoft Word or its PDF output skips cleanly when either is absent, so this suite
passes on Linux CI, on a machine without Office, and on a fresh clone.

That split is deliberate and it is the same one the sibling ``pptx2svg`` draws.  Tests
that check *our own* reading of a document run everywhere.  Tests that check what Word
does run only where Word is, and are the only tests that can say we are wrong rather than
merely inconsistent.
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

FIXTURES = REPO / "tests" / "fixtures"

sys.path.insert(0, str(REPO / "tools"))
#: Word is sandboxed; the oracle lives in the Office group container, defined once in
#: tools/oracle.py.  See the fourth note in the AppleScript header.
from oracle import ORACLE_DIR, SCRIPT, WordBusy, run  # noqa: E402


def pytest_addoption(parser) -> None:
    parser.addoption(
        "--run-slow",
        action="store_true",
        default=False,
        help="also run the tests marked slow (the raster fidelity scores: minutes, and Word's exports)",
    )
    parser.addoption(
        "--pdfium",
        action="store_true",
        default=False,
        help=(
            "also run the tests marked pdfium, which score against the fidelity harness's old "
            "instrument (Word's PDF drawn by pdfium).  Scores default to the svg truth, as "
            "tools/fidelity.py does; this is the suite's `--truth both`."
        ),
    )
    parser.addoption(
        "--exact",
        action="store_true",
        default=False,
        help=(
            "score at the exact glyph size too, in the fidelity tests that render ours.  They "
            "default to the device glyph size, as tools/fidelity.py does; this is the suite's "
            "`--glyph-size both`."
        ),
    )
    parser.addoption(
        "--update-snapshots",
        action="store_true",
        default=False,
        help="rewrite the committed SVG snapshots under tests/vrt/ (see tests/test_vrt.py)",
    )


#: The one xdist group every test that drives Word joins.  Word is a single instance per
#: machine, and :func:`oracle.export` opens a document and then exports ``active
#: document``: two exports at once would export each other's pages, and each begins by
#: quitting Word (:func:`oracle.recover`) under the other.  (Across processes, and with
#: docx-agent's oracle, the Word lock in ``tools/oracle.py`` keeps them apart.)
WORD_GROUP = "word"
#: Fixtures that export through the oracle, launching Word when the export is not cached.
#: A test that requests one, however indirectly, is marked ``word`` here even if it forgot
#: the marker.
WORD_FIXTURES = frozenset({"oracle_pdf"})
#: Modules whose tests share work computed on first use rather than at import (every
#: pytest-xdist worker collects every module): each module's tests join an xdist group of
#: their own, so that one worker does that work once instead of each worker for its share.
LAZY_MODULES = frozenset({"test_script.py", "test_page_top.py", "test_tables.py", "test_label.py", "test_justify.py",
                          "test_hyphenation.py"})
#: How the controller is distributing tests, handed to its workers (which xdist starts
#: with ``dist`` reset to ``"no"``) through the environment they inherit.
_DIST_ENV = "DOCX2SVG_XDIST_DIST"


def pytest_cmdline_main(config) -> None:
    """Make ``-n`` mean ``--dist loadgroup`` unless another mode was asked for.

    Runs after pytest-xdist's own (``tryfirst``) hook, which turns ``-n`` into ``load``.
    ``loadgroup`` is ``load`` for every test without an ``xdist_group``, so it changes
    nothing for them; it is what keeps the ``word`` tests on one worker without every
    caller having to remember the flag."""
    if not hasattr(config.option, "dist") or hasattr(config, "workerinput"):
        return
    if config.option.dist == "load":  # also an explicit --dist load: indistinguishable here
        config.option.dist = "loadgroup"
    os.environ[_DIST_ENV] = config.option.dist


@pytest.hookimpl(tryfirst=True)  # before xdist's own, which reads the groups into node ids
def pytest_collection_modifyitems(config, items) -> None:
    """Tests marked ``slow`` run only with ``--run-slow``: the default run stays the size it
    was before Phase 5 (about five minutes) and needs nothing but the repository.

    And tests that drive Word stay off concurrent workers.  Every such test is marked
    ``word`` (``tests/test_oracle.py``, the tests of Word's export of the layout sweep;
    and any test requesting a fixture of :data:`WORD_FIXTURES`, marked here), and joins
    one xdist group, which ``loadgroup`` -- what ``-n`` means here
    (:func:`pytest_cmdline_main`) -- runs on a single worker, one test at a time.  Any
    other distribution ignores groups, so under one such a test fails rather than races
    (:func:`pytest_runtest_setup`).  Serially nothing changes.  Nothing else in the suite
    launches Word: the fidelity and converter tests read the oracle's cached exports and
    skip without them."""
    for item in items:
        if WORD_FIXTURES.intersection(getattr(item, "fixturenames", ())) and not item.get_closest_marker("word"):
            item.add_marker(pytest.mark.word)
    if hasattr(config, "workerinput"):
        marked = [item for item in items if item.get_closest_marker("word")]
        for item in marked:
            item.add_marker(pytest.mark.xdist_group(WORD_GROUP))
        for item in items:
            if item.path.name in LAZY_MODULES and not item.get_closest_marker("word"):
                item.add_marker(pytest.mark.xdist_group(item.path.name))
                marked.append(item)
        if marked and os.environ.get(_DIST_ENV) == "loadgroup":
            # A worker re-parses the command line, so under a bare -n it does not know the
            # controller is grouping, and would leave the group out of the node ids that
            # xdist's scheduler reads it from.
            config.option.loadgroup = True
    if config.getoption("--run-slow"):
        return
    skip = pytest.mark.skip(reason="slow: run with --run-slow")
    for item in items:
        if "slow" in item.keywords:
            item.add_marker(skip)


def pytest_runtest_setup(item) -> None:
    if item.get_closest_marker("pdfium") and not item.config.getoption("--pdfium"):
        pytest.skip("scores against the old pdfium instrument: run with `pytest --pdfium`")
    if (hasattr(item.config, "workerinput") and item.get_closest_marker("word")
            and os.environ.get(_DIST_ENV) != "loadgroup"):
        pytest.fail("this test drives Microsoft Word, which is one instance per machine: run it serially, "
                    "or in parallel with --dist loadgroup")


#: What a layout reports when a face its document names is not installed: a paragraph (or a
#: table cell's) cannot be measured, or a line number has no advances.  On a machine without Office's
#: faces (every CI runner but Windows) a test marked ``faces`` that meets one skips: it was
#: written against a layout in those faces and says nothing about the code without them.
_FACE_ABSENT = ("layout-stopped:unmeasurable", "layout-stopped:no face metrics", "line-numbers-not-drawn")


@pytest.fixture(autouse=True)
def _skip_where_faces_are_absent(request, monkeypatch):
    if request.node.get_closest_marker("faces") is None:
        return
    import docx2svg

    lay_out = docx2svg._lay_out

    def checked(source, options):
        result = lay_out(source, options)
        for code, message, *_ in result[0].warnings:
            if code in _FACE_ABSENT or (code.startswith("layout-stopped:") and "cannot be measured" in message):
                pytest.skip(f"a face this document is laid out in is not installed here: [{code}] {message}")
        return result

    monkeypatch.setattr(docx2svg, "_lay_out", checked)


@pytest.fixture
def update_snapshots(request) -> bool:
    """Whether ``--update-snapshots`` was passed; see :mod:`tests.test_vrt`."""
    return bool(request.config.getoption("--update-snapshots"))


@pytest.fixture(scope="session")
def layout_sweep_path() -> Path:
    path = FIXTURES / "layout-sweep.docx"
    if not path.exists():
        pytest.skip("run tools/make_layout_sweep.py first")
    return path


def word_is_available() -> bool:
    if sys.platform != "darwin":
        return False
    if not Path("/Applications/Microsoft Word.app").exists():
        return False
    return shutil.which("osascript") is not None


@pytest.fixture(scope="session")
def oracle_pdf(layout_sweep_path: Path) -> Path:
    """Word's PDF export of the layout sweep, exported once per session.

    Drives Word when the export is not cached, so every test that requests it is marked
    ``word`` (:data:`WORD_FIXTURES`) and none runs beside another under ``-n``.

    An existing PDF is reused rather than regenerated: an export takes several seconds
    of a foreground application stealing focus, and the fixture only changes when
    ``make_layout_sweep.py`` is rerun.  Delete the PDF to force a fresh one.
    """
    pdf = ORACLE_DIR / "layout-sweep.pdf"
    if pdf.exists():
        return pdf
    if not word_is_available():
        pytest.skip("Microsoft Word is not available; the oracle cannot run here")
    # Under the machine-wide Word lock, with the lock-file sweep from the AppleScript
    # header before and after (a `~$` file left by a failed attempt wedges the next one,
    # and the failure looks like a broken document): tools/oracle.py's ``run``.  A Word
    # someone else is using, or a lock held past its timeout, skips.
    try:
        run(SCRIPT, layout_sweep_path.read_bytes(), ORACLE_DIR / layout_sweep_path.name, pdf, timeout=300)
    except WordBusy as busy:
        pytest.skip(f"Word is in use: {busy}")
    except RuntimeError as error:
        pytest.fail(str(error))
    return pdf


@pytest.fixture(scope="session")
def oracle_lines(oracle_pdf: Path):
    pytest.importorskip("pypdfium2")
    sys.path.insert(0, str(REPO / "tools"))
    from read_layout_sweep import read_pages  # noqa: E402

    return read_pages(oracle_pdf)
