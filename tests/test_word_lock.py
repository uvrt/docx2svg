"""The machine-wide Word lock (``tools/oracle.py``, :func:`oracle.word`), without Word.

The lock is a file the test points elsewhere, and Word's process and ``pkill`` are
stood in for: nothing here launches, quits or looks at the real Word.
"""

from __future__ import annotations

import fcntl
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))

import oracle  # noqa: E402


@pytest.fixture
def word(tmp_path, monkeypatch):
    """A lock file and an oracle directory of the test's own, and a stand-in Word:
    ``state["running"]`` says whether it runs; ``state["killed"]`` counts ``pkill``s."""
    state = {"running": False, "killed": 0}
    monkeypatch.setattr(oracle, "GROUP_CONTAINER", tmp_path)
    monkeypatch.setattr(oracle, "WORD_LOCK", tmp_path / "word-oracle.lock")
    monkeypatch.setattr(oracle, "ORACLE_DIR", tmp_path / "docx2svg-oracle")
    monkeypatch.setattr(oracle, "word_running", lambda: state["running"])
    monkeypatch.setattr(oracle.time, "sleep", lambda seconds: None)

    def run(command, **kwargs):
        assert command[:2] == ["pkill", "-x"], command
        state["killed"] += 1
        state["running"] = False

    monkeypatch.setattr(oracle.subprocess, "run", run)
    (tmp_path / "docx2svg-oracle").mkdir()
    return state


def _held_elsewhere(path: Path):
    """The lock held through another open file description, as another process holds it."""
    handle = open(path, "a+")
    fcntl.flock(handle, fcntl.LOCK_EX | fcntl.LOCK_NB)
    return handle


def test_the_lock_is_docx_agent_s():
    assert oracle.WORD_LOCK == Path.home() / "Library" / "Group Containers" / "UBF8T346G9.Office" / "word-oracle.lock"


def test_the_lock_is_held_while_word_works_and_released_after(word):
    with oracle.word():
        with pytest.raises(BlockingIOError):
            _held_elsewhere(oracle.WORD_LOCK)
        with oracle.word():  # re-entrant
            pass
        with pytest.raises(BlockingIOError):
            _held_elsewhere(oracle.WORD_LOCK)
    _held_elsewhere(oracle.WORD_LOCK).close()


def test_a_held_lock_is_waited_for_and_then_given_up(word, monkeypatch):
    clock = iter(range(0, 10_000, 5))
    monkeypatch.setattr(oracle.time, "monotonic", lambda: next(clock))
    other = _held_elsewhere(oracle.WORD_LOCK)
    try:
        with pytest.raises(oracle.WordBusy):
            with oracle.word(timeout=60):
                pytest.fail("ran under a lock someone else holds")
    finally:
        other.close()
    with oracle.word(timeout=60):
        pass


def test_a_word_this_process_did_not_start_is_never_quit(word, tmp_path):
    """Word running when the lock is taken is someone's: refused, and left running; so is
    one found by a :func:`oracle.recover` called outside the lock, which clears only the
    oracle directory's ``~$`` files."""
    word["running"] = True
    with pytest.raises(oracle.WordBusy):
        with oracle.word():
            pytest.fail("ran beside a Word in use")
    assert word == {"running": True, "killed": 0}
    stale = oracle.ORACLE_DIR / "~$probe.docx"
    stale.write_bytes(b"")
    elsewhere = tmp_path / "~$theirs.docx"
    elsewhere.write_bytes(b"")
    oracle.recover()
    assert word == {"running": True, "killed": 0}
    assert not stale.exists() and elsewhere.exists()


def test_the_word_this_process_started_is_quit_before_the_lock_is_released(word):
    with oracle.word():
        word["running"] = True  # an export launched it
        oracle.recover()  # under the lock: ours
        assert word["killed"] == 1
        word["running"] = True  # launched again, left running by a failure
    assert word == {"running": False, "killed": 2}


def test_an_export_runs_under_the_lock(word, monkeypatch):
    """:func:`oracle.run`, as every export and re-save goes, holds the lock around the
    script and the recoveries around it."""
    def osascript(command, **kwargs):
        if command[0] == "pkill":
            word["killed"] += 1
            word["running"] = False
            return
        with pytest.raises(BlockingIOError):
            _held_elsewhere(oracle.WORD_LOCK)
        word["running"] = True
        Path(command[3]).write_bytes(b"%PDF")

    monkeypatch.setattr(oracle.subprocess, "run", osascript)
    out = oracle.run(oracle.SCRIPT, b"docx", oracle.ORACLE_DIR / "probe.docx", oracle.ORACLE_DIR / "probe.pdf")
    assert out.read_bytes() == b"%PDF" and not (oracle.ORACLE_DIR / "probe.docx").exists()
    assert word["running"] is False
    _held_elsewhere(oracle.WORD_LOCK).close()
