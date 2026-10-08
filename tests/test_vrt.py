"""Snapshot VRT: render every committed document and compare with the committed SVG.

The sibling ``pptx2svg``'s arrangement (its ``tests/test_vrt.py``), for the same reason:
a regression net that runs everywhere -- no Word, no font file, no rasteriser -- and
answers one question, *did the output change?*  It cannot say the output is right;
``tests/test_render.py`` (every glyph against Word's PDF) and ``tools/fidelity.py`` (the
raster) do that.  Never rebaseline a failure before knowing which of the two it is.

What makes byte equality a fair test here:

* **No host font is read.**  The faces' numbers come from the recording
  ``tests/fixtures/render-observations.json`` (``tools/render_record.RecordedFonts``),
  so a snapshot is the same on a machine with Word's faces and on one without.
* **Hash seeds.**  ``PYTHONHASHSEED`` is randomised per process; a render that ordered
  output by a ``set`` would differ between processes and never within one, so
  :func:`test_a_second_process_with_a_different_hash_seed_renders_the_same_bytes`
  renders in two subprocesses under two seeds.
* **Line endings and encoding.**  ``.gitattributes`` pins ``tests/vrt/**/*.svg`` to LF;
  every read and write here is UTF-8 bytes; the render emits no carriage return.
* **Pictures** are compared by the SHA-256 of their bytes, not their base64
  (:func:`normalise`), so a snapshot does not carry the fixture's images a second time.

Rebaselining::

    python -m pytest tests/test_vrt.py --update-snapshots

The flag skips rather than passes, so a run that somehow acquired it cannot report
success.  A failing render is written to ``tests/vrt/diffs/`` (gitignored).
"""

from __future__ import annotations

import base64
import hashlib
import json
import os
import re
import subprocess
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
FIXTURES = REPO / "tests" / "fixtures"
SNAPSHOT_DIR = REPO / "tests" / "vrt"
DIFF_DIR = SNAPSHOT_DIR / "diffs"
REBASELINE = "python -m pytest tests/test_vrt.py --update-snapshots"

#: The committed documents, as ``tests/fixtures`` holds them.
DOCUMENTS = (
    "layout-sweep.docx", "style-document.docx",
    "samplelib/sample-blank.docx", "samplelib/sample-long.docx", "samplelib/sample-resume.docx",
    "samplelib/sample-simple.docx",
    "wordto/sample-1page.docx", "wordto/sample-5pages.docx", "wordto/sample-10pages.docx",
    "wordto/sample-with-images.docx", "wordto/sample-with-table.docx",
)

#: Generated documents whose snapshots hold headers, footers and fields: the story probe's
#: mode-15 document, the page-selection probe's even/odd document and the field probe's
#: computed document (``tools/read_story_probe.py``), measured with the faces recorded in
#: ``tests/fixtures/story-observations.json``.
PROBES = ("story-15", "select-even", "fields-computed")

_IMAGE = re.compile(r'href="data:([^;"]+);base64,([^"]*)"')


def normalise(document: str) -> str:
    """Each embedded picture's base64 replaced by the SHA-256 of its bytes."""
    def digest(match: re.Match) -> str:
        data = base64.b64decode(match.group(2))
        return f'href="data:{match.group(1)};sha256,{hashlib.sha256(data).hexdigest()}"'

    return _IMAGE.sub(digest, document)


def render(name: str) -> list[str]:
    """Every page of a committed document, measured with the recorded faces."""
    sys.path.insert(0, str(REPO / "tools"))
    import render_record

    from docx2svg import convert_docx_to_svg

    if name in PROBES:
        import read_story_probe

        faces = json.loads(read_story_probe.OBSERVATIONS.read_text(encoding="utf-8"))["faces"]
        source = dict(read_story_probe.documents())[name]
    else:
        faces = json.loads((FIXTURES / "render-observations.json").read_text(encoding="utf-8"))["faces"]
        source = FIXTURES / name
    fonts = render_record.RecordedFonts(faces)
    return [normalise(page) for page in convert_docx_to_svg(source, render_record.options(fonts))]


def directory_of(name: str) -> Path:
    return SNAPSHOT_DIR / name.replace("/", "-").removesuffix(".docx")


def snapshot_path(name: str, page: int) -> Path:
    return directory_of(name) / f"page-{page:02d}.svg"


def read_snapshot(path: Path) -> str:
    return path.read_bytes().decode("utf-8").replace("\r\n", "\n")


def write_snapshot(path: Path, document: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(document.encode("utf-8"))


def describe_difference(expected: str, actual: str) -> str:
    limit = min(len(expected), len(actual))
    offset = next((i for i in range(limit) if expected[i] != actual[i]), limit)
    start = max(0, offset - 70)
    return (f"first difference at character {offset} (snapshot {len(expected)} chars, render {len(actual)})\n"
            f"  snapshot: ...{ascii(expected[start:offset + 70])}\n"
            f"  render:   ...{ascii(actual[start:offset + 70])}")


@pytest.mark.parametrize("name", DOCUMENTS + PROBES)
def test_svg_matches_the_committed_snapshot(name: str, update_snapshots: bool):
    documents = render(name)
    directory = directory_of(name)
    if update_snapshots:
        for stale in sorted(directory.glob("page-*.svg")):
            stale.unlink()
        for number, document in enumerate(documents, start=1):
            write_snapshot(snapshot_path(name, number), document)
        pytest.skip(f"--update-snapshots: rewrote {len(documents)} snapshot(s)")
    committed = sorted(directory.glob("page-*.svg"))
    assert committed, f"no snapshots in {directory}: capture them with `{REBASELINE}` and read them first"
    assert len(committed) == len(documents), f"{name} renders {len(documents)} page(s), {len(committed)} committed"
    failures = []
    for number, document in enumerate(documents, start=1):
        path = snapshot_path(name, number)
        expected = read_snapshot(path)
        if document != expected:
            dump = DIFF_DIR / directory.name / path.name
            write_snapshot(dump, document)
            failures.append(f"{path}\n{describe_difference(expected, document)}\n  render written to {dump}")
    assert not failures, "\n\n".join(failures) + f"\n\nIf the change is intended, rebaseline with `{REBASELINE}`."


def test_every_snapshot_directory_has_a_document():
    expected = {directory_of(name).name for name in DOCUMENTS + PROBES}
    found = {path.name for path in SNAPSHOT_DIR.iterdir() if path.is_dir() and path.name != "diffs"}
    assert found == expected


def test_the_render_emits_no_carriage_returns():
    for name in ("samplelib/sample-long.docx", "style-document.docx"):
        assert not any("\r" in page for page in render(name))


def test_a_second_process_with_a_different_hash_seed_renders_the_same_bytes():
    """Two subprocesses, two seeds, the same bytes -- and the committed ones."""
    name = "style-document.docx"
    script = (
        "import sys, json; sys.path.insert(0, {tests!r}); sys.path.insert(0, {src!r});"
        "import test_vrt; print(json.dumps(test_vrt.render({name!r})))"
    ).format(tests=str(REPO / "tests"), src=str(REPO / "src"), name=name)
    outputs = []
    for seed in ("1", "2"):
        environment = dict(os.environ, PYTHONHASHSEED=seed)
        result = subprocess.run([sys.executable, "-c", script], capture_output=True, text=True, env=environment,
                                check=True, cwd=str(REPO))
        outputs.append(json.loads(result.stdout))
    assert outputs[0] == outputs[1]
    committed = [read_snapshot(path) for path in sorted(directory_of(name).glob("page-*.svg"))]
    assert outputs[0] == committed
