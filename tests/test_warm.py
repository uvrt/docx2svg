"""``docx2svg.warm`` and the font scan: a process's one-time work done in advance, or taken
from an earlier process, changes no output.

The warm-up fills the caches a first conversion fills -- the modules, the font folders'
index, the faces a document uses -- with the same values, so a conversion after it is the
same bytes as one without it; the scan it returns (``docx2svg.fonts.FontScan``) is taken
by a later process only for the files whose modification time and size are unchanged.
The made-up fonts here are built in memory (``test_font_scan``'s builder); the end-to-end
test lays out a committed fixture with whatever faces the machine has, in subprocesses, so
each starts cold.
"""

from __future__ import annotations

import json
import os
import pickle
import struct
import subprocess
import sys
from pathlib import Path

import pytest
from test_font_scan import _font, _name_table

import docx2svg
from docx2svg import fonts

REPO = Path(__file__).resolve().parent.parent
DOCUMENT = REPO / "tests" / "fixtures" / "wordto" / "sample-with-table.docx"

_HEAD = struct.pack(">IIIIHH", 0x00010000, 0, 0, 0x5F0F3CF5, 0, 1000) + b"\0" * 34


def _os2(bold: bool = False, italic: bool = False, weight: int = 400) -> bytes:
    """An ``OS/2`` table as long as fsSelection (offset 62) needs: the weight and width
    classes and the bold and italic bits."""
    table = bytearray(78)
    struct.pack_into(">HH", table, 4, weight, 5)
    struct.pack_into(">H", table, 62, (0x20 if bold else 0) | (0x01 if italic else 0))
    return bytes(table)


def _face(family: str, **style) -> bytes:
    return _font({b"name": _name_table(family), b"head": _HEAD, b"OS/2": _os2(**style)})


def _collection(*faces: bytes) -> bytes:
    """A ``.ttc`` holding ``faces``, each font's table offsets moved to where it lands."""
    header = 12 + 4 * len(faces)
    out, offsets = b"", []
    for face in faces:
        start = header + len(out)
        count = struct.unpack_from(">H", face, 4)[0]
        moved = bytearray(face)
        for i in range(count):
            entry = 12 + 16 * i
            struct.pack_into(">I", moved, entry + 8, struct.unpack_from(">I", face, entry + 8)[0] + start)
        offsets.append(start)
        out += bytes(moved)
    return struct.pack(">4sHHI", b"ttcf", 1, 0, len(faces)) + struct.pack(f">{len(faces)}I", *offsets) + out


def _read(face) -> dict:
    """What the scan asks of a face, or the error each answer raises."""
    out = {}
    for name in ("style", "rasteriser_families", "css", "variable", "names"):
        try:
            out[name] = getattr(face, name)
        except Exception as error:  # noqa: BLE001 -- the error is the answer compared
            out[name] = type(error)
    try:
        out["families"] = face.families()
    except Exception as error:  # noqa: BLE001
        out["families"] = type(error)
    return out


@pytest.fixture
def fresh_scan(monkeypatch):
    """No file read yet, as in a new process (and the index forgotten after)."""
    monkeypatch.setattr(fonts, "_SCANNED", {})
    monkeypatch.setattr(fonts, "_RASTERISER", {})
    fonts.installed_index.cache_clear()
    yield
    fonts.installed_index.cache_clear()


def test_the_header_scan_answers_as_the_whole_file(tmp_path):
    """Read at their offsets, the scan's three tables give each face the same names, style
    and CSS keys as the whole file does, and a file that cannot be read fails alike."""
    files = {
        "single.ttf": _face("Warm Probe", bold=True),
        "pair.ttc": _collection(_face("Warm Pair"), _face("Warm Pair", italic=True, weight=300)),
        "headless.ttf": _font({b"name": _name_table("Warm Headless")}),
        "short.ttf": _face("Warm Short")[:40],
        "empty.ttf": b"",
        "past-the-end.ttc": _collection(_face("Warm Cut"))[:60],
    }
    for name, data in files.items():
        (tmp_path / name).write_bytes(data)
        whole = [(number, face.offset, _read(face)) for number, face in fonts._faces_in(tmp_path / name)]
        scanned = [(number, face.offset, _read(face)) for number, face in fonts._scanned_faces(tmp_path / name)]
        assert scanned == whole, name
    # A table the scan did not read is read from the file when asked for.
    (face,) = [face for _, face in fonts._scanned_faces(tmp_path / "single.ttf")]
    assert face.table(b"head") == _HEAD and face.data == files["single.ttf"]


def test_a_changed_file_is_read_again(tmp_path, fresh_scan, monkeypatch):
    """A scan is taken for the files it recorded that are unchanged; a file rewritten is
    read afresh, and the index finds its new name."""
    (tmp_path / "a.ttf").write_bytes(_face("Warm Kept"))
    (tmp_path / "b.ttf").write_bytes(_face("Warm Before"))
    scan = pickle.loads(pickle.dumps(fonts.font_scan((tmp_path,), "test")))
    assert set(scan.files) == {str(tmp_path / "a.ttf"), str(tmp_path / "b.ttf")}

    (tmp_path / "b.ttf").write_bytes(_face("Warm After, longer"))
    monkeypatch.setattr(fonts, "_SCANNED", {})
    assert fonts.load_scan(scan, "test") == 1
    read = []
    real = fonts._scanned_faces
    monkeypatch.setattr(fonts, "_scanned_faces", lambda path: read.append(path.name) or real(path))
    index = fonts.installed_index((tmp_path,))
    assert read == ["b.ttf"]
    assert set(index) == {("warm kept", False, False), ("warm after, longer", False, False)}


def test_a_scan_of_another_version_is_not_taken(tmp_path, fresh_scan):
    (tmp_path / "a.ttf").write_bytes(_face("Warm Version"))
    scan = fonts.font_scan((tmp_path,), "0.0.1")
    assert fonts.load_scan(scan, "0.0.2") == 0
    assert fonts.load_scan({"files": {}}, "0.0.1") == 0
    assert fonts.load_scan(scan, "0.0.1") == 1


def test_a_run_keeps_its_character_formats_out_of_a_pickle():
    """A run's formats are made once per slot and shared by its characters; they hold the
    document, so a pickled ``Resolved`` leaves them out."""
    from docx2svg.resolve import character_format, resolve_run

    document = docx2svg.parse_package(DOCUMENT.read_bytes())
    paragraph = next(block for block in document.body if getattr(block, "runs", None))
    resolved = resolve_run(document, paragraph, paragraph.runs[0])
    first = character_format(resolved, "a", document)
    assert character_format(resolved, "b", document) is first
    assert character_format(resolved, "一", document).slot == "eastAsia"
    copied = pickle.loads(pickle.dumps(resolved))
    assert "_formats" not in copied.__dict__ and copied == resolved
    assert character_format(copied, "a", document) == first


_SCRIPT = """
import hashlib, json, pickle, sys
sys.path.insert(0, {src!r})
import docx2svg
from docx2svg import fonts
mode, scan_path = sys.argv[1], sys.argv[2]
read = []
real = fonts._scanned_faces
fonts._scanned_faces = lambda path: read.append(str(path)) or real(path)
if mode == "warm":
    with open(scan_path, "wb") as handle:
        pickle.dump(docx2svg.warm(source={document!r}), handle)
elif mode == "scan":
    with open(scan_path, "rb") as handle:
        scan = pickle.load(handle)
    docx2svg.warm(scan=scan)
    unrecorded = [path for path in fonts._listed(docx2svg._font_dirs(docx2svg.ConvertOptions()))
                  if path not in scan.files]
    assert sorted(set(read)) == sorted(unrecorded), (len(read), len(unrecorded))
options = docx2svg.ConvertOptions()
svgs = docx2svg.convert_docx_to_svg({document!r}, options)
print(json.dumps({{"pages": [hashlib.sha256(svg.encode()).hexdigest() for svg in svgs],
                  "warnings": [str(warning) for warning in options.warnings],
                  "coverage": repr(options.coverage)}}))
"""


def test_a_warm_process_renders_the_same_bytes(tmp_path):
    """Cold, warmed with the document, and warmed from the first warm-up's scan: three
    processes, the same pages, warnings and coverage -- and the third reads no font file
    the scan recorded."""
    script = _SCRIPT.format(src=str(REPO / "src"), document=str(DOCUMENT))
    scan = str(tmp_path / "scan.pickle")
    outputs = []
    for mode in ("cold", "warm", "scan"):
        result = subprocess.run([sys.executable, "-c", script, mode, scan], capture_output=True, text=True,
                                check=False, cwd=str(REPO), env=dict(os.environ))
        assert result.returncode == 0, result.stderr
        outputs.append(json.loads(result.stdout))
    assert outputs[0]["pages"]
    assert outputs[1] == outputs[0]
    assert outputs[2] == outputs[0]


def test_warm_leaves_the_options_as_they_were():
    options = docx2svg.ConvertOptions()
    scan = docx2svg.warm(DOCUMENT, options)
    assert options.warnings == [] and options.coverage is None
    assert isinstance(scan, fonts.FontScan) and scan.version == docx2svg.__version__
