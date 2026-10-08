"""What the parsed model cannot stand for: paragraphs whose layout is not modelled.

:func:`paragraph_features` says, per paragraph, which of these it holds, read from
``word/document.xml`` directly: frames, note references, picture bullets -- and fields.
The paginator can treat such a paragraph as an obstacle (``paginate.flow(features=...)``)
and the renderer can say so rather than draw text Word does not draw.  Moved here
unchanged from ``tools/baselines.py``, which now calls it, when the renderer needed it at
run time.

**Fields are laid out now** (ROADMAP.md, "Headers, footers and fields -- measured"): the
parser no longer keeps an instruction as run text, draws every field's cached result and
computes ``PAGE``, ``NUMPAGES`` and ``SECTIONPAGES``, so the renderer ignores ``field``.
The baseline and page scorers (``tools/baselines.py``, ``tools/pages.py``) still leave a
field's paragraph out of scope, as their pinned counts were measured.
"""

from __future__ import annotations

import io
import zipfile
from xml.etree import ElementTree

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def paragraph_features(package: bytes) -> list[str | None]:
    """For every paragraph in document order, cells included (``tools/baselines.blocks``
    order), why the parsed model cannot stand for it (``None`` when nothing is wrong):
    ``drawing (picture bullet)``, ``note reference``, ``field`` or ``frame``.

    Walks the body the way the parser does: paragraphs and tables, cells' paragraphs and
    nested tables in order.
    """
    w = _W
    archive = zipfile.ZipFile(io.BytesIO(package))
    root = ElementTree.fromstring(archive.read("word/document.xml"))
    out: list[str | None] = []

    # List levels whose label is a picture (``w:lvlPicBulletId``): a drawing on the line.
    pictures: set[tuple[str, str]] = set()
    if "word/numbering.xml" in archive.namelist():
        numbering = ElementTree.fromstring(archive.read("word/numbering.xml"))
        abstract = {a.get(w + "abstractNumId"): {lvl.get(w + "ilvl") for lvl in a.findall(w + "lvl")
                                                 if lvl.find(w + "lvlPicBulletId") is not None}
                    for a in numbering.findall(w + "abstractNum")}
        for num in numbering.findall(w + "num"):
            ref = num.find(w + "abstractNumId")
            for level in abstract.get(ref.get(w + "val") if ref is not None else None, ()):
                pictures.add((num.get(w + "numId"), level))

    def feature(p) -> str | None:
        num = p.find(f"{w}pPr/{w}numPr")
        if num is not None:
            num_id, ilvl = num.find(w + "numId"), num.find(w + "ilvl")
            key = (num_id.get(w + "val") if num_id is not None else None,
                   ilvl.get(w + "val") if ilvl is not None else "0")
            if key in pictures:
                return "drawing (picture bullet)"
        names = {node.tag for node in p.iter()}
        if w + "footnoteReference" in names or w + "endnoteReference" in names:
            return "note reference"
        if w + "instrText" in names or w + "fldSimple" in names:
            return "field"
        if p.find(f"{w}pPr/{w}framePr") is not None:
            return "frame"
        return None

    def walk(parent):
        for node in parent:
            if node.tag == w + "p":
                out.append(feature(node))
            elif node.tag == w + "tbl":
                for row in node.findall(w + "tr"):
                    for cell in row.findall(w + "tc"):
                        walk(cell)

    walk(root.find(w + "body"))
    return out
