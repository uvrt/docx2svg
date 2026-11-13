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
import re
import zipfile
from xml.etree import ElementTree

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def kept_features(package: bytes, scan, digests: list, previous: dict | None) -> tuple[list[str | None], dict]:
    """:func:`paragraph_features` of ``package``, and what it found per body element by the
    element's digest (round 57): ``scan`` is where the main part's body elements are
    (:class:`docx2svg.relayout.Scan`), ``digests`` theirs, in order, and ``previous`` what an
    earlier version's call returned.  An element found there is not read again; the others
    are read alone, in the main part's namespaces -- or, where most are new, the whole part
    is read once.  The same list :func:`paragraph_features` makes."""
    w = _W
    archive = zipfile.ZipFile(io.BytesIO(package))
    pictures = frozenset(_picture_levels(archive))
    xml = scan.xml
    by_digest: dict = {}
    known = {digest: found for (kind, digest), found in (previous or {}).items() if kind == pictures}
    missing = sum(1 for digest in digests if digest not in known)
    if missing > len(digests) // 4:
        root = ElementTree.fromstring(xml)
        body = root.find(w + "body")
        children = list(body) if body is not None else []
        if len(children) != len(scan.spans) or any(
                child.tag.rsplit("}", 1)[-1] != name for child, (name, _, _) in zip(children, scan.spans)):
            return paragraph_features(package), {}
        for child, digest in zip(children, digests):
            known.setdefault(digest, tuple(_features(child, pictures)))
    else:
        # Each element alone, in the part's root and body start tags (their namespaces).
        opening = xml[:scan.body]
        names = re.findall(rb"<([A-Za-z_][\w.:-]*)", opening)
        if len(names) < 2:
            return paragraph_features(package), {}
        closing = b"</" + names[-1] + b"></" + names[0] + b">"
        for (_, start, end), digest in zip(scan.spans, digests):
            if digest in known:
                continue
            try:
                root = ElementTree.fromstring(opening + xml[start:end] + closing)
            except ElementTree.ParseError:
                return paragraph_features(package), {}
            body = root.find(w + "body")
            element = body[0] if body is not None and len(body) else None
            known[digest] = tuple(_features(element, pictures)) if element is not None else ()
    out: list[str | None] = []
    for digest in digests:
        found = known[digest]
        out.extend(found)
        by_digest[(pictures, digest)] = found
    return out, by_digest


def _picture_levels(archive) -> set[tuple[str | None, str | None]]:
    """List levels whose label is a picture (``w:lvlPicBulletId``), as ``(numId, ilvl)``."""
    w = _W
    pictures: set[tuple[str | None, str | None]] = set()
    if "word/numbering.xml" in archive.namelist():
        numbering = ElementTree.fromstring(archive.read("word/numbering.xml"))
        abstract = {a.get(w + "abstractNumId"): {lvl.get(w + "ilvl") for lvl in a.findall(w + "lvl")
                                                 if lvl.find(w + "lvlPicBulletId") is not None}
                    for a in numbering.findall(w + "abstractNum")}
        for num in numbering.findall(w + "num"):
            ref = num.find(w + "abstractNumId")
            for level in abstract.get(ref.get(w + "val") if ref is not None else None, ()):
                pictures.add((num.get(w + "numId"), level))
    return pictures


def _features(node, pictures) -> list[str | None]:
    """:func:`paragraph_features` of one body element: a paragraph's, a table's cells'
    paragraphs' in order, nothing for another element."""
    w = _W
    out: list[str | None] = []

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

    walk([node])
    return out


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
    pictures: set[tuple[str | None, str | None]] = set()
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
