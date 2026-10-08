"""Element paths: what ``data-docx-path`` and every layout object's ``path`` name.

Every element the SVG draws from the document carries ``data-docx-path``, and the layout
objects behind it carry the same string: :attr:`docx2svg.layout.Line.path` (the
paragraph a line is of), :attr:`~docx2svg.layout.Span.path` (the run),
:attr:`~docx2svg.layout.Float.path` (the floating drawing),
:attr:`~docx2svg.layout.Stop.path` (the obstacle the layout stopped at).  A path is a
chain of steps from its part's root, counted as the parser counts them when it reads the
part (:class:`Steps`, :func:`row_level`), so that :func:`resolve_path` finds the element
again in any ElementTree-compatible tree of the same part -- the standard library's, or
lxml's -- and a caller that edits the document can find what was drawn:

* ``w:name[k]``: the k-th child with that **local name**, whatever its namespace (the
  prefix written is always ``w:`` for WordprocessingML's own steps), comments and
  processing instructions not counted -- XPath's ``*[local-name()='name'][k]``;
* ``w:tr[k]`` and ``w:tc[k]``: the k-th row of a table, the k-th cell of a row, counted
  **through** row- and cell-level content controls (``w:sdt/w:sdtContent``) and
  ``w:customXml``, and counting rows deleted as a revision (which are not drawn);
* ``w:footnote[@w:id=N]``, ``w:endnote[@w:id=N]``: a note, by its id;
* ``wp:anchor[k]`` after a run: its k-th floating drawing, ``wp:inline[k]`` its k-th
  inline DrawingML object (a ``w:drawing`` or the ``mc:Choice`` of an
  ``mc:AlternateContent`` holding one); a group's members (``wps:wsp[k]``,
  ``pic:pic[k]``, ``wpg:grpSp[k]``, ``wpg:graphicFrame[k]``) are counted among the
  group's own children;
* a step without an index (``w:body``, ``wps:txbx``, ``w:txbxContent``): the first
  element with that local name below (the part's root itself when it is the first step:
  ``w:hdr``, ``w:ftr``, ``w:footnotes``);
* ``label``: the list label drawn for the paragraph before it (resolves to the paragraph);
* ``w:sectPr[k]/w:cols`` (a column separator): the k-th section's ``w:cols``, the
  sections' ``w:sectPr`` taken in document order.

The part a path is in is the main document part for the body, the story's part for a
header or footer (:attr:`Line.part <docx2svg.layout.Line.part>`, ``data-docx-part``), and
the footnotes or endnotes part for a path that starts at a note: :func:`path_part`.
:func:`locate` does both against a package.

A few paths name something the document does not hold, and resolve to ``None``: Word's
own footnote separator where the document has none (``w:footnotes/separator``).
"""

from __future__ import annotations

import re
from typing import Iterator

from .xmlutil import attr, child, children, local_name

__all__ = ["Steps", "locate", "path_part", "resolve_path", "row_level", "split_path"]


class Steps:
    """The parser's step counter: 1-based indexes among same-named siblings (by local
    name), as XPath counts them, each step written after ``prefix``."""

    def __init__(self, prefix: str) -> None:
        self.prefix = prefix
        self.counts: dict[str, int] = {}

    def __call__(self, node) -> str:
        name = local_name(node.tag)
        self.counts[name] = self.counts.get(name, 0) + 1
        step = f"w:{name}[{self.counts[name]}]"
        return f"{self.prefix}/{step}" if self.prefix else step


def row_level(parent, name: str) -> Iterator:
    """``parent``'s ``w:tr`` or ``w:tc`` children, also those inside a ``w:sdt``'s
    ``w:sdtContent`` or a ``w:customXml`` (row- and cell-level containers)."""
    for node in children(parent):
        local = local_name(node.tag)
        if local == name:
            yield node
        elif local == "sdt":
            content = child(node, "sdtContent")
            if content is not None:
                yield from row_level(content, name)
        elif local == "customXml":
            yield from row_level(node, name)


_STEP = re.compile(r"(?:(?P<prefix>\w+):)?(?P<name>\w+)(?:\[(?P<index>\d+)\]|\[@w:id=(?P<id>-?\d+)\])?")


def split_path(path: str) -> list[tuple[str | None, str, int | None, str | None]]:
    """A path's steps as ``(prefix, local name, 1-based index, note id)``; ``ValueError``
    for a string that is not a path."""
    out = []
    for step in path.split("/"):
        match = _STEP.fullmatch(step)
        if match is None:
            raise ValueError(f"not a docx2svg path step: {step!r} in {path!r}")
        index = match.group("index")
        out.append((match.group("prefix"), match.group("name"), int(index) if index else None, match.group("id")))
    return out


def _run_objects(run) -> Iterator:
    """A run's objects, as the parser meets them: each child that may be a drawing."""
    from .parse.document import _RUN_OBJECTS

    for node in children(run):
        if local_name(node.tag) in _RUN_OBJECTS:
            yield node


def _anchors(run) -> list:
    """A run's floating drawings, in order (``wp:anchor[k]``)."""
    from .parse.drawing import drawing_element

    out = []
    for node in _run_objects(run):
        drawing = drawing_element(node)
        anchor = child(drawing, "anchor") if drawing is not None else None
        if anchor is not None:
            out.append(anchor)
    return out


def _inlines(run) -> list:
    """A run's inline DrawingML objects, in order (``wp:inline[k]``): each object the
    parser marks ``drawing:...`` (one ``wp:inline`` with an extent)."""
    from .parse.document import _object
    from .parse.drawing import drawing_element

    out = []
    for node in _run_objects(run):
        if _object(node).startswith("drawing:"):
            out.append(child(drawing_element(node), "inline"))
    return out


def _members(node):
    """Where a drawing's group members are counted: the group itself, or the group an
    anchor or inline drawing shows."""
    if local_name(node.tag) in ("anchor", "inline"):
        data = child(child(node, "graphic"), "graphicData")
        content = next(iter(children(data)), None)
        if content is not None and local_name(content.tag) in ("wgp", "grpSp"):
            return content
    return node


def _sections(root) -> list:
    """Every section's ``w:sectPr`` in a main document part, in document order: each
    paragraph's (``w:pPr/w:sectPr``), then the body's own."""
    body = child(root, "body")
    out = [found for node in root.iter() if isinstance(node.tag, str) and local_name(node.tag) == "pPr"
           for found in children(node, "sectPr")]
    final = child(body, "sectPr")
    return out + ([final] if final is not None else [])


def resolve_path(root, path: str):
    """The element ``path`` names in a part whose root element is ``root``, or ``None``.

    ``root`` is the part's root element (``w:document``, ``w:hdr``, ``w:ftr``,
    ``w:footnotes``, ``w:endnotes``) as the standard library's ElementTree or lxml parses
    it.  A path is resolved by the rules of :mod:`docx2svg.paths`; a path the part does
    not hold (an edited tree, a synthetic step) resolves to ``None``."""
    if root is None or not path:
        return None
    try:
        steps = split_path(path)
    except ValueError:
        return None
    node = root
    for number, (prefix, name, index, note) in enumerate(steps):
        if prefix is None:
            # A pseudo-step: the list label is drawn for its paragraph; anything else
            # (Word's default separator) is not in the document.
            if name == "label" and local_name(node.tag) == "p":
                continue
            return None
        if number == 0 and name == "sectPr" and index is not None and local_name(root.tag) == "document":
            sections = _sections(root)
            node = sections[index - 1] if 0 < index <= len(sections) else None
        elif note is not None:
            node = next((c for c in children(node, name) if attr(c, "id") == note), None)
        elif index is None:
            if number == 0 and local_name(node.tag) == name:
                continue
            node = next((d for d in node.iter() if d is not node and isinstance(d.tag, str)
                         and local_name(d.tag) == name), None)
        else:
            if name in ("tr", "tc"):
                candidates = list(row_level(node, name))
            elif name == "anchor" and local_name(node.tag) == "r":
                candidates = _anchors(node)
            elif name == "inline" and local_name(node.tag) == "r":
                candidates = _inlines(node)
            else:
                candidates = list(children(_members(node), name))
            node = candidates[index - 1] if 0 < index <= len(candidates) else None
        if node is None:
            return None
    return node


def path_part(package, path: str, part: str | None = None) -> str | None:
    """The part a path is in: ``part`` when the drawn object says (a header's or footer's
    line, :attr:`Line.part <docx2svg.layout.Line.part>`), the footnotes or endnotes part
    for a note's path, else the main document part.  ``package`` is a
    :class:`docx2svg.opc.Package`."""
    from .opc import REL_ENDNOTES, REL_FOOTNOTES

    if part:
        return part
    main = package.main_document_part
    first = path.split("/", 1)[0]
    for prefix, relationship in (("w:footnote", REL_FOOTNOTES), ("w:endnote", REL_ENDNOTES)):
        if first == f"{prefix}s" or first.startswith(f"{prefix}["):
            related = package.related(main, relationship)
            return related[0] if related else None
    return main


def locate(source, path: str, part: str | None = None):
    """``(part, element)`` for a drawn object's path, ``element`` parsed by the standard
    library from ``source`` (a :class:`docx2svg.opc.Package`, the ``.docx``'s bytes or a
    file name); ``element`` is ``None`` where the path names nothing in the part.

    For many lookups, parse each part once and call :func:`resolve_path`."""
    from .opc import Package
    from .xmlutil import parse_xml

    package = source if isinstance(source, Package) else Package.open(source)
    where = path_part(package, path, part)
    if where is None or not package.exists(where):
        return where, None
    return where, resolve_path(parse_xml(package.read(where)), path)
