"""Read a ``w:drawing`` into :class:`docx2svg.model.Anchor` and :class:`~docx2svg.model.Graphic`.

A drawing is either inline (``wp:inline``: a character of its line, which
:mod:`docx2svg.parse.document` records as a mark with its extent) or floating
(``wp:anchor``, ECMA-376 20.4.2.3): positioned against the page, a margin, the column,
its paragraph, line or character, stacked above or below the text, and possibly wrapped
around.  This module reads the anchor as written -- nothing is resolved here -- and what
it shows: a picture (``pic:pic``), a shape (``wps:wsp``, with its geometry, fill,
outline and a text box's content), a group of them (``wpg:wgp``), a chart (``c:chart``) or
SmartArt (``dgm:relIds``), or something else, named by its ``graphicData/@uri``.

Word writes a shape or a group inside ``mc:AlternateContent``, a ``mc:Choice`` that
``Requires`` the ``wps`` or ``wpg`` namespace and a VML ``mc:Fallback``; the choice is the
DrawingML one, and :func:`drawing_element` finds it.
"""

from __future__ import annotations

import dataclasses
from typing import Callable
from xml.etree.ElementTree import Element

from ooxml_common.drawingml import read

from ..model import Anchor, AnchorPosition, Graphic
from ..xmlutil import attr, attr_int, child, children, local_name

_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_R_EMBED = _R + "embed"


def drawing_element(node: Element) -> Element | None:
    """The ``w:drawing`` a run's object is: the node itself, or the first ``mc:Choice``
    of an ``mc:AlternateContent`` that holds one (Word's ``wps`` / ``wpg`` choice)."""
    name = local_name(node.tag)
    if name == "drawing":
        return node
    if name == "AlternateContent":
        for choice in children(node, "Choice"):
            found = child(choice, "drawing")
            if found is not None:
                return found
    return None


def _int(element: Element | None, name: str, default: int = 0) -> int:
    value = attr_int(element, name)
    return default if value is None else value


def _bool(element: Element | None, name: str, default: bool = False) -> bool:
    value = attr(element, name)
    if value is None:
        return default
    return value not in ("0", "false", "off")


def _position(element: Element | None) -> AnchorPosition:
    if element is None:
        return AnchorPosition()
    offset = child(element, "posOffset")
    align = child(element, "align")
    value = None
    if offset is not None and offset.text and offset.text.strip().lstrip("-").isdigit():
        value = int(offset.text.strip())
    return AnchorPosition(attr(element, "relativeFrom"), value,
                          align.text.strip() if align is not None and align.text else None)


def read_anchor(drawing: Element, path: str, blocks: Callable[[Element, str], list]) -> Anchor | None:
    """The ``wp:anchor`` of a ``w:drawing`` (``None`` for an inline one); ``blocks`` reads a
    text box's ``w:txbxContent`` into paragraphs and tables, given its path."""
    anchor = child(drawing, "anchor")
    if anchor is None:
        return None
    extent = child(anchor, "extent")
    effect = child(anchor, "effectExtent")
    wrap_element = next((c for c in children(anchor) if local_name(c.tag).startswith("wrap")), None)
    simple = None
    if _bool(anchor, "simplePos"):
        point = child(anchor, "simplePos")
        simple = (_int(point, "x"), _int(point, "y"))
    doc_pr = child(anchor, "docPr")
    graphic_data = child(child(anchor, "graphic"), "graphicData")
    return Anchor(
        extent=(_int(extent, "cx"), _int(extent, "cy")),
        effect=tuple(_int(effect, side) for side in ("l", "t", "r", "b")),  # type: ignore[arg-type]
        h=_position(child(anchor, "positionH")),
        v=_position(child(anchor, "positionV")),
        simple=simple,
        wrap=local_name(wrap_element.tag) if wrap_element is not None else None,
        wrap_text=attr(wrap_element, "wrapText"),
        distance=tuple(_int(anchor, name) for name in ("distT", "distB", "distL", "distR")),  # type: ignore[arg-type]
        behind=_bool(anchor, "behindDoc"),
        relative_height=_int(anchor, "relativeHeight"),
        allow_overlap=_bool(anchor, "allowOverlap", True),
        layout_in_cell=_bool(anchor, "layoutInCell", True),
        locked=_bool(anchor, "locked"),
        graphic=read_graphic_data(graphic_data, path, blocks) if graphic_data is not None else None,
        identifier=attr(doc_pr, "id"),
        polygon=_polygon(wrap_element),
    )


def _polygon(wrap: Element | None) -> tuple[tuple[int, int], ...] | None:
    """A ``wrapTight`` / ``wrapThrough``'s ``wp:wrapPolygon``: ``wp:start`` and each
    ``wp:lineTo``, in order."""
    polygon = child(wrap, "wrapPolygon") if wrap is not None else None
    if polygon is None:
        return None
    points = tuple((_int(point, "x"), _int(point, "y")) for point in children(polygon)
                   if local_name(point.tag) in ("start", "lineTo"))
    return points or None


def read_graphic_data(graphic_data: Element, path: str, blocks: Callable[[Element, str], list]) -> Graphic:
    """``a:graphicData``'s content, the top-level graphic of a drawing."""
    content = next(iter(children(graphic_data)), None)
    if content is None:
        return Graphic("other", uri=attr(graphic_data, "uri"), path=path)
    graphic = _graphic(content, path, blocks)
    if graphic is None:
        return Graphic("other", uri=attr(graphic_data, "uri"), path=path)
    return graphic


def _xfrm(element: Element | None) -> dict:
    if element is None:
        return {}
    off, ext = child(element, "off"), child(element, "ext")
    out = {"offset": (_int(off, "x"), _int(off, "y")), "extent": (_int(ext, "cx"), _int(ext, "cy")),
           "rotation": _int(element, "rot"), "flip_h": _bool(element, "flipH"), "flip_v": _bool(element, "flipV")}
    ch_off, ch_ext = child(element, "chOff"), child(element, "chExt")
    if ch_ext is not None:
        out["child_offset"] = (_int(ch_off, "x"), _int(ch_off, "y"))
        out["child_extent"] = (_int(ch_ext, "cx"), _int(ch_ext, "cy"))
    return out


#: The effects the shared renderer draws (``ooxml_common.drawingml.effect``); any other
#: child of ``a:effectLst`` is warned of.
_DRAWN_EFFECTS = {"outerShdw", "innerShdw", "glow", "softEdge"}


def _unsupported(sp_pr: Element | None) -> list[str]:
    out = []
    effects = child(sp_pr, "effectLst")
    for node in children(effects):
        if local_name(node.tag) not in _DRAWN_EFFECTS:
            out.append(f"effect {local_name(node.tag)}")
    for name in ("effectDag", "scene3d", "sp3d"):
        if child(sp_pr, name) is not None:
            out.append(name)
    return out


def _paint(sp_pr: Element | None, style: Element | None = None) -> dict:
    """What ``spPr`` (and a ``wps:style``) say about painting, read by the shared reader
    (``ooxml_common.drawingml.read``) and left unresolved: the fill, the outline, the
    effects, the theme references and the geometry."""
    return {"geometry": read.parse_geometry_spec(sp_pr), "fill": read.parse_fill(sp_pr),
            "line": read.parse_line(child(sp_pr, "ln")), "effects": read.parse_effect_list(child(sp_pr, "effectLst")),
            "style": read.parse_shape_style(style)}


def _graphic(element: Element, path: str, blocks: Callable[[Element, str], list]) -> Graphic | None:
    name = local_name(element.tag)
    if name == "chart" and element.get(_R + "id"):
        # ``c:chart r:id``: the chart part, related from the part the drawing is in.
        return Graphic("chart", relationship=element.get(_R + "id"), path=path)
    if name == "relIds" and element.get(_R + "dm"):
        # ``dgm:relIds``: SmartArt, named by its data model; the drawing Word cached for
        # it is found from there (``ooxml_common.drawingml.diagram``).
        return Graphic("diagram", relationship=element.get(_R + "dm"), path=path)
    if name == "pic":
        sp_pr = child(element, "spPr")
        blip = child(child(element, "blipFill"), "blip")
        unsupported = _unsupported(sp_pr)
        if child(child(element, "blipFill"), "srcRect") is not None and any(
                attr(child(child(element, "blipFill"), "srcRect"), side) not in (None, "0")
                for side in ("l", "t", "r", "b")):
            unsupported.append("cropped")
        paint = _paint(sp_pr, child(element, "style"))
        paint.pop("fill")
        return Graphic("picture", relationship=blip.get(_R_EMBED) if blip is not None else None,
                       unsupported=tuple(unsupported), path=path, **paint, **_xfrm(child(sp_pr, "xfrm")))
    if name == "wsp":
        sp_pr = child(element, "spPr")
        body = child(element, "bodyPr")
        box = child(child(element, "txbx"), "txbxContent")
        text: tuple = ()
        if box is not None:
            text = tuple(blocks(box, f"{path}/wps:txbx/w:txbxContent"))
        body_values: dict = {}
        if body is not None:
            for key in ("lIns", "tIns", "rIns", "bIns"):
                if attr(body, key) is not None:
                    body_values[key] = attr_int(body, key)
            for key in ("anchor", "vert", "wrap", "rot", "anchorCtr"):
                if attr(body, key) is not None:
                    body_values[key] = attr(body, key)
            fit = next((local_name(c.tag) for c in children(body)
                        if local_name(c.tag) in ("noAutofit", "spAutoFit", "normAutofit")), None)
            if fit:
                body_values["fit"] = fit
        unsupported = _unsupported(sp_pr)
        if child(element, "linkedTxbx") is not None:
            unsupported.append("linked text box")
        warp = attr(child(body, "prstTxWarp"), "prst")
        if warp not in (None, "textNoShape"):
            unsupported.append(f"text warp {warp}")
        return Graphic("shape", text=text, body=body_values, unsupported=tuple(unsupported), path=path,
                       text_rect=read.parse_text_rect(sp_pr) if box is not None else None,
                       **_paint(sp_pr, child(element, "style")), **_xfrm(child(sp_pr, "xfrm")))
    if name == "graphicFrame":
        # A group's chart or diagram: the frame's ``wpg:xfrm`` places what its
        # ``a:graphicData`` holds.
        data = child(child(element, "graphic"), "graphicData")
        content = next(iter(children(data)), None) if data is not None else None
        inner = _graphic(content, path, blocks) if content is not None else None
        if inner is None:
            return None
        return dataclasses.replace(inner, **_xfrm(child(element, "xfrm")))
    if name in ("wgp", "grpSp"):
        group_pr = child(element, "grpSpPr")
        members = []
        counts: dict[str, int] = {}
        for node in children(element):
            local = local_name(node.tag)
            if local in ("wsp", "pic", "grpSp", "graphicFrame"):
                counts[local] = counts.get(local, 0) + 1
                prefix = {"wsp": "wps", "pic": "pic", "grpSp": "wpg", "graphicFrame": "wpg"}[local]
                member = _graphic(node, f"{path}/{prefix}:{local}[{counts[local]}]", blocks)
                members.append(member if member is not None else Graphic("other", uri=local, path=path))
        return Graphic("group", children=tuple(members), unsupported=tuple(_unsupported(group_pr)), path=path,
                       fill=read.parse_fill(group_pr), effects=read.parse_effect_list(child(group_pr, "effectLst")),
                       **_xfrm(child(group_pr, "xfrm")))
    return None
