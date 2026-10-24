"""Read a run's VML object (``w:pict``, ``w:object``, an ``mc:AlternateContent``'s VML
``mc:Fallback``) as far as the layout can take it, instead of stopping at it.

VML (ECMA-376 Part 4, 19.1) is how Word writes a watermark (a ``v:shape`` with a
``v:textpath``, in a header's ``w:sdt`` of the ``Watermarks`` gallery), an embedded
object's preview picture (``v:shape`` with ``v:imagedata``) and the shapes of older
documents.  A shape's CSS ``style`` says where it is and how large:

* **absolutely positioned** (``position:absolute``): a floating drawing, read into the
  same :class:`~docx2svg.model.Anchor` a ``wp:anchor`` is -- ``width`` / ``height`` its
  extent; ``mso-position-horizontal`` / ``-vertical`` an alignment (``left``, ``center``,
  ``right``, ``top``, ``bottom``, ``inside``, ``outside``), else ``margin-left`` /
  ``margin-top`` an offset, against ``mso-position-horizontal-relative`` /
  ``-vertical-relative`` (``margin``, ``page``, ``text`` -- the column and the paragraph
  --, ``char``, ``line`` and the margin areas, as ``wp:positionH`` / ``wp:positionV``
  name them); ``z-index`` below 0 behind the text; ``w10:wrap/@type`` the wrap
  (``square``, ``tight``, ``through``, ``topAndBottom``, ``none``; none stated: in front
  of the text, ``wrapNone``);
* **inline** (anything else): a character of its line as tall and wide as its box, as a
  ``wp:inline`` is -- the mark ``drawing:cx:cy:0:0:0:0``.

What it shows: a picture (``v:imagedata/@r:id``: :class:`~docx2svg.model.Graphic`
``picture``), a text path (``v:textpath/@string``: ``textpath``, drawn by
:mod:`docx2svg.drawing` as its text stretched to the box, in its fill colour and opacity
and turned by ``rotation``; ``body["fit"]`` whether it fits its shape, ``fitshape``), or
anything else (``other``, drawn as a marked placeholder of its box).  Where Word draws a
text path that fits its shape is measured (:data:`docx2svg.drawing.TEXT_PATH_INSETS`); one
that does not is not.

Standard library only.
"""

from __future__ import annotations

import re
from xml.etree.ElementTree import Element

from ..model import Anchor, AnchorPosition, Graphic
from ..xmlutil import attr, child, children, local_name

_VML = "urn:schemas-microsoft-com:vml"
_R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
#: The VML elements that draw something (``v:shapetype`` only defines one).
_SHAPES = ("shape", "rect", "roundrect", "oval", "line", "polyline", "arc", "curve", "image", "group")
#: EMU per unit of a CSS length (ECMA-376 Part 4, 19.1.2.19); a bare number is px (96 an inch).
_EMU = {"pt": 12700, "in": 914400, "cm": 360000, "mm": 36000, "px": 9525, "pc": 152400, "emu": 1}
_RELATIVE_H = {"margin": "margin", "page": "page", "text": "column", "char": "character",
               "left-margin-area": "leftMargin", "right-margin-area": "rightMargin",
               "inner-margin-area": "insideMargin", "outer-margin-area": "outsideMargin"}
_RELATIVE_V = {"margin": "margin", "page": "page", "text": "paragraph", "line": "line",
               "top-margin-area": "topMargin", "bottom-margin-area": "bottomMargin",
               "inner-margin-area": "insideMargin", "outer-margin-area": "outsideMargin"}
_WRAPS = {"square": "wrapSquare", "tight": "wrapTight", "through": "wrapThrough",
          "topAndBottom": "wrapTopAndBottom", "none": "wrapNone"}
#: VML's colour names (ECMA-376 Part 4, 19.1.2.3: the sixteen HTML 4 names).
_COLOURS = {"black": "000000", "silver": "C0C0C0", "gray": "808080", "white": "FFFFFF", "maroon": "800000",
            "red": "FF0000", "purple": "800080", "fuchsia": "FF00FF", "green": "008000", "lime": "00FF00",
            "olive": "808000", "yellow": "FFFF00", "navy": "000080", "blue": "0000FF", "teal": "008080",
            "aqua": "00FFFF"}


def style_of(element: Element | None) -> dict[str, str]:
    """A VML ``style`` attribute's declarations, by lowercased property."""
    out: dict[str, str] = {}
    for declaration in (attr(element, "style") or "").split(";"):
        name, _, value = declaration.partition(":")
        if name.strip() and value.strip():
            out[name.strip().lower()] = value.strip().strip('"').replace("&quot;", "")
    return out


def length_emu(value: str | None) -> int | None:
    """A CSS length (``412.4pt``, ``1in``, ``0``) in EMU; ``None`` where it is not one."""
    if not value:
        return None
    found = re.fullmatch(r"\s*(-?[0-9]*\.?[0-9]+)\s*([a-z]*)\s*", value.lower())
    if found is None:
        return None
    number, unit = float(found.group(1)), found.group(2) or "px"
    if unit not in _EMU:
        return None
    return round(number * _EMU[unit])


def colour(value: str | None) -> str | None:
    """A VML colour (``#C0C0C0``, ``silver``, ``#ccc``, ``black [3213]``) as ``RRGGBB``."""
    if not value:
        return None
    value = value.split()[0].strip().lower()
    if value.startswith("#"):
        digits = value[1:]
        if len(digits) == 3:
            digits = "".join(c * 2 for c in digits)
        return digits.upper() if re.fullmatch(r"[0-9a-f]{6}", digits) else None
    return _COLOURS.get(value)


def fraction(value: str | None, default: float = 1.0) -> float:
    """A VML fraction: ``.5``, ``50%`` or ``32768f`` (sixty-five-thousand-five-hundred-and-
    thirty-sixths)."""
    if not value:
        return default
    value = value.strip().lower()
    try:
        if value.endswith("f"):
            return float(value[:-1]) / 65536
        if value.endswith("%"):
            return float(value[:-1]) / 100
        return float(value)
    except ValueError:
        return default


def vml_shape(node: Element) -> tuple[Element, str] | None:
    """The first VML element a run's object draws -- in ``w:pict`` or ``w:object``, or
    the ``w:pict`` of an ``mc:AlternateContent``'s ``mc:Fallback`` -- and its path from
    the object (:mod:`docx2svg.paths`)."""
    name = local_name(node.tag)
    if name == "AlternateContent":
        fallback = child(node, "Fallback")
        inner = next((c for c in children(fallback) if local_name(c.tag) in ("pict", "object")), None)
        found = vml_shape(inner) if inner is not None else None
        return (found[0], f"mc:Fallback/w:{local_name(inner.tag)}[1]/{found[1]}") if found else None
    if name not in ("pict", "object"):
        return None
    counts: dict[str, int] = {}
    for element in children(node):
        local = local_name(element.tag)
        counts[local] = counts.get(local, 0) + 1
        if element.tag.startswith("{" + _VML + "}") and local in _SHAPES:
            return element, f"v:{local}[{counts[local]}]"
    return None


#: The VML elements that are a preset geometry, and the shape types Word writes for one
#: (``#_x0000_t1`` a rectangle, ``t202`` a text box): drawn as the DrawingML preset.
_PRESETS = {"rect": "rect", "roundrect": "roundRect", "oval": "ellipse"}
_SHAPE_TYPES = {"_x0000_t1": "rect", "_x0000_t202": "rect", "_x0000_t2": "roundRect", "_x0000_t3": "ellipse"}
#: A text box's insets where ``v:textbox/@inset`` states none: 0.1 in left and right,
#: 0.05 in top and bottom (ECMA-376 Part 4, 19.1.2.22).
_INSETS = (91440, 45720, 91440, 45720)


def _shape(shape: Element, preset: str, path: str, blocks, common: dict) -> Graphic:
    """A VML rectangle, rounded rectangle, oval or text box as the DrawingML shape it is:
    its fill (``fillcolor``, ``filled``, ``v:fill/@opacity``; white by default), its
    outline (``strokecolor``, ``strokeweight``, ``stroked``; black, 0.75 pt by default)
    and a text box's content (``v:textbox/w:txbxContent``, with its insets)."""
    from ooxml_common.drawingml import source as s
    from ooxml_common.drawingml.model import ColorTransform, SrgbColor

    fill_element = child(shape, "fill")
    if attr(shape, "filled") in ("f", "false") or attr(fill_element, "on") in ("f", "false"):
        fill = s.SourceNoFill()
    else:
        alpha = fraction(attr(fill_element, "opacity"))
        transforms = [ColorTransform("alpha", alpha * 100000)] if alpha < 1 else []
        fill = s.SourceSolidFill(SrgbColor(colour(attr(shape, "fillcolor")) or "FFFFFF", transforms))
    stroke = child(shape, "stroke")
    if attr(shape, "stroked") in ("f", "false") or attr(stroke, "on") in ("f", "false"):
        line = s.SourceOutline(fill=s.SourceNoFill())
    else:
        weight = length_emu(attr(shape, "strokeweight") or "0.75pt") or 9525
        line = s.SourceOutline(width=weight, fill=s.SourceSolidFill(
            SrgbColor(colour(attr(shape, "strokecolor") or attr(stroke, "color")) or "000000")))
    box = child(shape, "textbox")
    content = child(box, "txbxContent")
    text: tuple = ()
    body: dict = {}
    if content is not None and blocks is not None:
        text = tuple(blocks(content, f"{path}/v:textbox/w:txbxContent"))
        insets = [length_emu(part) for part in (attr(box, "inset") or "").split(",")]
        for key, default, value in zip(("lIns", "tIns", "rIns", "bIns"), _INSETS, insets + [None] * 4):
            body[key] = value if value is not None else default
    elif content is not None:
        body = dict(zip(("lIns", "tIns", "rIns", "bIns"), _INSETS))
    return Graphic("shape", geometry=("preset", preset, {}), fill=fill, line=line, text=text, body=body,
                   extent=(0, 0), **common)


#: The shape type Word writes for a watermark's text (``o:spt`` 136, plain text): its
#: ``v:shapetype`` states ``v:textpath fitshape="t"`` (Word for Mac 16.106's watermark,
#: measured), which a shape of the type inherits where its own shape type is not at hand.
_FITTED_TYPES = {"_x0000_t136"}


def _fits(shape: Element, textpath: Element, types: dict) -> bool:
    """Whether a text path is stretched to fill its shape (``fitshape``, ECMA-376 Part 4,
    19.1.2.20): as its own ``v:textpath`` says, else its ``v:shapetype``'s, else whether
    the type is :data:`_FITTED_TYPES`'s."""
    own = attr(textpath, "fitshape")
    kind = (attr(shape, "type") or "").lstrip("#")
    if own is None and kind in types:
        own = attr(child(types[kind], "textpath"), "fitshape")
    if own is None:
        return kind in _FITTED_TYPES
    return own in ("t", "true", "1", "on")


def _graphic(shape: Element, path: str, blocks=None, types: dict | None = None) -> Graphic:
    style = style_of(shape)
    rotation = 0
    try:
        rotation = round(float(style.get("rotation", "0").rstrip("fd")) * 60000)
    except ValueError:
        rotation = 0
    flip = style.get("flip", "")
    common = {"rotation": rotation, "flip_h": "x" in flip, "flip_v": "y" in flip, "path": path}
    image = child(shape, "imagedata")
    if image is not None and (image.get(_R_ID) or attr(image, "id")):
        return Graphic("picture", relationship=image.get(_R_ID) or attr(image, "id"), **common)
    textpath = child(shape, "textpath")
    preset = _PRESETS.get(local_name(shape.tag)) or _SHAPE_TYPES.get((attr(shape, "type") or "").lstrip("#"))
    if preset is not None and (textpath is None or not attr(textpath, "string")):
        return _shape(shape, preset, path, blocks, common)
    if textpath is not None and attr(textpath, "string"):
        text_style = style_of(textpath)
        fill = child(shape, "fill")
        filled = attr(shape, "filled") not in ("f", "false")
        body = {"string": attr(textpath, "string"),
                "font": text_style.get("font-family") or "Calibri",
                "bold": text_style.get("font-weight") in ("bold", "700", "800", "900"),
                "italic": text_style.get("font-style") == "italic",
                "fill": (colour(attr(shape, "fillcolor")) or "FFFFFF") if filled else None,
                "opacity": fraction(attr(fill, "opacity")) if fill is not None else 1.0,
                "stroke": colour(attr(shape, "strokecolor")) or "000000" if attr(shape, "stroked") not in (
                    "f", "false") else None,
                "fit": _fits(shape, textpath, types or {})}
        return Graphic("textpath", body=body, **common)
    return Graphic("other", uri=f"vml:{local_name(shape.tag)}", path=path)


def read_vml(node: Element, path: str, blocks=None) -> tuple[str, object] | None:
    """A run's VML object: ``("anchor", Anchor)`` for an absolutely positioned shape,
    ``("inline", (mark, picture relationship or None))`` for one in the line, or
    ``None`` where it is not VML or states no size.  ``blocks(element, path)`` reads a
    text box's ``w:txbxContent`` (as :func:`docx2svg.parse.drawing.read_anchor`'s)."""
    found = vml_shape(node)
    if found is None:
        return None
    shape, inner = found
    path = f"{path}/{inner}"
    style = style_of(shape)
    width, height = length_emu(style.get("width")), length_emu(style.get("height"))
    if not width or not height or width < 0 or height < 0:
        return None
    types = {attr(element, "id"): element for element in children(node) if local_name(element.tag) == "shapetype"
             and attr(element, "id")}
    graphic = _graphic(shape, path, blocks, types)
    if style.get("position") != "absolute":
        shown = graphic.relationship if graphic.kind == "picture" else graphic if graphic.kind in (
            "shape", "textpath") else None
        return "inline", (f"drawing:{width}:{height}:0:0:0:0", shown)

    def position(kind: str, relatives: dict, default: str, aligns: tuple, offset_key: str) -> AnchorPosition:
        relative = relatives.get(style.get(f"mso-position-{kind}-relative", default), relatives[default])
        align = style.get(f"mso-position-{kind}")
        if align in aligns:
            return AnchorPosition(relative=relative, align=align)
        return AnchorPosition(relative=relative, offset=(length_emu(style.get(offset_key)) or 0)
                              + (length_emu(style.get("left" if kind == "horizontal" else "top")) or 0))

    h = position("horizontal", _RELATIVE_H, "text", ("left", "center", "right", "inside", "outside"), "margin-left")
    v = position("vertical", _RELATIVE_V, "text", ("top", "center", "bottom", "inside", "outside"), "margin-top")
    wrap_element = next((c for c in children(shape) if local_name(c.tag) == "wrap"), None)
    wrap = _WRAPS.get(attr(wrap_element, "type") or "none", "wrapNone")
    try:
        z = int(style.get("z-index", "0"))
    except ValueError:
        z = 0
    return "anchor", Anchor(extent=(width, height), h=h, v=v, wrap=wrap, behind=z < 0,
                            relative_height=abs(z) % (1 << 31), graphic=graphic,
                            identifier=attr(shape, "id"), vml=True)
