"""``w:rPr`` and ``w:pPr`` as flat maps of *declared* properties, for the cascade.

The dataclasses in :mod:`docx2svg.model` carry a curated, typed view of a few properties.
The cascade needs something else: every property a level *declared*, and nothing it did
not, so that "absent" and "declared false" stay distinguishable -- the whole of toggle
resolution turns on that difference.  So each level is read into a ``dict`` keyed by a
flat name, and :mod:`docx2svg.resolve` merges those dicts.

Key vocabulary
--------------

Run properties (``w:rPr``):

* ``rFonts.ascii`` / ``rFonts.hAnsi`` / ``rFonts.eastAsia`` / ``rFonts.cs`` -- a
  :class:`FontRef`.  The explicit face and the theme reference of one slot are read
  *together*, because Word treats them as one property (measured, ROADMAP.md, "Style
  inheritance"): on one element the theme attribute wins over the explicit one, and a
  level that declares either replaces both as inherited from below.
* ``rFonts.hint`` -- ``default`` / ``eastAsia`` / ``cs``.
* toggles (:data:`TOGGLES`) and other on/off properties -- ``bool``.
* ``sz``, ``szCs``, ``kern``, ``spacing``, ``w``, ``position`` -- ``int``.
* ``color``, ``u``, ``vertAlign``, ``highlight``, ``effect`` -- ``str`` (the ``w:val``).
* ``bdr`` -- ``(val, sz, space)``: the run border's style, width in eighths of a point
  and distance from the text in points, as ``pBdr``'s sides are read.
* ``shd`` -- ``(val, color, fill)``: shading's pattern, pattern colour and fill, as
  written (``auto`` kept).  Drawn, not laid out (Phase 5).
* ``lang.val`` / ``lang.eastAsia`` / ``lang.bidi`` -- ``str``.

Paragraph properties (``w:pPr``):

* ``spacing.before`` / ``.after`` / ``.line`` / ``.beforeLines`` / ``.afterLines`` --
  ``int``; ``spacing.lineRule`` -- ``str``; ``spacing.beforeAutospacing`` /
  ``.afterAutospacing`` -- ``bool``.
* ``ind.left`` / ``.right`` / ``.firstLine`` / ``.hanging`` -- ``int`` (``w:start`` /
  ``w:end`` are read as ``left`` / ``right``).
* ``jc``, ``textAlignment`` -- ``str``; ``outlineLvl`` -- ``int``.
* ``numPr.numId`` / ``numPr.ilvl`` -- ``int``.
* ``pBdr.top`` / ``.bottom`` / ``.between``... -- ``(val, sz, space, color)``: style,
  width in eighths of a point, distance from the text in points, colour.
* ``shd`` -- ``(val, color, fill)``, as for a run.
* ``tabs`` -- a tuple of :class:`~docx2svg.model.TabStop`, *including* ``clear`` stops,
  which the cascade uses to remove an inherited stop.
* ``framePr`` -- a ``dict`` of its attributes by local name, as written.
* ``keepNext``, ``keepLines``, ``pageBreakBefore``, ``widowControl``,
  ``contextualSpacing``, ``bidi``, ``snapToGrid``, ``suppressAutoHyphens``,
  ``mirrorIndents`` -- ``bool``.

Table, row and cell properties (``w:tblPr``, ``w:trPr``, ``w:tcPr``), read the same way
for a table, its table style and the style's conditional formats (Phase 6, tables):

* widths -- ``tblW``, ``tblInd``, ``tblCellSpacing``, ``tcW``, ``wBefore``, ``wAfter``
  and each margin (``tblCellMar.left``, ``tcMar.top``...) -- ``(w, type)``: the value
  and ``dxa`` / ``pct`` / ``auto`` / ``nil`` as written (``dxa`` where the type is not);
* borders -- ``tblBorders.top`` ... ``.insideV``, ``tcBorders.top`` ... ``.tl2br`` --
  ``(val, sz, space, color)`` as ``pBdr``'s sides;
* ``jc``, ``tblLayout``, ``vAlign``, ``vMerge`` (``restart`` / ``continue``),
  ``hMerge``, ``textDirection`` -- ``str``; ``gridSpan``, ``gridBefore``,
  ``gridAfter`` -- ``int``; ``trHeight`` -- ``(val, hRule)``, the rule ``atLeast`` where
  unwritten; ``shd`` as for a run; ``tblpPr`` -- a ``dict`` of its attributes (a floating
  table); ``cantSplit``, ``tblHeader``, ``noWrap``, ``hidden``, ``bidiVisual`` -- ``bool``.
"""

from __future__ import annotations

from dataclasses import dataclass
from xml.etree.ElementTree import Element

from ooxml_common.drawingml import read

from ..model import TabStop
from ..xmlutil import attr, attr_bool, attr_int, child, children, local_name

#: ECMA-376 Part 1, 17.7.3: the run properties the spec calls *toggle properties*.
#: How Word combines them across the style hierarchy is measured, not taken from the
#: spec -- see :mod:`docx2svg.resolve.cascade`.
TOGGLES = frozenset({
    "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow",
    "emboss", "imprint", "vanish",
})

_RUN_BOOLS = TOGGLES | {"webHidden", "specVanish", "noProof", "snapToGrid", "rtl", "cs", "oMath"}
_RUN_INTS = ("sz", "szCs", "kern", "spacing", "w", "position")
_RUN_STRINGS = ("color", "u", "vertAlign", "highlight", "effect", "em")

_PARA_BOOLS = frozenset({
    "keepNext", "keepLines", "pageBreakBefore", "widowControl", "contextualSpacing", "bidi",
    "snapToGrid", "suppressAutoHyphens", "mirrorIndents", "suppressLineNumbers",
    "adjustRightInd", "wordWrap", "kinsoku", "overflowPunct", "autoSpaceDE", "autoSpaceDN",
})

SLOTS = ("ascii", "hAnsi", "eastAsia", "cs")
#: The theme attribute that pairs with each explicit slot.  Note ``cstheme``'s spelling.
THEME_ATTRIBUTE = {"ascii": "asciiTheme", "hAnsi": "hAnsiTheme",
                   "eastAsia": "eastAsiaTheme", "cs": "cstheme"}


@dataclass(frozen=True)
class FontRef:
    """One ``w:rFonts`` slot as one element declared it.

    ``theme`` is an ``ST_Theme`` value (``minorHAnsi``, ``majorBidi``...), ``face`` an
    explicit family name.  Both may be present on one element; the theme reference then
    wins (ECMA-376 17.3.2.26, and measured: case ``f04``).
    """

    face: str | None = None
    theme: str | None = None

    def __str__(self) -> str:
        if self.theme:
            return f"theme {self.theme}" + (f" (over {self.face!r})" if self.face else "")
        return repr(self.face)


def _on(element: Element) -> bool:
    return attr_bool(element)


def read_run_properties(element: Element | None) -> dict[str, object]:
    """Every property ``element`` (a ``w:rPr``) declares, and nothing else."""
    out: dict[str, object] = {}
    if element is None:
        return out
    for node in children(element):
        name = local_name(node.tag)
        if name == "rFonts":
            for slot in SLOTS:
                face = attr(node, slot)
                theme = attr(node, THEME_ATTRIBUTE[slot])
                if face is not None or theme is not None:
                    out[f"rFonts.{slot}"] = FontRef(face, theme)
            hint = attr(node, "hint")
            if hint is not None:
                out["rFonts.hint"] = hint
        elif name == "lang":
            for key in ("val", "eastAsia", "bidi"):
                value = attr(node, key)
                if value is not None:
                    out[f"lang.{key}"] = value
        elif name == "bdr":
            out[name] = (attr(node, "val"), attr_int(node, "sz", 0) or 0, attr_int(node, "space", 0) or 0)
        elif name == "shd":
            out[name] = (attr(node, "val"), attr(node, "color"), attr(node, "fill"))
        elif name in _RUN_BOOLS:
            out[name] = _on(node)
        elif name in _RUN_INTS:
            value = attr_int(node, "val")
            if value is not None:
                out[name] = value
        elif name in _RUN_STRINGS:
            value = attr(node, "val")
            if value is not None:
                out[name] = value
        elif name == "textFill":
            # Word 2010's text effects (``w14:textFill``, ``w14:textOutline``): DrawingML
            # fills and outlines in the w14 namespace, which the shared reader reads by
            # their local names.
            out[name] = read.parse_fill(node)
        elif name == "textOutline":
            out[name] = read.parse_line(node)
        # rStyle is structural, not a property: the cascade reads it separately.
    return out


def _tabs(node: Element) -> tuple[TabStop, ...]:
    stops = []
    for tab in children(node, "tab"):
        position = attr_int(tab, "pos")
        if position is not None:
            stops.append(TabStop(position, attr(tab, "val") or "left", attr(tab, "leader")))
    return tuple(stops)


def read_paragraph_properties(element: Element | None) -> dict[str, object]:
    """Every property ``element`` (a ``w:pPr``) declares, except its ``w:rPr``/``w:sectPr``."""
    out: dict[str, object] = {}
    if element is None:
        return out
    for node in children(element):
        name = local_name(node.tag)
        if name == "spacing":
            for key in ("before", "after", "line", "beforeLines", "afterLines"):
                value = attr_int(node, key)
                if value is not None:
                    out[f"spacing.{key}"] = value
            rule = attr(node, "lineRule")
            if rule is not None:
                out["spacing.lineRule"] = rule
            for key in ("beforeAutospacing", "afterAutospacing"):
                if attr(node, key) is not None:
                    out[f"spacing.{key}"] = attr_bool(node, key)
        elif name == "ind":
            for key, aliases in (("left", ("left", "start")), ("right", ("right", "end")),
                                 ("firstLine", ("firstLine",)), ("hanging", ("hanging",))):
                for alias in aliases:
                    value = attr_int(node, alias)
                    if value is not None:
                        out[f"ind.{key}"] = value
                        break
        elif name == "numPr":
            numbering = attr_int(child(node, "numId"), "val")
            level = attr_int(child(node, "ilvl"), "val")
            if numbering is not None:
                out["numPr.numId"] = numbering
            if level is not None:
                out["numPr.ilvl"] = level
        elif name == "tabs":
            out["tabs"] = _tabs(node)
        elif name == "framePr":
            # A frame: its attributes by local name (``dropCap``, ``lines``, ``hSpace``,
            # ``hAnchor``...), as written (:mod:`docx2svg.paginate`, drop caps).
            out["framePr"] = {local_name(key): value for key, value in node.attrib.items()}
        elif name == "pBdr":
            # Per side: (style, width in eighths of a point, space in points, colour).  A
            # border takes vertical room, so it is part of what the line stack needs; the
            # colour decides whether two paragraphs' borders are one box.
            for side in children(node):
                out[f"pBdr.{local_name(side.tag)}"] = (
                    attr(side, "val"), attr_int(side, "sz", 0) or 0, attr_int(side, "space", 0) or 0,
                    attr(side, "color"),
                )
        elif name == "shd":
            out[name] = (attr(node, "val"), attr(node, "color"), attr(node, "fill"))
        elif name in ("jc", "textAlignment"):
            value = attr(node, "val")
            if value is not None:
                out[name] = value
        elif name == "outlineLvl":
            value = attr_int(node, "val")
            if value is not None:
                out[name] = value
        elif name in _PARA_BOOLS:
            out[name] = _on(node)
    return out


# -- tables ------------------------------------------------------------------------------


def _width(node: Element) -> tuple[int, str]:
    """``(w:w, w:type)``: a ``pct`` value written as ``"50%"`` is read in fiftieths."""
    kind = attr(node, "type") or "dxa"
    raw = attr(node, "w") or "0"
    if raw.endswith("%"):
        try:
            return round(float(raw[:-1]) * 50), "pct"
        except ValueError:
            return 0, kind
    try:
        return int(float(raw)), kind
    except ValueError:
        return 0, kind


def _border(node: Element) -> tuple:
    return (attr(node, "val"), attr_int(node, "sz", 0) or 0, attr_int(node, "space", 0) or 0, attr(node, "color"))


_TABLE_WIDTHS = frozenset({"tblW", "tblInd", "tblCellSpacing", "tcW", "wBefore", "wAfter"})
_TABLE_STRINGS = frozenset({"jc", "vAlign", "textDirection", "hMerge"})
_TABLE_INTS = frozenset({"gridSpan", "gridBefore", "gridAfter", "tblStyleRowBandSize", "tblStyleColBandSize"})
_TABLE_BOOLS = frozenset({"cantSplit", "tblHeader", "noWrap", "hidden", "bidiVisual", "tcFitText", "hideMark"})


def read_table_properties(element: Element | None) -> dict[str, object]:
    """Every property ``element`` -- a ``w:tblPr``, ``w:trPr`` or ``w:tcPr`` -- declares
    (the vocabulary is shared: a row's ``jc`` and a table's are both ``jc``)."""
    out: dict[str, object] = {}
    if element is None:
        return out
    for node in children(element):
        name = local_name(node.tag)
        if name in _TABLE_WIDTHS:
            out[name] = _width(node)
        elif name in ("tblCellMar", "tcMar"):
            for side in children(node):
                key = {"start": "left", "end": "right"}.get(local_name(side.tag), local_name(side.tag))
                out[f"{name}.{key}"] = _width(side)
        elif name in ("tblBorders", "tcBorders"):
            for side in children(node):
                key = {"start": "left", "end": "right"}.get(local_name(side.tag), local_name(side.tag))
                out[f"{name}.{key}"] = _border(side)
        elif name == "tblLayout":
            out[name] = attr(node, "type") or "autofit"
        elif name == "trHeight":
            out[name] = (attr_int(node, "val", 0) or 0, attr(node, "hRule") or "atLeast")
        elif name == "vMerge":
            out[name] = attr(node, "val") or "continue"
        elif name == "shd":
            out[name] = (attr(node, "val"), attr(node, "color"), attr(node, "fill"))
        elif name == "tblpPr":
            out[name] = {local_name(key): value for key, value in node.attrib.items()}
        elif name in _TABLE_STRINGS:
            value = attr(node, "val")
            if value is not None:
                out[name] = value
        elif name in _TABLE_INTS:
            value = attr_int(node, "val")
            if value is not None:
                out[name] = value
        elif name in _TABLE_BOOLS:
            out[name] = _on(node)
    return out
