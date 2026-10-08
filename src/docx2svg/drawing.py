"""What a drawing's graphic paints: DrawingML shapes, groups and pictures as SVG, in
device px, through ``ooxml-common``'s shared DrawingML renderers.

``ooxml_common.drawingml`` reads the XML (:mod:`docx2svg.parse.drawing` calls its reader),
applies colour transforms and draws geometry, fills, outlines, arrowheads and effects --
the code pptx2svg draws a slide's shapes with, taking the application as a parameter where
Word and PowerPoint measurably differ (``ooxml_common.drawingml.rules.WORD``: colour
composition, gradient geometry, dashes and caps, arrowheads, the default join, the
pattern phase; ROADMAP.md, "DrawingML drawn by the shared renderers").  What is Word's
alone stays here: a theme colour through ``w:clrSchemeMapping``; a shape's theme style
references (``wps:style``) through the theme's format scheme, ``phClr`` the reference's
colour; a group's fill for a member that says ``a:grpFill``; where the shape goes (the
anchor's rules, ROADMAP.md "Floating drawings -- measured") and the order it paints in.

What it cannot draw honestly (a picture fill, a 3-D scene, an effect the shared renderer
does not have) it says so, and the caller warns and, where nothing could be drawn, draws
a marked placeholder of the extent.

Measured on ``tools/make_drawing_probe.py`` (ROADMAP.md, "Floating drawings --
measured"): a shape is drawn at its extent truncated to whole twips on each axis; a
group's children in their child coordinates scaled from ``a:chOff`` / ``a:chExt`` onto
that box.  On ``tools/make_dml_probe.py``: the Word rules above, the style references
(1-based indexes, ``phClr``) and the group fill.
"""

from __future__ import annotations

import dataclasses
from dataclasses import dataclass, field
from fractions import Fraction

from ooxml_common.drawingml import color as dml_color
from ooxml_common.drawingml import effect as dml_effect
from ooxml_common.drawingml import fill as dml_fill
from ooxml_common.drawingml import geometry as dml_geometry
from ooxml_common.drawingml import model as m
from ooxml_common.drawingml import source as s
from ooxml_common.drawingml.rules import WORD
from ooxml_common.drawingml.svg import Defs, num

from .model import Graphic

#: EMU per device px (914,400 EMU an inch, 300 px an inch).
EMU_PX = Fraction(300, 914400)
#: The device grid the renderers draw on: Word's PDF export's (ROADMAP.md, Phase 0.5).
DPI = 300


@dataclass
class Primitive:
    """One thing a drawing paints, in page device px: a ``path`` (SVG path data, filled
    or stroked), an ``image`` (a picture part in a box), a ``placeholder``, or ``markup``
    (an element the shared renderer wrote whole: an arrowhead).  ``begin`` and ``end``
    bracket a shape's primitives in a group carrying its effects (``attrs``, a filter).

    A path's paint is ``attrs``, the SVG attributes the shared renderer wrote; ``fill``
    and ``stroke`` are its colours as ``RRGGBB`` where they are one colour (``None`` for a
    gradient, a pattern, or none), for the tools that read the model."""

    kind: str
    path: str
    d: str | None = None
    fill: str | None = None
    stroke: str | None = None
    stroke_width: Fraction | None = None
    x: Fraction = Fraction(0)
    y: Fraction = Fraction(0)
    width: Fraction = Fraction(0)
    height: Fraction = Fraction(0)
    relationship: str | None = None
    #: An SVG transform about the box (rotation, flips), or ``None``.
    transform: str | None = None
    what: str = ""
    attrs: str = ""
    markup: str = ""


@dataclass
class Drawn:
    """What :func:`draw` makes of a graphic: its primitives in paint order, the text
    boxes to lay out (``(graphic, x, y, width, height)`` of the shape's box, px), and
    what could not be drawn faithfully (warning codes)."""

    primitives: list[Primitive] = field(default_factory=list)
    text_boxes: list[tuple] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


# -- colour ---------------------------------------------------------------------------

#: ``w:clrSchemeMapping``'s attributes and values spelled as DrawingML's colour-map slots
#: and theme slots: Word writes ``t1="dark1"`` where a slide master writes ``tx1="dk1"``.
_MAPPED_KEYS = {"t1": "tx1", "t2": "tx2", "hyperlink": "hlink", "followedHyperlink": "folHlink"}
_MAPPED = {"light1": "lt1", "dark1": "dk1", "light2": "lt2", "dark2": "dk2", "hyperlink": "hlink",
           "followedHyperlink": "folHlink"}
#: The map Word uses where the settings state none.
DEFAULT_COLOR_MAP = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}


def color_map_slots(color_map: dict) -> dict:
    """``w:clrSchemeMapping`` as ``{colour-map slot: theme slot}`` (``tx1`` -> ``dk1``)."""
    out = dict(DEFAULT_COLOR_MAP)
    for key, value in color_map.items():
        out[_MAPPED_KEYS.get(key, key)] = _MAPPED.get(value, value)
    return out


class Colours:
    """Colour, fill, outline and effect resolution for one document: its theme's colours,
    ``w:clrSchemeMapping`` and format scheme, under Word's colour rules
    (``ooxml_common.drawingml.color.WORD``)."""

    def __init__(self, theme: dict, color_map: dict, formats=None, warnings: list[str] | None = None):
        self.theme = theme
        self.slots = color_map_slots(color_map)
        self.formats = formats
        self.warnings = warnings if warnings is not None else []

    def resolve(self, choice, placeholder: m.ResolvedColor | None = None) -> m.ResolvedColor | None:
        """A colour choice as a resolved colour, or ``None`` (warned).  ``placeholder`` is
        what ``phClr`` means: a style reference's colour."""
        if choice is None:
            return None
        if choice.kind == "srgb" and choice.linear is not None:
            return _upper(dml_color.apply_transforms_linear(choice.linear, choice.transforms, dml_color.WORD))
        if choice.kind == "srgb":
            base = choice.hex
        elif choice.kind == "system":
            base = choice.last_color or (choice.value if len(choice.value) == 6 else None)
        elif choice.scheme == "phClr":
            if placeholder is None:
                self.warnings.append("drawing-colour-unknown:phClr")
                return None
            resolved = dml_color.apply_transforms(placeholder.hex, choice.transforms, dml_color.WORD)
            if placeholder.alpha < 1 and all(t.kind != "alpha" for t in choice.transforms):
                resolved = m.ResolvedColor(resolved.hex, placeholder.alpha)
            return _upper(resolved)
        else:
            base = self.theme.get(self.slots.get(choice.scheme, choice.scheme))
            if base is None:
                self.warnings.append(f"drawing-colour-unknown:{choice.scheme}")
                return None
        if not base or len(base.lstrip("#")) != 6:
            self.warnings.append(f"drawing-colour-unknown:{choice.kind}")
            return None
        return _upper(dml_color.apply_transforms(base, choice.transforms, dml_color.WORD))

    def fill(self, fill, placeholder=None, group=None):
        """A fill as read (``ooxml_common.drawingml.source``) as the model's, or ``None``
        where it cannot be drawn (warned).  ``group`` is what ``a:grpFill`` means, the
        enclosing group's fill (none when no group states one: measured)."""
        if fill is None:
            return None
        if isinstance(fill, s.SourceNoFill):
            return m.NoFill()
        if isinstance(fill, s.SourceSolidFill):
            colour = self.resolve(fill.color, placeholder)
            return m.SolidFill(colour) if colour is not None else None
        if isinstance(fill, s.SourceGradientFill):
            stops = []
            for stop in fill.stops:
                colour = self.resolve(stop.color, placeholder)
                if colour is not None:
                    stops.append(m.GradientStop(stop.position, colour))
            if not stops:
                return None
            return m.GradientFill(stops, angle=fill.angle / 60000, gradient_type=fill.gradient_type,
                                  center_x=fill.center_x, center_y=fill.center_y, path=fill.path, focus=fill.focus,
                                  scaled=fill.scaled, rotate_with_shape=fill.rotate_with_shape)
        if isinstance(fill, s.SourcePatternFill):
            foreground = self.resolve(fill.foreground_color, placeholder)
            background = self.resolve(fill.background_color, placeholder)
            if foreground is None or background is None:
                return None
            return m.PatternFill(fill.preset, foreground, background)
        if isinstance(fill, s.SourceGroupFill):
            return group if group is not None else m.NoFill()
        self.warnings.append("drawing-not-drawn:blipFill")
        return None

    def style_fill(self, reference):
        """A ``wps:style`` ``a:fillRef``: the theme's fill style ``idx`` (1-based; from
        1001 the background fill styles), ``phClr`` its colour."""
        if reference is None or reference.idx == 0 or self.formats is None:
            return None
        styles, index = ((self.formats.bg_fill_styles, reference.idx - 1001) if reference.idx >= 1000
                         else (self.formats.fill_styles, reference.idx - 1))
        if not 0 <= index < len(styles):
            return None
        return self.fill(styles[index], self.resolve(reference.color))

    def outline(self, line, reference=None):
        """A shape's ``a:ln`` over its ``a:lnRef`` (the theme's line style ``idx``,
        ``phClr`` its colour), as the model's outline, or ``None`` for no outline."""
        base = None
        if reference is not None and 0 < reference.idx and self.formats is not None:
            if reference.idx - 1 < len(self.formats.line_styles):
                base = self.formats.line_styles[reference.idx - 1]
        if line is None and base is None:
            return None
        placeholder = self.resolve(reference.color) if base is not None else None
        merged = s.SourceOutline()
        for layer in (base, line):
            for name in ("width", "fill", "dash_style", "custom_dash", "line_cap", "line_join", "head_end",
                         "tail_end", "compound"):
                value = getattr(layer, name) if layer is not None else None
                if value is not None:
                    setattr(merged, name, value)
        paint = self.fill(merged.fill, placeholder)
        if paint is None or isinstance(paint, m.NoFill):
            return None
        if not isinstance(paint, (m.SolidFill, m.GradientFill)):
            self.warnings.append("drawing-not-drawn:pattern outline")
            return None
        if merged.compound not in (None, "sng"):
            self.warnings.append(f"drawing-outline-approximate:{merged.compound}")
        return m.Outline(width=merged.width if merged.width is not None else DEFAULT_LINE_EMU, fill=paint,
                         dash_style=merged.dash_style or "solid", custom_dash=merged.custom_dash,
                         line_cap=merged.line_cap, line_join=merged.line_join, head_end=merged.head_end,
                         tail_end=merged.tail_end, compound=merged.compound)

    def effects(self, effects, reference=None):
        """``a:effectLst``, or else the theme's effect style ``a:effectRef`` names
        (1-based: Word draws ``idx="3"`` as the third, measured), as the model's."""
        if effects is None and reference is not None and reference.idx > 0 and self.formats is not None:
            if reference.idx - 1 < len(self.formats.effect_styles):
                effects = self.formats.effect_styles[reference.idx - 1]
        if effects is None:
            return None
        out = m.EffectList()
        shadow = effects.outer_shadow
        if shadow is not None and (colour := self.resolve(shadow.color)) is not None:
            out.outer_shadow = m.OuterShadow(shadow.blur_radius, shadow.distance, shadow.direction / 60000, colour,
                                             shadow.alignment, shadow.rotate_with_shape)
        inner = effects.inner_shadow
        if inner is not None and (colour := self.resolve(inner.color)) is not None:
            out.inner_shadow = m.InnerShadow(inner.blur_radius, inner.distance, inner.direction / 60000, colour)
        if effects.glow is not None and (colour := self.resolve(effects.glow.color)) is not None:
            out.glow = m.Glow(effects.glow.radius, colour)
        if effects.soft_edge is not None:
            out.soft_edge = m.SoftEdge(effects.soft_edge.radius)
        return None if out.is_empty() else out


def _upper(colour: m.ResolvedColor) -> m.ResolvedColor:
    """``#RRGGBB``, as this package has always written a colour."""
    return m.ResolvedColor("#" + colour.hex.lstrip("#").upper(), colour.alpha)


# -- placement ------------------------------------------------------------------------


@dataclass(frozen=True)
class _Turn:
    """One box's rotation and flips about its centre."""

    cx: float
    cy: float
    rotation: float
    flip_h: bool
    flip_v: bool

    def _scale(self) -> str:
        return f"scale({-1 if self.flip_h else 1} {-1 if self.flip_v else 1})"

    def forward(self) -> str:
        parts = [f"translate({num(self.cx)} {num(self.cy)})"]
        if self.rotation:
            parts.append(f"rotate({num(self.rotation)})")
        if self.flip_h or self.flip_v:
            parts.append(self._scale())
        return " ".join(parts + [f"translate({num(-self.cx)} {num(-self.cy)})"])

    def inverse(self) -> str:
        parts = [f"translate({num(self.cx)} {num(self.cy)})"]
        if self.flip_h or self.flip_v:
            parts.append(self._scale())
        if self.rotation:
            parts.append(f"rotate({num(-self.rotation)})")
        return " ".join(parts + [f"translate({num(-self.cx)} {num(-self.cy)})"])


def _turn(graphic: Graphic, x, y, width, height) -> _Turn | None:
    if not (graphic.rotation or graphic.flip_h or graphic.flip_v):
        return None
    return _Turn(float(x + width / 2), float(y + height / 2), graphic.rotation / 60000, graphic.flip_h,
                 graphic.flip_v)


def _transform(turns: tuple) -> str | None:
    """The turns (outermost first) as one SVG transform."""
    return " ".join(turn.forward() for turn in turns) or None


def _frame(turns: tuple) -> dml_fill.ShapeFrame:
    """Where an element under ``turns`` sits on the page: its own rotation and flips,
    and the page's coordinates carried into its user space."""
    own = turns[-1] if turns else None
    return dml_fill.ShapeFrame(rotation=own.rotation if own else 0.0, flip_h=bool(own and own.flip_h),
                               flip_v=bool(own and own.flip_v),
                               page_transform=" ".join(turn.inverse() for turn in reversed(turns)) or None)


# -- a graphic ------------------------------------------------------------------------

#: ``a:ln`` without ``@w``: Word's default outline width, 0.75 pt.
DEFAULT_LINE_EMU = 9525
#: Presets whose ``shape`` path gradient is a ``rect`` one: rectangles (measured on
#: ``rect``); on any other a ``shape`` gradient is drawn as rings of the box's ellipse
#: (measured on ``ellipse``).
_RECTANGULAR = {"rect", "flowChartProcess"}


def draw(graphic: Graphic | None, x: Fraction, y: Fraction, width: Fraction, height: Fraction, path: str, *,
         theme: dict, color_map: dict, formats=None, defs=None, drawn: Drawn | None = None, frames=None,
         part: str | None = None) -> Drawn:
    """What ``graphic`` paints in the box ``x, y, width, height`` (px).  ``defs`` is where
    gradients, patterns and filters go (an ``ooxml_common.drawingml.svg.SvgDefs``: the
    page's); ``formats`` the theme's format scheme; ``frames`` draws a chart or a
    SmartArt diagram (:class:`docx2svg.chart.Frames`), related from ``part`` (the
    document's main part where ``None``)."""
    drawn = drawn if drawn is not None else Drawn()
    colours = Colours(theme, color_map, formats, drawn.warnings)
    _draw(graphic, x, y, width, height, path, colours, defs if defs is not None else Defs(), drawn, None, None, (),
          (frames, part))
    return drawn


def _draw(graphic, x, y, width, height, path, colours: Colours, defs, drawn: Drawn, group_fill, group_box,
          turns: tuple, frames: tuple = (None, None)) -> None:
    warnings = drawn.warnings
    if graphic is not None and graphic.kind in ("chart", "diagram") and frames[0] is not None:
        # Drawn upright in its box -- a chart's or a diagram's frame states no rotation --
        # and turned with the groups it is in.
        found = frames[0].draw(graphic, x, y, width, height, path, defs, warnings, frames[1])
        if found is not None:
            outer = _transform(turns)
            for primitive in found:
                if outer:
                    primitive.transform = f"{outer} {primitive.transform}"
            drawn.primitives.extend(found)
            return
    if graphic is None or graphic.kind in ("other", "chart", "diagram"):
        drawn.primitives.append(Primitive("placeholder", path, x=x, y=y, width=width, height=height,
                                          what="drawing"))
        what = (graphic.uri or (graphic.kind if graphic.kind != "other" else "graphic")) if graphic else "none"
        warnings.append(f"drawing-not-drawn:{what.rsplit('/', 1)[-1]}")
        return
    for name in graphic.unsupported:
        warnings.append(f"drawing-not-drawn:{name}")
    turn = _turn(graphic, x, y, width, height)
    here = turns + ((turn,) if turn else ())
    transform = _transform(here)
    if graphic.kind == "picture":
        drawn.primitives.append(Primitive("image", path, x=x, y=y, width=width, height=height,
                                          relationship=graphic.relationship, transform=transform))
        _outline(graphic, x, y, width, height, path, colours, defs, drawn, here)
        return
    if graphic.kind == "group":
        cw, ch = graphic.child_extent
        ox, oy = graphic.child_offset
        sx = width / cw if cw else Fraction(1)
        sy = height / ch if ch else Fraction(1)
        # A member's ``a:grpFill`` is this group's fill spread over this group's box (a
        # gradient runs across the whole group, measured); a group that states none
        # passes its own group's on.
        own = colours.fill(graphic.fill, group=group_fill)
        if own is not None:
            group_fill, group_box = own, (x, y, width, height)
        begin = len(drawn.primitives)
        for member in graphic.children:
            mx, my = member.offset
            mw, mh = member.extent
            _draw(member, x + (mx - ox) * sx, y + (my - oy) * sy, mw * sx, mh * sy, member.path or path,
                  colours, defs, drawn, group_fill, group_box, here, frames)
        _wrap(drawn, begin, colours.effects(graphic.effects), defs, path)
        return
    # A shape: its fill, then its outline, in a group carrying its effects.
    style = graphic.style
    if graphic.fill is not None:
        fill = colours.fill(graphic.fill, group=group_fill)
    else:
        fill = colours.style_fill(style.fill_ref if style is not None else None)
    box = group_box if isinstance(graphic.fill, s.SourceGroupFill) and group_box is not None else (x, y, width, height)
    paths = geometry_paths(graphic.geometry, x, y, width, height, warnings)
    begin = len(drawn.primitives)
    if fill is not None and not isinstance(fill, m.NoFill):
        attrs = _fill_attrs(fill, defs, box, here, graphic.geometry)
        solid = fill.color.hex[1:] if isinstance(fill, m.SolidFill) else None
        for geometry_path in paths:
            if geometry_path.fill != "none" and geometry_path.d:
                if geometry_path.fill != "norm":
                    warnings.append(f"drawing-fill-approximate:{geometry_path.fill}")
                drawn.primitives.append(Primitive("path", path, d=geometry_path.d, fill=solid, attrs=attrs,
                                                  transform=transform))
    _outline(graphic, x, y, width, height, path, colours, defs, drawn, here, paths)
    _wrap(drawn, begin, colours.effects(graphic.effects, style.effect_ref if style is not None else None), defs,
          path)
    if graphic.text:
        drawn.text_boxes.append((graphic, x, y, width, height))


def _fill_attrs(fill, defs, box, turns: tuple, geometry) -> str:
    name = geometry[1] if geometry is not None and geometry[0] == "preset" else "rect" if geometry is None else None
    frame = dataclasses.replace(_frame(turns), rectangular=name in _RECTANGULAR)
    return dml_fill.render_fill_attrs(fill, defs, tuple(float(v) for v in box), rules=WORD, dpi=DPI, frame=frame)


def _wrap(drawn: Drawn, begin: int, effects, defs, path: str) -> None:
    """Put the primitives from ``begin`` on in a group that carries ``effects``."""
    if effects is None or begin == len(drawn.primitives):
        return
    attrs = dml_effect.render_effects(effects, defs, dpi=DPI)
    if attrs:
        drawn.primitives.insert(begin, Primitive("begin", path, attrs=attrs))
        drawn.primitives.append(Primitive("end", path))


def geometry_paths(geometry: tuple | None, x, y, width, height, warnings: list[str]) -> list:
    """The geometry's paths over the box (``ooxml_common.drawingml.geometry.GeometryPath``:
    path data, fill mode, stroked): a preset from ECMA-376's complete table, its
    adjustments overridden by the shape's ``a:avLst``, or a custom geometry, its guides
    evaluated at the box's size in px."""
    w, h, ox, oy = float(width), float(height), float(x), float(y)
    if geometry is None:
        geometry = ("preset", "rect", {})
    if geometry[0] == "preset":
        found = dml_geometry.preset_path_data(geometry[1], w, h, geometry[2], x=ox, y=oy)
        if found is None:
            warnings.append(f"drawing-geometry-unknown:{geometry[1]}")
            return []
        return found
    return dml_geometry.spec_path_data(geometry[1], w, h, x=ox, y=oy)


def _outline(graphic: Graphic, x, y, width, height, path: str, colours: Colours, defs, drawn: Drawn, turns: tuple,
             paths=None) -> None:
    style = graphic.style
    outline = colours.outline(graphic.line, style.line_ref if style is not None else None)
    if outline is None:
        return
    if paths is None:
        paths = geometry_paths(graphic.geometry, x, y, width, height, drawn.warnings)
    box = (float(x), float(y), float(width), float(height))
    attrs = dml_fill.render_outline_attrs(outline, defs, rules=WORD, dpi=DPI, box=box, frame=_frame(turns))
    solid = outline.fill.color.hex[1:] if isinstance(outline.fill, m.SolidFill) else None
    transform = _transform(turns)
    for geometry_path in paths:
        if not (geometry_path.stroke and geometry_path.d):
            continue
        d, heads = geometry_path.d, []
        if outline.head_end is not None or outline.tail_end is not None:
            ends = dml_geometry.path_ends(d)
            if ends is not None:
                heads, start, end = dml_fill.render_arrowheads(outline, ends, dpi=DPI)
                d = dml_geometry.trim_path(d, start, end)
        drawn.primitives.append(Primitive("path", path, d=d, stroke=solid, stroke_width=outline.width * EMU_PX,
                                          attrs='fill="none" ' + attrs, transform=transform))
        drawn.primitives.extend(Primitive("markup", path, markup=markup, transform=transform) for markup in heads)


#: Automatic text on a fill darker than this is drawn white (the fill's Rec. 601 luma, 0-255).
#: Measured on ``make_dml_probe.py``'s ``auto`` family: greys ``48`` (72) and darker and
#: ``E00000`` (67) turn it white; grey ``4C`` (76), ``008000`` (75.1) and ``4472C4`` (110)
#: leave it black.  The threshold lies between 72 and 75.1.
AUTOMATIC_WHITE_BELOW = 74


def automatic_text_colour(graphic: Graphic, colours: Colours) -> str | None:
    """What a shape's text whose colour is automatic is drawn in, ``RRGGBB``, or ``None``
    for the default black: its ``wps:style`` ``a:fontRef`` colour where it names one
    (whatever the fill: ``tx1`` on a dark fill stays black), else white on a solid fill
    darker than :data:`AUTOMATIC_WHITE_BELOW` (measured, ``make_dml_probe.py``)."""
    style = graphic.style
    if style is not None and style.font_ref is not None and style.font_ref.color is not None:
        resolved = colours.resolve(style.font_ref.color)
        return resolved.hex[1:] if resolved is not None else None
    fill = colours.fill(graphic.fill) if graphic.fill is not None else colours.style_fill(
        style.fill_ref if style is not None else None)
    if isinstance(fill, m.SolidFill):
        red, green, blue = (int(fill.color.hex[k:k + 2], 16) for k in (1, 3, 5))
        if 0.299 * red + 0.587 * green + 0.114 * blue < AUTOMATIC_WHITE_BELOW:
            return "FFFFFF"
    return None


def outline_width(graphic: Graphic, colours: Colours) -> int:
    """The width (EMU) a shape's outline states -- its own ``a:ln@w`` over its style's line
    reference's -- whether or not the outline is drawn; 0 where neither states one."""
    width = None
    reference = graphic.style.line_ref if graphic.style is not None else None
    formats = colours.formats
    if reference is not None and 0 < reference.idx and formats is not None \
            and reference.idx - 1 < len(formats.line_styles):
        width = formats.line_styles[reference.idx - 1].width
    if graphic.line is not None and graphic.line.width is not None:
        width = graphic.line.width
    return width or 0


def text_area_edges(graphic: Graphic, colours: Colours, width: Fraction, height: Fraction) -> tuple:
    """How far in from the left, top, right and bottom of a text box's shape (``width`` by
    ``height`` EMU) its text area starts, before ``wps:bodyPr``'s insets, in EMU.

    Measured by ``make_text_box_probe.py`` (both settings): the text is laid out in the
    geometry's **text rectangle** -- a ``roundRect``'s keeps clear of its corners, an
    ``ellipse``'s is the inscribed rectangle, a custom geometry's is its ``a:rect``
    (``ooxml_common.drawingml.geometry.text_rect``) -- **less half the outline's width** on
    every side, the width the shape or its style's line reference states, drawn or
    ``a:noFill`` (a 1 pt style line under ``a:noFill`` moves the text 2 px in and down at
    300 dpi, as a drawn one does).  The half is whole EMU, truncated: a 635 EMU outline
    under an inset of 72 twips puts the text 72 twips in, not 73
    (``make_text_box_top_probe.py``, ``edge``; the layout rounds each side to twips).
    """
    w, h = float(width), float(height)
    left, top, right, bottom = dml_geometry.text_rect(graphic.geometry, w, h, rect=graphic.text_rect)
    half = Fraction(round(outline_width(graphic, colours)) // 2)
    return (Fraction(left) + half, Fraction(top) + half, Fraction(w - right) + half, Fraction(h - bottom) + half)
