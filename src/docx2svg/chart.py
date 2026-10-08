"""Charts (``c:chart``) and SmartArt (``dgm:relIds``) in a ``w:drawing``, drawn by
``ooxml-common``'s chart layout and DrawingML scene renderers.

Neither is drawn from the document's own XML.  A chart part is data and styling, laid out
by :class:`ooxml_common.chart.layout.ChartBuilder` into the drawable scene -- the code
pptx2svg draws a slide's charts with -- under Word's measured rules
(:data:`ooxml_common.chart.rules.WORD`); a SmartArt diagram is the laid-out drawing Word
cached for it (``dsp:drawing``), found by :func:`ooxml_common.drawingml.diagram.
diagram_drawing_part` and read as a shape tree.  Both are then drawn by
:func:`ooxml_common.drawingml.elements.render_element` into the drawing's box.

What is Word's and not the shared code's is here: the chart part and the diagram's parts
are related from the part the drawing sits in (the document, a header, a footer); colours
resolve through the document's theme and ``w:clrSchemeMapping`` under Word's colour rules
(:class:`docx2svg.drawing.Colours`); a theme typeface (``+mn-lt``) is the document theme's;
and the faces and sizes Word gives a title that states none, and the line box it sets
DrawingML text in, were measured (``tools/make_chart_probe.py``,
``tools/make_smartart_probe.py``; ROADMAP.md, F.20 and F.21).

The scene renderers draw at 96 px an inch; the fragment is scaled onto the page's device
grid (300 an inch) by its transform.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from fractions import Fraction

from ooxml_common.chart import layout as chart_layout
from ooxml_common.chart import read as chart_read
from ooxml_common.chart import rules as chart_rules
from ooxml_common.drawingml import scene as m
from ooxml_common.drawingml import source as s
from ooxml_common.drawingml.context import RenderContext
from ooxml_common.drawingml.elements import render_element
from ooxml_common.drawingml.source_tree import SourceRunProperties
from ooxml_common.opc import OpcPackage
from ooxml_common.text.fontmap import metrics_for
from ooxml_common.text.measure import DefaultTextMeasurer

from .drawing import EMU_PX, Colours, Primitive

#: The renderers' grid (``ooxml_common.units.DEFAULT_DPI``) onto the page's.
_SCALE = Fraction(300, 96)


@dataclass
class Frames:
    """What a chart or a diagram needs from the document: its package (the ``.docx``'s
    bytes, opened once), the parsed document (theme, colour map, fonts) and the part a
    drawing sits in.  :meth:`draw` is what :func:`docx2svg.drawing.draw` calls for a
    ``chart`` or ``diagram`` graphic."""

    package: bytes | None
    document: object
    _opened: OpcPackage | None = field(default=None, repr=False)

    def opc(self) -> OpcPackage | None:
        if self._opened is None and self.package is not None:
            self._opened = OpcPackage.open(self.package)
        return self._opened

    def owner(self, part: str | None, path: str = "") -> str | None:
        """The part a drawing sits in: ``part`` (a header's, a footer's), else the notes
        part for a note's drawing (``path`` ``w:footnote[...]/...``, as
        :func:`docx2svg.paths.path_part` reads it), else the main document part."""
        if part:
            return part
        package = self.opc()
        if package is None:
            return None
        main = next((rel.target_part for rel in package.relationships("").values()
                     if rel.type.endswith("/officeDocument") and rel.target_part), None)
        first = path.split("/", 1)[0]
        for prefix in ("footnote", "endnote"):
            if main is not None and first.startswith(f"w:{prefix}["):
                return next((rel.target_part for rel in package.relationships(main).values()
                             if rel.type.endswith(f"/{prefix}s") and rel.target_part), None)
        return main

    def draw(self, graphic, x, y, width, height, path: str, defs, warnings: list[str],
             part: str | None = None) -> list[Primitive] | None:
        """The primitives ``graphic`` (a chart or a diagram) paints in the box ``x, y,
        width, height`` (device px), or ``None`` where it cannot be drawn (warned)."""
        package = self.opc()
        owner = self.owner(part, path)
        if package is None or owner is None:
            warnings.append(f"drawing-not-drawn:{graphic.kind}")
            return None
        width_emu, height_emu = width / EMU_PX, height / EMU_PX
        if graphic.kind == "chart":
            element = chart_element(package, owner, graphic.relationship, float(width_emu), float(height_emu),
                                    self.document, warnings)
        else:
            from .diagram import diagram_element

            element = diagram_element(package, owner, graphic.relationship, float(width_emu), float(height_emu),
                                      self.document, warnings)
        if element is None:
            return None
        return [render(element, x, y, path, defs, graphic.kind)]


def render(element, x, y, path: str, defs, what: str = "chart") -> Primitive:
    """``element`` (whose box starts at the origin) drawn at ``x, y`` (device px), its
    definitions -- gradients, markers, filters -- registered on the page's ``defs``
    under ids that page has not used."""
    rules = chart_rules.WORD.drawing
    context = RenderContext(rules=rules, font_mapping={}, measurer=WordTextMeasurer(kerning=rules.kerning))
    context._next_id = getattr(defs, "_next_id", 0)
    markup = render_element(element, context)
    if hasattr(defs, "_next_id"):
        defs._next_id = context._next_id
    for definition in context.defs:
        defs.add_def(definition)
    from .svg import number

    return Primitive("markup", path, markup=markup, transform=f"translate({number(x)} {number(y)}) "
                                                               f"scale({number(_SCALE)})", what=what)


class WordTextMeasurer(DefaultTextMeasurer):
    """``ooxml-common``'s measurer with Word's line box: a face's own ascent and descent
    (``hhea``, from the shared tables), where PowerPoint's is 1.2 em whatever the face.
    Measured on ``make_smartart_probe.py``: three lines of 19 pt Aptos at 90% spacing step
    20.88 pt, which is 0.9 of Aptos's 1.2207 em, where 1.2 em gives 20.52.  A face with no
    table keeps the shared measurer's answers.

    Built with Word's kerning (``DrawingRules.kerning``): the legacy ``kern`` table, which
    is what Word charges where it kerns (``tools/make_wrap_kern_probe.py``) -- not the
    OpenType feature the shared measurer charges by default."""

    def line_height_ratio(self, font_family=None, font_family_ea=None) -> float:
        metrics = metrics_for(font_family) or metrics_for(font_family_ea)
        if metrics is None:
            return super().line_height_ratio(font_family, font_family_ea)
        return (metrics.ascender + abs(metrics.descender)) / metrics.units_per_em

    def ascender_ratio(self, font_family=None, font_family_ea=None) -> float:
        metrics = metrics_for(font_family) or metrics_for(font_family_ea)
        if metrics is None:
            return super().ascender_ratio(font_family, font_family_ea)
        return metrics.ascender / metrics.units_per_em


# -- charts -----------------------------------------------------------------------------

CHART_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"


def chart_element(package: OpcPackage, owner: str, relationship: str | None, width_emu: float, height_emu: float,
                  document, warnings: list[str]):
    """The chart part ``relationship`` names from ``owner``, laid out in a frame of
    ``width_emu`` x ``height_emu`` as the scene's :class:`~ooxml_common.drawingml.scene.
    ChartElement`, or ``None`` (warned)."""
    part = package.related_part(owner, relationship) if relationship else None
    if part is None or not package.has_part(part):
        warnings.append("chart-unreadable:no chart part")
        return None
    try:
        xml = package.read_xml(part)
    except Exception:  # noqa: BLE001 -- a malformed part is reported, not raised
        warnings.append("chart-unreadable:not well-formed")
        return None
    source = chart_read.parse_chart_space(xml) if xml is not None else None
    if source is None:
        warnings.append("chart-unreadable:no c:chart")
        return None
    plots = chart_layout.drawable_plots(source)
    if not plots:
        kinds = ",".join(sorted({plot.kind for plot in source.plots})) or "nothing"
        warnings.append(f"chart-unsupported-type:{kinds}")
        return None
    if width_emu <= 0 or height_emu <= 0:
        warnings.append("chart-unreadable:zero-sized")
        return None
    colours = Colours(document.theme_colors, document.color_map, document.theme_formats, warnings)
    if source.color_map_override is not None:
        # ``c:clrMapOvr`` replaces the document's mapping for this chart, slot by slot.
        colours.slots = dict(colours.slots, **source.color_map_override.mapping)
    style = chart_style(source, document, colours)
    if has_chart_style(xml):
        # Word 365's chart style draws every default axis line, tick mark and gridline in
        # ``898989`` (measured, ``make_chart_text_probe.py``), where a chart without one
        # draws them black.
        style = replace(style, line_color=CHART_STYLE_LINE)
    # Which title each text body is: the chart's (or its automatic one, which has none),
    # or an axis' -- they take different defaults -- and its paragraphs' ``a:defRPr``s.
    titles = {id(axis.title.rich): ("axis", _title_defaults(node))
              for axis, node in zip(source.axes, _axis_nodes(xml)) if axis.title is not None}
    chart_title = ("chart", _title_defaults(_child_path(xml, "chart", "title")))
    base = chart_title_size(source, xml)

    def resolve_text(rich, text, size, align):
        kind, stated = titles.get(id(rich), chart_title) if rich is not None else ("auto", [])
        return title_text(document, colours, style, rich, text, size, align, stated=stated,
                          base_size=style.font_size if kind == "axis" else base)

    builder = chart_layout.ChartBuilder(
        source, plots[0], width_pt=width_emu / chart_layout.EMU_PER_POINT,
        height_pt=height_emu / chart_layout.EMU_PER_POINT, style=style, resolve_fill=colours.fill,
        resolve_outline=lambda outline: _outline(colours, outline), resolve_text=resolve_text,
        resolve_typeface=lambda typeface: theme_typeface(document, typeface), plots=plots, rules=chart_rules.WORD)
    children, data = builder.build()
    if builder.flattened_three_d:
        warnings.append("chart-3d-flattened:" + ",".join(builder.flattened_three_d))
    if builder.default_title_wanted:
        # Word's own "Chart Title", in its interface's language, which a document does
        # not say; its band is kept.
        warnings.append("chart-title-not-drawn:title in Word's own words")
    for reason in dict.fromkeys(builder.axis_titles_not_drawn):
        warnings.append("chart-axis-title-not-drawn:" + AXIS_TITLE_REASONS.get(reason, reason))
    frame = m.Transform(offset_x=0, offset_y=0, extent_width=width_emu, extent_height=height_emu)
    return m.ChartElement(transform=frame, chart=data, child_transform=replace(frame), children=children)


def chart_style(source, document, colours: Colours) -> chart_layout.ChartStyle:
    """The chart's text and series colour defaults in Word: the document theme's minor
    face (measured: every axis and legend label of a chart that states no face), the
    size the chart's ``c:txPr`` states or 10 pt, ``tx1``, and the theme's accents."""
    scheme = getattr(document, "font_scheme", None)
    minor = scheme.minor.latin or None if scheme is not None else None
    minor_ea = scheme.minor.east_asian or None if scheme is not None else None
    text = colours.resolve(s.SchemeColor(scheme="tx1")) or m.ResolvedColor("#000000")
    accents = chart_layout.accent_colors(lambda key: colours.resolve(s.SchemeColor(scheme=key)))
    return chart_layout.ChartStyle(font_family=minor, font_size=chart_layout.default_font_size(source), color=text,
                                   accents=accents, font_family_ea=minor_ea)


def theme_typeface(document, typeface: str | None) -> str | None:
    """``+mn-lt`` / ``+mj-lt`` (and ``-ea``, ``-cs``) as the document theme's face; any
    other typeface as it is."""
    if not typeface or not typeface.startswith("+"):
        return typeface
    scheme = getattr(document, "font_scheme", None)
    if scheme is None:
        return None
    fonts = scheme.major if typeface.startswith("+mj") else scheme.minor
    return {"lt": fonts.latin, "ea": fonts.east_asian, "cs": fonts.complex_script}.get(typeface[-2:]) or None


def _outline(colours: Colours, outline):
    """A chart element's ``a:ln``: ``None`` for none (no line, or ``a:noFill``)."""
    if outline is None:
        return None
    paint = colours.fill(outline.fill) if outline.fill is not None else None
    if isinstance(paint, m.NoFill):
        return None
    return m.Outline(width=outline.width if outline.width is not None else DEFAULT_CHART_LINE_EMU,
                     fill=paint if isinstance(paint, (m.SolidFill, m.GradientFill)) else None,
                     dash_style=outline.dash_style or "solid",
                     custom_dash=list(outline.custom_dash) if outline.custom_dash else None,
                     line_cap=outline.line_cap, line_join=outline.line_join, head_end=outline.head_end,
                     tail_end=outline.tail_end, compound=outline.compound)


#: ``a:ln`` without ``@w`` in a chart part.
DEFAULT_CHART_LINE_EMU = 9525


def _child_path(node, *names):
    from ooxml_common.xmlutil import child

    for name in names:
        node = child(node, name)
    return node


def _axis_nodes(chart_space) -> list:
    """The plot area's axes, in the order :func:`ooxml_common.chart.read.parse_chart_space`
    reads them."""
    from ooxml_common.chart.read import AXIS_ELEMENTS
    from ooxml_common.xmlutil import children, local_name

    return [node for node in children(_child_path(chart_space, "chart", "plotArea"))
            if local_name(node.tag) in AXIS_ELEMENTS]


def _title_defaults(owner) -> list[bool]:
    """Per paragraph of the title ``c:title`` under ``owner`` (an axis; or the chart's,
    given the ``c:title`` itself), whether it states an ``a:defRPr`` at all -- an empty
    one counts, which is why the XML is asked and not the reader's model."""
    from ooxml_common.xmlutil import child, children, local_name

    title = owner if owner is not None and local_name(owner.tag) == "title" else child(owner, "title")
    rich = child(child(title, "tx"), "rich")
    return [child(child(paragraph, "pPr"), "defRPr") is not None for paragraph in children(rich, "p")]


def _stated_text_size(body) -> float | None:
    """``c:txPr``'s first ``a:defRPr@sz`` (pt), or ``None`` where it states none."""
    for paragraph in (body.paragraphs if body is not None else []):
        defaults = paragraph.properties.default_run_properties if paragraph.properties else None
        if defaults is not None:
            return defaults.font_size or None
    return None


#: ``c14:style``'s namespace: the chart style Word 365 states, which sets a title's default
#: size (:func:`chart_title_size`).
C14 = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart"


def chart_title_size(source, chart_space) -> float:
    """The size Word gives a chart title that takes the chart's text -- one whose
    paragraph states an empty ``a:defRPr``, or Word's automatic title (measured,
    ``make_chart_text_probe.py``): **1.2 times** the size the chart's ``c:txPr`` states
    (8 pt gives 9.6, 10 gives 12, 12 gives 14.4, 18 gives 21.6, with the chart style or
    without); where it states none, **18 pt** in a chart that states a ``c14:style``
    (Word 365's, in ``mc:AlternateContent``; ``c:style`` alone does not count) and 10 pt in
    one that does not."""
    stated = _stated_text_size(source.text_properties)
    if stated is not None:
        return 1.2 * stated
    return 18.0 if has_chart_style(chart_space) else 10.0


def has_chart_style(chart_space) -> bool:
    """Whether the chart states Word 365's chart style, ``c14:style`` (which Word writes in
    ``mc:AlternateContent``, a ``c:style`` its fallback); ``c:style`` alone is not it."""
    return any(node.tag == f"{{{C14}}}style" for node in chart_space.iter())


#: What an axis title not drawn is, by :attr:`ChartBuilder.axis_titles_not_drawn`'s reason.
AXIS_TITLE_REASONS = {"default": "axis title in Word's own words", "chart": "axis title in a chart of this kind",
                      "placement": "axis title turned or placed otherwise"}


#: The default axis, tick and gridline colour under Word 365's chart style.
CHART_STYLE_LINE = "#898989"


def title_text(document, colours: Colours, style, rich, text: str, size: float, align: str,
               stated: list[bool] | None = None, base_size: float | None = None) -> m.TextBody:
    """A chart's or an axis' title text in Word: each run's own face, size and weight over
    its paragraph's ``a:defRPr``, over Word's defaults for a title -- which depend on
    whether the paragraph states an ``a:defRPr`` at all (measured, ``make_chart_probe.py``,
    ``make_chart_text_probe.py``):

    * with one, even an empty one, the chart's minor face at ``base_size`` -- a chart
      title's :func:`chart_title_size`, an axis title's the chart's text size (``c:txPr``'s
      or 10 pt) -- **bold**; and so is Word's automatic title (``rich`` ``None``: the series'
      name);
    * without one, :data:`TITLE_FACE` at ``size`` (18 pt), not bold -- whatever face and
      size the chart's ``c:txPr`` states.
    """
    paragraphs = []
    source_paragraphs = (rich.paragraphs if rich is not None else []) or [None]
    for index, paragraph in enumerate(source_paragraphs):
        defaults = paragraph.properties.default_run_properties if paragraph is not None and paragraph.properties \
            else None
        if rich is None or defaults is not None or (stated is not None and index < len(stated) and stated[index]):
            base = SourceRunProperties(typeface=style.font_family,
                                       font_size=style.font_size if base_size is None else base_size, bold=True)
        else:
            base = SourceRunProperties(typeface=TITLE_FACE, font_size=size, bold=False)
        runs = (paragraph.runs if paragraph is not None else []) or [None]
        out = []
        for source_run in runs:
            props = _merge_run(base, defaults, source_run.properties if source_run is not None else None)
            out.append(m.TextRun(text=source_run.text if source_run is not None else text,
                                 properties=_run_properties(document, colours, props, style)))
        paragraphs.append(m.Paragraph(runs=out, properties=m.ParagraphProperties(alignment=align)))
    return m.TextBody(paragraphs=paragraphs, body_properties=chart_layout.CHART_TEXT_BODY)


def _merge_run(*layers):
    out = SourceRunProperties()
    for layer in layers:
        if layer is None:
            continue
        for name in out.__dataclass_fields__:
            value = getattr(layer, name)
            if value is not None:
                setattr(out, name, value)
    return out


def _run_properties(document, colours: Colours, props, style) -> m.RunProperties:
    colour = colours.resolve(props.color) if props.color is not None else None
    return m.RunProperties(font_size=props.font_size, font_family=theme_typeface(document, props.typeface),
                           font_family_ea=theme_typeface(document, props.typeface_ea) if props.typeface_ea
                           else style.font_family_ea,
                           bold=bool(props.bold), italic=bool(props.italic), color=colour or style.color)


#: The face of a chart title whose paragraph states no ``a:defRPr``, whatever the theme,
#: the chart's ``c:txPr`` and the document's own faces (measured: ``make_chart_probe.py``,
#: and documents whose default faces were Arial, Times New Roman and Verdana under themes
#: of Georgia and Aptos).
TITLE_FACE = "Arial"
