"""SmartArt (``dgm:relIds`` in a ``w:drawing``): the laid-out drawing Word cached for it,
as ``ooxml-common``'s scene.

A diagram is authored as data and a layout definition, and laying it out is a diagram
engine of its own.  Word lays it out whenever it opens the document and writes the result
on every save as a drawing part (``dsp:drawing``): a shape tree in DrawingML, found from
the data model by :func:`ooxml_common.drawingml.diagram.diagram_drawing_part` and read by
:func:`ooxml_common.drawingml.read_tree.parse_shape_tree`.  Drawing a document's SmartArt
is drawing that cache.  What is Word's here is how a shape of it resolves: its colours
through the document's theme and ``w:clrSchemeMapping`` (:class:`docx2svg.drawing.
Colours`), its theme style references (``dsp:style``) through the theme's format scheme,
and its text -- the face the style's ``a:fontRef`` names (the theme's minor or major
face), in the reference's colour, at the size the cache states.
"""

from __future__ import annotations

from ooxml_common.drawingml import diagram as dml_diagram
from ooxml_common.drawingml import scene as m
from ooxml_common.drawingml import source as s
from ooxml_common.drawingml import source_tree as st
from ooxml_common.drawingml.read_tree import parse_shape_tree
from ooxml_common.xmlutil import attr, child, local_name

from .drawing import Colours

#: DrawingML's run size where neither the run nor its paragraph states one, pt.
DEFAULT_TEXT_PT = 18.0


def diagram_element(package, owner: str, relationship: str | None, width_emu: float, height_emu: float,
                    document, warnings: list[str]):
    """The drawing Word cached for the SmartArt whose data model ``relationship`` names
    from ``owner``, as a scene group filling a frame of ``width_emu`` x ``height_emu``, or
    ``None`` (warned)."""
    data_part = package.related_part(owner, relationship) if relationship else None
    if data_part is None or not package.has_part(data_part):
        warnings.append("diagram-unreadable:no data model")
        return None
    drawing_part = dml_diagram.diagram_drawing_part(package, owner, data_part)
    if drawing_part is None:
        warnings.append("diagram-no-cached-drawing")
        return None
    try:
        drawing = package.read_xml(drawing_part)
    except Exception:  # noqa: BLE001 -- a malformed part is reported, not raised
        drawing = None
    sp_tree = child(drawing, "spTree") if drawing is not None else None
    if sp_tree is None:
        warnings.append("diagram-unreadable:no shape tree")
        return None
    frame = m.Transform(offset_x=0, offset_y=0, extent_width=width_emu, extent_height=height_emu)
    colours = Colours(document.theme_colors, document.color_map, document.theme_formats, warnings)
    fonts = iter([attr(child(child(node, "style"), "fontRef"), "idx") for node in sp_tree.iter()
                  if local_name(node.tag) == "sp"])
    children = [found for node in parse_shape_tree(sp_tree)
                if (found := _element(node, colours, document, fonts, warnings)) is not None]
    if not children:
        warnings.append("diagram-no-cached-drawing")
        return None
    return m.GroupElement(transform=frame, child_transform=dml_diagram.diagram_child_transform(sp_tree, frame),
                          children=children)


def _transform(transform: s.SourceTransform | None) -> m.Transform:
    if transform is None:
        return m.Transform()
    return m.Transform(offset_x=transform.offset_x, offset_y=transform.offset_y, extent_width=transform.width,
                       extent_height=transform.height, rotation=transform.rotation / 60000,
                       flip_h=transform.flip_horizontal, flip_v=transform.flip_vertical)


def _geometry(geometry) -> m.Geometry:
    if geometry is None:
        return m.PresetGeometry(preset="rect")
    if isinstance(geometry, s.SourceCustomGeometry):
        return m.CustomGeometry(paths=list(geometry.paths))
    return m.PresetGeometry(preset=geometry.preset, adjust_values=dict(geometry.adjust_values))


def _element(node, colours: Colours, document, fonts, warnings: list[str]):
    if getattr(node, "hidden", False):
        if isinstance(node, st.SourceShape):
            next(fonts, None)
        return None
    if isinstance(node, st.SourceGroup):
        children = [found for item in node.children
                    if (found := _element(item, colours, document, fonts, warnings)) is not None]
        transform = _transform(node.transform)
        inner = node.child_transform
        child_transform = m.Transform(offset_x=inner.offset_x, offset_y=inner.offset_y, extent_width=inner.width,
                                      extent_height=inner.height) if inner is not None else transform
        return m.GroupElement(transform=transform, child_transform=child_transform, children=children,
                              effects=colours.effects(node.effects))
    if isinstance(node, st.SourceConnector):
        style = node.style
        return m.ConnectorElement(transform=_transform(node.transform), geometry=_geometry(node.geometry),
                                  outline=colours.outline(node.outline, style.line_ref if style else None),
                                  effects=colours.effects(node.effects, style.effect_ref if style else None))
    if isinstance(node, st.SourceShape):
        font_ref = next(fonts, None)
        style = node.style
        fill = colours.fill(node.fill) if node.fill is not None else colours.style_fill(
            style.fill_ref if style else None)
        return m.ShapeElement(
            transform=_transform(node.transform), geometry=_geometry(node.geometry), fill=fill,
            outline=colours.outline(node.outline, style.line_ref if style else None),
            effects=colours.effects(node.effects, style.effect_ref if style else None),
            text_body=_text_body(node.text_body, style, font_ref, colours, document),
            text_transform=_transform(node.text_transform) if node.text_transform is not None else None,
            text_rect=node.text_rect)
    warnings.append(f"diagram-not-drawn:{getattr(node, 'kind', 'element')}")
    return None


def _text_body(body: st.SourceTextBody | None, style, font_ref: str | None, colours: Colours, document):
    if body is None or not any(run.text for paragraph in body.paragraphs for run in paragraph.runs):
        return None
    scheme = getattr(document, "font_scheme", None)
    theme = (scheme.major if font_ref == "major" else scheme.minor) if scheme is not None else None
    face = theme.latin or None if theme is not None else None
    face_ea = theme.east_asian or None if theme is not None else None
    reference = style.font_ref if style is not None else None
    colour = colours.resolve(reference.color) if reference is not None and reference.color is not None else None
    paragraphs = []
    for paragraph in body.paragraphs:
        props = paragraph.properties or st.SourceParagraphProperties()
        defaults = props.default_run_properties
        runs = []
        for run in paragraph.runs:
            merged = _merge(defaults, run.properties)
            runs.append(m.TextRun(text=run.text, properties=_run(merged, face, face_ea, colour, colours, document)))
        end = paragraph.end_para_run_properties
        paragraphs.append(m.Paragraph(
            runs=runs,
            properties=m.ParagraphProperties(
                alignment=props.align, line_spacing=props.line_spacing,
                space_before=props.space_before or m.PercentSpacing(0),
                space_after=props.space_after or m.PercentSpacing(0), level=props.level or 0,
                bullet=props.bullet if not isinstance(props.bullet, st.SourceBlipBullet) else None,
                bullet_font=_face(props.bullet_font, document),
                bullet_color=colours.resolve(props.bullet_color) if props.bullet_color is not None else None,
                bullet_size_pct=props.bullet_size_pct, bullet_size_points=props.bullet_size_points,
                margin_left=props.margin_left, indent=props.indent, tab_stops=list(props.tab_stops or [])),
            end_para_run_properties=_run(_merge(defaults, end), face, face_ea, colour, colours, document)
            if end is not None else None))
    return m.TextBody(paragraphs=paragraphs, body_properties=_body(body.properties))


def _merge(*layers) -> st.SourceRunProperties:
    out = st.SourceRunProperties()
    for layer in layers:
        if layer is None:
            continue
        for name in out.__dataclass_fields__:
            value = getattr(layer, name)
            if value is not None:
                setattr(out, name, value)
    return out


def _face(typeface: str | None, document) -> str | None:
    from .chart import theme_typeface

    return theme_typeface(document, typeface)


def _run(props: st.SourceRunProperties, face, face_ea, colour, colours: Colours, document) -> m.RunProperties:
    resolved = colours.resolve(props.color) if props.color is not None else colour
    return m.RunProperties(
        font_size=props.font_size or DEFAULT_TEXT_PT, font_family=_face(props.typeface, document) or face,
        font_family_ea=_face(props.typeface_ea, document) or face_ea,
        font_family_cs=_face(props.typeface_cs, document), bold=bool(props.bold), italic=bool(props.italic),
        underline=bool(props.underline), underline_style=props.underline_style,
        strikethrough=bool(props.strikethrough), baseline=props.baseline or 0.0,
        color=resolved or colours.resolve(s.SchemeColor(scheme="tx1")))


def _body(props: st.SourceTextBodyProperties | None) -> m.BodyProperties:
    out = m.BodyProperties()
    if props is None:
        return out
    for name, value in (("margin_left", props.margin_left), ("margin_right", props.margin_right),
                        ("margin_top", props.margin_top), ("margin_bottom", props.margin_bottom),
                        ("anchor", props.anchor), ("wrap", props.wrap), ("auto_fit", props.auto_fit),
                        ("font_scale", props.font_scale), ("ln_spc_reduction", props.ln_spc_reduction),
                        ("num_col", props.num_col), ("vert", props.vert),
                        ("rotation", props.rotation / 60000 if props.rotation is not None else None),
                        ("default_tab_size", props.default_tab_size)):
        if value is not None:
            setattr(out, name, value)
    return out
