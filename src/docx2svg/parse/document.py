"""Read ``word/document.xml`` and ``word/styles.xml`` into :mod:`docx2svg.model`.

Scope, and why it is this small
-------------------------------

This reader covers what the fixtures exercise: sections, paragraph and run properties,
tabs, indents, spacing, and -- for the style cascade -- the styles part, the theme's font
scheme, numbering, ``w:themeFontLang``, and tables as far as their structure and style
(``Document.body``).  Each paragraph and run also carries every property it *declared*
as a flat map (:mod:`docx2svg.parse.properties`), which is what :mod:`docx2svg.resolve`
merges.  Headers and footers are read as their blocks; a field's instruction is never
text: a computed field (``PAGE``, ``NUMPAGES``, ``SECTIONPAGES``) is one run carrying it
(:class:`docx2svg.model.Field`), and any other field is its cached result's runs.
Footnotes are read for the paginator; ``w:drawing`` no further than recording that a run
holds one, its extent, and a floating one's anchor.

That is on purpose and it is the project's first methodological commitment: **one measured
number beats a large unrunnable skeleton.**  Every element this reader handles is one
that a test compares against Word's own PDF; nothing here is written on the strength of
having read the schema.  ROADMAP.md says which phase adds each of the rest and what must
be measured before it.

The one piece of real format subtlety already present is section discovery, which is not
where a first reading of the schema suggests it is -- see :func:`_sections`.
"""

from __future__ import annotations

import dataclasses
from dataclasses import dataclass, field
from typing import AbstractSet, BinaryIO, Callable
from xml.etree.ElementTree import Element

from ooxml_common.opc import OpcPackage
from ooxml_common.xmlutil import attr, local_name, parse_xml
from ooxml_common.xmlutil import element_child as child
from ooxml_common.xmlutil import element_children as children

from ..model import (
    Anchor,
    Document,
    Field,
    Graphic,
    Table,
    Indent,
    PageMargins,
    PageSize,
    Paragraph,
    ParagraphProperties,
    PictureBullet,
    Run,
    RunProperties,
    Section,
    Spacing,
    TabStop,
)
from ..opc import REL_ENDNOTES, REL_FOOTNOTES, REL_NUMBERING, REL_SETTINGS, REL_STYLES, REL_THEME, main_part, read_part
from ..paths import Steps, row_level
from ..xmlutil import attr_bool, attr_int
from .drawing import PictureRef, drawing_element, picture_crop, read_anchor, read_graphic_data
from .vml import read_vml
from .properties import read_paragraph_properties, read_run_properties, read_table_properties, text_colour, theme_colours
from .styles import (
    parse_compat_options, parse_compat_settings, parse_compatibility_mode, parse_default_tab_stop,
    parse_color_map, parse_endnote_settings, parse_even_and_odd_headers, parse_hyphenation, parse_numbering, parse_picture_bullets, parse_settings,
    parse_styles, parse_theme, parse_theme_colors, parse_theme_formats,
)

#: Elements inside a ``w:r`` that are content but are not text.  Recorded by index so a
#: later layout stage knows a tab happened *between* two specific characters; folding
#: them into the run's text would make a tab indistinguishable from a space, and a tab
#: is a jump to a stop rather than a width.
_RUN_MARKS = {"tab", "br", "cr", "noBreakHyphen", "softHyphen"}
#: Inline objects that occupy a line without being text.  Recorded as marks so that a
#: later stage knows the line is not only text -- a drawing's height moves every line
#: below it, and nothing about drawings has been measured yet.
_RUN_OBJECTS = {"drawing", "pict", "object", "AlternateContent"}


#: Body-level elements that mark a range or a spot and hold no content: nothing is lost by
#: not reading them, so they are not reported.
_MARKERS = frozenset({"bookmarkStart", "bookmarkEnd", "proofErr", "permStart", "permEnd",
                      "commentRangeStart", "commentRangeEnd", "moveFromRangeStart", "moveFromRangeEnd",
                      "moveToRangeStart", "moveToRangeEnd"})


def _tab_stops(properties: Element | None) -> tuple[TabStop, ...]:
    tabs = child(properties, "tabs")
    if tabs is None:
        return ()
    stops = []
    for element in children(tabs, "tab"):
        position = attr_int(element, "pos")
        if position is None:
            continue
        stops.append(
            TabStop(
                position_twips=position,
                alignment=attr(element, "val") or "left",
                leader=attr(element, "leader"),
            )
        )
    # Word sorts stops by position regardless of document order, and a document that
    # lists them out of order is legal.  Sorting here means no consumer has to.
    return tuple(sorted(stops, key=lambda stop: stop.position_twips))


def _indent(properties: Element | None) -> Indent | None:
    element = child(properties, "ind")
    if element is None:
        return None
    # ``w:start``/``w:end`` are the ISO-strict spellings of ``w:left``/``w:right`` and
    # both occur in the wild; Word 2010+ writes the transitional pair.  Reading only
    # one spelling loses the indent silently, which is the worst possible failure for
    # a quantity this project measures to a hundredth of a point.
    left = attr_int(element, "left")
    if left is None:
        left = attr_int(element, "start")
    right = attr_int(element, "right")
    if right is None:
        right = attr_int(element, "end")
    return Indent(
        left=left,
        right=right,
        first_line=attr_int(element, "firstLine"),
        hanging=attr_int(element, "hanging"),
    )


def _spacing(properties: Element | None) -> Spacing | None:
    element = child(properties, "spacing")
    if element is None:
        return None
    return Spacing(
        before=attr_int(element, "before"),
        after=attr_int(element, "after"),
        line=attr_int(element, "line"),
        line_rule=attr(element, "lineRule"),
    )


def _run_properties(properties: Element | None) -> RunProperties | None:
    if properties is None:
        return None
    fonts = child(properties, "rFonts")
    bold = child(properties, "b")
    italic = child(properties, "i")
    strike = child(properties, "strike")
    style = child(properties, "rStyle")
    color = child(properties, "color")
    size = child(properties, "sz")
    underline = child(properties, "u")
    spacing = child(properties, "spacing")
    vertical = child(properties, "vertAlign")
    return RunProperties(
        ascii_font=attr(fonts, "ascii"),
        high_ansi_font=attr(fonts, "hAnsi"),
        east_asian_font=attr(fonts, "eastAsia"),
        complex_script_font=attr(fonts, "cs"),
        size_half_points=attr_int(size, "val"),
        # ``attr_bool`` rather than a truth test on the element: an absent ``w:val``
        # means *true*, so ``<w:b/>`` is bold and ``<w:b w:val="0"/>`` is not.
        bold=attr_bool(bold) if bold is not None else None,
        italic=attr_bool(italic) if italic is not None else None,
        strike=attr_bool(strike) if strike is not None else None,
        underline=attr(underline, "val"),
        color=text_colour(color),
        character_spacing=attr_int(spacing, "val"),
        vertical_align=attr(vertical, "val"),
        style_id=attr(style, "val"),
        declared=read_run_properties(properties),
    )


def _paragraph_properties(properties: Element | None) -> ParagraphProperties | None:
    if properties is None:
        return None
    numbering = child(properties, "numPr")
    style = child(properties, "pStyle")
    alignment = child(properties, "jc")
    widow = child(properties, "widowControl")
    return ParagraphProperties(
        style_id=attr(style, "val"),
        alignment=attr(alignment, "val"),
        indent=_indent(properties),
        spacing=_spacing(properties),
        tab_stops=_tab_stops(properties),
        numbering_id=attr_int(child(numbering, "numId"), "val"),
        numbering_level=attr_int(child(numbering, "ilvl"), "val"),
        keep_next=child(properties, "keepNext") is not None
        and attr_bool(child(properties, "keepNext")),
        keep_lines=child(properties, "keepLines") is not None
        and attr_bool(child(properties, "keepLines")),
        page_break_before=child(properties, "pageBreakBefore") is not None
        and attr_bool(child(properties, "pageBreakBefore")),
        widow_control=attr_bool(widow) if widow is not None else None,
        run_properties=_run_properties(child(properties, "rPr")),
        declared=read_paragraph_properties(properties),
    )


def _object(node: Element) -> str:
    """The mark an inline object leaves in its run: ``drawing:cx:cy:l:t:r:b`` (EMU) for a
    ``w:drawing`` holding one ``wp:inline`` -- its extent and effect extent, which the
    line breaker and the line height use -- and ``drawing`` for anything else (an
    anchored, floating object, :attr:`Run.anchors`; VML; an embedded object).  A drawing
    inside ``mc:AlternateContent`` is its DrawingML choice (:func:`drawing.drawing_element`)."""
    drawing = drawing_element(node)
    if drawing is None:
        return "drawing"
    inline = [child_ for child_ in drawing if local_name(child_.tag) == "inline"]
    if len(inline) != 1:
        return "drawing"
    extent = child(inline[0], "extent")
    effect = child(inline[0], "effectExtent")
    if extent is None:
        return "drawing"
    values = [attr_int(extent, "cx", 0) or 0, attr_int(extent, "cy", 0) or 0]
    values += [attr_int(effect, side, 0) or 0 if effect is not None else 0 for side in ("l", "t", "r", "b")]
    return "drawing:" + ":".join(str(v) for v in values)


_W14_PARA_ID = "{http://schemas.microsoft.com/office/word/2010/wordml}paraId"
_R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"


def _anchor(node: Element, path: str, fields: "_Fields") -> Anchor | None:
    """A floating drawing (:attr:`Run.anchors`), its text boxes' content read as a part's
    blocks (with ``fields``, as the part's other fields)."""
    drawing = drawing_element(node)
    if drawing is None:
        return None
    return read_anchor(drawing, path, lambda element, where: _text_box_story(element, where, fields))


def _picture(node: Element) -> str | None:
    """The relationship id of the picture an inline ``w:drawing`` shows, or ``None``; with
    its crop where ``a:srcRect`` cuts it (:class:`~docx2svg.parse.drawing.PictureRef`)."""
    drawing = drawing_element(node)
    if drawing is None or child(drawing, "inline") is None:
        return None
    for element in drawing.iter():
        if local_name(element.tag) == "pic":
            blip = child(child(element, "blipFill"), "blip")
            embed = blip.get(_R_EMBED) if blip is not None else None
            if embed:
                crop = picture_crop(element)
                return PictureRef(embed, crop) if crop else embed
    for element in drawing.iter():
        if local_name(element.tag) == "blip" and element.get(_R_EMBED):
            return element.get(_R_EMBED)
    return None


def _inline_graphic(node: Element, path: str, fields: "_Fields"):
    """What an inline ``w:drawing`` that is not a picture shows -- a shape, a group, a
    chart or SmartArt (:class:`docx2svg.model.Graphic`) -- or ``None``."""
    drawing = drawing_element(node)
    inline = child(drawing, "inline") if drawing is not None else None
    graphic_data = child(child(inline, "graphic"), "graphicData")
    if graphic_data is None:
        return None
    graphic = read_graphic_data(graphic_data, path, lambda element, where: _text_box_story(element, where, fields))
    return graphic if graphic.kind in ("shape", "group", "chart", "diagram") else None


#: The parser's step counter and its row- and cell-level walk, shared with
#: :func:`docx2svg.paths.resolve_path`, which finds the element a path names by the same
#: counts.
_Paths = Steps


#: Fields whose result the layout computes from its own pagination and the document; every
#: other field's result is drawn as Word cached it (ROADMAP.md, "Headers, footers and
#: fields").  :data:`CROSS_REFERENCES` are computed in the body only; ``STYLEREF`` in the
#: body and in headers and footers, not in a text box or a note (round 21).
COMPUTED_FIELDS = frozenset({"PAGE", "NUMPAGES", "SECTIONPAGES", "PAGEREF", "REF", "SEQ", "STYLEREF"})
#: Computed in the body's own text (cells included), from its bookmarks and its order
#: (``make_xref_field_probe.py``); in a header, a footer, a note or a text box drawn as
#: cached, as before.
CROSS_REFERENCES = frozenset({"PAGEREF", "REF", "SEQ"})


def field_keyword(instruction: str) -> str:
    """A field instruction's keyword, upper-cased (``PAGE`` of `` page \\* roman``)."""
    words = instruction.split()
    return words[0].upper() if words else ""


def merge_format(instruction: str) -> bool:
    """Whether the instruction carries ``\\* MERGEFORMAT``."""
    return "MERGEFORMAT" in instruction.upper().replace(" ", "")


@dataclass
class _Open:
    """A complex field between its ``begin`` and its ``end``."""

    #: Whether everything around it is drawn (it is not inside another field's
    #: instruction or computed result).
    visible: bool
    path: str
    instruction: list[str] = field(default_factory=list)
    #: ``instr`` until ``separate``, ``result`` after it.
    state: str = "instr"
    computed: bool = False
    #: The ``w:rPr`` of the first run of the instruction and of the cached result.
    instr_properties: Element | None = None
    instr_path: str | None = None
    result_properties: Element | None = None
    result: list[str] = field(default_factory=list)
    #: The cached result as ``(text, w:rPr)`` per piece of text.
    cached: list = field(default_factory=list)


class _Fields:
    """The complex fields open at this point of a part: a field's ``begin`` and ``end``
    may be in different runs, and a result (a table of contents) may span paragraphs."""

    def __init__(self, cached: list[str] | None = None, *, cross_references: bool = False,
                 styleref: bool = False) -> None:
        self.stack: list[_Open] = []
        #: The keyword of every field whose cached result is drawn.
        self.cached: list[str] = cached if cached is not None else []
        #: Whether :data:`CROSS_REFERENCES` are computed here: the body's own text, not a
        #: text box's.
        self.cross_references = cross_references
        #: Whether ``STYLEREF`` is computed here: the body's own text, a header's, a footer's.
        self.styleref = styleref
        #: The bookmarks met in the paragraph being read: ``(kind, id, name, run index)``.
        self.bookmarks: list | None = None

    def visible(self) -> bool:
        return all(f.state == "result" and not f.computed for f in self.stack)

    def in_toc_result(self) -> bool:
        """Whether this point is inside a ``TOC`` field's result (its entries)."""
        return any(f.state == "result" and field_keyword("".join(f.instruction)) == "TOC" for f in self.stack)

    def computes(self, instruction: str) -> bool:
        keyword = field_keyword(instruction)
        if keyword == "STYLEREF":
            return self.styleref
        return keyword in COMPUTED_FIELDS and (keyword not in CROSS_REFERENCES or self.cross_references)


def _computed_run(instruction: str, result: str, properties: Element | None, path: str, cached=()) -> Run:
    return Run(text=result, properties=_run_properties(properties), path=path,
               field=Field(instruction.strip(), field_keyword(instruction),
                           tuple((text, _run_properties(element)) for text, element in cached)))


def _runs(element: Element, path: str = "", fields: _Fields | None = None) -> list[Run]:
    """A ``w:r`` as the runs it contributes: its drawn text and marks, less what is inside a
    field's instruction or a computed field's cached result; and the computed field a
    ``w:fldChar w:fldCharType="end"`` in it closes (:data:`COMPUTED_FIELDS`)."""
    fields = fields if fields is not None else _Fields()
    out: list[Run] = []
    text_parts: list[str] = []
    marks: list[tuple[int, str]] = []
    drawings: list[str | PictureRef | Graphic | None] = []
    anchors: list[Anchor] = []
    length = 0
    properties = child(element, "rPr")
    plain = fields.visible()

    def flush() -> None:
        nonlocal text_parts, marks, drawings, anchors, length
        if text_parts or marks:
            out.append(Run(text="".join(text_parts), properties=_run_properties(properties), breaks=tuple(marks),
                           path=path, drawings=tuple(drawings), anchors=tuple(anchors)))
        text_parts, marks, drawings, anchors, length = [], [], [], [], 0

    #: Each kind of run object met so far (``w:pict``...), for a VML object's path.
    seen: dict[str, int] = {}
    for node in children(element):
        name = local_name(node.tag)
        if name == "fldChar":
            kind = attr(node, "fldCharType")
            if kind == "begin":
                fields.stack.append(_Open(fields.visible(), path))
            elif kind == "separate" and fields.stack:
                top = fields.stack[-1]
                top.state = "result"
                top.computed = fields.computes("".join(top.instruction))
            elif kind == "end" and fields.stack:
                top = fields.stack.pop()
                instruction = "".join(top.instruction)
                # A field with no result (no ``separate``) is computed as one with a result.
                computed = top.computed if top.state == "result" else fields.computes(instruction)
                if top.visible and fields.visible() and not computed:
                    fields.cached.append(field_keyword(instruction))
                if top.visible and fields.visible() and computed:
                    flush()
                    chosen = top.instr_properties
                    if merge_format(instruction) and top.result_properties is not None:
                        chosen = top.result_properties
                    out.append(_computed_run(instruction, "".join(top.result), chosen, top.instr_path or top.path,
                                             top.cached))
            continue
        if name == "instrText":
            # An instruction is never drawn: its text is the field's code (Phase 3 kept it
            # as run text, and measured it).
            if fields.stack and fields.stack[-1].state == "instr":
                top = fields.stack[-1]
                top.instruction.append(node.text or "")
                if top.instr_properties is None and top.instr_path is None:
                    top.instr_properties = properties
                    top.instr_path = path
            continue
        if not fields.visible():
            top = fields.stack[-1]
            if top.computed and top.state == "result" and all(
                    f.state == "result" and not f.computed for f in fields.stack[:-1]) and name in ("t", "delText"):
                top.result.append(node.text or "")
                top.cached.append((node.text or "", properties))
                if top.result_properties is None:
                    top.result_properties = properties if properties is not None else Element("rPr")
            continue
        if name in ("t", "delText"):
            # ``xml:space="preserve"`` matters and ElementTree keeps the text either
            # way; what must NOT happen is stripping it here.  A run of three spaces is
            # three spaces.
            piece = node.text or ""
            text_parts.append(piece)
            length += len(piece)
        elif name in _RUN_MARKS:
            marks.append((length, attr(node, "type") or name))
        elif name == "footnoteReference":
            # Which note, for the paginator: a note is laid out on its reference's page;
            # its number is drawn here (:mod:`docx2svg.notes`), unless a custom mark
            # follows.
            custom = ":custom" if attr(node, "customMarkFollows") in ("1", "true", "on") else ""
            marks.append((length, f"footnoteReference:{attr(node, 'id')}{custom}"))
        elif name == "endnoteReference":
            # Which note: its number is drawn here (:mod:`docx2svg.notes`); a custom mark
            # (``w:customMarkFollows``) draws none, the run after it being the mark.
            custom = ":custom" if attr(node, "customMarkFollows") in ("1", "true", "on") else ""
            marks.append((length, f"endnoteReference:{attr(node, 'id')}{custom}"))
        elif name in ("endnoteRef", "footnoteRef", "separator", "continuationSeparator"):
            # A note's own number; a separator's line (:mod:`docx2svg.notes`).
            marks.append((length, name))
        elif name in _RUN_OBJECTS:
            # A floating drawing is ``anchor:<k>``, the k-th of :attr:`Run.anchors`; an
            # inline one ``drawing:...`` with its picture in :attr:`Run.drawings`; anything
            # else (VML, an embedded object) ``drawing``.
            found = _anchor(node, f"{path}/wp:anchor[{sum(1 for a in anchors if not a.vml) + 1}]", fields)
            if found is not None:
                marks.append((length, f"anchor:{len(anchors)}"))
                anchors.append(found)
                continue
            seen[name] = seen.get(name, 0) + 1
            vml = read_vml(node, f"{path}/{'mc' if name == 'AlternateContent' else 'w'}:{name}[{seen[name]}]",
                           lambda element, where: _text_box_story(element, where, fields)) \
                if drawing_element(node) is None else None
            if vml is not None:
                # VML (:mod:`docx2svg.parse.vml`): a watermark's text path, an embedded
                # object's picture, an older shape -- floating, or a character of its line.
                _kind, value = vml
                if isinstance(value, Anchor):  # ``_kind`` "anchor"
                    marks.append((length, f"anchor:{len(anchors)}"))
                    anchors.append(value)
                else:
                    marks.append((length, value[0]))
                    drawings.append(value[1])
                continue
            mark = _object(node)
            marks.append((length, mark))
            if mark.startswith("drawing:"):
                drawings.append(_picture(node) or _inline_graphic(node, f"{path}/wp:inline[{len(drawings) + 1}]",
                                                                   fields))
    flush()
    if not out and plain and fields.visible() and not any(
            local_name(node.tag) in ("fldChar", "instrText") for node in children(element)):
        # A run of nothing (properties alone) is still a run, as it always was.
        out.append(Run(text="", properties=_run_properties(properties), path=path))
    return out


def _run(element: Element, path: str = "") -> Run:
    """A ``w:r`` outside any field, as one run (its text, marks and objects)."""
    found = _runs(element, path)
    return found[0] if found else Run(text="", properties=_run_properties(child(element, "rPr")), path=path)


def _simple_field(node: Element, step: "_Paths", fields: _Fields) -> list[Run]:
    """A ``w:fldSimple``: a computed field as one run (formatted as its result's first run
    with ``\\* MERGEFORMAT``, else as the paragraph's text: ``make_field_probe.py``, ``f5``);
    any other, its cached result's runs."""
    here = step(node)
    inner_step = _Paths(here)
    instruction = attr(node, "instr") or ""
    inner = [run for run in children(node, "r")]
    if fields.computes(instruction):
        if not fields.visible():
            return []
        cached = [((t.text or ""), child(run, "rPr")) for run in inner for t in children(run, "t")]
        text = "".join(piece for piece, _ in cached)
        properties = child(inner[0], "rPr") if inner and merge_format(instruction) else None
        return [_computed_run(instruction, text, properties, here, cached)]
    out: list[Run] = []
    if fields.visible():
        fields.cached.append(field_keyword(instruction))
    for run in inner:
        out.extend(_runs(run, inner_step(run), fields))
    return out


#: Run-level containers the final view of a document draws the content of (ROADMAP.md,
#: "Revisions -- measured"): an insertion, a move's destination, a content control's
#: ``w:sdtContent``, and those that wrap runs without changing their layout.
_RUN_CONTAINERS = frozenset({"hyperlink", "ins", "moveTo", "smartTag", "customXml", "sdt", "dir", "bdo"})


def _inline(element: Element, step: "_Paths", fields: _Fields, runs: list) -> None:
    """``element``'s runs into ``runs``, descending into :data:`_RUN_CONTAINERS`; a
    deletion and a move's source (``w:del``, ``w:moveFrom``) are not drawn."""
    for node in children(element):
        name = local_name(node.tag)
        if name == "r":
            runs.extend(_runs(node, step(node), fields))
        elif name == "fldSimple":
            runs.extend(_simple_field(node, step, fields))
        elif name in ("bookmarkStart", "bookmarkEnd"):
            step(node)
            if fields.bookmarks is not None and fields.visible():
                kind = "start" if name == "bookmarkStart" else "end"
                fields.bookmarks.append((kind, attr(node, "id"), attr(node, "name"), len(runs)))
        elif name in _RUN_CONTAINERS:
            here = step(node)
            if name == "sdt":
                content = child(node, "sdtContent")
                if content is not None:
                    _inline(content, _Paths(f"{here}/w:sdtContent[1]"), fields, runs)
            elif name == "hyperlink" and fields.in_toc_result():
                # A table of contents' entry (Run.toc_link): a link in a TOC field's result.
                start = len(runs)
                _inline(node, _Paths(here), fields, runs)
                runs[start:] = [dataclasses.replace(run, toc_link=True) for run in runs[start:]]
            else:
                _inline(node, _Paths(here), fields, runs)
        else:
            step(node)


def _mark_removed(element: Element) -> bool:
    """Whether a paragraph's mark is deleted, or moved away, in the final view."""
    mark = child(child(element, "pPr"), "rPr")
    return mark is not None and (child(mark, "del") is not None or child(mark, "moveFrom") is not None)


def _paragraph(element: Element, table_style: str | None = None,
               conditions: tuple[str, ...] = (), path: str = "", fields: _Fields | None = None) -> Paragraph:
    runs: list[Run] = []
    fields = fields if fields is not None else _Fields()
    outer, fields.bookmarks = fields.bookmarks, []
    try:
        _inline(element, _Paths(""), fields, runs)
        bookmarks = tuple(fields.bookmarks)
    finally:
        fields.bookmarks = outer
    return Paragraph(
        runs=tuple(runs),
        properties=_paragraph_properties(child(element, "pPr")),
        table_style_id=table_style,
        table_conditions=conditions,
        path=path,
        para_id=element.get(_W14_PARA_ID),
        mark_deleted=_mark_removed(element),
        section=child(child(element, "pPr"), "sectPr"),
        bookmarks=bookmarks,
    )


def _empty(paragraph: Paragraph) -> bool:
    """Whether nothing of the paragraph is drawn but its mark."""
    return all(not run.text and not run.breaks and run.field is None for run in paragraph.runs)


def _joined(first: Paragraph, following: Paragraph) -> Paragraph:
    """``first``, whose mark is deleted, joined to ``following``: the runs of both, under
    ``following``'s paragraph properties and section break (measured in Word's final view,
    every compatibility mode).  The joined paragraph keeps ``first``'s path and id."""
    count = len(first.runs)
    joins = ((count,) if following.runs else ()) + tuple(count + k for k in following.joins)
    bookmarks = first.bookmarks + tuple((kind, mark, name, count + k) for kind, mark, name, k in following.bookmarks)
    return dataclasses.replace(following, runs=first.runs + following.runs, path=first.path,
                               para_id=first.para_id, joins=joins, bookmarks=bookmarks)


def _join(blocks: list) -> list:
    """The final view of a container's blocks -- what Accept All leaves: each paragraph
    whose mark is deleted joined to the paragraph after it (a chain into the last); before
    a table, to the first paragraph of the table's first cell (:func:`_joined_into`); one
    with no paragraph after it stays."""
    out: list = []
    for block in reversed(blocks):
        if isinstance(block, Paragraph) and block.mark_deleted:
            following = out[-1] if out else None
            if isinstance(following, Paragraph):
                out[-1] = _joined(block, following)
                continue
            if isinstance(following, Table) and block.section is None:
                table = _joined_into(block, following)
                if table is not None:
                    out[-1] = table
                    continue
                if _empty(block):
                    continue
        out.append(block)
    return out[::-1]


def _joined_into(first: Paragraph, table: Table) -> Table | None:
    """``table`` with ``first`` (a paragraph whose mark is deleted, just before it) joined to
    the first paragraph of its first cell -- a nested table's, where the cell begins with
    one -- as Word's Accept All moves it (measured, round 54 lane J: ROADMAP.md,
    "Revisions -- measured", R.3): ``first``'s runs first, under the cell paragraph's
    properties, the cell paragraph's path and id kept (it is the cell's paragraph); ``None``
    where the first cell holds no paragraph to join."""
    if not table.rows or not table.rows[0]:
        return None
    cell = table.rows[0][0]
    if not cell:
        return None
    head: Paragraph | Table | None = cell[0]
    if isinstance(head, Paragraph):
        head = dataclasses.replace(_joined(first, head), path=head.path, para_id=head.para_id)
    elif isinstance(head, Table):
        head = _joined_into(first, head)
    if head is None or not isinstance(head, (Paragraph, Table)):
        return None
    rows = ((((head,) + tuple(cell[1:])),) + tuple(table.rows[0][1:]),) + tuple(table.rows[1:])
    return dataclasses.replace(table, rows=rows, joined=True)


def _block_list(parent: Element, table_style: str | None, conditions: tuple[str, ...], path: str,
                fields: _Fields) -> list:
    out: list = []
    step = _Paths(path)
    for node in children(parent):
        name = local_name(node.tag)
        here = step(node)
        _block_of(node, name, here, table_style, conditions, fields, out)
    return out


def _block_of(node: Element, name: str, here: str, table_style: str | None, conditions: tuple[str, ...],
              fields: _Fields, out: list) -> None:
    """The blocks of one of a container's elements (:func:`_block_list`) into ``out``."""
    if name == "p":
        out.append(_paragraph(node, table_style, conditions, here, fields))
    elif name == "tbl":
        # A table every row of which is deleted is not drawn at all.
        out.extend(_table_blocks(node, here, fields))
    elif name == "sdt":
        # A block-level content control: its content, as the container's.
        content = child(node, "sdtContent")
        if content is not None:
            out.extend(_block_list(content, table_style, conditions, f"{here}/w:sdtContent[1]", fields))
    elif name == "customXml":
        out.extend(_block_list(node, table_style, conditions, here, fields))


class KeptBlocks:
    """The body's top-level blocks parsed before, kept to build the next version's model
    from (round 41): a paragraph's or table's model is made from its element's bytes, the
    namespaces in scope (the main part's bytes before ``w:body``'s content) and the theme's
    colours as the settings map them -- nothing else of the package -- so a block whose
    bytes and context are the same is the same model, and is taken as it was rather than
    built again.  A paragraph is kept by those alone and takes its new place
    (:attr:`~docx2svg.model.Paragraph.path`, the only thing in it its place sets: its runs'
    paths are its own); a table or a content control by its place too (its cells'
    paragraphs' paths start with it).

    Kept only where its parse changed none of the complex fields open at it -- it opened and
    closed its own, if any: a table of contents' line inside the field's result, not the
    paragraph that opens it -- and taken only where the fields open at it are in the same
    state (each one's state, whether it is computed, its keyword: all a block's parse reads
    of them); the field keywords it adds to ``Document.cached_fields`` are kept with it.
    Every top-level block of a document is its own object: one taken is a copy of the one
    kept (a paragraph) or made anew from its cells (a table, :func:`_following_properties`),
    and an element whose bytes occur twice in a document is built the second time.  Nothing
    a layout does changes a block's model (``tests/test_kept_blocks.py``).

    Holds the blocks of the last document parsed with it (each parse keeps those it used or
    made, and drops the rest): one version's model, which the layout it goes with holds
    anyway.  Process-local: a pickle of it is an empty one.  Used by one parse at a time (a
    layout held by one task); where two parses ran at once, each would keep what it found --
    the same models -- and the last one's set stands."""

    __slots__ = ("_entries", "hits", "built")

    def __init__(self) -> None:
        self._entries: dict = {}
        #: Of the last parse: the body's elements taken as kept, and built.
        self.hits = 0
        self.built = 0

    def __len__(self) -> int:
        return len(self._entries)

    def __reduce__(self):
        return (KeptBlocks, ())

    def clear(self) -> None:
        self._entries = {}


#: The body's elements that are blocks of the model, or hold them.
_BLOCK_ELEMENTS = frozenset({"p", "tbl", "sdt", "customXml"})


def _taken(block, path: str | None):
    """A kept top-level block, as this document's own object: a paragraph copied (at
    ``path``, where kept by its bytes alone), a table as it is (its document's
    :func:`_following_properties` makes it anew)."""
    if isinstance(block, Paragraph):
        out = object.__new__(Paragraph)
        out.__dict__.update(block.__dict__)
        if path is not None:
            object.__setattr__(out, "path", path)
        return out
    return block


def _open_state(stack: list) -> tuple:
    """What a block's parse reads of the complex fields open at it (:class:`_Fields`:
    ``visible``, ``in_toc_result``, the result a computed field collects): each one's state,
    whether it is computed, and its keyword."""
    return tuple((f.state, f.computed, field_keyword("".join(f.instruction))) for f in stack)


def _open_snapshot(stack: list) -> tuple:
    """The complex fields open, as objects and everything a block's parse may change of them:
    the same before and after a block, the block changed none of them (it opened and closed
    its own fields only)."""
    return tuple((id(f), f.state, f.computed, len(f.instruction), len(f.result), len(f.cached),
                  id(f.instr_properties), f.instr_path, id(f.result_properties)) for f in stack)


def _kept_block_list(body: Element, fields: _Fields, kept: KeptBlocks, segments: list, context: bytes) -> list:
    """The body's blocks (:func:`_block_list` of ``w:body``), each top-level element's
    taken from ``kept`` where its segment's digest and ``context`` are those it was built
    from (:class:`KeptBlocks`), else built; ``kept`` then holds this body's."""
    nodes = children(body)
    if len(nodes) != len(segments) or any(local_name(node.tag) != segment.name
                                         for node, segment in zip(nodes, segments)):
        # The scan of the bytes and the parse disagree: nothing is taken, nor kept.
        kept.clear()
        kept.hits, kept.built = 0, len(nodes)
        return _block_list(body, None, (), "w:body", fields)
    before = kept._entries
    now: dict = {}
    used: set = set()
    hits = built = 0
    out: list = []
    step = _Paths("w:body")
    for node, segment in zip(nodes, segments):
        here = step(node)
        name = segment.name
        if name not in _BLOCK_ELEMENTS:
            continue
        key = (context, name, segment.digest, None if name == "p" else here, _open_state(fields.stack))
        entry = now.get(key) or before.get(key)
        if entry is not None and id(entry) not in used:
            used.add(id(entry))
            now[key] = entry
            blocks, cached = entry
            fields.cached.extend(cached)
            out.extend(_taken(block, here if name == "p" else None) for block in blocks)
            hits += 1
            continue
        open_before = _open_snapshot(fields.stack)
        start, cached_start = len(out), len(fields.cached)
        _block_of(node, name, here, None, (), fields, out)
        built += 1
        if _open_snapshot(fields.stack) == open_before and key not in now:
            entry = (tuple(out[start:]), tuple(fields.cached[cached_start:]))
            used.add(id(entry))
            now[key] = entry
    kept._entries = now
    kept.hits, kept.built = hits, built
    return out


def _blocks(parent: Element, table_style: str | None = None,
            conditions: tuple[str, ...] = (), path: str = "", fields: _Fields | None = None) -> list:
    """A part's (or a cell's) paragraphs and tables, in order, in the final view of its
    tracked changes (:func:`_join`).  ``fields`` carries the complex fields open across
    them: one per part."""
    fields = fields if fields is not None else _Fields()
    return _join(_block_list(parent, table_style, conditions, path, fields))


def _following_properties(blocks: list, following: ParagraphProperties | None = None
                          ) -> tuple[list, ParagraphProperties | None]:
    """``blocks`` with each paragraph whose mark is deleted and that joined nothing drawn
    under the paragraph properties of the paragraph after it in the story -- the first of
    a table's first cell, of the next cell, of the body after the table -- or under none of
    its own when no paragraph follows (measured in Word's final view); and the first
    paragraph's properties, for the blocks before these."""
    out: list = []
    for block in reversed(blocks):
        if isinstance(block, Paragraph):
            if block.mark_deleted:
                block = dataclasses.replace(block, properties=following)
            following = block.properties
        elif isinstance(block, Table):
            rows = []
            for row in reversed(block.rows):
                cells = []
                for cell in reversed(row):
                    cell_blocks, following = _following_properties(list(cell), following)
                    cells.append(tuple(cell_blocks))
                rows.append(tuple(cells[::-1]))
            block = dataclasses.replace(block, rows=tuple(rows[::-1]))
        out.append(block)
    return out[::-1], following


def _story(parent: Element, path: str = "", fields: _Fields | None = None) -> list:
    """A story's blocks (a part's, a note's, a text box's) in the final view."""
    return _following_properties(_blocks(parent, path=path, fields=fields))[0]


def _text_box_story(parent: Element, path: str, fields: _Fields) -> list:
    """A text box's blocks: the part's fields, but no cross-reference computed in it (its
    bookmarks are not the body's, nor are its sequences counted with the body's)."""
    outer, fields.cross_references = fields.cross_references, False
    styleref, fields.styleref = fields.styleref, False
    try:
        return _story(parent, path=path, fields=fields)
    finally:
        fields.cross_references = outer
        fields.styleref = styleref


def _table_blocks(element: Element, path: str = "", fields: _Fields | None = None) -> list:
    """A ``w:tbl``'s blocks in the final view: the table (:func:`_table`), or -- where a
    cell's last paragraph mark is deleted in its row's first cell, and the cell holds
    paragraphs before it -- those paragraphs moved out of the table before that row, the
    table split there as Word's Accept All leaves it (:func:`_cell_ends`)."""
    table, moved, grid_of = _table(element, path, fields)
    if not moved:
        return [table] if table.rows else []
    out: list = []
    start = 0
    for row, blocks in moved:
        if row > start:
            out.append(_table_part(table, start, row, grid_of))
        out.extend(blocks)
        start = row
    out.append(_table_part(table, start, len(table.rows), grid_of))
    return out


def _table_part(table: Table, start: int, end: int, grid_of: Callable) -> Table:
    """Rows ``start``..``end`` of ``table``, a table of their own, its grid the columns
    some cell of theirs covers (:func:`_vacated_columns`, applied to the part)."""
    grid = grid_of(start, end)
    return Table(style_id=table.style_id, rows=table.rows[start:end], path=table.path, grid=grid,
                 properties=table.properties, row_properties=table.row_properties[start:end],
                 cell_properties=table.cell_properties[start:end], look=table.look,
                 conditions=table.conditions[start:end], cell_paths=table.cell_paths[start:end], joined=True)


def _last_mark_removed(container: Element) -> bool:
    """Whether the last paragraph of a cell (or a block-level control's content) has its
    mark deleted: the cell's last block is that paragraph."""
    for node in reversed(list(children(container))):
        name = local_name(node.tag)
        if name == "p":
            return _mark_removed(node)
        if name == "tbl":
            return False
        if name == "sdt":
            content = child(node, "sdtContent")
            if content is not None and any(local_name(n.tag) in ("p", "tbl", "sdt", "customXml")
                                           for n in children(content)):
                return _last_mark_removed(content)
        elif name == "customXml":
            if any(local_name(n.tag) in ("p", "tbl", "sdt", "customXml") for n in children(node)):
                return _last_mark_removed(node)
    return False


def _cell_ends(cells: list) -> tuple[list[list[int]], set[int]]:
    """What Word's Accept All makes of a row's cells (its ``w:tc`` elements, deleted ones
    left out) whose last paragraph mark is deleted (measured, round 54 lane J: ROADMAP.md,
    "Revisions -- measured", R.3): the cell takes the next cell in -- its last paragraph
    joined to the next cell's first, the next cell's other blocks after it, the next cell
    gone, its own span and width kept -- and a cell with no next cell in the row is gone
    with its content.  Returns the row's cells as groups of indexes (the first of each the
    cell kept, the others taken into it) and the indexes of the cells gone."""
    groups: list[list[int]] = []
    gone: set[int] = set()
    n = 0
    while n < len(cells):
        group = [n]
        while _last_mark_removed(cells[group[-1]]) and group[-1] + 1 < len(cells):
            group.append(group[-1] + 1)
        if _last_mark_removed(cells[group[-1]]):
            gone.update(group)
        else:
            groups.append(group)
            gone.update(group[1:])
        n = group[-1] + 1
    return groups, gone


def _cell_taken(blocks: list, taken: list) -> list:
    """A cell's blocks with the next cell's (``taken``) taken in: its last paragraph (its
    mark deleted) joined to theirs first (:func:`_joined`, under the next one's properties;
    the joined paragraph keeps the cell's path), the rest after."""
    if blocks and taken and isinstance(blocks[-1], Paragraph) and isinstance(taken[0], Paragraph):
        return blocks[:-1] + [_joined(blocks[-1], taken[0])] + list(taken[1:])
    return list(blocks) + list(taken)


def _moved_out(blocks: list) -> list:
    """A cell's blocks moved out of its table: no longer under the table's style."""
    return [dataclasses.replace(block, table_style_id=None, table_conditions=())
            if isinstance(block, Paragraph) else block for block in blocks]


def _table(element: Element, path: str = "", fields: _Fields | None = None) -> tuple[Table, list, Callable]:
    """A table, the blocks the final view moves out of it (``[(row, blocks), ...]``: before
    that row, :func:`_table_blocks`) and its grid for a run of its rows.  The table: its rows and cells, the table style, and what its geometry is read from
    -- ``w:tblGrid`` and every ``w:tblPr``, ``w:trPr`` and ``w:tcPr`` as declared.

    The cell paragraphs carry the table's style id (``""`` when it names none), because
    the table style is a level of their cascade.  A ``w:tc`` inside a row-level
    container (``w:sdt``, ``w:customXml``) is read as the row's cell.
    """
    properties = child(element, "tblPr")
    style = attr(child(properties, "tblStyle"), "val") or ""
    look = _table_look(child(properties, "tblLook"))
    # The final view is the document with every revision accepted (round 54, R110): a
    # deleted row (``w:trPr/w:del``) is not drawn, whether or not its content is deleted too;
    # nor is a cell with a pending deletion (``w:cellDel``: a tracked column deletion, a
    # tracked merge's cell merged away), the grid columns only such cells covered going with
    # it -- as Word's Accept All leaves the table (measured, ROADMAP.md "Tables -- measured",
    # stage 7f).  Word's *No Markup* screen still draws such a cell until it is accepted.
    # Paths keep their places among the document's rows and cells.
    # A row all of whose cells are deleted goes with them.
    row_elements = [(i, row) for i, row in enumerate(_row_level(element, "tr"))
                    if child(child(row, "trPr"), "del") is None
                    and not _all_cells_deleted(row)]
    grid_nodes = list(children(child(element, "tblGrid"), "gridCol"))
    # A cell whose last paragraph mark is deleted takes the next cell in, or goes at its
    # row's end (:func:`_cell_ends`): the cells gone are left out as deleted ones are.
    plans = []
    gone_cells: set[int] = set()
    for i, row in row_elements:
        shown = [cell for cell in _row_level(row, "tc") if not _cell_deleted(cell)]
        groups, gone = _cell_ends(shown)
        plans.append((shown, groups))
        gone_cells.update(id(shown[n]) for n in gone)
    rows, row_properties, cell_properties, conditions, cell_paths = [], [], [], [], []
    moved: list = []
    for k, ((i, row), (shown, groups)) in enumerate(zip(row_elements, plans)):
        kept = {id(shown[group[0]]) for group in groups}
        taken = {id(shown[group[0]]): [shown[n] for n in group[1:]] for group in groups}
        cells: list[tuple[int, Element]] = []
        properties_of: list[dict] = []
        for j, cell in enumerate(_row_level(row, "tc")):
            if id(cell) in kept:
                cells.append((j, cell))
                properties_of.append(_accepted_cell_properties(cell))
            elif cells and _cell_deleted(cell) and _content_kept(cell):
                _absorb(properties_of[-1], read_table_properties(child(cell, "tcPr")))
        row_conditions = tuple(cell_conditions(look, k, n, len(row_elements), len(cells)) for n in range(len(cells)))
        paths = tuple(f"{path}/w:tr[{i + 1}]/w:tc[{j + 1}]" for j, _ in cells)
        row_blocks = []
        for n, (j, cell) in enumerate(cells):
            blocks = _blocks(cell, style, row_conditions[n], paths[n], fields)
            others = taken[id(cell)]
            if others:
                if n == 0 and len(blocks) > 1:
                    # The row's first cell's blocks before its last paragraph leave the
                    # table, before the row (measured: Word's Accept All).
                    moved.append((k, _moved_out(blocks[:-1])))
                    blocks = blocks[-1:]
                elements = list(_row_level(row, "tc"))
                for other in others:
                    there = f"{path}/w:tr[{i + 1}]/w:tc[{elements.index(other) + 1}]"
                    blocks = _cell_taken(blocks, _blocks(other, style, row_conditions[n], there, fields))
            row_blocks.append(tuple(blocks))
        rows.append(tuple(row_blocks))
        row_properties.append(read_table_properties(child(row, "trPr")))
        cell_properties.append(tuple(properties_of))
        conditions.append(row_conditions)
        cell_paths.append(paths)

    def grid_of(start: int, end: int) -> tuple[int, ...]:
        vacated = _vacated_columns([row for _, row in row_elements[start:end]], len(grid_nodes), gone_cells)
        return tuple(attr_int(column, "w", 0) or 0 for k, column in enumerate(grid_nodes) if k not in vacated)

    table = Table(style_id=style or None, rows=tuple(rows), path=path, grid=grid_of(0, len(rows)),
                  properties=read_table_properties(properties), row_properties=tuple(row_properties),
                  cell_properties=tuple(cell_properties), look=look, conditions=tuple(conditions),
                  cell_paths=tuple(cell_paths))
    return table, moved, grid_of


def _all_cells_deleted(row: Element) -> bool:
    cells = list(_row_level(row, "tc"))
    return bool(cells) and all(_cell_deleted(cell) for cell in cells)


#: Run content a reader sees (a deleted cell whose content holds some is kept: its room).
_SHOWN = frozenset({"t", "tab", "br", "cr", "drawing", "pict", "object", "sym", "noBreakHyphen", "ptab"})


def _content_kept(cell: Element) -> bool:
    """Whether a deleted cell keeps anything: content not deleted (outside ``w:del`` and
    ``w:moveFrom``) or a paragraph mark of its own not deleted.  Word's Accept All then
    widens the cell before it by its columns and width, its content going with it -- an
    empty cell, or one whose text alone is deleted, too -- and removes only a cell whose
    content and marks are all deleted (measured, round 54: ROADMAP.md, stage 7f; lane M's
    ``m-empty`` and ``c-textdel``).  The same rule as docx-agent's accept."""
    def shown(node: Element, nested: bool) -> bool:
        for item in children(node):
            name = local_name(item.tag)
            if name in ("del", "moveFrom", "tcPr", "pPr", "rPr"):
                continue                      # deleted content, or properties
            if name == "p" and not nested and not _mark_removed(item):
                return True                   # a mark of the cell's own, kept
            if (name in _SHOWN and (name != "t" or item.text)) or shown(item, nested or name == "tbl"):
                return True
        return False

    return shown(cell, False)


def _absorb(previous: dict, deleted: dict) -> None:
    """``previous`` (a cell's properties) widened by a deleted cell's columns and width."""
    previous["gridSpan"] = max(1, previous.get("gridSpan", 1) or 1) + max(1, deleted.get("gridSpan", 1) or 1)
    mine, theirs = previous.get("tcW"), deleted.get("tcW")
    if mine and theirs and mine[1] == theirs[1] == "dxa":
        previous["tcW"] = (mine[0] + theirs[0], "dxa")


def _cell_deleted(cell: Element) -> bool:
    """Whether ``cell`` is a pending tracked deletion (``w:tcPr/w:cellDel``)."""
    return child(child(cell, "tcPr"), "cellDel") is not None


def _accepted_cell_properties(cell: Element) -> dict:
    """A cell's properties with its revisions accepted: a ``w:cellMerge`` (a tracked merge
    down) gives the cell its vertical merge, as accepting it does -- ``rest`` a merge's
    first cell, ``cont`` the cells below it."""
    properties = child(cell, "tcPr")
    out = read_table_properties(properties)
    merge = attr(child(properties, "cellMerge"), "vMerge")
    if merge in ("rest", "cont"):
        out["vMerge"] = "restart" if merge == "rest" else "continue"
    return out


def _vacated_columns(rows: list[Element], count: int, gone: AbstractSet[int] = frozenset()) -> frozenset[int]:
    """The grid columns (of ``count``) that only cells with a pending deletion cover, in
    every one of ``rows`` that covers them at all: accepting the deletions removes them
    from the grid, the other columns keeping their widths (Word's Accept All, measured).
    A column some row skips (``w:gridBefore``, ``w:gridAfter``) is kept.  ``gone``: the
    ``id`` of other cells the final view leaves out (:func:`_cell_ends`)."""
    if not count:
        return frozenset()
    covered: dict[int, bool] = {}
    for row in rows:
        properties = child(row, "trPr")
        column = attr_int(child(properties, "gridBefore"), "val", 0) or 0
        skipped = set(range(column))
        kept_before = False
        for cell in _row_level(row, "tc"):
            span = max(1, attr_int(child(child(cell, "tcPr"), "gridSpan"), "val", 1) or 1)
            # A deleted cell whose content is kept goes to the cell before it (_absorb).
            deleted = (_cell_deleted(cell) and not (kept_before and _content_kept(cell))) or id(cell) in gone
            kept_before = kept_before or not _cell_deleted(cell)
            for c in range(column, column + span):
                covered[c] = covered.get(c, True) and deleted
            column += span
        after = attr_int(child(properties, "gridAfter"), "val", 0) or 0
        skipped |= set(range(column, column + after))
        for c in skipped:
            covered[c] = False
    return frozenset(c for c, only_deleted in covered.items() if only_deleted and c < count)


_row_level = row_level


#: ``w:tblLook/@w:val``'s bits (the pre-2010 spelling of the same flags).
_LOOK_BITS = {"firstRow": 0x20, "lastRow": 0x40, "firstColumn": 0x80, "lastColumn": 0x100,
              "noHBand": 0x200, "noVBand": 0x400}


def _table_look(element: Element | None) -> dict[str, bool]:
    """Which of the table style's conditional formats the table switches on."""
    if element is None:
        # ECMA-376's default look: header row, first column and row banding on.
        return {"firstRow": True, "lastRow": False, "firstColumn": True, "lastColumn": False,
                "noHBand": False, "noVBand": True}
    look = {}
    mask = attr(element, "val")
    bits = int(mask, 16) if mask and all(c in "0123456789abcdefABCDEF" for c in mask) else 0
    for name, bit in _LOOK_BITS.items():
        explicit = attr(element, name)
        look[name] = attr_bool(element, name) if explicit is not None else bool(bits & bit)
    return look


def cell_conditions(look: dict[str, bool], row: int, column: int, rows: int, columns: int) -> tuple[str, ...]:
    """The conditional formats for cell (row, column), in ECMA-376 17.7.6's order.

    Whole table first, then banding, then first/last column, then first/last row, then
    the corner cells -- each later one overriding the earlier.  Band sizes are taken as
    1 (``w:tblStyleRowBandSize`` is not read).  Measured on one case only: in
    ``sample-simple.docx`` the top-left cell, where a bold first row meets a bold first
    column, is drawn bold -- so conditional formats override one another inside the
    table level rather than toggling.
    """
    first_row, last_row = look["firstRow"] and row == 0, look["lastRow"] and row == rows - 1
    first_col, last_col = look["firstColumn"] and column == 0, look["lastColumn"] and column == columns - 1
    out = ["wholeTable"]
    if not look["noVBand"] and not (first_col or last_col):
        index = column - (1 if look["firstColumn"] else 0)
        out.append("band1Vert" if index % 2 == 0 else "band2Vert")
    if not look["noHBand"] and not (first_row or last_row):
        index = row - (1 if look["firstRow"] else 0)
        out.append("band1Horz" if index % 2 == 0 else "band2Horz")
    out += [name for name, on in (("firstCol", first_col), ("lastCol", last_col),
                                  ("firstRow", first_row), ("lastRow", last_row)) if on]
    if (first_row or last_row) and (first_col or last_col):
        out.append(("nw" if first_row else "sw") + "Cell" if first_col
                   else ("ne" if first_row else "se") + "Cell")
    return tuple(out)


#: The old name, kept for docx-agent until it moves.
_conditions = cell_conditions


def _page_size(section_properties: Element) -> PageSize:
    element = child(section_properties, "pgSz")
    return PageSize(
        width_twips=attr_int(element, "w", 12240) or 12240,
        height_twips=attr_int(element, "h", 15840) or 15840,
        orientation=attr(element, "orient"),
    )


def _margins(section_properties: Element) -> PageMargins:
    element = child(section_properties, "pgMar")
    if element is None:
        return PageMargins()
    # A stated 0 is 0: a header at ``w:header="0"`` is drawn at the page's top edge
    # (``make_story_probe.py``, ``distance 0``), not at the default distance.
    return PageMargins(
        top=_or(attr_int(element, "top"), 1440),
        right=_or(attr_int(element, "right"), 1440),
        bottom=_or(attr_int(element, "bottom"), 1440),
        left=_or(attr_int(element, "left"), 1440),
        header=_or(attr_int(element, "header"), 720),
        footer=_or(attr_int(element, "footer"), 720),
        gutter=_or(attr_int(element, "gutter"), 0),
    )


def _or(value: int | None, default: int) -> int:
    return default if value is None else value


def _section(section_properties: Element, first: int, last: int) -> Section:
    cols = child(section_properties, "cols")
    columns: tuple[tuple[int, int], ...] = ()
    if cols is not None and attr(cols, "equalWidth") in ("0", "false"):
        columns = tuple((attr_int(col, "w", 0) or 0, attr_int(col, "space", 0) or 0)
                        for col in children(cols, "col"))
    return Section(
        page_size=_page_size(section_properties),
        margins=_margins(section_properties),
        start_type=attr(child(section_properties, "type"), "val"),
        first_paragraph=first,
        last_paragraph=last,
        column_count=max(1, attr_int(cols, "num", 1) or 1) if cols is not None else 1,
        column_space=attr_int(cols, "space", 720) if cols is not None and attr(cols, "space") else 720,
        columns=columns,
        column_separator=attr(cols, "sep") in ("1", "true", "on") if cols is not None else False,
        headers=_references(section_properties, "headerReference"),
        footers=_references(section_properties, "footerReference"),
        title_page=child(section_properties, "titlePg") is not None
        and attr_bool(child(section_properties, "titlePg")),
        page_number_start=attr_int(child(section_properties, "pgNumType"), "start"),
        vertical_alignment=attr(child(section_properties, "vAlign"), "val"),
        page_number_format=attr(child(section_properties, "pgNumType"), "fmt"),
        endnote_format=attr(child(child(section_properties, "endnotePr"), "numFmt"), "val"),
        endnote_start=attr_int(child(child(section_properties, "endnotePr"), "numStart"), "val"),
        endnote_restart=attr(child(child(section_properties, "endnotePr"), "numRestart"), "val"),
        footnote_format=attr(child(child(section_properties, "footnotePr"), "numFmt"), "val"),
        footnote_start=attr_int(child(child(section_properties, "footnotePr"), "numStart"), "val"),
        footnote_restart=attr(child(child(section_properties, "footnotePr"), "numRestart"), "val"),
        footnote_position=attr(child(child(section_properties, "footnotePr"), "pos"), "val"),
        line_numbering=_line_numbering(child(section_properties, "lnNumType")),
    )


def _line_numbering(element: Element | None) -> tuple | None:
    """``w:lnNumType`` as :attr:`Section.line_numbering` holds it."""
    if element is None:
        return None
    count_by = attr_int(element, "countBy")
    if count_by is not None and count_by < 1:
        return None
    return (count_by or 1, attr_int(element, "start") or 0, attr_int(element, "distance"),
            attr(element, "restart") or "newPage")


_R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def _references(section_properties: Element, name: str) -> tuple[tuple[str, str], ...]:
    return tuple((attr(node, "type") or "default", rid)
                 for node in children(section_properties, name) if (rid := node.get(_R_ID)))


def _sections(body: Element, paragraphs: list[Paragraph]) -> list[Section]:
    """Discover sections.

    **The awkward part of the format.**  A naive reading looks for ``w:sectPr`` under
    ``w:body`` and finds exactly one -- the last.  Every *earlier* section's ``w:sectPr``
    is a child of the ``w:pPr`` of the last paragraph belonging to it, which means the
    element that describes a section sits at the end of that section rather than at its
    start, and a document with three page sizes has two of them hidden inside paragraph
    properties.

    Two consequences the rest of this project depends on.  First, a reader that only
    looks under ``w:body`` silently treats a multi-section document as single-section,
    which is a whole-document layout error with no error message.  Second, the paragraph
    carrying a ``w:sectPr`` is a real paragraph that still renders -- it belongs to the
    section it terminates, not to the next one, which is why ``last_paragraph`` below is
    inclusive of it.
    """
    sections: list[Section] = []
    first = 0
    for index, paragraph in enumerate(paragraphs):
        # The final view's paragraphs (:func:`_join`): a section break on a paragraph whose
        # mark is deleted is gone with the mark, its section one with the next.
        section_properties = paragraph.section
        if section_properties is not None:
            sections.append(_section(section_properties, first, index + 1))
            first = index + 1

    final = child(body, "sectPr")
    if final is not None:
        sections.append(_section(final, first, len(paragraphs)))
    elif not sections:
        # A body with no ``w:sectPr`` at all is legal; Word applies its own defaults.
        # Recording the fact beats inventing a page size without saying so.
        sections.append(
            Section(
                page_size=PageSize(12240, 15840),
                margins=PageMargins(),
                first_paragraph=0,
                last_paragraph=len(paragraphs),
            )
        )
    return sections


def _doc_defaults(styles_xml: bytes) -> tuple[RunProperties | None, ParagraphProperties | None]:
    root = parse_xml(styles_xml)
    defaults = child(root, "docDefaults")
    if defaults is None:
        return None, None
    run_default = child(child(defaults, "rPrDefault"), "rPr")
    paragraph_default = child(child(defaults, "pPrDefault"), "pPr")
    return _run_properties(run_default), _paragraph_properties(paragraph_default)


def parse_document(document_xml: bytes, styles_xml: bytes | None = None, *, kept: KeptBlocks | None = None,
                   segments: list | None = None, context: bytes = b"") -> Document:
    """Parse ``word/document.xml`` (and optionally ``word/styles.xml``).  With ``kept``
    (round 41, :class:`KeptBlocks`), ``segments`` the body's elements as
    :func:`docx2svg.relayout.segments` reads them from these bytes (read here where not
    given) and ``context`` what else the body's models are made from (the theme's colours,
    :func:`parse_package`): the blocks kept from the last parse with ``kept`` are taken
    where their bytes and context are the same -- the same document either way."""
    root = parse_xml(document_xml)
    body = child(root, "body")
    if body is None:
        return Document(warnings=["document-no-body"])

    cached: list[str] = []
    fields = _Fields(cached, cross_references=True, styleref=True)
    if kept is not None and segments is None:
        from ..relayout import segments as read_segments

        segments = read_segments(document_xml)
    if kept is not None and segments is not None:
        from ..relayout import body_context

        found = body_context(document_xml)
        listed = _kept_block_list(body, fields, kept, segments, context + found) if found is not None \
            else _block_list(body, None, (), "w:body", fields)
        body_blocks = _following_properties(_join(listed))[0]
    else:
        body_blocks = _story(body, path="w:body", fields=fields)
    paragraphs = [block for block in body_blocks if isinstance(block, Paragraph)]
    document = Document(
        paragraphs=paragraphs,
        sections=_sections(body, paragraphs),
        body=body_blocks,
        cached_fields=cached,
    )

    # Every top-level element that is neither a paragraph nor the section properties is
    # content this reader drops.  Saying so, with a stable code, is the difference
    # between "not implemented" and "silently wrong".
    for node in children(body):
        name = local_name(node.tag)
        if name in ("p", "sectPr", "sdt", "customXml") or name in _MARKERS:
            # A block-level content control's content is read as the body's.
            continue
        if name == "tbl":
            # Laid out (:mod:`docx2svg.table`); what of one is not modelled, the layout
            # says where it stops (``layout-stopped:table``).
            continue
        code = f"body-element-unsupported:{name}"
        if code not in document.warnings:
            document.warnings.append(code)

    if styles_xml is not None:
        run_default, paragraph_default = _doc_defaults(styles_xml)
        document.default_run_properties = run_default
        document.default_paragraph_properties = paragraph_default
        document.styles = parse_styles(styles_xml)

    return document


def parse_footnotes(footnotes_xml: bytes) -> tuple[dict[int, list], dict[str, list], dict[str, int]]:
    """``word/footnotes.xml``: the notes by id, and the separators by type
    (``separator``, ``continuationSeparator``, ``continuationNotice``), each as its
    paragraphs, and the separators' ids."""
    notes, separators, ids = parse_notes(footnotes_xml, "footnote")
    return ({key: [b for b in blocks if isinstance(b, Paragraph)] for key, blocks in notes.items()},
            {key: [b for b in blocks if isinstance(b, Paragraph)] for key, blocks in separators.items()}, ids)


def parse_notes(notes_xml: bytes, kind: str) -> tuple[dict[int, list], dict[str, list], dict[str, int]]:
    """``word/footnotes.xml`` (``kind`` ``footnote``) or ``word/endnotes.xml``
    (``endnote``): the notes by id and the separators by type (``separator``,
    ``continuationSeparator``, ``continuationNotice``), each as its blocks, and the
    separators' ids."""
    notes: dict[int, list] = {}
    separators: dict[str, list] = {}
    ids: dict[str, int] = {}
    for node in children(parse_xml(notes_xml), kind):
        blocks = _story(node, path=f"w:{kind}[@w:id={attr(node, 'id')}]")
        note_type = attr(node, "type")
        note_id = attr_int(node, "id")
        if note_type in ("separator", "continuationSeparator", "continuationNotice"):
            separators[note_type] = blocks
            if note_id is not None:
                ids[note_type] = note_id
        elif note_id is not None:
            notes[note_id] = blocks
    return notes, separators, ids


def image_pixels(data: bytes) -> tuple[int, int] | None:
    """A picture's width and height in pixels (PNG, JPEG, GIF, BMP), or ``None`` for a
    format not read (``ooxml_common.imagemeta`` reads the headers)."""
    from ooxml_common.imagemeta import pixels_and_density

    found = pixels_and_density(data)
    return (found[0], found[1]) if found else None


def parse_package(source: str | bytes | BinaryIO, *, kept: KeptBlocks | None = None,
                  segments: list | None = None) -> Document:
    """Open a ``.docx`` and parse its main document part.  Theme colours (``w:themeColor``,
    ``w:themeFill`` with their tints and shades) read as Word draws them
    (:func:`docx2svg.parse.properties.theme_colours`).

    ``kept`` (round 41, :class:`KeptBlocks`): the body's blocks of the last document parsed
    with it are taken where their bytes, the namespaces around them and the theme's colours
    are the same, the others built; ``segments`` is :func:`docx2svg.relayout.segments` of
    the main part (``word/document.xml``), where the caller has read it.  The document is
    the same with ``kept`` or without."""
    package = OpcPackage.open(source)
    main = main_part(package)
    theme_parts = package.related_parts_of_type(main, REL_THEME)
    settings_parts = package.related_parts_of_type(main, REL_SETTINGS)
    colours = parse_theme_colors(read_part(package, theme_parts[0])) if theme_parts else {}
    mapping = parse_color_map(read_part(package, settings_parts[0])) if settings_parts else {}
    context = b""
    if kept is not None:
        if main.lstrip("/") != _MAIN_PART:
            # The segments a caller reads are of ``word/document.xml``.
            kept, segments = None, None
        else:
            # The body's models read the theme's colours, through the settings' mapping.
            context = repr((sorted(colours.items()), sorted(mapping.items()))).encode("utf-8", "surrogatepass")
    with theme_colours(word_theme_colour(colours, mapping)):
        return _parse_package(package, main, kept=kept, segments=segments, context=context)


_MAIN_PART = "word/document.xml"


#: ``w:themeColor`` / ``w:themeFill`` values -> the theme slot (``a:clrScheme``'s child).
_THEME_SLOTS = {"dark1": "dk1", "light1": "lt1", "dark2": "dk2", "light2": "lt2", "hyperlink": "hlink",
                "followedHyperlink": "folHlink", **{f"accent{k}": f"accent{k}" for k in range(1, 7)}}
#: The four names ``w:clrSchemeMapping`` maps, its attribute for each, and its default.
_MAPPED = {"text1": ("t1", "dark1"), "background1": ("bg1", "light1"),
           "text2": ("t2", "dark2"), "background2": ("bg2", "light2")}


def word_theme_colour(theme_colors: dict, color_map: dict):
    """``(name, tint, shade) -> RRGGBB``: a WordprocessingML theme colour -- ``text1`` and
    the other three through ``w:clrSchemeMapping`` -- with its tint or shade as Word draws
    it (``ooxml_common.themetint.word_theme_tint``, measured); ``None`` for a name the
    theme does not define."""
    from ooxml_common.themetint import word_theme_tint

    def resolve(name: str, tint: str | None, shade: str | None) -> str | None:
        if name in _MAPPED:
            key, default = _MAPPED[name]
            name = color_map.get(key) or default
        base = theme_colors.get(_THEME_SLOTS.get(name, ""))
        return word_theme_tint(base, tint, shade) if base else None

    return resolve


def _parse_package(package: OpcPackage, main: str, *, kept: KeptBlocks | None = None,
                   segments: list | None = None, context: bytes = b"") -> Document:
    styles_parts = package.related_parts_of_type(main, REL_STYLES)
    styles_xml = read_part(package, styles_parts[0]) if styles_parts else None
    document = parse_document(read_part(package, main), styles_xml, kept=kept, segments=segments, context=context)
    if not styles_parts:
        document.warnings.append("styles-part-missing")
    theme_parts = package.related_parts_of_type(main, REL_THEME)
    if theme_parts:
        theme_xml = read_part(package, theme_parts[0])
        document.font_scheme = parse_theme(theme_xml)
        document.theme_colors = parse_theme_colors(theme_xml)
        document.theme_formats = parse_theme_formats(theme_xml)
    numbering_parts = package.related_parts_of_type(main, REL_NUMBERING)
    if numbering_parts:
        numbering_xml = read_part(package, numbering_parts[0])
        numbering = parse_numbering(numbering_xml, document.styles)
        document.numbering = numbering.levels
        document.numbering_lists = numbering.lists
        document.numbering_restarts = numbering.restarts
        for identifier, relationship in parse_picture_bullets(numbering_xml).items():
            image = package.related_part(numbering_parts[0], relationship) if relationship else None
            pixels = image_pixels(read_part(package, image)) if image and package.has_part(image) else None
            document.picture_bullets[identifier] = PictureBullet(relationship, numbering_parts[0], pixels)
    for section in document.sections:
        for _, relationship_id in section.headers + section.footers:
            if relationship_id not in document.stories:
                part = package.related_part(main, relationship_id)
                if part is not None and package.has_part(part):
                    root = parse_xml(read_part(package, part))
                    # Paragraphs and tables, with the fields open across them.
                    document.stories[relationship_id] = _story(root, path=local_name(root.tag).join(("w:", "")),
                                                                fields=_Fields(document.cached_fields, styleref=True))
                    document.story_parts[relationship_id] = part
    footnote_parts = package.related_parts_of_type(main, REL_FOOTNOTES)
    if footnote_parts:
        document.footnotes, document.footnote_separators, document.footnote_separator_ids = parse_footnotes(
            read_part(package, footnote_parts[0]))
    endnote_parts = package.related_parts_of_type(main, REL_ENDNOTES)
    if endnote_parts:
        document.endnotes, document.endnote_separators, document.endnote_separator_ids = parse_notes(
            read_part(package, endnote_parts[0]), "endnote")
    settings_parts = package.related_parts_of_type(main, REL_SETTINGS)
    if settings_parts:
        for name, value in settings_read(read_part(package, settings_parts[0])).items():
            setattr(document, name, value)
    return document


def settings_read(settings_xml: bytes) -> dict:
    """Everything the layout reads of the settings part, by the :class:`Document` field it
    sets -- all of it: a settings part that reads the same lays out the same (round 39: a
    layout from a previous one compares the settings by this, :func:`settings_digest`)."""
    endnote_position, endnote_named = parse_endnote_settings(settings_xml)
    return {"theme_font_lang": parse_settings(settings_xml),
            "compatibility_mode": parse_compatibility_mode(settings_xml),
            "compat_options": parse_compat_options(settings_xml),
            "compat_settings": parse_compat_settings(settings_xml),
            "default_tab_stop": parse_default_tab_stop(settings_xml),
            "even_and_odd_headers": parse_even_and_odd_headers(settings_xml),
            "hyphenation": parse_hyphenation(settings_xml),
            "color_map": parse_color_map(settings_xml),
            "endnote_position": endnote_position,
            "endnote_named": endnote_named,
            "footnote_named": parse_endnote_settings(settings_xml, "footnote")[1]}


def settings_digest(settings_xml: bytes) -> bytes:
    """A digest of :func:`settings_read` -- of what the layout reads of the settings, not of
    their bytes (a revision id, a proofing state, a document variable, the way the XML is
    written read alike) -- the same in every process (sets in order, dicts as written)."""
    import hashlib

    def canonical(value):
        if isinstance(value, (set, frozenset)):
            return ("set", sorted(map(repr, value)))
        if isinstance(value, dict):
            return ("dict", [(repr(k), canonical(v)) for k, v in value.items()])
        return repr(value)

    try:
        read = settings_read(settings_xml)
    except Exception:  # noqa: BLE001 -- unreadable: compared by its bytes, as any part
        return b"bytes:" + hashlib.blake2b(settings_xml, digest_size=16).digest()
    text = repr([(name, canonical(value)) for name, value in read.items()])
    return b"read:" + hashlib.blake2b(text.encode("utf-8", "surrogatepass"), digest_size=16).digest()
