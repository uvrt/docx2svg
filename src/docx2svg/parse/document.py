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
from typing import BinaryIO
from xml.etree.ElementTree import Element

from ..model import (
    Anchor,
    Document,
    Field,
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
from ..opc import REL_ENDNOTES, REL_FOOTNOTES, REL_NUMBERING, REL_SETTINGS, REL_STYLES, REL_THEME, Package
from ..paths import Steps, row_level
from ..xmlutil import attr, attr_bool, attr_int, child, children, local_name, parse_xml
from .drawing import drawing_element, read_anchor, read_graphic_data
from .properties import read_paragraph_properties, read_run_properties, read_table_properties
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
        color=attr(color, "val"),
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
    node = drawing_element(node)
    if node is None:
        return "drawing"
    inline = [child_ for child_ in node if local_name(child_.tag) == "inline"]
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
    """The relationship id of the picture an inline ``w:drawing`` shows, or ``None``."""
    node = drawing_element(node)
    if node is None or child(node, "inline") is None:
        return None
    for element in node.iter():
        if local_name(element.tag) == "blip" and element.get(_R_EMBED):
            return element.get(_R_EMBED)
    return None


def _inline_graphic(node: Element, path: str, fields: "_Fields"):
    """What an inline ``w:drawing`` that is not a picture shows -- a shape, a group, a
    chart or SmartArt (:class:`docx2svg.model.Graphic`) -- or ``None``."""
    node = drawing_element(node)
    inline = child(node, "inline") if node is not None else None
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
#: fields").  :data:`CROSS_REFERENCES` are computed in the body only.
COMPUTED_FIELDS = frozenset({"PAGE", "NUMPAGES", "SECTIONPAGES", "PAGEREF", "REF", "SEQ"})
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

    def __init__(self, cached: list[str] | None = None, *, cross_references: bool = False) -> None:
        self.stack: list[_Open] = []
        #: The keyword of every field whose cached result is drawn.
        self.cached: list[str] = cached if cached is not None else []
        #: Whether :data:`CROSS_REFERENCES` are computed here: the body's own text, not a
        #: text box's.
        self.cross_references = cross_references
        #: The bookmarks met in the paragraph being read: ``(kind, id, name, run index)``.
        self.bookmarks: list | None = None

    def visible(self) -> bool:
        return all(f.state == "result" and not f.computed for f in self.stack)

    def computes(self, instruction: str) -> bool:
        keyword = field_keyword(instruction)
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
    drawings: list[str | None] = []
    anchors: list[tuple] = []
    length = 0
    properties = child(element, "rPr")
    plain = fields.visible()

    def flush() -> None:
        nonlocal text_parts, marks, drawings, anchors, length
        if text_parts or marks:
            out.append(Run(text="".join(text_parts), properties=_run_properties(properties), breaks=tuple(marks),
                           path=path, drawings=tuple(drawings), anchors=tuple(anchors)))
        text_parts, marks, drawings, anchors, length = [], [], [], [], 0

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
            found = _anchor(node, f"{path}/wp:anchor[{len(anchors) + 1}]", fields)
            if found is not None:
                marks.append((length, f"anchor:{len(anchors)}"))
                anchors.append(found)
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
    """The final view of a container's blocks: each paragraph whose mark is deleted
    joined to the paragraph after it (a chain into the last); one with no paragraph
    after it stays, and one with nothing drawn before a table is not drawn at all."""
    out: list = []
    for block in reversed(blocks):
        if isinstance(block, Paragraph) and block.mark_deleted:
            following = out[-1] if out else None
            if isinstance(following, Paragraph):
                out[-1] = _joined(block, following)
                continue
            if isinstance(following, Table) and _empty(block) and block.section is None:
                continue
        out.append(block)
    return out[::-1]


def _block_list(parent: Element, table_style: str | None, conditions: tuple[str, ...], path: str,
                fields: _Fields) -> list:
    out: list = []
    step = _Paths(path)
    for node in children(parent):
        name = local_name(node.tag)
        here = step(node)
        if name == "p":
            out.append(_paragraph(node, table_style, conditions, here, fields))
        elif name == "tbl":
            table = _table(node, here, fields)
            # A table every row of which is deleted is not drawn at all.
            if table.rows:
                out.append(table)
        elif name == "sdt":
            # A block-level content control: its content, as the container's.
            content = child(node, "sdtContent")
            if content is not None:
                out.extend(_block_list(content, table_style, conditions, f"{here}/w:sdtContent[1]", fields))
        elif name == "customXml":
            out.extend(_block_list(node, table_style, conditions, here, fields))
    return out


def _blocks(parent: Element, table_style: str | None = None,
            conditions: tuple[str, ...] = (), path: str = "", fields: _Fields | None = None) -> list:
    """A part's (or a cell's) paragraphs and tables, in order, in the final view of its
    tracked changes (:func:`_join`).  ``fields`` carries the complex fields open across
    them: one per part."""
    fields = fields if fields is not None else _Fields()
    return _join(_block_list(parent, table_style, conditions, path, fields))


def _following_properties(blocks: list, following: ParagraphProperties | None = None) -> tuple[list, object]:
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
    try:
        return _story(parent, path=path, fields=fields)
    finally:
        fields.cross_references = outer


def _table(element: Element, path: str = "", fields: _Fields | None = None) -> Table:
    """A table: its rows and cells, the table style, and what its geometry is read from
    -- ``w:tblGrid`` and every ``w:tblPr``, ``w:trPr`` and ``w:tcPr`` as declared.

    The cell paragraphs carry the table's style id (``""`` when it names none), because
    the table style is a level of their cascade.  A ``w:tc`` inside a row-level
    container (``w:sdt``, ``w:customXml``) is read as the row's cell.
    """
    properties = child(element, "tblPr")
    style = attr(child(properties, "tblStyle"), "val") or ""
    look = _table_look(child(properties, "tblLook"))
    # A deleted row (``w:trPr/w:del``) is not drawn in the final view, whether or not its
    # content is deleted too (measured); its path keeps its place among the rows.
    row_elements = [(i, row) for i, row in enumerate(_row_level(element, "tr"))
                    if child(child(row, "trPr"), "del") is None]
    rows, row_properties, cell_properties, conditions = [], [], [], []
    for k, (i, row) in enumerate(row_elements):
        cells = list(_row_level(row, "tc"))
        row_conditions = tuple(_conditions(look, k, j, len(row_elements), len(cells)) for j in range(len(cells)))
        rows.append(tuple(
            tuple(_blocks(cell, style, row_conditions[j], f"{path}/w:tr[{i + 1}]/w:tc[{j + 1}]", fields))
            for j, cell in enumerate(cells)
        ))
        row_properties.append(read_table_properties(child(row, "trPr")))
        cell_properties.append(tuple(read_table_properties(child(cell, "tcPr")) for cell in cells))
        conditions.append(row_conditions)
    grid = tuple(attr_int(column, "w", 0) or 0 for column in children(child(element, "tblGrid"), "gridCol"))
    return Table(style_id=style or None, rows=tuple(rows), path=path, grid=grid,
                 properties=read_table_properties(properties), row_properties=tuple(row_properties),
                 cell_properties=tuple(cell_properties), look=look, conditions=tuple(conditions))


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


def _conditions(look: dict[str, bool], row: int, column: int, rows: int, columns: int) -> tuple[str, ...]:
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
    return tuple((attr(node, "type") or "default", node.get(_R_ID))
                 for node in children(section_properties, name) if node.get(_R_ID))


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


def parse_document(document_xml: bytes, styles_xml: bytes | None = None) -> Document:
    """Parse ``word/document.xml`` (and optionally ``word/styles.xml``)."""
    root = parse_xml(document_xml)
    body = child(root, "body")
    if body is None:
        return Document(warnings=["document-no-body"])

    cached: list[str] = []
    body_blocks = _story(body, path="w:body", fields=_Fields(cached, cross_references=True))
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
    from ooxml_common.imagemeta import _pixels_and_density

    found = _pixels_and_density(data)
    return (found[0], found[1]) if found else None


def parse_package(source: str | bytes | BinaryIO) -> Document:
    """Open a ``.docx`` and parse its main document part."""
    package = Package.open(source)
    main = package.main_document_part
    styles_parts = package.related(main, REL_STYLES)
    styles_xml = package.read(styles_parts[0]) if styles_parts else None
    document = parse_document(package.read(main), styles_xml)
    if not styles_parts:
        document.warnings.append("styles-part-missing")
    theme_parts = package.related(main, REL_THEME)
    if theme_parts:
        theme_xml = package.read(theme_parts[0])
        document.font_scheme = parse_theme(theme_xml)
        document.theme_colors = parse_theme_colors(theme_xml)
        document.theme_formats = parse_theme_formats(theme_xml)
    numbering_parts = package.related(main, REL_NUMBERING)
    if numbering_parts:
        numbering_xml = package.read(numbering_parts[0])
        numbering = parse_numbering(numbering_xml, document.styles)
        document.numbering = numbering.levels
        document.numbering_lists = numbering.lists
        document.numbering_restarts = numbering.restarts
        for identifier, relationship in parse_picture_bullets(numbering_xml).items():
            image = package.part_by_id(numbering_parts[0], relationship) if relationship else None
            pixels = image_pixels(package.read(image)) if image and package.exists(image) else None
            document.picture_bullets[identifier] = PictureBullet(relationship, numbering_parts[0], pixels)
    for section in document.sections:
        for _, relationship_id in section.headers + section.footers:
            if relationship_id not in document.stories:
                part = package.part_by_id(main, relationship_id)
                if part is not None and package.exists(part):
                    root = parse_xml(package.read(part))
                    # Paragraphs and tables, with the fields open across them.
                    document.stories[relationship_id] = _story(root, path=local_name(root.tag).join(("w:", "")),
                                                                fields=_Fields(document.cached_fields))
                    document.story_parts[relationship_id] = part
    footnote_parts = package.related(main, REL_FOOTNOTES)
    if footnote_parts:
        document.footnotes, document.footnote_separators, document.footnote_separator_ids = parse_footnotes(
            package.read(footnote_parts[0]))
    endnote_parts = package.related(main, REL_ENDNOTES)
    if endnote_parts:
        document.endnotes, document.endnote_separators, document.endnote_separator_ids = parse_notes(
            package.read(endnote_parts[0]), "endnote")
    settings_parts = package.related(main, REL_SETTINGS)
    if settings_parts:
        settings_xml = package.read(settings_parts[0])
        document.theme_font_lang = parse_settings(settings_xml)
        document.compatibility_mode = parse_compatibility_mode(settings_xml)
        document.compat_options = parse_compat_options(settings_xml)
        document.compat_settings = parse_compat_settings(settings_xml)
        document.default_tab_stop = parse_default_tab_stop(settings_xml)
        document.even_and_odd_headers = parse_even_and_odd_headers(settings_xml)
        document.hyphenation = parse_hyphenation(settings_xml)
        document.color_map = parse_color_map(settings_xml)
        document.endnote_position, document.endnote_named = parse_endnote_settings(settings_xml)
        _, document.footnote_named = parse_endnote_settings(settings_xml, "footnote")
    return document
