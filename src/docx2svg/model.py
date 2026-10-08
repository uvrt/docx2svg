"""The source model: what the file says, before any layout.

Deliberately *unresolved*, the same split the sibling ``pptx2svg`` draws between its
``parse`` and ``resolve`` stages.  A :class:`Paragraph` here carries the indent that was
written on it and nothing inherited from a style, a numbering definition or
``w:docDefaults``; a :class:`Run` carries the face its own ``w:rFonts`` named and not the
one it will end up drawn in.  Resolution is a separate stage because the cascade is where
the bugs are, and a stage that has not run yet cannot be blamed for them.

The model stops where layout begins.  There is no ``Line`` and no ``Page`` type: those
are *computed*, not read, and inventing a place for them here would be the first step
towards pretending the file says where a line breaks.  It does not.
"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class PageSize:
    """``w:pgSz``, in twips, as authored.

    Note what this is not: the extent of the page Word exports.  See
    :func:`docx2svg.units.device_page_extent_pt`.
    """

    width_twips: int
    height_twips: int
    #: ``w:orient``; ``None`` when unstated, which means portrait.
    orientation: str | None = None


@dataclass(frozen=True)
class PageMargins:
    """``w:pgMar``, in twips.

    ``header`` and ``footer`` are distances from the *page edge*, not from the text
    column, and they interact with ``top``/``bottom`` in a way that is not a simple
    maximum.  Carried unresolved for that reason.
    """

    top: int = 1440
    right: int = 1440
    bottom: int = 1440
    left: int = 1440
    header: int = 720
    footer: int = 720
    gutter: int = 0


@dataclass(frozen=True)
class Section:
    """``w:sectPr``.

    A section is the unit of page geometry, and a document has at least one.  The
    awkward part of the format is where the element lives: the *last* section's
    ``w:sectPr`` is a child of ``w:body``, and every earlier one is a child of the
    ``w:pPr`` of the last paragraph *in* that section.  So sections are discovered by
    walking paragraphs, not by looking in one place -- see
    :func:`docx2svg.parse.document.parse_document`.
    """

    page_size: PageSize
    margins: PageMargins
    #: ``w:type``: ``nextPage``, ``continuous``, ``evenPage``, ``oddPage``, ``nextColumn``.
    #: ``None`` means ``nextPage``.
    start_type: str | None = None
    #: Index of the first paragraph in this section, into ``Document.paragraphs``.
    first_paragraph: int = 0
    #: One past the last.  A section may legitimately be empty.
    last_paragraph: int = 0
    #: ``w:cols``: the number of text columns, the space between them in twips, and --
    #: when ``w:equalWidth="0"`` -- each column's own ``(width, space after)``.  A
    #: column's width is what a line of this section may fill (ROADMAP.md, Phase 3).
    column_count: int = 1
    column_space: int = 720
    columns: tuple[tuple[int, int], ...] = ()
    #: ``w:cols/@w:sep``: a line drawn between the columns.
    column_separator: bool = False
    #: ``w:headerReference`` / ``w:footerReference``: ``(type, relationship id)`` as
    #: stated here (a type not stated is inherited from the section before), and
    #: ``w:titlePg``.  The parts' paragraphs are ``Document.stories``.
    headers: tuple[tuple[str, str], ...] = ()
    footers: tuple[tuple[str, str], ...] = ()
    title_page: bool = False
    #: ``w:pgNumType``: ``@w:start`` (the number its first page restarts at; ``None``
    #: continues from the section before) and ``@w:fmt`` (how ``PAGE`` shows it).
    page_number_start: int | None = None
    page_number_format: str | None = None
    #: ``w:vAlign``: how the body sits between the margins (``top`` when unstated).  Not
    #: laid out: a section that states ``center``, ``both`` or ``bottom`` is warned of.
    vertical_alignment: str | None = None
    #: ``w:endnotePr``: how the section's endnotes are numbered -- ``w:numFmt``,
    #: ``w:numStart`` and ``w:numRestart`` as stated (``None``: unstated).
    endnote_format: str | None = None
    endnote_start: int | None = None
    endnote_restart: str | None = None
    #: ``w:footnotePr``, as ``endnotePr``, and ``w:pos`` (``pageBottom`` or
    #: ``beneathText``); the document's own (in the settings) is not taken
    #: (:mod:`docx2svg.notes`).
    footnote_format: str | None = None
    footnote_start: int | None = None
    footnote_restart: str | None = None
    footnote_position: str | None = None
    #: ``w:lnNumType``: ``(countBy, start, distance in twips or None for auto, restart)`` as
    #: Word reads it -- ``countBy`` 1, ``start`` 0 and ``restart`` ``newPage`` unstated
    #: (``make_line_number_probe.py``); ``None``: no line numbers.
    line_numbering: tuple | None = None


@dataclass(frozen=True)
class TabStop:
    position_twips: int
    #: ``left``, ``center``, ``right``, ``decimal``, ``bar``, ``num``, ``clear``.
    alignment: str = "left"
    leader: str | None = None


@dataclass(frozen=True)
class Indent:
    """``w:ind``, in twips.

    ``first_line`` and ``hanging`` are mutually exclusive in practice and are carried
    separately rather than collapsed to a signed number, because which one was written
    survives into the numbering cascade and a collapsed value loses it.
    """

    left: int | None = None
    right: int | None = None
    first_line: int | None = None
    hanging: int | None = None


@dataclass(frozen=True)
class Spacing:
    """``w:spacing``, in twips except ``line_rule``.

    ``w:line`` is in twips when ``w:lineRule`` is ``exact`` or ``atLeast``, and in
    *240ths of a line* when it is ``auto`` -- so ``w:line="360"`` is either 18 pt or
    1.5x depending on a sibling attribute.  Storing the rule alongside is the only way
    the value means anything.
    """

    before: int | None = None
    after: int | None = None
    line: int | None = None
    line_rule: str | None = None


@dataclass(frozen=True)
class RunProperties:
    """``w:rPr``, as written.  ``None`` everywhere means "inherit"."""

    #: ``w:rFonts``.  Four slots, chosen per *character* by its script, not per run.
    ascii_font: str | None = None
    high_ansi_font: str | None = None
    east_asian_font: str | None = None
    complex_script_font: str | None = None
    #: ``w:sz`` in half-points, kept in the authored unit.
    size_half_points: int | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: str | None = None
    strike: bool | None = None
    color: str | None = None
    #: ``w:spacing`` inside ``w:rPr`` -- inter-character tracking, in twips.  Not the
    #: paragraph ``w:spacing``; the format reuses the name for an unrelated quantity.
    character_spacing: int | None = None
    #: ``w:vertAlign``: ``superscript``, ``subscript``, ``baseline``.
    vertical_align: str | None = None
    style_id: str | None = None
    #: Every property the element declared, as a flat map for the cascade
    #: (:mod:`docx2svg.parse.properties` lists the keys).  The typed fields above are a
    #: convenience view of a few of them; this is what :mod:`docx2svg.resolve` reads.
    declared: dict = field(default_factory=dict, hash=False)


@dataclass(frozen=True)
class ParagraphProperties:
    """``w:pPr``, as written."""

    style_id: str | None = None
    #: ``w:jc``: ``left``/``start``, ``center``, ``right``/``end``, ``both``, ``distribute``.
    alignment: str | None = None
    indent: Indent | None = None
    spacing: Spacing | None = None
    tab_stops: tuple[TabStop, ...] = ()
    #: ``w:numPr`` -- the list this paragraph belongs to, unresolved.
    numbering_id: int | None = None
    numbering_level: int | None = None
    #: ``w:keepNext``, ``w:keepLines``, ``w:pageBreakBefore``, ``w:widowControl``.
    #: All four are pagination constraints, and all four are why pagination is not a
    #: simple running total of line heights.
    keep_next: bool = False
    keep_lines: bool = False
    page_break_before: bool = False
    widow_control: bool | None = None
    #: The paragraph *mark*'s ``w:rPr``.  The mark is a character: it has a face and a
    #: size, takes part in the height of a line that holds nothing but spaces
    #: (``vertical.line_items``), gives a list label its properties, and Word draws it
    #: (as a space) in its PDF.
    run_properties: RunProperties | None = None
    declared: dict = field(default_factory=dict, hash=False)


@dataclass(frozen=True)
class Field:
    """A field whose result the layout computes (``PAGE``, ``NUMPAGES``, ``SECTIONPAGES``,
    and in the body ``PAGEREF``, ``REF`` and ``SEQ``): its instruction as written
    (``PAGE \\* roman``) and its keyword.

    The run that carries it holds, as its text, the result Word cached in the file -- what
    is drawn where the value cannot be computed (a page count past where the layout
    stops, a number format that depends on the application's language).  Its properties
    are the ones Word formats the result in (``make_field_probe.py``): the instruction's
    first run's, or with ``\\* MERGEFORMAT`` the cached result's first run's.
    """

    instruction: str
    keyword: str
    #: The result Word cached, as ``(text, RunProperties)`` per run: what ``\\* MERGEFORMAT``
    #: carries over to a recomputed ``REF`` word by word (:mod:`docx2svg.fields`).
    cached: tuple = field(default=(), compare=False)


@dataclass(frozen=True)
class AnchorPosition:
    """``wp:positionH`` or ``wp:positionV``: what the position is measured from
    (``relativeFrom``) and either an offset from it (``wp:posOffset``, EMU) or an
    alignment in it (``wp:align``)."""

    relative: str | None = None
    offset: int | None = None
    align: str | None = None


@dataclass(frozen=True)
class Graphic:
    """What a drawing shows, as the file says it (DrawingML, ``a:graphic``), unresolved.

    ``kind`` is ``picture`` (``pic:pic``; ``relationship`` names its image), ``shape``
    (``wps:wsp``), ``group`` (``wpg:wgp`` / ``wpg:grpSp``; ``children``), ``chart``
    (``c:chart``; ``relationship`` names the chart part), ``diagram`` (SmartArt's
    ``dgm:relIds``; ``relationship`` names its data model) or ``other`` (a canvas, an
    embedded object...: ``uri`` says which).  Geometry, colours and text are
    kept as read; :mod:`docx2svg.drawing` draws them.
    """

    kind: str
    #: ``a:xfrm``: offset and extent (EMU, in the parent group's child coordinates; the
    #: anchor's own for the top-level graphic), rotation (60000ths of a degree), flips.
    offset: tuple[int, int] = (0, 0)
    extent: tuple[int, int] = (0, 0)
    rotation: int = 0
    flip_h: bool = False
    flip_v: bool = False
    relationship: str | None = None
    #: The geometry as ``ooxml_common.drawingml.read.parse_geometry_spec`` reads it:
    #: ``("preset", name, {adjustment: formula})``, ``("custom", spec)``, or ``None``.
    geometry: tuple | None = field(default=None, compare=False)
    #: ``spPr``'s fill (a group's: ``grpSpPr``'s), outline and effects, and the shape's
    #: ``wps:style`` theme references, as ``ooxml_common.drawingml.read`` reads them
    #: (``ooxml_common.drawingml.source`` types, unresolved); ``None`` where not stated.
    fill: object = field(default=None, compare=False)
    line: object = field(default=None, compare=False)
    effects: object = field(default=None, compare=False)
    style: object = field(default=None, compare=False)
    #: A text box's content (``w:txbxContent``): paragraphs and tables.
    text: tuple = ()
    #: ``wps:bodyPr``: ``lIns`` ``tIns`` ``rIns`` ``bIns`` (EMU), ``anchor``, ``vert``,
    #: ``wrap``, ``rot``, and ``fit``, the auto-fit element's name (``noAutofit``,
    #: ``spAutoFit``, ``normAutofit``).
    body: dict = field(default_factory=dict, hash=False, compare=False)
    #: A custom geometry's text rectangle (``a:custGeom/a:rect``: four guide names or
    #: literals); a preset's is the specification's.  ``None`` where there is none.
    text_rect: tuple | None = field(default=None, compare=False)
    children: tuple["Graphic", ...] = ()
    #: A group's ``a:chOff`` and ``a:chExt``.
    child_offset: tuple[int, int] = (0, 0)
    child_extent: tuple[int, int] = (0, 0)
    #: What is there and not drawn: effects, a fill or outline kind not read, a
    #: ``graphicData`` of another kind.
    unsupported: tuple[str, ...] = ()
    uri: str | None = None
    #: The structural path of the drawing's content, for ``data-docx-path``.
    path: str = field(default="", compare=False)


@dataclass(frozen=True)
class Anchor:
    """A floating drawing, ``wp:anchor`` (ECMA-376 20.4.2.3), as written.

    Where it goes (``h``, ``v``, or ``simple`` when ``simplePos="1"``), how large it is
    (``extent``, EMU, and ``effect``, the effect extent ``l t r b``), how text goes around
    it (``wrap``: ``wrapNone``, ``wrapSquare``, ``wrapTight``, ``wrapThrough``,
    ``wrapTopAndBottom``, or ``None``; ``wrap_text`` and ``distance`` -- ``distT``,
    ``distB``, ``distL``, ``distR``), how it stacks (``behind`` the text, and
    ``relative_height``), and what it shows (``graphic``).
    """

    extent: tuple[int, int]
    effect: tuple[int, int, int, int] = (0, 0, 0, 0)
    h: AnchorPosition = AnchorPosition()
    v: AnchorPosition = AnchorPosition()
    simple: tuple[int, int] | None = None
    wrap: str | None = None
    wrap_text: str | None = None
    distance: tuple[int, int, int, int] = (0, 0, 0, 0)
    behind: bool = False
    relative_height: int = 0
    allow_overlap: bool = True
    layout_in_cell: bool = True
    locked: bool = False
    graphic: Graphic | None = None
    #: ``wp:docPr/@id``.
    identifier: str | None = None
    #: ``wp:wrapPolygon`` of a ``wrapTight`` or ``wrapThrough``: its points, in 21,600ths of
    #: the extent on each axis, from ``wp:start`` through every ``wp:lineTo``.
    polygon: tuple[tuple[int, int], ...] | None = field(default=None, compare=False)

    @property
    def moves_text(self) -> bool:
        """Whether text is laid out around it: every wrap but ``wrapNone`` (behind or in
        front of the text)."""
        return self.wrap not in (None, "wrapNone")


@dataclass(frozen=True)
class Run:
    text: str
    properties: RunProperties | None = None
    #: ``w:tab``, ``w:br``, ``w:cr`` and friends, in document order, interleaved with
    #: the text by index.  A tab is not whitespace -- it is a jump to a stop -- so it
    #: cannot be folded into :attr:`text` without losing the distinction.
    breaks: tuple[tuple[int, str], ...] = ()
    #: Where the run is, as an XPath-like structural path from its paragraph
    #: (``w:hyperlink[1]/w:r[2]``): what an SVG element drawn from it names in
    #: ``data-docx-path`` (Phase 5).  Not part of the run's value.
    path: str = field(default="", compare=False)
    #: For each inline drawing's mark (``drawing:...``) in :attr:`breaks`, in order, the
    #: relationship id of the picture it shows (``a:blip/@r:embed``), the shape or group it
    #: shows (:class:`Graphic`), or ``None`` for anything else.  Not part of the run's value.
    drawings: tuple = field(default=(), compare=False)
    #: Each floating drawing (:class:`Anchor`) in the run; its place in the text is the
    #: mark ``anchor:<index>`` in :attr:`breaks`.  Not part of the run's value.
    anchors: tuple["Anchor", ...] = field(default=(), compare=False)
    #: A field the layout computes (:class:`Field`); ``None`` for text, and for a field
    #: whose result is drawn as Word cached it (its result's runs are then ordinary runs).
    #: A note's number (:mod:`docx2svg.notes`): its glyphs advance by whole device px
    #: and are drawn on whole device px (:func:`docx2svg.linebreak.pieces`).
    note_number: bool = field(default=False, compare=False)
    field: Field | None = field(default=None, compare=False)


@dataclass(frozen=True)
class Paragraph:
    runs: tuple[Run, ...] = ()
    properties: ParagraphProperties | None = None
    #: ``w:tblStyle`` of the innermost table this paragraph sits in, ``""`` for a table
    #: that names none, ``None`` outside tables.  A table style is a level of the
    #: cascade for the paragraphs in its cells, and nothing else records the link.
    table_style_id: str | None = None
    #: The table style's conditional formats (``w:tblStylePr/@w:type``) that apply to
    #: the cell this paragraph is in, in the order they are applied -- decided from the
    #: cell's position and the table's ``w:tblLook``.  Empty outside tables.
    table_conditions: tuple[str, ...] = ()
    #: The paragraph's structural path in its part (``w:body/w:p[3]``,
    #: ``w:body/w:tbl[1]/w:tr[2]/w:tc[1]/w:p[1]``), 1-based among same-named siblings, and
    #: ``w14:paraId`` when Word wrote one.  An id from the file is not unique (Word copies
    #: it with the paragraph), so an SVG element names both (Phase 5).  Neither is part of
    #: the paragraph's value.
    path: str = field(default="", compare=False)
    para_id: str | None = field(default=None, compare=False)
    #: A paragraph of the notes the layout places in the flow (:mod:`docx2svg.notes`):
    #: ``endnote`` for a note's, ``separator`` or ``continuationSeparator`` for a
    #: separator's; ``None`` for the document's own (a footnote's paragraphs are not
    #: placed in the flow: :mod:`docx2svg.notes`).
    note: str | None = field(default=None, compare=False)
    #: The final view of tracked changes (ROADMAP.md, "Revisions -- measured"): a
    #: paragraph whose mark is deleted is joined to the paragraph after it, and these are
    #: the indexes into :attr:`runs` where each joined paragraph's runs begin;
    #: ``mark_deleted`` says the mark is deleted and there was no paragraph to join.
    joins: tuple[int, ...] = field(default=(), compare=False)
    mark_deleted: bool = field(default=False, compare=False)
    #: The ``w:sectPr`` this paragraph ends a section with (an ``Element``), or ``None``.
    section: object = field(default=None, compare=False, repr=False)
    #: The bookmarks that start or end in the paragraph, in document order:
    #: ``("start" | "end", w:id, w:name or None, run index)`` -- the index into :attr:`runs`
    #: of the run the mark stands before (``len(runs)`` after the last).  What ``PAGEREF``
    #: and ``REF`` find their target by (:mod:`docx2svg.fields`).
    bookmarks: tuple = field(default=(), compare=False)

    @property
    def text(self) -> str:
        return "".join(run.text for run in self.runs)


@dataclass(frozen=True)
class Table:
    """``w:tbl``: rows of cells of blocks, and what their geometry is read from.

    ``rows[r][c]`` is cell ``c`` of row ``r`` (``w:tc`` in document order, so a cell
    spanning grid columns is one entry); ``row_properties[r]`` and
    ``cell_properties[r][c]`` are what their ``w:trPr`` and ``w:tcPr`` declare, and
    ``properties`` what the ``w:tblPr`` does (:func:`docx2svg.parse.properties.
    read_table_properties`), all unresolved: the table style's are resolved with them in
    :mod:`docx2svg.table`.
    """

    style_id: str | None = None
    rows: tuple[tuple[tuple["Paragraph | Table", ...], ...], ...] = ()
    #: Its structural path (``w:body/w:tbl[1]``), as :attr:`Paragraph.path`.
    path: str = field(default="", compare=False)
    #: ``w:tblGrid``: every grid column's width, twips.
    grid: tuple[int, ...] = field(default=(), compare=False)
    properties: dict = field(default_factory=dict, compare=False)
    row_properties: tuple[dict, ...] = field(default=(), compare=False)
    cell_properties: tuple[tuple[dict, ...], ...] = field(default=(), compare=False)
    #: ``w:tblLook``, as :func:`docx2svg.parse.document._table_look` reads it.
    look: dict = field(default_factory=dict, compare=False)
    #: The table style's conditional formats that apply to each cell, in order.
    conditions: tuple[tuple[tuple[str, ...], ...], ...] = field(default=(), compare=False)


@dataclass(frozen=True)
class Style:
    """One ``w:style``, unresolved: what *it* declares, not what it inherits."""

    style_id: str
    #: ``paragraph``, ``character``, ``table`` or ``numbering``.
    kind: str
    name: str | None = None
    based_on: str | None = None
    #: ``w:next`` and ``w:link`` are editing behaviour -- the style of the paragraph
    #: typed after this one, and the paired character style.  Neither changes how an
    #: existing document renders; a run whose ``w:rStyle`` names a *paragraph* style is
    #: not redirected to its linked character style (measured: case ``t22``).
    next_style: str | None = None
    link: str | None = None
    #: ``w:default="1"``: the style applied where none is named.
    default: bool = False
    paragraph: dict = field(default_factory=dict, hash=False)
    run: dict = field(default_factory=dict, hash=False)
    #: Table styles only: ``w:tblStylePr`` type -> (declared pPr, declared rPr, tblPr,
    #: trPr, tcPr).
    conditional: dict = field(default_factory=dict, hash=False)
    #: Table styles only: what the style's own ``w:tblPr``, ``w:trPr`` and ``w:tcPr``
    #: declare (:func:`docx2svg.parse.properties.read_table_properties`).
    table: dict = field(default_factory=dict, hash=False)
    table_row: dict = field(default_factory=dict, hash=False)
    table_cell: dict = field(default_factory=dict, hash=False)


@dataclass
class StyleSheet:
    """``word/styles.xml``: ``w:docDefaults`` and the styles, by id."""

    run_defaults: dict = field(default_factory=dict)
    paragraph_defaults: dict = field(default_factory=dict)
    styles: dict[str, Style] = field(default_factory=dict)

    def default(self, kind: str) -> Style | None:
        for style in self.styles.values():
            if style.kind == kind and style.default:
                return style
        return None


@dataclass(frozen=True)
class ThemeFonts:
    """``a:majorFont`` or ``a:minorFont``: three typefaces and the per-script list."""

    latin: str = ""
    east_asian: str = ""
    complex_script: str = ""
    #: ``a:font script="Hebr" typeface="..."`` -- consulted through ``w:themeFontLang``.
    scripts: dict = field(default_factory=dict, hash=False)


@dataclass(frozen=True)
class FontScheme:
    major: ThemeFonts
    minor: ThemeFonts


@dataclass(frozen=True)
class NumberingLevel:
    level: int
    format: str | None = None
    text: str | None = None
    #: The level's ``w:pPr`` -- a cascade level for the paragraph (its indent).
    paragraph: dict = field(default_factory=dict, hash=False)
    #: The level's ``w:rPr`` -- applies to the list *label* only, not the text.
    run: dict = field(default_factory=dict, hash=False)
    #: ``w:suff``: what follows the label -- ``tab`` (the default), ``space`` or
    #: ``nothing``.
    suffix: str | None = None
    #: ``w:start`` -- for a ``w:num``'s level, its ``w:lvlOverride``'s start where it
    #: states one: the overriding ``w:lvl``'s ``w:start``, else ``w:startOverride``.
    start: int | None = None
    #: ``w:lvlRestart``: the level (1-based) whose items restart this one; ``0`` never;
    #: ``None`` unstated -- any shallower level's.
    restart: int | None = None
    #: ``w:isLgl``: this level's label shows every level's number in decimal.
    legal: bool = False
    #: ``w:lvlJc``: ``left`` (``None``), ``center`` or ``right`` -- how the label stands
    #: at its start: from it, centred on it, or ending at it.
    alignment: str | None = None
    #: ``w:lvlPicBulletId``: the label is the picture of that ``w:numPicBullet``
    #: (:attr:`Document.picture_bullets`), not its text.
    picture_bullet: int | None = None


@dataclass(frozen=True)
class PictureBullet:
    """A ``w:numPicBullet``: the picture a list level may show as its label.

    ``relationship`` is the picture's relationship id in ``part`` (the numbering part:
    ``v:imagedata/@r:id`` of a VML ``w:pict``, ``a:blip/@r:embed`` of a ``w:drawing``);
    ``pixels`` the picture's width and height in pixels, which is what sizes it
    (:func:`docx2svg.linebreak.picture_bullet_units`), or ``None`` where its format is
    not read.
    """

    relationship: str | None
    part: str | None = None
    pixels: tuple[int, int] | None = None


@dataclass(frozen=True)
class Hyphenation:
    """``w:autoHyphenation`` is on: what else ``settings.xml`` says about it."""

    #: ``w:hyphenationZone`` in twips; ``None`` when unstated.
    zone: int | None = None
    #: ``w:consecutiveHyphenLimit``; ``0`` (or unstated) is no limit.
    limit: int = 0
    #: ``w:doNotHyphenateCaps``: words in capitals are not hyphenated.
    no_caps: bool = False


@dataclass
class Document:
    paragraphs: list[Paragraph] = field(default_factory=list)
    sections: list[Section] = field(default_factory=list)
    #: ``w:docDefaults`` from ``word/styles.xml``, unmerged.
    default_run_properties: RunProperties | None = None
    default_paragraph_properties: ParagraphProperties | None = None
    #: Paragraphs and tables in document order.  ``paragraphs`` above holds only the
    #: top-level paragraphs, which is what section discovery indexes into.
    body: list = field(default_factory=list)
    #: ``word/styles.xml``; ``None`` when the package has no styles part.
    styles: StyleSheet | None = None
    #: The theme's font scheme; ``None`` when the package has no theme part.
    font_scheme: FontScheme | None = None
    #: The theme's colour scheme (``a:clrScheme``): ``dk1``, ``lt1``, ``accent1``... ->
    #: ``RRGGBB``; and ``w:settings/w:clrSchemeMapping`` (``bg1`` -> ``light1``...), empty
    #: when the settings state none.  What a DrawingML ``a:schemeClr`` means.
    theme_colors: dict = field(default_factory=dict)
    color_map: dict = field(default_factory=dict)
    #: The theme's format scheme (``a:fmtScheme``, an ``ooxml_common.drawingml.source.
    #: SourceFormatScheme``), what a shape's ``wps:style`` indexes; ``None`` without a theme.
    theme_formats: object = None
    #: ``w:numbering``: numId -> level -> definition, overrides applied.
    numbering: dict = field(default_factory=dict)
    #: numId -> the list it counts in: the ``w:abstractNumId`` of its abstract definition
    #: (the first of those sharing its ``w:nsid``).  Instances of one list share counts.
    numbering_lists: dict = field(default_factory=dict)
    #: ``(numId, ilvl)`` of every level a ``w:startOverride`` restarts at the instance's
    #: first item there.
    numbering_restarts: frozenset = frozenset()
    #: Paragraph path -> list label, where a story's labels are counted with others'
    #: (text boxes count as one story): consulted before the story's own count.
    list_labels: dict = field(default_factory=dict, compare=False)
    #: ``w:numPicBullet``: id -> :class:`PictureBullet`.
    picture_bullets: dict = field(default_factory=dict)
    #: ``w:settings/w:themeFontLang`` (``val`` / ``eastAsia`` / ``bidi``).  Decides which
    #: of the theme's per-script fonts an East Asian or complex-script theme reference
    #: means; without it Word ignores the theme for those slots (measured).
    theme_font_lang: dict = field(default_factory=dict)
    #: ``w:settings/w:compat``'s ``compatibilityMode`` (12 = Word 2007, 14 = 2010, 15 =
    #: 2013 and later); ``None`` when unstated -- no settings part, no ``w:compat``, or
    #: one without the setting.  Unresolved: what "unstated" means is measured per rule
    #: (so far it always behaves as 12 and 14 do; ROADMAP.md, "The page top").
    compatibility_mode: int | None = None
    #: ``w:settings/w:defaultTabStop`` in twips; ``None`` when unstated.  The interval of
    #: the tab stops a paragraph has beyond its own.
    default_tab_stop: int | None = None
    #: The legacy ``w:compat`` options that are on, by element name
    #: (``doNotUseHTMLParagraphAutoSpacing``...).
    compat_options: frozenset = frozenset()
    #: ``w:compat/w:compatSetting`` other than the mode, by name, their ``w:val`` as
    #: written (``overrideTableStyleFontSizeAndJustification`` -> ``"1"``).
    compat_settings: dict = field(default_factory=dict)
    #: ``word/footnotes.xml``: note id -> its paragraphs; and the separators by type
    #: (``separator``, ``continuationSeparator``) -> theirs.  A body run records a
    #: reference as the mark ``footnoteReference:<id>`` (:attr:`Run.breaks`).
    footnotes: dict = field(default_factory=dict)
    footnote_separators: dict = field(default_factory=dict)
    #: The footnote separators' ids by type, and the notes the settings' ``w:footnotePr``
    #: names (``w:footnote/@w:id``), as ``endnote_separator_ids`` and ``endnote_named``.
    footnote_separator_ids: dict = field(default_factory=dict)
    footnote_named: frozenset = frozenset()
    #: ``word/endnotes.xml``, as ``footnotes``: note id -> its blocks, and the separators
    #: by type; a body run records a reference as the mark ``endnoteReference:<id>``, a
    #: note its own number (``w:endnoteRef``) as ``endnoteRef``.
    endnotes: dict = field(default_factory=dict)
    endnote_separators: dict = field(default_factory=dict)
    #: The separator notes' ids by type (``w:endnote/@w:type``).
    endnote_separator_ids: dict = field(default_factory=dict)
    #: ``w:settings/w:endnotePr``: ``w:pos`` (``docEnd`` when unstated, or ``sectEnd``),
    #: and the notes it names (``w:endnote/@w:id``) -- the separators Word takes from the
    #: endnotes part; any it does not name, it replaces with its own.
    endnote_position: str | None = None
    endnote_named: frozenset = frozenset()
    #: Header and footer parts: relationship id (from ``word/document.xml``) -> their
    #: blocks (paragraphs and tables), and -> the part's name (``story_parts``).
    stories: dict = field(default_factory=dict)
    story_parts: dict = field(default_factory=dict)
    #: ``w:settings/w:evenAndOddHeaders``: even-numbered pages take the ``even`` stories.
    even_and_odd_headers: bool = False
    #: ``w:settings/w:autoHyphenation`` and the settings that shape it (ROADMAP.md, 3.9):
    #: :class:`Hyphenation`, or ``None`` when Word does not hyphenate automatically.
    hyphenation: "Hyphenation | None" = None
    #: The keyword of every field drawn as Word cached it (not computed), in every part,
    #: in order met: what the renderer warns of where Word computes it again.
    cached_fields: list[str] = field(default_factory=list)
    #: Everything the reader met and did not model, as stable code strings, so a caller
    #: can fail a build on the ones it cares about.  The sibling project's
    #: ``ConvertOptions.warnings`` channel, brought over on day one rather than retrofitted.
    warnings: list[str] = field(default_factory=list)
