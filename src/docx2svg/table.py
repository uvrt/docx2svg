"""Tables: where every cell is, what its text may fill, and how tall its row is.

ROADMAP.md, "Tables -- measured", holds the probes behind every rule here, with their
scores and the hypotheses they refuted.  The model, in the order Word needs it:

* **Resolution** (:func:`resolve_table`): the table style's ``w:basedOn`` chain (the
  default table style when the table names none), then the table's own ``w:tblPr``; a
  row's and a cell's properties likewise, with the style's conditional formats for the
  cell (``w:tblStylePr``) between the style and the cell's own ``w:tcPr``.
* **The grid** (:func:`grid_lines`): the grid columns' widths as authored
  (``w:tblGrid``), placed by the table's indent and alignment.
* **A cell's text** starts at its left grid line plus the larger of its left margin and
  the inner half of its left border, **snapped to a whole device pixel**; it may fill up
  to its right grid line less the larger of its right margin and the inner half of its
  right border, measured from that snapped start (:func:`cell_box`).
* **A row** is as tall as its tallest cell -- margins and paragraphs -- and at least
  (``atLeast``) or exactly (``exact``) its ``w:trHeight``; each horizontal border adds its
  width once, between the rows it parts.

Everything horizontal is computed in twips (every input is one) and converted once;
everything vertical in exact device px, as the body's stack is.  Standard library only.
"""

from __future__ import annotations

import dataclasses

import math
from dataclasses import dataclass, field
from fractions import Fraction

from .linebreak import UNITS_PER_TWIP, twips_to_units
from .model import Document, Table
from .resolve.cascade import style_chain
from .vertical import LAYOUT_UNIT_PX, LINES_TALL, quantise, round_half_up, twips_to_px

#: The cell margin Word uses on a side nothing declares, twips: left and right as below,
#: top and bottom none.  Measured (``make_table_geometry_probe.py``, ``margins none``):
#: with no ``w:tblCellMar`` anywhere -- no table style either -- a cell's text starts
#: 2 px in and its line is 20 twips shorter than the grid, with borders or without.  (A
#: document's own default table style, Word's "Normal Table", states 108.)
DEFAULT_SIDE_MARGIN = 10

#: Border styles that are not drawn (and take no room).
NO_BORDER = frozenset({None, "nil", "none"})

#: What of a table the model lays out, stage by stage (ROADMAP.md, "Tables --
#: measured"): ``grid`` (widths, indent, alignment, margins, borders), ``content`` (the
#: cell text's size and spacing rules), ``rows`` (``w:trHeight``, ``w:vAlign``),
#: ``merges`` (``w:gridSpan``, ``w:vMerge``, ``w:gridBefore`` / ``w:gridAfter``),
#: ``borders`` (cell borders against their neighbours', shading, banding) and ``pages``
#: (rows across pages).  A table that needs a stage not listed is an obstacle
#: (:class:`Unsupported`), so each stage was switched on only with its probe.
STAGES = frozenset({"grid", "content", "rows", "merges", "borders", "pages"})

SIDES = ("top", "left", "bottom", "right")


def drawn(border) -> bool:
    return bool(border) and border[0] not in NO_BORDER


def border_twips(border) -> int:
    """A border line's width in whole twips, truncated (``w:sz`` is eighths of a point:
    ``w:sz`` 5 is 12 twips, 27 is 67), as a paragraph border's is
    (:func:`docx2svg.vertical.border_px`).  Measured (``make_table_geometry_probe.py``):
    with the exact 12.5 twips a ``w:sz`` 5 table's middle cell starts a pixel right of
    where Word draws it in mode 15, and the drawn widths (``w:sz`` 27: 13 px, not 14)
    are the truncated twips'."""
    if not drawn(border):
        return 0
    return border[1] * 20 // 8 * LINES_TALL.get(border[0], 1)


def compatible_mode(document: Document) -> bool:
    """Whether the document lays tables out as Word 2013 and later do (mode 15): the table
    starts at its indent with its left border's outer edge, rather than with its first
    cell's text (ROADMAP.md, "Tables -- measured", the grid)."""
    return document.compatibility_mode is not None and document.compatibility_mode >= 15


# -- resolution --------------------------------------------------------------------------


def table_style_id(document: Document, table: Table) -> str | None:
    """The table style that applies: the named one, or the default table style when the
    table names none or one that is not a table style (as the cascade does for the
    paragraphs in its cells)."""
    sheet = document.styles
    if sheet is None:
        return None
    if table.style_id:
        style = sheet.styles.get(table.style_id)
        if style is not None and style.kind == "table":
            return table.style_id
    default = sheet.default("table")
    return default.style_id if default else None


@dataclass
class Cell:
    """One ``w:tc`` of one row, resolved and placed on the grid."""

    row: int
    #: Its index among its row's cells (``Table.rows[row][index]``).
    index: int
    #: The first grid column it covers, and how many (``w:gridSpan``).
    column: int
    span: int
    blocks: tuple
    #: Its resolved ``w:tcPr`` (style, conditional formats, then its own).
    properties: dict
    #: ``restart``, ``continue`` or ``None`` (``w:vMerge``).
    merge: str | None = None
    path: str = ""
    #: Margins, twips: ``top``, ``left``, ``bottom``, ``right``.
    margins: dict = field(default_factory=dict)
    #: Each side's border before the conflicts between neighbours are settled, and
    #: whether the cell itself states it (``w:tcBorders``) rather than the table.
    borders: dict = field(default_factory=dict)
    own: dict = field(default_factory=dict)
    #: The border Word draws on its left and right edges once the neighbours' are
    #: settled (:func:`edge_winner`), and which side declared it: ``"left"`` (the cell on
    #: the edge's left, or the table) or ``"right"``.
    left_edge: tuple = (None, "left")
    right_edge: tuple = (None, "left")


@dataclass
class Row:
    index: int
    properties: dict
    cells: list[Cell]
    grid_before: int = 0
    grid_after: int = 0


@dataclass
class ResolvedTable:
    table: Table
    properties: dict
    rows: list[Row]
    #: Whether any level of the table's cascade states ``w:tblInd``.
    indent_stated: bool = False
    #: A table nested in a cell: placed as a mode-15 table is, in every mode
    #: (:func:`grid_origin`, :func:`percent_widths`).
    nested: bool = False
    #: Each grid column's narrowest content, twips (:func:`column_minimums`): what a table
    #: in percent over cells in ``dxa`` shares by (:func:`percent_widths`).
    minimums: tuple | None = None
    #: ``(row, cell index) -> (narrowest, widest)`` content of every cell, twips
    #: (:func:`cell_extent`): what an autofit table is sized by (:func:`autofit_widths`).
    extents: dict | None = None

    @property
    def columns(self) -> int:
        return len(self.table.grid)


def _merge(chain, attribute: str, direct: dict, conditions: tuple[str, ...] = (), index: int | None = None) -> dict:
    """``attribute`` of each style in ``chain``, root first, each followed by its
    conditional formats' part ``index`` for ``conditions`` (the banding ones only where
    the style turns banding on: :func:`docx2svg.resolve.cascade.banded`); then
    ``direct``."""
    from .resolve.cascade import banded

    conditions = banded(chain, conditions)
    out: dict = {}
    for style in chain:
        out.update(getattr(style, attribute))
        if index is not None:
            for condition in conditions:
                if condition in style.conditional:
                    out.update(style.conditional[condition][index])
    out.update(direct)
    return out


def _margin(value) -> int:
    """A cell margin in twips: ``dxa`` as written, anything else none."""
    if value is None:
        return 0
    width, kind = value
    return width if kind in ("dxa", None) else 0


def resolve_table(document: Document, table: Table) -> ResolvedTable:
    style_id = table_style_id(document, table)
    chain = style_chain(document.styles, style_id, "table")
    properties = _merge(chain, "table", table.properties)
    indent_stated = "tblInd" in table.properties or any("tblInd" in style.table for style in chain)
    rows: list[Row] = []
    count = len(table.rows)
    for r, cells in enumerate(table.rows):
        row_direct = table.row_properties[r] if r < len(table.row_properties) else {}
        row_conditions = table.conditions[r][0] if r < len(table.conditions) and table.conditions[r] else ()
        row_properties = _merge(chain, "table_row", row_direct, row_conditions, 3)
        grid_before = row_properties.get("gridBefore", 0) or 0
        grid_after = row_properties.get("gridAfter", 0) or 0
        column = grid_before
        placed: list[Cell] = []
        for c, blocks in enumerate(cells):
            direct = table.cell_properties[r][c] if r < len(table.cell_properties) and c < len(
                table.cell_properties[r]) else {}
            conditions = table.conditions[r][c] if r < len(table.conditions) and c < len(table.conditions[r]) else ()
            props = _merge(chain, "table_cell", direct, conditions, 4)
            span = max(1, props.get("gridSpan", 1) or 1)
            cell = Cell(r, c, column, span, tuple(blocks), props, props.get("vMerge"),
                        f"{table.path}/w:tr[{r + 1}]/w:tc[{c + 1}]")
            cell.margins = {side: _margin(props.get(f"tcMar.{side}", properties.get(f"tblCellMar.{side}")))
                            if (f"tcMar.{side}" in props or f"tblCellMar.{side}" in properties)
                            else (DEFAULT_SIDE_MARGIN if side in ("left", "right") else 0)
                            for side in SIDES}
            placed.append(cell)
            column += span
        rows.append(Row(r, row_properties, placed, grid_before, grid_after))
    resolved = ResolvedTable(table, properties, rows, indent_stated)
    columns = resolved.columns or max((cell.column + cell.span for row in rows for cell in row.cells), default=0)
    for row in rows:
        for cell in row.cells:
            cell.borders, cell.own = _cell_borders(properties, cell, row, count, columns)
        for k, cell in enumerate(row.cells):
            left = row.cells[k - 1] if k else None
            if left is None:
                cell.left_edge = (cell.borders.get("left"), "right" if cell.own.get("left") else "left")
            else:
                cell.left_edge = edge_winner((left.borders.get("right"), left.own.get("right")),
                                             (cell.borders.get("left"), cell.own.get("left")))
                left.right_edge = cell.left_edge
        if row.cells:
            last = row.cells[-1]
            last.right_edge = (last.borders.get("right"), "left")
    return resolved


def edge_winner(first, second) -> tuple:
    """The border Word draws where two cells' borders meet on one edge: ``first`` is the
    left (or upper) cell's ``(border, stated by the cell)``, ``second`` the right (or
    lower) one's.  Returns ``(border, "left" | "right")`` -- which side declared it, the
    table counting as the left.

    Measured by ``make_table_border_probe.py`` (colours per border, widths 2 to 24,
    ``single`` / ``double``, ``nil``): **the wider wins** -- a ``double`` line counting
    three widths -- whichever cell or level states it; ``nil`` or no border loses to any
    border (a ``nil`` cell border under a table's ``insideV`` of 24 leaves the 24); at
    equal widths the left or upper cell's.
    """
    a, a_own = first
    b, b_own = second
    if not drawn(b):
        return a, "left"
    if not drawn(a):
        return b, "right" if b_own else "left"
    if border_twips(b) > border_twips(a):
        return b, "right" if b_own else "left"
    return a, "left"


def inner_part(border) -> int:
    """The part of a border on a cell's side of its grid line, whole twips: half its
    width, the smaller half of an odd width (``make_table_border_probe.py``: a border of
    ``w:sz`` 18, 45 twips, between two cells with no margins leaves each 22)."""
    return border_twips(border) // 2


def outer_part(border) -> int:
    """The part of a table's outer border beyond its grid line: the larger half of an odd
    width (``make_table_geometry_probe.py``: in mode 15 a ``w:sz`` 2 left border, 5
    twips, puts the first grid line 3 twips in)."""
    return -(-border_twips(border) // 2)


def _cell_borders(table_properties: dict, cell: Cell, row: Row, rows: int, columns: int) -> tuple[dict, dict]:
    """Each side's border as the cell and the table state it: the cell's own (its
    ``w:tcBorders``, conditional formats included) over the table's -- the outer border
    on the table's edge, ``insideH`` / ``insideV`` inside; and which the cell states."""
    out = {}
    own_sides = {}
    edge = {"top": cell.row == 0, "bottom": cell.row == rows - 1,
            "left": cell.column == 0, "right": cell.column + cell.span >= columns}
    inside = {"top": "insideH", "bottom": "insideH", "left": "insideV", "right": "insideV"}
    for side in SIDES:
        own = cell.properties.get(f"tcBorders.{side}")
        own_sides[side] = own is not None
        if own is not None:
            out[side] = own
            continue
        key = side if edge[side] else inside[side]
        out[side] = table_properties.get(f"tblBorders.{key}")
    return out, own_sides


# -- the grid ----------------------------------------------------------------------------


def _edge_border(resolved: ResolvedTable, row: Row, side: str):
    """The border on one outer vertical edge of a row: its first cell's left, or its
    last cell's right."""
    cells = row.cells
    if not cells:
        return None
    return (cells[0] if side == "left" else cells[-1]).borders.get(side)


def left_offset(cell: Cell) -> int:
    """Twips from a cell's left grid line to where its text starts: the larger of its
    left margin and the inner part of *its own* left border -- not of the one Word draws
    there when a neighbour's is wider (``make_table_border_probe.py``: a cell whose left
    border is ``w:sz`` 4 beside a neighbour's 24 starts its text as if the 4 were
    drawn)."""
    return max(cell.margins["left"], inner_part(cell.borders.get("left")))


def budget_left(cell: Cell) -> int:
    """Twips from a cell's left grid line to where its line budget starts: the larger of
    its left margin and the inner part of the border Word draws there (the wider of its
    own and its neighbour's).  Where that is wider than its own, the text still starts
    at :func:`left_offset`, and the line ends that much sooner."""
    return max(cell.margins["left"], inner_part(cell.left_edge[0]))


def right_offset(cell: Cell) -> int:
    """Twips from where a cell's text may end to its right grid line: the larger of its
    right margin and the near part of the border Word draws there."""
    return max(cell.margins["right"], inner_part(cell.right_edge[0]))


def grid_origin(document: Document, resolved: ResolvedTable, column_width: int) -> int:
    """Twips from the text column's left edge to the table's first grid line.

    Measured (``make_table_geometry_probe.py``):

    * **mode 15**: the left border's *outer* edge is at ``w:tblInd``, so the first grid
      line is its outer half further in;
    * **below 15, with ``w:tblInd`` stated** (by the table or its style): the first
      cell's *text* is at the indent, so the grid line is its left offset further out;
    * **below 15, with no ``w:tblInd`` anywhere**: the first grid line is at the margin;
    * **centred** and **right-aligned** tables (``w:jc``) ignore the indent: below 15
      the grid lines are centred, or the last cell's text ends at the right margin; in
      mode 15 the table's outer edges, borders included, are centred or end there.
    """
    properties = resolved.properties
    first_row = resolved.rows[0] if resolved.rows else None
    width = sum(table_widths(document, resolved, column_width))
    indent = _margin(properties.get("tblInd"))
    jc = properties.get("jc") or "left"
    mode15 = compatible_mode(document) or resolved.nested
    first = first_row.cells[0] if first_row and first_row.cells else None
    last = first_row.cells[-1] if first_row and first_row.cells else None
    # The outer parts of the first row's outer borders: beyond the grid.
    outer_left = outer_part(first.left_edge[0]) if first else 0
    outer_right = outer_part(last.right_edge[0]) if last else 0
    if jc in ("center", "centre"):
        if mode15:
            outer = width + outer_left + outer_right
            return Fraction(column_width - outer, 2) + outer_left
        return Fraction(column_width - width, 2)
    if jc in ("right", "end"):
        if resolved.nested:
            # Its last grid line on the cell's text end (make_nested_table_probe.py).
            return column_width - width
        if mode15:
            return column_width - outer_right - width
        return column_width + (right_offset(last) if last else 0) - width
    if resolved.nested:
        # A negative indent is none (make_nested_table_probe.py, tblInd -200).
        return max(indent, 0) + outer_left
    if mode15:
        return indent + outer_left
    if resolved.indent_stated:
        return indent - (left_offset(first) if first else 0)
    return 0


def grid_lines(document: Document, resolved: ResolvedTable, column_width: int) -> list:
    """Twips from the text column's left edge to every grid line, left to right."""
    x = grid_origin(document, resolved, column_width)
    out = [x]
    for width in table_widths(document, resolved, column_width):
        x += width
        out.append(x)
    return out


@dataclass(frozen=True)
class CellBox:
    """Where a cell's text goes, horizontally.

    ``text_left`` is in device px from the page's left edge: a whole pixel, as Word holds
    it -- in its layout unit, so 740 px is 740.0004 (727,450 units), as Word's PDF prints
    every cell's text start;
    ``budget`` is what a line may fill, twips; ``left`` / ``right`` its grid lines,
    exact px from the page's left edge.
    """

    text_left: int
    budget: int
    left: Fraction
    right: Fraction


def text_left_units(document: Document, resolved: ResolvedTable, page_left_twips: int, lines: list,
                    cell: Cell) -> int:
    """Where a cell's text starts, in Word's layout unit from the page's left edge,
    before it is snapped to a whole device pixel.

    The grid line plus :func:`left_offset`, rounded to the unit -- except below mode 15
    for a left-aligned table with no ``w:tblInd`` anywhere, whose position Word keeps as
    two lengths: the table's edge, the margin less the inner half of the first cell's
    left border (:func:`inner_part`), and the text's distance from that edge.  Each is
    truncated to the unit (the edge's subtraction rounding it up), and at a tie
    (1548 twips is 322.5 px) that decides the pixel: measured over 17 border widths and
    16 margins about the tie (``make_table_geometry_probe.py``), the text is drawn a
    pixel right exactly where the half's unit fraction is 0.6 or more (``w:sz`` 1, 2, 5,
    6, 9, 10, 18), and every other case agrees with plain rounding.
    """
    offset = lines[cell.column] + left_offset(cell)
    jc = resolved.properties.get("jc") or "left"
    if (not compatible_mode(document) and not resolved.nested and not resolved.indent_stated
            and jc not in ("center", "centre", "right", "end")
            and resolved.rows and resolved.rows[0].cells):
        half = inner_part(resolved.rows[0].cells[0].borders.get("left"))
        edge = math.ceil((page_left_twips - half) * UNITS_PER_TWIP)
        return edge + math.floor((half + offset) * UNITS_PER_TWIP)
    return twips_to_units(page_left_twips + offset)


def cell_box(document: Document, resolved: ResolvedTable, page_left_twips: int, lines: list, cell: Cell) -> CellBox:
    """The cell's text start (whole device px) and line budget (twips).

    ``page_left_twips`` is the text column's left edge from the page's, ``lines`` the
    grid lines (:func:`grid_lines`).  The text starts at the grid line plus
    :func:`left_offset` (:func:`text_left_units`), rounded to a whole device pixel --
    every probe cell's text is drawn at a whole pixel -- and a line may reach from there
    the grid distance less both offsets, so its right edge moves with that rounding.
    """
    left = lines[cell.column]
    right = lines[min(cell.column + cell.span, len(lines) - 1)]
    start = text_left_units(document, resolved, page_left_twips, lines, cell)
    budget = right - left - budget_left(cell) - right_offset(cell)
    return CellBox(quantise(Fraction(round_half_up(start * LAYOUT_UNIT_PX))), int(budget),
                   twips_to_px(page_left_twips + left), twips_to_px(page_left_twips + right))


# -- widths --------------------------------------------------------------------------------


def column_widths(resolved: ResolvedTable) -> tuple[int, ...]:
    """The grid columns' widths, twips.

    ``w:tblGrid`` as authored -- unless every cell of every row states its width
    (``w:tcW`` in ``dxa``) and those widths put the grid lines somewhere else, when Word
    lays the table out on the cells' widths: measured with a grid of 1000/3000/2000
    under cells of 2000 each, and 2000s under 1000/3000/2000, in both layouts
    (``make_table_geometry_probe.py``, ``widths``)."""
    grid = tuple(resolved.table.grid)
    lines: set[int] | None = None
    count = 0
    for row in resolved.rows:
        x = sum(grid[:row.grid_before])
        edges = []
        for cell in row.cells:
            width = cell.properties.get("tcW")
            if not width or width[1] != "dxa" or width[0] <= 0:
                return grid
            x += width[0]
            edges.append(x)
        count = max(count, len(edges))
        row_lines = set(edges)
        lines = row_lines if lines is None else lines | row_lines
    if not resolved.rows or lines is None:
        return grid
    ordered = sorted(lines)
    derived = tuple(b - a for a, b in zip([0] + ordered, ordered))
    if len(derived) != len(grid) or any(row.grid_before or row.grid_after for row in resolved.rows):
        return grid
    return derived


#: ``w:tblW`` in ``pct``: fiftieths of a percent.
PERCENT_WHOLE = 5000


def in_percent(resolved: ResolvedTable) -> bool:
    width = resolved.properties.get("tblW")
    return bool(width) and width[1] == "pct"


def percent_widths(document: Document, resolved: ResolvedTable, column_width: int) -> tuple[int, ...]:
    """The columns of a table whose width is stated in percent (``w:tblW`` ``pct``), twips,
    whatever ``w:tblGrid`` says; :class:`Unsupported` where the probe did not measure it.

    Measured by ``make_pct_table_probe.py`` (ROADMAP.md, "Tables -- measured", stage 7a),
    autofit tables of two to six columns whose content fits, modes none, 14 and 15:

    * **100 percent** is the text column and the first cell's left and the last cell's
      right offset (:func:`left_offset`: the margin, or the border's inner half where
      that is wider) below mode 15 -- 9,380 twips of a 9,164-twip column with margins of
      108, 9,174 with none and borders of ``w:sz`` 4, 9,764 with 300 -- and the column
      less the outer halves of the outer borders in mode 15 (9,154), whatever the
      margins;
    * the table is ``w:tblW`` of that, and its columns share it **in proportion to the
      cells' ``w:tcW`` in percent, whatever is in them** (1,000 / 2,000 pct under 5,000:
      a third and two thirds; a 16-letter word beside a label, or 15 words beside none,
      change nothing), each grid line at the floor of its share of the whole,
      cumulatively (the grid Word writes for 818 / 849 / 849 / 850 / 851 / 783 pct of
      9,576 twips: 1,566, 1,626, 1,626, 1,628, 1,630, 1,500);
    * where the shares in percent add up to more than ``w:tblW``, the columns before the
      last take what they say and the last what is left (1,000 / 4,500 of 5,000: 1,877 and
      7,502 twips).

    A column narrower than a word in it is widened by Word to hold it (782-twip shares
    of three columns labelled with a 671-twip word: 888 twips each), as autofit does: that
    stops the layout as a word wider than its column does (:func:`flow_table`).  **Cells
    in ``dxa`` under a width in percent are shared by their widths and their content**
    (:func:`dxa_shares`, with :attr:`ResolvedTable.minimums`).  Not modelled: a row whose
    cells differ from the first's, a cell across columns, cells in percent and in ``dxa``
    together, ``w:tcW`` in percent under ``w:tblW`` auto, a fixed layout.
    """
    table_width = resolved.properties["tblW"][0]
    rows = resolved.rows
    if not rows or not rows[0].cells or not isinstance(table_width, int) or table_width <= 0:
        raise Unsupported("a table width in percent that is not a number")
    if (resolved.properties.get("tblLayout") or "autofit") == "fixed":
        raise Unsupported("a fixed table with a width in percent (not measured)")
    shares = [cell.properties.get("tcW") for cell in rows[0].cells]
    kinds = {share[1] if share else None for share in shares}
    if kinds == {"dxa"} and resolved.minimums is None:
        raise Unsupported("autofit: a table in percent whose cells' widths are in twips, shared by their content")
    if kinds not in ({"pct"}, {"dxa"}) or any(share[0] <= 0 for share in shares):
        raise Unsupported("a table in percent whose cells' widths are not all in percent")
    columns = len(resolved.table.grid)
    for row in rows:
        if (row.grid_before or row.grid_after or len(row.cells) != columns
                or any(cell.span != 1 for cell in row.cells)
                or [cell.properties.get("tcW") for cell in row.cells] != shares):
            raise Unsupported("a table in percent whose rows differ, or with a cell across columns (not measured)")
    first, last = rows[0].cells[0], rows[0].cells[-1]
    if compatible_mode(document) or resolved.nested:
        # A nested table's whole is its cell's text area less its outer borders' outer
        # halves in every mode (make_nested_table_probe.py, 5000 pct: 4,274 of 4,284).
        whole = column_width - outer_part(first.left_edge[0]) - outer_part(last.right_edge[0])
    else:
        whole = column_width + left_offset(first) + right_offset(last)
    total = whole * table_width // PERCENT_WHOLE
    values = [share[0] for share in shares]
    if kinds == {"dxa"}:
        return dxa_shares(values, total, resolved.minimums)
    cumulative = [sum(values[:k + 1]) for k in range(len(values))]
    if kinds == {"pct"} and cumulative[-1] > table_width:
        edges = [whole * c // PERCENT_WHOLE for c in cumulative[:-1]] + [total]
    else:
        edges = [total * c // cumulative[-1] for c in cumulative]
    return tuple(b - a for a, b in zip([0] + edges, edges))


def dxa_shares(preferred: list, total: int, minimums) -> tuple[int, ...]:
    """The columns of a table in percent whose cells state their widths in ``dxa``
    (``preferred``), sharing ``total`` twips: in proportion to those widths where they add
    up to no more than it; where they add up to more, **each gives up the excess in
    proportion to its width less its narrowest content** (``minimums``, twips:
    :func:`column_minimums`).  Each grid line at the floor of its cumulative share.

    Measured by ``make_pct_table_probe.py`` (stage 7a) and ``make_nested_table_probe.py``:
    1,000 / 3,000 twips under 3,500 pct (6,566 twips) are drawn a quarter and three
    quarters; 4,788 / 4,788 with a 16-letter word (1,524 twips) in one and a 671-twip label
    in the other 3,470 / 3,096 (the rule: 3,467 / 3,099); with a nested table whose
    narrowest content is 2,127 twips in one, 3,624 / 2,942 (the rule: 3,628 / 2,938) --
    each within the pixel Word drew it at."""
    columns = len(preferred)
    if minimums is None or len(minimums) != columns:
        raise Unsupported("autofit: a table in percent whose cells' widths are in twips, shared by their content")
    wanted = sum(preferred)
    if wanted <= total:
        shares = [Fraction(total * p, wanted) for p in preferred]
    else:
        room = [p - m for p, m in zip(preferred, minimums)]
        if any(r < 0 for r in room) or sum(room) <= 0 or sum(minimums) > total:
            raise Unsupported("autofit: a table in percent narrower than its content (not measured)")
        excess = wanted - total
        shares = [p - Fraction(excess) * r / sum(room) for p, r in zip(preferred, room)]
    edges, x = [], Fraction(0)
    for share in shares:
        x += share
        edges.append(math.floor(x))
    edges[-1] = total
    return tuple(b - a for a, b in zip([0] + edges, edges))


def paragraph_minimum(document: Document, paragraph, section, advances, metrics) -> Fraction:
    """The narrowest a paragraph's text goes, twips: its widest word (the pieces between
    two spaces, breaks or tabs), and its left and right indents."""
    from . import linebreak

    try:
        pieces, _lines = linebreak.break_paragraph(document, paragraph, section, advances, metrics)
    except linebreak.Unmeasurable:
        raise Unsupported("a cell paragraph that cannot be measured") from None
    widest = word = Fraction(0)
    for piece in pieces:
        if piece.kind in (linebreak.SPACE, linebreak.BREAK, linebreak.TAB):
            word = Fraction(0)
            continue
        word += Fraction(piece.width)
        widest = max(widest, word)
    from .resolve import resolve_paragraph

    pp = resolve_paragraph(document, paragraph)
    indents = (pp.get("ind.left", 0) or 0) + (pp.get("ind.right", 0) or 0)
    return widest / UNITS_PER_TWIP + max(indents, 0)


#: Twips an empty cell's content takes in Word's autofit, beyond its margins: cells with
#: nothing in them, and one stated 150 twips wide, are drawn 222 twips wide with margins
#: of 108 -- 221 puts a grid line a pixel from Word's (``make_autofit_width_probe.py``,
#: ``auto`` and ``widen``, every mode).
EMPTY_CONTENT = 6


def paragraph_extent(document: Document, paragraph, section, advances, metrics) -> tuple[Fraction, Fraction]:
    """The narrowest and the widest a paragraph's text goes in Word's autofit, twips: its
    widest word, and its longest line unbroken (a break starting a new one), each with
    the indents in front of it and its right indent.

    Measured (``make_autofit_width_probe.py``, ``indent``): **below mode 15 a first-line
    indent counts for nothing** and a hanging one takes its width off the left indent --
    every word counts from the paragraph's nearest line start, ``left + min(firstLine,
    0)``: a 953-twip word in a 1,440-twip cell with a first-line indent of 720 keeps the
    cell's width and is broken inside, as a filesamples document's calendar breaks its
    day names; with a left indent of 720 the cell is widened to hold both.  **In mode 15
    each line counts from where it starts**: the first word and the first line from
    ``left + firstLine``, the rest from ``left``."""
    from . import linebreak
    from .resolve import resolve_paragraph

    try:
        pieces = linebreak.pieces(document, paragraph, advances, metrics)
    except linebreak.Unmeasurable:
        raise Unsupported("a cell paragraph that cannot be measured") from None
    if any(piece.kind in (linebreak.TAB, linebreak.OBJECT) for piece in pieces):
        raise Unsupported("autofit: a tab or a drawing in a cell sized by its content (not measured)")
    pp = resolve_paragraph(document, paragraph)
    left = pp.get("ind.left", 0) or 0
    right = pp.get("ind.right", 0) or 0
    first = (pp.get("ind.firstLine", 0) or 0) - (pp.get("ind.hanging", 0) or 0)
    mode15 = compatible_mode(document)
    first_start = left + first if mode15 else left + min(first, 0)
    rest_start = left if mode15 else first_start
    widest = longest = Fraction(0)
    word = line = trailing = Fraction(0)
    start, first_word, first_line = first_start, True, True
    for piece in pieces + [linebreak.Piece(linebreak.BREAK, "\n")]:
        if piece.kind == linebreak.BREAK:
            longest = max(longest, start + (line - trailing) / UNITS_PER_TWIP)
            line = trailing = Fraction(0)
            first_line = False
        if piece.kind in (linebreak.SPACE, linebreak.BREAK):
            if word:
                widest = max(widest, (first_start if first_word else rest_start) + word / UNITS_PER_TWIP)
                first_word = False
            word = Fraction(0)
            if piece.kind == linebreak.SPACE:
                line += Fraction(piece.width)
                trailing += Fraction(piece.width)
            start = first_start if first_line else rest_start
            continue
        word += Fraction(piece.width)
        line += Fraction(piece.width)
        trailing = Fraction(0)
    return max(widest, Fraction(0)) + right, max(longest, Fraction(0)) + right


def cell_extent(document: Document, cell: Cell, section, advances, metrics) -> tuple[Fraction, Fraction]:
    """A cell's narrowest and widest content in Word's autofit, twips, with its left and
    right offsets (:func:`left_offset`, :func:`right_offset`): the widest of its
    paragraphs' (:func:`paragraph_extent`); an empty cell :data:`EMPTY_CONTENT`.  A table
    nested in it is not measured for autofit."""
    low = high = Fraction(0)
    for item in cell.blocks:
        if isinstance(item, Table):
            raise Unsupported("autofit: a table nested in a cell sized by its content (not measured)")
        narrow, wide = paragraph_extent(document, item, section, advances, metrics)
        low, high = max(low, narrow), max(high, wide)
    if not high:
        low = high = Fraction(EMPTY_CONTENT)
    offsets = left_offset(cell) + right_offset(cell)
    return low + offsets, high + offsets


def column_minimums(document: Document, resolved: ResolvedTable, section, advances, metrics) -> tuple:
    """Each grid column's narrowest content, twips: the widest of its cells' narrowest
    paragraph or nested table (:func:`table_minimum`), with the cell's left and right
    offsets.  A cell spanning columns is not measured."""
    columns = len(resolved.table.grid)
    out = [Fraction(0)] * columns
    for row in resolved.rows:
        for cell in row.cells:
            if cell.merge == "continue":
                continue
            if cell.span != 1 or cell.column >= columns:
                raise Unsupported("autofit: a table in percent with a merged cell, shared by content (not measured)")
            content = Fraction(0)
            for item in cell.blocks:
                if isinstance(item, Table):
                    content = max(content, table_minimum(document, item, section, advances, metrics))
                else:
                    content = max(content, paragraph_minimum(document, item, section, advances, metrics))
            out[cell.column] = max(out[cell.column], content + left_offset(cell) + right_offset(cell))
    return tuple(out)


def table_minimum(document: Document, table: Table, section, advances, metrics) -> Fraction:
    """The narrowest a nested table goes, twips: its columns' narrowest content and the
    outer halves of its outer borders (``make_nested_table_probe.py``: a nested table of
    3,000 twips and one of 1,000 whose cells hold the same words weigh alike)."""
    resolved = resolve_table(document, table)
    minimums = column_minimums(document, resolved, section, advances, metrics)
    first = resolved.rows[0].cells[0] if resolved.rows and resolved.rows[0].cells else None
    last = resolved.rows[0].cells[-1] if resolved.rows and resolved.rows[0].cells else None
    edges = (Fraction(border_twips(first.left_edge[0]), 2) if first else 0) + (
        Fraction(border_twips(last.right_edge[0]), 2) if last else 0)
    return sum(minimums, Fraction(0)) + edges


def table_widths(document: Document, resolved: ResolvedTable, column_width: int) -> tuple[int, ...]:
    """The grid columns' widths, twips: :func:`autofit_widths` for an autofit table whose
    cells' content has been measured (:attr:`ResolvedTable.extents`), else
    :func:`percent_widths` for a table whose width is stated in percent, else
    :func:`column_widths`."""
    if resolved.extents is not None:
        return autofit_widths(document, resolved, column_width)
    if in_percent(resolved):
        return percent_widths(document, resolved, column_width)
    return column_widths(resolved)


def _outer_edges(resolved: ResolvedTable) -> tuple:
    first = resolved.rows[0].cells[0] if resolved.rows and resolved.rows[0].cells else None
    last = resolved.rows[0].cells[-1] if resolved.rows and resolved.rows[0].cells else None
    return first, last


def autofit_room(document: Document, resolved: ResolvedTable, column_width: int) -> int:
    """The widest an autofit table goes, twips: below mode 15 the text column and the
    last cell's right offset (its grid from the margin, its last cell's text ending at the
    right margin: 9,272 twips of a 9,164-twip column with margins of 108); in mode 15, and
    nested in a cell, the column less the outer halves of the outer borders (9,154)
    (``make_autofit_width_probe.py``, ``share``)."""
    first, last = _outer_edges(resolved)
    if compatible_mode(document) or resolved.nested:
        return column_width - (outer_part(first.left_edge[0]) if first else 0) - (
            outer_part(last.right_edge[0]) if last else 0)
    return column_width + (right_offset(last) if last else 0)


def autofit_widths(document: Document, resolved: ResolvedTable, column_width: int) -> tuple[int, ...]:
    """The columns of an autofit table sized from its content, twips.

    Measured by ``make_autofit_width_probe.py`` (ROADMAP.md, "Tables -- measured", stage
    7b), modes none, 14 and 15, every cell's narrowest and widest content from
    :func:`cell_extent`:

    * **a column is as wide as its cells' width in ``dxa``, or where none states one as
      its widest content**, and never narrower than its narrowest content (a word wider
      than its width in ``dxa`` widens it, and the table);
    * **where that is wider than the room** (:func:`autofit_room`), the columns with no
      width give way first -- each keeps its narrowest content and takes a share of what
      is left **in proportion to its widest content less its narrowest**
      (60 words beside 10: 7,430 twips of 9,272, the rule 7,430); where there are none,
      the columns in ``dxa`` share it the same way, by their width less their narrowest;
    * **a table width in ``dxa`` or in percent** (of :func:`percent_widths`' whole) wider
      than the columns widens the columns with no width in proportion to their widest
      content, or, where every column is in ``dxa``, all of them in proportion to their
      widths; narrower, it is shared as the room is;
    * a column sized by its content is as wide as it to the twip above; each grid line of
      a table whose columns grew or gave way at the floor of its cumulative width.

    **Narrower than its content** -- the columns' narrowest content adding up to more than
    the room -- measured by ``make_autofit_over_probe.py``, modes none, 14 and 15: the
    table keeps to the room, whatever width it states in ``dxa``, each column its offsets
    and a share of the rest **in proportion to its narrowest content** (its widest word),
    the words broken inside where wider (nine columns of two-word labels: 1,157 / 955 /
    1,022 ... twips, the rule 1,157 / 955 / 1,023).  Below mode 15 a width in percent
    narrower than the content is drawn past the margin, its words broken by a rule not
    settled: a stop.

    A table whose cells state their widths in ``dxa`` and hold their content in the room
    keeps its stated grid (:func:`column_widths`).  Not measured, and so
    :class:`Unsupported`: a cell across columns wider than they are, ``w:gridBefore`` or
    ``w:gridAfter`` beside cells to be sized, a column only spanned, cells in percent, a
    column in ``dxa`` that must give way beside ones with no width at their narrowest,
    a cell with ``w:noWrap`` and no width that must give way, a table width in ``dxa``
    wider than the room.  (``w:noWrap``
    in a cell in ``dxa`` changes nothing: such a cell of 300 twips is widened to its
    widest word, as one without.)
    """
    extents = resolved.extents
    columns = len(resolved.table.grid)
    rows = resolved.rows
    room = autofit_room(document, resolved, column_width)
    stated = resolved.properties.get("tblW")
    target = None
    if stated and stated[1] == "dxa" and isinstance(stated[0], int) and stated[0] > 0:
        target = stated[0]
    elif stated and stated[1] == "pct" and isinstance(stated[0], int) and stated[0] > 0:
        first, last = _outer_edges(resolved)
        if compatible_mode(document) or resolved.nested:
            whole = room
        else:
            whole = column_width + left_offset(first) + right_offset(last)
        target = whole * stated[0] // PERCENT_WHOLE
    cells = [(row, cell) for row in rows for cell in row.cells if cell.merge != "continue"]

    def dxa_width(cell: Cell):
        width = cell.properties.get("tcW")
        if width and width[1] == "dxa" and isinstance(width[0], int) and width[0] > 0:
            return width[0]
        if width and width[1] == "pct" and width[0]:
            raise Unsupported("autofit: a cell width in percent in a table not in percent (not measured)")
        return None

    if target is None and all(dxa_width(cell) is not None and extents[(row.index, cell.index)][0] <= dxa_width(cell)
                              for row, cell in cells):
        stated_grid = column_widths(resolved)
        if sum(stated_grid) <= room:
            # Word keeps the widths it wrote while the content holds (stage 7).
            return stated_grid
    if any(row.grid_before or row.grid_after for row in rows):
        raise Unsupported("autofit: w:gridBefore or w:gridAfter in a table sized by its content (not measured)")
    lows = [Fraction(0)] * columns
    #: Each column's cell offsets (margins, or borders' inner halves) where its narrowest
    #: content is: what it keeps when content wider than the room is shared.
    offsets = [Fraction(0)] * columns
    highs: list = [None] * columns
    dxas: list = [None] * columns
    spanned = []
    for row, cell in cells:
        low, high = extents[(row.index, cell.index)]
        width = dxa_width(cell)
        if cell.span != 1:
            spanned.append((cell, low, high, width))
            continue
        c = cell.column
        if c >= columns:
            raise Unsupported("autofit: a row with more cells than the grid (not measured)")
        if low > lows[c]:
            offsets[c] = Fraction(left_offset(cell) + right_offset(cell))
        lows[c] = max(lows[c], low)
        if width is not None:
            dxas[c] = max(dxas[c] or 0, width)
        else:
            highs[c] = max(highs[c] or Fraction(0), high)
    if any(d is None and h is None for d, h in zip(dxas, highs)):
        raise Unsupported("autofit: a grid column no cell of its own sizes (not measured)")
    if any(d is not None and h is not None for d, h in zip(dxas, highs)):
        raise Unsupported("autofit: a column of cells in dxa and cells with no width (not measured)")
    auto = [d is None for d in dxas]
    wants = [max(Fraction(d) if d is not None else h, low) for d, h, low in zip(dxas, highs, lows)]
    for cell, low, high, width in spanned:
        need = max(Fraction(width) if width is not None else high, low)
        if need > sum(wants[cell.column:cell.column + cell.span]):
            raise Unsupported("autofit: a cell across columns wider than they are (measured, not settled)")
    total = sum(wants)
    narrowest = sum(lows)
    percent_past = (target is not None and stated[1] == "pct"
                    and not (compatible_mode(document) or resolved.nested))
    if percent_past and narrowest > target:
        # Below mode 15 Word draws it past the margin, about as wide as its content
        # (5000 pct over three columns of 3,552: 3,552 / 3,552 / 3,547), its words broken
        # by a rule not settled.
        raise Unsupported("autofit: a table in percent narrower than its content (measured, not settled)")
    if narrowest > room and not spanned and not percent_past:
        # Narrower than its content: the table keeps to the room, whatever width it states,
        # each column its margins and a share of the rest in proportion to its narrowest
        # content -- its widest word, broken inside where it is wider than that.
        content = [low - offset for low, offset in zip(lows, offsets)]
        left = Fraction(room) - sum(offsets)
        if left <= 0 or sum(content) <= 0:
            raise Unsupported("autofit: a table whose margins are wider than the room (not measured)")
        # Each grid line at the nearest twip below mode 15, at the floor in it (five columns
        # of 1,854.4 twips: the third's text a pixel left of Word's either way round).
        return _edges([offset + left * part / sum(content) for offset, part in zip(offsets, content)],
                      nearest=not (compatible_mode(document) or resolved.nested))
    goal = target if target is not None else min(total, Fraction(room))
    if target is not None and target > room and not (stated[1] == "pct"):
        raise Unsupported("autofit: a table width in dxa wider than the room (not measured)")
    if goal > total:
        grow = [c for c in range(columns) if auto[c]] or list(range(columns))
        base = sum(wants[c] for c in grow)
        widths = [wants[c] + (goal - total) * wants[c] / base if c in grow else wants[c] for c in range(columns)]
    elif goal < total:
        if any(cell.properties.get("noWrap") and auto[cell.column] for _row, cell in cells if cell.span == 1):
            raise Unsupported("autofit: a cell with w:noWrap and no width giving way (not measured)")
        if any(auto) and sum(w for w, a in zip(wants, auto) if not a) + sum(
                low for low, a in zip(lows, auto) if a) <= goal:
            share = [c for c in range(columns) if auto[c]]
        elif not any(auto):
            share = list(range(columns))
        else:
            raise Unsupported("autofit: columns in dxa giving way beside columns at their narrowest (not measured)")
        left = goal - sum(wants[c] for c in range(columns) if c not in share) - sum(lows[c] for c in share)
        slack = sum(wants[c] - lows[c] for c in share)
        if left < 0 or slack <= 0:
            raise Unsupported("autofit: a table narrower than its content (not measured)")
        widths = [lows[c] + left * (wants[c] - lows[c]) / slack if c in share else wants[c] for c in range(columns)]
    else:
        # Each as wide as its content, to the twip above: what fits in it fits on its line.
        widths = [math.ceil(want) for want in wants]
    edges, x = [], Fraction(0)
    for width in widths:
        x += width
        # A table widened to its stated width: each grid line at the nearest twip (6,000
        # twips over content of 952 and 1,534: 2,298, drawn a pixel right of 2,297).
        edges.append(round_half_up(x) if goal > total else math.floor(x))
    return tuple(b - a for a, b in zip([0] + edges, edges))


def _edges(widths, *, nearest: bool) -> tuple[int, ...]:
    """Columns of ``widths``, each grid line at the nearest twip of its cumulative width, or
    at its floor."""
    edges, x = [], Fraction(0)
    for width in widths:
        x += width
        edges.append(round_half_up(x) if nearest else math.floor(x))
    return tuple(b - a for a, b in zip([0] + edges, edges))


# -- the table in the flow -----------------------------------------------------------------


@dataclass
class CellParagraph:
    """One paragraph of a cell, broken into lines at the cell's width."""

    paragraph: object
    #: Its index in ``baselines.blocks`` order (table cells counted), as a
    #: :class:`docx2svg.paginate.Para`'s.
    block: int
    resolved: object
    pieces: list
    lines: list
    #: Line geometry from the cell's text start (:class:`docx2svg.linebreak.Geometry`).
    geometry: object
    label: str | None
    heights: list
    #: Space before and after, twips, as the body stacks them; its borders' room above
    #: and below; contextual spacing and style, for the gap to its neighbour.
    before: int = 0
    after: int = 0
    top_border: Fraction = Fraction(0)
    border: Fraction = Fraction(0)
    contextual: bool = False
    style: str | None = None
    undrawn: frozenset = frozenset()
    #: The paragraph style that applies (the default one where none is named).
    effective_style: str | None = None
    #: The cell's end mark, hidden (:func:`hides_mark`): the paragraph takes no room.
    hidden: bool = False
    #: Its floating drawings, as :func:`docx2svg.paginate.anchors_of` lists a body
    #: paragraph's.
    anchors: tuple = ()
    #: Line -> how far it went down, px, below a drawing it could not go beside or into
    #: the band of one text wraps above and below (:func:`wrap_cell`).
    drops: dict = field(default_factory=dict)
    #: ``(text index, run, index in the run)`` -> ``(x, y)`` of a drawing text wraps
    #: around in the cell, px: from the page's left edge, and from below the row's top
    #: border (:func:`wrap_cell`).
    positioned: dict = field(default_factory=dict)
    #: A table nested in the cell, stacked as a paragraph of one line as tall as the
    #: table (:func:`nested_entry`); ``None`` for a paragraph.
    table: "TableFlow | None" = None


@dataclass(frozen=True)
class NestedHeight:
    """The one "line" a nested table takes in its cell's stack: as tall as the table."""

    pitch: Fraction


@dataclass
class CellFlow:
    cell: Cell
    box: CellBox
    paragraphs: list[CellParagraph]
    #: The paragraphs' height, px: the first one's space before to the last one's after.
    content: Fraction = Fraction(0)
    #: A cell merged down (``w:vMerge="restart"``): how many rows it covers, and the height
    #: of those rows from below its top border to the edge below the last (its lines are
    #: aligned in that); a ``continue`` cell below it is ``merged``, and draws nothing.
    rows: int = 1
    span_height: Fraction | None = None
    merged: bool = False
    #: Its lines as the cell stacks them (:func:`cell_stack`), and the space after the last.
    stack: list = field(default_factory=list)
    after: Fraction = Fraction(0)


@dataclass
class RowFlow:
    row: Row
    cells: list[CellFlow]
    #: The border on the edge above the row, px (the table's top border for the first).
    top_border: Fraction = Fraction(0)
    #: From below that border to the edge below: the tallest cell with its margins, or
    #: the row's ``w:trHeight``.
    height: Fraction = Fraction(0)


@dataclass
class TableFlow:
    """A top-level table as the paginator and the page walk see it."""

    table: Table
    resolved: ResolvedTable
    rows: list[RowFlow]
    #: The border below the last row, px.
    bottom_border: Fraction = Fraction(0)
    #: Every grid line, px from the page's left edge, exact.
    lines: list = field(default_factory=list)
    #: The style of the paragraph before the table and the default paragraph style, for
    #: contextual spacing at a cell's edges (:func:`cell_stack`).
    previous_style: str | None = None
    default_style: str | None = None
    #: The text column's left edge from the page's, twips, and every grid line from that
    #: edge, twips (:func:`grid_lines`): what the drawing rounds (``layout._row_borders``).
    page_left_twips: int = 0
    grid_twips: list = field(default_factory=list)
    #: Where a floating table is positioned (:func:`floating_of`); ``None`` in the flow.
    floating: "Floating | None" = None


class Unsupported(Exception):
    """A table the model does not lay out yet; ``args[0]`` says why."""


def border_px(border) -> Fraction:
    """The room a border line takes across its edge, px: its width in whole twips."""
    return twips_to_px(border_twips(border))


def cell_blocks(table: Table):
    """Every paragraph of a table in ``baselines.blocks`` order, with the cell it is in
    (``(row, cell)``); a nested table's paragraphs count, with ``None`` for their cell."""
    for r, row in enumerate(table.rows):
        for c, cell in enumerate(row):
            for item in cell:
                if isinstance(item, Table):
                    for _ in cell_blocks(item):
                        yield None, None
                else:
                    yield (r, c), item


def _edge_width(cells_above: list[Cell], cells_below: list[Cell], columns: int) -> Fraction:
    """The border on the edge between two rows (either may be empty: the table's top or
    bottom), px: the widest one stated there, by the cells on either side."""
    widths = [border_px(cell.borders.get("bottom")) for cell in cells_above]
    widths += [border_px(cell.borders.get("top")) for cell in cells_below]
    return max(widths, default=Fraction(0))


def _unmodelled(resolved: ResolvedTable) -> str | None:
    """What a table needs that :data:`STAGES` does not model yet, if anything."""
    cells = [cell for row in resolved.rows for cell in row.cells]
    if "rows" not in STAGES:
        if any(row.properties.get("trHeight") for row in resolved.rows):
            return "a stated row height (w:trHeight)"
        if any((cell.properties.get("vAlign") or "top") != "top" for cell in cells):
            return "a cell aligned vertically (w:vAlign)"
    if "merges" not in STAGES:
        if any(cell.merge or cell.span > 1 for cell in cells) or any(row.grid_before or row.grid_after
                                                                     for row in resolved.rows):
            return "merged cells (w:gridSpan, w:vMerge, w:gridBefore, w:gridAfter)"
    if "borders" not in STAGES:
        if any(cell.own.get(side) for cell in cells for side in SIDES) or any(
                cell.properties.get("shd") for cell in cells):
            return "a cell's own borders or shading"
        if any(key.startswith("tblBorders.") and drawn(value) and value[0] != "single"
               for key, value in resolved.properties.items()):
            return "a table border that is not a single line"
    return None


def autofit_unmodelled(resolved: ResolvedTable) -> str | None:
    """What of an autofit table Word would size from its content, which the model does
    not: it lays an autofit table out on the widths the file states, as Word does while
    they hold (``make_table_autofit_probe.py``, ROADMAP.md "Tables -- measured", stage
    7).  A cell with no ``w:tcW`` is not one of them -- Word makes its column as wide as
    its content and margins, whatever ``w:tblGrid`` says -- nor a cell width in percent
    or ``auto`` in a table whose width is not in percent (not measured).  (A table
    width in percent is shared out by :func:`percent_widths`.)  (A word
    wider than its column, the third case, is found as the cell is broken.)"""
    if in_percent(resolved):
        # Shared out as its cells' widths say (:func:`percent_widths`, stage 7a).
        return None
    if any(not cell.properties.get("tcW") or cell.properties["tcW"][1] != "dxa"
           for row in resolved.rows for cell in row.cells if cell.merge != "continue"):
        return "autofit: a cell with no width in twips (w:tcW), sized to its content"
    return None


#: Twips text keeps off a floating table at least, whatever ``*FromText`` says: a
#: ``rightFromText`` of 0, 5 or 10 keeps it 10 twips off, 15 and more what they say
#: (``make_float_table_probe.py``, ``edge``, every mode).
MIN_FROM_TEXT = 10
#: Twips a floating table's edge with no border reaches past its grid line, for the text
#: beside it: a table with no borders keeps text 15 twips further off than its grid line
#: and ``rightFromText`` put it, where one with borders keeps it the border's outer half
#: off (``make_float_table_probe.py``, ``edge``: ``rightFromText`` 0-500 with no borders).
UNBORDERED_EDGE = 15


@dataclass(frozen=True)
class Floating:
    """A floating table (``w:tblpPr``) as :func:`floating_of` reads it: what it is
    positioned against vertically (``text``, ``margin`` or ``page``), by ``y`` twips or
    aligned (``y_spec``), and how far text keeps off each side, twips.  Across, the
    table is laid out in its frame (the column, or the ``page``) as a table indented
    ``x`` twips, or aligned ``x_spec``, is in the column (:func:`flow_table`)."""

    vertical: str
    y: int = 0
    y_spec: str | None = None
    left: int = 0
    right: int = 0
    top: int = 0
    bottom: int = 0
    page: bool = False
    x: int = 0
    x_spec: str | None = None


def floating_of(document: Document, table: Table, properties: dict) -> Floating | None:
    """What ``w:tblpPr`` makes of a table: ``None`` where it is laid out in the flow;
    :class:`Unsupported` where the probe did not measure it.

    Measured by ``make_float_table_probe.py`` (ROADMAP.md, "Floating tables"), no
    settings and modes 14 and 15:

    * **Across**, ``horzAnchor`` absent, ``text`` and ``margin`` are the column, ``page``
      the page; ``tblpX`` is where the table's indent would be (below mode 15 its first
      cell's text, in mode 15 its left border's outer edge), and ``tblpXSpec`` ``left``,
      ``center`` and ``right`` place it as ``w:jc`` would in that frame.
    * **Down**, ``vertAnchor`` ``text`` is the top of the paragraph after the table --
      where the paragraph before ends after its space after, above the paragraph's own
      space before -- plus ``tblpY``; ``margin`` (and ``vertAnchor`` absent) is the top
      margin, ``page`` the page's top, each plus ``tblpY`` or aligned by ``tblpYSpec``
      ``top``, ``center`` or ``bottom``.  **Below mode 15 a table against the margin or
      the page at ``tblpY`` 0 goes where ``text`` would put it.**
    * **Not floating**: with every position 0 (no ``tblpX``, ``tblpY`` or either spec)
      and ``vertAnchor`` not ``text``, the table is laid out in the flow, as if it had no
      ``w:tblpPr``, in every mode.
    """
    spec = properties.get("tblpPr")
    if spec is None:
        return None

    def number(key: str) -> int:
        try:
            return int(spec.get(key) or 0)
        except ValueError:
            raise Unsupported(f"a floating table's {key} that is not a number") from None

    vertical = spec.get("vertAnchor") or "margin"
    horizontal = spec.get("horzAnchor")
    x, y = number("tblpX"), number("tblpY")
    x_spec, y_spec = spec.get("tblpXSpec"), spec.get("tblpYSpec")
    if vertical not in ("text", "margin", "page") or horizontal not in (None, "text", "margin", "page"):
        raise Unsupported("a floating table positioned against a frame not measured")
    if not x and not y and x_spec is None and y_spec is None and vertical != "text":
        if horizontal is not None:
            raise Unsupported("a floating table at 0, 0 with a horzAnchor (not measured)")
        return None
    if x_spec not in (None, "left", "center", "right") or y_spec not in (None, "top", "center", "bottom"):
        raise Unsupported("a floating table aligned inside or outside (not measured)")
    if y_spec is not None and vertical == "text":
        raise Unsupported("a floating table aligned against the text (not measured)")
    if "jc" in table.properties or "tblInd" in table.properties:
        raise Unsupported("a floating table with its own w:jc or w:tblInd (not measured)")
    if not compatible_mode(document) and vertical != "text" and not y and y_spec is None:
        vertical = "text"
    return Floating(vertical, y, y_spec, number("leftFromText"), number("rightFromText"), number("topFromText"),
                    number("bottomFromText"), horizontal == "page", x, x_spec)


def floating_top(floating: Floating, section, paragraph_top: Fraction, height: Fraction) -> Fraction:
    """A floating table's top, exact px from the page's: ``paragraph_top`` is the top of
    the paragraph after it (:func:`floating_of`), ``height`` the table's, borders
    included."""
    if floating.vertical == "text":
        return paragraph_top + twips_to_px(floating.y)
    m = section.margins
    page_height = twips_to_px(section.page_size.height_twips)
    if floating.vertical == "margin":
        start = twips_to_px(abs(m.top))
        extent = page_height - start - twips_to_px(abs(m.bottom))
    else:
        start, extent = Fraction(0), page_height
    if floating.y_spec == "center":
        return start + (extent - height) / 2
    if floating.y_spec == "bottom":
        return start + extent - height
    return start + twips_to_px(floating.y)


def floating_edges(flow: TableFlow) -> tuple[Fraction, Fraction]:
    """How far a floating table's box reaches past its outer grid lines, left and right,
    twips, for the text beside it: its outer borders' outer halves, exactly, or
    :data:`UNBORDERED_EDGE` where there is none (``make_float_table_probe.py``)."""
    rows = flow.resolved.rows
    first = rows[0].cells[0] if rows and rows[0].cells else None
    last = rows[0].cells[-1] if rows and rows[0].cells else None

    def edge(border) -> Fraction:
        # The border's exact half: a w:sz 6 border (15 twips) keeps text 7.5 twips off.
        return Fraction(border_twips(border), 2) if drawn(border) else Fraction(UNBORDERED_EDGE)

    return (edge(first.left_edge[0]) if first else Fraction(0), edge(last.right_edge[0]) if last else Fraction(0))


def flow_table(document: Document, table: Table, section, advances, metrics, *, block: int,
               labels: dict | None = None, features: list | None = None, cache: dict | None = None,
               previous=None, nested: bool = False) -> TableFlow:
    """Lay a top-level table out for the flow: its grid, each cell's paragraphs broken
    at the cell's width with their line heights, and each row's height.  Raises
    :class:`Unsupported` for what is not modelled yet (a nested or floating table, a
    vertical merge, an autofit table whose content would widen a column, a paragraph the
    model cannot measure)."""
    from dataclasses import replace as _replace

    from . import lines as line_model
    from . import linebreak
    from .model import PageMargins, PageSize
    from .resolve import resolve_paragraph
    from .resolve.cascade import _paragraph_style_id as paragraph_style_id
    from .vertical import paragraph_borders_px

    default = document.styles.default("paragraph") if document.styles is not None else None
    default_style = default.style_id if default else None
    previous_style = paragraph_style_id(document.styles, previous) if previous is not None else None

    resolved = resolve_table(document, table)
    resolved.nested = nested
    unmodelled = _unmodelled(resolved)
    if unmodelled:
        raise Unsupported(unmodelled)
    floating = floating_of(document, table, resolved.properties)
    if floating is not None and nested:
        raise Unsupported("a floating table nested in a cell (not measured)")
    if floating is not None:
        # Laid out in its frame as a table with that indent, or that alignment, is in the
        # column (:func:`floating_of`).
        if floating.page:
            section = _replace(section, margins=_replace(section.margins, left=0, right=0, gutter=0))
        resolved.properties = {**resolved.properties, "tblInd": (floating.x - (floating.x > 0), "dxa"),
                               "jc": floating.x_spec or "left"}
        resolved.indent_stated = True
    page_left = section.margins.left + section.margins.gutter
    column_width = linebreak.column_width_twips(section)
    if in_percent(resolved) and resolved.rows and resolved.rows[0].cells and all(
            (cell.properties.get("tcW") or (0, None))[1] == "dxa" for cell in resolved.rows[0].cells):
        # Shared by the cells' widths and their narrowest content (:func:`dxa_shares`).
        resolved.minimums = column_minimums(document, resolved, section, advances, metrics)
    autofit = (resolved.properties.get("tblLayout") or "autofit") != "fixed"
    kinds = {(cell.properties.get("tcW") or (0, None))[1] for cell in resolved.rows[0].cells} if (
        resolved.rows and resolved.rows[0].cells) else set()
    if autofit and not (in_percent(resolved) and kinds in ({"pct"}, {"dxa"})):
        # Sized from its content (:func:`autofit_widths`) -- or, where that cannot be
        # measured, laid out on the widths it states, as before.
        try:
            resolved.extents = {(row.index, cell.index): cell_extent(document, cell, section, advances, metrics)
                                for row in resolved.rows for cell in row.cells if cell.merge != "continue"}
        except Unsupported:
            if in_percent(resolved) or autofit_unmodelled(resolved):
                raise
            resolved.extents = None
    lines = grid_lines(document, resolved, column_width)
    if autofit and resolved.extents is None:
        unsized = autofit_unmodelled(resolved)
        if unsized:
            raise Unsupported(unsized)
    # Each cell's paragraphs and nested tables, with their index in ``baselines.blocks``
    # order (a nested table's: its first paragraph's).
    entries_of: dict[tuple[int, int], list] = {}
    index = block
    for r, row_cells in enumerate(table.rows):
        for c, cell_items in enumerate(row_cells):
            for item in cell_items:
                entries_of.setdefault((r, c), []).append((index, item))
                index += sum(1 for _ in cell_blocks(item)) if isinstance(item, Table) else 1
    rows: list[RowFlow] = []
    for row in resolved.rows:
        flows = []
        for cell in row.cells:
            box = cell_box(document, resolved, page_left, lines, cell)
            cell_section = _replace(section, page_size=PageSize(max(box.budget, 1), section.page_size.height_twips),
                                    margins=PageMargins(section.margins.top, 0, section.margins.bottom, 0,
                                                        section.margins.header, section.margins.footer, 0),
                                    column_count=1, columns=())
            placed: list[CellParagraph] = []
            items = entries_of.get((row.index, cell.index), []) if cell.merge != "continue" else []
            resolved_items = [resolve_paragraph(document, paragraph) if not isinstance(paragraph, Table) else None
                              for _, paragraph in items]
            for k, (number, paragraph) in enumerate(items):
                if isinstance(paragraph, Table):
                    before = items[k - 1][1] if k and not isinstance(items[k - 1][1], Table) else None
                    placed.append(nested_entry(document, paragraph, cell, page_left, lines, section, advances,
                                               metrics, block=number, labels=labels, features=features, cache=cache,
                                               previous=before))
                    continue
                reason = features[number] if features and number < len(features) else None
                if reason not in (None, "note reference"):
                    raise Unsupported(f"a {reason} in a cell")
                if any(kind == "drawing" for run in paragraph.runs for _, kind in run.breaks):
                    raise Unsupported("a drawing in a cell")
                why = cell_anchors_unmodelled(document, paragraph)
                if why:
                    raise Unsupported(why)
                pp = resolved_items[k]
                label = (labels or {}).get(number)
                pieces, broken = line_model.broken(document, paragraph, cell_section, advances, metrics, label)
                if pieces is None:
                    raise Unsupported("a cell paragraph that cannot be measured")
                if autofit and any(line.emergency for line in broken) and resolved.extents is None and (
                        paragraph_minimum(document, paragraph, cell_section, advances, metrics) > box.budget):
                    # A word wider than its column, which Word would widen.  One wider only
                    # than a first line's room is broken inside, the column kept: a
                    # first-line indent does not count below mode 15 (:func:`paragraph_extent`).
                    raise Unsupported("autofit: a word wider than its column")
                shares = line_model.shares_of(document, paragraph, pieces, broken)
                objects = line_model.objects_of(pieces, broken)
                object_runs = line_model.object_runs_of(pieces, broken)
                heights = []
                for n, chars in enumerate(shares):
                    if objects[n]:
                        height = line_model.object_line_height(document, paragraph, pp, chars, metrics, objects[n],
                                                               first_line=n == 0, runs=object_runs[n])
                    else:
                        height = line_model.line_height(document, paragraph, pp, chars, metrics, first_line=n == 0,
                                                        cache=cache)
                    if height is None:
                        raise Unsupported("no face metrics in a cell")
                    heights.append(height)
                neighbours = [resolved_items[j] if 0 <= j < len(items) else None for j in (k - 1, k + 1)]
                top_border, border = paragraph_borders_px(pp, *neighbours)
                from .paginate import _line_of_sources, anchors_of

                anchors = anchors_of(paragraph, _line_of_sources(pieces, broken), len(broken))
                placed.append(CellParagraph(
                    paragraph, number, pp, pieces, broken,
                    linebreak.geometry(document, paragraph, cell_section, pp), label, heights,
                    before=pp.get("spacing.before", 0) or 0, after=pp.get("spacing.after", 0) or 0,
                    top_border=top_border, border=border, contextual=bool(pp.get("contextualSpacing")),
                    style=paragraph.properties.style_id if paragraph.properties else None,
                    effective_style=paragraph_style_id(document.styles, paragraph),
                    undrawn=frozenset(n for n, chars in enumerate(shares)
                                      if n < len(shares) - 1 and all(char.isspace() for _, char in chars)),
                    anchors=anchors))
            if placed and placed[-1].table is None and hides_mark(cell, placed[-1].paragraph):
                placed[-1].hidden = True
            if (len(placed) > 1 and placed[-2].table is not None and placed[-1].table is None
                    and not placed[-1].pieces):
                # An empty last paragraph after a nested table takes no room: the cell ends
                # at the table's bottom border (make_nested_table_probe.py, whatever its
                # space before).
                placed[-1].hidden = True
            reach = wrap_cell(document, cell, box, placed, previous_style, default_style, page_left, lines, section,
                              advances, metrics, cache)
            flow = CellFlow(cell, box, placed, merged=cell.merge == "continue")
            flow.stack, flow.after = cell_stack(placed, previous_style, default_style)
            flow.content = content_height(placed, previous_style, default_style)
            if reach is not None:
                # The row holds a drawing text wraps around in it (``make_cell_anchor_probe.py``).
                flow.content = max(flow.content, reach - twips_to_px(cell.margins["top"]))
            flows.append(flow)
        rows.append(RowFlow(row, flows))
    columns = len(lines) - 1
    groups = _merge_groups(rows)
    for r, row_flow in enumerate(rows):
        above = [flow.cell for flow in rows[r - 1].cells] if r else []
        row_flow.top_border = _edge_width(above, [flow.cell for flow in row_flow.cells if not flow.merged], columns)
        row_flow.height = row_height(row_flow)
    for (r, flow), last in groups:
        # A merged cell taller than its rows makes the last of them taller.
        flow.rows = last - r + 1
        available = sum((rows[k].height for k in range(r, last + 1)), Fraction(0)) + sum(
            (rows[k].top_border for k in range(r + 1, last + 1)), Fraction(0))
        needed = flow.content + twips_to_px(flow.cell.margins["top"]) + twips_to_px(flow.cell.margins["bottom"])
        if needed > available:
            rows[last].height += needed - available
            available = needed
        flow.span_height = available
    bottom = _edge_width([flow.cell for flow in rows[-1].cells], [], columns) if rows else Fraction(0)
    return TableFlow(table, resolved, rows, bottom, [twips_to_px(page_left + x) for x in lines],
                     previous_style, default_style, page_left, list(lines), floating)


def nested_entry(document: Document, table: Table, cell: Cell, page_left, lines: list, section, advances, metrics,
                 *, block: int, labels, features, cache, previous) -> CellParagraph:
    """A table nested in ``cell`` as its cell's stack holds it: one "line" as tall as the
    table, laid out in the cell's text area as a mode-15 table is in the column, in every
    mode (:attr:`ResolvedTable.nested`).

    Measured by ``make_nested_table_probe.py`` (ROADMAP.md, "Tables -- measured", "A
    table nested in a cell"): its left border's outer edge at the cell's text start (or
    its indent, a negative one none), centred in the text area or its last grid line on
    the text's end; stacked as a paragraph's lines are -- at the cell's content top, below
    the paragraph before it with that paragraph's space after, the paragraph after it
    right below its bottom border with its own space before (:func:`cell_stack`: its own
    space before and after are none) -- and an empty last paragraph after it takes no
    room."""
    from dataclasses import replace as _replace

    text_left = page_left + lines[cell.column] + left_offset(cell)
    text_right = page_left + lines[min(cell.column + cell.span, len(lines) - 1)] - right_offset(cell)
    inner = _replace(section, margins=_replace(section.margins, left=text_left, gutter=0,
                                               right=section.page_size.width_twips - text_right),
                     column_count=1, columns=())
    flow = flow_table(document, table, inner, advances, metrics, block=block, labels=labels, features=features,
                      cache=cache, previous=previous, nested=True)
    height = place_table(flow, Fraction(0), Fraction(10 ** 7), first_on_page=False,
                         mode15=(document.compatibility_mode or 0) >= 15).end
    return CellParagraph(None, block, {}, [], [], None, None, [NestedHeight(height)], undrawn=frozenset({0}),
                         table=flow)


def cell_anchors_unmodelled(document: Document, paragraph) -> str | None:
    """Why a cell paragraph's floating drawings are not laid out (``None`` where they are).

    Measured by ``make_cell_anchor_probe.py`` (ROADMAP.md, "Floating drawings --
    measured", F.16): a drawing text does not wrap around (``wrapNone``) is drawn where
    ``layout._Placer._cell_anchors`` puts it and moves nothing -- not the cell's text,
    not its row, however far past the cell it reaches.  One text wraps around is laid out
    by :func:`wrap_cell` where it is positioned in the cell (F.17).  Below mode 15 one
    positioned against the page that text wraps around -- Word moves the table's rows clear
    of it -- and a position that probe did not measure are not modelled."""
    from .floating import cell_position_modelled, in_cell

    mode15 = (document.compatibility_mode or 0) >= 15
    for run in paragraph.runs:
        for anchor in run.anchors:
            if anchor.moves_text and not in_cell(anchor, mode15):
                return "a floating drawing text wraps around in a cell, positioned against the page"
            if anchor.moves_text and anchor.wrap not in CELL_WRAPS:
                return f"a floating drawing in a cell with {anchor.wrap}"
            why = cell_position_modelled(anchor, mode15)
            if why:
                return f"a floating drawing in a cell positioned by {why}"
    return None


#: The wraps text goes around a drawing in a cell by: beside it, and above and below it.
CELL_WRAPS = ("wrapSquare", "wrapTight", "wrapThrough", "wrapTopAndBottom")


def _stack_tops(paragraphs: list, margin_top: Fraction, previous_style, default_style) -> dict:
    """``(paragraph index, line) -> (top, pitch, paragraph top)`` of a cell's stacked
    lines, px from below the row's top border, as ``layout._Placer._table`` places them:
    a paragraph's top is where the one before ends after its space after (the first
    one's, below the cell's top margin)."""
    stack, _after = cell_stack(paragraphs, previous_style, default_style)
    index = {id(paragraph): k for k, paragraph in enumerate(paragraphs)}
    out = {}
    line_y = paragraph_top = margin_top
    before = None
    for placed in stack:
        if placed.number == 0 and before is not None:
            same = placed.paragraph.style == before.style
            paragraph_top = line_y + twips_to_px(0 if (before.contextual and same) else before.after)
        elif placed.number == 0:
            paragraph_top = line_y
        top = line_y + placed.above
        out[(index[id(placed.paragraph)], placed.number)] = (top, placed.pitch, paragraph_top)
        line_y = top + placed.pitch + placed.below
        before = placed.paragraph
    return out


def _derive(document: Document, paragraph: CellParagraph, source) -> None:
    """A cell paragraph's heights, undrawn lines and anchors' lines, from its lines."""
    from . import lines as line_model
    from .paginate import _line_of_sources, anchors_of

    shares = line_model.shares_of(document, paragraph.paragraph, paragraph.pieces, paragraph.lines)
    heights = []
    for n, (line, chars) in enumerate(zip(paragraph.lines, shares)):
        height = source.height(line, n == 0)
        if height is None:
            raise Unsupported("no face metrics in a cell")
        heights.append(height)
    paragraph.heights = heights
    paragraph.undrawn = frozenset(n for n, chars in enumerate(shares)
                                  if n < len(shares) - 1 and all(char.isspace() for _, char in chars))
    paragraph.anchors = anchors_of(paragraph.paragraph, _line_of_sources(paragraph.pieces, paragraph.lines),
                                   len(paragraph.lines))


def wrap_cell(document: Document, cell: Cell, box: CellBox, placed: list, previous_style, default_style,
              page_left: int, lines: list, section, advances, metrics, cache) -> Fraction | None:
    """Lay a cell's text out around the drawings anchored in it that text wraps around,
    and return how far down the lowest of them reaches (px from below the row's top
    border; ``None`` where there are none).

    Measured by ``make_cell_anchor_probe.py`` (ROADMAP.md, "Floating drawings --
    measured", F.17), for a drawing positioned in the cell (:func:`docx2svg.floating.in_cell`):
    the drawing is positioned from the cell laid out without it, as ``layout`` positions
    one text does not wrap around, and the cell's lines then go beside it
    (``wrapSquare``, ``wrapTight``, ``wrapThrough``) or below it (``wrapTopAndBottom``)
    as the body's go beside and below a drawing on the page (:mod:`docx2svg.wrap`,
    :func:`docx2svg.paginate.clear_of_bands`), in the cell's text area as their column;
    the row reaches down to the drawing's foot, in mode 15 to its ``distB`` below it."""
    from . import floating
    from . import wrap as wrap_model
    from .layout import anchor_character_units, units_px, wrapped_positions
    from .linebreak import Line, break_line, twips_to_units
    from .paginate import Band, ParaSource, clear_of_bands, multiple_extra

    mode15 = (document.compatibility_mode or 0) >= 15
    paragraphs = [paragraph for paragraph in placed if not paragraph.hidden]
    wrapping = [(k, entry) for k, paragraph in enumerate(paragraphs) for entry in paragraph.anchors
                if entry[3].wrap in CELL_WRAPS]
    if not wrapping:
        return None
    if (cell.properties.get("vAlign") or "top") != "top" or cell.merge:
        raise Unsupported("a floating drawing text wraps around in a merged or vertically aligned cell")
    px = floating.TWIP_PX
    margin_top = twips_to_px(cell.margins["top"])
    width, height_twips = section.page_size.width_twips, section.page_size.height_twips
    m = section.margins
    text_left = page_left + lines[cell.column] + left_offset(cell)
    text_right = page_left + lines[min(cell.column + cell.span, len(lines) - 1)] - right_offset(cell)
    box_left = page_left + lines[cell.column] + inner_part(cell.left_edge[0])
    tops = _stack_tops(paragraphs, margin_top, previous_style, default_style)
    wraps, bands = [], []
    foot = Fraction(0)
    for k, (source, run, index, anchor, here) in wrapping:
        paragraph = paragraphs[k]
        top, pitch, paragraph_top = tops[(k, here)]
        line = paragraph.lines[here]
        positions = wrapped_positions(paragraph.pieces, line, paragraph.geometry, advances, first_line=here == 0,
                                      alignment=paragraph.resolved.get("jc"),
                                      last_line=here == len(paragraph.lines) - 1)
        character = twips_to_px(text_left) + units_px(anchor_character_units(
            paragraph.pieces, line, positions, source, paragraph.geometry, first_line=here == 0, run=run))
        frames = floating.Frames(width, height_twips, text_left, width - text_right, 0, height_twips, m.header,
                                 m.footer, 1, mode15, paragraph_top / px, (paragraph_top if here == 0 else top) / px,
                                 (top + pitch) / px, character / px, box_left, 0, cell=True)
        x, y = floating.horizontal(anchor, frames), floating.vertical(anchor, frames)
        if x is None or y is None:
            raise Unsupported("a floating drawing in a cell that cannot be positioned")
        if not mode15:
            # On a whole twip, the nearest, as on the page (``paginate.page_wraps``).
            x = Fraction(math.floor(x + Fraction(1, 2)))
        cx, cy = (floating.emu_twips(v) for v in anchor.extent)
        el, et, er, eb = (floating.emu_twips(v) for v in anchor.effect)
        dt, db, dl, dr = (floating.emu_twips(v) * px for v in anchor.distance)
        paragraph.positioned[(source, run, index)] = (x * px, y * px)
        # The row holds the drawing down to its foot, in mode 15 with its distB too
        # (``make_cell_anchor_probe.py``, ``distB 228600``: 75 px lower in mode 15 only).
        foot = max(foot, (y + cy + eb) * px + (db if mode15 else 0))
        key = (paragraph.block, source, run, index)
        if anchor.wrap == "wrapTopAndBottom":
            bands.append(Band(y * px - dt, (y + cy) * px + db, y * px, key))
            continue
        polygon = None
        if anchor.wrap in ("wrapTight", "wrapThrough") and anchor.polygon:
            polygon = tuple(((x + int(Fraction(cx * a, 21600))) * px, (y + int(Fraction(cy * b, 21600))) * px)
                            for a, b in anchor.polygon)
            if not mode15:
                dt = db = Fraction(0)
        if polygon is not None:
            top_px, bottom_px = min(b for _a, b in polygon) - dt, max(b for _a, b in polygon) + db
        else:
            top_px, bottom_px = (y - et) * px - dt, (y + cy + eb) * px + db
        wraps.append(wrap_model.Wrap(top_px, bottom_px, (x - el) * px, (x + cx + er) * px, dl, dr,
                                     anchor.wrap_text or "bothSides", y * px, key, polygon, dt, db, x * px))
    exact = twips_to_px(text_left)
    column = wrap_model.Column(exact, box.text_left, Fraction(twips_to_units(box.budget)), mode15,
                               twips_to_px(box.budget))
    for k, paragraph in enumerate(paragraphs):
        pieces = paragraph.pieces
        n = len(pieces)
        source = ParaSource(document, paragraph.paragraph, paragraph.resolved, pieces, paragraph.geometry,
                            advances, metrics, cache if cache is not None else {}, True)
        after = twips_to_px(paragraph.after)
        number = 0
        while number < len(paragraph.lines):
            top = _stack_tops(paragraphs, margin_top, previous_style, default_style)[(k, number)][0]
            top -= paragraph.drops.get(number, 0)
            first = number == 0
            start = paragraph.lines[number].start
            plain = Line(n, n) if start >= n else break_line(pieces, start, paragraph.geometry, advances, first=first)
            height = source.height(plain, first)
            if height is None:
                raise Unsupported("no face metrics in a cell")
            last = number == len(paragraph.lines) - 1

            def reach(h, last=last) -> Fraction:
                if mode15:
                    return h.pitch - multiple_extra(h)
                return h.pitch + (after if last else 0)

            y = top
            if bands:
                # Below mode 15 a paragraph's last line reaches down to the end of its space
                # after; in mode 15 to the end of its pitch (as ``paginate._scan``).
                y = clear_of_bands(y, height.pitch + (after if last and not mode15 else 0), bands)
            line = plain
            for _attempt in range(64):
                if not wraps:
                    break
                segs, beside = wrap_model.segments(wraps, y, y + reach(height), column,
                                                   y + height.pitch - multiple_extra(height))
                if segs is None:
                    line = plain
                    break
                found = wrap_model.break_across(pieces, start, segs, paragraph.geometry, advances, first=first)
                if found is not None:
                    line = found
                    break
                below = min(w.bottom for w in beside)
                if below <= y:
                    break
                y = below
            if y != top:
                paragraph.drops[number] = y - top
            else:
                paragraph.drops.pop(number, None)
            old = paragraph.lines[number]
            if isinstance(line, wrap_model.WrappedLine) or (line.start, line.end) != (old.start, old.end):
                tail: list = []
                if line.end < n or (line.forced and n and pieces[n - 1].char == "\n"):
                    tail = source.rest(line.end, first=False)
                paragraph.lines = paragraph.lines[:number] + [line] + tail
                _derive(document, paragraph, source)
            number += 1
    return foot


def hides_mark(cell: Cell, paragraph) -> bool:
    """Whether ``paragraph``, a cell's last, takes no room in its row: the cell carries
    ``w:hideMark`` and the paragraph is empty -- nothing but its mark
    (:func:`docx2svg.paginate.is_empty`: an empty run is nothing, a space or a tab is
    something).

    Measured by ``make_hide_mark_probe.py`` (48 cases, no settings, modes 12, 14 and 15,
    which agree): such a paragraph counts for nothing -- not its line, whatever its mark's
    size (6, 10, 11 or 48 pt), nor its space before or after -- so a row of such cells
    is its ``w:trHeight`` (``atLeast`` or with no ``w:hRule``) plus the cells' margins, and
    with no ``w:trHeight`` its margins alone (0 px without).  Only the last paragraph: of
    two empty paragraphs the first is a line, and so is a no-break space before an empty
    one.  Only its own cell: an empty cell beside it without the element keeps the row a
    line tall.  The same first, last and alone in a table, under ``w:vAlign`` and thick
    inside borders."""
    from .paginate import is_empty

    return bool(cell.properties.get("hideMark")) and is_empty(paragraph)


@dataclass(frozen=True)
class StackedLine:
    """One line of a cell as its stack places it: the room above it that is its own (the
    paragraph's gap and top border, on its first line), its pitch, and the room below
    (its paragraph's bottom border, on its last).  ``box_before`` and ``box_after`` are
    what of the space around it is in its line box, as the body's paragraphs have it
    (:func:`docx2svg.vertical.paragraph_gap_px`): the kept part of the space before on
    the first line, the border and the space after on the last."""

    paragraph: CellParagraph
    number: int
    above: Fraction
    pitch: Fraction
    below: Fraction
    box_before: Fraction = Fraction(0)
    box_after: Fraction = Fraction(0)


def cell_stack(paragraphs: list[CellParagraph], previous_style: str | None = None,
               default_style: str | None = None) -> tuple[list[StackedLine], Fraction]:
    """Each line of a cell's paragraphs, stacked as the body stacks paragraphs -- the
    first one's space before counts, neighbours' spacing collapses to the larger, and
    contextual spacing between two of one style is none -- and the last paragraph's
    space after, which counts too (``make_table_content_probe.py``).

    **Contextual spacing at the cell's edges** (a contextual first or last paragraph):
    the first paragraph loses its space before when it is of the style of the paragraph
    before the table; the last loses its space after when it is of the *default*
    paragraph style, whatever follows the table -- as if the cell ended in a paragraph of
    that style.  Measured on one-cell tables, every pairing of Normal, a contextual style
    and another before and after the table (36 cases)."""
    from .vertical import paragraph_gap_px

    # A hidden end mark (:func:`hides_mark`) is not in the stack at all.
    paragraphs = [paragraph for paragraph in paragraphs if not paragraph.hidden]
    out: list[StackedLine] = []
    previous: CellParagraph | None = None
    for index, paragraph in enumerate(paragraphs):
        if previous is None:
            before = paragraph.before
            if ("content" in STAGES and paragraph.contextual and previous_style is not None
                    and paragraph.effective_style == previous_style):
                before = 0
            gap = box = twips_to_px(before)
        else:
            same = paragraph.style == previous.style
            own_before = 0 if (paragraph.contextual and same) else paragraph.before
            prev_after = 0 if (previous.contextual and same) else previous.after
            gap, box = paragraph_gap_px(prev_after, own_before)
        following = paragraphs[index + 1] if index + 1 < len(paragraphs) else None
        if following is None:
            own_after = 0 if ("content" in STAGES and paragraph.contextual
                              and paragraph.effective_style == default_style) else paragraph.after
        else:
            own_after = 0 if (paragraph.contextual and following.style == paragraph.style) else paragraph.after
        last = len(paragraph.heights) - 1
        for n, height in enumerate(paragraph.heights):
            out.append(StackedLine(
                paragraph, n, ((gap + paragraph.top_border) if n == 0 else Fraction(0)) + paragraph.drops.get(n, 0),
                height.pitch,
                paragraph.border if n == last else Fraction(0),
                box_before=(box + paragraph.top_border) if n == 0 else Fraction(0),
                box_after=(paragraph.border + twips_to_px(own_after)) if n == last else Fraction(0)))
        previous = paragraph
    after = Fraction(0)
    if previous is not None:
        after = twips_to_px(0 if ("content" in STAGES and previous.contextual
                                  and previous.effective_style == default_style) else previous.after)
    return out, after


def content_height(paragraphs: list[CellParagraph], previous_style: str | None = None,
                   default_style: str | None = None) -> Fraction:
    stacked, after = cell_stack(paragraphs, previous_style, default_style)
    return sum((line.above + line.pitch + line.below for line in stacked), Fraction(0)) + after


def _merge_groups(rows: list[RowFlow]) -> list:
    """Every cell merged down (``w:vMerge="restart"``), with the last row its merge
    reaches: the rows below it holding a ``continue`` cell at its grid column."""
    out = []
    for r, row in enumerate(rows):
        for flow in row.cells:
            if flow.cell.merge != "restart":
                continue
            last = r
            while last + 1 < len(rows) and any(
                    other.cell.merge == "continue" and other.cell.column == flow.cell.column
                    for other in rows[last + 1].cells):
                last += 1
            out.append(((r, flow), last))
    return out


def row_height(row: RowFlow) -> Fraction:
    """The row's height between its top border and the edge below.

    Measured by ``make_table_row_probe.py`` (three-row tables, a one-line cell beside a
    three-line one, four settings that agree):

    * with no ``w:trHeight``: the tallest cell, its paragraphs and its top and bottom
      margins;
    * ``atLeast`` (and a height with no ``w:hRule``): at least the stated height *plus
      the cell's top and bottom margins* -- the height is compared with the paragraphs
      alone (margins 100/50 under 1000 twips: a row 240 px, not 208);
    * ``exact``: the stated height *less the border on the row's top edge* -- the table's
      top border for the first row (``w:sz`` 24 under ``exact`` 800: every line exact),
      ``insideH`` below it -- whatever the content, which is cut.
    """
    stated = row.row.properties.get("trHeight")
    value, rule = stated if stated else (0, "atLeast")
    if rule == "exact" and value:
        return twips_to_px(value) - row.top_border
    floor = twips_to_px(value) if value else Fraction(0)
    # A cell merged down counts over all its rows (:func:`flow_table`), not in its first.
    return max((max(flow.content if not flow.cell.merge else Fraction(0), floor)
                + twips_to_px(flow.cell.margins["top"]) + twips_to_px(flow.cell.margins["bottom"])
                for flow in row.cells), default=Fraction(0))


# -- across pages -----------------------------------------------------------------------


@dataclass(frozen=True)
class RowPiece:
    """The part of a row on one page: every cell's stacked lines ``start[c]`` to
    ``end[c]`` (:attr:`CellFlow.stack`), below the edge at ``top`` whose border is
    ``top_border`` tall, ``height`` down to the edge below."""

    row: int
    start: tuple
    end: tuple
    top: Fraction
    top_border: Fraction
    height: Fraction
    #: A header row (``w:tblHeader``) repeated at the top of a page the table continues on.
    header: bool = False
    #: The row began on an earlier page.
    continued: bool = False
    #: The row goes on to the next page.
    split: bool = False
    #: The table as laid out where this row is drawn, when that is not the table's own
    #: flow: beside a drawing text wraps around (:func:`docx2svg.paginate.table_beside`).
    flow: object = None
    #: A cell merged down whose merge goes across a page: ``(row, cell, start, end)`` --
    #: its stacked lines ``start`` to ``end`` (of ``rows[row].cells[cell]``, the merge's
    #: first cell), drawn from this piece's top in its column, running on down through
    #: the merge's rows on the page (:func:`place_table`).  Empty where the whole merge
    #: is on one page: its first row's piece draws it, as before.
    carry: tuple = ()

    @property
    def whole(self) -> bool:
        return not self.continued and not self.split


@dataclass
class TablePage:
    """What of a table goes on one page (:func:`place_table`)."""

    pieces: list
    #: Below the last piece and the border under it.
    end: Fraction
    #: ``(row, per-cell stacked line)`` where the table goes on, or ``None`` when it
    #: ends on this page.  An empty per-cell tuple is the row from its start.
    next: tuple | None


class Overflow(Exception):
    """A row the model cannot place on an empty page (a line taller than the page):
    ``args[0]`` is the row, ``args[1]`` the :class:`TablePage` of
    what was placed before it."""


def header_rows(flow: TableFlow) -> int:
    """How many rows at the table's start are header rows (``w:tblHeader``), repeated on
    every page the table continues on."""
    count = 0
    for row in flow.rows:
        if not row.row.properties.get("tblHeader"):
            break
        count += 1
    return count


def slice_height(cell: CellFlow, start: int, end: int, continued: bool = False) -> Fraction:
    """The height of a cell's stacked lines ``start`` to ``end``, each with its room above
    and below -- also the first on a page the row goes on to: a paragraph starting there
    keeps its space before (``make_table_pages_probe.py``, ``240 before every line``) --
    and the space after the last paragraph where the slice ends the cell."""
    total = Fraction(0)
    for k in range(start, end):
        line = cell.stack[k]
        total += line.above + line.pitch + line.below
    if end == len(cell.stack):
        total += cell.after
    return total


def _edge_below(flow: TableFlow, r: int) -> Fraction:
    return flow.rows[r + 1].top_border if r + 1 < len(flow.rows) else flow.bottom_border


def _margins(cell: CellFlow) -> Fraction:
    return twips_to_px(cell.cell.margins["top"]) + twips_to_px(cell.cell.margins["bottom"])


def _split_counts(flow: TableFlow, row: RowFlow, starts: tuple, room: Fraction, continued: bool,
                  widow_control: bool) -> tuple:
    """Per cell, how many of its lines from ``starts`` fit in ``room`` (below the top
    border and down to the border under the row): lines whole, the cell's margins
    included.  With ``widow_control`` (mode 15), a paragraph split there keeps two lines
    on each side, or moves whole.  A cell merged down counts none here (its lines run
    through the merge's rows: :func:`place_table`)."""
    out = []
    for cell, start in zip(row.cells, starts):
        if _merging(cell):
            out.append(0)
            continue
        count = 0
        while start + count < len(cell.stack) and (
                _margins(cell) + slice_height(cell, start, start + count + 1, continued) <= room):
            count += 1
        if widow_control and 0 < count and start + count < len(cell.stack):
            line = cell.stack[start + count]
            paragraph = line.paragraph
            widow = paragraph.resolved.get("widowControl")
            if widow is None or widow:
                lines = len(paragraph.heights)
                first = start + count - line.number  # the paragraph's first line in the stack
                on_page = start + count - max(first, start)
                left = lines - line.number
                if left < 2 and lines >= 2:
                    count -= 2 - left
                    on_page -= 2 - left
                if 0 < on_page < 2 and first >= start:
                    count -= on_page
        out.append(max(0, count))
    return tuple(out)


def cut_piece(flow: TableFlow, piece: RowPiece, ends: tuple) -> RowPiece:
    """``piece``, a row split at a page's foot, ending instead at each cell's stacked line
    ``ends`` (where the paginator split it), its height theirs."""
    row = flow.rows[piece.row]
    # A merge's cell keeps what it draws here (its lines are the carry's, or its own).
    ends = tuple(own if _merging(cell) else end for cell, own, end in zip(row.cells, piece.end, ends))
    height = max((_margins(cell) + slice_height(cell, start, end, piece.continued)
                  for cell, start, end in zip(row.cells, piece.start, ends) if not _merging(cell)),
                 default=Fraction(0))
    return dataclasses.replace(piece, end=ends, height=max(height, piece.height) if any(
        _merging(cell) for cell in row.cells) else height)


def _merging(cell: CellFlow) -> bool:
    """A cell of a merge down: its first (holding the merge's lines) or one below it."""
    return cell.merged or cell.rows > 1


@dataclass
class _Merge:
    """A cell merged down, as :func:`place_table` places it: the merge's first row and
    cell, its last row, and its column."""

    row: int
    cell: int
    last: int
    column: int

    def flow(self, table: TableFlow) -> CellFlow:
        return table.rows[self.row].cells[self.cell]

    def index_in(self, table: TableFlow, r: int) -> int | None:
        """The merge's cell in row ``r``: its first, or the ``continue`` one at its column."""
        if r == self.row:
            return self.cell
        for k, cell in enumerate(table.rows[r].cells):
            if cell.merged and cell.cell.column == self.column:
                return k
        return None


def _merges(flow: TableFlow) -> list[_Merge]:
    out = []
    for (r, cell), last in _merge_groups(flow.rows):
        if last > r:
            out.append(_Merge(r, flow.rows[r].cells.index(cell), last, cell.cell.column))
    return out


def _fitting(cell: CellFlow, start: int, top: Fraction, limit: Fraction, continued: bool) -> int:
    """How many of a cell's stacked lines from ``start`` end at or above ``limit`` when the
    first starts at ``top``."""
    end = start
    while end < len(cell.stack) and top + slice_height(cell, start, end + 1, continued) <= limit:
        end += 1
    return end


def place_table(flow: TableFlow, y: Fraction, bottom: Fraction, *, row: int = 0, part: tuple = (),
                first_on_page: bool, mode15: bool, stop: int | None = None) -> TablePage:
    """See :func:`_place_table`; a table whose merges down all lie on the page, or that
    has none, is placed as stage 6 measured, unchanged."""
    merges = _merges(flow)
    if not merges:
        return _place_table(flow, y, bottom, row=row, part=part, first_on_page=first_on_page, mode15=mode15,
                            stop=stop)
    return _place_table(flow, y, bottom, row=row, part=part, first_on_page=first_on_page, mode15=mode15,
                        stop=stop, merges=merges)


def _place_table(flow: TableFlow, y: Fraction, bottom: Fraction, *, row: int = 0, part: tuple = (),
                 first_on_page: bool, mode15: bool, stop: int | None = None, merges: list | None = None) -> TablePage:
    """What of a table goes on the page from ``y``, starting at ``row`` (from each cell's
    stacked line ``part`` when the row began on an earlier page), with the page's text
    ending at ``bottom``.

    Measured by ``make_table_pages_probe.py``:

    * a row stays on the page when it ends, with the border under it, at or above the
      page's foot (inclusive);
    * a row that does not is **split** at the foot: each cell keeps the lines that end,
      with its bottom margin and that border, above it, and the rest go on at the top
      of the next page -- under the row's top border and its cells' top margins again,
      a paragraph starting there with its space before; in mode 15 a split paragraph
      keeps two lines on each side (widow control), in modes 12 and 14 and unstated it
      may leave one; where a cell with lines left would keep none, the row moves whole;
    * a ``w:cantSplit`` row moves whole -- and one taller than a page, at a page's top,
      is split in mode 15 and runs off the page below it;
    * on a page the table continues on, its header rows (``w:tblHeader`` from the first)
      come first.

    Measured by ``make_table_foot_probe.py``:

    * **a header row is never split** at the foot: it moves whole, as a ``w:cantSplit``
      row does; and **in mode 15 the table's header rows are not left at a page's foot
      with no row after them** -- the table moves to the next page (with no
      ``settings.xml`` they stay, and are repeated over the next page's rows);
    * **a row holding a cell merged down splits as any row does**: its own cells as
      above, and the merged cell's lines running on through the merge's rows -- those
      that end above the page's break stay, the rest go on at the next page's top, from
      the top of the merge's first row there (:attr:`RowPiece.carry`).  The merge's last
      row is as tall as its own cells or the merged cell's lines still to place,
      whichever is taller, on each page.

    With ``stop``, the rows before it only: the page's part of them, and ``(stop, ())``
    next with no bottom edge where they all fit.
    """
    rows = flow.rows
    pieces: list[RowPiece] = []
    header_count = header_rows(flow)
    merges = merges or []
    y0 = y
    #: Per merge met on this page (by ``id``): :class:`_OnPage`.
    state: dict[int, _OnPage] = {}
    if first_on_page and (row > 0 or part) and 0 < header_count <= row:
        for h in range(header_count):
            header = rows[h]
            cells = tuple(len(cell.stack) for cell in header.cells)
            pieces.append(RowPiece(h, tuple(0 for _ in cells), cells, y, header.top_border, header.height,
                                   header=True))
            y += header.top_border + header.height

    def finish(end: Fraction, after: tuple | None) -> TablePage:
        """The page as placed: each merge on it given the lines that end above the last
        piece where its end is not known yet, and its carry where it goes across the
        page."""
        for m in merges:
            found = state.get(id(m))
            if found is not None and found.end is None:
                merged = m.flow(flow)
                last = pieces[-1] if pieces else None
                limit = (last.top + last.top_border + last.height if last else y0) - twips_to_px(
                    merged.cell.margins["bottom"])
                found.end = _fitting(merged, found.start, found.top, limit, True)
        return TablePage(_carried(flow, pieces, merges, state), end, after)

    def leave(after: tuple) -> TablePage:
        """Nothing more of the table on this page: it goes on at ``after``."""
        if mode15 and header_count and pieces and not first_on_page and all(
                piece.row < header_count and not piece.header for piece in pieces):
            # Mode 15: the table's header rows are not left at the foot with no row after
            # them; the table starts the next page (``make_table_foot_probe.py``).
            return TablePage([], y0, (0, ()))
        return finish(y + (_edge_below(flow, pieces[-1].row) if pieces else Fraction(0)), after)

    while row < len(rows):
        if stop is not None and row >= stop:
            return finish(y, (row, ()))
        current = rows[row]
        ends = tuple(len(cell.stack) for cell in current.cells)
        continued = bool(part)
        starts = part or tuple(0 for _ in current.cells)
        content_top = y + current.top_border
        active = [m for m in merges if m.row <= row <= m.last]
        for m in active:
            if id(m) not in state:
                index = m.index_in(flow, row)
                if continued and index is not None:
                    start = starts[index]
                elif row == m.row:
                    start = 0
                else:
                    start = _placed_before(flow, m, row)
                state[id(m)] = _OnPage(start, content_top + twips_to_px(m.flow(flow).cell.margins["top"]),
                                       len(pieces))
        if continued:
            height = max((_margins(cell) + slice_height(cell, start, len(cell.stack), True)
                          for cell, start in zip(current.cells, starts) if not _merging(cell)), default=Fraction(0))
        else:
            height = row_height(current) if active else current.height
        for m in active:
            if row == m.last:
                merged, found = m.flow(flow), state[id(m)]
                height = max(height, found.top + slice_height(merged, found.start, len(merged.stack), True)
                             + twips_to_px(merged.cell.margins["bottom"]) - content_top)
        edge = _edge_below(flow, row)
        nothing_yet = first_on_page and not any(not piece.header for piece in pieces)
        if y + current.top_border + height + edge <= bottom:
            pieces.append(RowPiece(row, starts, ends, y, current.top_border, height, continued=continued))
            y += current.top_border + height
            for m in active:
                if row == m.last:
                    state[id(m)].end = len(m.flow(flow).stack)
            row, part = row + 1, ()
            continue
        if "pages" not in STAGES:
            raise Overflow(row, TablePage(pieces, y, (row, part)))
        header = row < header_count and not continued
        if (current.row.properties.get("cantSplit") or header) and not (nothing_yet and mode15):
            if nothing_yet:
                pieces.append(RowPiece(row, starts, ends, y, current.top_border, height, continued=continued))
                y += current.top_border + height
                row, part = row + 1, ()
                continue
            return leave((row, part))
        room = bottom - edge - y - current.top_border
        counts = _split_counts(flow, current, starts, room, continued, mode15)
        if any(count == 0 and start < len(cell.stack)
               for count, start, cell in zip(counts, starts, current.cells) if not _merging(cell)):
            # A cell that would keep none of its lines holds the whole row back
            # (mode 15: a six-line paragraph with room for one line moves, and the one-line
            # cell beside it with it).
            counts = tuple(0 for _ in counts)
        split_ends = tuple(start + count for start, count in zip(starts, counts))
        split_height = max((_margins(cell) + slice_height(cell, start, end, continued)
                            for cell, start, end in zip(current.cells, starts, split_ends) if not _merging(cell)),
                           default=Fraction(0))
        # The merges' lines that end above the page's break: within this row's part for a
        # merge that goes on below the row, down to the page's foot for one the row ends.
        kept = {}
        for m in active:
            merged, found = m.flow(flow), state[id(m)]
            margin = twips_to_px(merged.cell.margins["bottom"])
            limit = (bottom - edge if row == m.last else content_top + split_height) - margin
            kept[id(m)] = _fitting(merged, found.start, found.top, limit, True)
            if row == m.last and kept[id(m)] > found.start:
                split_height = max(split_height, found.top + slice_height(merged, found.start, kept[id(m)], True)
                                   + margin - content_top)
        # The merge's last row goes on to this page only for lines that reach below its
        # top: those above it are in the rows above.
        grows = any(row == m.last and kept[id(m)] > _fitting(
            m.flow(flow), state[id(m)].start, state[id(m)].top, y - twips_to_px(m.flow(flow).cell.margins["bottom"]),
            True) for m in active)
        if not any(counts) and not grows:
            if nothing_yet:
                raise Overflow(row, TablePage(pieces, y, (row, part)))
            return leave((row, part))
        pieces.append(RowPiece(row, starts, split_ends, y, current.top_border, split_height, continued=continued,
                               split=True))
        y += current.top_border + split_height
        for m in active:
            state[id(m)].end = kept[id(m)]
        return finish(y + edge, (row, _next_part(flow, row, split_ends, active, state)))
    return finish(y + flow.bottom_border, None)


@dataclass
class _OnPage:
    """A merge on one page: the first of its lines here, where that line's top is, the
    index of its first piece here, and the line after its last here (``None`` until
    known)."""

    start: int
    top: Fraction
    first: int
    end: int | None = None


def _next_part(flow: TableFlow, row: int, ends: tuple, active: list, state: dict) -> tuple:
    """Where row ``row`` goes on: each own cell's next stacked line, and each merge's next
    line in its cell of the row."""
    out = list(ends)
    for m in active:
        index = m.index_in(flow, row)
        if index is not None:
            out[index] = state[id(m)].end
    return tuple(out)


def _placed_before(flow: TableFlow, merge: _Merge, row: int) -> int:
    """How many of a merge's lines its rows above ``row`` hold, laid out one after
    another: where a page starts at a row inside the merge, the rows before it were
    placed whole on the page before (a merge going across three pages, not measured,
    is taken as if its rows above were on one)."""
    merged = merge.flow(flow)
    region = sum((row_height(flow.rows[k]) for k in range(merge.row, row)), Fraction(0)) + sum(
        (flow.rows[k].top_border for k in range(merge.row + 1, row)), Fraction(0))
    limit = region - twips_to_px(merged.cell.margins["bottom"])
    return _fitting(merged, 0, twips_to_px(merged.cell.margins["top"]), limit, False)


def _carried(flow: TableFlow, pieces: list, merges: list, state: dict) -> list:
    """``pieces`` with each merge that does not lie whole on this page given its lines
    here, on the first of its pieces here (:attr:`RowPiece.carry`); its first row, where
    that is here, then draws none of them itself."""
    out = list(pieces)
    for m in merges:
        found = state.get(id(m))
        if found is None or found.first >= len(out):
            continue
        piece = out[found.first]
        end = found.end if found.end is not None else found.start
        whole = piece.row == m.row and not piece.continued and found.start == 0 and end == len(m.flow(flow).stack)
        if whole:
            # The merge's lines all lie on this page: drawn by its first row, as before --
            # also where that row is split here.
            if piece.end[m.cell] != end:
                ends = list(piece.end)
                ends[m.cell] = end
                out[found.first] = dataclasses.replace(piece, end=tuple(ends))
            continue
        if piece.row == m.row:
            starts, ends = list(piece.start), list(piece.end)
            starts[m.cell] = ends[m.cell] = found.start
            piece = dataclasses.replace(piece, start=tuple(starts), end=tuple(ends))
        out[found.first] = dataclasses.replace(piece, carry=piece.carry + ((m.row, m.cell, found.start, end),))
    return out
