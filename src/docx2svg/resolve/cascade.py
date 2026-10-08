"""The property cascade: what a paragraph, a run or a paragraph mark resolves to, and why.

Every value comes back with its **origin** -- the level of the hierarchy that decided it
and, for a style, which style in the ``w:basedOn`` chain declared it.  That is the point
of this module as much as the values are: a later defect ("this line is 1 px low") is
diagnosable only if one can ask where the 11 pt on that line came from.  See
:meth:`Resolved.explain`.

The hierarchy, as Word 16.106 applies it
----------------------------------------

Measured by ``tools/make_style_probe.py`` / ``tools/read_style_probe.py``; ROADMAP.md,
"Style inheritance", has the cases.  Lowest precedence first:

1. ``w:docDefaults`` (``w:rPrDefault`` / ``w:pPrDefault``);
2. the **table style** of the innermost table, for paragraphs in its cells (the default
   table style when the table names none), followed by its conditional formats for the
   cell (``w:tblStylePr``: banding, first/last row and column, corners), each by override
   -- a real document (``sample-simple.docx``) draws the cell where a bold first row
   meets a bold first column bold, so they do not toggle one another;
3. **numbering** -- the list level's ``w:pPr`` -- *here* when the ``w:numPr`` comes from
   the paragraph style;
4. the **paragraph style** (the default paragraph style when none is named, or when the
   name matches no paragraph style);
5. **numbering** again -- *here* when the ``w:numPr`` is direct formatting, which is why
   a directly numbered paragraph takes the list's indent over its style's (case
   ``p02``).  ECMA-376 17.7.2 puts numbering below the paragraph style unconditionally;
6. the **character style** (runs only).  The *default* character style contributes
   nothing, not even through ``w:basedOn`` (cases ``g01``, ``g09``), and an ``rStyle``
   naming a style that is not a character style is ignored (``t22``);
7. **direct formatting**.

Within one style's ``w:basedOn`` chain every property, toggles included, is a plain
override: the style closest to the one applied wins.

Toggle properties -- where Word is not ECMA-376
-----------------------------------------------

ECMA-376 17.7.3 says a toggle property (``w:b``, ``w:i``, ``w:caps``, ``w:vanish``...)
in a style *flips* the value accumulated so far and ``false`` leaves it unchanged; only
direct formatting sets it absolutely.  What Word does, from 30 + 33 cases:

* each style **level** -- table, paragraph, character -- first resolves its own value
  through its ``w:basedOn`` chain by override (a bold style based on a bold style is
  bold, ``t08``; ``w:b w:val="0"`` based on a bold style is not bold, ``t09``);
* then, starting from the ``w:docDefaults`` value (false when absent), **every level
  whose value differs from the docDefaults value flips the result**.  With docDefaults
  silent that is the spec's XOR -- paragraph bold plus character bold is *not* bold
  (``t03``), three levels bold is bold (``t29``), ``false`` changes nothing (``t11``).
  With docDefaults bold it is not: a bold paragraph style plus a bold character style
  stay bold (``g07``), and a style saying ``false`` turns bold *off* (``g10``, ``g11``);
* direct formatting sets the value absolutely (``t04``, ``t05``, ``t07``, ``g04``).

The same rule holds for italic (``t15``), caps (``t17``) and vanish (``t20``); the other
toggles are assumed to follow it and have not been drawn.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from ..model import Document, NumberingLevel, Paragraph, Run, Style, StyleSheet
from ..parse.properties import TOGGLES

#: Level names, in precedence order (lowest first).
DOC_DEFAULTS = "docDefaults"
TABLE = "table style"
NUMBERING = "numbering"
PARAGRAPH = "paragraph style"
CHARACTER = "character style"
DIRECT = "direct"
APPLICATION = "application default"
THEME = "theme"


@dataclass(frozen=True)
class Origin:
    """Where one value came from."""

    level: str
    #: The style applied at this level (for a style level).
    style_id: str | None = None
    #: The style in the ``w:basedOn`` chain that actually declared the value, when it is
    #: not ``style_id`` itself.
    declared_in: str | None = None
    #: Anything else worth saying: a numbering level, a theme slot, a toggle's effect.
    detail: str | None = None

    def __str__(self) -> str:
        text = self.level
        if self.style_id:
            text += f" {self.style_id!r}"
            if self.declared_in and self.declared_in != self.style_id:
                text += f" (from {self.declared_in!r})"
        if self.detail:
            text += f" [{self.detail}]"
        return text


@dataclass(frozen=True)
class Resolved:
    """Resolved properties plus, for each, the origins that decided it.

    ``origins[key]`` is a tuple: its *last* entry is the level that decided a plain
    property; for a toggle it lists every level that declared the property, each
    marked with whether it flipped the value.
    """

    values: dict = field(default_factory=dict)
    origins: dict = field(default_factory=dict)

    def __getitem__(self, key: str):
        return self.values[key]

    def __contains__(self, key: str) -> bool:
        return key in self.values

    def get(self, key: str, default=None):
        return self.values.get(key, default)

    def origin(self, key: str) -> Origin | None:
        chain = self.origins.get(key)
        return chain[-1] if chain else None

    def explain(self, keys=None) -> str:
        """One line per property: ``b = False  <- paragraph style 'PB' [flips] ...``."""
        lines = []
        for key in sorted(self.values if keys is None else keys):
            if key not in self.values:
                lines.append(f"{key} = (unset)")
                continue
            chain = self.origins.get(key, ())
            lines.append(f"{key} = {self.values[key]!s}  <- " + "; ".join(map(str, chain)))
        return "\n".join(lines)


# -- styles --------------------------------------------------------------------------


def style_chain(sheet: StyleSheet | None, style_id: str | None, kind: str) -> list[Style]:
    """The ``w:basedOn`` chain of ``style_id``, root first, stopping at a cycle.

    Returns ``[]`` for an id that names no style of ``kind``.  For ``character`` the
    default character style is dropped from the chain: Word applies nothing it declares
    (cases ``g01`` and ``g09`` of the probe).
    """
    if sheet is None or style_id is None:
        return []
    chain: list[Style] = []
    seen: set[str] = set()
    current = sheet.styles.get(style_id)
    if current is None or current.kind != kind:
        return []
    while current is not None and current.style_id not in seen and current.kind == kind:
        seen.add(current.style_id)
        chain.append(current)
        current = sheet.styles.get(current.based_on) if current.based_on else None
    chain.reverse()
    if kind == "character":
        chain = [style for style in chain if not style.default]
    return chain


def _paragraph_style_id(sheet: StyleSheet | None, paragraph: Paragraph) -> str | None:
    declared = paragraph.properties.style_id if paragraph.properties else None
    if sheet is None:
        return None
    if declared is not None:
        style = sheet.styles.get(declared)
        if style is not None and style.kind == "paragraph":
            return declared
    # No pStyle, or one naming no paragraph style (case t23): the default applies.
    default = sheet.default("paragraph")
    return default.style_id if default else None


def _table_style_id(sheet: StyleSheet | None, paragraph: Paragraph) -> str | None:
    if paragraph.table_style_id is None or sheet is None:
        return None
    if paragraph.table_style_id:
        style = sheet.styles.get(paragraph.table_style_id)
        if style is not None and style.kind == "table":
            return paragraph.table_style_id
    default = sheet.default("table")
    return default.style_id if default else None


def banded(chain: list[Style], conditions: tuple[str, ...]) -> tuple[str, ...]:
    """``conditions`` less the banding ones the table style does not turn on: row bands
    (``band1Horz`` / ``band2Horz``) only where its chain states ``w:tblStyleRowBandSize``,
    column bands only where it states ``w:tblStyleColBandSize``.  Measured by
    ``make_table_border_probe.py`` (``style``): a table style with a ``band1Horz`` shading
    and no band size draws no band; Word's own table styles all state a size of 1.
    (A size above 1 -- bands of several rows -- is not modelled: bands are one row.)"""
    from ..table import STAGES

    if "borders" not in STAGES:
        return conditions
    rows = any((style.table.get("tblStyleRowBandSize") or 0) > 0 for style in chain)
    columns = any((style.table.get("tblStyleColBandSize") or 0) > 0 for style in chain)
    return tuple(c for c in conditions if not ((c.endswith("Horz") and not rows) or (c.endswith("Vert") and not columns)))


# -- merging -------------------------------------------------------------------------

_Layer = dict  # key -> (value, Origin)


def _flatten(chain: list[Style], level: str, part: str, applied: str,
             conditions: tuple[str, ...] = ()) -> _Layer:
    """One style level: its chain merged root-first by override, origins kept.

    For a table style, each style's conditional formats (``w:tblStylePr``) that apply to
    the cell follow its own properties, in ``conditions`` order, also by override.
    """
    layer: _Layer = {}
    index = 0 if part == "paragraph" else 1
    for style in chain:
        declared = style.paragraph if part == "paragraph" else style.run
        _merge(layer, {k: (v, Origin(level, applied, style.style_id)) for k, v in declared.items()})
        for condition in conditions:
            if condition in style.conditional:
                declared = style.conditional[condition][index]
                origin = Origin(level, applied, style.style_id, condition)
                _merge(layer, {k: (v, origin) for k, v in declared.items()})
    return layer


def _layer(declared: dict, origin: Origin) -> _Layer:
    return {key: (value, origin) for key, value in declared.items()}


def _merge(accumulated: _Layer, layer: _Layer) -> None:
    """Override ``accumulated`` with ``layer``, with the two non-plain properties:

    * ``ind.firstLine`` and ``ind.hanging`` are one quantity written two ways; a level
      declaring either removes the other as inherited;
    * ``tabs`` accumulate: a level adds its stops, and a ``clear`` stop removes the
      inherited stop at that position.
    """
    for key, (value, origin) in layer.items():
        if key == "ind.firstLine":
            accumulated.pop("ind.hanging", None)
        elif key == "ind.hanging":
            accumulated.pop("ind.firstLine", None)
        if key == "tabs":
            stops = {stop.position_twips: stop for stop in accumulated.get("tabs", ((), None))[0]}
            for stop in value:
                if stop.alignment == "clear":
                    stops.pop(stop.position_twips, None)
                else:
                    stops[stop.position_twips] = stop
            value = tuple(sorted(stops.values(), key=lambda stop: stop.position_twips))
        accumulated[key] = (value, origin)


def _finish(layer: _Layer, toggles: dict | None = None) -> Resolved:
    values = {key: value for key, (value, _) in layer.items()}
    origins = {key: (origin,) for key, (_, origin) in layer.items()}
    for key, (value, chain) in (toggles or {}).items():
        values[key] = value
        origins[key] = chain
    return Resolved(values, origins)


# -- numbering -----------------------------------------------------------------------


def _numbering_level(document: Document, num_id, ilvl) -> NumberingLevel | None:
    if not num_id:  # numId 0 removes numbering
        return None
    return document.numbering.get(num_id, {}).get(ilvl or 0)


# -- the three entry points ------------------------------------------------------------


def resolve_paragraph(document: Document, paragraph: Paragraph) -> Resolved:
    """The paragraph's own properties (``w:pPr``), through the whole hierarchy."""
    sheet = document.styles
    layer: _Layer = {}
    if sheet is not None:
        _merge(layer, _layer(sheet.paragraph_defaults, Origin(DOC_DEFAULTS)))

    table_id = _table_style_id(sheet, paragraph)
    if table_id:
        chain = style_chain(sheet, table_id, "table")
        _merge(layer, _flatten(chain, TABLE, "paragraph", table_id, banded(chain, paragraph.table_conditions)))

    style_id = _paragraph_style_id(sheet, paragraph)
    style_layer = _flatten(style_chain(sheet, style_id, "paragraph"), PARAGRAPH, "paragraph", style_id)
    direct = paragraph.properties.declared if paragraph.properties else {}

    # Where the list's indent sits depends on where the numPr came from (case p02/p03).
    direct_numbering = "numPr.numId" in direct
    source = direct if direct_numbering else {k: v for k, (v, _) in style_layer.items()}
    level = _numbering_level(document, source.get("numPr.numId"),
                             direct.get("numPr.ilvl", source.get("numPr.ilvl")))
    numbering_layer = {}
    if level is not None:
        detail = f"numId {source.get('numPr.numId')} level {level.level}"
        numbering_layer = _layer(level.paragraph, Origin(NUMBERING, detail=detail))

    if not direct_numbering:
        _merge(layer, numbering_layer)
    _merge(layer, style_layer)
    if direct_numbering:
        _merge(layer, numbering_layer)
    _merge(layer, _layer(direct, Origin(DIRECT)))
    _autospacing(document, layer)
    return _finish(layer)


#: What ``w:beforeAutospacing`` / ``w:afterAutospacing`` make the space: 14 pt.
AUTOSPACING_TWIPS = 280


def _autospacing(document: Document, layer: _Layer) -> None:
    """``w:beforeAutospacing`` / ``w:afterAutospacing`` replace the space they sit beside
    with **14 pt, whatever is stated**, at every size and line rule -- unless the settings
    turn on ``w:doNotUseHTMLParagraphAutoSpacing``, under which the stated value holds.

    Measured by ``tools/make_autospacing_probe.py``: stated 0, 100, 400 and 1000 twips at
    8, 11, 20 and 36 pt, single, 1.5 and ``exact`` lines, beside neighbours of 0..1000
    twips, with no settings, in mode 14 and in mode 15 (ROADMAP.md, "Autospacing --
    measured").  The flag's origin is the value's origin.
    """
    if "doNotUseHTMLParagraphAutoSpacing" in document.compat_options:
        return
    for side in ("before", "after"):
        flag = layer.get(f"spacing.{side}Autospacing")
        if flag and flag[0]:
            origin = flag[1]
            layer[f"spacing.{side}"] = (AUTOSPACING_TWIPS, Origin(
                origin.level, origin.style_id, origin.declared_in,
                f"w:{side}Autospacing: 14 pt, whatever is stated"))


def autospaced(document: Document, resolved: Resolved, side: str) -> bool:
    """Whether ``side`` (``"before"`` / ``"after"``) of a resolved paragraph is autospaced
    -- its flag is on and the settings do not turn autospacing off -- so that its 14 pt
    is subject to :func:`docx2svg.vertical.autospace_kept` against its neighbour."""
    return (bool(resolved.get(f"spacing.{side}Autospacing"))
            and "doNotUseHTMLParagraphAutoSpacing" not in document.compat_options)


def _resolve_run_properties(document: Document, paragraph: Paragraph, character_style: str | None,
                            direct: dict) -> Resolved:
    sheet = document.styles
    defaults = sheet.run_defaults if sheet is not None else {}

    levels: list[_Layer] = []
    table_id = _table_style_id(sheet, paragraph)
    if table_id:
        chain = style_chain(sheet, table_id, "table")
        levels.append(_flatten(chain, TABLE, "run", table_id, banded(chain, paragraph.table_conditions)))
    style_id = _paragraph_style_id(sheet, paragraph)
    paragraph_level = _flatten(style_chain(sheet, style_id, "paragraph"), PARAGRAPH, "run", style_id)
    if table_id and legacy_table_size(document, table_id):
        table_level = levels[0]

        def gives_way(key: str, value) -> bool:
            if key not in ("sz", "szCs") or value[0] != LEGACY_DEFAULT_SIZE:
                return False
            # The table style's size, or else w:docDefaults' -- unless that is 10 pt,
            # which counts as unset too, and then the 12 pt stays.
            return key in table_level or defaults.get(key) not in (None, LEGACY_UNSET_DEFAULT)

        paragraph_level = {key: value for key, value in paragraph_level.items() if not gives_way(key, value)}
    levels.append(paragraph_level)
    if character_style:
        levels.append(_flatten(style_chain(sheet, character_style, "character"), CHARACTER, "run",
                               character_style))

    layer: _Layer = _layer({k: v for k, v in defaults.items() if k not in TOGGLES},
                           Origin(DOC_DEFAULTS))
    for level in levels:
        _merge(layer, {k: v for k, v in level.items() if k not in TOGGLES})
    _merge(layer, _layer({k: v for k, v in direct.items() if k not in TOGGLES}, Origin(DIRECT)))

    toggles: dict = {}
    declared = set(defaults) | set(direct) | {key for level in levels for key in level}
    for key in sorted(declared & TOGGLES):
        base = bool(defaults.get(key, False))
        value = base
        chain = [Origin(DOC_DEFAULTS, detail=f"{base}")] if key in defaults else []
        for level in levels:
            if key in level:
                declared_value, origin = level[key]
                flips = bool(declared_value) != base
                if flips:
                    value = not value
                chain.append(Origin(origin.level, origin.style_id, origin.declared_in,
                                    f"{declared_value}, {'flips' if flips else 'no flip'}"))
        if key in direct:
            value = bool(direct[key])
            chain.append(Origin(DIRECT, detail=f"{value}, absolute"))
        toggles[key] = (value, tuple(chain))
    return _finish(layer, toggles)


#: The size Word treated as "no size stated" before Word 2013, in half points: 12 pt.
LEGACY_DEFAULT_SIZE = 24
#: A ``w:docDefaults`` size that counts as unset under the same rule: 10 pt, ECMA-376's
#: default (measured: over ``w:docDefaults`` of 8, 9, 10.5, 11, 11.5, 13 and 14 pt a 12 pt
#: Normal gives way to them in a silent table style's cells; over 10 pt it stays 12 pt).
LEGACY_UNSET_DEFAULT = 20


def legacy_table_size(document: Document, table_id: str) -> bool:
    """Whether a table cell's text in a 12 pt paragraph style takes the table style's size
    instead (:data:`LEGACY_DEFAULT_SIZE`).

    **Finding 6**, measured by ``make_table_style_size_probe.py`` and a sweep of the
    Normal style's and ``w:docDefaults``' sizes: with no mode stated and in modes 12 and
    14, unless the ``overrideTableStyleFontSizeAndJustification`` compatibility setting
    is on, in a table whose style states any paragraph or run property (a ``w:spacing``,
    a ``w:color``, a size, an alignment), a paragraph whose styles give it **12 pt**
    takes the size the table style gives it -- the table style's own, or where it states
    none, ``w:docDefaults``' (unless that is 10 pt: :data:`LEGACY_UNSET_DEFAULT`) -- as
    though 12 pt, the pre-2007 default, meant "not set".
    Any other size the paragraph style gives stays: Normal at 11, 10 or 14 pt beats a
    table style's 9 or 15 pt, and a table style's 9 pt beats Normal's 12 pt (over
    ``w:docDefaults`` of 10 or 11 pt alike); a paragraph style stating 16 pt keeps it.
    A table style stating only table properties changes nothing, and nor does mode 15.
    The style probe's p06-p08 (Normal 11 pt, a table style's 14 pt: Normal wins) are this
    rule's other side.  Alignment is not affected by a centred Normal (it centres the
    cells of every table style); whether a stated *left* alignment gives way is not
    measured.  (``filesamples/sample1``'s Normal is 12 pt and its ``TableGrid`` states a
    line spacing: its cells are ``w:docDefaults``' 11 pt.)
    """
    from ..table import STAGES

    if "content" not in STAGES:
        return False
    mode = document.compatibility_mode
    if mode is not None and mode >= 15:
        return False
    if document.compat_settings.get("overrideTableStyleFontSizeAndJustification", "0") in ("1", "true", "on"):
        return False
    return any(style.paragraph or style.run for style in style_chain(document.styles, table_id, "table"))


def resolve_run(document: Document, paragraph: Paragraph, run: Run) -> Resolved:
    """A run's properties (``w:rPr``), through the whole hierarchy."""
    properties = run.properties
    return _resolve_run_properties(
        document, paragraph,
        properties.style_id if properties else None,
        properties.declared if properties else {},
    )


def resolve_mark(document: Document, paragraph: Paragraph) -> Resolved:
    """The paragraph mark's run properties.

    The mark is resolved like a run with no character style: docDefaults, table style,
    paragraph style, then the ``w:pPr/w:rPr`` written for it.  It matters for layout --
    it is the height of an empty line (``vertical.line_items``) and the properties a list
    label starts from -- and Word draws it, as a space, in its PDF.
    """
    mark = paragraph.properties.run_properties if paragraph.properties else None
    return _resolve_run_properties(document, paragraph, None, mark.declared if mark else {})
