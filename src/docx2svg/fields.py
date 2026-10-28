"""The text of a computed field: ``PAGE``, ``NUMPAGES``, ``SECTIONPAGES``, ``PAGEREF``,
``REF`` and ``SEQ``.

The parser keeps a computed field as one run (:class:`docx2svg.model.Field`), its text
the result Word cached; the layout asks :func:`field_text` for the text on a page from its
own pagination.  Every rule is Word's as ``make_field_probe.py`` measured it (sixteen
sections of page-number formats and switches, each field cached as a value Word could
not have computed):

* ``PAGE`` is the page's number in ``w:pgNumType/@w:fmt`` -- unless a switch formats it;
  ``NUMPAGES`` (the pages of the document, blank ones counted) and ``SECTIONPAGES`` (the
  pages the section has content on) are arabic whatever the section's format;
* ``\\* roman`` / ``Roman`` (the switch's first letter's case is the result's),
  ``alphabetic`` / ``ALPHABETIC``, ``Arabic``, ``ArabicDash`` (``- 4 -``), ``Hex``
  (upper case) and ``\\# "000"`` override the section's format (a picture's ``#`` is a
  space where the number has no digit: ``\\# "#0"`` draws 2 as a space and 2); ``MERGEFORMAT`` and
  ``CHARFORMAT`` do not (they choose the result's run format: the parser's); ``Upper``,
  ``Lower``, ``FirstCap`` and ``Caps`` change case after;
* ``ordinal``, ``cardinalText`` and ``ordinalText`` (and the switches ``Ordinal``,
  ``CardText``, ``OrdText``, ``DollarText``) are words in **the application's language**,
  not the document's: an ``en-GB`` document is drawn ``1e``, ``een``, ``eerste`` by a
  Dutch Word.  The file cannot say which, so they are not computed (:data:`LANGUAGE`).

**Cross-references**, in the body (``make_xref_field_probe.py``: every case, both
settings), each formatted as a ``PAGE`` is unless said:

* ``PAGEREF name`` is the page the bookmark *starts* on -- the line its start stands in,
  a paragraph that runs on to the next page giving the later page -- as ``PAGE`` shows
  that page: its number in **its own section's** format (``xi`` for a bookmark on the
  second page of a ``lowerRoman`` section from 10, from a decimal section), switches
  applied; nested in a table of contents' result too.  ``\\p`` ("above", "on page 3") is
  the application's language, and a missing bookmark Word's error text in it: not
  computed.
* ``SEQ id`` counts the body's ``SEQ`` fields of that identifier -- case ignored
  (``figure`` counts with ``Figure``) -- in document order, table cells in their place:
  ``\\n`` (or nothing) adds one, ``\\c`` repeats the last, ``\\r n`` sets ``n``, ``\\h``
  adds one and draws nothing; ``\\*`` formats.  ``\\s`` (restart at a heading level) is
  not computed, nor is any field of an identifier that has one.  The switches may stand
  before the identifier (``SEQ \\# "00" Item`` is ``SEQ Item \\# "00"``), a switch that
  takes an argument taking the next token whatever it is; a ``\\#`` after a ``\\*`` is
  Word's error text (round 16's ``f4_probe.py``; :func:`sequence_switches`).
* ``REF name`` is the bookmark's text **in its own runs' formatting** -- direct and
  character-style formatting carried, the source paragraph's style not (a run in a
  Cambria heading is drawn in the destination's Calibri) -- a ``SEQ`` in it as counted;
  with ``\\* CHARFORMAT`` all of it in the instruction's first run's; with ``\\*
  MERGEFORMAT`` the *k*-th word in the *k*-th word's of the cached result, words past
  them and the spaces in the source's.  ``Upper``, ``Lower``, ``FirstCap``, ``Caps``
  apply.  A bookmark over two paragraphs (Word draws both, paragraph break and all),
  a missing one, or other switches (``\\n``, ``\\w``, ``\\p``...) are not computed.

Standard library only.
"""

from __future__ import annotations

import dataclasses
import re

from .linebreak import _roman
from .model import Paragraph, Run, Table

#: Formats and switches whose words are the application's language (see the module
#: docstring): not computed, the cached result drawn, and a warning.
LANGUAGE = frozenset({"ordinal", "cardinaltext", "ordinaltext", "dollartext"})

#: ``chicago``'s symbols, repeated once more at every cycle (``*``, ``†``, ``‡``, ``§``,
#: ``**``...): the first two measured.
CHICAGO = ("*", "\u2020", "\u2021", "\u00a7")


class Uncomputable(Exception):
    """The field's text cannot be computed from the file (``args[0]`` says why)."""


def _letters(value: int) -> str:
    """``a`` .. ``z``, then ``aa`` .. ``zz``, ``aaa``: the letter repeated once more every
    26 (measured: 27 ``AA``, 52 ``zz``, 53 ``aaa``, 255 ``uuuuuuuuuu``); a space for 0."""
    if value <= 0:
        return " "
    return chr(ord("a") + (value - 1) % 26) * ((value - 1) // 26 + 1)


def format_value(value: int, fmt: str | None) -> str:
    """A number in a ``w:pgNumType/@w:fmt`` (ECMA-376 ``ST_NumberFormat``).  Raises
    :class:`Uncomputable` for a format in the application's language or not measured."""
    fmt = fmt or "decimal"
    if fmt == "decimal":
        return str(value)
    if fmt in ("upperRoman", "lowerRoman"):
        # 4000 is MMMM, 4001 MMMMI; 0 is drawn as a space (measured).
        roman = _roman(value) if value > 0 else " "
        return roman.upper() if fmt == "upperRoman" else roman
    if fmt in ("upperLetter", "lowerLetter"):
        letters = _letters(value)
        return letters.upper() if fmt == "upperLetter" else letters
    if fmt == "numberInDash":
        return f"- {value} -"
    if fmt == "decimalZero":
        return f"{value:02d}"
    if fmt == "hex":
        return f"{value:X}"
    if fmt == "chicago" and value > 0:
        return CHICAGO[(value - 1) % 4] * ((value - 1) // 4 + 1)
    if fmt.lower() in LANGUAGE:
        raise Uncomputable(f"{fmt} is in the application's language")
    raise Uncomputable(f"the page number format {fmt!r} is not measured")


_SWITCH = re.compile(r'\\([*#@!])\s*("[^"]*"|\S+)')


def switches(instruction: str) -> list[tuple[str, str]]:
    """``[(kind, argument)]`` of an instruction's switches (``\\* roman``, ``\\# "000"``)."""
    return [(kind, argument.strip('"')) for kind, argument in _SWITCH.findall(instruction)]


def _picture(value: int, picture: str) -> str:
    """A numeric picture made of ``0`` and ``#`` only: the digits right-aligned in it, a
    place with no digit drawn ``0`` for a ``0`` and a space for a ``#`` (measured,
    ``f4_probe.py``: ``"00"`` 01, ``"#0"`` and 2 a space then 2, ``"##0"`` and 1 two
    spaces then 1, ``"0#"`` 03, ``"#"`` 2; a number longer than the picture all its
    digits, ``"#0"`` 120)."""
    if not picture or any(char not in "0#" for char in picture):
        raise Uncomputable(f"the numeric picture {picture!r} is not measured")
    digits = str(value)
    places = picture[:max(0, len(picture) - len(digits))]
    return "".join("0" if char == "0" else " " for char in places) + digits


def field_text(instruction: str, keyword: str, *, page: int, pages: int | None, section_pages: int | None,
               section_format: str | None) -> str:
    """What Word draws for a computed field on a page numbered ``page`` (before
    formatting), in a document of ``pages`` pages whose section has content on
    ``section_pages`` of them (``None``: not known -- past where the layout stops).

    For ``PAGEREF``, ``page`` and ``section_format`` are the bookmark's page's: its number
    as ``PAGE`` shows it there (:attr:`docx2svg.PageInfo.number`) and its section's
    ``page_number_format``.  Raises :class:`Uncomputable`."""
    if keyword == "PAGEREF":
        if re.search(r"\\p\b", _unquoted(instruction)):
            raise Uncomputable("\\p says where the bookmark is in the application's language")
        keyword = "PAGE"
    if keyword == "PAGE":
        value, fmt = page, section_format
    elif keyword == "NUMPAGES":
        if pages is None:
            raise Uncomputable("the page count is not known past where the layout stops")
        value, fmt = pages, None
    elif keyword == "SECTIONPAGES":
        if section_pages is None:
            raise Uncomputable("the section's page count is not known past where the layout stops")
        value, fmt = section_pages, None
    else:
        raise Uncomputable(f"{keyword} is not computed")
    return number_text(instruction, value, fmt)


def number_text(instruction: str, value: int, fmt: str | None = None) -> str:
    """``value`` formatted by the instruction's switches, else in ``fmt`` (a
    ``w:pgNumType/@w:fmt``; decimal when ``None``), then its case switches applied."""
    text = None
    case = None
    for kind, argument in switches(instruction):
        if kind == "#":
            text = _picture(value, argument)
            continue
        if kind != "*":
            continue
        word = argument.lower()
        if word in ("mergeformat", "charformat"):
            continue
        if word in ("upper", "lower", "firstcap", "caps"):
            case = word
        elif word == "roman":
            text = format_value(value, "upperRoman" if argument[:1].isupper() else "lowerRoman")
        elif word == "alphabetic":
            text = format_value(value, "upperLetter" if argument[:1].isupper() else "lowerLetter")
        elif word == "arabic":
            text = str(value)
        elif word == "arabicdash":
            text = f"- {value} -"
        elif word == "hex":
            text = f"{value:X}"
        elif word in LANGUAGE or word in ("ordinal", "cardtext", "ordtext"):
            raise Uncomputable(f"\\* {argument} is in the application's language")
        else:
            raise Uncomputable(f"the switch \\* {argument} is not measured")
    if text is None:
        text = format_value(value, fmt)
    if case == "upper":
        text = text.upper()
    elif case == "lower":
        text = text.lower()
    elif case in ("firstcap", "caps"):
        text = text[:1].upper() + text[1:] if case == "firstcap" else " ".join(w[:1].upper() + w[1:]
                                                                              for w in text.split(" "))
    return text


def apply_case(text: str, instruction: str) -> str:
    """``text`` with the instruction's ``\\* Upper``, ``Lower``, ``FirstCap`` or ``Caps``."""
    for kind, argument in switches(instruction):
        word = argument.lower()
        if kind != "*":
            continue
        if word == "upper":
            text = text.upper()
        elif word == "lower":
            text = text.lower()
        elif word == "firstcap":
            text = text[:1].upper() + text[1:]
        elif word == "caps":
            text = " ".join(w[:1].upper() + w[1:] for w in text.split(" "))
    return text


def _unquoted(instruction: str) -> str:
    return re.sub(r'"[^"]*"', '""', instruction)


def _argument(instruction: str) -> str | None:
    """A field's first argument: ``REF Name \\h`` -> ``Name``; quotes removed."""
    words = re.findall(r'"[^"]*"|\S+', instruction)
    return words[1].strip('"') if len(words) > 1 and not words[1].startswith("\\") else None


def _flags(instruction: str) -> list[tuple[str, str | None]]:
    """The letter switches of an instruction (``\\h``, ``\\r 5``), with a number after one."""
    return [(letter.lower(), number or None)
            for letter, number in re.findall(r"\\([A-Za-z])(?:\s+(\d+))?", _unquoted(instruction))]


# -- cross-references --------------------------------------------------------------------


def body_paragraphs(blocks) -> list[Paragraph]:
    """Every paragraph of ``blocks`` in document order, table cells in their place -- the
    order the layout numbers them in (``layout._all_paragraphs``)."""
    out: list[Paragraph] = []

    def walk(item) -> None:
        if isinstance(item, Table):
            for row in item.rows:
                for cell in row:
                    for inner in cell:
                        walk(inner)
        else:
            out.append(item)

    for block in blocks:
        walk(block)
    return out


def bookmarks(paragraphs: list[Paragraph]) -> dict[str, tuple[tuple[int, int], tuple[int, int] | None]]:
    """Bookmark name -> ``((paragraph, run index) of its start, that of its end or
    None)``, the first of a name winning, paragraphs indexed as :func:`body_paragraphs`."""
    starts: dict[str, tuple[str | None, tuple[int, int]]] = {}
    ends: dict[str | None, tuple[int, int]] = {}
    for block, paragraph in enumerate(paragraphs):
        for kind, mark, name, run in paragraph.bookmarks:
            if kind == "start" and name and name not in starts:
                starts[name] = (mark, (block, run))
            elif kind == "end":
                ends.setdefault(mark, (block, run))
    return {name: (start, ends.get(mark)) for name, (mark, start) in starts.items()}


def sequence_values(paragraphs: list[Paragraph]) -> dict[tuple[int, int], str | Uncomputable]:
    """``(paragraph, run) -> text`` of every ``SEQ`` field, counted as Word counts them (see
    the module docstring); an :class:`Uncomputable` where it cannot be."""
    fields = [(block, k, run.field.instruction) for block, paragraph in enumerate(paragraphs)
              for k, run in enumerate(paragraph.runs) if run.field is not None and run.field.keyword == "SEQ"]
    parsed = [(block, k, instruction, sequence_switches(instruction)) for block, k, instruction in fields]
    restarted = {(name or "").upper() for _, _, _, (name, found) in parsed if any(key == "s" for key, _ in found)}
    counts: dict[str, int] = {}
    out: dict[tuple[int, int], str | Uncomputable] = {}
    for block, k, instruction, (name, found) in parsed:
        out[(block, k)] = sequence_step(instruction, (name or "").upper(), found, counts, restarted)
    return out


#: ``SEQ``'s switches that take the token after them as their argument, whatever it is
#: (measured, ``f4_probe.py``: ``SEQ \\r Mm`` is Word's "no sequence given" error, ``Mm``
#: read as ``\\r``'s argument); ``\\c``, ``\\h`` and ``\\n`` take none.
SEQ_ARGUMENTS = frozenset("#*@rs")


def sequence_switches(instruction: str) -> tuple[str | None, list[tuple[str, str | None]]]:
    """A ``SEQ`` instruction's identifier and its switches in order, ``[(key, argument)]``:
    the key the switch's character (a letter lower-cased), the argument the token after a
    switch that takes one (:data:`SEQ_ARGUMENTS`), quotes removed.  The identifier is the
    first other token, quotes removed, **wherever the switches stand**: Word reads
    ``SEQ \\# "00" Item``, ``SEQ \\#"00" Item``, ``SEQ \\r 7 Item``, ``SEQ \\* roman Item``,
    ``SEQ \\h Item`` and ``SEQ \\# "00" "Item"`` as ``SEQ Item ...`` (measured,
    ``f4_probe.py``).  ``None`` when there is none."""
    tokens = re.findall(r'"[^"]*"|\\\S|[^\s"\\]+', instruction)
    name: str | None = None
    found: list[tuple[str, str | None]] = []
    k = 1 if tokens and tokens[0].upper() == "SEQ" else 0
    while k < len(tokens):
        token = tokens[k]
        k += 1
        if token.startswith("\\"):
            key = token[1] if token[1] in "#*@!" else token[1].lower()
            argument = None
            if key in SEQ_ARGUMENTS and k < len(tokens):
                argument = tokens[k].strip('"')
                k += 1
            found.append((key, argument))
        elif name is None:
            name = token.strip('"')
    return name, found


def sequence_step(instruction: str, name: str, found: list[tuple[str, str | None]], counts: dict[str, int],
                  restarted: frozenset[str] | set[str] = frozenset()) -> str | Uncomputable:
    """One ``SEQ`` field counted: ``counts`` (identifier, upper case -> its value) moved on
    as Word moves it, and the field's text, or an :class:`Uncomputable` -- ``name`` and
    ``found`` are :func:`sequence_switches`' (the name upper-cased).  A ``\\#`` after a
    ``\\*`` is Word's error text in the application's language, the sequence counted all
    the same (measured: ``SEQ Item \\* MERGEFORMAT \\# "00"``; ``\\# "00" \\* ARABIC``
    draws 2, the later switch winning)."""
    if not name:
        return Uncomputable("a SEQ field names no sequence")
    if name in restarted:
        return Uncomputable(f"the sequence {name!r} restarts at heading levels (\\s)")
    flags = dict(found)
    if "r" in flags:
        restart = flags["r"]
        if restart is None or not restart.isdigit():
            return Uncomputable("\\r takes a number")
        counts[name] = int(restart)
    elif "c" not in flags:
        counts[name] = counts.get(name, 0) + 1
    value = counts.get(name, 0)
    if "h" in flags:
        return ""
    keys = [key for key, _ in found]
    if "*" in keys and "#" in keys[keys.index("*"):]:
        return Uncomputable("a \\# after a \\* is Word's error text, in the application's language")
    try:
        return number_text(instruction, value)
    except Uncomputable as why:
        return why


_REF_FLAGS = frozenset({"h"})
_LAYOUT_FIELDS = frozenset({"PAGE", "NUMPAGES", "SECTIONPAGES", "PAGEREF"})


def reference_runs(paragraphs: list[Paragraph], run: Run, marks: dict) -> tuple[Run, ...]:
    """What Word draws for a ``REF`` field's run: the bookmark's text as runs (see the
    module docstring), each with the field's path.  ``marks`` is :func:`bookmarks`.
    Raises :class:`Uncomputable`."""
    field = run.field
    if field is None:
        raise Uncomputable("the run is not a field's")
    instruction = field.instruction
    name = _argument(instruction)
    if not name or name not in marks:
        raise Uncomputable("its bookmark is not in the body (Word draws an error in the application's language)")
    for letter, _ in _flags(instruction):
        if letter not in _REF_FLAGS:
            raise Uncomputable(f"the switch \\{letter} is not computed")
    (start_block, start_run), end = marks[name]
    if end is None:
        raise Uncomputable("its bookmark has no end")
    end_block, end_run = end
    if end_block != start_block:
        raise Uncomputable("its bookmark spans paragraphs (Word draws them, paragraph marks and all)")
    source = paragraphs[start_block].runs[start_run:end_run]
    for piece in source:
        if piece.breaks or piece.drawings or piece.anchors:
            raise Uncomputable("its bookmark holds a tab, a break or an object")
        if piece.field is not None and piece.field.keyword in _LAYOUT_FIELDS:
            raise Uncomputable("its bookmark holds a page-number field")
    words = [argument.lower() for kind, argument in switches(instruction) if kind == "*"]
    for word in words:
        if word not in ("mergeformat", "charformat", "upper", "lower", "firstcap", "caps"):
            raise Uncomputable(f"the switch \\* {word} is not measured on REF")
    text = apply_case("".join(piece.text for piece in source), instruction)
    if "charformat" in words:
        chars = [run.properties] * len(text)
    else:
        chars = [piece.properties for piece in source for _ in piece.text]
        if "mergeformat" in words and field.cached:
            old = [properties for piece, properties in field.cached for _ in piece]
            old_text = "".join(piece for piece, _ in field.cached)
            old_words = [old[match.start()] for match in re.finditer(r"\S+", old_text)]
            for index, match in enumerate(re.finditer(r"\S+", text)):
                if index < len(old_words):
                    chars[match.start():match.end()] = [old_words[index]] * (match.end() - match.start())
    if not text:
        return (dataclasses.replace(run, text="", field=None),)
    out: list[Run] = []
    begin = 0
    for index in range(1, len(text) + 1):
        if index == len(text) or chars[index] is not chars[begin]:
            out.append(Run(text=text[begin:index], properties=chars[begin], path=run.path))
            begin = index
    return tuple(out)


def cross_references(blocks) -> tuple[list, list[tuple[str, str]]]:
    """``blocks`` (the body) with every ``SEQ`` field's text counted and every ``REF``
    field's run replaced by the runs Word draws for it; and ``(keyword, why)`` for each
    field drawn as cached because it cannot be computed.  ``PAGEREF`` needs the layout:
    left to it."""
    paragraphs = body_paragraphs(blocks)
    problems: list[tuple[str, str]] = []
    sequences = sequence_values(paragraphs)
    replaced: dict[int, list] = {}
    for (block, k), value in sequences.items():
        if isinstance(value, Uncomputable):
            problems.append(("SEQ", value.args[0]))
            continue
        edited = replaced.setdefault(block, list(paragraphs[block].runs))
        edited[k] = dataclasses.replace(edited[k], text=value)
    counted = [dataclasses.replace(p, runs=tuple(replaced[b])) if b in replaced else p
               for b, p in enumerate(paragraphs)]
    marks = bookmarks(counted)
    expanded: dict[int, Paragraph] = {}
    for block, paragraph in enumerate(counted):
        if not any(run.field is not None and run.field.keyword == "REF" for run in paragraph.runs):
            continue
        runs: list = []
        index_of: dict[int, int] = {}
        for k, run in enumerate(paragraph.runs):
            index_of[k] = len(runs)
            if run.field is None or run.field.keyword != "REF":
                runs.append(run)
                continue
            try:
                runs.extend(reference_runs(counted, run, marks))
            except Uncomputable as why:
                problems.append(("REF", why.args[0]))
                runs.append(run)
        index_of[len(paragraph.runs)] = len(runs)
        # A REF replaced by several runs shifts the bookmarks after it.
        moved = tuple((kind, mark, name, index_of.get(k, k)) for kind, mark, name, k in paragraph.bookmarks)
        expanded[block] = dataclasses.replace(paragraph, runs=tuple(runs), bookmarks=moved)
    final = [expanded.get(b, p) for b, p in enumerate(counted)]
    return _rebuilt(blocks, iter(final)), problems


def _rebuilt(blocks, paragraphs):
    """``blocks`` with its paragraphs, in :func:`body_paragraphs` order, taken from
    ``paragraphs``."""
    def walk(item):
        if isinstance(item, Table):
            rows = tuple(tuple(tuple(walk(inner) for inner in cell) for cell in row) for row in item.rows)
            return dataclasses.replace(item, rows=rows)
        return next(paragraphs)

    return [walk(block) for block in blocks]


# -- STYLEREF ----------------------------------------------------------------------------

#: Word's texts for a ``STYLEREF`` it cannot resolve, in the application's language
#: (measured with Word for Mac 16 in English and in Dutch, round 21's ``styleref``
#: probes; the document's own language changes nothing): ``none`` -- the style is in the
#: document's styles but no paragraph of the body has it; ``undefined`` -- no style of
#: that name (``{name}`` is the name as the instruction writes it); ``unnamed`` -- the
#: instruction names none.  A text not measured in a language is not computed.
STYLEREF_ERRORS = {
    "en-US": {
        "none": "Error! No text of specified style in document.",
        "undefined": "Error! Use the Home tab to apply {name} to the text that you want to appear here.",
        "unnamed": "Error! No style name given.",
    },
    "nl-NL": {
        "none": "Fout! Geen tekst met de opgegeven stijl in het document.",
    },
}


@dataclasses.dataclass(frozen=True)
class StyleReference:
    """A ``STYLEREF`` instruction read: the style ``name`` as written (quotes removed, not
    trimmed -- Word does not trim it), whether ``\\l`` searches a page from its foot up, and
    whether ``\\p`` asks where the text is (``above`` / ``below`` in the body)."""

    name: str
    last: bool = False
    position: bool = False


#: The paragraph number switches -- ``\\n``, ``\\r``, ``\\w`` (and ``\\s``, which Word
#: reads as one of them) -- draw a list label's text in its own ways, not computed.
_NUMBER_SWITCHES = frozenset("nrws")


def style_reference(instruction: str) -> StyleReference:
    """A ``STYLEREF`` instruction (:class:`StyleReference`).  Raises :class:`Uncomputable`
    for a switch that is not computed: the paragraph number ones (``\\n`` ``\\r`` ``\\w``
    ``\\s``), a ``\\*`` format other than ``Upper``, ``Lower``, ``FirstCap``, ``Caps``,
    ``MERGEFORMAT`` and ``CHARFORMAT``, an unknown one, or a name not written first."""
    tokens = re.findall(r'"[^"]*"|\\\S|[^\s"\\]+', instruction)
    k = 1 if tokens and tokens[0].upper() == "STYLEREF" else 0
    name: str | None = None
    last = position = False
    first = True
    while k < len(tokens):
        token = tokens[k]
        k += 1
        if not token.startswith("\\"):
            if name is not None or not first:
                raise Uncomputable("its style name is not the first thing after STYLEREF, or is not one word")
            name = token[1:-1] if token.startswith('"') and token.endswith('"') and len(token) > 1 else token
            first = False
            continue
        first = False
        key = token[1]
        if key == "*":
            argument = tokens[k].strip('"').lower() if k < len(tokens) else ""
            k += 1
            if argument not in ("mergeformat", "charformat", "upper", "lower", "firstcap", "caps"):
                raise Uncomputable(f"the switch \\* {argument} is not measured on STYLEREF")
        elif key.lower() == "l":
            last = True
        elif key.lower() == "p":
            position = True
        elif key.lower() == "t":
            pass  # alone it changes nothing (measured); with a number switch, not computed
        elif key.lower() in _NUMBER_SWITCHES:
            raise Uncomputable(f"the paragraph number switch \\{key} is not computed")
        else:
            raise Uncomputable(f"the switch \\{key} is not measured on STYLEREF")
    return StyleReference(name or "", last, position)


def find_style(name: str, styles) -> tuple[str, frozenset[str]] | None:
    """The style a ``STYLEREF`` names: ``(kind, style ids)`` -- ``paragraph`` or
    ``character`` -- or ``None`` where no style has that name.  ``styles`` is the document's
    :class:`~docx2svg.model.StyleSheet` (or ``None``).

    The one place a field's style name is looked up.  Measured (English Word): the stored
    name (``w:name``), case ignored -- ``heading 1``, ``Heading 1`` and ``HEADING 1`` alike
    -- not the style id, not with spaces around it; a number ``1`` .. ``9`` is the built-in
    ``heading N``.  Dutch Word matches a built-in style by its Dutch name instead (``Kop 1``):
    the conversion's locale decides whether a lookup that finds nothing is Word's error
    (:func:`style_reference_target`)."""
    wanted = name.lower()
    if name.isdigit() and 1 <= int(name) <= 9:
        wanted = f"heading {int(name)}"
    kinds: dict[str, set] = {}
    for style_id, style in (styles.styles.items() if styles is not None else ()):
        if style.name is not None and style.name.lower() == wanted and style.kind in ("paragraph", "character"):
            kinds.setdefault(style.kind, set()).add(style_id)
    if not kinds:
        if name.isdigit() and 1 <= int(name) <= 9:
            return "paragraph", frozenset()
        return None
    kind = "paragraph" if "paragraph" in kinds else "character"
    return kind, frozenset(kinds[kind])


def style_reference_target(reference: StyleReference, styles, language: str) -> tuple[str, frozenset[str]] | str:
    """What a ``STYLEREF`` searches for (:func:`find_style`), or Word's error text where it
    finds no style (:data:`STYLEREF_ERRORS`).  Raises :class:`Uncomputable` for the error
    text of a language not measured, a style found by a name Word may not match in the
    conversion's language (built-in names are its interface language's), and the default
    character style (Word draws text of the page's other runs: not computed)."""
    errors = STYLEREF_ERRORS.get(language, {})
    if not reference.name:
        if "unnamed" not in errors:
            raise Uncomputable("the instruction names no style: Word's error text, in the application's language")
        return errors["unnamed"]
    found = find_style(reference.name, styles)
    if found is None:
        if "undefined" not in errors:
            raise Uncomputable("no style has that name in English: Word's interface language decides")
        return errors["undefined"].format(name=reference.name)
    kind, ids = found
    if kind == "character" and styles is not None and any(styles.styles[i].default for i in ids):
        raise Uncomputable("STYLEREF to the default character style is not measured")
    return kind, ids


@dataclasses.dataclass(frozen=True)
class Occurrence:
    """Text in the style a ``STYLEREF`` searches for: the paragraph (:func:`body_paragraphs`
    index) it is in, the character it starts at, and its text -- an :class:`Uncomputable`
    where it holds what is not text."""

    block: int
    offset: int
    text: str | Uncomputable


#: Fields whose text in a paragraph depends on the page it is drawn on.
_PAGED = frozenset({"PAGE", "NUMPAGES", "SECTIONPAGES", "PAGEREF", "STYLEREF"})


def _occurrence_text(runs) -> str | Uncomputable:
    """The text of ``runs`` as ``STYLEREF`` shows it: a page break is a space (measured: a
    heading split by one), list labels are not part of it, a floating drawing's anchor is
    nothing; a tab, another break or an inline object is not measured."""
    out = []
    for run in runs:
        if run.field is not None and run.field.keyword in _PAGED:
            return Uncomputable("its text holds a page-number field")
        if run.properties is not None and run.properties.declared.get("vanish"):
            return Uncomputable("its text holds hidden text")
        text = run.text
        for position, kind in sorted(run.breaks, reverse=True):
            if kind.startswith("anchor:"):
                continue  # a floating drawing: not text (measured: a text box's anchor)
            if kind != "page":
                return Uncomputable("its text holds a tab, a break or an object")
            text = text[:position] + " " + text[position:]
        out.append(text)
    return "".join(out)


def style_occurrences(paragraphs: list[Paragraph], kind: str, ids: frozenset[str], default_paragraph: str | None,
                      skip: frozenset[int] = frozenset()) -> list[Occurrence]:
    """Every stretch of ``paragraphs`` (:func:`body_paragraphs`) in the style ``ids`` in
    document order: a paragraph of a paragraph style, whole (an empty one too: measured, its
    text is empty), or each run of runs of a character style (measured: two words of it
    with others between are two).  ``default_paragraph``: the id of the paragraph style a
    paragraph naming none has; ``skip``: paragraphs not searched (notes placed in the flow)."""
    out: list[Occurrence] = []
    for block, paragraph in enumerate(paragraphs):
        if block in skip:
            continue
        if kind == "paragraph":
            style = paragraph.properties.style_id if paragraph.properties is not None else None
            if (style or default_paragraph) in ids:
                out.append(Occurrence(block, 0, _occurrence_text(paragraph.runs)))
            continue
        offset = 0
        stretch: list = []
        start = 0
        for run in [*paragraph.runs, None]:
            inside = run is not None and run.properties is not None and run.properties.style_id in ids
            if inside:
                if not stretch:
                    start = offset
                stretch.append(run)
            elif stretch:
                out.append(Occurrence(block, start, _occurrence_text(stretch)))
                stretch = []
            if run is not None:
                offset += len(run.text)
    return out


def style_reference_text(instruction: str, found: Occurrence | None, target: tuple | str,
                         language: str) -> str | Uncomputable:
    """A ``STYLEREF``'s text: the occurrence found, its case switches applied; Word's error
    text where the target is one, or where nothing in the body has the style
    (:data:`STYLEREF_ERRORS`)."""
    if isinstance(target, str):
        return target
    if found is None:
        errors = STYLEREF_ERRORS.get(language, {})
        if "none" not in errors:
            return Uncomputable("no text of the style in the document: Word's error text, in the application's "
                                "language")
        return errors["none"]
    if isinstance(found.text, Uncomputable):
        return found.text
    return apply_case(found.text, instruction)
