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
  (upper case) and ``\\# "000"`` override the section's format; ``MERGEFORMAT`` and
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
  not computed, nor is any field of an identifier that has one.
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
    """A numeric picture made of ``0`` and ``#`` only: zero-padded to its ``0``s."""
    if not picture or any(char not in "0#" for char in picture):
        raise Uncomputable(f"the numeric picture {picture!r} is not measured")
    return f"{value:0{picture.count('0')}d}"


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
    restarted = {(_argument(instruction) or "").upper() for _, _, instruction in fields
                 if any(letter == "s" for letter, _ in _flags(instruction))}
    counts: dict[str, int] = {}
    out: dict[tuple[int, int], str | Uncomputable] = {}
    for block, k, instruction in fields:
        name = (_argument(instruction) or "").upper()
        if not name:
            out[(block, k)] = Uncomputable("a SEQ field names no sequence")
            continue
        if name in restarted:
            out[(block, k)] = Uncomputable(f"the sequence {name!r} restarts at heading levels (\\s)")
            continue
        flags = dict(_flags(instruction))
        if "r" in flags and flags["r"] is not None:
            counts[name] = int(flags["r"])
        elif "c" not in flags:
            counts[name] = counts.get(name, 0) + 1
        value = counts.get(name, 0)
        if "h" in flags:
            out[(block, k)] = ""
            continue
        try:
            out[(block, k)] = number_text(instruction, value)
        except Uncomputable as why:
            out[(block, k)] = why
    return out


_REF_FLAGS = frozenset({"h"})
_LAYOUT_FIELDS = frozenset({"PAGE", "NUMPAGES", "SECTIONPAGES", "PAGEREF"})


def reference_runs(paragraphs: list[Paragraph], run: Run, marks: dict) -> tuple[Run, ...]:
    """What Word draws for a ``REF`` field's run: the bookmark's text as runs (see the
    module docstring), each with the field's path.  ``marks`` is :func:`bookmarks`.
    Raises :class:`Uncomputable`."""
    instruction = run.field.instruction
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
        if "mergeformat" in words and run.field.cached:
            old = [properties for piece, properties in run.field.cached for _ in piece]
            old_text = "".join(piece for piece, _ in run.field.cached)
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
        runs = replaced.setdefault(block, list(paragraphs[block].runs))
        runs[k] = dataclasses.replace(runs[k], text=value)
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
