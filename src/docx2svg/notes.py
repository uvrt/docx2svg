"""Endnotes, placed in the flow: numbered, and put where Word puts them.

An endnote reference draws its note's number where it stands, in the reference run's
properties, and the note goes after the text -- after the document's last paragraph, or,
under ``w:endnotePr/w:pos`` ``sectEnd``, after each section's -- under a separator.  So
the notes are laid out as paragraphs of the flow: :func:`with_endnotes` returns the
document with every reference's number spliced into its run's text and the notes'
paragraphs (their own numbers spliced in too) placed after the text, each group opened
by the separator's paragraph.  Everything after that -- line breaking, the page stack,
pagination, drawing -- is the body's, and is measured as the body's is.  Measured by
``make_endnote_probe.py`` (ROADMAP.md, "Endnotes -- measured"):

* **The number** is the note's place among the document's references, from ``1``, in
  the section's format (``w:numFmt``; ``lowerRoman`` when unstated) -- the *section's*
  ``w:endnotePr``, not the document's: the settings' format and first number are not
  taken, in mode 14 or 15.  A section that restarts (``w:numRestart`` ``eachSect``), and
  the first, start at its ``w:numStart``; the others go on counting.
* **Where**: after the last paragraph of the document, on its page -- no space but the
  paragraphs' own -- or of each section with ``sectEnd``, before its break.
* **The separator** is the endnotes part's ``separator`` note, **when the settings'
  ``w:endnotePr`` names it** (``w:endnote/@w:id``); otherwise -- no settings part, or one
  that does not name it -- Word's own: a paragraph of the default style with no space
  after and single spacing, its formatting in the file ignored.  Its line is drawn as a
  strikethrough of its run's face and size is (``OS/2`` strikeout position and size, at
  the device size), 2,880 twips long from where the paragraph's text starts
  (:data:`SEPARATOR_TWIPS`); the continuation separator's, from there to the column's
  right edge, whatever the alignment.
* **On a page the notes go on to**, the continuation separator's paragraph first
  (:func:`docx2svg.paginate.flow`'s ``continuation``).

**Footnotes** are numbered here too (:func:`with_footnotes`), but not placed in the flow:
their references keep their marks for the paginator, which keeps each note's room at the
foot of its reference's page (ROADMAP.md 4.3), and the layout draws them there under the
separator (:meth:`docx2svg.layout._Placer._footnotes`).  Measured by
``make_footnote_draw_probe.py`` (ROADMAP.md, "Footnotes drawn -- measured"): numbered
by the section's ``w:footnotePr`` (``decimal`` unstated), a section that goes on counting
by the note's place among all the document's references (:func:`footnote_numbers`); the
number drawn as an endnote's is; the separators the part's when the settings name them,
Word's own otherwise, and in mode 15 a page a note goes on to opens with the separator
rather than the continuation separator (:func:`footnote_separator_kind`).
"""

from __future__ import annotations

import dataclasses

from .linebreak import format_number
from .model import Document, Paragraph, ParagraphProperties, Run, Spacing, Table

#: The separator's line, twips: two inches, whatever the paragraph's indent or face.
SEPARATOR_TWIPS = 2880
#: The number format when a section states none (Word's default for endnotes).
DEFAULT_FORMAT = "lowerRoman"


def number_text(value: int, fmt: str | None) -> str:
    """A note's number in a ``w:numFmt`` (``lowerRoman`` when ``None``)."""
    fmt = fmt or DEFAULT_FORMAT
    if fmt == "chicago":
        return "*†‡§"[(value - 1) % 4] * ((value - 1) // 4 + 1)
    return format_number(value, fmt)


def default_separator(kind: str, notes: str = "endnote") -> Paragraph:
    """Word's own separator paragraph: the default style, no space after, single."""
    properties = ParagraphProperties(spacing=Spacing(after=0, line=240, line_rule="auto"),
                                     declared={"spacing.after": 0, "spacing.line": 240, "spacing.lineRule": "auto"})
    return Paragraph(runs=(Run(text="", breaks=((0, kind),), path="w:r[1]"),), properties=properties,
                     path=f"w:{notes}s/{kind}", note=kind)


def separator(document: Document, kind: str, notes: str = "endnote") -> list:
    """The separator's paragraphs as Word draws them (:func:`default_separator` unless
    the settings name the part's own); ``notes``: ``endnote`` or ``footnote``."""
    note_id = getattr(document, f"{notes}_separator_ids").get(kind)
    blocks = getattr(document, f"{notes}_separators").get(kind)
    if blocks and note_id is not None and note_id in getattr(document, f"{notes}_named"):
        return [dataclasses.replace(block, note=kind) if isinstance(block, Paragraph) else block
                for block in blocks]
    return [default_separator(kind, notes)]


def _splice(run: Run, texts: dict[int, str], keep: bool = False) -> Run:
    """``run`` with ``texts`` (mark index -> text) put in its text where those marks are,
    and the marks removed (``keep``: left where their text starts); a mark after one keeps
    its place after its text."""
    inserted = [(k, position, texts[k]) for k, (position, _) in enumerate(run.breaks) if k in texts]
    out, last = "", 0
    for _, position, value in inserted:
        out += run.text[last:position] + value
        last = position
    out += run.text[last:]
    moved = tuple((position + sum(len(value) for k, at, value in inserted
                                  if at < position or (at == position and k < j)), kind)
                  for j, (position, kind) in enumerate(run.breaks) if keep or j not in texts)
    return dataclasses.replace(run, text=out, breaks=moved)


def _split(run: Run, texts: dict[int, str], keep: bool = False) -> list[Run]:
    """``run`` as runs, with ``texts`` (mark index -> text) put where those marks are, each
    a run of its own marked as a note's number (:attr:`Run.note_number`), and the marks
    removed (``keep``: at the start of their number's run); every other mark stays with
    the text around it."""
    out: list[Run] = []
    last = 0
    pending: list[tuple[int, str]] = []

    def piece(start: int, stop: int) -> None:
        marks = tuple((position - start, kind) for position, kind in pending)
        if run.text[start:stop] or marks:
            out.append(dataclasses.replace(run, text=run.text[start:stop], breaks=marks))

    for k, (position, kind) in enumerate(run.breaks):
        if k in texts:
            piece(last, position)
            pending = []
            out.append(dataclasses.replace(run, text=texts[k], breaks=((0, kind),) if keep else (), drawings=(),
                                           anchors=(), note_number=True))
            last = position
        else:
            pending.append((position, kind))
    piece(last, len(run.text))
    return out


def _number_runs(paragraph: Paragraph, kind: str, numbers, keep: bool = False) -> Paragraph:
    """``paragraph`` with every mark ``kind`` (``endnoteReference:<id>``, ``endnoteRef``)
    replaced by its number, ``numbers(mark) -> str`` (``keep``: the mark kept, at the
    number's start)."""
    runs = []
    changed = False
    for run in paragraph.runs:
        texts = {k: numbers(mark) for k, (_, mark) in enumerate(run.breaks) if mark.startswith(kind)}
        if texts and any(mark.startswith(("drawing", "anchor:")) for _, mark in run.breaks):
            # A run that also holds drawings keeps its marks' order with them: the
            # number goes in its text (laid out as text is).
            changed = True
            runs.append(_splice(run, texts, keep))
        elif texts:
            changed = True
            runs.extend(_split(run, texts, keep))
        else:
            runs.append(run)
    return dataclasses.replace(paragraph, runs=tuple(runs)) if changed else paragraph


def _section_numbers(document: Document) -> list[int]:
    """Each top-level block's section."""
    out = []
    top_level = 0
    for item in document.body:
        section = len(document.sections) - 1
        for number, s in enumerate(document.sections):
            if s.first_paragraph <= top_level < s.last_paragraph:
                section = number
                break
        out.append(section)
        if not isinstance(item, Table):
            top_level += 1
    return out


def with_endnotes(document: Document) -> tuple[Document, set[int]]:
    """The document with its endnotes numbered and placed (see the module), and the ids
    of the blocks placed, which the file does not hold where they now stand."""
    if not any(kind.startswith("endnoteReference:") for paragraph in _paragraphs(document.body)
               for run in paragraph.runs for _, kind in run.breaks):
        return document, set()
    sections = _section_numbers(document)
    counter = 0
    current = None
    #: Per section, ``[(note id, its number text)]`` in the order referenced.
    referenced: dict[int, list[tuple[int, str]]] = {}
    body = []

    def numbered(item, section: int):
        nonlocal counter, current
        if isinstance(item, Table):
            rows = tuple(tuple(tuple(numbered(block, section) for block in cell) for cell in row) for row in item.rows)
            return dataclasses.replace(item, rows=rows)
        if section != current:
            s = document.sections[section]
            if current is None or s.endnote_restart == "eachSect":
                counter = s.endnote_start if s.endnote_start is not None else 1
            current = section
        fmt = document.sections[section].endnote_format

        def number(mark: str) -> str:
            nonlocal counter
            parts = mark.split(":")
            if len(parts) > 2 and parts[2] == "custom":
                referenced.setdefault(section, []).append((int(parts[1]), ""))
                return ""
            text = number_text(counter, fmt)
            counter += 1
            try:
                referenced.setdefault(section, []).append((int(parts[1]), text))
            except ValueError:
                pass
            return text

        return _number_runs(item, "endnoteReference:", number)

    for item, section in zip(document.body, sections):
        body.append(numbered(item, section))
    placed: set[int] = set()
    by_end: dict[int, list] = {}
    position = document.endnote_position or "docEnd"
    for section, notes in referenced.items():
        target = section if position == "sectEnd" else len(document.sections) - 1
        by_end.setdefault(target, []).extend(notes)
    groups: dict[int, list] = {}
    for target, notes in by_end.items():
        blocks = list(separator(document, "separator"))
        for note_id, text in notes:
            for block in document.endnotes.get(note_id, []):
                if isinstance(block, Paragraph):
                    block = _number_runs(block, "endnoteRef", lambda _mark, text=text: text)
                    block = dataclasses.replace(block, note="endnote")
                blocks.append(block)
        groups[target] = blocks
        placed.update(id(block) for block in blocks)
    # Each group after the last block of its section (the document's last, docEnd).
    out = []
    for k, (item, section) in enumerate(zip(body, sections)):
        out.append(item)
        last_of_section = k == len(body) - 1 or sections[k + 1] != section
        if last_of_section and section in groups:
            out.extend(groups.pop(section))
    for blocks in groups.values():
        out.extend(blocks)
    return _renumbered(document, out, placed), placed


def _paragraphs(blocks) -> list[Paragraph]:
    out = []
    for item in blocks:
        if isinstance(item, Table):
            for row in item.rows:
                for cell in row:
                    out.extend(_paragraphs(cell))
        else:
            out.append(item)
    return out


def _renumbered(document: Document, body: list, placed: set[int]) -> Document:
    """``document`` with ``body``: its top-level paragraphs, and each section's range of
    them grown by the notes placed in it."""
    paragraphs = [item for item in body if not isinstance(item, Table)]
    sections = []
    top_level = 0
    original = 0
    bounds = []
    # Map each original top-level index to its new one.
    new_index = {}
    for item in body:
        if isinstance(item, Table):
            continue
        if id(item) not in placed:
            new_index[original] = top_level
            original += 1
        top_level += 1
    for number, s in enumerate(document.sections):
        first = new_index.get(s.first_paragraph, len(paragraphs)) if s.first_paragraph < original else len(paragraphs)
        bounds.append(first)
    for number, s in enumerate(document.sections):
        last = bounds[number + 1] if number + 1 < len(bounds) else len(paragraphs)
        sections.append(dataclasses.replace(s, first_paragraph=bounds[number], last_paragraph=last))
    return dataclasses.replace(document, body=body, paragraphs=paragraphs, sections=sections)


# -- footnotes ---------------------------------------------------------------------------

#: The number format when a section states none (Word's default for footnotes).
DEFAULT_FOOTNOTE_FORMAT = "decimal"


def footnote_separator_kind(document: Document, kind: str) -> str:
    """The separator a page's footnotes open with: on a page a note goes on to, the
    continuation separator -- except in mode 15, where Word opens it with the separator
    itself (``make_footnote_draw_probe.py``, ``overflow`` and ``separators``: the 2,880-twip
    line, the part's 20 pt one where it is named, not the column-wide one)."""
    if kind == "continuationSeparator" and (document.compatibility_mode or 0) >= 15:
        return "separator"
    return kind


def footnote_references(document: Document) -> list[tuple[int, int, bool, bool]]:
    """Every footnote reference of the body in order, table cells included: ``(note id,
    section, whether it is a custom mark, whether it is in a table)``."""
    out = []
    sections = _section_numbers(document)

    def walk(item, section, in_table=False):
        if isinstance(item, Table):
            for row in item.rows:
                for cell in row:
                    for block in cell:
                        walk(block, section, True)
            return
        for run in item.runs:
            for _, mark in run.breaks:
                if mark.startswith("footnoteReference:"):
                    parts = mark.split(":")
                    try:
                        out.append((int(parts[1]), section, len(parts) > 2 and parts[2] == "custom", in_table))
                    except ValueError:
                        pass

    for item, section in zip(document.body, sections):
        walk(item, section)
    return out


def footnote_numbers(document: Document, pages: dict | None = None) -> list[str]:
    """Each footnote reference's number (:func:`footnote_references` order; ``""`` for a
    custom mark).  Measured by ``make_footnote_draw_probe.py`` (``format``, ``chain``,
    ``eachpage``, ``docpr``, ``docpage``), every setting alike:

    * the **section's** ``w:footnotePr`` is taken, the document's (in the settings) is
      not -- neither its format, its first number nor its restart;
    * the format is ``w:numFmt`` (``decimal`` when unstated);
    * a section that does not restart numbers its notes by their place among **all the
      document's** references, from its own ``w:numStart`` (1 when unstated): after a
      section restarting, the next goes on from the whole document's count (``chain``:
      the fourth section's notes are 7 and 8, after 1, 2, 9, 10 and 1, 2);
    * ``eachSect`` numbers them by their place in the section, ``eachPage`` by their place
      on their page (``pages``: note id -> page; the section's count where unknown), each
      from ``w:numStart``.
    """
    out = []
    index = 0
    by_section: dict[int, int] = {}
    by_page: dict = {}
    for note_id, section_number, custom, _ in footnote_references(document):
        if custom:
            out.append("")
            continue
        index += 1
        by_section[section_number] = by_section.get(section_number, 0) + 1
        section = document.sections[section_number]
        restart = section.footnote_restart or "continuous"
        start = section.footnote_start if section.footnote_start is not None else 1
        if restart == "eachSect":
            place = by_section[section_number]
        elif restart == "eachPage":
            page = pages.get(note_id) if pages else None
            key = (section_number, "page", page) if page is not None else (section_number, "section")
            by_page[key] = by_page.get(key, 0) + 1
            place = by_page[key]
        else:
            place = index
        out.append(number_text(start - 1 + place, section.footnote_format or DEFAULT_FOOTNOTE_FORMAT))
    return out


def with_footnotes(document: Document, pages: dict | None = None) -> Document:
    """The document with every footnote reference's number spliced into its run's text
    (a run of its own, :attr:`Run.note_number`, still holding the reference's mark for the
    paginator) and every note's own number (``w:footnoteRef``) into its paragraph
    (:func:`footnote_numbers`)."""
    references = footnote_references(document)
    if not references:
        return document
    numbers = iter(footnote_numbers(document, pages))
    by_note: dict[int, str] = {}

    def numbered(item):
        if isinstance(item, Table):
            rows = tuple(tuple(tuple(numbered(block) for block in cell) for cell in row) for row in item.rows)
            return dataclasses.replace(item, rows=rows)

        def number(mark: str) -> str:
            text = next(numbers, "")
            try:
                by_note.setdefault(int(mark.split(":")[1]), text)
            except ValueError:
                pass
            return text

        return _number_runs(item, "footnoteReference:", number, keep=True)

    body = [numbered(item) for item in document.body]
    paragraphs = [item for item in body if not isinstance(item, Table)]
    footnotes = {note_id: [_number_runs(block, "footnoteRef", lambda _mark, text=by_note.get(note_id, ""): text)
                           if isinstance(block, Paragraph) else block for block in blocks]
                 for note_id, blocks in document.footnotes.items()}
    return dataclasses.replace(document, body=body, paragraphs=paragraphs, footnotes=footnotes)
