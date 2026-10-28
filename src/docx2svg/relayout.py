"""Laying a document out again from a previous layout, from where an edit reaches (round 19, R1).

A live view of a report shows its pages while an agent writes it, one paragraph or one
cell at a time; laying the whole document out after each edit costs a layout of every
page.  :func:`docx2svg.lay_out_docx` with ``previous`` (a layout of an earlier version of
the same document) instead:

1. **finds what changed** -- every part but the main one compared by digest, and the
   body's top-level elements compared by a digest of their bytes, together with what each
   block is measured from beyond its own XML (its list labels, what the flow cannot
   model in it): the blocks before the first that differs and after the last are the
   previous layout's;
2. **measures only those** -- the flow (:func:`docx2svg.paginate.flow`) takes a
   paragraph's lines and line heights and a table's flow from the previous layout's
   :class:`Measures` wherever what they are measured from is the same;
3. **lays the pages out again from the page holding each changed block** -- or the
   page before, where that page's end may depend on it (the block starts the page, or a
   chain of keep-with-next paragraphs from the page's start reaches it) or the page
   starts inside a paragraph or table -- with the pagination's state as it stood there
   (:attr:`docx2svg.paginate.Pagination.states`), **until a page starts at the start of an
   unchanged block in the same place, on the same page, with the same page number and the
   same footnote carried over** as a page of the previous layout did; from there on the
   previous layout's pages are taken as they are, their elements' paths renumbered where
   blocks were inserted or deleted before them;
4. **says which pages to draw again** (:attr:`Layout.changed_pages`): the pages whose
   drawing differs from the previous layout's page of the same number, and the pages it
   did not have -- a page whose only change is its elements' ``data-docx-path`` is not one.

**Fields** (round 20).  A body with page-number fields -- a table of contents' lines, each
a ``PAGEREF`` to its heading's bookmark, a ``PAGE``, a ``NUMPAGES`` -- is laid out in
passes, the fields' values found by one laid out by the next until they are the ones it
was laid out with (:meth:`docx2svg.layout._Placer.run_full`, four passes at most).  A
layout from a previous one makes the same passes (:class:`Again`), each from the same
pass of the previous layout (:attr:`State.passes`: the first from its first, every later
one from its last, the one drawn) and in **stretches** (:func:`windows`): the blocks the
edit changed, and each block elsewhere whose text the pass sets differs from the
previous layout's -- a ``TOC`` line whose page number changed, the ``SEQ`` numbers after an
inserted caption and the ``REF`` to them, a ``PAGE`` on a page that moved.  Each stretch is
laid out from the page holding it until a page starts in step with the previous pass
again (a hard page break after it, typically: after a table of contents, the pages after
the break are the previous pass's, moved by the pages added or removed before it), so a
pass's pagination, and the values its fields find, are those of a full layout's same
pass; and as many passes are made.  ``SEQ`` and ``REF``, which depend on the document's
order and not on the pages, are found once from the new document: a block whose value
changed is a stretch like an edited one; so is a ``STYLEREF`` in the body (round 21).  A
``STYLEREF`` in a header or footer shows text of the page's own (:class:`State`'s
``story_refs``): a page kept from the previous layout whose header or footer text changed --
a heading edited, restyled or moved -- is drawn again.  Where such a header's height on some
page differs from its cached one, the full layout's passes lay that page out under it, and
a layout from a previous one is a whole one (``styleref-height``).

The result is the layout a full layout of the new bytes makes, page for page, byte for
byte in every page's SVG (``tests/test_relayout.py`` holds it to that over random edits).
Where an edit cannot be scoped, the whole document is laid out again -- with the
measurements of unchanged blocks still reused where the document's styles and settings
are the same -- and :attr:`Layout.relaid` says why (:data:`FALLBACKS`) and
:attr:`Layout.relaid_cause` what: the element, the field or the part.
"""

from __future__ import annotations

import dataclasses
import hashlib
import io
import re
import zipfile
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Callable

from . import paginate
from .paginate import Position

if TYPE_CHECKING:
    from .fonts import InstalledFonts
    from .model import Paragraph, Table

#: Why a layout from a previous one lays the whole document out again, by
#: :attr:`Layout.relaid` code.  With ``measured``, the measurements of the blocks that did
#: not change are still taken from the previous layout.
FALLBACKS = {
    "no-state": "the previous layout keeps nothing to lay out again from (made by another docx2svg, or "
                "without keeping it)",
    "version": "the previous layout was made by another docx2svg",
    "options": "the font folders, substitutes, measurement or locale differ from the previous layout's",
    "styles": "the styles part changed",
    "numbering": "the numbering definitions changed",
    "settings": "the settings part changed",
    "theme": "the theme changed",
    "fonts": "the font table or an embedded font changed",
    "headers-footers": "a header or footer changed",
    "notes": "the footnotes or endnotes part changed",
    "parts": "another part the drawing reads changed (a picture, a chart), or a relationship was removed",
    "sections": "a section's properties changed (a section break added, removed or edited)",
    "note-references": "footnote or endnote references were added, removed or moved: the notes are numbered again",
    "footnote-restart": "footnote numbers restart on each page",
    "line-numbers": "a section numbers its lines",
    "columns": "a section has several text columns",
    "stopped": "the layout stops (CARRY_ON is off)",
    "fields-open": "an edited block opens or closes a field another block holds",
    "joined": "an edited block's paragraph mark is deleted as a revision (it joins the paragraph after it)",
    "body": "the body's elements could not be read for comparing",
    "page-count": "the page count changed and a header or footer shows NUMPAGES or SECTIONPAGES",
    "styleref-height": "a header's or footer's STYLEREF makes it taller or shorter on a page than with its cached "
                       "result, which moves that page's body",
}

#: Parts that do not change what is drawn: properties, custom XML, comments.
_UNDRAWN = re.compile(r"^(docProps/|customXml/|word/comments[A-Za-z]*\.xml$|word/people\.xml$|"
                      r"word/_rels/comments[A-Za-z]*\.xml\.rels$|word/webSettings\.xml$)")
MAIN = "word/document.xml"
MAIN_RELS = "word/_rels/document.xml.rels"
_PART_REASONS = (
    (re.compile(r"^word/styles(WithEffects)?\.xml$"), "styles"),
    (re.compile(r"^word/numbering\.xml$"), "numbering"),
    (re.compile(r"^word/settings\.xml$"), "settings"),
    (re.compile(r"^word/theme/"), "theme"),
    (re.compile(r"^word/(fontTable\.xml|fonts/)"), "fonts"),
    (re.compile(r"^word/(header|footer)\d*\.xml$|^word/_rels/(header|footer)\d*\.xml\.rels$"), "headers-footers"),
    (re.compile(r"^word/(footnotes|endnotes)\.xml$"), "notes"),
)

#: The version of what :class:`State` keeps; a state of another is not read.
STATE_FORMAT = 3


def _digest(data: bytes) -> bytes:
    return hashlib.blake2b(data, digest_size=16).digest()


# -- the body's elements ---------------------------------------------------------------------------

_TAG = re.compile(rb'<(/?)([A-Za-z_][\w.:-]*)((?:\s+[^\s=/>]+\s*=\s*(?:"[^"]*"|\'[^\']*\'))*)\s*(/?)>')
_BODY = re.compile(rb"<(?:[A-Za-z_][\w.-]*:)?body(?:\s[^>]*)?>")
_BEGIN = re.compile(rb'fldCharType\s*=\s*["\']begin["\']')
_END = re.compile(rb'fldCharType\s*=\s*["\']end["\']')
_DELETED_MARK = re.compile(rb"<(?:\w+:)?pPr\b(?:(?!</(?:\w+:)?pPr>).)*<(?:\w+:)?del\b", re.S)
_IDS = re.compile(rb'\s(?:\w+:)?(?:paraId|textId)\s*=\s*"[^"]*"')
_NOTE_REFERENCE = re.compile(rb'<(?:\w+:)?(footnote|endnote)Reference\b[^>]*?\b(?:\w+:)?id\s*=\s*["\'](-?\d+)["\']')


@dataclass(frozen=True)
class Segment:
    """One element of the body (``w:p``, ``w:tbl``, ``w:sdt``, ``w:sectPr``, a bookmark...),
    as its bytes say: its local name, a digest of its bytes, whether it is ``plain`` -- its
    fields open and close inside it -- the
    footnote and endnote references it holds, in order, and whether it holds a section's
    properties."""

    name: str
    digest: bytes
    plain: bool
    notes: tuple
    section: bool = False
    #: A paragraph mark deleted as a revision: the paragraph joins the one after it.
    joined: bool = False
    #: A digest of its bytes less its paragraphs' and rows' ``w14:paraId`` and
    #: ``w14:textId``, which a layout draws (``data-docx-id``) but measures nothing by:
    #: what the flow's measurements are kept by (:class:`Measures`), so that stamping ids
    #: on every paragraph (docx-agent's first edit) measures nothing again.
    measure: bytes = b""


def segments(xml: bytes) -> list[Segment] | None:
    """The body's elements of ``word/document.xml``, in order, read from its bytes; ``None``
    where they cannot be read that way (a comment, a processing instruction or CDATA in
    the body, or a tag the scan does not parse)."""
    found = elements(xml)
    return None if found is None else [_segment(name, data) for name, data in found]


def elements(xml: bytes) -> list[tuple[str, bytes]] | None:
    """The body's elements of ``word/document.xml`` as ``(local name, bytes)``, in order
    (:func:`segments`)."""
    found = _BODY.search(xml)
    if found is None:
        return None
    start = found.end()
    end = xml.rfind(b"</", start)
    if end < 0:
        return None
    body = xml[start:end]
    if b"<!--" in body or b"<?" in body or b"<![CDATA[" in body:
        return None
    out: list[tuple[str, bytes]] = []
    depth = 0
    begin = 0
    position = 0
    name = ""
    for match in _TAG.finditer(body):
        if body[position:match.start()].count(b"<"):
            return None  # a "<" the scan did not read as a tag
        position = match.end()
        closing, tag, empty = match.group(1), match.group(2), match.group(4)
        if closing:
            depth -= 1
            if depth < 0:
                break
            if depth == 0:
                out.append((name, body[begin:match.end()]))
        elif empty:
            if depth == 0:
                out.append((tag.split(b":")[-1].decode(), body[match.start():match.end()]))
        else:
            if depth == 0:
                begin, name = match.start(), tag.split(b":")[-1].decode()
            depth += 1
    if depth != 0 and depth != -1:
        return None
    return out


def _segment(name: str, data: bytes) -> Segment:
    joined = _DELETED_MARK.search(data) is not None
    plain = len(_BEGIN.findall(data)) == len(_END.findall(data))
    notes = tuple((kind.decode(), int(number)) for kind, number in _NOTE_REFERENCE.findall(data))
    return Segment(name, _digest(data), plain, notes, name == "sectPr" or b"sectPr" in data, joined,
                   _digest(_IDS.sub(b"", data)))


# -- the parts -------------------------------------------------------------------------------------


def part_digests(data: bytes) -> tuple[dict, bytes | None]:
    """``({part name: digest}, the main part's bytes)``: every part that can change what is
    drawn but the main part, which :func:`segments` reads."""
    out = {}
    main = None
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        for info in archive.infolist():
            name = info.filename
            if name.endswith("/"):
                continue
            if name == MAIN:
                main = archive.read(name)
            elif not _UNDRAWN.match(name):
                out[name] = _digest(archive.read(name))
    return out, main


_RELATIONSHIP = re.compile(rb"<(?:\w+:)?Relationship\b[^>]*>")
_ATTRIBUTE = re.compile(rb'(\w+)\s*=\s*"([^"]*)"')


def relationships(data: bytes | None) -> dict:
    """The main part's relationships, ``id -> (type, target, mode)``, from their bytes."""
    out = {}
    for element in _RELATIONSHIP.findall(data or b""):
        attributes = dict(_ATTRIBUTE.findall(element))
        out[attributes.get(b"Id")] = (attributes.get(b"Type"), attributes.get(b"Target"),
                                      attributes.get(b"TargetMode"))
    return out


def parts_reason(before: dict, after: dict) -> str | None:
    """Why the parts ``after`` differ from ``before`` in a way an edit of the body cannot be
    scoped past (:data:`FALLBACKS`), or ``None``: only parts the main part's new
    relationships add, or the main part's relationships, gained."""
    found = parts_cause(before, after)
    return found[0] if found is not None else None


def parts_cause(before: dict, after: dict) -> tuple[str, str] | None:
    """:func:`parts_reason`, with the part that differs: ``(reason, part name)``."""
    for name in sorted(set(before) | set(after)):
        if before.get(name) == after.get(name) or name == MAIN_RELS:
            continue
        if name not in before and name not in (MAIN_RELS,):
            # A new part: only an edited block can refer to it (a new relationship).
            if any(pattern.match(name) for pattern, _ in _PART_REASONS):
                return next(reason for pattern, reason in _PART_REASONS if pattern.match(name)), name
            continue
        if name == "[Content_Types].xml":
            continue  # each kept part's type is compared below
        for pattern, reason in _PART_REASONS:
            if pattern.match(name):
                return reason, name
        return "parts", name
    return None


# -- what the flow measured ------------------------------------------------------------------------


def _texts(paragraphs) -> bytes:
    """A digest of the paragraphs' text as the model holds it: what the layout sets in a
    block beyond its own XML -- a footnote's number, a ``SEQ`` or ``REF`` field's value, a
    field's result left out where its instruction began in another block -- changes it."""
    return _digest("\x00".join("\x01".join(run.text for run in paragraph.runs)
                                for paragraph in paragraphs).encode("utf-8", "surrogatepass"))


class Measures:
    """What :func:`docx2svg.paginate.flow` measured, by what it measured it from: a
    top-level paragraph's ``(pieces, lines, shares, heights, more)`` by its element's
    digest as measured (:attr:`Segment.measure`), its place in that element, its section
    and its list label (``more``: what follows from the lines alone, found once); a
    table's ``(flow, mode, approximation)`` by its element's whole digest -- its flow holds
    its cells' paragraphs, their paths and ids -- its path, its section, its cells' labels
    and features, its place and the style of the paragraph before it.  ``previous`` is the layout's before
    (taken from), ``kept`` this layout's (every entry it used or made: the next layout's
    ``previous``)."""

    def __init__(self, previous: dict | None = None) -> None:
        self.previous = previous or {}
        self.kept: dict = {}
        #: ``id(block) -> (digest, place)`` for the flow's top-level blocks: as measured
        #: (:attr:`Segment.measure`), what measurements are kept by, and whole.
        self.keys: dict[int, tuple] = {}
        self.blocks: dict[int, tuple] = {}
        self.elements: list = []
        self.hits = 0

    def bind(self, document, digests: dict) -> None:
        """Key the top-level blocks of ``document`` (the document the flow lays out) by
        ``digests``: the element key (``w:body/w:p[3]``) -> its segment's digest."""
        self.keys = {}
        self.blocks = {}
        #: Per top-level block, in order: the body element it is in (``w:body/w:p[3]``), or
        #: ``None`` (an endnote's paragraph, a block whose element was not read).
        self.elements = []
        seen: dict = {}
        element: str | tuple[str, str]
        for block in document.body:
            path = block.path
            steps = path.split("/")
            if steps[0] == "w:body" and len(steps) > 1:
                element = "/".join(steps[:2])
                found = digests.get(element)
                if found is None:
                    self.elements.append(None)
                    continue
                digest, measured = found
            else:
                # An endnote's paragraph, placed in the flow: its part is compared whole.
                element = digest = measured = ("note", path)
            place = seen.get(element, 0)
            seen[element] = place + 1
            self.keys[id(block)] = (measured, place)
            self.blocks[id(block)] = (digest, place)
            self.elements.append(element if isinstance(element, str) else None)

    def paragraph_key(self, paragraph, section: int, label):
        base = self.keys.get(id(paragraph))
        if base is None or paragraph.joins:
            # A paragraph joined to the next (its mark deleted) is more than its element.
            return None
        return ("p", base, section, label, _texts([paragraph]))

    def table_key(self, table, section: int, block: int, labels: dict, features, before):
        # A table's flow holds its model -- its cells' paragraphs, their paths and ids, which
        # the page draws -- so it is kept by its whole digest and its path.
        base = self.blocks.get(id(table))
        if base is None:
            return None
        from .model import Paragraph
        from .table import cell_blocks

        count = sum(1 for _ in cell_blocks(table))
        # The flow reads the paragraph before only for its style.
        previous = ((before.properties.style_id if before.properties else None,)
                    if isinstance(before, Paragraph) else None)
        from .layout import _all_paragraphs

        return ("t", base, table.path, section, block, _texts(_all_paragraphs([table])),
                tuple(labels.get(k) for k in range(block, block + count)),
                tuple(features[k] if features and k < len(features) else None for k in range(block, block + count)),
                previous)

    def get(self, key):
        found = self.kept.get(key)
        if found is None:
            found = self.previous.get(key)
            if found is not None:
                self.kept[key] = found
                self.hits += 1
        return found

    def put(self, key, value) -> None:
        self.kept[key] = value


# -- what a layout keeps ---------------------------------------------------------------------------


@dataclass
class State:
    """What a layout keeps so that the next version of its document can be laid out from it
    (:attr:`docx2svg.layout.Layout.state`): the bytes and what they were compared by, the
    pagination and what each of its pages placed, and what the flow measured.  Picklable,
    with the layout."""

    #: ``(docx2svg version, STATE_FORMAT)``: a state of another is not read.
    version: tuple
    #: What the layout was measured with besides the document (:func:`options_key`);
    #: ``None`` where it cannot be compared (the next layout is a whole one: ``options``).
    options: tuple | None
    data: bytes
    parts: dict
    relationships: dict
    #: The body's elements (:func:`segments`); ``None`` where they could not be read.
    segments: list | None
    #: Per top-level block of the flow: ``((digest, place), labels, features)``, and the
    #: body element it is in (:attr:`Measures.elements`).
    signatures: list
    elements: list
    #: Per item of the flow: its top-level block, and whether it holds a page's decision to
    #: end before the block after it (keep with next; a table, which may keep its rows).
    item_blocks: list
    chains: list
    #: The pagination (:class:`docx2svg.paginate.Pagination`), and per page it laid out
    #: ``(first page, end page, log, floats)``: the pages placed (blank ones before it
    #: included), the warnings placing them recorded (:attr:`Layout.log`), and how many
    #: floating drawings were placed before it.
    result: paginate.Pagination
    iterations: list
    #: :attr:`docx2svg.paginate.Flow.approximated`, in order.
    approximated: list
    #: :attr:`Measures.kept`.
    memo: dict
    #: The faces measured with (:class:`docx2svg.fonts.InstalledFonts`): the next layout's.
    fonts: InstalledFonts | None
    #: Whether a header or footer shows ``NUMPAGES`` or ``SECTIONPAGES``; the page count,
    #: and each section's pages (``SECTIONPAGES``).
    count_fields: bool
    page_count: int
    section_pages: dict
    #: Why the next layout from this one lays the whole document out again whatever
    #: changes (:data:`FALLBACKS`), or ``None``.
    blocked: str | None
    #: Each pass the layout made (:class:`Pass`), the last the one drawn: one, unless the
    #: body has page-number fields.
    passes: list = field(default_factory=list)
    #: Page (0-based) -> a digest of what it draws (:func:`drawing`), as found.
    drawings: dict = field(default_factory=dict)
    #: The 0-based pages this layout laid out again, made from a previous one (all of them
    #: where it laid out the whole document); with the pages kept but moved or drawn again.
    laid: range | None = None
    redrawn: frozenset = frozenset()
    #: The 0-based pages laid out again, exactly (:attr:`laid` spans them: from a previous
    #: layout, the stretches laid out again may be apart -- a table of contents and an edit).
    fresh: frozenset = frozenset()
    #: Page (0-based) -> the text of each ``STYLEREF`` its header and footer show
    #: (:meth:`docx2svg.layout._Placer.story_references_on`): a page kept by the next layout
    #: is drawn again where they differ.
    story_refs: dict = field(default_factory=dict)


def options_key(options) -> tuple | None:
    """What a conversion's ``options`` measure with besides the document, comparable;
    ``None`` where they measure with the caller's own ``advances``, ``metrics`` or
    ``decorations`` (which nothing here can compare)."""
    if options.advances is not None or options.metrics is not None or options.decorations is not None:
        return None
    from . import _user_dirs
    from .fonts import default_font_dirs

    dirs = tuple(str(path) for path in tuple(default_font_dirs()) + _user_dirs(options, subfolders=True))
    return (dirs, bool(options.substitute_fonts), tuple(sorted((options.font_substitutes or {}).items())),
            options.locale)


def _version() -> tuple:
    from . import __version__

    return (__version__, STATE_FORMAT)


def body_digests(segs: list[Segment] | None) -> dict:
    """``w:body/w:p[3]`` (the element key a block's path starts with) -> ``(its digest, its
    digest as measured)`` (:attr:`Segment.measure`)."""
    out: dict = {}
    counts: dict = {}
    for segment in segs or ():
        counts[segment.name] = counts.get(segment.name, 0) + 1
        out[f"w:body/w:{segment.name}[{counts[segment.name]}]"] = (segment.digest, segment.measure)
    return out


def _top_level(document) -> list[tuple[Paragraph | Table, int, int | None]]:
    """``(block, first block index, how many)`` of each top-level block of the flow."""
    out = []
    pairs = paginate._blocks_of_body(document)
    for k, (block, index) in enumerate(pairs):
        end = pairs[k + 1][1] if k + 1 < len(pairs) else None
        out.append((block, index, end))
    return out


def signatures(document, memo: Measures, labels: dict, features) -> tuple[list, dict]:
    """Per top-level block of the flow ``document``: what it is laid out from beyond the
    document's parts -- its element's digest and its place in it, its paragraphs' list
    labels and what the flow cannot model in them; and block index -> top-level block."""
    out = []
    ordinal: dict = {}
    blocks = _top_level(document)
    total = len(features) if features else 0
    for k, (block, first, end) in enumerate(blocks):
        end = end if end is not None else max(total, first + 1)
        indexes = range(first, end)
        for index in indexes:
            ordinal[index] = k
        out.append((memo.blocks.get(id(block)), tuple(labels.get(i) for i in indexes),
                    tuple(features[i] if features and i < len(features) else None for i in indexes)))
    return out, ordinal


def _item_blocks(items, ordinal: dict) -> list:
    return [ordinal.get(item.block, -1) for item in items]


def _chains(items) -> list:
    return [bool(getattr(item, "keep_next", False)) or isinstance(item, paginate.TableItem) for item in items]


def _story_fields(document, keywords: tuple = ("NUMPAGES", "SECTIONPAGES")) -> bool:
    """Whether a header or footer shows one of the fields ``keywords``."""
    from .layout import _all_paragraphs

    for blocks in document.stories.values():
        for paragraph in _all_paragraphs(blocks):
            if any(run.field is not None and run.field.keyword in keywords for run in paragraph.runs):
                return True
    return False


def blocked(placer, original, by_page: bool) -> str | None:
    """Why no layout from this document's can lay out only what an edit reaches
    (:data:`FALLBACKS`), from what it is: ``None`` where nothing stands in the way."""
    if by_page:
        return "footnote-restart"
    if any(section.line_numbering for section in original.sections):
        return "line-numbers"
    if any(paginate.is_multi_column(section) for section in original.sections):
        return "columns"
    if not paginate.CARRY_ON:
        return "stopped"
    return None


def record(placer, layout, data: bytes, parts: dict, main: bytes | None, fonts, options: tuple | None,
           approximated: list, original, by_page: bool) -> State:
    """The :class:`State` of a layout just made by ``placer``."""
    segs = placer.segments
    memo = placer.memo
    labels = getattr(placer.items, "labels", {})
    sigs, ordinal = signatures(placer.document, memo, labels, placer.features_used)
    result = placer.result
    reason = blocked(placer, original, by_page)
    if reason is None and result.stopped is not None:
        reason = "stopped"
    if reason is None and segs is None:
        reason = "body"
    if reason is None and options is None:
        reason = "options"
    section_pages = {k: placer._section_pages(result, k) for k in range(len(placer.document.sections))}
    passes = [Pass(found, block_texts(document), list(messages if messages is not None else
                                                     getattr(items, "approximated", ())),
                   _item_blocks(items, ordinal), _chains(items))
              for document, items, found, messages in getattr(placer, "passes", ())]
    if not passes or passes[-1].result is not result:
        reason = reason or "body"
    if any(getattr(items, "page_stories", None) for _, items, _, _ in getattr(placer, "passes", ())):
        reason = reason or "styleref-height"
    story_refs = {index: placer.story_references_on(index, page.info) for index, page in enumerate(layout.pages)}
    return State(_version(), options, data, parts, relationships(_read_rels(data)), segs, sigs,
                 list(memo.elements), _item_blocks(placer.items, ordinal), _chains(placer.items), result,
                 [entry[1:] for entry in placer.placed], list(approximated), dict(memo.kept), fonts,
                 _story_fields(placer.document), len(layout.pages), section_pages, reason, passes,
                 story_refs={k: v for k, v in story_refs.items() if v})


def _read_rels(data: bytes) -> bytes | None:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        try:
            return archive.read(MAIN_RELS)
        except KeyError:
            return None


# -- deciding -------------------------------------------------------------------------------------


@dataclass
class Plan:
    """What :func:`plan` found: the previous layout's state, and either why the whole
    document is laid out again (``reason``) or nothing standing in the way.  ``measured``:
    the previous layout's measurements may be reused (the parts they depend on are the
    same).  ``cause``: :attr:`docx2svg.layout.Layout.relaid_cause`."""

    state: State | None
    reason: str | None
    measured: bool
    parts: dict
    main: bytes | None
    segments: list | None
    cause: dict | None = None


_PARA_ID = re.compile(rb'\s(?:\w+:)?paraId\s*=\s*"([^"]*)"')
_FIELD_PIECE = re.compile(rb'fldCharType\s*=\s*["\'](begin|separate|end)["\']|'
                          rb'<(?:\w+:)?instrText\b[^>]*?(?:/>|>([^<]*)</(?:\w+:)?instrText>)')


def open_field(data: bytes) -> str | None:
    """The instruction of the first field the element ``data`` opens and does not close
    (``TOC \\o "1-3" \\h``), trimmed; ``None`` where it opens none (it closes one opened
    before it, or its fields are its own)."""
    stack: list[list[bytes]] = []
    for kind, text in _FIELD_PIECE.findall(data):
        if kind == b"begin":
            stack.append([])
        elif kind == b"end":
            if stack:
                stack.pop()
        elif not kind and stack:
            stack[-1].append(text)
    if not stack:
        return None
    from xml.sax.saxutils import unescape

    return " ".join(unescape(b"".join(stack[0]).decode("utf-8", "replace"), {"&quot;": '"', "&apos;": "'"}
                             ).split()) or None


def where(xml: bytes | None, index: int, with_field: bool = False) -> dict:
    """Where the body element ``index`` of ``xml`` (``word/document.xml``) is, for
    :attr:`docx2svg.layout.Layout.relaid_cause`: ``{"at": its path, "para_id": its first
    paragraph's w14:paraId}`` (the id where there is one), with ``"field"`` -- the
    instruction of the field it opens (:func:`open_field`) -- where asked for and found."""
    found = elements(xml) if xml is not None else None
    if found is None or not 0 <= index < len(found):
        return {}
    name, data = found[index]
    count = sum(1 for other, _ in found[:index + 1] if other == name)
    out: dict = {"at": f"w:body/w:{name}[{count}]"}
    ident = _PARA_ID.search(data)
    if ident is not None:
        out["para_id"] = ident.group(1).decode()
    if with_field:
        instruction = open_field(data)
        if instruction is not None:
            out["field"] = instruction
    return out


def _main_of(data: bytes | None) -> bytes | None:
    if data is None:
        return None
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        try:
            return archive.read(MAIN)
        except KeyError:
            return None


def plan(previous, data: bytes, options: tuple | None) -> Plan:
    """Compare the new bytes with what ``previous`` (a :class:`docx2svg.layout.Layout`)
    keeps."""
    parts, main = part_digests(data)
    segs = segments(main) if main is not None else None
    state = getattr(previous, "state", None)

    def fallback(reason: str, measured: bool, **cause) -> Plan:
        return Plan(state if isinstance(state, State) and reason not in ("no-state", "version") else None,
                    reason, measured, parts, main, segs, {"reason": reason, **cause})

    if not isinstance(state, State):
        return fallback("no-state", False)
    if state.version != _version():
        return fallback("version", False)
    if options is None or state.options != options:
        return fallback("options", False)
    found = parts_cause(state.parts, parts)
    if found is not None:
        return fallback(found[0], False, part=found[1])
    before, after = state.relationships, relationships(_read_rels(data))
    if any(key not in after or after[key] != value for key, value in before.items()):
        return fallback("parts", False, part=MAIN_RELS)
    if segs is None or state.segments is None:
        return fallback("body", True, part=MAIN)
    if state.blocked is not None:
        return fallback(state.blocked, True)
    old, new = state.segments, segs
    a = 0
    while a < min(len(old), len(new)) and old[a] == new[a]:
        a += 1
    b = 0
    while b < min(len(old), len(new)) - a and old[-1 - b] == new[-1 - b]:
        b += 1
    changed_old, changed_new = old[a:len(old) - b], new[a:len(new) - b]

    def at(test, with_field: bool = False) -> dict:
        # The first element of the new document that is why, else of the previous one.
        for k, segment in enumerate(changed_new):
            if test(segment):
                return where(main, a + k, with_field)
        for k, segment in enumerate(changed_old):
            if test(segment):
                return where(_main_of(state.data), a + k, with_field)
        return {}

    if any(segment.section for segment in changed_old + changed_new):
        return fallback("sections", True, **at(lambda segment: segment.section))
    if any(segment.joined for segment in changed_old + changed_new) or (a and old[a - 1].joined):
        place = at(lambda segment: segment.joined) or where(_main_of(state.data), a - 1)
        return fallback("joined", True, **place)
    if any(not segment.plain for segment in changed_old + changed_new):
        return fallback("fields-open", True, **at(lambda segment: not segment.plain, True))
    if [n for s in changed_old for n in s.notes] != [n for s in changed_new for n in s.notes]:
        return fallback("note-references", True, **at(lambda segment: bool(segment.notes)))
    return Plan(state, None, True, parts, main, segs)


def renumbering(old: list[Segment], new: list[Segment], old_element: str | None,
                new_element: str | None) -> dict:
    """``name -> (first old index renumbered, by how much)``: how the top-level steps of
    the paths of the blocks kept from a previous layout move (``w:body/w:p[40]`` becomes
    ``w:body/w:p[41]`` after a paragraph is inserted before it).  ``old_element`` and
    ``new_element`` are the body element the first kept block is in, before and after:
    every kept block is after it, and the elements before it are what moved them."""
    if old_element is None or new_element is None:
        return {}
    keys_old, keys_new = body_digests(old), body_digests(new)
    at_old, at_new = list(keys_old).index(old_element), list(keys_new).index(new_element)
    out = {}
    for name in {s.name for s in old} | {s.name for s in new}:
        before = sum(1 for s in old[:at_old] if s.name == name)
        shift = sum(1 for s in new[:at_new] if s.name == name) - before
        if shift:
            out[name] = (before + 1, shift)
    return out


_STEP = re.compile(r"w:body/w:([A-Za-z]+)\[(\d+)\]")


def renumbered(value, moves: dict, seen: dict | None = None):
    """``value`` -- a page, or anything a page holds -- with every body path's top-level
    step moved as ``moves`` (:func:`renumbering`) says, in its strings (paths, and the
    warnings that name them); a copy where anything changed, ``value`` itself where
    nothing did."""
    if not moves:
        return value
    seen = {} if seen is None else seen
    key = id(value)
    out: object
    if key in seen:
        return seen[key]

    def step(match: re.Match) -> str:
        name, index = match.group(1), int(match.group(2))
        move = moves.get(name)
        if move is None or index < move[0]:
            return match.group(0)
        return f"w:body/w:{name}[{index + move[1]}]"

    if isinstance(value, str):
        out = _STEP.sub(step, value) if "w:body/" in value else value
    elif isinstance(value, list):
        items = [renumbered(item, moves, seen) for item in value]
        out = items if any(a is not b for a, b in zip(items, value)) else value
    elif isinstance(value, tuple):
        members = tuple(renumbered(item, moves, seen) for item in value)
        out = members if any(a is not b for a, b in zip(members, value)) else value
    elif isinstance(value, dict):
        pairs = {renumbered(k, moves, seen): renumbered(v, moves, seen) for k, v in value.items()}
        out = pairs if any(a is not b or pairs[a] is not value[b] for a, b in zip(pairs, value)) else value
    elif dataclasses.is_dataclass(value) and not isinstance(value, type):
        changes = {}
        for item in dataclasses.fields(value):
            current = getattr(value, item.name)
            moved = renumbered(current, moves, seen)
            if moved is not current:
                changes[item.name] = moved
        extra = {k: v for k, v in vars(value).items() if k not in {f.name for f in dataclasses.fields(value)}} \
            if hasattr(value, "__dict__") else {}
        for name, current in extra.items():
            moved = renumbered(current, moves, seen)
            if moved is not current:
                changes[name] = moved
        if changes:
            import copy

            out = copy.copy(value)
            for name, moved in changes.items():
                object.__setattr__(out, name, moved)
        else:
            out = value
    else:
        out = value
    seen[key] = out
    return out


# -- drawings --------------------------------------------------------------------------------------

_DATA = re.compile(r' data-docx-[a-z-]+="[^"]*"')


def drawing(page, data: bytes, fonts) -> bytes:
    """A digest of what ``page`` draws: its SVG (``docx2svg.svg.render_page`` at the
    default glyph size, its pictures from ``data``) less its ``data-docx-*`` attributes,
    which say where things come from and draw nothing."""
    from . import _images
    from .svg import render_page

    svg = render_page(page, images=_images(data), warn=lambda *args: None,
                      names=fonts.drawing_name if fonts is not None else None)
    return _digest(_DATA.sub("", svg).encode())


# -- laying out again ------------------------------------------------------------------------------


class Relaid(Exception):
    """Raised by :meth:`Again.run` where it finds, part way, that the whole document must
    be laid out again: ``args[0]`` is why (:data:`FALLBACKS`)."""


def _common(old: list, new: list) -> tuple[int, int]:
    """The lengths of ``old`` and ``new``'s common start and (after it) common end."""
    a = 0
    while a < min(len(old), len(new)) and old[a] == new[a]:
        a += 1
    b = 0
    while b < min(len(old), len(new)) - a and old[-1 - b] == new[-1 - b]:
        b += 1
    return a, b


def _first(values: list, at_least: int) -> int:
    """The first index whose value is at least ``at_least`` (``len(values)`` for none)."""
    return next((k for k, value in enumerate(values) if value >= at_least), len(values))


def windows(old_sigs: list, sigs: list, old_texts: list, texts: list) -> list[tuple[int, int, int, int]]:
    """Where a pass of the new document differs from the same pass of the previous one, as
    ``(first, end, first before, end before)`` top-level block ordinals, in order: the
    blocks an edit changed (between the common start and end of the blocks'
    :func:`signatures`), and each block elsewhere whose text the layout set differs --
    a field's result (a ``TOC`` line's ``PAGEREF``, a ``PAGE``, a ``SEQ`` number, a ``REF``)
    -- each with its neighbours (laid out with it: borders, spacing, the style before a
    table), those that touch joined."""
    n, o = len(sigs), len(old_sigs)
    a, b = _common(old_sigs, sigs)
    offset = n - o
    marks: list[tuple[int, int]] = []
    edit = a < n - b or a < o - b
    if edit:
        marks.append((a, n - b))
    for index in range(a):
        if old_texts[index] != texts[index]:
            marks.append((index, index + 1))
    for index in range(max(n - b, a), n):
        if old_texts[index - offset] != texts[index]:
            marks.append((index, index + 1))
    joined: list[list[int]] = []
    for low, high in sorted(marks):
        low, high = max(low - 1, 0), min(high + 1, n)
        if joined and low <= joined[-1][1]:
            joined[-1][1] = max(joined[-1][1], high)
        else:
            joined.append([low, high])
    out = []
    for low, high in joined:
        if edit and low <= a and high >= n - b:
            # The window holding the edit: before it as it was, after it moved.
            out.append((low, high, low, high - offset))
        elif high <= a:
            out.append((low, high, low, high))
        else:
            out.append((low, high, low - offset, high - offset))
    return out


@dataclass
class Window:
    """One stretch of the flow a pass lays out again (:func:`windows`, as flow items): from
    ``changed`` (the first item) to ``suffix_new`` (the first after it that is unchanged),
    ``suffix_old`` that item in the previous pass; ``shift`` how many items the items after
    it moved by, ``offset`` how many top-level blocks."""

    changed: int
    suffix_new: int
    suffix_old: int
    shift: int
    offset: int


@dataclass
class Pass:
    """One pass of a layout (:meth:`docx2svg.layout._Placer.run_full` lays a body with
    page-number fields out again until their values settle): its pagination, per top-level
    block a digest of the text the pass laid out (:func:`block_texts`: the fields' values),
    the flow's approximations, per flow item its top-level block, and per item whether it
    holds a page's decision to end before the block after it."""

    result: paginate.Pagination
    texts: list
    approximated: list
    item_blocks: list
    chains: list


def block_texts(document) -> list:
    """Per top-level block of the flow ``document``: a digest of its paragraphs' text as the
    model holds it (:func:`_texts`): what a pass sets in a block beyond its XML."""
    from .layout import _all_paragraphs

    return [_texts(_all_paragraphs([block])) for block, _, _ in _top_level(document)]


@dataclass
class Spliced:
    """A pass's pagination made from the previous layout's same pass (:meth:`Again._splice`):
    the whole pagination, the flow's approximations over it, and how it is made, in order:
    ``("kept", first, end, shift, delta, offset, new first)`` -- the previous pass's
    iterations ``first`` to ``end``, their flow items moved by ``shift`` items (``offset``
    top-level blocks) and their pages by ``delta``, at the new iteration ``new first`` --
    or ``("fresh", first, end)``, the iterations laid out again."""

    result: paginate.Pagination
    messages: list
    segments: list
    windows: list


class Again:
    """One layout of a document made again from a previous layout's :class:`State`
    (:meth:`docx2svg.layout._Placer.run`).

    It goes through the passes a full layout makes (:meth:`docx2svg.layout._Placer.run_full`:
    the body laid out with its fields' cached results, then again with the page numbers
    found, until they are the ones it was laid out with, four passes at most): each pass
    is laid out from the same pass of the previous layout (the first from the first, every
    later one from the last, the one drawn), only where it differs (:func:`windows`) -- each
    stretch from the page holding its first block until a page starts where a page of the
    previous pass did, in the same state -- so that each pass's pagination, and the values
    its fields find, are a full layout's.  Of the last pass the pages laid out again are
    placed, the others taken from the previous layout.  After :meth:`run`,
    ``placed_pages`` holds the 0-based pages placed again and ``laid`` spans them."""

    def __init__(self, plan: Plan, previous) -> None:
        assert plan.state is not None, "a plan with no reason holds the previous state"
        self.plan = plan
        self.state: State = plan.state
        self.previous = previous
        self.placed_pages: frozenset = frozenset()
        self.laid = range(0)
        #: Pages kept from the previous layout but moved, or with their headers and
        #: footers drawn again: compared, as the pages placed again are.
        self.redrawn: set = set()
        #: How many passes the layout made, and how many stretches of the drawn one it laid
        #: out again (:attr:`docx2svg.layout.Layout.relaid` is ``None``).
        self.passes = 0
        self.stretches = 0

    def run(self, placer, source, original, features, by_page):
        from .layout import substitute_document
        from .notes import with_footnotes

        state = self.state
        reason = blocked(placer, original, by_page)
        if reason is not None:
            raise Relaid(reason)
        placer.features_used = features
        memo = placer.memo
        document = with_footnotes(original)
        values: dict = {}
        passes = state.passes
        last = len(passes) - 1
        placer.passes = []
        layout = placer.layout
        spliced = None
        for attempt in range(4):
            memo.bind(document, placer.digests)
            items = paginate.flow(document, placer.advances, placer.metrics, features=features, memo=memo)
            if not items or items.columns:
                raise Relaid("columns" if items and items.columns else "body")
            base = 0 if attempt == 0 else last
            spliced = self._splice(document, items, memo, features, passes[base])
            placer.document, placer.items = document, items
            placer.passes.append((document, items, spliced.result, spliced.messages))
            found = placer._body_field_values(items, spliced.result)
            if placer._story_heights(items, spliced.result):
                # A full layout's next pass lays some page out with its own header height.
                raise Relaid("styleref-height")
            self.passes = attempt + 1
            if found == values:
                break
            values = found
            document = substitute_document(with_footnotes(original), values)
        else:
            layout.warn("field-layout-unsettled", "the body's page-number fields did not settle on one "
                        "layout in four passes; the last is drawn")
        drawn_document, items = placer.document, placer.items
        if base != last:
            # The pass drawn was laid out from a pass of the previous layout that was not
            # drawn (the first, where the body's fields were all removed): from the one drawn.
            memo.bind(drawn_document, placer.digests)
            items = paginate.flow(drawn_document, placer.advances, placer.metrics, features=features, memo=memo)
            spliced = self._splice(drawn_document, items, memo, features, passes[last])
            placer.items = items
            placer.passes[-1] = (drawn_document, items, spliced.result, spliced.messages)
        return self._place(placer, drawn_document, items, spliced, document)

    # -- one pass --------------------------------------------------------------------------

    def _splice(self, document, items, memo: Measures, features, before: Pass) -> Spliced:
        """The pagination of ``document`` (whose flow is ``items``), from ``before``: its
        pages where nothing on them or before them differs, laid out again elsewhere."""
        state = self.state
        sigs, ordinal = signatures(document, memo, getattr(items, "labels", {}), features)
        texts = block_texts(document)
        blocks = _item_blocks(items, ordinal)
        old_blocks = before.item_blocks
        if len(state.signatures) != len(before.texts) or len(sigs) != len(texts):
            raise Relaid("body")
        found = []
        shift = 0
        for low, high, low_old, high_old in windows(state.signatures, sigs, before.texts, texts):
            changed, suffix_new = _first(blocks, low), _first(blocks, high)
            changed_old, suffix_old = _first(old_blocks, low_old), _first(old_blocks, high_old)
            if changed - changed_old != shift or changed > suffix_new:
                raise Relaid("body")
            shift = suffix_new - suffix_old
            found.append(Window(changed, suffix_new, suffix_old, shift, high - high_old))
        if len(items) - len(old_blocks) != shift:
            raise Relaid("body")
        result = before.result
        starts = result.starts
        if not starts:
            raise Relaid("body")
        at = {start: j for j, start in enumerate(starts)}
        movable = _movable(result, document)
        segments: list = []
        pieces: list = []
        j = shift = delta = offset = 0
        w = 0
        while w < len(found):
            window = found[w]
            changed_old = window.changed - shift
            # The page to lay out again from: the one holding the first changed item -- or one
            # before, where that page starts inside a paragraph or table, or where its end may
            # be decided by the change (the change starts the page, or a chain of
            # keep-with-next blocks from its start reaches it) -- and never one before where
            # the stretch before lined up (its pages up to there are this document's).
            k = max((i for i, start in enumerate(starts) if start.item <= changed_old), default=0)
            while k > j:
                start = starts[k]
                # Inside a paragraph only where its lines were the flow's, there as here (one
                # broken beside a drawing on the page before is laid out with that page).
                inside = start.part or (start.line and not result.states[k][4])
                if inside or all(before.chains[i] for i in range(start.item, changed_old)):
                    k -= 1
                    continue
                break
            k = max(k, j)
            if k > j:
                segments.append(("kept", j, k, shift, delta, offset))
            page, number, carry = result.states[k][:3]
            start = starts[k]
            resume = (Position(start.item + shift, start.line, start.part), page + delta, number + delta, carry)

            def lines_up(position, page, number, carry, w=w, k=k) -> bool:
                # A page starts at the start of an unchanged block past the stretch (and not
                # inside a later one), where a page of the previous pass started, with the same
                # footnote carried over -- on the same page, or on another where the pages
                # after it can move by as many (the page numbers going on with them, the same
                # headers and footers).
                if position.part or position.item < found[w].suffix_new:
                    return False
                v = w
                while v + 1 < len(found) and found[v + 1].changed <= position.item:
                    v += 1
                if position.item < found[v].suffix_new:
                    return False
                old = Position(position.item - found[v].shift, position.line)
                jj = at.get(old)
                if position.line and not (jj is not None and result.states[jj][4]
                                          and paginate._pristine(items[position.item])):
                    return False
                if jj is None or jj <= k:
                    return False
                old_page, old_number, old_carry = result.states[jj][:3]
                if old_carry != carry or page - old_page != number - old_number:
                    return False
                return page == old_page or movable[jj](page - old_page)

            again = paginate.paginate(document, items, resume=resume, until=lines_up)
            if again.stopped is not None:
                raise Relaid("stopped")
            pieces.append((again, len(getattr(items, "approximated", ()))))
            segments.append(("fresh", len(pieces) - 1))
            if again.until_at is None:
                j = len(starts)
                break
            position = again.until_at[0]
            v = w
            while v + 1 < len(found) and found[v + 1].changed <= position.item:
                v += 1
            shift, offset = found[v].shift, found[v].offset
            j = at[Position(position.item - shift, position.line)]
            delta = again.until_at[1] - result.states[j][0]
            w = v + 1
        else:
            if j < len(starts):
                segments.append(("kept", j, len(starts), shift, delta, offset))
        return self._merge(before, items, segments, pieces, found)

    def _merge(self, before: Pass, items, segments: list, pieces: list, found: list) -> Spliced:
        """The pagination of the whole document: ``before``'s iterations where ``segments``
        keeps them -- their flow positions moved by their ``shift`` items, their pages and
        page numbers by ``delta`` -- and ``pieces``' (the paginations laid out again) between,
        with the flow's approximations over them, in order."""
        from .paginate import Pagination

        result = before.result
        old_messages = before.approximated
        counts = [state[3] for state in result.states] + [len(old_messages)]
        new_messages = list(getattr(items, "approximated", ()))
        first_piece = pieces[0][0] if pieces else None
        flow_end = first_piece.states[0][3] if first_piece is not None and first_piece.states else (
            pieces[0][1] if pieces else len(new_messages))
        messages = list(new_messages[:flow_end])
        merged = Pagination()
        states: list = []
        starts: list = []
        ends: list = []
        info: dict = {}
        pages: dict = {}
        keyed: dict = {"bands": {}, "wraps": {}, "notes": {}, "frames": {}}
        out_segments: list = []

        def add(state, chunk, by: int = 0) -> None:
            states.append((state[0] + by, state[1] + by, state[2], len(messages), state[4]))
            messages.extend(message for message in chunk if message not in messages)

        for segment in segments:
            if segment[0] == "fresh":
                piece, end_count = pieces[segment[1]]
                piece_counts = [state[3] for state in piece.states] + [end_count]
                out_segments.append(("fresh", len(states), len(states) + len(piece.states)))
                for index in range(len(piece.states)):
                    add(piece.states[index], new_messages[piece_counts[index]:piece_counts[index + 1]])
                starts.extend(piece.starts)
                ends.extend(piece.ends)
                info.update(piece.info)
                for name, values in keyed.items():
                    values.update(getattr(piece, name))
                for index, found_pages in piece.pages.items():
                    pages.setdefault(index, []).extend(found_pages)
                continue
            _, first, end, shift, delta, offset = segment
            out_segments.append(("kept", first, end, shift, delta, offset, len(states)))
            for index in range(first, end):
                add(result.states[index], old_messages[counts[index]:counts[index + 1]], delta)
                start = result.starts[index]
                starts.append(Position(start.item + shift, start.line, start.part))
                if index < len(result.ends):
                    stop = result.ends[index]
                    ends.append(Position(stop.item + shift, stop.line, stop.part))
            low = result.states[first][0]
            high = result.states[end][0] if end < len(result.states) else None
            for page, value in result.info.items():
                if page >= low and (high is None or page < high):
                    info[page + delta] = dataclasses.replace(value, number=value.number + delta) if delta else value
            for name, values in keyed.items():
                for page, value in getattr(result, name).items():
                    if page >= low and (high is None or page < high):
                        values[page + delta] = value
            last_item = result.starts[end].item if end < len(result.starts) else None
            for index, found_pages in result.pages.items():
                if index < result.starts[first].item or (last_item is not None and index > last_item):
                    continue
                kept = [page + delta for page in found_pages if page >= low and (high is None or page < high)]
                if kept:
                    pages.setdefault(index + shift, []).extend(kept)
        merged.states = states
        merged.starts = starts
        merged.ends = ends
        merged.info = dict(sorted(info.items()))
        merged.pages = dict(sorted(pages.items()))
        for name, values in keyed.items():
            setattr(merged, name, dict(sorted(values.items())))
        return Spliced(merged, messages, out_segments, found)

    # -- the pass drawn ----------------------------------------------------------------------

    def _place(self, placer, document, items, spliced: Spliced, last_document):
        """Place the pages of the pass drawn that it laid out again, take the others from the
        previous layout (moved, their paths renumbered, their headers and footers drawn again
        where their numbers changed), and record what each page warned of, in page order."""
        state = self.state
        merged = spliced.result
        layout = placer.layout
        placer.result = merged
        placer.field_values, placer.footnote_pages = {}, None
        placer.references = paginate.story_references(last_document)
        placer.approximated = spliced.messages
        for message in spliced.messages:
            layout.approximate("pagination", message, None, None)
        placer._notes(items)
        details = placer._details(items, lazy=True)
        segments = spliced.segments
        self.stretches = sum(1 for segment in segments if segment[0] == "fresh")
        # The stretches laid out again, placed first, each with its warnings kept aside:
        # they are recorded in page order once the pages before them are.  ``floats``: the
        # floating drawings placed before each stretch, as a full layout counts them.
        floats = 0
        placed: dict = {}
        fresh_pages: set = set()
        for n, segment in enumerate(segments):
            if segment[0] == "kept":
                _, first, end, *_ = segment
                placed[n] = floats - state.iterations[first][3]
                if end < len(state.iterations):
                    floats += state.iterations[end][3] - state.iterations[first][3]
                continue
            _, first, end = segment
            first_page = merged.states[first][0]
            placer.__dict__["_float_count"] = [floats]
            placer.placed = []
            saved, layout.pages = layout.pages, []
            with _captured(layout):
                end_page = placer.place_pages(items, details, merged, range(first, end), first_page)
            placed[n] = (layout.pages, placer.placed)
            layout.pages = saved
            floats = placer.__dict__["_float_count"][0]
            fresh_pages.update(range(first_page, end_page))
            if end < len(merged.states) and merged.states[end][0] != end_page:
                raise Relaid("body")
        self.placed_pages = frozenset(fresh_pages)
        self.laid = range(min(fresh_pages), max(fresh_pages) + 1) if fresh_pages else range(0)
        last = segments[-1] if segments else None
        if last is not None and last[0] == "kept":
            total = state.page_count + last[4]
        else:
            total = max(fresh_pages) + 1 if fresh_pages else state.page_count
        sections = {s: placer._section_pages(merged, s) for s in range(len(document.sections))}
        counted = (state.count_fields or _story_fields(document)) and (
            total != state.page_count or sections != state.section_pages)
        numbered = _story_fields(document, ("PAGE", "NUMPAGES", "SECTIONPAGES"))
        layout.pages = []
        placer.placed = []
        for n, segment in enumerate(segments):
            if segment[0] == "fresh":
                fresh, entries = placed[n]
                layout.pages.extend(fresh)
                for entry in entries:
                    layout.replay(entry[3])
                    placer.placed.append(entry)
                continue
            _, first, end, shift, delta, offset, new_first = segment
            moves = self._moves(placer.memo.elements, first, offset) if n else {}
            floats_offset = placed[n]
            for index in range(first, end):
                first_page, end_page, log, floats_before, marks = state.iterations[index]
                pages = [renumbered(page, moves) for page in self.previous.pages[first_page:end_page]]
                log = renumbered(log, moves)
                # A STYLEREF in its header or footer whose text there changed (round 21).
                styled = state.story_refs.get(end_page - 1, ()) != placer.story_references_on(
                    end_page - 1 + delta, merged.info.get(end_page - 1 + delta))
                if delta or counted or styled:
                    self.redrawn.update(range(first_page + delta, end_page + delta))
                    if delta and not numbered and not counted and not styled:
                        pages, log = _moved(pages, delta), _moved_log(log, delta)
                    else:
                        pages, log, marks = self._restoried(placer, pages, log, marks, delta)
                layout.pages.extend(pages)
                layout.replay(log)
                placer.placed.append((new_first + index - first, first_page + delta, end_page + delta, log,
                                      floats_before + floats_offset, marks))
        return layout

    def _moves(self, memo_elements: list, first: int, offset: int) -> dict:
        """:func:`renumbering` for the pages kept from iteration ``first`` of the previous
        pass drawn on, whose top-level blocks moved by ``offset``: by the first block on them
        that is in a body element."""
        state = self.state
        start = state.result.starts[first]
        found = state.item_blocks[start.item] if start.item < len(state.item_blocks) else len(state.elements)
        for ordinal in range(max(found, 0), len(state.elements)):
            old = state.elements[ordinal]
            new = memo_elements[ordinal + offset] if 0 <= ordinal + offset < len(memo_elements) else None
            if old is not None and new is not None:
                # Both read: a plan with no reason compared them (:func:`plan`).
                assert state.segments is not None and self.plan.segments is not None
                return renumbering(state.segments, self.plan.segments, old, new)
            if (old is None) != (new is None):
                raise Relaid("body")
        return {}

    def _restoried(self, placer, pages: list, log: list, marks, delta: int) -> tuple:
        """``pages`` -- one iteration's, kept from the previous layout -- moved ``delta``
        pages on, and their headers and footers drawn again (a page number or the page
        count they show changed); with the iteration's warnings and marks to match."""
        import copy

        pages, log = _moved(pages, delta), _moved_log(log, delta)
        page = pages[-1]
        if page.info is None or page.info.blank or marks is None:
            return pages, log, marks
        before, after = marks
        if before[6] != after[6]:
            raise Relaid("page-count")  # its header or footer made drawing definitions
        again = copy.copy(page)
        tails = {}
        for name, start, end in zip(("lines", "rules", "pictures", "floats", "boxes"), before[:5], after[:5]):
            items = getattr(page, name)
            setattr(again, name, list(items[:start]))
            tails[name] = items[end:]
        layout = placer.layout
        with _captured(layout) as story:
            placer._draw_stories(again, again.info)
        from .layout import _marks

        drawn = _marks(again, log[:before[5]] + story)
        if drawn[6] != before[6]:
            raise Relaid("page-count")
        for name, tail in tails.items():
            getattr(again, name).extend(tail)
        return pages[:-1] + [again], log[:before[5]] + story + log[after[5]:], (before, drawn)


def changed_pages(previous, layout, again: "Again | None", data: bytes, fonts, reason: str | None) -> list[int]:
    """The 1-based pages of ``layout`` whose drawing differs from ``previous``'s page of the
    same number (:func:`drawing`), and those ``previous`` did not have.  ``again``: how the
    layout was made from ``previous`` -- the pages it did not lay out again or draw again
    are ``previous``'s own, or copies of them with their paths renumbered, which draw the
    same; ``None`` where every page was laid out."""
    placed = None if again is None else set(again.placed_pages) | again.redrawn
    old_pages = previous.pages
    state = getattr(previous, "state", None)
    old_data = state.data if isinstance(state, State) else None
    old_fonts = state.fonts if isinstance(state, State) else fonts
    old_drawings = state.drawings if isinstance(state, State) else {}
    drawings = layout.state.drawings if isinstance(layout.state, State) else {}
    # A picture or another part may have changed under an equal page: compare drawings.
    alike = reason not in ("parts", "no-state", "version")
    out = []
    for index, page in enumerate(layout.pages):
        if index >= len(old_pages):
            out.append(index + 1)
            continue
        old = old_pages[index]
        if page is old or (placed is not None and index not in placed):
            if index in old_drawings:
                drawings[index] = old_drawings[index]
            continue
        if old_data is None:
            out.append(index + 1)
            continue
        if alike and page == old:
            continue
        if index not in old_drawings:
            old_drawings[index] = drawing(old, old_data, old_fonts)
        if index not in drawings:
            drawings[index] = drawing(page, data, fonts)
        if drawings[index] != old_drawings[index]:
            out.append(index + 1)
    return out


def _movable(before, document) -> list:
    """Per iteration of ``before``'s pagination: whether the pages from it on can move by
    ``delta`` pages (``movable[j](delta)``) and be what a layout makes of them there: none
    of them starts a section or is a blank page Word inserts, and moving them keeps every
    page's headers and footers (an even number of pages, where even and odd pages have
    their own)."""
    out: list[Callable[[int], bool] | None] = [None] * len(before.starts)
    plain = True
    for j in range(len(before.starts) - 1, -1, -1):
        page = before.states[j][0]
        end = before.states[j + 1][0] if j + 1 < len(before.starts) else (max(before.info) + 1 if before.info else page)
        for number in range(page, end):
            info = before.info.get(number)
            if info is None or info.blank or info.section_first:
                plain = False
        def movable(delta: int, ok: bool = plain) -> bool:
            return ok and (delta % 2 == 0 or not document.even_and_odd_headers)

        out[j] = movable
    return out


def _moved(pages: list, delta: int) -> list:
    """``pages`` moved ``delta`` pages on: their places and the numbers they show."""
    if not delta:
        return pages
    import copy

    out = []
    for page in pages:
        moved = copy.copy(page)
        moved.number = page.number + delta
        if page.info is not None:
            moved.info = dataclasses.replace(page.info, number=page.info.number + delta)
        out.append(moved)
    return out


def _moved_log(log: list, delta: int) -> list:
    """A page's warnings (:attr:`Layout.log`) moved ``delta`` pages on."""
    if not delta:
        return log
    out = []
    for kind, *args in log:
        # The page is the third argument of each kind (``warn``, ``approximate``, ``placeholder``).
        if args[2] is not None:
            args[2] += delta
        out.append((kind, *args))
    return out


class _captured:
    """While it is entered, what a layout warns of goes to lists of its own (the ``as``
    value is its :attr:`Layout.log`), so it can be recorded later, in order."""

    def __init__(self, layout) -> None:
        self.layout = layout

    def __enter__(self) -> list:
        layout = self.layout
        self.saved = (layout.warnings, layout.approximations, layout.placeholders, layout.log)
        layout.warnings, layout.approximations, layout.placeholders, layout.log = [], [], [], []
        return layout.log

    def __exit__(self, *exc) -> None:
        layout = self.layout
        layout.warnings, layout.approximations, layout.placeholders, layout.log = self.saved
