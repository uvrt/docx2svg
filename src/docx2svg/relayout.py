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
3. **lays the pages out again from the page holding the first changed block** -- or the
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

The result is the layout a full layout of the new bytes makes, page for page, byte for
byte in every page's SVG (``tests/test_relayout.py`` holds it to that over random edits).
Where an edit cannot be scoped, the whole document is laid out again -- with the
measurements of unchanged blocks still reused where the document's styles and settings
are the same -- and :attr:`Layout.relaid` says why: see :data:`FALLBACKS`.
"""

from __future__ import annotations

import dataclasses
import hashlib
import io
import re
import zipfile
from dataclasses import dataclass, field

from . import paginate

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
    "page-fields": "the body has page-number fields (PAGE, NUMPAGES, SECTIONPAGES, PAGEREF: a TOC's page "
                   "numbers): their values depend on the whole layout",
    "cross-references": "the body has SEQ or REF fields, counted over the whole body",
    "footnote-restart": "footnote numbers restart on each page",
    "line-numbers": "a section numbers its lines",
    "columns": "a section has several text columns",
    "stopped": "the layout stops (CARRY_ON is off)",
    "fields-open": "an edited block opens or closes a field another block holds",
    "joined": "an edited block's paragraph mark is deleted as a revision (it joins the paragraph after it)",
    "body": "the body's elements could not be read for comparing",
    "page-count": "the page count changed and a header or footer shows NUMPAGES or SECTIONPAGES",
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
STATE_FORMAT = 1


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
    out: list[Segment] = []
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
                out.append(_segment(name, body[begin:match.end()]))
        elif empty:
            if depth == 0:
                out.append(_segment(tag.split(b":")[-1].decode(), body[match.start():match.end()]))
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
    for name in sorted(set(before) | set(after)):
        if before.get(name) == after.get(name) or name == MAIN_RELS:
            continue
        if name not in before and name not in (MAIN_RELS,):
            # A new part: only an edited block can refer to it (a new relationship).
            if any(pattern.match(name) for pattern, _ in _PART_REASONS):
                return next(reason for pattern, reason in _PART_REASONS if pattern.match(name))
            continue
        if name == "[Content_Types].xml":
            continue  # each kept part's type is compared below
        for pattern, reason in _PART_REASONS:
            if pattern.match(name):
                return reason
        return "parts"
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
    #: What the layout was measured with besides the document (:func:`options_key`).
    options: tuple
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
    result: object
    iterations: list
    #: :attr:`docx2svg.paginate.Flow.approximated`, in order.
    approximated: list
    #: :attr:`Measures.kept`.
    memo: dict
    #: The faces measured with (:class:`docx2svg.fonts.InstalledFonts`): the next layout's.
    fonts: object
    #: Whether a header or footer shows ``NUMPAGES`` or ``SECTIONPAGES``; the page count,
    #: and each section's pages (``SECTIONPAGES``).
    count_fields: bool
    page_count: int
    section_pages: dict
    #: Why the next layout from this one lays the whole document out again whatever
    #: changes (:data:`FALLBACKS`), or ``None``.
    blocked: str | None
    #: Page (0-based) -> a digest of what it draws (:func:`drawing`), as found.
    drawings: dict = field(default_factory=dict)
    #: The 0-based pages this layout laid out again, made from a previous one (all of them
    #: where it laid out the whole document); with the pages kept but moved or drawn again.
    laid: range | None = None
    redrawn: frozenset = frozenset()


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


def _top_level(document) -> list[tuple[object, int, int]]:
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
    from .layout import PAGE_FIELDS, _all_paragraphs

    if by_page:
        return "footnote-restart"
    if any(section.line_numbering for section in original.sections):
        return "line-numbers"
    if any(paginate.is_multi_column(section) for section in original.sections):
        return "columns"
    if not paginate.CARRY_ON:
        return "stopped"
    for paragraph in _all_paragraphs(original.body):
        for run in paragraph.runs:
            if run.field is not None:
                if run.field.keyword in PAGE_FIELDS:
                    return "page-fields"
                if run.field.keyword in ("SEQ", "REF"):
                    return "cross-references"
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
    return State(_version(), options, data, parts, relationships(_read_rels(data)), segs, sigs,
                 list(memo.elements), _item_blocks(placer.items, ordinal), _chains(placer.items), result,
                 [entry[1:] for entry in placer.placed], list(approximated), dict(memo.kept), fonts,
                 _story_fields(placer.document), len(layout.pages), section_pages, reason)


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
    same)."""

    state: State | None
    reason: str | None
    measured: bool
    parts: dict
    main: bytes | None
    segments: list | None


def plan(previous, data: bytes, options: tuple | None) -> Plan:
    """Compare the new bytes with what ``previous`` (a :class:`docx2svg.layout.Layout`)
    keeps."""
    parts, main = part_digests(data)
    segs = segments(main) if main is not None else None
    state = getattr(previous, "state", None)
    if not isinstance(state, State):
        return Plan(None, "no-state", False, parts, main, segs)
    if state.version != _version():
        return Plan(None, "version", False, parts, main, segs)
    if options is None or state.options != options:
        return Plan(state, "options", False, parts, main, segs)
    reason = parts_reason(state.parts, parts)
    if reason is not None:
        return Plan(state, reason, False, parts, main, segs)
    before, after = state.relationships, relationships(_read_rels(data))
    if any(key not in after or after[key] != value for key, value in before.items()):
        return Plan(state, "parts", False, parts, main, segs)
    if segs is None or state.segments is None:
        return Plan(state, "body", True, parts, main, segs)
    if state.blocked is not None:
        return Plan(state, state.blocked, True, parts, main, segs)
    old, new = state.segments, segs
    a = 0
    while a < min(len(old), len(new)) and old[a] == new[a]:
        a += 1
    b = 0
    while b < min(len(old), len(new)) - a and old[-1 - b] == new[-1 - b]:
        b += 1
    changed_old, changed_new = old[a:len(old) - b], new[a:len(new) - b]
    if any(segment.section for segment in changed_old + changed_new):
        return Plan(state, "sections", True, parts, main, segs)
    if any(segment.joined for segment in changed_old + changed_new) or (a and old[a - 1].joined):
        return Plan(state, "joined", True, parts, main, segs)
    if any(not segment.plain for segment in changed_old + changed_new):
        return Plan(state, "fields-open", True, parts, main, segs)
    if [n for s in changed_old for n in s.notes] != [n for s in changed_new for n in s.notes]:
        return Plan(state, "note-references", True, parts, main, segs)
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
        items = tuple(renumbered(item, moves, seen) for item in value)
        out = items if any(a is not b for a, b in zip(items, value)) else value
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


class Again:
    """One layout of a document made again from a previous layout's :class:`State`
    (:meth:`docx2svg.layout._Placer.run`).  After :meth:`run`, ``placed_pages`` is the
    range of pages laid out again (the others are the previous layout's)."""

    def __init__(self, plan: Plan, previous) -> None:
        self.plan = plan
        self.state = plan.state
        self.previous = previous
        self.placed_pages = range(0)
        #: Pages kept from the previous layout but moved, or with their headers and
        #: footers drawn again: compared, as the pages placed again are.
        self.redrawn: set = set()

    def run(self, placer, source, original, features, by_page):
        from .notes import with_footnotes
        from .paginate import Position

        state = self.state
        reason = blocked(placer, original, by_page)
        if reason is not None:
            raise Relaid(reason)
        placer.features_used = features
        document = with_footnotes(original)
        memo = placer.memo
        memo.bind(document, placer.digests)
        items = paginate.flow(document, placer.advances, placer.metrics, features=features, memo=memo)
        placer.document, placer.items = document, items
        if not items or items.columns:
            raise Relaid("columns" if items and items.columns else "body")
        sigs, ordinal = signatures(document, memo, getattr(items, "labels", {}), features)
        old_sigs = state.signatures
        a, b = _common(old_sigs, sigs)
        # A block's neighbours are laid out with it (borders, spacing, a drop cap, the
        # style before a table): one more on each side.
        low = max(a - 1, 0)
        high_new = min(len(sigs) - b + 1, len(sigs))
        high_old = len(old_sigs) - (len(sigs) - high_new)
        blocks = _item_blocks(items, ordinal)
        changed = _first(blocks, low)
        suffix_new = _first(blocks, high_new)
        suffix_old = _first(state.item_blocks, high_old)
        shift = suffix_new - suffix_old
        if len(items) - suffix_new != len(state.item_blocks) - suffix_old or changed > suffix_new:
            raise Relaid("body")
        before = state.result
        starts = before.starts
        if not starts:
            raise Relaid("body")
        # The page to lay out again from: the one holding the first changed item -- or one
        # before, where that page starts inside a paragraph or table, or where its end may
        # be decided by the change (the change starts the page, or a chain of
        # keep-with-next blocks from its start reaches it).
        k = max((j for j, start in enumerate(starts) if start.item <= changed), default=0)
        while k > 0:
            start = starts[k]
            # Inside a paragraph only where its lines were the flow's, there as here (one
            # broken beside a drawing on the page before is laid out with that page).
            inside = start.part or (start.line and not before.states[k][4])
            if inside or all(state.chains[i] for i in range(start.item, changed)):
                k -= 1
                continue
            break
        page, number, carry = before.states[k][:3]
        at = {start: j for j, start in enumerate(starts)}
        movable = _movable(before, document)

        def lines_up(position, page, number, carry) -> bool:
            # A page starts at the start of an unchanged block, where a page of the
            # previous layout started, with the same footnote carried over -- on the same
            # page, or on another where the pages after it can move by as many (the page
            # numbers going on with them, the same headers and footers).
            if position.part or position.item < suffix_new:
                return False
            if position.line and not (before.states[at.get(Position(position.item - shift, position.line), 0)][4]
                                      and paginate._pristine(items[position.item])):
                return False
            j = at.get(Position(position.item - shift, position.line))
            if j is None or j <= k:
                return False
            old_page, old_number, old_carry = before.states[j][:3]
            if old_carry != carry or page - old_page != number - old_number:
                return False
            return page == old_page or movable[j](page - old_page)

        again = paginate.paginate(document, items, resume=(starts[k], page, number, carry), until=lines_up)
        if again.stopped is not None:
            raise Relaid("stopped")
        new_count = len(again.starts)
        if again.until_at is not None:
            j = at[Position(again.until_at[0].item - shift, again.until_at[0].line)]
            delta = again.until_at[1] - before.states[j][0]
        else:
            j, delta = len(starts), 0
        self.moves = self._moves(memo, sigs, items, shift, j) if j < len(starts) else {}
        merged, messages = self._merge(before, again, k, j, shift, delta, items)
        moves = self.moves
        placer.result = placer.result_for_fields = merged
        placer.references = paginate.story_references(document)
        placer.approximated = messages
        placer.field_values, placer.footnote_pages = {}, None
        layout = placer.layout
        for message in messages:
            layout.approximate("pagination", message, None, None)
        placer._notes(items)
        details = placer._details(items, lazy=True)
        first_page = before.states[k][0]
        # The pages the edit reaches, placed with their warnings kept aside: they are
        # recorded in page order once the pages before them are.
        placer.__dict__["_float_count"] = [state.iterations[k][3]]
        with _captured(layout):
            end_page = placer.place_pages(items, details, merged, range(k, k + new_count), first_page)
        placed, layout.pages, fresh = placer.placed, [], layout.pages
        placer.placed = []
        self.placed_pages = range(first_page, end_page)
        if j < len(starts) and before.states[j][0] + delta != end_page:
            raise Relaid("body")
        total = end_page + (state.page_count - before.states[j][0] if j < len(starts) else 0)
        sections = {s: placer._section_pages(merged, s) for s in range(len(document.sections))}
        counted = (state.count_fields or _story_fields(document)) and (
            total != state.page_count or sections != state.section_pages)
        numbered = _story_fields(document, ("PAGE", "NUMPAGES", "SECTIONPAGES"))
        for index in range(k):
            first, end, log, floats, marks = state.iterations[index]
            pages = self.previous.pages[first:end]
            if counted:
                pages, log, marks = self._restoried(placer, pages, log, marks, 0)
                self.redrawn.update(range(first, end))
            layout.pages.extend(pages)
            layout.replay(log)
            placer.placed.append((index, first, end, log, floats, marks))
        layout.pages.extend(fresh)
        for entry in placed:
            layout.replay(entry[3])
            placer.placed.append(entry)
        for index in range(j, len(starts)):
            first, end, log, floats, marks = state.iterations[index]
            pages = [renumbered(page, moves) for page in self.previous.pages[first:end]]
            log = renumbered(log, moves)
            if delta or counted:
                self.redrawn.update(range(first + delta, end + delta))
                if delta and not numbered and not counted:
                    pages, log = _moved(pages, delta), _moved_log(log, delta)
                else:
                    pages, log, marks = self._restoried(placer, pages, log, marks, delta)
            layout.pages.extend(pages)
            layout.replay(log)
            placer.placed.append((k + new_count + index - j, first + delta, end + delta, log, floats, marks))
        return layout

    def _moves(self, memo: Measures, sigs: list, items, shift: int, j: int) -> dict:
        """:func:`renumbering` for the pages kept from iteration ``j`` of the previous
        pagination on: by the first block on them that is in a body element."""
        state = self.state
        start = state.result.starts[j]
        first = state.item_blocks[start.item] if start.item < len(state.item_blocks) else len(state.elements)
        offset = len(sigs) - len(state.signatures)
        for ordinal in range(max(first, 0), len(state.elements)):
            old, new = state.elements[ordinal], memo.elements[ordinal + offset] \
                if 0 <= ordinal + offset < len(memo.elements) else None
            if old is not None and new is not None:
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

    def _merge(self, before, again, k: int, j: int, shift: int, delta: int, items) -> tuple:
        """The pagination of the whole document: ``before``'s pages up to iteration ``k``,
        ``again``'s, and ``before``'s from iteration ``j`` on -- their flow positions moved
        by ``shift`` items, their pages and page numbers by ``delta`` -- and the flow's
        approximations over them, in order."""
        from .paginate import Pagination, Position

        def moved(position):
            return Position(position.item + shift, position.line, position.part)

        old_messages = self.state.approximated
        counts = [state[3] for state in before.states] + [len(old_messages)]
        new_messages = list(getattr(items, "approximated", ()))
        new_counts = [state[3] for state in again.states] + [len(new_messages)]
        flow_messages = new_messages[:new_counts[0]] if again.states else new_messages
        merged = Pagination()
        messages = list(flow_messages)
        states = []

        def add(state, chunk, by: int = 0):
            states.append((state[0] + by, state[1] + by, state[2], len(messages), state[4]))
            messages.extend(message for message in chunk if message not in messages)

        for index in range(k):
            add(before.states[index], old_messages[counts[index]:counts[index + 1]])
        for index in range(len(again.states)):
            add(again.states[index], new_messages[new_counts[index]:new_counts[index + 1]])
        moves = self.moves
        for index in range(j, len(before.starts)):
            add(before.states[index], renumbered(old_messages[counts[index]:counts[index + 1]], moves), delta)
        merged.states = states
        merged.starts = before.starts[:k] + again.starts + [moved(p) for p in before.starts[j:]]
        merged.ends = before.ends[:k] + again.ends + [moved(p) for p in before.ends[j:]]
        first = before.states[k][0]
        last = before.states[j][0] if j < len(before.starts) else None

        def pick(old: dict, new: dict, info: bool = False) -> dict:
            out = {page: value for page, value in old.items() if page < first}
            out.update(new)
            if last is not None:
                for page, value in old.items():
                    if page >= last:
                        out[page + delta] = dataclasses.replace(value, number=value.number + delta) \
                            if info and delta else value
            return out

        merged.info = dict(sorted(pick(before.info, again.info, True).items()))
        merged.bands = pick(before.bands, again.bands)
        merged.wraps = pick(before.wraps, again.wraps)
        merged.notes = pick(before.notes, again.notes)
        merged.frames = pick(before.frames, again.frames)
        start = before.starts[k]
        end = before.starts[j] if j < len(before.starts) else None
        pages = {index: found for index, found in before.pages.items() if index < start.item}
        pages.update(again.pages)
        if start.line:
            # The item the pages laid out again start inside: its lines before, as they were.
            pages[start.item] = before.pages.get(start.item, [])[:start.line] + again.pages.get(start.item, [])
        if end is not None:
            pages.update({index + shift: [page + delta for page in found] if delta else found
                          for index, found in before.pages.items() if index > end.item
                          or (index == end.item and not end.line)})
            if end.line:
                # The item the kept pages start inside: its lines before, as laid out again.
                kept = [page + delta for page in before.pages.get(end.item, [])[end.line:]]
                pages[end.item + shift] = again.pages.get(end.item + shift, [])[:end.line] + kept
        merged.pages = pages
        return merged, messages


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
    out = [None] * len(before.starts)
    plain = True
    for j in range(len(before.starts) - 1, -1, -1):
        page = before.states[j][0]
        end = before.states[j + 1][0] if j + 1 < len(before.starts) else (max(before.info) + 1 if before.info else page)
        for number in range(page, end):
            info = before.info.get(number)
            if info is None or info.blank or info.section_first:
                plain = False
        out[j] = (lambda delta, ok=plain: ok and (delta % 2 == 0 or not document.even_and_odd_headers))
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
