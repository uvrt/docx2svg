"""Open Packaging Conventions reader.

A ``.docx`` is the same container a ``.pptx`` is: a ZIP of XML parts wired together by
relationship files, with ``[Content_Types].xml`` saying what each part is.  Only the
*contents* differ -- the root part is ``word/document.xml`` rather than
``ppt/presentation.xml``, and the relationship types below are the Word ones.

That is the first and most concrete piece of evidence for the shared-package question
this project has to settle: this module and ``pptx2svg/opc.py`` do the same job on the
same bytes, and the divergence between them is a list of URI constants.  ROADMAP.md,
"Phase 1", records what the whole of that overlap measured out to be.
"""

from __future__ import annotations

import posixpath
import zipfile
from dataclasses import dataclass, field
from typing import BinaryIO

from .xmlutil import attr, children, parse_xml

CONTENT_TYPES_PART = "[Content_Types].xml"

_OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_OFFICE_DOCUMENT = f"{_OFFICE}/officeDocument"
REL_STYLES = f"{_OFFICE}/styles"
REL_NUMBERING = f"{_OFFICE}/numbering"
REL_SETTINGS = f"{_OFFICE}/settings"
REL_FONT_TABLE = f"{_OFFICE}/fontTable"
REL_THEME = f"{_OFFICE}/theme"
REL_HEADER = f"{_OFFICE}/header"
REL_FOOTER = f"{_OFFICE}/footer"
REL_FOOTNOTES = f"{_OFFICE}/footnotes"
REL_ENDNOTES = f"{_OFFICE}/endnotes"
REL_IMAGE = f"{_OFFICE}/image"
REL_HYPERLINK = f"{_OFFICE}/hyperlink"


@dataclass(frozen=True)
class Relationship:
    id: str
    type: str
    target: str
    #: ``"External"`` for a hyperlink or a linked image; such a target is a URI and
    #: must never be resolved against a part path.
    target_mode: str = "Internal"

    @property
    def external(self) -> bool:
        return self.target_mode == "External"


@dataclass
class Package:
    """An opened OPC container.

    Parts are read lazily and cached: a document with fifty images should not cost
    fifty megabytes of memory to ask what its page size is.
    """

    archive: zipfile.ZipFile
    content_types: dict[str, str] = field(default_factory=dict)
    default_types: dict[str, str] = field(default_factory=dict)
    _cache: dict[str, bytes] = field(default_factory=dict, repr=False)
    _rels: dict[str, dict[str, Relationship]] = field(default_factory=dict, repr=False)

    # -- construction -------------------------------------------------------

    @classmethod
    def open(cls, source: str | bytes | BinaryIO) -> "Package":
        import io

        if isinstance(source, bytes):
            source = io.BytesIO(source)
        package = cls(archive=zipfile.ZipFile(source))
        package._read_content_types()
        return package

    def _read_content_types(self) -> None:
        try:
            root = parse_xml(self.archive.read(CONTENT_TYPES_PART))
        except KeyError:
            # Not fatal for reading: nothing in this project dispatches on content
            # type yet, and a package missing it is one Word would repair rather than
            # refuse.  Recording the absence beats raising here.
            return
        for element in children(root):
            extension = attr(element, "Extension")
            part_name = attr(element, "PartName")
            content_type = attr(element, "ContentType") or ""
            if extension is not None:
                self.default_types[extension.lower()] = content_type
            elif part_name is not None:
                self.content_types[part_name.lstrip("/")] = content_type

    # -- parts --------------------------------------------------------------

    def read(self, part: str) -> bytes:
        part = part.lstrip("/")
        if part not in self._cache:
            self._cache[part] = self.archive.read(part)
        return self._cache[part]

    def exists(self, part: str) -> bool:
        return part.lstrip("/") in set(self.archive.namelist())

    def content_type(self, part: str) -> str:
        part = part.lstrip("/")
        if part in self.content_types:
            return self.content_types[part]
        extension = posixpath.splitext(part)[1].lstrip(".").lower()
        return self.default_types.get(extension, "")

    # -- relationships ------------------------------------------------------

    @staticmethod
    def rels_part_for(part: str) -> str:
        """``word/document.xml`` -> ``word/_rels/document.xml.rels``.

        The package-level relationships are the special case: the "part" is the empty
        string and its rels live at ``_rels/.rels``.
        """
        part = part.lstrip("/")
        directory, name = posixpath.split(part)
        return posixpath.join(directory, "_rels", f"{name}.rels")

    def relationships(self, part: str = "") -> dict[str, Relationship]:
        part = part.lstrip("/")
        if part in self._rels:
            return self._rels[part]
        rels_part = self.rels_part_for(part)
        found: dict[str, Relationship] = {}
        try:
            root = parse_xml(self.archive.read(rels_part))
        except KeyError:
            self._rels[part] = found
            return found
        for element in children(root, "Relationship"):
            identifier = attr(element, "Id") or ""
            found[identifier] = Relationship(
                id=identifier,
                type=attr(element, "Type") or "",
                target=attr(element, "Target") or "",
                target_mode=attr(element, "TargetMode") or "Internal",
            )
        self._rels[part] = found
        return found

    def resolve(self, part: str, relationship: Relationship) -> str:
        """A relationship target as a package part path.

        Targets are relative to the *directory* of the part that declares them, and may
        walk upward (``../media/image1.png`` from ``word/document.xml``).  An external
        target is returned unchanged; callers must check :attr:`Relationship.external`
        before treating the result as a part.
        """
        if relationship.external:
            return relationship.target
        if relationship.target.startswith("/"):
            return relationship.target.lstrip("/")
        directory = posixpath.dirname(part.lstrip("/"))
        return posixpath.normpath(posixpath.join(directory, relationship.target))

    def related(self, part: str, relationship_type: str) -> list[str]:
        return [
            self.resolve(part, relationship)
            for relationship in self.relationships(part).values()
            if relationship.type == relationship_type and not relationship.external
        ]

    def part_by_id(self, part: str, relationship_id: str) -> str | None:
        relationship = self.relationships(part).get(relationship_id)
        if relationship is None or relationship.external:
            return None
        return self.resolve(part, relationship)

    # -- the document -------------------------------------------------------

    @property
    def main_document_part(self) -> str:
        """The part ``_rels/.rels`` names as the office document.

        Not hard-coded to ``word/document.xml``.  The convention is near-universal but
        it *is* a convention -- the relationship is the normative statement -- and a
        reader that hard-codes it silently reads the wrong part in a package that
        renames it, which is a class of bug worth not having.
        """
        for relationship in self.relationships().values():
            if relationship.type == REL_OFFICE_DOCUMENT:
                return self.resolve("", relationship)
        return "word/document.xml"
