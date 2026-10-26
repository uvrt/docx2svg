"""Open Packaging Conventions: :class:`ooxml_common.opc.OpcPackage`, and what Word adds.

A ``.docx`` is the same container a ``.pptx`` is: a ZIP of XML parts wired together by
relationship files, with ``[Content_Types].xml`` saying what each part is.  Only the
*contents* differ -- the root part is ``word/document.xml`` rather than
``ppt/presentation.xml``, and the relationship types below are the Word ones.  So the
reader is ooxml-common's, shared with pptx2svg, and what stays here is that difference:
the Word relationship types, and three readings this project has always made of a
package where :class:`~ooxml_common.opc.OpcPackage` answers otherwise --

* :func:`main_part`: the part ``_rels/.rels`` names as the office document, else
  ``word/document.xml`` (``OpcPackage.presentation_part`` falls back to PowerPoint's);
* :func:`read_part`: a part's bytes, raising :class:`KeyError` for a part the package
  does not hold, as reading the ZIP did (``OpcPackage.read`` returns ``None``);
* :func:`content_type`: what ``[Content_Types].xml`` says, ``""`` where it says nothing
  (``OpcPackage.content_type`` falls back to a type the extension suggests).

:class:`Package` is the reader this module used to be, as names over ``OpcPackage``, kept
for one release for whoever imports it (``docx2svg.Package``).
"""

from __future__ import annotations

import posixpath

from ooxml_common.opc import (
    CONTENT_TYPES_PART,
    REL_HYPERLINK,
    REL_IMAGE,
    REL_OFFICE_DOCUMENT,
    REL_THEME,
    OpcPackage,
    Relationship,
    normalize_part_path,
    rels_path_for,
)

__all__ = [
    "CONTENT_TYPES_PART", "REL_ENDNOTES", "REL_FONT_TABLE", "REL_FOOTER", "REL_FOOTNOTES", "REL_HEADER",
    "REL_HYPERLINK", "REL_IMAGE", "REL_NUMBERING", "REL_OFFICE_DOCUMENT", "REL_SETTINGS", "REL_STYLES",
    "REL_THEME", "OpcPackage", "Package", "Relationship", "content_type", "main_part", "read_part",
]

_OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_STYLES = f"{_OFFICE}/styles"
REL_NUMBERING = f"{_OFFICE}/numbering"
REL_SETTINGS = f"{_OFFICE}/settings"
REL_FONT_TABLE = f"{_OFFICE}/fontTable"
REL_HEADER = f"{_OFFICE}/header"
REL_FOOTER = f"{_OFFICE}/footer"
REL_FOOTNOTES = f"{_OFFICE}/footnotes"
REL_ENDNOTES = f"{_OFFICE}/endnotes"


def main_part(package: OpcPackage) -> str:
    """The part ``_rels/.rels`` names as the office document.

    Not hard-coded to ``word/document.xml``.  The convention is near-universal but
    it *is* a convention -- the relationship is the normative statement -- and a
    reader that hard-codes it silently reads the wrong part in a package that
    renames it, which is a class of bug worth not having.
    """
    for relationship in package.relationships("").values():
        if relationship.type == REL_OFFICE_DOCUMENT and relationship.target_part is not None:
            return relationship.target_part
    return "word/document.xml"


def read_part(package: OpcPackage, part: str) -> bytes:
    """``part``'s bytes; :class:`KeyError` where the package has no such part."""
    data = package.read(part)
    if data is None:
        raise KeyError(f"There is no item named {part.lstrip('/')!r} in the archive")
    return data


def content_type(package: OpcPackage, part: str) -> str:
    """``part``'s type as ``[Content_Types].xml`` states it (an override, else its
    extension's default), ``""`` where it states none: a picture whose type the package
    does not say is not drawn as one."""
    part = normalize_part_path(part)
    if part in package.content_types:
        return package.content_types[part]
    extension = posixpath.splitext(part)[1].lstrip(".").lower()
    return package.default_types.get(extension, "")


class Package(OpcPackage):
    """The reader this module used to be, as names over :class:`OpcPackage`: kept for one
    release for whoever imports ``docx2svg.Package`` or ``docx2svg.opc.Package``.  Its
    relationships are ooxml-common's :class:`~ooxml_common.opc.Relationship`
    (``target_part``, ``is_external``), and it no longer holds the ZIP open (``archive``)."""

    @property
    def main_document_part(self) -> str:
        return main_part(self)

    def read(self, part: str) -> bytes:  # type: ignore[override]
        return read_part(super(), part)  # type: ignore[arg-type]

    def exists(self, part: str) -> bool:
        return self.has_part(part)

    def content_type(self, part: str) -> str:  # type: ignore[override]
        return content_type(self, part)

    @staticmethod
    def rels_part_for(part: str) -> str:
        return rels_path_for(normalize_part_path(part))

    def resolve(self, part: str, relationship: Relationship) -> str:
        return relationship.target if relationship.is_external else relationship.target_part or ""

    def related(self, part: str, relationship_type: str) -> list[str]:
        return self.related_parts_of_type(part, relationship_type)

    def part_by_id(self, part: str, relationship_id: str) -> str | None:
        return self.related_part(part, relationship_id)
