"""OOXML -> source model.  Deliberately unresolved; see :mod:`docx2svg.model`."""

from .document import parse_document, parse_package

__all__ = ["parse_document", "parse_package"]
