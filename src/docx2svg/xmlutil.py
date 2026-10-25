"""Namespace-blind XML helpers: :mod:`ooxml_common.xmlutil`'s, and the two Word adds.

OOXML namespaces are declared per part and the *prefixes* are not stable -- Word writes
``w:``, LibreOffice writes ``w:`` too, and a generator may write anything.  What is
stable is the local name.  Every lookup in this project is therefore by local name, the
same approach the sibling ``pptx2svg`` takes, and for the same reason: it removes a whole
class of bug in which a valid document parses to nothing because a prefix differed.

The cost is that two elements with the same local name in different namespaces become
indistinguishable.  In WordprocessingML that is not a live risk for the parts this
project reads; if it becomes one, :func:`qualified_name` is here to disambiguate at the
one site that needs it rather than everywhere.

The generic helpers are :mod:`ooxml_common.xmlutil`'s, which the rest of docx2svg imports
directly; ``children`` and ``child`` are its ``element_children`` and ``element_child``,
which step over comments and processing instructions, since docx-agent hands this project
lxml trees that keep them.  What stays here is WordprocessingML's: :func:`attr_int` (an
integer attribute, a measurement such as ``12pt`` falling back) and :func:`attr_bool`
(``ST_OnOff``, where an absent attribute means true).  The generic names are re-exported
so that ``docx2svg.xmlutil`` keeps working; :func:`descendants` is now ooxml-common's,
which leaves out the element it starts from.
"""

from __future__ import annotations

from xml.etree.ElementTree import Element

from ooxml_common.xmlutil import attr, descendants, local_name, parse_xml
from ooxml_common.xmlutil import element_child as child
from ooxml_common.xmlutil import element_children as children

__all__ = [
    "attr", "attr_bool", "attr_int", "child", "children", "descendants", "local_name", "parse_xml",
    "qualified_name",
]


def qualified_name(tag: str) -> tuple[str, str]:
    """``(namespace, local)``; namespace is ``""`` for an unqualified tag."""
    if tag.startswith("{"):
        namespace, _, rest = tag[1:].partition("}")
        return namespace, local_name(rest)
    return "", local_name(tag)


def attr_int(element: Element | None, name: str, default: int | None = None) -> int | None:
    value = attr(element, name)
    if value is None:
        return default
    try:
        return int(value)
    except ValueError:
        # An attribute the schema types as an integer but a generator wrote as a
        # measurement ("12pt").  Better to fall back than to fail the whole document.
        return default


def attr_bool(element: Element | None, name: str = "val", default: bool = False) -> bool:
    """An OOXML on/off attribute.

    The schema's ``ST_OnOff`` accepts ``1``/``0``, ``true``/``false`` and
    ``on``/``off``, and -- the part that catches people -- **an absent attribute means
    true**.  ``<w:b/>`` is bold; only ``<w:b w:val="0"/>`` is not.  This is the opposite
    of the usual XML convention and is why toggles go through this function.
    """
    if element is None:
        return default
    value = attr(element, name)
    if value is None:
        return True
    return value.lower() in ("1", "true", "on")
