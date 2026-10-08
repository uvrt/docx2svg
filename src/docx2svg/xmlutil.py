"""Namespace-blind XML helpers.

OOXML namespaces are declared per part and the *prefixes* are not stable -- Word writes
``w:``, LibreOffice writes ``w:`` too, and a generator may write anything.  What is
stable is the local name.  Every lookup in this project is therefore by local name, the
same approach the sibling ``pptx2svg`` takes, and for the same reason: it removes a whole
class of bug in which a valid document parses to nothing because a prefix differed.

The cost is that two elements with the same local name in different namespaces become
indistinguishable.  In WordprocessingML that is not a live risk for the parts this
project reads; if it becomes one, :func:`qualified_name` is here to disambiguate at the
one site that needs it rather than everywhere.
"""

from __future__ import annotations

import functools

from typing import Iterator
from xml.etree.ElementTree import Element, fromstring


def parse_xml(data: bytes | str) -> Element:
    if isinstance(data, str):
        data = data.encode("utf-8")
    return fromstring(data)


@functools.lru_cache(maxsize=4096)
def local_name(tag: str) -> str:
    """``{ns}w:body`` or ``{ns}body`` -> ``body``.  Cached: a document has a few hundred
    distinct tags and attribute names, and parsing asks for them millions of times."""
    if tag.startswith("{"):
        tag = tag.split("}", 1)[1]
    return tag.rsplit(":", 1)[-1]


def qualified_name(tag: str) -> tuple[str, str]:
    """``(namespace, local)``; namespace is ``""`` for an unqualified tag."""
    if tag.startswith("{"):
        namespace, _, rest = tag[1:].partition("}")
        return namespace, local_name(rest)
    return "", local_name(tag)


def children(element: Element | None, name: str | None = None) -> Iterator[Element]:
    if element is None:
        return
    for child in element:
        if not isinstance(child.tag, str):  # comments, processing instructions
            continue
        if name is None or local_name(child.tag) == name:
            yield child


def child(element: Element | None, name: str) -> Element | None:
    for found in children(element, name):
        return found
    return None


def descendants(element: Element | None, name: str) -> Iterator[Element]:
    if element is None:
        return
    for node in element.iter():
        if isinstance(node.tag, str) and local_name(node.tag) == name:
            yield node


def attr(element: Element | None, name: str, default: str | None = None) -> str | None:
    """Attribute by local name.

    Attributes carry the namespace of their *declaration*, not of their element, so
    ``w:val`` on a ``w:sz`` is ``{...wordprocessingml...}val`` while ``val`` on an
    unqualified element is bare.  Both spellings occur; matching on the local name
    covers both.
    """
    if element is None:
        return default
    value = element.get(name)
    if value is not None:
        return value
    for key, candidate in element.attrib.items():
        if local_name(key) == name:
            return candidate
    return default


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
