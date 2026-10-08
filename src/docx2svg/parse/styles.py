"""The parts the cascade reads besides the body: styles, theme fonts, numbering, settings.

Each is read as *declared* -- a style's own properties, not its inherited ones; a
numbering level's own ``w:pPr``, not merged with the paragraph's.  Merging is
:mod:`docx2svg.resolve`'s job and nothing here anticipates it, so a wrong answer can be
blamed on one stage.
"""

from __future__ import annotations

import dataclasses
from dataclasses import dataclass, field

from ..model import FontScheme, NumberingLevel, Style, StyleSheet, ThemeFonts
from ..xmlutil import attr, attr_bool, attr_int, child, children, local_name, parse_xml
from .properties import _on, read_paragraph_properties, read_run_properties, read_table_properties


def parse_styles(styles_xml: bytes) -> StyleSheet:
    root = parse_xml(styles_xml)
    sheet = StyleSheet()
    defaults = child(root, "docDefaults")
    sheet.run_defaults = read_run_properties(child(child(defaults, "rPrDefault"), "rPr"))
    sheet.paragraph_defaults = read_paragraph_properties(child(child(defaults, "pPrDefault"), "pPr"))
    for element in children(root, "style"):
        style_id = attr(element, "styleId")
        if style_id is None:
            continue
        paragraph = read_paragraph_properties(child(element, "pPr"))
        run = read_run_properties(child(element, "rPr"))
        default = attr(element, "default")
        style = Style(
            style_id=style_id,
            kind=attr(element, "type") or "paragraph",
            name=attr(child(element, "name"), "val"),
            based_on=attr(child(element, "basedOn"), "val"),
            next_style=attr(child(element, "next"), "val"),
            link=attr(child(element, "link"), "val"),
            default=default is not None and default.lower() in ("1", "true", "on"),
            paragraph=paragraph,
            run=run,
            conditional={
                kind: (read_paragraph_properties(child(node, "pPr")),
                       read_run_properties(child(node, "rPr")),
                       read_table_properties(child(node, "tblPr")),
                       read_table_properties(child(node, "trPr")),
                       read_table_properties(child(node, "tcPr")))
                for node in children(element, "tblStylePr")
                if (kind := attr(node, "type"))
            },
            table=read_table_properties(child(element, "tblPr")),
            table_row=read_table_properties(child(element, "trPr")),
            table_cell=read_table_properties(child(element, "tcPr")),
        )
        # A duplicate id is legal-looking and Word keeps the first; so does this.
        sheet.styles.setdefault(style_id, style)
    return sheet


def _theme_fonts(element) -> ThemeFonts:
    return ThemeFonts(
        latin=attr(child(element, "latin"), "typeface") or "",
        east_asian=attr(child(element, "ea"), "typeface") or "",
        complex_script=attr(child(element, "cs"), "typeface") or "",
        scripts={
            attr(font, "script"): attr(font, "typeface") or ""
            for font in children(element, "font")
            if attr(font, "script")
        },
    )


def parse_theme(theme_xml: bytes) -> FontScheme | None:
    root = parse_xml(theme_xml)
    scheme = child(child(root, "themeElements"), "fontScheme")
    if scheme is None:
        return None
    return FontScheme(
        major=_theme_fonts(child(scheme, "majorFont")),
        minor=_theme_fonts(child(scheme, "minorFont")),
    )


def parse_theme_colors(theme_xml: bytes) -> dict[str, str]:
    """The theme's colour scheme (``a:clrScheme``): each slot (``dk1``, ``lt1``, ``dk2``,
    ``lt2``, ``accent1``-``accent6``, ``hlink``, ``folHlink``) as ``RRGGBB`` -- an
    ``a:srgbClr``'s value, or an ``a:sysClr``'s ``lastClr``."""
    scheme = child(child(parse_xml(theme_xml), "themeElements"), "clrScheme")
    out: dict[str, str] = {}
    for slot in children(scheme):
        for colour in children(slot):
            value = attr(colour, "val") if local_name(colour.tag) == "srgbClr" else attr(colour, "lastClr")
            if value:
                out[local_name(slot.tag)] = value.upper()
    return out


def parse_theme_formats(theme_xml: bytes):
    """The theme's format scheme (``a:fmtScheme``): the fills, outlines and effects a
    shape's ``wps:style`` references by index, read by the shared reader
    (``ooxml_common.drawingml.read.parse_format_scheme``)."""
    from ooxml_common.drawingml.read import parse_format_scheme

    return parse_format_scheme(child(child(parse_xml(theme_xml), "themeElements"), "fmtScheme"))


def parse_color_map(settings_xml: bytes) -> dict[str, str]:
    """``w:clrSchemeMapping``: ``bg1`` -> ``light1`` and the rest, as stated."""
    element = child(parse_xml(settings_xml), "clrSchemeMapping")
    if element is None:
        return {}
    return {local_name(key): value for key, value in element.attrib.items()}


def _level(element) -> NumberingLevel:
    legal = child(element, "isLgl")
    return NumberingLevel(
        level=attr_int(element, "ilvl", 0) or 0,
        format=attr(child(element, "numFmt"), "val"),
        text=attr(child(element, "lvlText"), "val"),
        paragraph=read_paragraph_properties(child(element, "pPr")),
        run=read_run_properties(child(element, "rPr")),
        suffix=attr(child(element, "suff"), "val"),
        start=attr_int(child(element, "start"), "val"),
        picture_bullet=attr_int(child(element, "lvlPicBulletId"), "val"),
        alignment=attr(child(element, "lvlJc"), "val"),
        restart=attr_int(child(element, "lvlRestart"), "val"),
        legal=legal is not None and _on(legal),
    )


_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def parse_picture_bullets(numbering_xml: bytes) -> dict[int, str | None]:
    """``w:numPicBullet``: id -> the relationship id of its picture (in the numbering
    part), from a VML ``w:pict`` (``v:imagedata/@r:id``) or a ``w:drawing``
    (``a:blip/@r:embed``); ``None`` where it names none."""
    out: dict[int, str | None] = {}
    for element in children(parse_xml(numbering_xml), "numPicBullet"):
        identifier = attr_int(element, "numPicBulletId")
        if identifier is None:
            continue
        found = None
        for node in element.iter():
            name = local_name(node.tag)
            if name == "imagedata" and node.get(_R + "id"):
                found = node.get(_R + "id")
            elif name == "blip" and node.get(_R + "embed"):
                found = node.get(_R + "embed")
            if found:
                break
        out[identifier] = found
    return out


@dataclass
class Numbering:
    """``w:numbering`` as the layout counts it (:func:`parse_numbering`)."""

    #: numId -> ilvl -> level, overrides applied.
    levels: dict = field(default_factory=dict)
    #: numId -> the abstract definition whose counts it shares.
    lists: dict = field(default_factory=dict)
    #: ``(numId, ilvl)`` restarted at the instance's first item there.
    restarts: frozenset = frozenset()


def parse_numbering(numbering_xml: bytes, styles: StyleSheet | None = None) -> Numbering:
    """numId -> ilvl -> level, the list each instance counts in, and where an override
    restarts it (``make_numbering_probe.py``; ROADMAP.md, "Numbering and lists").

    * Abstract definitions sharing a ``w:nsid`` are one: the first in the part, its levels
      and its counts (a later one's levels are not used).
    * ``w:numStyleLink`` takes the levels of the definition under the numbering style's
      ``w:numId``; the instance still counts in its own definition's list.
    * ``w:lvlOverride/w:lvl`` replaces the level; ``w:startOverride`` is the level's start
      in this instance, and restarts the list's count at its first item at that level --
      at the overriding ``w:lvl``'s ``w:start`` where there is one.  A ``w:lvl`` alone
      restarts nothing.
    """
    root = parse_xml(numbering_xml)
    abstract: dict[int, dict[int, NumberingLevel]] = {}
    canonical: dict[int, int] = {}
    links: dict[int, str] = {}
    first_by_nsid: dict[str, int] = {}
    for element in children(root, "abstractNum"):
        identifier = attr_int(element, "abstractNumId")
        if identifier is None:
            continue
        nsid = attr(child(element, "nsid"), "val")
        if nsid and nsid.upper() in first_by_nsid:
            canonical[identifier] = first_by_nsid[nsid.upper()]
            continue
        if nsid:
            first_by_nsid[nsid.upper()] = identifier
        canonical[identifier] = identifier
        abstract[identifier] = {level.level: level for level in map(_level, children(element, "lvl"))}
        link = attr(child(element, "numStyleLink"), "val")
        if link:
            links[identifier] = link
    instances = {attr_int(element, "numId"): element for element in children(root, "num")}

    def source_levels(identifier: int, seen: frozenset = frozenset()) -> dict[int, NumberingLevel]:
        identifier = canonical.get(identifier, identifier)
        link = links.get(identifier)
        if link is not None and styles is not None and identifier not in seen:
            style = styles.styles.get(link)
            linked = style.paragraph.get("numPr.numId") if style is not None else None
            element = instances.get(linked)
            target = attr_int(child(element, "abstractNumId"), "val") if element is not None else None
            if target is not None:
                return source_levels(target, seen | {identifier})
        return abstract.get(identifier, {})

    out = Numbering()
    restarts = set()
    for identifier, element in instances.items():
        source = attr_int(child(element, "abstractNumId"), "val")
        if identifier is None or source is None:
            continue
        levels = dict(source_levels(source))
        for override in children(element, "lvlOverride"):
            ilvl = attr_int(override, "ilvl", 0) or 0
            replacement = child(override, "lvl")
            start = attr_int(child(override, "startOverride"), "val")
            if replacement is not None:
                level = _level(replacement)
                levels[level.level] = level
                ilvl = level.level
            elif start is not None:
                base = levels.get(ilvl) or NumberingLevel(level=ilvl)
                levels[ilvl] = dataclasses.replace(base, start=start)
            if start is not None:
                restarts.add((identifier, ilvl))
        out.levels[identifier] = levels
        out.lists[identifier] = canonical.get(source, source)
    out.restarts = frozenset(restarts)
    return out


def parse_settings(settings_xml: bytes) -> dict[str, str]:
    """``w:themeFontLang``'s attributes -- the only setting the cascade reads."""
    element = child(parse_xml(settings_xml), "themeFontLang")
    if element is None:
        return {}
    return {key: value for key in ("val", "eastAsia", "bidi") if (value := attr(element, key))}


def parse_endnote_settings(settings_xml: bytes, kind: str = "endnote") -> tuple[str | None, frozenset]:
    """``w:endnotePr`` (``kind`` ``footnote``: ``w:footnotePr``): ``w:pos`` (``None``
    when unstated) and the ids of the notes it names (``w:endnote/@w:id``: the
    separators)."""
    element = child(parse_xml(settings_xml), f"{kind}Pr")
    if element is None:
        return None, frozenset()
    named = frozenset(value for value in (attr_int(node, "id") for node in children(element, kind))
                      if value is not None)
    return attr(child(element, "pos"), "val"), named


def parse_default_tab_stop(settings_xml: bytes) -> int | None:
    """``w:defaultTabStop/@w:val`` in twips, or ``None`` when the settings do not state it."""
    return attr_int(child(parse_xml(settings_xml), "defaultTabStop"), "val")


def parse_even_and_odd_headers(settings_xml: bytes) -> bool:
    """``w:evenAndOddHeaders``: even-numbered pages take the ``even`` header and footer
    (``make_story_select_probe.py``)."""
    element = child(parse_xml(settings_xml), "evenAndOddHeaders")
    return element is not None and attr_bool(element)


def parse_hyphenation(settings_xml: bytes):
    """``w:autoHyphenation`` with ``w:hyphenationZone``, ``w:consecutiveHyphenLimit`` and
    ``w:doNotHyphenateCaps``, as a :class:`~docx2svg.model.Hyphenation`; ``None`` when
    automatic hyphenation is off (the element absent, or ``w:val`` false)."""
    from ..model import Hyphenation

    root = parse_xml(settings_xml)
    auto = child(root, "autoHyphenation")
    if auto is None or not attr_bool(auto, "val", default=True):
        return None
    caps = child(root, "doNotHyphenateCaps")
    return Hyphenation(zone=attr_int(child(root, "hyphenationZone"), "val"),
                       limit=attr_int(child(root, "consecutiveHyphenLimit"), "val") or 0,
                       no_caps=caps is not None and attr_bool(caps, "val", default=True))


def parse_compatibility_mode(settings_xml: bytes) -> int | None:
    """``w:compat/w:compatSetting[@w:name="compatibilityMode"]/@w:val``, or ``None`` when
    the settings do not state one (no ``w:compat``, an empty one as Word 12 writes, or
    one holding only legacy options).  What an unstated mode means is a layout question,
    answered where it is measured (:func:`docx2svg.vertical.keeps_space_before_at_page_top`).
    """
    for setting in children(child(parse_xml(settings_xml), "compat"), "compatSetting"):
        if attr(setting, "name") == "compatibilityMode":
            return attr_int(setting, "val")
    return None


def parse_compat_options(settings_xml: bytes) -> frozenset[str]:
    """The legacy ``w:compat`` options that are on (``w:doNotUseHTMLParagraphAutoSpacing``,
    ``w:suppressSpBfAfterPgBrk``...), by local name.  ``w:compatSetting`` is not one of
    them (:func:`parse_compatibility_mode` reads the mode)."""
    return frozenset(
        local_name(option.tag) for option in children(child(parse_xml(settings_xml), "compat"))
        if local_name(option.tag) != "compatSetting" and attr_bool(option, "val", default=True)
    )


def parse_compat_settings(settings_xml: bytes) -> dict[str, str]:
    """Every ``w:compat/w:compatSetting`` but the mode, name -> ``w:val`` as written."""
    return {name: attr(setting, "val") or "" for setting in children(child(parse_xml(settings_xml), "compat"),
                                                                       "compatSetting")
            if (name := attr(setting, "name")) and name != "compatibilityMode"}
