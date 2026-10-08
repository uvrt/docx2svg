"""Write a laid-out page as SVG: a transcription of :mod:`docx2svg.layout`, nothing more.

Every number in the document comes from the layout; this module only formats it.

* **Coordinates are Word's device pixels** (1/300 inch, the grid its PDF export draws
  on: ROADMAP.md, Phase 0.5): the ``viewBox`` is the page's extent in whole device
  pixels and ``width``/``height`` say it in points, so baselines are integers and a pen
  position is exact to four decimals (the layout unit is 0.00102 px).
* **Every glyph has its own x.**  A ``<text>`` element lists one x per character, so
  the rasteriser's own kerning, shaping and ligatures cannot move a glyph from where
  the model put it (SVG 1.1 10.4: each absolutely positioned character starts a new
  text chunk).  Word applies kerning only where ``w:kern`` asks, and the layout has
  already added it.
* **Identity**: every element drawn from the document carries ``data-docx-path``, the
  structural path of the paragraph, run or table it came from (``w:body/w:p[3]/w:r[2]``),
  and a paragraph's group also ``data-docx-id``, its ``w14:paraId`` when the file has one
  -- which is not unique, because Word copies it with the paragraph; hence both.  A
  header's or footer's paragraph group also says which (``data-docx-story``) and from
  which part (``data-docx-part="word/header1.xml"``), its paths being the part's own
  (``w:hdr/w:p[1]``).  A **note separator** is a group of its own -- its path
  ``w:footnotes/separator`` (``continuationSeparator``, ``w:endnotes/...``) where Word's
  own is drawn, the note's paragraph (``w:footnote[@w:id=-1]/w:p[1]``) where the
  settings name the part's -- holding one empty line group; its line is a
  ``<rect data-docx-kind="separator">`` with the path of its run.  Word's own resolves to
  no element and has no ``data-docx-id``: it is not a paragraph of the text.
* **Fonts are named, not embedded**: ``font-family`` is the face the cascade resolved,
  ``font-weight`` and ``font-style`` its bold and italic -- except for a face that name
  does not reach in a rasteriser (a superfamily member such as Calibri Light, filed
  under "Calibri" at weight 300): then a fallback list, ``font-family="Calibri Light,
  Calibri"``, and the face's own weight, stretch and style
  (:class:`docx2svg.fonts.DrawingName`), which find the original file.

Output is a function of the layout alone: no set iteration, no clock, no platform
formatting (numbers go through :func:`number`).
"""

from __future__ import annotations

import base64
import re
from fractions import Fraction
from xml.sax.saxutils import escape, quoteattr

from .layout import Layout, Page, Span
from .vertical import PX_PER_PT

#: What the glyph outlines are scaled to.  ``device``: the size rounded to whole device
#: pixels, which is what Word draws its ink at (an 11 pt glyph at 46 px, while the pen
#: advances by 45.83 px -- Phase 2); ``exact``: the size itself.  Pen positions are the
#: model's either way.  ROADMAP.md, "Phase 5 -- measured", has the measurement that
#: chose the default.
GLYPH_SIZES = ("device", "exact")


def number(value) -> str:
    """A length in device px, to four decimals, half up, no trailing zeros."""
    scaled = Fraction(value) * 10000
    whole = (scaled.numerator * 2 + scaled.denominator) // (2 * scaled.denominator)
    sign = "-" if whole < 0 else ""
    whole = abs(whole)
    integer, fraction = divmod(whole, 10000)
    if not fraction:
        return f"{sign}{integer}"
    return f"{sign}{integer}.{fraction:04d}".rstrip("0")


def font_size_px(half_points: int, glyph_size: str) -> Fraction:
    exact = Fraction(half_points, 2) * PX_PER_PT
    if glyph_size == "exact":
        return exact
    return Fraction((exact.numerator * 2 + exact.denominator) // (2 * exact.denominator))


def _color(value: str | None, default: str = "000000") -> str:
    return "#" + (value or default).upper()


def _family(face: str) -> str:
    return quoteattr(face)


_PLAIN_FAMILY = re.compile(r"[A-Za-z_][A-Za-z0-9_-]*( [A-Za-z_][A-Za-z0-9_-]*)*")


def _family_list(families: list[str]) -> str:
    """A CSS family list: a name of plain words as it is, any other quoted."""
    def one(name: str) -> str:
        if _PLAIN_FAMILY.fullmatch(name):
            return name
        return "'" + name.replace("\\", "\\\\").replace("'", "\\'") + "'"

    return quoteattr(", ".join(one(name) for name in families))


def _font(span: Span, names) -> list[str]:
    """The ``font-*`` attributes naming the span's face (:class:`docx2svg.fonts.DrawingName`)."""
    name = names(span.face, span.bold, span.italic) if names is not None else None
    if name is None:
        out = [f"font-family={_family(span.face)}"]
        if span.bold:
            out.append('font-weight="bold"')
        if span.italic:
            out.append('font-style="italic"')
        return out
    families = [span.face] + ([name.family] if name.family.lower() != span.face.lower() else [])
    out = [f"font-family={_family_list(families)}"]
    if name.weight != 400:
        out.append(f'font-weight="{name.weight}"')
    if name.stretch != "normal":
        out.append(f'font-stretch="{name.stretch}"')
    if name.style != "normal":
        out.append(f'font-style="{name.style}"')
    return out


def _paint(span: Span, size_px, defs) -> str:
    """A span's text effects (``w14:textFill`` / ``w14:textOutline``) as SVG paint, drawn
    by the shared renderers under Word's rules; the gradient spans the span's text."""
    from ooxml_common.drawingml import fill as dml_fill
    from ooxml_common.drawingml import model as dml
    from ooxml_common.drawingml.rules import WORD
    from ooxml_common.drawingml.svg import Defs

    defs = defs if defs is not None else Defs()
    fill, outline = span.paint
    box = (float(span.xs[0]), float(span.y - size_px * Fraction(4, 5)), float(span.end - span.xs[0]), float(size_px))
    if fill is None:
        attrs = f'fill="{_color(span.color)}"'
    else:
        attrs = dml_fill.render_fill_attrs(fill, defs, box, rules=WORD, dpi=300)
    if outline is not None:
        attrs += " " + dml_fill.render_outline_attrs(outline, defs, rules=WORD, dpi=300, box=box)
    return attrs if not isinstance(fill, dml.NoFill) or outline is not None else 'fill="none"'


def _text(span: Span, glyph_size: str, names=None, defs=None) -> str:
    chars = "".join(span.chars)
    name = names(span.face, span.bold, span.italic) if names is not None else None
    if name is not None and getattr(name, "chars", None):
        # A symbol face laid out from recorded metrics: its symbols' Unicode equivalents,
        # each at the position Word's face gives the symbol.
        chars = "".join(name.chars.get(char, char) for char in span.chars)
    font = _font(span, names)
    attributes = [
        f"data-docx-path={quoteattr(span.path)}",
        f'x="{" ".join(number(x) for x in span.xs)}"',
        f'y="{number(span.y)}"',
        font[0],
        f'font-size="{number(font_size_px(span.half_points, glyph_size))}"',
    ] + font[1:]
    if getattr(span, "paint", None):
        attributes.append(_paint(span, font_size_px(span.half_points, glyph_size), defs))
    else:
        attributes.append(f'fill="{_color(span.color)}"')
    if span.kind != "text":
        attributes.append(f'data-docx-kind="{span.kind}"')
    if chars != chars.strip() or "  " in chars:
        attributes.append('xml:space="preserve"')
    return f"<text {' '.join(attributes)}>{escape(chars)}</text>"


def _rect(x, y, width, height, fill: str, *, extra: str = "") -> str:
    return (f'<rect x="{number(x)}" y="{number(y)}" width="{number(width)}" height="{number(height)}"'
            f' fill="{fill}"{extra}/>')


#: The rules painted before the text, in this order (a stable sort keeps each kind's own
#: order): what text sits on must be under it, in SVG's painter's order as in Word's.
#: Word's PDF (``sample-with-table``, ``sample-simple``) draws each table row's borders,
#: then each cell's shading followed by that cell's text; the model's borders and shading
#: share no pixel, so drawing every table border before any shading is the same picture.
#: Nested backgrounds go outermost first: a cell's shading, a paragraph's, a run's shading
#: or highlight.  Everything else -- paragraph borders, underlines, strikes -- Word draws
#: after the text it belongs to, and so does this.  ROADMAP.md, "Phase 5 -- measured", 5.12.
_UNDER_TEXT = {"table-border": 0, "cell-shading": 1, "paragraph-shading": 2, "shading": 3, "highlight": 3}


def render_page(page: Page, *, glyph_size: str = "device", images=None, width: int | None = None,
                height: int | None = None, warn=None, names=None) -> str:
    """One page as a standalone SVG document.

    ``names(face, bold, italic) -> DrawingName | None`` names a face its document name
    does not reach (:meth:`docx2svg.fonts.InstalledFonts.drawing_name`); without it,
    every face is named as the document names it.
    ``images(relationship id) -> (content type, bytes) | None`` supplies inline
    pictures; a picture it cannot supply, or in a format an SVG cannot carry (EMF,
    WMF), is drawn as a placeholder of its extent and reported through
    ``warn(code, message)``.
    """
    if glyph_size not in GLYPH_SIZES:
        raise ValueError(f"glyph_size must be one of {GLYPH_SIZES}")
    if width is None and height is None:
        size = (f' width="{number(Fraction(page.width_pt).limit_denominator(1000))}pt"'
                f' height="{number(Fraction(page.height_pt).limit_denominator(1000))}pt"')
    else:
        if width is None:
            width = round(Fraction(height * page.width_px, page.height_px))
        if height is None:
            height = round(Fraction(width * page.height_px, page.width_px))
        size = f' width="{width}" height="{height}"'
    out = [
        '<svg xmlns="http://www.w3.org/2000/svg" version="1.1"'
        f'{size} viewBox="0 0 {page.width_px} {page.height_px}" data-docx-page="{page.number + 1}">',
        f'<rect width="{page.width_px}" height="{page.height_px}" fill="#FFFFFF"/>',
    ]
    defs = getattr(page, "defs", None)
    if defs is None:
        from ooxml_common.drawingml.svg import Defs

        defs = Defs()
    head = len(out)
    # Word's order (:meth:`docx2svg.layout.Page.paint`): floating drawings behind the text
    # under everything else, table borders and shading included; those in front over
    # everything but the stop band; a header's or footer's drawings with it, before the body.
    for step in page.paint():
        if step[0] == "floats":
            for placed in step[1]:
                out.extend(_float(placed, glyph_size, names, images, warn, defs))
        elif step[0] == "under":
            out.extend(_rules([rule for rule in page.rules if rule.kind in _UNDER_TEXT]))
        elif step[0] == "lines":
            out.extend(_lines(step[1], glyph_size, names, defs))
        elif step[0] == "over":
            out.extend(_rules([rule for rule in page.rules if rule.kind not in _UNDER_TEXT], ordered=False))
        else:
            for picture in page.pictures:
                out.append(_picture(picture.path, picture.relationship, picture.x, picture.y, picture.width,
                                    picture.height, getattr(picture, "part", None), images, warn))
    if defs.defs:
        out.insert(head, "<defs>" + "".join(defs.defs) + "</defs>")
    if page.stop is not None:
        out.append(_stop(page))
    out.append("</svg>")
    return "\n".join(out) + "\n"


def _rules(rules, *, ordered: bool = True) -> list[str]:
    """Filled rectangles: the kinds drawn under the text in :data:`_UNDER_TEXT`'s order."""
    if ordered:
        rules = sorted(rules, key=lambda rule: _UNDER_TEXT.get(rule.kind, 0))
    return [_rect(rule.x, rule.y, rule.width, rule.height, _color(rule.color),
                  extra=f" data-docx-path={quoteattr(rule.path)} data-docx-kind=\"{rule.kind}\"") for rule in rules]


def _lines(lines, glyph_size: str, names, defs=None) -> list[str]:
    """Lines of text, grouped by paragraph (and story)."""
    out: list[str] = []
    current_path = None
    for line in lines:
        story = getattr(line, "story", None)
        key = (line.path, story, getattr(line, "part", None))
        if key != current_path:
            if current_path is not None:
                out.append("</g>")
            out.append(f"<g data-docx-path={quoteattr(line.path)}"
                       + (f" data-docx-id={quoteattr(line.paragraph_id)}" if line.paragraph_id else "")
                       + (f' data-docx-story="{story}" data-docx-part={quoteattr(line.part)}' if story else "")
                       + ">")
            current_path = key
        out.append(f'<g data-docx-line="{line.number}" data-docx-baseline="{line.baseline}">')
        out.extend(_text(span, glyph_size, names, defs) for span in line.spans if span.chars)
        out.append("</g>")
    if current_path is not None:
        out.append("</g>")
    return out


_IMAGE_TYPES = ("image/png", "image/jpeg", "image/gif", "image/svg+xml")


def _picture(path: str, relationship, x, y, width, height, source, images, warn, transform: str | None = None) -> str:
    """A picture part in its box, or a placeholder of the box where it cannot be carried."""
    found = None
    if images and relationship:
        found = images(relationship, source) if source else images(relationship)
    extra = f' transform="{transform}"' if transform else ""
    if found is not None and found[0] in _IMAGE_TYPES:
        content_type, data = found
        href = f"data:{content_type};base64,{base64.b64encode(data).decode('ascii')}"
        return (f'<image data-docx-path={quoteattr(path)} x="{number(x)}" y="{number(y)}" width="{number(width)}"'
                f' height="{number(height)}" preserveAspectRatio="none"{extra} href="{href}"/>')
    if warn is not None and relationship is not None:
        what = found[0] if found is not None else "a missing part"
        warn("picture-not-drawn", f"a picture in {what} is drawn as a placeholder of its extent ({path})")
    return _placeholder(x, y, width, height, path, "drawing" if relationship is None else "picture format")


def _float(placed, glyph_size: str, names, images, warn, defs=None) -> list[str]:
    """A floating drawing: its primitives (:mod:`docx2svg.drawing`), then its text box's
    shading, text and rules, in one group."""
    story = f' data-docx-story="{placed.story}" data-docx-part={quoteattr(placed.part)}' if placed.story else ""
    out = [f'<g data-docx-path={quoteattr(placed.path)} data-docx-float="{"behind" if placed.behind else "front"}"'
           f"{story}>"]
    for primitive in placed.primitives:
        transform = f' transform="{primitive.transform}"' if primitive.transform else ""
        if primitive.kind == "path":
            if primitive.attrs:
                paint = primitive.attrs
            else:
                paint = f'fill="{_color(primitive.fill)}"' if primitive.fill else 'fill="none"'
                if primitive.stroke:
                    paint += f' stroke="{_color(primitive.stroke)}" stroke-width="{number(primitive.stroke_width)}"'
            out.append(f"<path data-docx-path={quoteattr(primitive.path)} d={quoteattr(primitive.d)} {paint}"
                       f"{transform}/>")
        elif primitive.kind == "markup":
            out.append(f"<g data-docx-path={quoteattr(primitive.path)}{transform}>{primitive.markup}</g>")
        elif primitive.kind == "begin":
            out.append(f"<g data-docx-path={quoteattr(primitive.path)} {primitive.attrs}>")
        elif primitive.kind == "end":
            out.append("</g>")
        elif primitive.kind == "image":
            out.append(_picture(primitive.path, primitive.relationship, primitive.x, primitive.y, primitive.width,
                                primitive.height, placed.part, images, warn, primitive.transform))
        else:
            out.append(_placeholder(primitive.x, primitive.y, primitive.width, primitive.height, primitive.path,
                                    primitive.what or "drawing"))
    out.extend(_rules([rule for rule in placed.rules if rule.kind in _UNDER_TEXT]))
    out.extend(_lines(placed.lines, glyph_size, names, defs))
    out.extend(_rules([rule for rule in placed.rules if rule.kind not in _UNDER_TEXT], ordered=False))
    out.append("</g>")
    return out


def _placeholder(x, y, width, height, path: str, what: str) -> str:
    return (f'<g data-docx-unsupported="{what}" data-docx-path={quoteattr(path)}>'
            + _rect(x, y, width, height, "#F2F2F2",
                    extra=' stroke="#9A9A9A" stroke-width="3" stroke-dasharray="12 8"')
            + "</g>")


def _stop(page: Page) -> str:
    """Where the layout stopped: a hatched band from the obstacle to the foot of the text
    area, across the column, saying what the obstacle is and that nothing after it is
    drawn -- rather than a page that looks finished and is not."""
    stop = page.stop
    left, right, _top, bottom = page.column
    y = stop.y
    height = max(Fraction(bottom) - y, Fraction(60))
    label = (f"docx2svg: not drawn: {stop.reason}. The layout stops here; "
             f"{stop.remaining} more item(s) of the document are not laid out.")
    parts = [f'<g data-docx-unsupported={quoteattr(stop.reason)} data-docx-path={quoteattr(stop.path)}'
             f' data-docx-remaining="{stop.remaining}">',
             _rect(left, y, right - left, height, "#F2F2F2",
                   extra=' stroke="#9A9A9A" stroke-width="3" stroke-dasharray="12 8"')]
    if stop.extent is not None:
        x0, y0, w0, h0 = stop.extent
        parts.append(_rect(x0, y0, w0, h0, "none", extra=' stroke="#9A9A9A" stroke-width="3"'))
    parts.append(f'<text x="{number(left + 24)}" y="{number(y + 48)}" font-family="sans-serif" font-size="33"'
                 f' fill="#6B6B6B">{escape(label)}</text>')
    parts.append("</g>")
    return "".join(parts)


def render(layout: Layout, *, glyph_size: str = "device", images=None) -> list[str]:  # noqa: D103
    return [render_page(page, glyph_size=glyph_size, images=images) for page in layout.pages]
