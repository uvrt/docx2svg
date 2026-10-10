"""The integers Word lays two symbol faces out with, and what their symbols are drawn as.

Both moved to ooxml-common, so pptx2svg lays out and draws an absent symbol face the same
way: the recorded integers to :mod:`ooxml_common.text.recorded_symbol_faces` (still
written by ``tools/record_symbol_faces.py``), the Unicode equivalents -- now every glyph of
Symbol, Wingdings, Wingdings 2, Wingdings 3 and Webdings, from the Unicode Consortium's
published tables -- to :mod:`ooxml_common.text.symbol_fonts`.  :class:`docx2svg.fonts.InstalledFonts`
falls back to these only where neither the face nor an embedded copy of it is found, and
reports it (``font-substituted``).
"""

from ooxml_common.text.recorded_symbol_faces import ADVANCES, DECORATIONS, METRICS
from ooxml_common.text.symbol_fonts import to_unicode

__all__ = ["ADVANCES", "DECORATIONS", "METRICS", "UNICODE"]

#: family -> {code point: what is drawn in its place}, for every code recorded.
UNICODE = {
    family: {
        code: text
        for code in sorted(advances)
        if (text := to_unicode(family, chr(code))) is not None
    }
    for family, advances in ADVANCES.items()
}
