"""Resolution: from what the file declares to what Word applies.

``parse`` reads each level of the hierarchy as declared; this stage merges them.  The
split is the sibling ``pptx2svg``'s, for the same reason: the cascade is where the bugs
are, and every resolved value here carries the level it came from
(:class:`~docx2svg.resolve.cascade.Resolved`), so a wrong one can be traced.
"""

from .cascade import AUTOSPACING_TWIPS, Origin, Resolved, autospaced, resolve_mark, resolve_paragraph, resolve_run, style_chain
from .fonts import (
    CharacterFormat, character_format, script_fields_taken, script_half_points, script_raise_half_points, slot_for,
    theme_face,
)

__all__ = [
    "AUTOSPACING_TWIPS",
    "CharacterFormat",
    "Origin",
    "Resolved",
    "autospaced",
    "character_format",
    "resolve_mark",
    "resolve_paragraph",
    "resolve_run",
    "script_fields_taken",
    "script_half_points",
    "script_raise_half_points",
    "slot_for",
    "style_chain",
    "theme_face",
]
