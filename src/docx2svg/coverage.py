"""How much of a document a layout covers: the summary a caller reads before trusting it.

The layout never guesses.  Where it meets what it cannot lay out yet it stops -- the body
at the first obstacle (``layout-stopped:<reason>``), a header, footer or text box at its
own (``story-stopped:<reason>``) -- and an absent face is laid out with its open
substitute or not at all (:mod:`docx2svg.fonts`).  Each of those is a warning; this is
their sum, so that "laid out and nothing was wrong" is told apart from "could not lay it
all out" without reading every warning:

* ``pages`` laid out and ``estimated_pages`` -- the layout's own count when it reached the
  end, else Word's count from ``docProps/app.xml`` as last saved (``estimate_source`` says
  which), else ``None``: past a stop the pages are not known, and none is invented;
* ``blocks`` (the body's top-level paragraphs and tables), ``blocks_laid_out`` and
  ``blocks_skipped`` -- a table the layout stopped in counts as skipped;
* ``stop`` -- the first body stop, with its reason, message, page and element path;
  ``story_stops`` -- every header, footer, note or text box drawn only up to a point;
* ``substituted_fonts`` and ``missing_fonts``.

``complete`` is true only when every block was laid out, no story stopped and no face was
missing.  A complete layout with substituted faces is complete: what it says about
pagination holds as far as the substitutes are metric compatible
(``substituted_fonts[...]["metric_compatible"]``).

Standard library only.
"""

from __future__ import annotations

import io
import re
import zipfile
from dataclasses import asdict, dataclass, field


@dataclass(frozen=True)
class StopInfo:
    """One stop: the warning's ``code`` (``layout-stopped:table``...), its ``reason`` (the
    part after the colon), its ``message``, the 1-based ``page`` and, for the body's stop,
    the ``path`` of the element it stopped at (:mod:`docx2svg.paths`)."""

    code: str
    reason: str
    message: str
    page: int | None = None
    path: str | None = None


@dataclass
class Coverage:
    complete: bool
    pages: int
    estimated_pages: int | None
    #: ``layout`` (the layout reached the end), ``app.xml`` (Word's count as last saved,
    #: which an edit since may have changed), or ``None`` (not known).
    estimate_source: str | None
    blocks: int
    blocks_laid_out: int
    blocks_skipped: int
    stop: StopInfo | None = None
    story_stops: list[StopInfo] = field(default_factory=list)
    #: ``{"family", "substitute", "metric_compatible"}`` per substituted face.
    substituted_fonts: list[dict] = field(default_factory=list)
    #: Faces asked for that are not installed, not embedded and have no substitute here.
    missing_fonts: list[str] = field(default_factory=list)

    def as_dict(self) -> dict:
        return asdict(self)

    def summary(self) -> str:
        """One line: what was laid out, and why not everything."""
        if self.complete:
            text = f"complete: {self.pages} page(s), {self.blocks} block(s)"
        else:
            estimate = f" of ~{self.estimated_pages}" if self.estimate_source == "app.xml" else ""
            text = (f"partial: {self.pages}{estimate} page(s), {self.blocks_laid_out} of {self.blocks} block(s) "
                    "laid out")
            if self.stop is not None:
                text += f"; stopped on page {self.stop.page} ({self.stop.reason})"
            if self.story_stops:
                text += f"; {len(self.story_stops)} header/footer/box stop(s)"
            if self.missing_fonts:
                text += f"; missing faces: {', '.join(self.missing_fonts)}"
        if self.substituted_fonts:
            text += "; substituted: " + ", ".join(f"{s['family']} -> {s['substitute']}"
                                                  + ("" if s["metric_compatible"] else " (approximate)")
                                                  for s in self.substituted_fonts)
        return text


def saved_pages(package: bytes | None) -> int | None:
    """``<Pages>`` of ``docProps/app.xml``: Word's page count when the file was last saved
    (any other writer's, or none)."""
    if not package:
        return None
    try:
        with zipfile.ZipFile(io.BytesIO(package)) as archive:
            text = archive.read("docProps/app.xml").decode("utf-8", "replace")
    except (KeyError, zipfile.BadZipFile, OSError):
        return None
    found = re.search(r"<(?:\w+:)?Pages>\s*(\d+)\s*</", text)
    return int(found.group(1)) if found and int(found.group(1)) > 0 else None


def _reason(code: str) -> str:
    return code.split(":", 1)[1] if ":" in code else code


def coverage_of(layout, document, package: bytes | None = None, fonts=None) -> Coverage:
    """The :class:`Coverage` of ``layout`` (:class:`docx2svg.layout.Layout`) of
    ``document``; ``fonts`` the :class:`docx2svg.fonts.InstalledFonts` it was measured
    with, if any."""
    body = list(document.body)
    stop = None
    for page in layout.pages:
        if page.stop is not None:
            message = next((m for c, m, p in layout.warnings if c.startswith("layout-stopped:")
                            and p == page.number + 1), "")
            code = next((c for c, m, p in layout.warnings if c.startswith("layout-stopped:")
                         and p == page.number + 1), f"layout-stopped:{page.stop.reason}")
            stop = StopInfo(code, page.stop.reason, message, page.number + 1, page.stop.path)
            break
    if stop is None:
        found = next(((c, m, p) for c, m, p in layout.warnings if c.startswith("layout-stopped:")), None)
        if found is not None:
            stop = StopInfo(found[0], _reason(found[0]), found[1], found[2])
    story_stops = [StopInfo(c, _reason(c), m, p) for c, m, p in layout.warnings if c.startswith("story-stopped:")]

    laid_out = len(body)
    if stop is not None:
        laid_out = 0
        if stop.path:
            for index, block in enumerate(body):
                path = getattr(block, "path", "")
                if path and (stop.path == path or stop.path.startswith(path + "/")):
                    laid_out = index
                    break
            else:
                laid_out = max(0, len(body) - (layout.pages[stop.page - 1].stop.remaining if stop.page else 0))

    # Any measuring object may stand in for InstalledFonts; one that substitutes nothing says so.
    substituted = [s.as_dict() for s in getattr(fonts, "substitutions", {}).values()]
    missing = sorted(getattr(fonts, "missing", ()))
    complete = stop is None and not story_stops and not missing
    pages = len(layout.pages)
    if stop is None:
        estimate, source = pages, "layout"
    elif (saved := saved_pages(package)) is not None:
        estimate, source = max(saved, pages), "app.xml"
    else:
        estimate, source = None, None
    return Coverage(complete, pages, estimate, source, len(body), laid_out, len(body) - laid_out, stop, story_stops,
                    substituted, missing)
