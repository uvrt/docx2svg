"""SVG to PNG, delegated to an existing rasteriser (the ``png`` extra: resvg-py).

The same arrangement as the sibling ``pptx2svg``'s ``png.py``, and the same code: the
backends, their order and probing, the call to each and the passes a plan is drawn in are
:mod:`ooxml_common.raster`'s.  resvg-py first (prebuilt wheels, no system library, and
the only backend that takes font arguments), cairosvg as a fallback for hosts that
already have Cairo.

**Fonts.**  The SVG names its faces and embeds none.  Word lays a document out with
Microsoft's faces, and this library measured every advance with them, so a raster is
only faithful where the rasteriser draws with the same files.  There is no bundle to be
reproducible with -- those faces are not ours to ship -- so resvg reads the host's
installed fonts, plus any ``font_dirs`` / ``font_files`` given (a document's embedded
faces among them).  A face the host lacks is reported by the layout, not here: without
its advances the layout cannot measure the paragraph and stops there with a warning.
"""

from __future__ import annotations

from typing import Sequence

from ooxml_common import raster
from ooxml_common.raster import (  # noqa: F401  (this module's public names, kept)
    Backend,
    RasterizerNotAvailable,
    available_backends,
)

_BACKEND_MODULES = raster.BACKENDS
_importable = raster.importable
_as_bytes = raster.as_bytes


def svg_to_png(
    svg: str,
    *,
    width: int | None = None,
    height: int | None = None,
    scale: float | None = None,
    background: str | None = None,
    backend: Backend = "auto",
    font_dirs: Sequence[str] | None = None,
    font_files: Sequence[str] | None = None,
    skip_system_fonts: bool = False,
    sans_serif_family: str | None = None,
) -> bytes:
    """Rasterise one SVG document to PNG bytes.

    ``width``/``height`` set the output size in pixels (give one and the aspect ratio is
    kept); ``scale`` multiplies the SVG's own size instead.  ``font_dirs`` and
    ``font_files`` are searched before the host's fonts, which ``skip_system_fonts``
    leaves out; ``sans_serif_family`` names the face the generic ``sans-serif`` means
    (the placeholders' labels use it).  Only resvg takes the font arguments.
    ``font_dirs`` left out reads ``OOXML_FONT_DIRS`` (``os.pathsep``-separated); an empty
    list means none.
    """
    from ooxml_common.fonts.office import user_font_dirs
    from ooxml_common.metafile import inline_metafile_images

    # A metafile drawn as SVG goes inline: resvg draws no text in an SVG it reads through
    # <image>.
    svg = inline_metafile_images(svg)
    font_dirs = [str(path) for path in user_font_dirs(font_dirs)]
    chosen = raster.choose_backend(
        backend,
        none_installed="no SVG rasterizer installed; run `pip install docx2svg[png]` for resvg-py "
        "(prebuilt wheels, no system dependencies)")
    if chosen == "resvg":
        from .fonts import drawing_plan

        plan = drawing_plan(svg, font_files, font_dirs)
        if plan is not None:
            return _render_plan(svg, plan, width=width, height=height, scale=scale, background=background,
                                font_dirs=font_dirs, font_files=font_files, skip_system_fonts=skip_system_fonts,
                                sans_serif_family=sans_serif_family)
        return _render_with_resvg(svg, width=width, height=height, scale=scale, background=background,
                                  font_dirs=font_dirs, font_files=font_files, skip_system_fonts=skip_system_fonts,
                                  sans_serif_family=sans_serif_family)
    if chosen == "cairosvg":
        return _render_with_cairosvg(svg, width=width, height=height, scale=scale, background=background)
    raise ValueError(f"unknown rasterizer backend: {backend!r}")  # pragma: no cover


def _render_with_resvg(svg: str, *, width, height, scale, background, font_dirs, font_files,
                       skip_system_fonts: bool, sans_serif_family: str | None = None) -> bytes:
    """resvg's call (:func:`ooxml_common.raster.render_with_resvg`), the generic
    ``sans-serif`` pointed at ``sans_serif_family`` where one is given."""
    return raster.render_with_resvg(
        svg, width=width, height=height, scale=scale, background=background, font_dirs=font_dirs,
        font_files=font_files, skip_system_fonts=skip_system_fonts,
        families={"sans_serif_family": sans_serif_family} if sans_serif_family else None,
        not_installed="resvg-py is not installed; run `pip install docx2svg[png]`")


def _render_plan(svg: str, plan, *, width, height, scale, background, font_dirs, font_files,
                 skip_system_fonts: bool, sans_serif_family: str | None) -> bytes:
    """Draw ``svg`` in the passes of ``plan`` (:func:`docx2svg.fonts.drawing_plan`), each
    over the last (:func:`ooxml_common.raster.draw_in_passes`): a pass that leaves files
    out is handed the files and the folders' files in the order resvg loads them, without
    those."""
    from ooxml_common.fonts.rasterdb import load_order

    def draw(number: int, document: str) -> bytes:
        withheld = plan.withheld[number]
        files, dirs = font_files, font_dirs
        if withheld:
            files = [path for path in load_order([str(path) for path in font_files or ()],
                                                 [str(path) for path in font_dirs or ()])
                     if path not in withheld]
            dirs = []
        return _render_with_resvg(document, width=width, height=height, scale=scale,
                                  background=background if number == 0 else None, font_dirs=dirs,
                                  font_files=files, skip_system_fonts=skip_system_fonts,
                                  sans_serif_family=sans_serif_family)

    return raster.draw_in_passes(svg, plan.stacks, plan.passes, draw)


def pass_svg(svg: str, plan, number: int, underlay: bytes | None) -> str:
    """``svg`` for pass ``number`` of ``plan`` (:func:`ooxml_common.raster.pass_svg`), as
    the sibling ``pptx2svg`` draws its passes: pass 0 is the SVG with the text of every
    later pass hidden; a later pass is the SVG with everything hidden but its own text,
    over ``underlay`` -- the passes before it, as PNG, drawn pixel for pixel beneath."""
    return raster.pass_svg(svg, plan.stacks, number, underlay)


def _render_with_cairosvg(svg: str, *, width, height, scale, background) -> bytes:
    return raster.render_with_cairosvg(svg, width=width, height=height, scale=scale, background=background)
