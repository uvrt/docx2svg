"""SVG to PNG, delegated to an existing rasteriser (the ``png`` extra: resvg-py).

The same arrangement as the sibling ``pptx2svg``'s ``png.py``, written here rather than
imported: resvg-py first (prebuilt wheels, no system library, and the only backend that
takes font arguments), cairosvg as a fallback for hosts that already have Cairo.

**Fonts.**  The SVG names its faces and embeds none.  Word lays a document out with
Microsoft's faces, and this library measured every advance with them, so a raster is
only faithful where the rasteriser draws with the same files.  There is no bundle to be
reproducible with -- those faces are not ours to ship -- so resvg reads the host's
installed fonts, plus any ``font_dirs`` / ``font_files`` given (a document's embedded
faces among them).  A face the host lacks is reported by the layout, not here: without
its advances the layout cannot measure the paragraph and stops there with a warning.
"""

from __future__ import annotations

from typing import Iterable, Literal, Sequence

Backend = Literal["resvg", "cairosvg", "auto"]


class RasterizerNotAvailable(RuntimeError):
    """No SVG-to-PNG backend is installed."""


_BACKEND_MODULES: tuple[tuple[str, str], ...] = (
    ("resvg", "resvg_py"),
    ("cairosvg", "cairosvg"),
)


def _importable(module: str) -> bool:
    """Whether ``module`` imports, treating *any* failure as unavailable: cairosvg with
    its wheel installed and libcairo missing raises ``OSError``, not ``ImportError``
    (the sibling project lost every caller to that once)."""
    try:
        __import__(module)
    except Exception:
        return False
    return True


def available_backends() -> list[str]:
    """The rasterisation backends that import right now, best first."""
    return [name for name, module in _BACKEND_MODULES if _importable(module)]


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

    font_dirs = [str(path) for path in user_font_dirs(font_dirs)]
    chosen = backend
    if chosen == "auto":
        found = available_backends()
        if not found:
            raise RasterizerNotAvailable(
                "no SVG rasterizer installed; run `pip install docx2svg[png]` for resvg-py "
                "(prebuilt wheels, no system dependencies)")
        chosen = found[0]  # type: ignore[assignment]
    if chosen == "resvg":
        return _render_with_resvg(svg, width=width, height=height, scale=scale, background=background,
                                  font_dirs=font_dirs, font_files=font_files, skip_system_fonts=skip_system_fonts,
                                  sans_serif_family=sans_serif_family)
    if chosen == "cairosvg":
        return _render_with_cairosvg(svg, width=width, height=height, scale=scale, background=background)
    raise ValueError(f"unknown rasterizer backend: {backend!r}")


def _render_with_resvg(svg: str, *, width, height, scale, background, font_dirs, font_files,
                       skip_system_fonts: bool, sans_serif_family: str | None = None) -> bytes:
    try:
        import resvg_py
    except ImportError as error:  # pragma: no cover - only without the extra
        raise RasterizerNotAvailable("resvg-py is not installed; run `pip install docx2svg[png]`") from error
    options: dict = {"svg_string": svg}
    if width is not None:
        options["width"] = int(width)
    if height is not None:
        options["height"] = int(height)
    if scale is not None:
        options["zoom"] = float(scale)
    if background is not None:
        options["background"] = background
    if font_dirs:
        options["font_dirs"] = [str(path) for path in font_dirs]
    if font_files:
        options["font_files"] = [str(path) for path in font_files]
    if skip_system_fonts:
        options["skip_system_fonts"] = True
    if sans_serif_family:
        options["sans_serif_family"] = sans_serif_family
    return _as_bytes(resvg_py.svg_to_bytes(**options))


def _as_bytes(value) -> bytes:
    """resvg-py has returned both ``bytes`` and ``list[int]`` across releases."""
    if isinstance(value, (bytes, bytearray)):
        return bytes(value)
    if isinstance(value, Iterable):
        return bytes(value)
    raise TypeError(f"unexpected rasterizer result type: {type(value)!r}")


def _render_with_cairosvg(svg: str, *, width, height, scale, background) -> bytes:
    try:
        import cairosvg
    except ImportError as error:  # pragma: no cover - only without the extra
        raise RasterizerNotAvailable("cairosvg is not installed") from error
    options: dict = {"bytestring": svg.encode("utf-8")}
    if width is not None:
        options["output_width"] = int(width)
    if height is not None:
        options["output_height"] = int(height)
    if scale is not None:
        options["scale"] = float(scale)
    if background is not None:
        options["background_color"] = background
    return cairosvg.svg2png(**options)
