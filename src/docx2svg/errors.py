"""What a layout raises when laying an element out fails."""

from __future__ import annotations


class LayoutError(RuntimeError):
    """Laying an element out failed unexpectedly: ``path`` is the element's path (a table's,
    ``w:body/w:tbl[3]``), ``reason`` what went wrong (the failure's type and message).  The
    failure itself is the exception's ``__cause__``.  A layout raises it in place of the
    failure, so a caller can name the element (docx-agent's ``layout_failed``); what the
    model does not lay out yet is never one (``ConvertOptions.warnings`` says so instead)."""

    def __init__(self, path: str, reason: str) -> None:
        super().__init__(path, reason)
        self.path = path
        self.reason = reason

    def __str__(self) -> str:
        return f"laying out {self.path} failed: {self.reason}"


def failure(path: str, error: BaseException) -> LayoutError:
    """``error`` met laying out the element at ``path``, as a :class:`LayoutError`."""
    return LayoutError(path, f"{type(error).__name__}: {error}")
