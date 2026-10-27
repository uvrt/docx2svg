"""The names the sibling repositories import from docx2svg: renaming one fails here.

Siblings' CI installs docx2svg from ``main``, so a rename that breaks one shows only at its
next run, not on the pull request that made it.  ``USED_BY`` is that surface, each name with
the siblings that import it (in their ``src``, ``tests`` and ``tools``); docs/api.md ("Used
by docx-agent") documents it, and ``test_the_documented_names_are_these`` holds the two to
the same names.

A sibling that starts importing a name: add it here and there.  Renaming one: keep the old
name as an alias for one release, with a comment saying who it is kept for (``OLD_NAMES``),
and move the sibling to the new one.
"""

from __future__ import annotations

import importlib
import re
from pathlib import Path

import pytest

DOC = Path(__file__).resolve().parent.parent / "docs/api.md"
SECTION = "## Used by docx-agent"

#: module -> {name: the siblings that import it}.
USED_BY: dict[str, dict[str, str]] = {
    "docx2svg": {
        "ConvertOptions": "docx-agent",
        "available_backends": "docx-agent",
        "convert_docx": "docx-agent",
        "convert_docx_to_layout": "docx-agent",
        "convert_docx_to_png": "docx-agent",
        "convert_docx_to_svg": "docx-agent",
        "drawn_faces": "docx-agent",
        "lay_out_docx": "docx-agent",
        "page_svgs": "docx-agent",
        "user_font_dirs": "docx-agent",
        "warm": "docx-agent",
    },
    "docx2svg.coverage": {
        "Coverage": "docx-agent",
    },
    "docx2svg.fields": {
        "Uncomputable": "docx-agent",
        "field_text": "docx-agent",
        "sequence_step": "docx-agent",
        "sequence_switches": "docx-agent",
    },
    "docx2svg.fonts": {
        "InstalledFonts": "docx-agent",
        "cloud_font_dirs": "docx-agent",
        "default_font_dirs": "docx-agent",
        "installed_index": "docx-agent",
        "rasteriser_files": "docx-agent",
    },
    "docx2svg.linebreak": {
        "ListCounters": "docx-agent",
    },
    "docx2svg.model": {
        "Paragraph": "docx-agent",
        "ParagraphProperties": "docx-agent",
        "Run": "docx-agent",
        "RunProperties": "docx-agent",
        "Table": "docx-agent",
    },
    "docx2svg.paginate": {
        "CARRY_ON": "docx-agent",
    },
    "docx2svg.parse.document": {
        "cell_conditions": "docx-agent",
        "parse_package": "docx-agent",
    },
    "docx2svg.parse.properties": {
        "read_paragraph_properties": "docx-agent",
        "read_run_properties": "docx-agent",
    },
    "docx2svg.parse.styles": {
        "Numbering": "docx-agent",
        "parse_numbering": "docx-agent",
        "parse_styles": "docx-agent",
    },
    "docx2svg.paths": {
        "path_part": "docx-agent",
        "resolve_path": "docx-agent",
    },
    "docx2svg.png": {
        "svg_to_png": "docx-agent",
    },
    "docx2svg.resolve": {
        "character_format": "docx-agent",
        "resolve_paragraph": "docx-agent",
        "resolve_run": "docx-agent",
    },
    "docx2svg.table": {
        "edge_winner": "docx-agent",
        "outer_part": "docx-agent",
        "resolve_table": "docx-agent",
        "table_style_id": "docx-agent",
    },
}

#: (module, old underscore name) -> the public name; aliases kept for one release.
OLD_NAMES: dict[tuple[str, str], str] = {
    ("docx2svg", "_lay_out"): "lay_out_docx",
    ("docx2svg.parse.document", "_conditions"): "cell_conditions",
}


def _used() -> list[tuple[str, str]]:
    return [(module, name) for module, names in USED_BY.items() for name in names]


@pytest.mark.parametrize(("module", "name"), _used())
def test_a_name_a_sibling_imports_is_there(module, name):
    owner = importlib.import_module(module)
    if hasattr(owner, name):
        return
    try:
        importlib.import_module(f"{module}.{name}")     # a submodule not imported yet
    except ImportError:
        pytest.fail(f"{module}.{name} is gone, and {USED_BY[module][name]} import it: keep the "
                    "old name as an alias for one release and move them to the new one")


def test_siblings_import_no_private_name():
    private = [f"{module}.{name}" for module, name in _used() if name.startswith("_")]
    assert not private, f"give these a public name (keep the old one as an alias): {private}"


def test_an_old_name_is_the_public_one():
    for (module, old), new in OLD_NAMES.items():
        owner = importlib.import_module(module)
        assert getattr(owner, old) is getattr(owner, new), f"{module}.{old} is not {new}"
        assert new in USED_BY.get(module, {}), f"{module}.{new} is missing from USED_BY"


def _documented() -> set[tuple[str, str]]:
    lines = DOC.read_text(encoding="utf-8").splitlines()
    if SECTION is not None:
        start, depth = lines.index(SECTION) + 1, len(SECTION) - len(SECTION.lstrip("#"))
        end = next((i for i in range(start, len(lines)) if lines[i].startswith("#")
                    and len(lines[i]) - len(lines[i].lstrip("#")) <= depth), len(lines))
        lines = lines[start:end]
    found = set()
    for line in lines:
        if line.startswith("- `"):
            module, *names = re.findall(r"`([^`]+)`", line)
            found.update((module, name) for name in names)
    return found


def test_the_documented_names_are_these():
    documented, used = _documented(), set(_used())
    assert not used - documented, f"add to {DOC.name}: {sorted(used - documented)}"
    extra = sorted(documented - used)
    assert not extra, f"add to USED_BY, or take out of {DOC.name}: {extra}"
