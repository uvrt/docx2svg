# Changelog

docx2svg has not been released to PyPI. The public repository starts from a single
snapshot commit; the development history before it is summarised here.
[`ROADMAP.md`](ROADMAP.md) records every measurement in detail.

## Unreleased

- **Python 3.14 and 3.15.** CI runs the suite on both, on Linux, macOS and Windows, and
  the classifiers declare them. `requires-python` stays `>=3.10`; no library change, and
  the `png` extra's resvg-py has wheels for both. CPython 3.14's Windows builds deflate
  with zlib-ng, whose bytes differ from zlib's for the same entries (both valid; what a
  package holds is unchanged), so CI's check that `tests/fixtures/` is current leaves the
  `.docx` out of its diff there, where `tests/test_parse.py` compares it part by part as
  it does everywhere.
- **An application's own font folder, from the environment too.** Without
  `ConvertOptions.font_dirs`, the `OOXML_FONT_DIRS` environment variable (folders
  separated by `os.pathsep`, shared with pptx2svg) is read, for layout and drawing alike,
  and by `svg_to_png`; an explicit empty list means none. Either is added to the folders
  Word uses and searched after them, as `font_dirs` always was, and the folders under it
  are now read too. Requires ooxml-common 0.8.
- **Fonts without Office.** An Office face that is neither installed nor embedded is laid
  out and drawn with its open metric compatible substitute, if that is installed: Carlito
  for Calibri, and Liberation Sans, Serif and Mono for Arial, Times New Roman and Courier
  New. Each pair was measured equal in advances and line metrics against Word's copy.
  Every substitution is reported (`font-substituted`), and none is made when the real face
  is present. An absent Symbol or Wingdings is laid out from the integers recorded from
  Word's copies (`docx2svg.recorded`, `tools/record_symbol_faces.py`), and its bullets are
  drawn as Unicode equivalents. `ConvertOptions.substitute_fonts` (`--no-substitute-fonts`)
  turns substitution off, and `ConvertOptions.font_substitutes` names approximate
  substitutes. The README has a new "Fonts" section.
- **Coverage.** `ConvertOptions.coverage` / `Layout.coverage` (`docx2svg.coverage`) say
  how much was laid out: the pages, the blocks laid out and skipped, the first stop with
  its reason and path, header, footer and text-box stops, and the faces substituted or
  missing. The CLI prints its summary.
- **A drawing text wraps around in a cell aligned vertically or merged** (ROADMAP F.25) is
  laid out as Word lays it out, where the table stopped: the cell laid out from its top,
  then its lines and drawings moved down together by `w:vAlign`, the drawing's foot
  counting in the height aligned; such a drawing kept inside its cell; and a line with no
  room beside a `wrapTight` / `wrapThrough` drawing stepping down a line at a time. A
  drawing positioned as no probe measured in a cell is **approximated** instead of
  stopping the table: `Coverage.approximations` lists each (reason, page, path),
  `Coverage.status` is `complete`, `approximate` or `partial`, and the warning is
  `layout-approximate:cell-drawing`.
- CI installs Carlito, Caladea and Liberation on Linux and macOS, so the layout code runs
  there past the first paragraph. `tests/test_substitutes.py` checks Carlito's layout of
  `layout-sweep.docx` glyph for glyph against Word's recorded Calibri numbers.

- Chart labels as Word draws them (ooxml-common 0.7): tick, category and legend text in
  its `c:txPr` colour (`tx1` at 65%, as Office writes it), data labels in their own number
  format (`€41.2m`), and a radar's category labels 4% of the radius off their vertex.

## 0.1.0 -- 2026-10-08 (initial public release)

Developed 2026-09-24 to 2026-10-07, phase by phase against Microsoft Word's PDF export:

- **Phase 0:** the Word oracle (AppleScript export to vector PDF, measured with
  pypdfium2), a generated fixture that is ours, and a loop that closes.
- **Phase 1:** the format-neutral half shared with pptx2svg extracted into
  [`ooxml-common`](https://github.com/uvrt/ooxml-common).
- **Phases 2-4:** the vertical model, line breaking to an exact inclusive budget,
  hyphenation (English, Dutch, German, French, from openly licensed patterns), and
  pagination with keeps, footnotes, endnotes, tables, text columns and floating drawings.
- **Phase 5:** SVG and PNG output, every glyph checked against Word's PDF; headers,
  footers and computed fields; charts and SmartArt; tracked changes in Word's final view.
- **API:** `convert_docx` (layout and SVGs from one layout), `data-docx-path` element
  paths with `docx2svg.paths.locate`, `Page.info`, list counters without a document.
