# Changelog

docx2svg has not been released to PyPI. Each version that reaches `main` is tagged
`v<version>`, from 0.2.0, the first tagged. The public repository starts from a single
snapshot commit; the development history before it is summarised here.
[`ROADMAP.md`](ROADMAP.md) records every measurement in detail.

## 0.6.0 -- unreleased

- **Pages and blocks, both ways** (production round 19, R3): which body blocks each page
  holds, and which pages each block is on, keyed by what the XML alone gives -- the path of
  the block's element in the main part, as `data-docx-path` counts it, and a paragraph's
  `w14:paraId`. `Page.blocks` lists the keys of every paragraph, table, table row and cell
  the page draws something of (a table's key, then its rows', cells' and their paragraphs'),
  once each, in the order the page draws them; `Layout.blocks_on(page)` the top-level ones;
  `Layout.pages_of(key)` every 1-based page a block is drawn on -- each page of a paragraph,
  table or row split across pages, each page a header row is drawn again on -- by path (a
  run's path is taken as its paragraph's) or by paraId. `paths.block_keys(path)` gives the
  keys of a drawn object's path. Read from what each page draws (its body lines, floating
  drawings and their text boxes, placeholders, stop), so nothing is added to the layout's
  work; headers, footers and notes are not the body's. `tools/report_docs.py` writes the
  made-up reports the tests use (headings, lists, tables with a header row, side panels in
  text boxes text wraps beside, a header and a `Page N of M` footer). Rendering is
  unchanged.

## 0.4.2 -- unreleased

- **A font scan skips a file without a `head` table again** (hotfix after S8): since
  ooxml-common 0.33.1 its `Face` raises `FontError`, not `TypeError`, for a font that lacks
  `head`, `hhea` or `hmtx`. `fonts.installed_index` skipped only `struct.error` and
  `TypeError`, so macOS's `Supplemental/NISC18030.ttf` (no `head`) stopped the whole scan on
  macOS. It now skips `FontError` too, as the other scans here already did; a test builds
  such a font in memory. Rendering is unchanged.

## 0.4.1 -- unreleased

- **A linter, and `py.typed`** (review, S8): ruff checks pyflakes and a few bug-finding rules
  (`[tool.ruff]` in pyproject.toml; no formatter, no import sorting), as a step of CI's build
  job, pinned in the dev extra. What it found is fixed: unused imports and locals (in
  `layout.py`, `lines.py`, `fonts.py`'s second `import os`, the tests and the tools) and
  f-strings without a placeholder. `units.emu_to_twips`, which nothing called, is gone. The
  wheel and sdist now carry `py.typed`, so a caller's type checker reads the annotations.
  mypy is not run here yet. Rendering is unchanged.

## 0.4.0 -- unreleased

- **One raster backend for both renderers** (review, S7): `png.py` was written after
  pptx2svg's and copied its mechanism. Finding the backends, calling resvg-py and cairosvg
  with their options, `RasterizerNotAvailable`, and drawing a page in passes
  (`pass_svg`) are now `ooxml_common.raster`, which pptx2svg uses too. `png.py` keeps what is
  Word's: the plan of which text goes in which pass (`fonts.drawing_plan`), and the
  generic `sans-serif` named for the placeholders' labels. Every public name stays, with
  what it raises. `RasterizerNotAvailable` is now ooxml-common's class, so it is pptx2svg's
  as well. Two things are now as pptx2svg had them. `backend="auto"` stops probing at the
  first backend that imports, instead of importing cairosvg too. A pass leaves out the
  texts it draws nothing of (`ooxml_common.raster.without_undrawn_text`), which resvg would
  lay out and not draw. PNGs are byte for byte what they were, per backend: every fixture
  document at 500 px, plus a small SVG with each size, scale, background and
  `sans_serif_family` option, through resvg and through cairosvg (118 images), and five
  pages drawn in two passes (a heading face filed under the body face's name), through resvg.
  `tools/fidelity.py` adds `ooxml_common/raster.py`'s source to the cache key of our raster.
  The ooxml-common floor is raised to 0.33.0.

## 0.3.0 -- unreleased

- **Packages and XML are read through ooxml-common** (review, S6): `docx2svg/opc.py` and
  `docx2svg/xmlutil.py` were copies of `ooxml_common.opc` and `ooxml_common.xmlutil` with
  other contracts (a missing part raised in one and was `None` in the other; `children`
  was an iterator in one and a list in the other). The ZIP, content types and
  relationships are now `ooxml_common.opc.OpcPackage`'s, and the XML helpers
  `ooxml_common.xmlutil`'s, with `children` and `child` its new `element_children` and
  `element_child` (ooxml-common 0.32.0, the dependency is raised to it), which step over
  comments and processing instructions as docx2svg's own did: docx-agent hands
  `paths.resolve_path` and the property readers lxml trees, which keep them. Where the
  shared reader answers otherwise, docx2svg says what it read before, in `docx2svg.opc`:
  `main_part` (the root relationship's office document, else `word/document.xml`),
  `read_part` (a missing part raises `KeyError`) and `content_type` (`""` where
  `[Content_Types].xml` says nothing, so a picture whose type the package does not state
  is still not drawn). `docx2svg.opc` keeps the Word relationship types;
  `docx2svg.xmlutil` keeps `attr_int` and `attr_bool` and re-exports the shared helpers.
  `paths.path_part` takes an `OpcPackage` or, as before, anything with
  `main_document_part` and `related`. Rendering output is byte-identical (the VRT
  snapshots, and 236 documents rendered before and after).
- **Kept for one release:** `docx2svg.Package` (`docx2svg.opc.Package`), now an
  `OpcPackage` with the old names on it (`main_document_part`, `read` raising, `exists`,
  `related`, `part_by_id`, `resolve`, `rels_part_for`, `content_type` giving `""`). Its
  relationships are ooxml-common's (`target_part`, `is_external`, where the old ones had
  `target_mode` and `external`), and it no longer holds the ZIP open (`archive` is gone).
  `docx2svg.xmlutil.descendants` is ooxml-common's, which leaves out the element it starts
  from; nothing here used it. An `OpcPackage` reads every part when it opens, where the old
  reader read each on first use: the 40-page timing document lays out and draws in the
  same time.

## 0.2.0 -- 2026-10-25

- **Versions that mean something** (review, N3): this is 0.2.0, the first version tagged
  (`v0.2.0`), and docx-agent now asks for `docx2svg>=0.2.0`. Each version that reaches
  `main` is tagged by the new tag workflow, and CI's `build` job checks that
  `pyproject.toml`'s version is the one this file's top version section names
  (`tools/check_version.py`). The dependency notes moved from `pyproject.toml` to
  CONTRIBUTING.md, "Dependencies". No code change.

- **The names docx-agent imports are declared and held** (review, N2): docs/api.md, "Used
  by docx-agent", lists the 44 names docx-agent's `src`, `tests` and `tools` import from
  here, and `tests/test_public_api.py` imports each one, so a rename fails this
  repository's CI rather than docx-agent's next run; the test holds the section and its
  list to the same names. `parse.document.cell_conditions` is the public name of
  `_conditions`, which stays as an alias for one release, as `_lay_out` does for
  `lay_out_docx`. docx2svg now uses ooxml-common 0.28.0's public
  `imagemeta.pixels_and_density` (the dependency is raised to it), and its tests and
  tools `fonts.sfnt.read_tables` and `text.measure.font_kern_units`. No behaviour change.
- **CI runs nightly** (02:47 UTC) and on demand, so a change in ooxml-common that breaks
  this suite shows the next morning; ooxml-common's own CI now runs this suite on its
  pull requests too. No code change.
- **The layout carries on past what it cannot lay out** (round 16). No single unsupported
  item stops the document or a header, footer or text box any more: each is laid out by an
  approximate rule (`layout-approximate:<reason>`) or drawn as a placeholder -- a dashed grey
  box of its stated or estimated height, `data-docx-placeholder="1"`, in the flow where it
  stands (`layout-placeholder:<reason>`, page and element path, `Coverage.placeholders`) --
  and everything after it is laid out. A table the measured rules refuse is laid out on its
  `w:tblGrid`, a floating table nested in a cell against its cell (as Word draws one,
  measured), one in columns in the flow, a frame as a
  paragraph, a cell paragraph or nested table it still cannot lay out as a placeholder in
  its cell; a row taller than the page is drawn whole and the table goes on. The table of
  every stop path, before and after, is in ROADMAP.md ("Carrying on past what is not laid
  out"). `paginate.CARRY_ON = False` stops as before. **Changes the output** of every
  document that stopped: it now has all its pages. The probe scorers keep the measured
  rules' scores (`strict_layout`); `tests/test_pages.py` likewise.
- **An absent face is laid out with a default substitute** (round 16): the face pptx2svg
  uses for it (`ooxml_common.text.fontmap.substitution_for`: Carlito for Calibri Light),
  then that face's Office original (Calibri), else the first installed face of its class
  (font table `w:family` / `w:pitch`, or its name), reported as approximate
  (`font-substituted`, `Substitution.kind`). Substitutes named are followed through
  (Calibri Light -> Calibri -> Carlito). A slot that names no face is Word's fallback (Times
  New Roman for complex script, MS Mincho for East Asian, as measured); a character a face
  lacks is measured in another face (`font-glyph-fallback`). None of these stops a
  paragraph any more.
- **A watermark's letters are kerned, its faces' outlines read, its fit and face as Word
  decides them** (round 16, F5c; measured in Word 16.106's PDF, ROADMAP.md): a fitted text
  path's letters are set at their advances with the face's pair kerning (`GPOS` `kern`, else
  the legacy table), one SVG `x` per letter -- "AVATAR" in Calibri was up to 11.8 pt off
  Word's glyphs, now 0.4. Outline boxes are exact and read from CFF faces too
  (`docx2svg.outlines`, standard library): an OpenType CFF face is drawn fitted instead of
  approximate. Italic in a face with no italic is the upright face slanted by 0.34 in font
  units, fitted slanted. Whether a text path fits is its shape type's `fitshape` alone (Word
  ignores the shape's own), and one whose type does not fit is **not drawn** (Word draws
  nothing). "Times New Roman" or no `font-family` is drawn in the theme's minor face.
  **Changes the output** of watermarks with kerned pairs, in CFF faces, in synthetic italic,
  in Times New Roman or none, and of text paths that do not fit.
- **A watermark's letters fill its shape, as Word draws them** (round 16, F5b; measured in
  Word 16.106's PDF of a watermark Word made): a text path that fits its shape (`fitshape`,
  from its own `v:textpath` or its `v:shapetype`, Word's `_x0000_t136` where that is not at
  hand) is stretched on both axes until its glyphs' outlines (`glyf` boxes,
  `fonts.Face.ink`) fill the shape's box less 0.1 in left and right and 0.05 in top and
  bottom, centred. It kept the face's proportions, so the letters were 116 pt tall in a
  194.6 pt box and 9 pt off centre; now 187.4 pt, centred, where Word draws 187.0 (along:
  430.4 pt, Word 430.5). No longer reported approximate where the face's outlines are at
  hand. The PNG is now also given the watermark's face (`drawn_faces`), which it drew in a
  fallback face before. **Changes the output** of every document with such a watermark.
- **VML is read** (round 16, `docx2svg.parse.vml`): a watermark's text path is drawn behind
  the text on every page of its header (stretched to its box, its colour, opacity and
  rotation; approximate), a header's other content goes on after it; an embedded object's
  preview picture, VML rectangles, ovals and text boxes are drawn; positioned VML floats,
  other VML is a character of its box, and VML stating no size is a placeholder.
- **Cropped pictures** (`a:srcRect`, negative sides padding) are drawn cropped, floating
  and inline, as pptx2svg crops (round 16). A floating one was drawn uncropped with
  `drawing-not-drawn:cropped`; an inline one uncropped, silently.
- **Footnotes on a page shared with columns** (round 16): the notes of a one-column section
  after a continuous break -- a table cell's or a paragraph's -- were laid out but not drawn
  on a page holding a section of several columns; they are drawn at the page's foot.
  `make_columns_probe.py`'s `notesregion` scores rise (`columns-observations.json`).
- **`SEQ` with its switches before the identifier** (round 16, F4; measured in Word 16.106):
  `SEQ \# "00" Item` counts `Item` and draws 01, 02, as Word does, where it was "names no
  sequence" (`fields.sequence_switches`, `fields.sequence_step`, which docx-agent's field
  update now shares). A switch's argument is the token after it whatever it is (`SEQ \r Mm`
  names none); a `\#` after a `\*` is not computed (Word draws an error in its language);
  a numeric picture's `#` with no digit is a space (`\# "#0"` draws " 2"), for every
  computed field.

- **Two faces the rasteriser files under one name are drawn apart** (round 10, item 4):
  Aptos Display and Aptos are both "Aptos" at 400 to resvg, so a page with headings in one
  and body text in the other drew both in whichever loaded first and warned
  `face-not-addressable`. `svg_to_png` now draws such a page in passes
  (`fonts.drawing_plan`, `png.pass_svg`): each face's text in a pass without the other
  family's files, over the passes before -- with Word's and the cloud cache's faces and
  with an application's folder holding both. docx-agent's render path, which calls
  `svg_to_png`, gets the same. Pages that need no pass are unchanged byte for byte.

- **A public API for laying out once and writing pages later** (round 10):
  `lay_out_docx(source, options)` returns `(layout, data, fonts)` -- every page laid out, the
  bytes, the faces it was measured with; picklable, so a layout can be kept or handed to
  another process -- `page_svgs(layout, data, options, fonts=..., device_size=...)` writes
  the selected pages byte for byte as `convert_docx_to_svg` does, and `user_font_dirs(options)`
  is the folders the rasteriser is handed after the faces the SVG names (docs/api.md).
  The private names docx-agent used (`_lay_out`, `_svgs`, `_user_dirs`) still work.
- **Rasterising loads only the fonts a page can use** (round 10, O2; ooxml-common 0.23.0,
  now the floor): `svg_to_png` hands resvg the files of its folders and list that the page
  can use, in resvg's load order (`ooxml_common.fonts.rasterdb`), instead of every file of
  every folder on every call. Same pixels: 845 page renders (fixtures, local documents,
  five font setups, the docx-agent render path) byte-identical to before. With system
  fonts on Linux (fontconfig's list) everything is handed over, as before.
- **EMF+ pictures are drawn from their EMF+ records** (round 7, L9 step 2; ooxml-common
  0.22.0, now the floor), as Windows PowerPoint draws them, no longer from a dual file's GDI
  fallback; an EMF+-only file is no longer a placeholder. A file whose EMF+ records draw
  nothing is drawn from its GDI records with a `metafile-records` warning. VRT snapshots
  unchanged.
- **EMF and WMF pictures are drawn** (round 7, L9; ooxml-common 0.21.0, now the floor):
  an inline or floating picture whose part is a metafile was a placeholder of its extent
  (`picture-not-drawn`); it is now drawn from its GDI records as SVG by ooxml-common's
  player, measured against Windows PowerPoint's drawing of the same records (Word's own
  drawing not measured). What the player leaves out -- EMF+ records among them, a dual
  file being drawn from its GDI records -- is a `metafile-records` warning; a metafile
  that draws nothing keeps the placeholder. `svg_to_png` writes these pictures inline
  (resvg draws no text in an SVG `<image>`). A picture bullet in EMF or WMF still stops
  the layout. VRT snapshots unchanged.
- **A chart's date categories are drawn as dates** (ooxml-common 0.17.0, now the floor): a
  category kept as a serial under a date code -- `[$-413]d\ mmmm\ yyyy`, as ooxml-edit's
  `date_format` writes it -- reads `5 januari 2026` where it read `46027`, as Word for Mac
  drew it (measured 16 October 2026). A code without a language prefix takes its names
  from `ConvertOptions.locale`'s language; a `/` is drawn as `/` (Word shows the viewer's
  region separator, which the locale does not hold). Word's automatic chart title is
  unchanged (already measured here, `chart_title_size`). Goldens unchanged.

- **The caller's locale: `ConvertOptions(locale=...)`.** It takes ooxml-common's shared
  `Locale` (0.15.0, now the floor): what Word takes from its machine rather than from the
  document. A decimal tab stop aligns on its decimal separator -- Word aligns on the
  system's (measured) -- and a chart's numbers are printed in its decimal and thousands
  separators (`ChartBuilder(locale=...)`). The default is `EN_US`, never the operating
  system's locale. **This changes the default output** of a decimal tab stop: it used to
  align on a hard-coded `,`, the measuring machine's separator, and now aligns on en-US's
  `.`, so `1234.5` after a decimal stop puts its `.` at the stop where it was drawn as
  after a right stop. No VRT snapshot holds a decimal stop, so none changed; a caller-built
  `Locale` with a `,` decimal draws byte for byte what was drawn before, and the probes
  measured in Word (`tools/read_render.py`, `tools/read_tab_leader_probe.py`) are scored
  in that locale (`render_record.MEASURING_LOCALE`). Charts' default output is unchanged.
- **An absent Symbol or Wingdings draws every glyph, from shared tables.** The recorded
  metrics moved to ooxml-common (`text.recorded_symbol_faces`, unchanged;
  `tools/record_symbol_faces.py` writes them there), and what an absent face's symbols are
  drawn as is now ooxml-common's `text.symbol_fonts`, built from the Unicode Consortium's
  tables, which pptx2svg uses too: every code whose glyph Unicode encodes -- Symbol's
  Greek and mathematical signs, all of Wingdings -- instead of the eleven bullet-library
  characters before. Wingdings' check (U+F0FC) stays U+2713, as Word and PowerPoint draw it.
  Requires ooxml-common with `text.symbol_fonts`.
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
