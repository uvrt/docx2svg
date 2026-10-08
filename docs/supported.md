# What docx2svg draws

Moved from the README. The measurements behind each line are in [`ROADMAP.md`](../ROADMAP.md).

**What it draws** is a transcription of a layout measured against Word's own PDF export,
glyph by glyph: text in its resolved face, size, weight, slant and colour, at the pen
position Word uses (every glyph of every committed document, and all 496,104 of a
175-page third-party one, at Word's position, baseline, face and size); underline,
strikethrough, highlight and shading; tab leaders; list labels; line numbers; paragraph borders and shading;
inline pictures; superscripts and subscripts; tables; text columns, balanced and ruled as Word
balances and rules them; footnotes and endnotes; headers and footers, with
`PAGE`, `NUMPAGES` and `SECTIONPAGES` computed from its own pages, `PAGEREF`, `REF` and
`SEQ` from its own pages and bookmarks, and every other field drawn as Word cached it; floating drawings, shapes and text boxes; charts, laid out as Word
lays them out, and SmartArt, from the drawing Word caches for it, inline and floating; the
content of content controls; and a document with tracked changes as Word's final (*No
Markup*) view draws it -- moved text at its destination, paragraphs joined across deleted
marks, deleted rows gone -- with comments taking no room. No markup view is drawn.
**What it does not draw yet** -- a chart's 3-D scene (drawn flat), a SmartArt diagram with
no cached drawing, a picture fill, a header's drawing the body wraps around, a floating table or a floating
drawing not positioned against its column in a section of several text columns, an
endnote in one in mode 15, and a footnote in one that a continuous break joins to another
section -- it says so: a
stable warning code, and where the layout cannot go on (a table's height is not
modelled), a marked band on the page where it stops and no page invented after it.

Faces are named in the SVG, not embedded. The layout reads the faces Word uses where
they are installed -- Office's cloud-font cache too, where Word 365 keeps Aptos Display --
(standard library only); a PNG is drawn with the same files.

## Status by area


| Area | State |
| --- | --- |
| Oracle (Word → PDF, measured with pypdfium2) | **Working**, and better than the sibling's — vector text, not pixels |
| Fixture corpus | One document, ours, generated, 2.6 kB |
| OPC package reader | Complete for what the fixture uses |
| Sections, page geometry | Read, and verified against the oracle |
| Paragraph/run properties, indents, tabs | Read, and verified against the oracle |
| Styles, `w:docDefaults` | `docDefaults` only; no inheritance cascade |
| Headers/footers, fields | **Drawn** (ROADMAP.md, "Headers, footers and fields — measured"): the story each page shows, where Word draws it; page numbers computed; other fields as cached |
| Footnotes, floating `w:drawing`, charts, SmartArt | **Drawn** (ROADMAP.md 4.13, "Floating drawings — measured", F.20, F.21): charts under Word's measured chart rules, SmartArt from the drawing Word caches. What is not drawn faithfully is reported as a warning, never dropped in silence |
| Text measurement | Vertical model measured (Phase 2). Advance widths from `ooxml-common`'s tables for the six faces checked against Word (`docx2svg.measure`), or from the installed faces |
| Line breaking | **Measured for Latin text** (ROADMAP.md, "Phase 3 — measured"): an exact, inclusive budget; every line of every probe and 7,164 / 7,167 real-document lines break where Word breaks them. Justified text in mode 15 compresses its spaces (3.8). **Automatic hyphenation** (`w:autoHyphenation`, 3.9) for English, Dutch, German and French: Word's zone, fit, fragment, limit, apostrophe and exclusion rules in each compatibility mode, its hyphen drawn where Word draws it, and in mode 15 no page ending in one (3.9.2); where a word breaks comes from the Moby Hyphenator list (public domain) for English and from Liang's algorithm over openly licensed patterns (`src/docx2svg/patterns/SOURCES.txt`), never from Word's proofing tools, and agrees with Word's break points for 88–95% of the 359 English probe words before the 42 it breaks otherwise are taken as observed exceptions (all of them after; only the first figure says anything of other words), 96% of the Dutch, 91% of the German and 86–88% of the French. East Asian breaking is not measured |
| Pagination | **Measured** (ROADMAP.md, "Phase 4 — measured"): every page top it reaches in the real documents, 210 / 218; it stops, rather than guess, at a table or a floating drawing |
| SVG and PNG output | **Phase 5** (ROADMAP.md, "Phase 5 — measured"): one SVG per page, a transcription of the layout; every glyph checked against Word's PDF; tables, floating drawings and columns reported and marked, not guessed |
| Tracked changes, content controls, comments | **Final view drawn** (ROADMAP.md, "Revisions — measured"): moves, deleted paragraph marks (joined under the next paragraph's properties; below mode 15 Word's lower-drawn lines too), content controls at every level, property changes, inserted and deleted rows and cells, as Word's *No Markup* view draws them, 70 / 70 probe cases in every setting; comments take no room |
| Text columns | **Measured** (ROADMAP.md, "Phase 4 — measured", 4.14): laid out column by column under the keeps, balanced above a continuous break, with their separator; every probe line Word's in every setting. Footnotes at each column's foot below mode 15, run on through the columns at the page's foot in mode 15 (4.15); drop caps, and endnotes below mode 15, in columns (4.16); floating drawings against their column (4.17) |

