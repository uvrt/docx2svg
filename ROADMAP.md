# Roadmap to rendering DOCX

Working plan for turning this repository from an oracle and a reader into a renderer.
Written to be picked up cold: every item says what is missing, how to approach it, **what
must be measured before it**, and how to know it is done.

Phases are ordered by *dependency and payoff*. Phase 0 comes first because everything
after it needs a way to tell "better" from "different" — the same reason it comes first in
the sibling [`pptx2svg`](https://github.com/uvrt/pptx2svg), whose roadmap says the oracle
"gates everything". That ordering is not a style choice. It is the only defence against
the failure mode that the entire DOCX rendering field exhibits, documented under *The
landscape* below: shipping a number that measures the whiteness of paper.

**Effort key:** S ≈ half a day · M ≈ 1–3 days · L ≈ 1–2 weeks · XL ≈ 3+ weeks.

**A note on what this file is for.** Constants carry the observations that produced them
and the residual that remains. Hypotheses that were refuted are recorded rather than
deleted — there are already two, and both were refuted by the first fixture. A constant
with no measurement behind it does not belong in this repository, and the fastest way to
notice one is to require that it be written down here with its evidence.

---

## Where DOCX is harder than PPTX, and where it is easier

This is the section to read first, because it is the reason this project is a sibling
rather than a feature of `pptx2svg`.

### Harder: layout is emergent, and errors propagate

A PPTX slide is a **set of absolutely positioned boxes**. Every shape carries an `a:off`
and an `a:ext`; text fills the box it was given. If the renderer gets one shape wrong,
that shape is wrong and the other eleven are fine. Errors are *local*.

A DOCX body is a **sequence of runs and a column**. Nothing in the file says where a line
breaks or where a page ends. Both are computed by Word from the text, the face's advance
widths, the column width, the tab stops, the paragraph's `keepNext`/`keepLines`/
`widowControl`, any floating object's wrap geometry, the space reserved for footnotes, and
a bundle of `w:compat` switches that differ by Word version. Get any of it wrong by one
character and the line breaks differently; once a line differs, every subsequent line on
the page differs, and once the page fills one line early or late, **every page after it is
wrong**. Errors are *global and compounding*.

That asymmetry sets the whole strategy. In `pptx2svg` you can render half the shapes and
score the result. Here, a renderer that is 99% right about line breaking is not 99% right
about the document — it is right until the first disagreement and then unrelated to the
oracle. So the measurements have to be built *upward* from quantities that cannot
propagate: page geometry, then indents and tabs (position, no flow), then the line advance
(vertical, one paragraph at a time), then line breaking, then pagination. Each phase below
is that ladder.

### Harder: there is no single "Word" to match

`w:compat` is a bundle of version-dependent layout switches, and Word honours them.
"Matching Word" is therefore ill-posed until you say *which* Word, with which compat
level, on which platform, with which fonts installed. LibreOffice implements these one bug
at a time; `docx-renderer` parses none of them. This project's answer is the oracle: the
target is Word 16.106 on this Mac, and that is stated rather than implied.

### Easier: the oracle is vector, and that is a large advantage

`pptx2svg` scores fidelity by rasterising both sides and computing SSIM plus a colour
histogram correlation. It has to: a slide is a picture, and "is this gradient right" is a
question about pixels.

**This project's oracle answers in numbers.** Word's PDF export carries text as vector
objects — verified, no raster anywhere — so `pypdfium2`'s `get_charbox` returns the ink box
of every glyph in points, to better than a thousandth of a point. And the hard problem here
*is* a question about numbers: where did the line break, and on which page did the
paragraph land. Both are read directly.

The measured consequence is already visible. `pptx2svg`'s best fixture scores SSIM 0.9897
— an excellent score that nonetheless cannot tell you *what* the remaining difference is.
This project's first fixture measured its indent ladder to **8e-6 pt**, which is float32
noise in pdfium's own coordinates, and would have caught an error of one twip (0.05 pt)
six thousand times over. A vector oracle does not merely score better; it says what is
wrong.

### Easier: one font model, no 3-D, no chart engine

`pptx2svg` carries 54 measured pattern presets, gradient laws, a tile-scale law, 187 preset
shape geometries, a chart renderer covering ten types, SmartArt, and EMF/WMF preview
extraction. A document body uses **none of that** until it contains a `w:drawing`, and
when it does, `w:drawing` wraps the same DrawingML `pptx2svg` already renders. There is no
3-D in a document body, no camera, no bevel. The text cascade is also simpler: a run's
face comes from `w:rFonts`' four script slots and the style chain, with no placeholder
inheritance from a layout and a master. (This paragraph originally also said "no theme
`+mn-lt`/`+mj-lt` indirection". **That was wrong**: `w:asciiTheme="minorHAnsi"` is the
same indirection, it is how real documents name their faces, and for East Asian and
complex-script slots it goes one step further, through `w:themeFontLang`. See *Style
inheritance — measured*.)

**Net:** the surface is smaller and the oracle is sharper, but the one thing that remains
is the thing nobody in the field has solved.

---

## The landscape — surveyed, and judged

Surveyed September 2026, before writing any layout code, and specifically to answer: what
is genuinely solved, and what does everyone approximate?

The rule this project takes from `pptx2svg`, which took it from `pptx-renderer`:
**method is worth borrowing, constants never are.** `pptx2svg` adopted SSIM and histogram
correlation as *gates* from a reference implementation and then measured every constant in
its own codebase itself. The same applies here, and the reference implementation makes the
point unusually well.

### `recon-vcs/docx-renderer`, reviewed at code level

TypeScript, browser-only, **renders to HTML + CSS in a live DOM** — not SVG, not canvas,
not PDF. A fork of `docx-preview-sync`, itself a fork of `docx-preview`/`docxjs`. ~10,800
LOC, Apache-2.0, 6 stars, single author, first commit 2026-07-08 and last push 2026-07-10.
A two-week-old solo project that is a dependency of one product, not an established
library — which is worth stating, because its README reads like an established library's.

**It paginates, and the mechanism is the interesting part.** Two stages: a static one that
builds layout regions from section properties and explicit breaks
(`modern-page-splitter.ts`, `page-builder.ts`), and a dynamic one that decides everything
else by **measuring the browser's own layout**. Every appended node triggers
`measurePageOverflow`, which is literally `clientHeight < scrollHeight`
(`overflow-measurer.ts:8-10`); on overflow the node is removed, a break recorded, and the
page split. The granularity is per token: `run-parser.ts:150-160` shreds every `w:t` into
characters for CJK ideographs and whitespace-delimited words otherwise. So pagination is a
**token-by-token DOM-reflow bisection**.

**It does not measure text. Chromium does.** No font metric tables, no fontkit, no
opentype.js, no FreeType. All advance widths, shaping, line breaking, justification and
hyphenation are Blink's, reached through CSS. Its one original contribution to metrics is
`gdi-line-height.ts`: it probes the resolved font stack on a canvas at 1024 px, reads
`fontBoundingBoxAscent`/`Descent`, and then reproduces what it believes is Word's GDI
rounding — **ceil ascent and descent separately, then sum** — writing the result back as an
explicit `line-height`.

**It claims fidelity to Word and measures it, and the measurement is close to
meaningless.** `README.md` states "Goal: 99% visual agreement with Microsoft Word";
`ACCURACY.md` reports 98.890% over a 5-page fixture against a Word PDF export. The metric
is `1 - mismatched/total` via pixelmatch after a **nearest-neighbour** downscale. The
research pass computed the actual ink coverage of its own reference pages:

| page | ink (non-white px) | reported score | mismatched | mismatch ÷ ink |
| --- | --- | --- | --- | --- |
| 1 | 0.53% | 99.486% | 0.514% | **97%** |
| 3 | 0.18% | 99.787% | 0.213% | **118%** |
| 5 | 15.25% | 97.101% | 2.899% | 19% |

On page 3 the number of differing pixels **exceeds the total ink on the page** — which is
what you see when every glyph is in the wrong place, since a shifted glyph counts twice.
A 99.8% score on a page that is 99.8% white paper carries almost no information.

**To its considerable credit, the author knows.** `docs/TEXT_ALIGNMENT_PLAN.md` opens with
「スコアだけ見ない」— don't just look at the score — and records that after fixing three
genuine bugs the score went *down* (99.531% → 99.486%) because some paragraphs had merely
been cancelling one bug against another. It also records a change that raised the score and
was **not committed** because there was no justification for the constant (+20 px on a
heading's margin). That is the right instinct, and it is rarer than it should be.

**Its real gaps are the ones that decide page breaks.** `keepNext`, `keepLines` and
`widowControl` are parsed and then discarded — `properties-parser.ts:268-270` reads
`case "keepLines": //TODO - maybe ignore`. Footnote space is not reserved before
pagination (`html-renderer-sync.ts:445`), so a page break can be chosen without the
footnote area Word reserved. Kerning is a TODO. `w:compat` is not parsed at all
(`settings.ts:26`: "TODO support more settings", five elements handled). And the README's
claim of golden-HTML comparison is false — `grep -rn golden` outside the README returns
nothing, the browser tests assert only that `pageErrors == []`, and the accuracy workflow
is wired to `branches: [main-tmp]`, so the headline number is not regenerated on `main`.

### What to borrow from it, and what not to

**Borrow — method:**

1. **Generate the fixture; ship the reference; ship the diff script.** It does all three.
   Almost nobody else does, vendors included. This repository does the same, with the
   difference that its fixture is regenerable from a committed script and its reference is
   measured rather than rasterised.
2. **Name the residual, and refuse a constant that has no derivation.** Its
   `TEXT_ALIGNMENT_PLAN.md` is a better account of the difficulty than anything any
   commercial vendor publishes.
3. **The GDI-rounding *question* is the right question.** The idea that Word's line advance
   is a rounded quantity, not a real number, is correct — see the measurement below. It
   arrived at the question honestly and it is the most sophisticated open attempt at it.

**Do not borrow — constants and architecture:**

1. **Its GDI model is refuted here.** `ceil(ascent_px) + ceil(descent_px)` for 11 pt Calibri
   at 300 dpi gives **57 px**. Word 16.106's PDF export uses **55.949 px**, measured over 70
   baselines. See *The line advance* below. The method (rounding happens) survives; the
   constant does not, which is exactly the rule.
2. **Delegating measurement to a layout engine caps you at that engine.** No amount of
   line-height correction fixes a paragraph Chromium wrapped one word differently from
   Word, and once a line differs the page does. This is the ceiling its own planning docs
   describe, and it is structural.
3. **A whole-page pixel ratio is not a fidelity metric for documents.** It scores paper.
   `pptx2svg` already rejected raw pixel difference for a related reason and moved to SSIM;
   this project goes further and does not rasterise at all.

### The rest of the field, one line each

| Project | Lays out itself? | Fidelity claim |
| --- | --- | --- |
| **docx4j** (Java) | No — XSL-FO → Apache FOP | None; ships `documents4j` (drives real Word) and MS Graph paths *because* the FO path is limited |
| **LibreOffice headless** | **Yes** — the only mature independent open-source Writer engine | Good on simple documents, degrades on complex; chases Word compat flags one bug at a time |
| **Pandoc** | No — discards page geometry entirely, by design | None; a maintainer says Word's page layout "is not really feasible" for pandoc |
| **docx-preview / docxjs** | No — browser | None stated |
| **mammoth** | No — semantic only (`Heading 1` → `<h1>`) | Deliberately none |
| **python-docx** | **Renders nothing** | n/a |
| **docx2pdf** | No — automates real Word, LibreOffice fallback | 100% by construction, where Word is present |
| **Aspose.Words** | **Yes**, from scratch, closed | "pixel-perfect"; no independent benchmark substantiates it |
| **Syncfusion DocIO** | Yes, closed | Notably *honest* about gaps (table auto-resize can shift page numbers) |
| **Spire.Doc** | Yes, closed | Shipped a "new engine-style layout" to fix page fidelity — a tacit admission |
| **ONLYOFFICE** | **Yes** — own C++ core, OOXML-native | Claimed better than LibreOffice; the claims are blog posts, not benchmarks |
| **SuperDoc** | **Yes** — a bespoke `layout-engine` paginating FlowBlocks, not CSS | Young, no published benchmark. The most interesting OSS case |
| **MS Graph `?format=pdf`** | No — real Word server-side | Full, by definition; OneDrive/SharePoint-hosted files only |
| **unoconv** | No — LibreOffice | Canonical repo archived 2023; `unoserver` is the live equivalent |

**DOCX → SVG specifically is essentially empty.** The only open-source hit is
`ryusui-hiro/document-svg`, created 2026-09-08, one star. Everything else advertising it is
a SaaS converter almost certainly vectorising an Aspose- or LibreOffice-produced PDF. There
is no mature purpose-built DOCX→SVG layout engine, which is both the opportunity and the
warning.

### Solved versus approximated

**Genuinely solved — correct and verifiable:**

1. **The container and the XML.** OPC, relationships, content types, the part graph.
   Deterministic, specified in ECMA-376, and everyone gets it right.
2. **Structural extraction.** Paragraph/run/table/style trees. Mammoth and Pandoc are
   *right* precisely because they stopped there.
3. **Static, non-reflowing geometry.** Page size, margins, section properties, explicit
   `w:br type="page"`, section-break parity. Arithmetic on twips.
4. **Fidelity by delegation.** Driving real Word is 100% correct by construction. It is
   the only approach with a proof — and it is what this project's oracle *is*.

**Everyone approximates, no exceptions:**

1. **Line breaking.** Nobody open-source reimplements Word's breaker. There is no
   verifiable Word-matching line breaker in the open, at all.
2. **Vertical metrics.** An undocumented algorithm, and Word is not one target — GDI and
   DirectWrite disagree.
3. **Where the page breaks.** Most implementations do not even implement the properties
   that decide it.
4. **`w:compat`.** Makes "matching Word" ill-posed until you say which Word.
5. **Floating objects and wrap.** CSS floats push text to one side; Word's wrap model is a
   different geometry.
6. **Justification, kerning, East Asian typography.** Delegated or dropped everywhere.
7. **Measuring fidelity at all.** There is **no published, reproducible pixel-diff
   benchmark of DOCX converters against real Word output.** This is the finding worth
   emphasising most, and it is the gap this repository's Phase 0 is aimed at.

**The skeptical conclusion: DOCX parsing is solved; DOCX layout is not, and cannot be
verified without Word in the loop.** Every project claiming otherwise is (a) delegating to
Word or LibreOffice, (b) closed and unmeasured, or (c) measuring itself with a metric that
mostly scores the whiteness of paper.

---

## Phase 0 — Word as the fidelity oracle — **done**

**Effort: M. Blocks everything else.** Done, and it is the reason anything below is
writable.

Microsoft Word 16.106 is installed on this machine and is scriptable, which makes the
reference implementation itself available as ground truth — strictly better than
LibreOffice, which is an approximation of Word and is not installed here anyway.

### 0.1 The oracle — verified working

`tools/word_export_pdf.applescript`. Three constraints were found the hard way, all three
of which cost the sibling project real time in their PowerPoint form; they are documented
in the script's header rather than here, so they travel with the code:

| Constraint | Consequence |
| --- | --- |
| **`open` does not return a document reference** | PowerPoint's dictionary has `open` yield the presentation; Word's does not. `set d to active document` *after* opening. The failure is a type error naming neither `open` nor the document, and reads as a broken input file. |
| **An AppleEvent timeout (−1712) is the default failure** | AppleScript's own deadline is two minutes and Word exceeds it on a first launch, a long document, or any dialog. Without `with timeout of N seconds` the script fails while Word is working normally. |
| **A `~$` lock file wedges the next attempt** | A run that died leaves `~$<name>.docx` behind; the next `open` either silently opens a read-only copy (whose `save as` writes nothing while the script exits 0) or raises a blocking dialog. Both look like a broken document. `pkill -x "Microsoft Word"` plus deleting the lock clears it. |

A fourth was found in Phase 2, and it is the one that did **not** carry over: **Word is
sandboxed, and its file grants are per app.** PowerPoint had been granted
`~/pptx2svg-oracle/` long ago, so the sibling never saw a prompt; Word had been granted
nothing, and every new probe file in `~/docx2svg-oracle/` raised a file-access dialog.
From the script's side a pending dialog is a hang — `open` blocks, −1712 after the full
180 s, no window visible to System Events, no `~$` lock — so it read as Word failing on
documents that had exported a minute earlier. The oracle now lives in the **Office group
container**, `~/Library/Group Containers/UBF8T346G9.Office/docx2svg-oracle/`, which every
Office app may read and write without a grant: verified silent (4 s through Word's own
`open` event), then nine probe exports in 69 s with no prompt. It needs no manual step on
any machine. `ORACLE_DIR` in `tools/oracle.py` is the one place the path is written; it
contains a space and is only ever passed as a single argv element.

Carried over unchanged from the PowerPoint oracle, and still biting: **always check the
PDF exists rather than trusting the exit status**, because several failure modes return
success and write nothing.

A fifth was found with tracked changes (*Revisions — measured*, R.0): **Word exports what
its window shows on opening, and for a document with revisions or comments that is *All
Markup*** — a markup pane of balloons beside the text (a deletion as "heeft verwijderd:
…", in the interface's language), the text moved aside to make room (73 px left, 363 px
down on an A4 page at 300 dpi), the glyphs at their size. Neither `w:revisionView
w:markup="0"` in `settings.xml` nor the document's `print revisions` changes the export.
The renderer draws only the final view, so `tools/oracle.py` exports every document that
holds a revision or a comment in Word's *No Markup* view (the script's `final`: the
window's `show revisions and comments` off and its `revisions view` final, a setting the
window does not outlive); a document with neither draws identically either way, so its
export keeps the name and cache entry it always had.

### 0.2 Why this oracle is better than the sibling's, measured

Word's PDF carries text as **vector** objects. `pypdfium2`'s `get_charbox(i)` returns the
ink box of every glyph in points. The `/BaseFont` entries name the embedded subsets, so
**Word's own font substitution is visible in the export**.

That last point is not incidental. The first probe document — `hello.docx`, naming Calibri
but carrying no `w:docDefaults` — came back with `AAAAAC+Calibri`, `AAAAAE+Aptos` **and**
`AAAAAG+Calibri`. Word resolved the theme's minor font for the runs that named no face, and
that substitution would have silently contaminated every measurement taken from it. The
sibling project's most expensive lesson was exactly this — five of its nine fixtures could
not be scored at all because the oracle and the renderer disagreed about which *face* to
draw, so any number measured font resolution rather than layout. So this project's first
fixture names its face in all four `w:rFonts` slots **and** in `w:docDefaults`, and
`tests/test_oracle.py::test_word_drew_the_face_the_document_asked_for` asserts the exported
`/BaseFont` set is exactly `{Calibri, Calibri-Bold}`. That test exists on day one rather
than being retrofitted after five fixtures were wasted.

**Corrected in the style-inheritance work:** the Aptos in `hello.docx` was not a run. It
was the *paragraph marks* -- Word draws each mark as a space glyph in its own text object
-- and those marks had no `w:rPr`, in a package with no styles part. Re-exported on
2026-09-24 from the same bytes, the marks come back **Calibri 12 pt**, not Aptos. So what
an unstyled mark resolves to depends on Word's state, not on the document (this
machine's `Normal.dotm` has an Aptos theme and 12 pt defaults; Word used the size and
not the theme). `docx2svg.resolve` reports such values with an `application default`
origin instead of guessing silently.

### 0.3 The fixture — ours, generated, and accepted by Word

`tools/make_layout_sweep.py` writes `tests/fixtures/layout-sweep.docx` (2.6 kB, five
parts, no third-party content). **The acceptance test for generated OOXML is that Word
opens and exports it** — well-formed is not the bar, because Word silently *repairs* a
document with wrong child order or a malformed content-type override, and a repaired
document exports under a different name or not at all.

The design principle, which every later fixture should inherit: **measure differences, not
positions.** Each block is a set of paragraphs differing in exactly one authored quantity,
all beginning with the same probe glyph. Subtracting one measurement from another cancels
the glyph's left side bearing, the face's vertical metrics, and Word's page-box rounding —
none of which this project may ship a font file to model. What survives the subtraction is
a number authored in twips in our own file that has to come back out of Word's PDF in
points.

### 0.4 What the loop measured

`tools/read_layout_sweep.py`, and `tests/test_oracle.py` as the gate.

| Quantity | Authored | Measured | Residual |
| --- | --- | --- | --- |
| Indent ladder, 5 steps | 0 / 18 / 36 / 54 / 72 pt | same | **8e-6 pt** |
| Tab stops, 3 | 72 / 144 / 216 pt | same | **8e-6 pt** |
| Baselines on the 1/300 in grid | — | 70 of 70 | < 0.01 px |
| Page break | — | after L32 of L70 | observation, pinned |

8e-6 pt is float32 noise in pdfium's coordinates on a 600 pt page; a twip is 0.05 pt, six
thousand times larger. These are exact.

### 0.5 The page box is not the page size — **measured, and a falsification test passed**

A4's `w:pgSz w:w="11906"` is 595.3 pt. Word's `/MediaBox` says **595.2**. The height is
authored 841.9 and exported **841.92**. Neither is truncation and neither is rounding to a
decimal place; the two residuals have opposite signs.

Both exported extents are **whole 1/300-inch device pixels**: 595.2 pt = 2480.0 px and
841.92 pt = 3508.0 px, which are exactly A4 at 300 dpi. The unrounded values are 2480.4167
and 3507.9167, so the residuals are −0.4167 px (−0.10 pt) and +0.0833 px (+0.02 pt) —
matching the observed −0.1 and +0.02 exactly.

**This was then tested against a case designed to kill it.** Section 2 of the fixture asks
for 10000 × 13000 twips, which is *exactly* 500 × 650 pt. A page box that were simply the
authored size would read 500 × 650; the device-pixel law predicts 2083 and 2708 px, that is
**499.92 × 649.92**. Word exported 499.92 × 649.92. The law survived.

Implemented as `docx2svg.units.device_page_extent_pt`, gated by `tests/test_units.py`.

**Why this matters more than 0.1 pt.** US Letter is a whole number of device pixels in both
directions and is unchanged by this rule — so a corpus of Letter documents would never have
revealed it, and every A4 comparison would have been wrong by up to 0.12 pt in a direction
that depends on the page size. That is precisely the kind of error that gets attributed to
something else, and it would have been attributed to text measurement.

**Not settled:** whether the 300 dpi is fixed or comes from a printer/display setting;
whether the rounding is round-half-even or round-half-away; and whether the same grid
applies on Windows Word. Three measurements each, on documents this project can generate.

### 0.6 The line advance — a ratio measured here, the accumulator reproduced in Phase 2

Seventy single-line paragraphs of 11 pt Calibri. Every baseline sits on a whole 1/300-inch
device pixel — the same grid as the page box, which is evidence of one mechanism rather
than two. The gaps between consecutive baselines are **56 px** with three exceptions of
**55 px**, roughly one in twenty.

That average is 55.949 px. Calibri's `hhea` gives ascender 1950, descender −550, lineGap 0,
unitsPerEm 2048, so `(1950 + 550 + 0) / 2048 = 2500/2048 = 1.220703125`, and at 11 pt that
is 13.427734375 pt = **55.948893 px**. The 55/56 alternation is *pure accumulated
rounding*: over 20 lines the exact total is 1118.98 px, which rounds to 1119 = 19 × 56 +
1 × 55. Exactly what is observed.

**Refuted: the GDI `ceil(ascent) + ceil(descent)` model.** `docx-renderer`'s
`gdi-line-height.ts` is built on it. At 11 pt Calibri and 300 dpi it gives **57 px**. The
measurement is 55.949. The *method* — that the advance is a rounded quantity rather than a
real number — is right and is worth keeping. The constant is wrong. (It may well be right
for the screen path at 96 dpi, which is what that project targets; what is refuted is
applying it to Word's PDF export at 300 dpi.)

**Also not yet settled, and this is the honest part.** Modelling the baselines as
`round(origin + n × 55.948893)` reproduces **61 of 70** and is off by exactly 1 px on the
other 9 — it drifts +1 at n = 10, comes back at the observed 55 px gap at n = 17, and
drifts again at n = 30. So the *ratio* is established and the *accumulator* is not: Word is
not accumulating a real number and rounding the total. Candidates: accumulation in a finer
fixed-point unit; rounding of the advance per line with a carried remainder; or a
per-paragraph origin that is itself snapped before the lines are laid out. **Measure before
choosing:** a probe document with one font size per page over 6–24 pt, and with
`w:line`/`w:lineRule` set to `exact`, `atLeast` and `auto` separately, will distinguish
them. That probe is Phase 2's first deliverable.

**Resolved in Phase 2** (see *Phase 2 — measured* below): all 70 baselines, and all 91
lines of the fixture, are reproduced to 0 px by `docx2svg.vertical`. (What Phase 2 left
at 1 px -- `auto` multiples, bordered lines, lines after paragraph spacing -- is closed by
*The line box — measured*.) The nine misses
were not an accumulator error at all — line tops *do* accumulate the exact ratio — but a
rounding that happens against the line's top edge or its bottom edge depending on the
fractional pixel of the pitch. The candidates listed above (a finer fixed-point unit, a
carried remainder, a snapped per-paragraph origin) were each tested and none is needed.

### 0.7 Two artefacts of our own measurement, recorded because they cost time

- **A baseline estimated from a whole line is not stable.** The first attempt took the
  modal ink-box bottom of every glyph on a line, and produced a *four-valued* line advance
  (13.20 / 13.35 / 13.44 / 13.53 pt) that read as a Word behaviour. It is not: round
  glyphs overshoot the baseline, so the modal bottom flips depending on whether a label's
  digits happen to be round. Measuring from **one glyph** that appears on every line gives
  two values, and they are the real ones. This is the fixture's own
  measure-differences principle applied to the vertical axis, and forgetting it cost an
  hour and nearly produced a fabricated constant.
- **A prefix match is not a matcher.** A test counting the pagination block by
  `text.startswith("L")` also counted the wrapping paragraph, which begins "Line
  breaking…". Match the label's shape.

### 0.8 What Phase 0 still owes

- **A regression net that runs without Word (S).** The sibling commits one SVG per slide
  and compares byte for byte. There is no output to snapshot here yet; when there is, the
  same arrangement applies, and `.gitattributes` already pins `eol=lf` reasoning.
- ~~**A `tools/oracle.py` that caches by input SHA-256 (S).**~~ **Done in Phase 2.**
  `oracle.export(docx_bytes, name=...)` writes `<name>-<sha16>.pdf` and re-exports only
  when the bytes change. `conftest.py` still reuses `layout-sweep.pdf` by name.
- ~~**The recovery helper (S).**~~ **Done in Phase 2.** `oracle.recover()` quits Word,
  *waits for it to be gone* (relaunching while the old Word was still quitting failed
  with exit 1 and nothing on stdout), and sweeps every `~$` lock. `conftest.py` and
  `read_layout_sweep.py` both call it.
- ~~**One Word, shared with docx-agent (S).**~~ **Done (2026-10-03).** docx-agent drives
  the same Word, and each project's recovery would quit Word under the other's export. So
  every launch, export, re-save and recovery here holds an advisory `flock` on
  `~/Library/Group Containers/UBF8T346G9.Office/word-oracle.lock`, the file docx-agent's
  oracle locks (`oracle.word()`, re-entrant; `oracle.run()` is the one place a script is
  run, used by `export`, `resave`, `conftest.py`'s `oracle_pdf` and
  `read_layout_sweep.py --export`). **Blocking with a timeout**, not docx-agent's skip:
  a probe batch waits out docx-agent's oracle rather than fail, polling once a second for
  up to `DOCX2SVG_WORD_LOCK_TIMEOUT` seconds (600), then `WordBusy`, which `oracle_pdf`
  turns into a skip. Whoever holds the lock quits the Word it started before releasing
  it, so **a Word running once the lock is held is someone else's** (after ten seconds'
  grace for one quitting): never quit, `WordBusy` instead. Under the lock `recover()`
  quits Word (ours: it was not running when the lock was taken); outside it, as the
  probe scripts' `finally` calls it, it takes the lock, leaves Word alone and deletes only
  the `~$` files in `docx2svg-oracle/`. A cached export takes no lock. Checked with a real
  export while another process held the lock: it waited for the release, exported, and
  left Word quit and no `~$` file (`tests/test_word_lock.py` holds the rules without Word).

---

## Phase 1 — The `pptx2svg` boundary: depend, vendor, extract, or reimplement — **decided, and executed: extracted into `ooxml-common`**

**Effort: S to decide, M–L to execute. Decide before Phase 2 writes any measurement code.**

### Executed — the decision, what moved, and what is left for Phase 3

**The decision.** Extract now, before Phase 3, rather than at the trigger recommended
below. That was the user's call and the reasoning below is kept as the record of what it
overrode: the risk it names -- freezing a text interface before either project has broken
a line against the other's -- is avoided by moving only what needed no interface design.

**The package is `ooxml-common`** (import name `ooxml_common`), a new repository beside
this one, built by `git filter-repo` from `pptx2svg` so that every moved file keeps its
history back to the sibling's first commit. Standard library only, like both consumers.

**What moved** -- the certainly-in list below, exactly, plus the three DrawingML data
modules that needed no value types:

| From `pptx2svg` | To | Lines |
| --- | --- | --- |
| `opc.py`, `xmlutil.py`, `units.py` | `ooxml_common.{opc,xmlutil,units}` | 421 |
| `fonts/` (6 modules) | `ooxml_common.fonts` | 2,455 (moved 2,539, less the 87-line deck half of `check.py`, plus a 3-line note) |
| `text/{kerning,metrics,fontmap,measure}.py` | `ooxml_common.text` | 10,636 |
| `guides.py`, `render/{pattern,preset_specs}.py` | `ooxml_common.drawingml` | 4,312 |

`fonts/check.py` split on the way: the report and `check_families` take family names and
moved; `check_deck`, `deck_families` and `resolved_families` walk the slide model and
stayed in `pptx2svg`. `pptx2svg` keeps every old import path as a re-export of the same
module object, and its suite, VRT snapshots and fidelity baselines are unchanged.

**What did not, and who decides it:**

* **`text/wrap.py` -- Phase 3.** Still the open question this phase named: the method
  transfers, the types (`m.Paragraph`, `m.RunProperties`, `m.TextRun`) do not. Phase 3
  breaks lines against Word, and what it needs from a paragraph is the protocol a shared
  line breaker takes. Until then `ooxml-common` stops below that line.
  *Answered by Phase 3:* a list of units that carry their own break rules, with the
  loop shared and the rules kept by each consumer -- proposed, with the evidence, in
  *Phase 3 — measured*, 3.7; not moved yet.
* **`render/fill.py`, `render/geometry.py` -- when DrawingML is drawn here (Phase 6).**
  They take `model.py`'s DrawingML value types; lifting those out of the slide model is
  the step that lets both follow, and it was not mechanical enough to do in passing.
* **`render/text.py`** stays PPTX-shaped (`a:bodyPr`), and **the fidelity harness** stays
  out, both as recommended below.
* **`pptx2svg-fonts`** keeps its name and its home in `pptx2svg`'s repository. It is
  found by import name only, so its location changes nothing for this project, and the
  moment to rename it, if ever, is before its first PyPI release.

**This project** imported nothing from `pptx2svg` before the extraction and imports
nothing from `ooxml-common` yet: `src/` is still standard-library only and
`dependencies = []` still says so. `pyproject.toml` now names `ooxml-common` as the
package to install beside it, and records the dependency line and CI step to add when
the first import lands -- Phase 2's advance widths or Phase 3's line breaker, whichever
comes first. Its duplicates of the shared modules -- `docx2svg.opc`, `xmlutil`, `units`
-- stay until then too: each carries Word-specific constants, and folding them into the
shared ones is a change to be made with the suite watching, not as part of a move.


This is the architecture question, and it is assessed here rather than assumed.

### What is actually shared, measured from the import graph

`pptx2svg` is 41,031 lines across `src/pptx2svg/`. Its modules divide cleanly, and the
division is **already visible in the imports** rather than needing to be designed:

**Zero coupling to the slide model — 18,295 lines (45%), importable today as-is:**

| Module | Lines | What it is |
| --- | --- | --- |
| `text/kerning.py` | 6,859 | Kerning class matrix measured from GPOS |
| `text/metrics.py` | 2,468 | Generated advance-width tables, 20 face-and-weight combinations |
| `fonts/` (6 modules) | 2,539 | Embedded-font decoding, SFNT reading, `fsType` policy, the bundle probe |
| `render/preset_specs.py` | 3,860 | 187 preset shape geometries |
| `text/fontmap.py` | 761 | Clone substitution with its grading system |
| `text/measure.py` | 548 | The measurer protocol and both implementations |
| `opc.py`, `xmlutil.py`, `units.py`, `guides.py`, `imagemeta.py`, `png.py`, `render/pattern.py` | 1,260 | Container, XML, units, geometry guides, the 54 pattern presets |

**Coupled to `model.py`, but only to its DrawingML half:** `render/fill.py` (442) touches
`m.Fill`, `m.GradientFill`, `m.PatternFill`, `m.Outline`, `m.TileInfo` — *value types*, not
slide structure. `render/geometry.py` (673) touches `m.PresetGeometry`,
`m.CustomGeometry`. DOCX gets these verbatim through `w:drawing`. Shareable after lifting
the DrawingML value types out of `model.py`.

**Genuinely PPTX-shaped, do not share:** `render/text.py` (1,630) touches
`m.BodyProperties`, `m.TextBody`, `m.Transform` — an `a:bodyPr` text box, which a document
body does not have. `resolve/view.py` (1,888), `resolve/chart.py` (8,998),
`parse/source.py` (725), `render/svg.py` (334).

**The interesting boundary case is `text/wrap.py` (473).** It imports `model` but touches
only `m.Paragraph`, `m.RunProperties`, `m.TextRun`. Those are DrawingML's spellings of
concepts DOCX has under `w:p`, `w:rPr`, `w:r` — the same concepts, different types. So the
*method* is shareable and the *types* are not. This is where a shared package has to decide
whether it defines its own paragraph protocol or stays below that line.

### The four options, and what each costs

**Reimplement.** Rejected. `text/kerning.py` alone is 6,859 lines of GPOS-derived kerning
classes, and `text/metrics.py` is 2,468 lines of measured advance widths for 20
face-and-weight combinations. Re-deriving those would take weeks and would produce a
*second* set of constants that could drift from the first — two numbers for the same
physical fact is strictly worse than one. And the measurements are the part of `pptx2svg`
with the most work behind them.

**Vendor (copy files in).** Rejected for the same reason at one remove: a copy is a fork,
and a constant fixed in one copy stays wrong in the other. The sibling's whole claim is
that every constant records its observations; duplicating them breaks that claim quietly.

**Depend on `pptx2svg` directly, as an editable path install.** Works today.
`pptx2svg` is not on PyPI, so this means `pip install -e ../pptx2svg`, which is exactly the
arrangement `pptx-agent` already uses (`render = ["pptx2svg>=0.1"]`, resolved from a path).
The cost is conceptual rather than technical: `docx2svg` would depend on a *PowerPoint*
renderer to measure text, which is an odd shape for a published package and would pull
`resolve/chart.py`'s 8,998 lines into the dependency closure for nothing.

**Extract a shared package (`ooxml-common`, or similar).** The tidy answer, and the one the
user leans towards. The import graph says it is buildable: 45% of `pptx2svg` already
imports nothing from `model.py`.

### Recommendation: depend now, extract at a named trigger — and the trigger is Phase 3

**Both halves of this, with the reasoning.**

*What belongs in it*, empirically rather than tidily:

- **Certainly in:** `opc.py`, `xmlutil.py`, `units.py`, `fonts/*`, `text/kerning.py`,
  `text/metrics.py`, `text/fontmap.py`, `text/measure.py`. These import nothing from the
  slide model and DOCX needs every one of them.
- **In, after lifting the DrawingML value types out of `model.py`:** `guides.py`,
  `render/preset_specs.py`, `render/pattern.py`, `render/geometry.py`, `render/fill.py`.
  DOCX needs these only once it renders `w:drawing`, which is Phase 6 — so this half can
  wait, and waiting is free.
- **The `pptx2svg-fonts` bundle:** in, unchanged. It is already a separate distribution
  and is font files under OFL, which both projects need identically for rasterisation.
- **Out:** the slide model, `resolve/*`, `render/text.py`, `render/svg.py`, and everything
  DOCX-flow-shaped on this side. The fidelity harness is **out too, and that is a
  correction to the obvious answer**: `pptx2svg` scores by SSIM on rasterised pages and
  this project measures glyph boxes in a vector PDF. They share the idea of an oracle and
  none of the code. Putting `tools/fidelity.py` in a shared package would be tidiness
  mistaken for reuse.
- **Genuinely undecided:** `text/wrap.py`. It is the piece DOCX most wants and the piece
  whose types least transfer. Extracting it requires inventing a paragraph protocol that
  neither project has today, and inventing it now would be designing a boundary from a
  guess — which is the thing this recommendation is trying to avoid.

*When*, which is the question that matters:

**Not yet.** Start `docx2svg` against `pptx2svg` as an editable path dependency, record
every module it actually imports, and extract when the list stops changing. The case:

- **The boundary is 90% obvious and 10% load-bearing.** The 18,295 coupling-free lines are
  not in doubt. `text/wrap.py` is, and it is the single most important module in the
  overlap for a project whose hard problem is line breaking. Extracting now means fixing
  its interface before this project has broken a single line of text against the oracle.
- **`pptx2svg` has 2,522 passing tests and measured constants throughout.** Restructuring
  it on speculation risks a codebase whose entire claim is that nothing in it is guessed.
  The asymmetry is stark: extracting early risks the sibling's credibility, extracting late
  costs a rename.
- **The git-history argument is real but smaller than it looks.** `git log --follow` and
  `git blame` both traverse renames, and a modern extraction with `git filter-repo`
  preserves the history of the moved files. What is genuinely lost is the *interleaving* —
  the ability to see that a kerning constant changed in the same commit as a wrap fix. That
  is worth something. It is not worth freezing an interface before it is understood.
- **Three editable installs is a real cost** (`pptx-agent` → `pptx2svg` → shared, plus
  `docx2svg` → shared), and it is a cost that is *incurred either way* if the extraction is
  right; deferring it does not avoid it.

**The trigger, stated now so it is not a judgement call later:** extract when this project
reaches Phase 3 (line breaking) and needs `text/wrap.py`'s method. At that point the import
list is known, the paragraph-protocol question has an answer driven by a real second
consumer, and the extraction is a description of observed fact rather than a design. Until
then `docx2svg` imports nothing from `pptx2svg` at all — which is why `src/` is
standard-library-only today and why `pyproject.toml` records the path-install arrangement
in a comment rather than as a dependency.

**If the reading is wrong, this is how it will show:** if Phase 2 finds itself importing
`text/metrics.py` and `text/measure.py` and nothing else, and nothing about the paragraph
protocol comes up, then the boundary *was* obvious and waiting cost a week. That is the
downside, and it is bounded. The opposite mistake — a shared package whose text interface
was designed before either consumer wrapped a line — is not bounded.

**Do not restructure `pptx2svg` while this is undecided.** This phase exists so the
decision is made with the evidence in front of it.

---

## Phase 2 — Text measurement and the vertical model (M–L)

**Blocks Phase 3, which blocks everything.** Nothing here draws.

The goal is a function from (face, size, string) to an advance width, and from (face, size,
`w:spacing`) to a line advance, both agreeing with Word to the device pixel.

**Must be measured before starting** — all three now measured; results, residuals and
refutations are under *Phase 2 — measured* below:

1. **The accumulator** (Phase 0.6). One probe document, one font size per page, 6–24 pt,
   `w:lineRule` in `exact` / `atLeast` / `auto` separately. This decides whether a line
   position is `round(origin + n·a)`, a carried-remainder accumulation, or something else.
   Nine baselines currently disagree with the simplest model and until that is resolved no
   pagination result is trustworthy.
2. **Whether the horizontal axis is quantised too.** Every vertical position measured so
   far is a whole 1/300-inch device pixel. The indent ladder's horizontal positions came
   back at 8e-6 pt from their authored values — 72.0000, 90.0000 — which *are* whole device
   pixels, so the fixture cannot yet distinguish "Word snaps x as well" from "these indents
   happened to be on the grid". **A probe with indents at 7, 13 and 37 twips settles it in
   one export**, and the answer changes how every advance width must be rounded. Do this
   first; it is an hour.
3. **Whether `w:rPr/w:spacing` (tracking) and kerning are applied before or after that
   rounding.** The sibling measured that PowerPoint applies the OpenType `kern` feature and
   that not modelling it cost up to 1.5% of a line's width. Word's answer must be measured,
   not inherited.

**Deliverables:** a measurer with the same protocol shape as `pptx2svg`'s
`TextMeasurer`, backed by its metric tables (see Phase 1); a vertical model with its
constants and their residuals recorded here; and a probe document plus reader for each, in
`tools/`, following the `make_*` / `read_*` pairing the sibling uses.

**Done when:** the predicted baseline of every line in a generated probe matches the
oracle's to 0 device pixels, across at least five font sizes and all three line rules.
Not "within a pixel" — the quantity is an integer and the model should produce integers.

### Phase 2 — measured

Four probes, each a `make_*` / `read_*` pair in `tools/`, all exported through
`tools/oracle.py`. A new instrument made them possible: **`tools/quartz_pdf.py` reads the
PDF content stream rather than pdfium's ink boxes.** Word's export goes through Quartz,
which draws the whole page under `0.24 0 0 0.24 cm` — that is, in **1/300-inch device
pixels** — and starts a new text object (with its pen position printed to 1e-4 px)
wherever Word positions a glyph on its own: a new run, a tab, the paragraph mark. So
baselines come out as integers, line widths come out exactly (the paragraph mark's pen
x), and giving every glyph its own run in alternating colours gives every glyph an exact
pen position. Inside a single `TJ`, positions carry Quartz's 1/1000-em quantisation
(≤ 0.023 px at 11 pt) and are used only where that is irrelevant.

Also read off the stream, and new: **the ink is drawn at the size rounded to whole device
pixels while the pen advances at the exact size.** 11 pt is 45.83 px and Quartz's `Tm`
scale is 46; 6.5 / 10.5 / 11.5 pt draw at 27 / 44 / 48. Glyph *outlines* in Word's PDF
are therefore up to 1% larger than their advances imply. Irrelevant to layout; relevant
to Phase 5, which must decide whether an SVG reproduces the ink or the advances. *Settled in
Phase 5: the SVG places every glyph from the advances and scales its outline to the
device-rounded size, Word's ink -- the raster cannot separate the two more finely than
pdfium and resvg differ; see "Phase 5 — measured", 5.4.*

#### 2.1 The horizontal axis is not quantised to the device grid — it is quantised to 1/4096 pt

`make_x_grid_probe.py` / `read_x_grid_probe.py`. Indents (`w:ind/@w:left` 0–24 twips in
1-twip steps, then 37, 53, 97, 113, 131, 1447; `@w:firstLine` 7, 13, 37) and tab stops
(1440 + 0–12 twips, and 1477) in Calibri 11 pt; glyph advances of `H` and `i` at 6.5 /
10.5 / 11.5 pt in Calibri, Arial and Times New Roman, once as one run and once as one run
per glyph.

| Candidate grid for a line start / tab stop | Worst residual, 48 positions |
| --- | --- |
| whole device pixel (1/300 in) | **0.500 px** — refuted: a 1-twip indent moves the line by 0.2086 px |
| the authored twip value, unrounded | 0.00043 px |
| the twip value rounded to **1/4096 pt** | **0.00006 px** (the print precision is 0.00005) |
| 1/8192 pt, 1/2048 pt, 1/1024 px, 1/1000 px, 1/1200 px, 1/65536 in | 0.00043 – 0.00254 px |

The deviations from the unrounded value repeat with a period of 5 twips (+0.0003, +0.0004,
−0.0004, −0.0002, 0 px), which is the signature of a grid that 1 twip (204.8 units) does
not divide. Two independent confirmations of the same unit follow below: tracking
(2.3) and the pitch of `auto` multiples (2.2).

**Glyph advances are the unrounded `advance / upm × size`**, per glyph and in total: every
per-glyph step is identical to 1e-4 px and none is a whole pixel (e.g. Calibri 6.5 pt `H`
steps 16.8742 px = 1276/2048 × 27.083). **A run boundary rounds nothing**: ten glyphs as one
run and as ten runs end at the same x to 1e-4 px. Right- and centre-aligned lines start at
fractional pixels (e.g. 2035.8, 1151.451), and the right margin they align to is the
*authored* page width (2180.417 px = 10466 twips), not the device-rounded page box.

So the answer to "does each glyph snap, or only the line start": **neither**. For a
2048-upm face at a half-point size an advance is `w × hp / 4096` pt — an exact integer in
the layout unit — so for Calibri, Arial, Times New Roman, Cambria, Aptos and Verdana the
1/4096 pt grid is invisible in advances and only shows in twip-authored quantities
(indents, tabs, tracking, `exact`/`atLeast` line heights). A 1000-upm face (CFF) would
show it per glyph; **not measured**.

#### 2.2 The line advance: exact accumulation, and a rounding against one edge of the line

`make_line_advance_probe.py` / `read_line_advance_probe.py`. One group per page, each
group 3–80 single-line paragraphs of `H` whose paragraph mark has the same size. Three
sweeps:

* **sizes** — Calibri, Aptos, Arial, Times New Roman, Cambria, Verdana × every half point
  6–12 pt, then 13, 14, 15, 16, 18, 20, 22, 24 pt; single spacing. 126 groups, 5,478
  baselines.
* **rules** — Calibri, Times New Roman, Arial × 6.5 / 8.5 / 11 / 14.5 / 20 pt × `auto`
  240 / 276 / 360 / 233, `exact` 240 / 277 / 301, `atLeast` 1 / 277 / 400.
* **dense** — Calibri, Times New Roman, Cambria × 6.5 / 9.5 / 11 / 14.5 / 20 pt × `exact`
  250 / 263 / 290 / 310 / 333 / 350 / 411 / 457, `atLeast` 290 / 333 / 350 / 411 / 457 /
  500, `auto` 200 / 220 / 250 / 264 / 288 / 300 / 320 / 400 / 480 — chosen so the
  fractional pixel of the pitch covers [0, 1) for each rule.

**The model** (`docx2svg.vertical`):

1. **Pitch.** `natural = (ascent + descent + lineGap) / upm × size`, from `hhea` (or the
   `OS/2` typo values under `USE_TYPO_METRICS`). `auto`: `natural × line / 240`;
   `exact`: `line` twips; `atLeast`: `max(natural, line twips)` — each rounded to
   1/4096 pt. Line tops accumulate it exactly: `top_k = top_0 + k·h`. Every one of the
   621 groups is consistent with that pitch; the feasible range around it is ~1e-4 of
   `h`, and it contains the exact value in every group.
2. **Rounding against an edge.** With `f = frac(h)` in device px:
   * `f ≥ ½` — **bottom-anchored**: `baseline = round(top + h) − round(below)`;
   * `f < ½` — **top-anchored**: `baseline = round(top) + round(above − f/2)` for a plain
     line, and `round(top) + floor(h) − round(below)` for a line whose extra space is
     above its text. (`exact`: `round(above − f/4)`; the `λ = 1` first chosen here was
     never discriminated and is refuted by `style-document.docx` — see *The line box*.)
3. **Where the text sits in the line.** `auto` single: `above = ascent + lineGap`
   (the gap goes above), `below = descent`. `auto` below 240: both scaled by
   `h / natural`. `atLeast` taller than the text: extra space above, `below = descent`.
   `exact`: `below = h / 5` whatever the face — the baseline is at 80% of the line
   (truncated to the layout unit, which only an `exact` line with space after shows).
   `auto` above 240: the extra is **below** the text, which makes it a line box with
   space below its text — the rule for those is *The line box*'s, not this one.

| Rule | Baselines exact | Groups exact |
| --- | --- | --- |
| `auto` 240 (single) and `atLeast` below natural | **6,144 / 6,144** | 141 / 141 |
| `auto` < 240 | **2,154 / 2,154** | 45 / 45 |
| `exact` | **6,705 / 6,705** | 165 / 165 |
| `atLeast` above natural | **4,386 / 4,386** | 135 / 135 |
| `auto` > 240 | 4,016 / 4,470 → **4,470 / 4,470** | 115 → **135 / 135** — closed by *The line box* |
| `layout-sweep.docx` (14 pt headings + 11 pt body, 2 pages, 2 sections) | **91 / 91** | — |

That is 19,389 baselines at 0 px across six faces, 21 sizes and all three rules, which is
the Phase 2 done-condition for the vertical model with one named exception -- since
closed: all 23,859 now (*The line box — measured*). The recorded
observations are `tests/fixtures/line-advance-observations.json` (measurements only), and
`tests/test_vertical.py` holds the model to every one without Word.

The fitted parameters, and how tightly the data pins them:

| Parameter | Chosen | Feasible interval over all groups |
| --- | --- | --- |
| top-anchored plain line: `round(above − λ·f)` | λ = ½ | [0.386, 0.592) over 81 groups |
| bottom-anchored: `round(below + δ)` | δ = 0 | [−0.005, 0.002) over 104 groups (single) |
| `atLeast` top-anchored: `round(above − λ·f)` | λ = 1 | [0.984, 1.371) over 57 groups |
| `atLeast` bottom-anchored δ | 0 | [−0.005, 0.077) |
| `exact` top-anchored λ | ~~1~~ **¼** | [−0.001, 2.000) here — **not discriminated**; the line-box probe narrows it to (0, ½), which is everything any `exact` line can show (*The line box*) |
| `exact` below fraction | 1/5 | every `exact` group: 11 pitches × 5 sizes × 3 faces |
| tie rule | half up | round-half-down leaves 41 of the dense sweep's 345 groups unexplained |

**Refuted, and kept here:**

| Hypothesis | Score |
| --- | --- |
| One rounding of the exact baseline, `round(top₀ + ascent + k·h)` (Phase 0's) | 4,040 / 5,478 baselines; 10 / 126 groups |
| Always top-anchored, `round(top) + round(ascent)` | 4,133 / 5,478 |
| Always bottom-anchored, `round(top + h) − round(descent)` | 4,714 / 5,478 |
| Split slack symmetrically on both edges (bottom-anchored `round(descent + (1−f)/2)`) | 5,116 / 5,478 |
| Round the pitch, the ascent, or both, to px, twips, pt, 1/8 pt, 1/100 pt, 1/1000 pt, 1/576 in, 1/96 in, 1/1024 in, EMU, 1/16 px or 1/64 px (× round / floor / ceil, components separately or summed) | best 1,356 / 1,826 (Calibri + Aptos) |
| GDI `ceil(ascent) + ceil(descent)` (docx-renderer) | 57 px at 11 pt against a 55.949 px pitch |
| Pitch from the *drawn* size (46 px for 11 pt) | 56.15 px against 55.949 |
| Times New Roman at the metrics of the copy Word embeds (`lineGap` 0, 1.107 em) | pitch is 1.1499 em in every group |
| An unquantised pitch (no 1/4096 pt) | 123 / 135 `auto`-multiple groups consistent, 134 / 135 quantised; exact 263 needed a round-half-*down* to fit, which then broke others |

**Settled since, by the line box: `auto` multiples above one (`w:line` > 240).** What
follows is what Phase 2 measured and refuted; the rule is under *The line box —
measured*: the multiple's extra is space **below the text** inside the line's box, and a
box with space below its text rounds in three parts. The probe named at the end of this
paragraph was run (Times New Roman 14.5 pt, 241–480, 480 baselines) and is part of
`make_line_box_probe.py`. Every such group is
top-anchored (phase 0), the pitch is right, and the baseline is `round(ascent + lineGap)`
below the pixel top in 115 of 135 groups; the other 20 are exactly 1 px off, never more:
Times New Roman / Arial 14.5 pt at 276–480 when `frac(h) > ½` (+1), Calibri 20 pt at every
multiple but 480 (+1), Cambria 9.5 pt (+1) and 14.5 pt except 320 (+1), Times New Roman
6.5 pt at 320 (+1), Calibri 20 pt at 480 (−1 against the natural line's offset). Tried and
failed: the natural single-spaced offset (118 / 135), scaling by `H/h` or `H/natural`, the
half-slack rule with the multiple's `f` (61 / 135) or the natural line's, bottom
references, `max`/`min` of top- and bottom-referenced offsets (best 126 / 135), and every
two-term linear combination of the fractional parts of `h`, `natural`, the extra space
and the rounded descent (best 128 / 135). **The probe that would decide it:** one face
and one size (Times New Roman 14.5 pt is the most sensitive), `w:line` swept 241–480 in
steps of 1, so the offset can be plotted against `frac(h)`, the extra space and its
fractional part separately — the current data has at most nine multiples per size, too
few to separate them.

**Other variables held fixed, so not yet known:** no `settings.xml` (so `w:compat` is
whatever Word assumes for a document without one); no `w:docGrid` type; no
`w:snapToGrid`; one line per paragraph and one size per line (a line mixing sizes, or a
paragraph's second line, has not been probed — the fixture's wrapping paragraph is four
lines of one size and matches); A4, 1-inch margins, portrait; Word 16.106 on macOS at the
PDF export's 300 dpi.

*Since measured, in the style-inheritance work:* a line mixing faces or sizes (largest
ascent and largest descent, separately; a list label's height not scaled by an `auto`
multiple), paragraph spacing between paragraphs (the larger of after and before), and
space before at a page top after a manual break (dropped). `settings.xml` with
`compatibilityMode` 14 or 15 moved no baseline of the generated inherited-style document.
See *Style inheritance — measured*.

#### 2.3 Kerning and tracking

`make_kerning_probe.py` / `read_kerning_probe.py`. `HHHHHHHH` (control), `AVAVAVAV`,
`ToTyToTy`, `LTWAYoVA` in Calibri, Arial and Times New Roman at 11.5 and 20 pt; plain, with
`w:kern w:val="2"`, with `w:spacing w:val="7"` and `"-5"` twips, and kern plus 7; each as
one run and as one run per glyph. Predictions from the faces' `kern` tables (read locally
with fontTools; nothing copied). **242 lines, worst residual 0.0001 px.**

* **Without `w:kern` Word does not kern at all.** Every plain width is the plain sum of
  advances. This is the opposite of PowerPoint body text, which the sibling measured
  kerning by default — the two applications differ, and neither answer transfers.
* **With `w:kern`, exactly the `kern` table's pairs**, in all three faces, and **across
  run boundaries**: ten single-glyph runs kern exactly like one run. (For these pairs the
  legacy `kern` table and GPOS `kern` agree; a pair on which they disagree has not been
  probed.) The threshold semantics of `w:kern/@w:val` (a minimum size) are not tested
  beyond "2 half points switches it on".
* **Tracking is added after every glyph, including the last, and is itself in 1/4096 pt**:
  +7 twips is 1433.6 units and comes back as 1434 (1.45874 px per glyph, not 1.45833).
  −5 twips (−1024 units) is exact.
* **Kerning and tracking add.** Since nothing horizontal is rounded to the device grid,
  "before or after the rounding" has no content on this axis: the only rounding is the
  1/4096 pt of the tracking amount itself, which happens before either is applied.

*Since measured, at the edge of a line* (`make_wrap_kern_probe.py` /
`read_wrap_kern_probe.py`, recorded in `tests/fixtures/wrap-kern-observations.json`,
`tests/test_wrap_kern.py`): pptx-agent's wrap-boundary probe asked where PowerPoint breaks
`Pass`, `Fail`, `Total`, `AVAWAY` and `Review` in Aptos, Calibri and Arial at 12, 18 and
24 pt; this asks Word, with the same words, faces, sizes and text widths -- each word's
plain, legacy-kerned and GPOS-kerned width, less 0.3 and 0.05 pt and plus 0.05 and 0.3 --
with `w:kern w:val="2"` and without. **With `w:kern`, the legacy table: 324 / 324 wrap
verdicts and every width kept on one line to 0.0001 pt**; GPOS's pairs get 317 (Aptos's
`ss`, in GPOS alone, is not charged: "Pass" fits where GPOS says it needs the 0.29 pt).
**Without `w:kern`, no kerning: 324 / 324.** A word fits exactly when it is no wider than
the line (wrapped at 0.014 pt over, kept at 0.002 pt under). Two variable faces, Noto Sans
JP and STIX Two Text, kerned, came out unkerned to 0.001 pt -- where PowerPoint charges
their GPOS pairs. ooxml-common carries both as a rule (`DrawingRules.kerning`: Word's the
legacy table, nothing for a variable face), and docx2svg's DrawingML text -- charts and
SmartArt, measured by ooxml-common -- now uses Word's.

#### 2.4 Which copy of a face Word lays out with

Word 16.106 ships its own Times New Roman (v7.00, `hhea.lineGap` 0) and macOS ships
another (v5.01, `lineGap` 87); their advances are identical and 25 Greek kern pairs differ.
**Word embeds v7.00 in the PDF** (read from the embedded subset's `head` table) **but lays
out with v5.01's metrics**: only `lineGap` 87 reproduces the pitch (2355/2048 em, the same
as Arial, which is why their single-spaced gaps were identical in every group), and only
v5.01's kern table reproduces `ΤΩΤΩΥμΥμ` kerned (−20.345 px; v7.00 has no such pairs).
So "the face" for layout is resolved through the system's font registry by name, and the
file drawn can be a different one. A renderer that measures with the file it embeds will
be wrong for any face installed twice. Calibri, Aptos and Cambria exist only in Word's
bundle here; Arial's two copies agree on `hhea`.

---

## Style inheritance — measured (Phase 6, first row; Suggested order step 2)

**Done: the cascade resolves every glyph of every document here to the face, weight,
slant and size Word drew it in — 1,770 probe glyphs and 31,057 document glyphs, no
exceptions.** Baselines of documents that state nothing directly reproduced to 0 px for
most lines, and the rest were within 1 px and traced to the vertical model, not the
cascade (below); since *The line box — measured*, every one of them is exact.

`docx2svg.resolve` is the stage. `resolve_paragraph`, `resolve_run` and `resolve_mark`
(the paragraph mark, which has a face and size, takes part in the height of a line
that holds nothing but spaces, and is drawn as a space in Word's PDF) return a
`Resolved`: values plus, for each, the **origins** that decided it — the level, the style applied, the style in the `basedOn`
chain that declared it, and for a toggle every level with whether it flipped:

```
b = False  <- paragraph style 'PB' [True, flips]; character style 'CB' [True, flips]
sz = 22  <- docDefaults
```

`character_format(resolved, char, document)` then picks the `w:rFonts` slot for one
character and resolves theme references. `parse/properties.py` reads every level as a
flat map of what it *declared* (so "absent" and "declared false" stay apart);
`parse/styles.py` reads styles (with `w:tblStylePr`), the theme's font scheme,
numbering and `w:themeFontLang`. Standard library only.

### The probes

`tools/make_style_probe.py` / `read_style_probe.py`: 12 documents, 172 cases, one
variable each, read off the PDF without a font file (face and weight from the
`/BaseFont` of the glyph's text object, size from its `Tm`, caps as uppercase text,
vanish as absent text, indents from the pen x). The reader parses the `.docx` with
`docx2svg` and resolves every run and mark itself; the generator's intent is not
consulted. **1,770 glyphs, 0 misses**; 104 glyphs are drawn in a *fallback* face (Georgia
has no Hebrew; Word draws East Asian text in MS Mincho when the slot names no East Asian
face), counted separately because that is substitution, not resolution.
`tests/test_resolve.py` holds the resolver to all of it from
`tests/fixtures/style-observations.json` without Word.

### Where Word is not ECMA-376

| Claim in ECMA-376 | What Word 16.106 drew | Cases |
| --- | --- | --- |
| 17.7.3: a toggle property (`b`, `i`, `caps`, `vanish`...) in a style **flips** the value accumulated so far, and `false` leaves it unchanged — throughout the style hierarchy | Within one style's **`basedOn` chain it is a plain override**: bold based on bold is bold; `b=0` based on bold is not bold. *Across* levels (table, paragraph, character style) each level's resolved value flips the result **iff it differs from the `w:docDefaults` value**. With docDefaults silent that is the spec's XOR (paragraph bold + character bold = not bold; three levels = bold; `false` = no change). With docDefaults bold, paragraph + character bold stay bold, and a style saying `false` turns bold **off**. Direct formatting is absolute, as the spec says. | t08–t14, g02–g11; a literal reading of 17.7.3 disagrees with Word in 17 of the 63 bold cases (printed by the reader) |
| 17.7.2: numbering sits below the paragraph style | Only when the `w:numPr` comes *from* the style. A **direct** `w:numPr` puts the list level's indent **above** the paragraph style's (720 over the style's 2160). | p02, p03 |
| 17.3.2.26: a character's `w:rFonts` slot is chosen by its Unicode range, Hebrew and Arabic being complex script | **Complex script is decided by the run's `w:cs` / `w:rtl`**, not the character. Unmarked Hebrew and Arabic are drawn at `w:sz`, bolded by `w:b` (not `w:bCs`), in the non-complex slot's face; with `w:rtl` or `w:cs`, at `w:szCs` in the `cs` slot's face. | f09, f13, f14, f20, f21, f24, f25, f27 |
| (theme fonts) `majorEastAsia` / `minorBidi` name the theme's `a:ea` / `a:cs` typefaces | They go **through `w:themeFontLang`**: the theme's per-script font for that language's script (`Hans` for zh-CN, `Hebr` for he-IL) if listed, else `a:ea`/`a:cs`. **Without `w:themeFontLang` Word ignores the theme for those slots** (drew Times New Roman for a `minorBidi` whose theme said Courier New, MS Mincho for `minorEastAsia`). | f15, f18, f23 in three settings variants |
| An absent `w:sz` is 10 pt | **12 pt**, with or without a styles part; absent paragraph spacing is **after 160, line 278 auto** (a 104 px pitch for 12 pt Calibri). Both match this machine's `Normal.dotm`, whose theme Word does *not* use (Calibri, not Aptos) — so these are Word's state, reported with `application default` origins. | n01–n04, `style-no-styles`, `style-empty-styles` |
| (no statement) paragraph spacing | Space after and the next paragraph's space before **collapse to the larger**, and a paragraph starting a page after a **manual page break loses its space before**. Both measured on whole documents (below): adding the two spacings instead scores 7 / 34, 1 / 20, 1 / 22 and 427 / 453; keeping space before at the page top scores 17 / 34 and 27 / 453. | sample-long (17 page tops), style-document |

**Agreeing with the spec, measured:** direct formatting is absolute (t04–t07, t30, g04);
a theme attribute beats the explicit face on the same element (f04); a level declaring
either the explicit or the theme attribute of a slot replaces both below it (f03, f05,
f06, f16, f17); the four slots inherit independently (f07); `w:hint="eastAsia"` sends
U+201C/D, §, — and, with a Chinese East Asian language, é to the East Asian slot (f11,
f12); a missing or unknown `pStyle` means the default paragraph style (t23); a
paragraph style with no `basedOn` does not inherit Normal (g06); the table style sits
below the paragraph style, including Normal (p06–p08: Normal's 11 pt beats the table's
14 pt — the "table style overrides Normal" behaviour described for LibreOffice was not
seen, in a document with no `overrideTableStyleFontSizeAndJustification` compat setting).

**`w:next` and `w:link` change nothing that renders.** An `rStyle` naming a linked
*paragraph* style is ignored, not redirected to its linked character style (t22). The
**default character style contributes nothing**, even through `basedOn` (g01, g09).

### Closing the loop: baselines of documents that state nothing directly

`tools/baselines.py` predicts every baseline of a document from resolved properties
only — spacing, line rule, and every item's face and size from `docx2svg.resolve`, face
metrics looked up by resolved name (`tools/face_metrics.py`, integers only) — and
compares with Word. Taken from the oracle: which paragraph each drawn line belongs to and
its page (Phases 3 and 4). `tools/read_baselines.py [--record]`;
`tests/test_baselines.py` pins every score and checks every glyph offline from
`tests/fixtures/baseline-observations.json`.

| Document | Baselines exact (then → with the line box) | Out of scope | Glyphs in the resolved face and size |
| --- | --- | --- | --- |
| `style-document.docx` — generated; theme fonts, `basedOn` chains, character styles, all three line rules, a list, a manual page break; **no run or paragraph states a face, size or spacing** | 24 / 34 → **34 / 34** | 0 | 1,361 / 1,361 |
| `layout-sweep.docx` (Phase 0), now through the resolver | **91 / 91** | 0 | 679 / 679 |
| `samplelib/sample-long.docx` — third-party, 18 pages, 17 manual breaks | 452 / 453 → **453 / 453** | 35 (drawings) | 27,458 / 27,458 |
| `samplelib/sample-resume.docx` | 18 / 20 → **20 / 20** | 0 | 724 / 724 |
| `samplelib/sample-simple.docx` | 18 / 22 → **22 / 22** | 7 (a table and what follows it) | 835 / 835 |
| `samplelib/sample-blank.docx` | 0 / 0 | — | — |
| **Real-world total** | 488 / 495 → **495 / 495** | 42 | **29,017 / 29,017** |

Every miss was exactly 1 px, and none was a cascade error — a wrong face or size moves a
line by several pixels and shows in the glyph column, which is clean. Where they came
from, as first recorded (each is now closed; *The line box — measured* says by what):

* **The three Title lines of the samples** (−1 each): a paragraph with a bottom border.
  *(Closed: the border and the space after are below the text in the Title line's box.)*
  The space the border takes (`w:space` + width, in 1/4096 pt) is right — every line
  after it is exact — but the bordered line itself sits 1 px low.
  `make_border_probe.py` / `read_border_probe.py` reproduces it: the line after a border
  is exact in **132 / 132**, the bordered line in 108 / 132 with the plain rule, and
  none of five rules tried (anchoring on the pitch including the border, etc.) does
  better. ~~Open.~~
* **Four sample lines under `auto` 276** (±1): the `auto` > 240 rule Phase 2 left at 115
  of 135 groups. Same class, no new information. *(Closed: the multiple's extra is
  below the text; three of the four also carry the heading's space before above it.)*
* **Ten lines of `style-document.docx`** (all −1): two `atLeast` lines whose pitch has a
  fractional pixel of exactly ½, list and mixed-size lines, three `exact` lines after
  them, one heading. Solving for the line-top offset each observed baseline allows shows
  **no single offset reconciles them** (two `atLeast` lines with the same predicted
  fractional top round opposite ways), so this is not an accumulated spacing error but
  the baseline-in-line rounding at or near a tie. ~~Open~~; the probe that would decide
  it is Phase 2's dense sweep extended with paragraph spacing between lines of different
  pitch, and `atLeast`/`exact` values whose pitch is exactly *n* + ½ px. *(Closed: six
  lines carry space after below their text, one a heading's leftover space before above
  it — which is why no single top offset reconciled them — and the three `exact` Code
  lines were Phase 2's undiscriminated `exact` λ, not spacing at all.)*

### What the real documents exposed that the probes did not

The third-party documents (`tests/fixtures/samplelib/`, provenance and licence in its
`PROVENANCE.md`) were the harder test, and six things came out of them:

1. **Table style conditional formatting is a cascade level.** 37 header and first-column
   glyphs of `sample-simple.docx` are bold Calibri only through `w:tblStylePr`
   (`firstRow`, `firstCol`), chosen by the cell's position and the table's `w:tblLook`.
   Now resolved, in ECMA-376's order; the corner cell (bold first row meets bold first
   column) is drawn bold, so conditional formats **override one another inside the table
   level** rather than toggling. Band sizes other than 1 are not read.
2. **Paragraph spacing collapses to the larger** of after and before (adding them: 1 / 20
   and 1 / 22 on the resume and simple documents) and is suppressed at the top of a page
   after a manual page break (keeping it: 27 / 453 on `sample-long.docx`, every chapter
   page 100 px low).
3. **A list label is part of the line**, resolved as the paragraph mark's properties plus
   the level's `w:rPr` (SymbolMT here).
4. **A line mixing faces takes the largest ascent and the largest descent separately**
   (not the tallest item). `make_mixed_line_probe.py`: Cambria + Symbol single-spaced
   is 56.24–56.33 px against 56.26 (separate maxima) and 56.15 (tallest).
5. **A list label's height is not scaled by an `auto` multiple**; the text's is.
   Cambria + Symbol label under `auto` 276 / 360: 64.29–64.38 / 83.10–83.19 px,
   reproduced as natural(all) + (multiple − 1) × natural(text) = 64.32 / 83.13; the same
   Symbol glyph as an inline *run* is scaled with the text (64.70). ~~One group is still
   0.012 px outside its interval (inline Symbol at 360).~~ *Closed by its baselines:* the
   extra is over the **tallest text item's own** natural height (SymbolMT, 56.150 px),
   not over the combined extent (56.262): 64.685 / 84.337 px, every one of the 440
   baselines of the probe exact, against 421 with the combined extent.
   `read_mixed_line_probe.py`.
6. **Word lays out Symbol with its own bundle's SymbolMT** (1.2251 em), not macOS's
   Symbol (1.0 em) — the opposite of Times New Roman (2.4). So "which copy of a face"
   is two observations, not a rule; `face_metrics.PREFER_BUNDLE` records the exception.

### Not settled, or not measured

* ~~The 1 px residuals above (bordered line; ties in `atLeast`/`exact`/mixed lines).~~
  Closed by *The line box — measured*.
* The toggles not drawn: `strike`, `dstrike`, `smallCaps`, `outline`, `shadow`,
  `emboss`, `imprint` are assumed to follow `b`/`i`/`caps`/`vanish`.
* ~~Space before at the top of a page after a *natural* page break (only manual breaks
  were measured).~~ Measured by the page-top probe: dropped after a natural or manual
  break, kept by a section's first paragraph, and by `pageBreakBefore` below mode 15;
  see *The page top — measured*.
* `overrideTableStyleFontSizeAndJustification` (set in the samples' `settings.xml`) and
  table style properties against Normal under it; `w:tblStyleRowBandSize`; numbering
  through `w:numStyleLink`; `w:latentStyles` (ignored).
* `w:hint` outside Latin-1 and the punctuation ranges; `w:hint="cs"`; complex script for
  scripts other than Hebrew and Arabic, and in a right-to-left paragraph.
* What an explicit East Asian face that is not East Asian becomes (Courier New was drawn
  as MS Mincho) and what Word draws for an empty theme slot — substitution, a later
  stage.
* Whether the 12 pt / after-160 / line-278 defaults come from `Normal.dotm` or are
  built in; the same bytes gave Aptos marks in Phase 0 and Calibri now.

---

## The line box — measured (Suggested order step 3)

**Done: every in-scope baseline of every probe and every document reproduces to 0 device
px**, and the three residuals Phase 2 and the style work left open were **one cause**:
Word does not round a baseline inside the line's pitch. It rounds it inside a **line
box** that also holds the space around the text — the paragraph's space before (what is
left of it after the previous paragraph's space after) above the text on the first line;
the space after and the bottom border below the text on the last line; and the extra
height of an `auto` multiple above one, which also goes below the text. The `auto`
multiples were lines with space below their text; the bordered Title was one too; and
seven of the ten `style-document.docx` misses carried space after or before in the box.
The other three were not spacing at all (the `exact` λ, below).

| Corpus | Before | After |
| --- | --- | --- |
| Phase 2 probes (sizes, rules, dense; 6 faces) | 23,405 / 23,859 (the 135 `auto` > 240 groups at 115) | **23,859 / 23,859** |
| Bordered-line probe (`make_border_probe.py`) | 240 / 264 | **264 / 264** |
| Mixed-line probe baselines (not scored before) | 387 / 440 | **440 / 440** |
| **Line-box probe** (`make_line_box_probe.py`, new: 52 documents, 497 pages) | 9,220 / 12,069 | **12,069 / 12,069** |
| `layout-sweep.docx` | 91 / 91 | **91 / 91** |
| `style-document.docx` | 24 / 34 | **34 / 34** |
| `samplelib/sample-long.docx` | 452 / 453 (35 out of scope) | **453 / 453** |
| `samplelib/sample-resume.docx` | 18 / 20 | **20 / 20** |
| `samplelib/sample-simple.docx` | 18 / 22 (7 out of scope) | **22 / 22** |
| **Real-world total** | 488 / 495 | **495 / 495** |

Out of scope, unchanged: tables, drawing paragraphs, and lines below a drawing on the
same page. No currently-exact baseline moved in any corpus. `tests/test_line_box.py`,
`test_vertical.py` and `test_baselines.py` hold all of it without Word.

### The probe

`make_line_box_probe.py` / `read_line_box_probe.py`. One line per paragraph (`Hx`, mark
the same face and size), every page led by a plain anchor with `w:pageBreakBefore`, pages
filled to 85% so nothing breaks naturally. One variable per sweep, each swept in steps of
one twip or one unit so every fractional pixel in [0, 1) occurs:

| Sweep | What varies | Faces × sizes | Baselines |
| --- | --- | --- | --- |
| `after` | space after 1–120 twips | Calibri, Times New Roman, Arial, Cambria × 6.5 / 11 / 14.5 / 20 pt | 1,996 |
| `before` | space before 1–120 twips | Calibri, Times New Roman, Cambria × the same | 1,497 |
| `border` | bottom border `w:space` 0–31 pt × `w:sz` 2 / 6 / 12 / 27, each followed by a plain line | Calibri, Times New Roman × 11 / 14.5 / 20 pt | 1,618 |
| `exact` | `w:line` 240–480 twips `exact` | Calibri 11 and 6.5, Courier New 10, Times New Roman 11 | 996 |
| `multiple` | `auto` `w:line` 241–480, two lines each (the probe Phase 2 named) | Times New Roman, Arial, Cambria 14.5; Calibri 20 and 11; Cambria 11 | 3,028 |
| `multiple-after` | `auto` 276 with space after 1–120 | Cambria, Calibri 11 | 250 |
| `combo-after` / `-before` / `-both` | `exact` 263/290/333, `atLeast` 350/411, `auto` 200/220/276/360 × spacing 1–48 twips | Calibri 11, Times New Roman 14.5 | 2,684 |

### The rule

`docx2svg.vertical.baseline_in_box(top, box)`, with `top` the exact top of the box and
`H` its height (space before + pitch + space after + border). Two regimes, decided by
`E`, everything in the box **below the text line** (multiple's extra + border + space
after):

1. **`E = 0`: Phase 2's rule, on the box.** `f = frac(H)`. `f ≥ ½`: `round(top + H) −
   round(below)`. Otherwise `exact` (no space before): `round(top) + round(above − f/4)`;
   a line with space above its text (`atLeast` above the text, or any space before):
   `round(top) + floor(H) − round(below)`; plain: `round(top) + round(above − f/2)`.
   `above` includes the space before.
2. **`E > 0`: always top-anchored, in three parts.** The space above the text `Xa` (space
   before, plus an `atLeast` line's extra) rounds to whole pixels; so does `E`; the text
   line — its own pitch `t` (the natural height for single, `atLeast` and multiples;
   the pitch for `exact` and multiples below one) — keeps what is left,
   `P = t + (Xa − round(Xa)) + (E − round(E))`, and sits on its bottom but no higher than
   its ascent: `round(top) + round(Xa) + max(round(A), round(P) − round(D))`. And never
   lower than its descent above the box's real bottom edge: `min(…, round(top + H) −
   round(D))`.

Found in that order from the data: the after sweep showed that with space after a line
is top-anchored whatever `f` is, with an offset that depends only on `frac(E)` — `57` for
Times New Roman 14.5 pt exactly when `frac(E) ∈ (0, ½)` — which is `round(t + E −
round(E)) − round(D)`; Arial at the same pitch stayed at `57 = round(A)` where that gives
56, hence the `max`. The combos then showed `atLeast` needs its extra rounded on its
own (Calibri 11 at 411 twips stays at 74 for every `E`, Times New Roman 14.5 follows `P`),
and 6 lines with `E` under half a pixel showed the bottom clamp.

Three smaller facts came out with it, each needed by some baselines and changing no
other:

* **Paragraph spacing collapses in twips.** The gap is the previous paragraph's space
  after plus `before − after` twips when positive, converted as that difference
  (`paragraph_gap_px`): 1 twip is 205 units, not 1,638 − 1,434 = 204. The previous
  paragraph keeps its whole space after in its own box.
* **A border's width is truncated to whole twips** (`border_px`): `w:sz` 27 is 67 twips.
* **An `exact` line's part below the baseline is `h/5` truncated to the layout unit.**

And one pitch fact, from the mixed-line probe's baselines: an `auto` multiple's extra is
over the **tallest single text item's** natural height (a label excluded), not over the
line's combined extent.

### Constants, observations, residuals

| Constant | Value | Evidence | Residual |
| --- | --- | --- | --- |
| `LAMBDA_EXACT` (top-anchored `exact`, nothing below the text) | ¼ | every `exact` baseline reproduces for λ ∈ (0, ½) and for nothing outside it: 0 misses 72 line-box and 555 Phase 2 baselines at ties, ½ misses 8, 1 (Phase 2's choice) misses 40 and the three Code lines of `style-document.docx` | **not expected to narrow further**: an `exact` line's rounding depends only on `w:line`, its pixel pattern repeats every 120 twips (25 px) up to one layout unit, and the sweep covers each residue twice. ¼ is the middle of the interval, not a measurement |
| `exact` below | `h/5` truncated to 1/4096 pt | 12,069 / 12,069; rounded: 11,961; unquantised: 11,949 (the `exact` + space after combos) | 0 |
| Border width | `floor(sz × 20 / 8)` twips | border sweep 1,618 / 1,618; keeping the half twip: 1,428 | 0 |
| Paragraph gap | after + q(before − after twips) | line-box 12,069 / 12,069; `max` of the two converted lengths: 3 misses (combo-both, at ties) | 0 |
| Multiple extra over | tallest text item | mixed-line 440 / 440 | 0 |
| Space after at the last line of a page | in its box | 148 such lines (each before a `pageBreakBefore` paragraph); left out: 102 / 148 | only measured before a `pageBreakBefore` page |

### Refuted, with scores

Line-box scores out of 11,572 leave out the probe's 497 page anchors (plain lines).

| Hypothesis | Score |
| --- | --- |
| Phase 2's model: each line rounds in its own pitch, spacing and borders stacked between lines | line-box probe 9,220 / 12,069; with the new `exact` and multiple rules but spacing still outside the box, 10,305 / 12,069; documents 27 / 34, 18 / 20, 18 / 22, 452 / 453 |
| Phase 2's rule applied to the whole box, `E > 0` included | line-box 9,019 / 11,572 (anchors excluded), Phase 2 probes 22,638 / 23,859, border 221 / 264, documents 469 / 620 |
| One rule for both regimes (the `E > 0` rule on bare lines) | 34,917 / 36,315 |
| `E > 0` without the ascent floor (`round(P) − round(D)` alone) | line-box 10,524 / 11,572, Phase 2 probes 23,013 / 23,859, documents 529 / 620 |
| `E > 0` without splitting off the space above (`Xa` rounded with the text) | 136 line-box misses |
| `E > 0` without the bottom clamp | 6 line-box misses (all `E` < 0.42 px) |
| Space before left out of the box | line-box 10,977 / 11,572; documents 616 / 620 |
| Space after and border left out of the box | line-box 10,201 / 11,572; border 240 / 264; documents 609 / 620 |
| The whole collapsed gap in the second paragraph's box (the first keeps only what exceeds it) | documents 32 / 34, 19 / 20, 19 / 22 |
| Offsets for `E > 0` of `ceil(A)`, `round(A + frac(H))`, `round(A + 1 − frac(E))` (each fitted the documents' lines first) | 1,068, 927, 938 / 1,497 on the first twelve `after` documents |
| Multiple extra over the combined extent (the earlier model) | 19 of the 44 inline-Symbol-under-a-multiple baselines off |

### Not settled, or not measured

* **Top, left/right and between borders**, a border shared by a group of consecutive
  paragraphs with identical borders (Word merges them), and borders with shading: none
  probed. Only `w:pBdr/w:bottom` with `w:val="single"` was.
* **`w:beforeLines` / `w:afterLines`**, and `w:contextualSpacing` inside the box beyond
  the documents' cases: not probed. (`w:beforeAutospacing` since: *Autospacing —
  measured*.)
* ~~**Space before at a page top after a natural page break, or on a `pageBreakBefore`
  paragraph**~~: measured since, with section starts and seven compatibility settings;
  see *The page top — measured*.
* **Mixed faces with space around them** are measured only as far as the documents
  have them (lists, one larger run); multi-line paragraphs likewise (the first and last
  line carry the spacing; the documents' two- to eleven-line paragraphs all agree).
* `w:docGrid` / `w:snapToGrid`, compatibility modes other than the default: held fixed,
  as in Phase 2.

The probe that would extend it: the same one-line-per-paragraph layout with
`w:pBdr/w:top` and `w:between` swept like `border`, and a `before` sweep whose paragraphs
each start a page (after a natural break and with `pageBreakBefore`).

## Nine more real documents — scored, not fitted (Suggested order step 3½)

**What it is:** a test of the vertical model on documents that **no probe shaped**:
from other Word versions and compatibility modes, and with **natural pagination**. The
model was **not changed**. Every miss is recorded below with its likely cause and the
probe that would decide it, so the next change can be attributed.

* **`tests/fixtures/wordto/`**: five documents from <https://word.to/samples/docx/>,
  committed (their terms say they are free for any use and not copyrighted;
  `wordto/PROVENANCE.md`). Word for Mac 14, the same template as `samplelib`,
  compatibility mode 14.
* **`filesamples`**: four documents from <https://filesamples.com/formats/docx>. They
  are **not committed**, because the site grants no licence (`tests/fixtures/
  filesamples/PROVENANCE.md`). They live in the gitignored `scratch/filesamples/`, and
  `tests/test_filesamples.py` skips where they are absent. They were written by Word 12
  (compatibility mode 12, an empty `w:compat`), Word 15 in mode 14, Word 14 in mode 14,
  and Word 15 in mode 15.

These documents are scored exactly as `samplelib` is. Which paragraph and page each
drawn line belongs to still comes from the oracle. Out of scope: table cells, drawing
paragraphs (now also picture bullets), lines below either on the same page, and, new
here, lines that are no body paragraph's text: footnotes, the table-of-contents field
result, and frames. When one of those appears, the rest of its page is out of scope too.

| Document | Word / mode | Pages | Baselines exact: the model | With the hypotheses below (diagnostic) | Out of scope | Glyphs in the resolved face and size |
| --- | --- | --- | --- | --- | --- | --- |
| `wordto/sample-1page.docx` | Mac 14 / 14 | 1 | 0 / 7 | **7 / 7** (document start) | 0 | 424 / 424 |
| `wordto/sample-5pages.docx` | Mac 14 / 14 | **1** | 0 / 26 | **26 / 26** (document start) | 30 (a table) | 1,461 / 1,469 ¹ |
| `wordto/sample-10pages.docx` | Mac 14 / 14 | **3, both breaks natural** | 39 / 70: pages 2 and 3 **39 / 39** | **70 / 70** (document start) | 19 (a table) | 3,795 / 3,797 ¹ |
| `wordto/sample-with-images.docx` | Mac 14 / 14 | 1 | 0 / 8 | **8 / 8** (document start) | 9 (a drawing) | 903 / 903 |
| `wordto/sample-with-table.docx` | Mac 14 / 14 | 1 | 0 / 5 | **5 / 5** (document start) | 39 (three tables) | 323 / 325 ¹ |
| **`wordto` total** | | | **39 / 116** | **116 / 116** | 97 | 6,906 / 6,918 |
| `filesamples/sample1.docx` | 12 / **none (12)** | 9 | 14 / 66 | 30 / 66 (pageBreakBefore); **46 / 66** (+ label ascent only) | 145 | 2,008 / 2,095 ² |
| `filesamples/sample2.docx` | Mac 15 / 14 | 1 | 6 / 10 | **10 / 10** (autospacing) | 1 | 454 / 454 |
| `filesamples/sample3.docx` | 14 / 14 | 2 | **26 / 26** | 26 / 26 | 33 | 1,950 / 1,951 ³ |
| `filesamples/sample4.docx` | 15 / **15** | **175, every break natural** | **5,551 / 5,551** | 5,551 / 5,551 | 700 (47 images) | **496,104 / 496,104** |
| **`filesamples` total** | | | **5,597 / 5,653** | **5,633 / 5,653** | 879 | 500,516 / 500,604 |
| **All nine** | | | **5,636 / 5,769** | **5,749 / 5,769** | | |
| Earlier real-world (`samplelib`), unchanged | | | 495 / 495 | 495 / 495 | 42 | 29,017 / 29,017 |

¹ The glyph checker loses its place in table rows (out of scope) whose first-column
cell is Calibri Bold. That cell's baseline is drawn 1 px below the row's Cambria cells,
so the row arrives as two lines in the wrong order. These are not cascade errors.
² 12 superscript and subscript glyphs, and 75 table-cell glyphs (finding 6). *Since
finding 5, the 12 are drawn at the resolved size: 2,020 / 2,095.*
³ Two text columns interleave (section with `w:cols`; below a table, out of scope).

*Since findings 1–4 were adopted, the model itself scores the diagnostic column,
5,749 / 5,769, and the switch is gone; see "Findings 1–4, probed and adopted". Since
finding 5, 5,755 / 5,769 (`sample1` 52 / 66); see "Superscripts, subscripts, position
and run borders — measured".*

**Font substitution: none, so every document was scored.** The oracle's embedded faces
were checked for each document. `sample2` names Helvetica Neue and Word drew macOS's
HelveticaNeue; Times New Roman is only in its East Asian and complex-script slots and
draws nothing. `sample3` names Tahoma only in styles no text uses; everything is drawn in
Calibri, Arial and SymbolMT. `sample1` names Ubuntu, which this machine lacks, and **Word
drew Ubuntu anyway: the document embeds it** (`w:embedRegular` and its siblings: Ubuntu
in four styles, Ubuntu Mono, and Tahoma). This is not substitution. The vertical metrics
come from the embedded parts (`face_metrics.embedded`, which removes ECMA-376 17.8.1's
32-byte obfuscation in memory; only the four integers leave), and every Ubuntu line
whose other properties are understood is exact. **Word 16.106 lays out with a face
embedded in the document when none is installed.** Whether an installed face takes
precedence over an embedded one was not tested; nothing here has both.

**Natural pagination: the vertical model holds across natural page tops; this is not
pagination.** Which line starts each page still comes from the oracle. `sample4` has 175
pages and one manual break. Of the 167 page tops that are natural breaks and carry text,
**104 continue a paragraph from the previous page**, which is new: every earlier fixture
broke only between paragraphs. 45 start a paragraph with no space before, and 18 are
out of scope (below an image); six pages hold only an image. All 149 in-scope tops are
exact. `sample-10pages` has two natural breaks,
and both pages are exact. Page 3 starts with a Heading 1 whose 24 pt space before Word
**drops** at the natural break, as the model does after a manual one. That is the first
measurement of the open question *space before at a page top after a natural break*: 1
case, dropped.

### What the nine documents exposed

1. **The document's first paragraph keeps its space before** (cause: **the vertical
   model's page-top rule**, which drops space before at every page top). This accounts
   for all 77 `wordto` misses. Each document opens with a Heading 1 (24 pt before), and
   every line of page 1 is exactly 100 px (24 pt) high. With the space kept in the first
   line's box (`hypotheses={"document start"}`), all 116 are exact. No earlier fixture
   could see it, because `samplelib` and the probes open with a Title or a plain line,
   which have no space before. **Probe:** a `before` sweep (1–120 twips, as in
   `make_line_box_probe.py`) on the *first* paragraph of a document, in compatibility
   modes 12, 14 and 15, and in a second section starting on a new page.
   **Adopted, refined:** the first paragraph of every *section* keeps it, in every
   compatibility setting. See *The page top — measured*.
2. **`w:pageBreakBefore` keeps the space before in modes 12 and 14, and drops it in
   15** (cause: **compatibility settings**). Every Heading 1 of `sample1` has
   `w:pageBreakBefore` and 24 pt before. In the original file (mode 12) Word keeps it,
   and five pages are 100 px high. The same file with only `compatibilityMode` set to 14
   keeps it too. Set to 15, it is dropped and the model's page tops are right. The earlier
   *manual break* observation (`samplelib`, mode 14: dropped) therefore does not extend
   to `pageBreakBefore` in mode 14. **Probe:** the same first-paragraph `before` sweep on
   paragraphs that start a page by `w:pageBreakBefore`, by a `w:br w:type="page"` in the
   previous paragraph, by a break alone in its own paragraph, and naturally, crossed with
   modes 12/14/15 and `w:suppressSpBfAfterPgBrk`.
   **Adopted:** kept below mode 15 and when no mode is stated, dropped in 15;
   `w:suppressSpBfAfterPgBrk` changes nothing. See *The page top — measured*.
3. **`w:beforeAutospacing` / `w:afterAutospacing` replace the stated spacing with 14 pt**
   (cause: **the cascade reads a property the model ignores**). `sample2`'s HTML-style
   heading states before = after = 100 twips with both autospacing flags. The model uses
   100 and misses the heading (−11 px) and the paragraph after it (−49 px). Reading both
   as 280 twips (`hypotheses={"autospacing"}`) makes all 10 exact. **Probe:** autospacing
   on/off × stated value × a neighbour's spacing, including the first paragraph and
   list items (HTML suppresses it between list items; Word's rule there is unknown).
   **Adopted, refined:** 14 pt at every size and line rule, but nothing between two
   list items or before the document's first paragraph, and the stated value under
   `w:doNotUseHTMLParagraphAutoSpacing`. See *Autospacing — measured*.
4. **A list label's descent does not enter the line height**. Its ascent does (cause:
   **the vertical model**). `sample1`'s bullets are SymbolMT (descent 450/2048 em),
   Courier New and Wingdings under Ubuntu (descent 0.189 em). The model makes each
   bullet line ~1 px too tall (69.87 px predicted; Word draws 69 between bullets).
   With the label's descent dropped (`"label ascent only"`), `sample1`'s list page goes
   from 0 to 16 exact out of 16 in scope. The earlier documents cannot tell the two
   rules apart, because Cambria's descent exceeds SymbolMT's (`make_mixed_line_probe.py`
   used Cambria), and they stay exact either way. The ArialMT that Word draws for the tab
   after every label here likewise adds no height. **Probe:** the mixed-line probe with
   label faces whose descent exceeds the text's (SymbolMT, Courier New, Wingdings, Arial
   labels under Ubuntu-like low-descent text, and under Calibri), single and `auto` 276.
   **Adopted:** in every setting, for labels taller than the text too. See *A list
   label's descent — measured*.
5. **Superscript and subscript runs, and a run border, make a line taller** (cause:
   **the vertical model**; the size half is **the cascade**). In `sample1`, the line
   with `w:vertAlign` superscript and subscript is followed by a pitch of 69 px, where
   66 is predicted. The next line, holding a run with `w:bdr` (a "box"), is 68 against
   66. Both still hold in modes 14 and 15. Every miss left on `sample1`'s page 2 (19
   lines) is below these two lines and carries their +5 px. The model sets neither the
   raised or lowered extent nor the run border in the line box. Separately, the
   resolved size ignores `w:vertAlign`: Word draws those glyphs at 2/3 size (33 px for
   12 pt), so 12 glyphs disagree. **Probe:** the mixed-line probe with a superscript or
   subscript run, sizes swept (and `w:position`), and with a `w:bdr` run, `w:sz` and
   `w:space` swept like `border`.
   **Probed; half refuted, half adopted.** A superscript or subscript does *not* make its
   line taller: it takes part at its own `w:sz`, on the baseline. The +3 px pitch after
   the script line was the next line's ascent, grown by the box's border, which *is* a
   rule, as `w:position` is. Word draws the glyphs at the face's own OS/2 script size
   (Ubuntu's 0.65, which rounds to 2/3 at 12 pt), not at 2/3. See *Superscripts,
   subscripts, position and run borders — measured*.
6. **In modes 12 and 14, text in a plain table style is not drawn at Normal's size**
   (cause: **cascade × compatibility**, in table cells, so out of scope for baselines).
   `sample1`'s Normal is 12 pt, its docDefaults 11 pt, and its `TableGrid` states no
   size. Cells are drawn at 46 px (and some at 42 px) where the cascade resolves 50 px,
   in the original (mode 12) and with mode 14. With mode 15 they are 50 px. This is the
   *overrideTableStyleFontSizeAndJustification* question the style work left open, now
   observed. **Probe:** `make_style_probe.py`'s p06–p08 extended with Normal ≠
   docDefaults and a table style with and without `w:sz`, in modes 12/14/15, with and
   without `overrideTableStyleFontSizeAndJustification`.
   **Settled** -- see *Tables -- measured*, stage 2: in modes below 15 without the
   override, a paragraph whose styles give it 12 pt takes the table style's size, or
   `w:docDefaults`' unless that is 10 pt; `sample1`'s cells are 11 pt, as Word drew them.
7. **Mode 12 lays out an embedded Ubuntu Mono differently** (cause: **compatibility
   settings**). In the original `sample1`, the three lines holding Ubuntu Mono are 67,
   71 and 69 px apart where 66 is predicted, and the paragraph breaks at different words.
   With mode 14 or 15 set, the breaks move and the lines are 66 apart, so exact but for
   finding 5's carry. Whether this is the face (monospaced, embedded, 1.0 em
   natural height) or mode 12's metric source is not decided. **Probe:** one paragraph
   of Ubuntu Mono (embedded) and of Courier New (installed), mixed with a proportional
   face, sizes swept, in modes 12 and 14.
   *Since finding 5, two more mode-12-only lines belong here:* the second line of the
   formatting paragraph (drawn 836, predicted 837) and the second line of the next one
   (1039 against 1040). Neither is explained by the scripts (the probe's mode 12 script
   lines are all exact, and 1039 holds none). Each follows a line holding Ubuntu's
   bold or italic styles, which points to the same question: which metrics mode 12 takes
   from an embedded face. The probe above, with Ubuntu's four styles, would decide it.
8. **Compatibility variants, for attribution.** `sample1` was also exported with only
   `compatibilityMode` changed (scratch copies, not committed): the model scores 14 / 66
   with mode 14 and 29 / 64 with mode 15 (the pageBreakBefore pages become exact); with
   the pageBreakBefore and label hypotheses, 47 / 66 with mode 14, and with the label
   hypothesis alone 45 / 64 with mode 15. What is left in every mode is finding 5.
   *With findings 1–4 adopted, the model itself scores those: 46 / 66 (mode 12), 47 /
   66 (mode 14), 45 / 64 (mode 15). With finding 5 as well: 52 / 66, **66 / 66** and
   **64 / 64**. What is left is mode 12's alone (finding 7).*
9. **The oracle side needed four repairs, none a model change.** A page holding only an
   image has no `/Font` resource (`quartz_pdf.read` failed on `sample4`). Quartz draws a
   raised run on its own baseline, so the run is folded back into its line by size and
   distance (`baselines.merge_raised`). A table row is drawn as one line across cells,
   so a table whose cells wrap is given every line up to the next paragraph. And text
   that no body paragraph holds (footnotes, a TOC field's result, a drop cap's frame) is
   skipped by looking ahead for the next paragraph's first line, with a partial match
   rolled back (`baselines.match_lines`). Field-instruction text is still parsed as run
   text (`parse/document.py` keeps `w:instrText`), and such paragraphs are out of scope
   (`baselines.paragraph_features`). The parser should drop it when fields are
   implemented. *Dropped since* Headers, footers and fields — measured *(H.3); the
   scorers still leave field paragraphs out of scope.* The six committed earlier documents' lines, scores and glyph counts are
   byte-for-byte unchanged by all of this.

The hypotheses in findings 1–4 were in `baselines.predict(hypotheses=...)` as
diagnostics, **off** in the model. Together they changed no exact baseline of any
committed document (all 736 committed in-scope baselines were exact under them). They
were recorded, not adopted: each needed its probe first, because each document shows
only one side of each rule. The next section adopts them one at a time, each on its own
probe.

---

## Findings 1–4, probed and adopted (Suggested order step 3¾)

The four hypotheses above were found by switching rules on and watching the documents
that suggested them go exact. That is how a hypothesis is found, not how one is tested.
So each got a generated probe that isolates it, was adopted **only** where the probe
agreed, went into the model (`src/docx2svg/`, not `tools/`), and was committed alone,
with a before/after table across every corpus. Probes that decide the page stack are
scored **through `baselines.predict`** (`tools/probe_documents.py`), so they test the
model's own code, not a copy of the rule in a reader.

### The page top — measured (findings 1 and 2)

`make_page_top_probe.py` / `read_page_top_probe.py`: 56 documents, 2,744 pages. In each,
a paragraph with space before starts a page in one of five ways: the **document's first
paragraph** (8 values, 1–480 twips, one per document); the paragraph after a **manual
break** alone in its paragraph; a **`w:pageBreakBefore`** paragraph; the first paragraph
of a **`nextPage` section** that is not the first; and whichever of 130 fillers, each
with space before, Word's own pagination puts at a page top (**natural**). Every kind
but the first is swept 1–120 twips. Seven settings: no `settings.xml`, one without
`w:compat`, an empty `w:compat` (what Word 12 writes), modes 12, 14 and 15 stated, and
mode 14 with `w:suppressSpBfAfterPgBrk`. Calibri 11 pt, one line per paragraph.
`tests/test_page_top.py` holds the model to every line.

**Finding 1, adopted and refined: the first paragraph of a *section* keeps its space
before**, in every setting. The documents could not separate "first in the document"
from "first in a section"; the probe does, and the section start keeps it exactly as the
document start does. It is kept in full and sits in the first line's box, as between two
paragraphs (`vertical.keeps_space_before_at_page_top`, `vertical.page_top_gap_px`).
After a manual break or a natural break it is still dropped.

| Page tops, every line of the page | Dropped everywhere (the old model) | First of a section keeps it |
| --- | --- | --- |
| document start (7 settings × 8 values) | 28 / 224 | **224 / 224** |
| section start (7 × 120 values) | 49 / 4,753 | **4,753 / 4,753** |
| manual break | 1,631 / 1,680 | 1,631 / 1,680 |
| natural | 4,984 / 4,991 | 4,984 / 4,991 |
| `pageBreakBefore` | 395 / 2,520 | 395 / 2,520 (finding 2) |
| `wordto` | 39 / 116 | **116 / 116** |
| every other corpus | unchanged | unchanged; no exact baseline moved |

Under the old rule every section start was low by its space before, rounded (1 px at 3
twips, 100 px at 480); only spaces under half a pixel came out exact.

**A dropped space before is still in the box** (new, from the same probe; committed on
its own). Where the space is dropped, 7 page tops per setting after a manual break, 7
`pageBreakBefore` tops in mode 15 and one natural top were 1 px *high*: Calibri 11 pt
drawn at 343 where every other page top is at 344. They are exactly the spaces of 12,
17, 36, 41, 65, 89 and 113 twips, and the rule that reproduces them is that the first
line's box still holds the space before, **hanging above the top margin**: it moves
nothing, but the baseline rounds in a box whose top is `margin − space`, so the space
decides which edge the line rounds against. 12 twips is 2.50041 px after the layout
unit, so the box top 297.4996 rounds down and the line lands at 343; 60 twips is
exactly 12.5 px, the top 287.5 rounds up, and the line stays at 344, as observed.
(`vertical.page_top_gap_px` returns the space for the box whether or not it is kept.)

| Page tops, every line of the page | Dropped space outside the box | Dropped space in the box |
| --- | --- | --- |
| manual break (7 × 120 values) | 1,631 / 1,680 | **1,680 / 1,680** |
| natural | 4,984 / 4,991 | **4,991 / 4,991** |
| `pageBreakBefore`, mode 15 | 353 / 360 | **360 / 360** |
| every other corpus | unchanged | unchanged; no exact baseline moved |

`samplelib`'s 17 chapter tops drop 480 twips, exactly 100 px, which rounds the same
either way; that is why the documents never showed it.

**Finding 2, adopted: `w:pageBreakBefore` keeps its space before below mode 15, and
drops it in 15.** Exactly as `sample1` suggested, and the probe adds what the one
document could not: **an unstated mode behaves as 12 and 14 do** (no `settings.xml`, a
settings part without `w:compat`, and Word 12's empty `w:compat` all keep it), and
`w:suppressSpBfAfterPgBrk` (in mode 14) changes nothing, for `pageBreakBefore` or for a
manual break. `docx2svg` now reads `compatibilityMode` (`Document.compatibility_mode`,
`None` when unstated) and `keeps_space_before_at_page_top` takes it. A manual break
drops the space in every setting, mode 12 included.

| `pageBreakBefore` page tops, every line | Dropped everywhere | Kept below 15 and unstated |
| --- | --- | --- |
| none / no `w:compat` / empty `w:compat` (3 × 120 values) | 21 / 1,080 | **1,080 / 1,080** |
| mode 12, 14, 14 + `suppressSpBfAfterPgBrk` (3 × 120) | 21 / 1,080 | **1,080 / 1,080** |
| mode 15 (120) | **360 / 360** | **360 / 360** |
| `filesamples/sample1` (mode 12, not committed) | 14 / 66 | **30 / 66** |
| `sample1` with only the mode set to 14 / 15 (scratch) | 14 / 66, 29 / 64 | **31 / 66**, 29 / 64 |
| every other corpus | unchanged | unchanged; no exact baseline moved |

Keeping it in mode 15 as well, or dropping it in every mode, each leaves 353 of the
other side's 360 lines per setting off (`tests/test_page_top.py`).

### Autospacing — measured (finding 3)

`make_autospacing_probe.py` / `read_autospacing_probe.py`: one document per setting (no
`settings.xml`; modes 14 and 15; mode 15 with `w:doNotUseHTMLParagraphAutoSpacing`), 90
pages, 317 lines each, every page led by a plain `pageBreakBefore` anchor. Each flag
alone with the stated value 0, 100, 400 or 1000 twips, at 8, 11, 20 and 36 pt, under
single, `auto` 1.5 and `exact` 40 pt; beside a neighbour's 0–1000 twips; three
consecutive autospaced paragraphs; three autospaced list items; an autospaced paragraph,
an item of one list, an item of another, an autospaced paragraph; list items with stated
spacing only; plain collapses; `w:contextualSpacing`; and the flag on the document's
first paragraph, a section's first, after a manual break and with `pageBreakBefore`.
List labels are Calibri, so they add nothing a line of text does not.
`tests/test_autospacing.py`.

**Finding 3, adopted and refined.** In the three settings that use it:

* **An autospaced side is 14 pt (280 twips), whatever is stated**, at every size and
  line rule tried: it does not scale with the text or the line. It then collapses with
  its neighbour exactly as a stated 280 would, and a section start or a
  `pageBreakBefore` paragraph (below mode 15) keeps it at a page top like any other
  space before. This is the cascade's: `resolve_paragraph` resolves `spacing.before` /
  `spacing.after` to 280, with the flag's origin and the reason
  (`resolve.cascade._autospacing`).
* **It is not there between two list items**, of the same list or of two lists; next
  to a paragraph that is not a list item it is. The first item's space after is out of
  its box too, which moves it by 1 px. Stated spacing between list items is not
  affected.
* **It is not there before the document's first paragraph** (drawn with no space, where
  the stated 100 twips would have given 21 px). A section's first paragraph keeps it.
  Both are about the neighbour, so they live in the stack: `vertical.autospace_kept`,
  with `resolve.autospaced` saying whether a side is autospaced at all.
* **Under `w:doNotUseHTMLParagraphAutoSpacing` the flags mean nothing**: the stated
  value holds (`stated` 240 / 240 on stated values), including between list items and
  at the document start.

| Lines, three settings together | Stated value (the old model) | 14 pt, no neighbour rule | The model |
| --- | --- | --- | --- |
| stated value × size × rule | 144 / 720 | 720 / 720 | **720 / 720** |
| neighbours 0–1000 twips | 84 / 105 | 105 / 105 | **105 / 105** |
| consecutive; contextual; section start | 6 / 39 | 39 / 39 | **39 / 39** |
| list items; list edges | 6 / 30 | 12 / 30 | **30 / 30** |
| document start | 0 / 6 | 0 / 6 | **6 / 6** |
| stated lists; plain collapses; manual break; `pageBreakBefore` | 35 / 51 | 51 / 51 | **51 / 51** |
| **total** | 275 / 951 | 927 / 951 | **951 / 951** |
| `filesamples/sample2` (not committed) | 6 / 10 | 10 / 10 | **10 / 10** |
| every other corpus | | | unchanged; no exact baseline moved |

**New, not adopted: `w:doNotUseHTMLParagraphAutoSpacing` makes every paragraph's space
after and the next one's space before *add*.** Not just autospaced ones: two plain
paragraphs with 100 after and 100 before are 200 apart, and so are stated list items.
With `collapse="sum"` every one of that document's 317 lines is exact (302 collapsing);
in the other settings adding misses 33. So the collapse to the larger that *Style
inheritance* measured is itself "HTML paragraph auto spacing", and this option turns it
off. It is outside the four findings, so the model does not read it yet; it is pinned
as a diagnostic in `tests/test_autospacing.py`. How it combines with contextual spacing
and page tops is not measured.

### A list label's descent — measured (finding 4)

`make_label_probe.py` / `read_label_probe.py`: three documents (no `settings.xml`, mode
12 as `sample1`, mode 15), 64 pages each. One group per page, 22 one-line paragraphs of
11 pt text in Tahoma, Arial, Cambria or Calibri (descent 0.207, 0.212, 0.222, 0.269 em),
the mark in the text's face and size, under single spacing and `auto` 276. Six labels:
SymbolMT and Wingdings bullets (descent 0.220 and 0.211: above some texts' and below
others'), Courier New `o` (0.300, with a *smaller* ascent than every text), Arial Black
`•` (0.310, ascent 1.101), and Courier New and SymbolMT at 20 pt — **labels taller than
the text** above and below. Two controls: the Courier New `o` at 11 and 20 pt as inline
*runs* of the text. `tests/test_label.py`.

**Finding 4, adopted as stated: a list label's ascent takes part in the line's extent,
its descent does not**, in every setting, at both sizes, under both line rules. A
Courier New label before Arial adds nothing at all (52.667 px drawn, Arial's own 52.70),
because its ascent is below Arial's and its descent is left out; Courier New 20 pt
before Tahoma gives the label's ascent over Tahoma's descent (78.857 drawn, 78.84 so,
94.38 with the label's descent). An `auto` multiple still adds its extra over the text
alone, on top of that extent. The same glyphs as inline runs count with their descent,
as text does. The rule is `vertical.line_extent(text_items, metrics, label_items)`;
the line-extent functions (`line_extent`, `tallest_natural`, `mixed_line_pitch`) moved
from `tools/baselines.py` into `docx2svg.vertical`, and `read_mixed_line_probe.py`
calls them too.

| Lines, three settings | Label counts with its descent (the old model) | Ascent only |
| --- | --- | --- |
| SymbolMT label | 288 / 528 | **528 / 528** |
| Wingdings label | 426 / 528 | **528 / 528** |
| Courier New label | 18 / 528 | **528 / 528** |
| Arial Black label | 18 / 528 | **528 / 528** |
| Courier New 20 pt label | 12 / 528 | **528 / 528** |
| SymbolMT 20 pt label | 21 / 528 | **528 / 528** |
| inline Courier New, 11 and 20 pt (controls) | 1,056 / 1,056 | 1,056 / 1,056 |
| mixed-line probe (Cambria under SymbolMT labels) | 440 / 440 | 440 / 440 |
| `filesamples/sample1` (mode 12, not committed) | 30 / 66 | **46 / 66** |
| every other corpus | | unchanged; no exact baseline moved |

The old model was right only where the text's descent is at least the label's, which is
every earlier fixture: Cambria (455/2048 em) under SymbolMT (450), and Calibri.

### What is left, and the diagnostic switch

`baselines.predict(hypotheses=...)` is **gone**: all four of its switches are now rules
of the model, and findings 5–7 never had one. There is no diagnostic path that differs
from the model; `tools/` calls `docx2svg` for every rule above. What remained of the nine
documents was findings 5–7: `filesamples/sample1` 46 / 66, every miss below its
super/subscript and bordered-run lines (finding 5) or on mode 12's Ubuntu Mono lines
(finding 7); table cells (finding 6) stay out of scope. Finding 5 has since had its
probe (next section). Two findings came out of
the probes themselves. A dropped space before stays in its line's box: adopted in a
commit of its own, because it is the other half of finding 1's page-top rule and moved
no exact baseline. `w:doNotUseHTMLParagraphAutoSpacing` makes paragraph spacing add:
recorded, not adopted.

| Corpus | Before findings 1–4 | After |
| --- | --- | --- |
| Phase 2 probes | 23,859 / 23,859 | 23,859 / 23,859 |
| Line-box probe | 12,069 / 12,069 | 12,069 / 12,069 |
| Mixed-line probe | 440 / 440 | 440 / 440 |
| **Page-top probe** (new) | 7,087 / 14,168 | **14,168 / 14,168** |
| **Autospacing probe** (new) | 578 / 1,268 | **1,253 / 1,268** (the 15: `doNotUseHTMLParagraphAutoSpacing`) |
| **Label probe** (new) | 1,839 / 4,224 | **4,224 / 4,224** |
| `layout-sweep.docx`, `style-document.docx` | 125 / 125 | 125 / 125 |
| `samplelib` | 495 / 495 | 495 / 495 |
| `wordto` | 39 / 116 | **116 / 116** |
| `filesamples` (not committed) | 5,597 / 5,653 | **5,633 / 5,653** |
| **Total** | 52,128 / 62,417 | **62,382 / 62,417** |

## Superscripts, subscripts, position and run borders — measured (finding 5)

`sample1` suggested that a `w:vertAlign` run makes its line ~3 px taller and a run with
`w:bdr` ~2 px, and that Word draws the script at 2/3 size. A probe that isolates each
variable **refutes the first and the third and confirms the second**, and adds
`w:position`:

* **A superscript or subscript takes part in its line at its own `w:sz`, on the
  baseline**: neither reduced nor raised nor lowered, whatever it looks like. In the
  text's size, all 5,184 such lines lie exactly on the control lines' baselines. The
  model already counted it so. `sample1`'s +3 px pitch after the script line was the
  *next* line's ascent, grown by the box's border.
* **A run border grows its run** by its space and width, above and below; the line
  takes the run's extent like any other, so a small bordered run grows it only where
  its border passes the text. An `auto` multiple scales the border too.
* **`w:position` moves its run's extent**: raised by `p` half points, it reaches `p`
  higher and `p` less far down. The multiple's extra is taken without it.
* **The glyphs are drawn at the face's OS/2 script size**, not at 2/3: `w:sz` ×
  `ySuperscriptYSize` / `unitsPerEm`, rounded half up to a half point. Ubuntu's is 0.65,
  and 0.65 × 24 = 15.6 rounds to 16, which is where "2/3" came from. That is the
  cascade's (`resolve.script_half_points`) and moves no line.

So nothing about a script run is in the line height, and the model's line rule for
scripts is unchanged. What moved lines is the border rule (and `sample1`'s page 2 with
it) and the position rule (the probe's).

### The probe

`make_script_probe.py` / `read_script_probe.py`, scored through `baselines.predict`.
One group per page, 12 one-line paragraphs, `Hxample n ` + the extra run (`Hxh`) + ` Hxh`
in the text's face and size, the mark likewise. Six faces with different ascent/descent
ratios (Calibri, Aptos, Times New Roman, Arial, Cambria, Courier New), four line rules
(`auto` 240 and 276, `exact` 360, `atLeast` 300), three settings (no `settings.xml`,
mode 12, mode 15). **Word drew the three settings identically, line for line**, so the
recording keeps modes 12 and 15 as references to the first and the tests score one.
`tests/test_script.py`.

| Document | What varies | Groups × 12 lines |
| --- | --- | --- |
| `script-*` | superscript, subscript, both, in the text's size; each alone twice and half the text's size; control; text at 11 and 20 pt | 384 |
| `position-*` | `w:position` ±1, 2, 3, 6, 12, 24, 48 half points; ±6..24 on a run half the text's size, ±6 on one twice it | 384 |
| `border-*` | `w:sz` 2, 4, 6, 8, 12, 18, 24, 27, 48, 96 at space 0; space 1, 2, 4, 8, 31 at `w:sz` 4; 24 × 12; four borders on runs half and twice the text's size | 432 |
| `script-sizes` | one line per `w:sz` 2..96 in twelve faces, a superscript and a subscript each, Word's pagination | 1,140 lines |

Superscript runs are drawn on baselines of their own; the reader folds them back into
their paragraph by rank on the page (the `k`-th extra run from the top belongs to the
`k // n`-th paragraph). Nearest does not work: a run raised by 24 pt under `exact` 18 pt
is drawn nearer the paragraph above. The drawn size comes from the run's advance (the
pen x of the next text object, exact to 1e-4 px, over the run's advance in font units);
the offset from its baseline.

### The rules, and what they refute

| Rule | Score | Refuted alternatives, same lines |
| --- | --- | --- |
| `vertAlign` at `w:sz`, on the baseline (`vertical.item_extent`) | 13,824 / 13,824 | at the drawn size: 11,232 (every script twice the text's size fails but under `exact`); "+3 px": 0 of the 5,184 same-size lines is taller than its control |
| `w:position` moves the run's extent (`item_extent`) | 13,824 / 13,824 | ignored (the model before): 4,167; also in the multiple's natural height: 11,130; a raised run keeping its descent: 13,404 |
| `w:bdr`: space + width above and below the run, in its natural height (`item_extent`, `item_natural_px`, `run_border_px`) | 15,552 / 15,552 | ignored: 3,702; outside the multiple's natural height: 12,399 (`auto` 276: 303 / 3,456); around the whole line: 0 of 864 half-size-run lines; width with its half twip: 15,039; width rounded to the unit, as `border_px`: 15,543 |
| drawn size: `w:sz` × OS/2 script size / upm, half up, ≥ 2; 3/5 when absurd (`resolve.script_half_points`) | 1,133 / 1,140 sizes | `w:sz`: 12; 2/3: 192; the OS/2 ratio rounded to a percent: 1,085, per mille: 1,113, 1/256: 1,071, 1/64: 790 |

The border width is truncated to whole twips, as a paragraph border's, and then
*truncated* to the layout unit (a paragraph border's is rounded): `w:sz` 27 is 67 twips,
13,721.6 units, and three lines of 432 sit at a tie that only 13,721 reproduces.

### Constants, observations, residuals

| Constant | Value | Evidence | Residual |
| --- | --- | --- | --- |
| Script size | `round_half_up(w:sz × ySuperscriptYSize / upm)`, min 2 | 1,133 / 1,140 over twelve faces at 0.416–0.65 | `w:sz` 90 in the six 0.65 faces is drawn at 59 (58.49), Baskerville Old Face 6 at 4 (2.49). Nothing tried reproduces 90 without breaking 10, 30, 50, 70 (x.499, drawn down) |
| `SCRIPT_SIZE_FALLBACK` | 3/5 | Helvetica Neue (0.204) and Galvji (1.228) at every size: feasible 0.5990–0.6011 | — |
| `SCRIPT_SIZE_RANGE` | [1/4, 1] | taken at 0.416, 0.528, 0.600, 0.650; not at 0.204, 1.228 | **bounds not measured**: lower in (0.204, 0.416], upper in [0.650, 1.228) |
| Subscript size | `ySubscriptYSize` | every face measured has it equal to `ySuperscriptYSize`, and every line draws both runs at one size | which field Word reads is not separated |
| Script offset | whole half points | all 2,280 drawn offsets are `round(k × 25/12)` px | **not modelled**; see below |

**The offset is recorded, not settled.** Word moves a script run by whole half points,
as `w:position` does, but by how many is not one rule. For faces whose
`ySuperscriptYOffset` is at most 0.39 em it is `round(w:sz × ySuperscriptYOffset / upm)`
in nearly every size (Cambria 95, Georgia 94, Impact 94, Aptos 91 of 95), and
`ySubscriptYOffset` likewise for the subscript (Aptos 95 / 95; Impact and Baskerville,
whose offset is 0, 95 / 95). For Calibri, Times New Roman, Arial and Courier New (0.42–
0.48 em) it is 2–6 of 95; they are raised by about a third of the size, not
monotonically (Calibri: 16 half points at `w:sz` 48 and 52, 19 at 56). Neither a linear
function of the size, of the drawn size, nor the difference of the two sizes' ascents
(in points, pixels, half points, twips, at the exact size or at whole ppem) fits every
face. No baseline depends on it; Phase 5 will. **Probe that would decide it:** the
size sweep with faces whose OS/2 offsets are edited in a font of our own (so one field
moves at a time), and a script beside a taller run, to see whether the line's extent
bounds the raise. *Measured in Phase 5 with faces of this project's own: whole half points; the
face's offset, but no more than the difference of the two sizes' ascents (descents);
fields refused below about a fifth of an em or with a negative subscript offset. 545 and
445 of 690 probe scripts exact, 661 and 670 within a half point; see "Phase 5 —
measured", 5.5.*

### Also found, not adopted

* **Baskerville Old Face is drawn taller than its `hhea` metrics.** Its size-sweep
  lines are 1–380 px low under the model (1536 + 512 per 2048, no line gap, no
  `USE_TYPO_METRICS`); its `usWin` metrics are 1805 + 531. It is not simply "Word uses
  `usWin`": Galvji and Charter also differ between the two and are drawn exactly by
  `hhea`. Its `hhea` sums to exactly one em, which may be what Word distrusts. It
  accounts for 135 of the size sweep's 136 misses (the Impact lines after it on the same
  page carry it).
  **Probe:** the Phase 2 single-line pitch probe in faces whose `hhea` and `usWin`
  differ, with and without `USE_TYPO_METRICS`.
* **One Helvetica Neue line (`w:sz` 80) is 1 px off** in the size sweep, cause unknown.
* **Which line of a paragraph its mark and its list label belong to** is not probed:
  each line now holds the characters Word put on it (`baselines.line_shares`), but the
  mark and the label still count on every line of the paragraph, as before. No corpus
  can tell yet. **Probe:** three-line paragraphs whose mark, or whose label, is twice
  the text's size. *Since probed: the mark counts on no line that holds text, and the
  label on the first line only; see "The paragraph mark in the line height — measured".*
* `face_metrics` finds "Aptos" in `Aptos-Light.ttf` (its typographic family is Aptos);
  its integers are the regular face's, so nothing moved.

### Scores

| Corpus | Before finding 5 | After |
| --- | --- | --- |
| Phase 2 probes | 23,859 / 23,859 | 23,859 / 23,859 |
| Line-box probe | 12,069 / 12,069 | 12,069 / 12,069 |
| Mixed-line probe | 440 / 440 | 440 / 440 |
| Page-top probe | 14,168 / 14,168 | 14,168 / 14,168 |
| Autospacing probe | 1,253 / 1,268 | 1,253 / 1,268 |
| Label probe | 4,224 / 4,224 | 4,224 / 4,224 |
| **Script probe: `vertAlign`** (new) | 13,824 / 13,824 | 13,824 / 13,824 |
| **Script probe: `w:position`** (new) | 4,167 / 13,824 | **13,824 / 13,824** |
| **Script probe: `w:bdr`** (new) | 3,702 / 15,552 | **15,552 / 15,552** |
| Script probe: size sweep's baselines (new) | 1,004 / 1,140 | 1,004 / 1,140 (Baskerville Old Face, above) |
| `layout-sweep.docx`, `style-document.docx` | 125 / 125 | 125 / 125 |
| `samplelib` | 495 / 495 | 495 / 495 |
| `wordto` | 116 / 116 | 116 / 116 |
| `filesamples` (not committed) | 5,633 / 5,653 | **5,639 / 5,653** (`sample1` 52 / 66) |
| **Total** | 85,079 / 106,757 | **106,592 / 106,757** |
| *Of the corpora before this probe* | *62,382 / 62,417* | ***62,388 / 62,417*** |
| `filesamples/sample1` glyphs in the resolved size | 2,008 / 2,095 | **2,020 / 2,095** (the rest: table cells, finding 6) |

What is left of the 35: 15 autospacing lines (`doNotUseHTMLParagraphAutoSpacing`) and 14
`sample1` lines, all mode 12's (finding 7): its Ubuntu Mono lines, the lines below them
on the page, and the two 1 px lines after Ubuntu's other styles. The same file in mode
14 or 15 is exact (66 / 66, 64 / 64). No currently-exact baseline moved in any corpus,
checked line by line at every commit. Each line now holds the characters Word put on it,
from the oracle (`baselines.line_shares`), which is what let `sample1`'s bordered line be
scored without making the rest of its paragraph taller; that change alone moved nothing.


---

## Local corpora — scored on one machine, recorded nowhere else

Some documents may be measured here but neither redistributed nor named: their terms
allow private use only. `filesamples` keeps its documents out of git but commits their
names, hashes and scores (`tests/test_filesamples.py`), which its source permits. A
**local corpus** keeps everything out: it is a directory under the gitignored `scratch/`
holding the documents, a manifest (`local-corpus.json`: each document's SHA-256 and the
scores it is pinned at, or why it is skipped), Word's recorded lines and any notes.
`tools/local_corpus.py record scratch/<name>` exports and scores every `.docx` there
and writes the manifest and the lines; `tests/test_local_corpora.py` scores every
document every manifest lists through `baselines.predict` and `check_glyphs`, offline,
and skips where there is none. Neither knows any corpus. A document for which Word
substituted a face (one it has neither installed nor embedded) is listed as skipped with
that reason, not scored. What such a corpus exposes that is general is recorded here in
general terms, with a generated probe; its scores and its lines stay beside it.

`filesamples` stays where it is: its test carries per-document commentary and an
embedded-face check that the generic manifest has no place for, and moving it would not
be simpler.

### A paragraph mark larger than its text — observed here, since probed and adopted

A table-heavy template exposed it: a one-line title of 24 pt text whose paragraph mark is
26 pt is drawn where the 24 pt text alone puts it, and the line below follows the text's
pitch. The model put the mark on the line like any other item, so it predicted the line
8 px low and carried 10 px to every line below on the page. Every earlier probe gave the
mark the text's size, and in an empty paragraph the mark is all there is, so nothing
could tell the two apart. A generated probe now settles it, and the rule is the model's:
see *The paragraph mark in the line height — measured*.

The same template's lines are otherwise out of scope, and honestly so: its first
paragraph anchors a floating drawing (`wp:anchor`, `wrapNone`), and the harness takes any
paragraph holding a drawing, and every line below it on its page, out of scope, before
the tables (Phase 6) take the rest. A scratch diagnostic that takes paragraphs whose only
drawings are `wrapNone` anchors back in scope finds, once the mark rule is in, no line of
the local corpora that such an anchor moves. That is a hint, not a measurement: the
anchors there are few, all `behindDoc`, and all in title paragraphs. **Not in scope
yet.** *Since* Floating drawings — measured *(F.1, F.4): in scope, drawn, and every
line of the corpora's paragraphs that anchor one exact.* **Probe:** generated anchored shapes (`wrapNone`, `wrapSquare`,
`wrapTopAndBottom`, behind and in front of the text, of several heights, anchored in the
first, a middle and the last line of a paragraph) in otherwise plain paragraphs.

## The paragraph mark in the line height — measured

**The paragraph mark takes part in a line's height only when the line holds nothing but
spaces** (an empty paragraph, or one of space characters alone), where it alone is the
line. On a line that holds text it takes no part at all: not by its size, not by its
face's ascent or descent, not on the paragraph's last line (where it sits) nor on any
other. **A space (U+0020) takes no part either**, whatever its size or face, alone or
between words; a no-break space (U+00A0) counts as text does. `vertical.line_items`;
`baselines.line_items` calls it. **A list label takes part in its paragraph's first line
only** (`vertical.label_items`), as it did before by its ascent alone (finding 4).

### The probe

`make_mark_probe.py` / `read_mark_probe.py`, scored through `baselines.predict`, in
`tests/test_mark.py`. One case per page: an 11 pt anchor line, the case paragraph, three
11 pt lines, so the case paragraph's baselines and the pitch below it are scored. The
first two documents are the original observation (Calibri, single spacing, no settings
and mode 15). The **sweep** crosses every case with five line rules on the case paragraph
(`auto` 240 and 276, `exact` 720, `atLeast` 640 — between the natural height of 24 pt
text and of a 30 pt mark in every face used — and `atLeast` 300, below both) and four
settings (no `settings.xml`, modes 12, 14, 15). **The four settings draw every line
alike**, so each open question below was asked of all three line rules and every mode.

| Family | Cases (text 24 pt) |
| --- | --- |
| size | six text faces (Calibri, Aptos, Times New Roman, Arial, Cambria, Courier New): mark 24 (control), 30, 16 pt; a paragraph of one 24 pt space whose mark is 30 pt; an empty paragraph whose mark is 30 pt; three-line paragraphs whose mark is 24 or 30 pt |
| face | the mark in a face reaching further at the *same* size: Arial Black under Courier New (+0.268 em above) and Calibri, Palatino Linotype under Times New Roman; deeper: Lucida Calligraphy under Comic Sans MS, Bradley Hand under Arial (+0.187 em below, 0.088 em less above); *smaller* and taller (Arial Black 20 pt under Courier New 24 pt), smaller and deeper (Bradley Hand 20 pt under Arial); shorter both ways (Courier New under Arial Black); a three-line paragraph; an empty paragraph in Arial Black |
| label | numbered paragraphs of Calibri and Courier New text: the label from a 30 pt mark (it takes the mark's properties), a level `w:sz` of 24 pt under a 30 pt mark, a level `w:sz` of 30 pt under a 24 pt mark, one line and three |
| space | Calibri, Times New Roman, Courier New: one 30 pt space under a 24 pt mark, one 26 pt space under a 16 pt mark, one 16 pt space under a 24 pt mark, one 30 pt tab under a 24 pt mark; 24 pt text with a 30 pt space after it, or a 30 pt space or tab inside it; a 24 pt space in Arial Black (0.15–0.27 em taller than the text's face) alone and between two words; a 30 pt no-break space alone and between two words |

### What each open question resolved to

* **Size or extent?** Extent does not matter either: a mark in a face reaching 0.27 em
  higher, or 0.19 em lower, than the text at the *same* size takes no part, and neither
  does a smaller mark that still reaches higher or lower than the text (`face-*`). "Larger"
  was never the condition; the mark is simply not on a line that holds text.
* **Paragraphs longer than one line:** no line. Counting the mark on every line (the model
  before) or on the last line only are both refuted (table below).
* **List labels:** the label still takes the mark's properties — a 30 pt mark makes a
  30 pt label, whose ascent counts on its line as before (finding 4), while the mark
  itself does not; a level `w:sz` of 24 pt under a 30 pt mark leaves the line at the
  text's height. **And the label counts on the first line only** — not a question the
  task asked, but the three-line numbered cases answer it: a 30 pt label, from the
  mark or from the level's `w:sz`, makes the first line taller and leaves the second
  and third at the text's pitch.
* **Smaller marks** take no part in any way, including a smaller mark whose face reaches
  lower than the text (`face-smallerdeep`) or higher (`face-smallertall`).
* **All three line rules, six text faces, four settings:** the same everywhere.
* **Where the mark does count:** an empty paragraph (a 30 pt mark in six faces; a 24 pt
  Arial Black mark) and a paragraph of spaces alone (`size-onlyspace`: a 24 pt space
  under a 30 pt mark is a 30 pt line).

* **Spaces** (found by the `space` family, not asked): a space larger than the text,
  after it or between its words, or of a taller face at the same size, leaves the line
  as tall as the text alone; a paragraph of one 30 pt space under a 24 pt mark is a
  24 pt line, and one 16 pt space under a 24 pt mark too. A no-break space is not a
  space here: 30 pt, alone or between words, it makes a 30 pt line.

### Scores

All on the same recording (the sweep: 10,780 lines over the four settings).

| Rule, same lines | First two documents | Sweep |
| --- | --- | --- |
| the mark on every line of its paragraph, spaces counted (the model before) | 56 / 120 | 7,416 / 10,780 |
| the mark on the paragraph's last line (and on a line of spaces) | 56 / 120 | 7,752 / 10,780 |
| **the mark only on a line of nothing but spaces**, spaces counted | **120 / 120** | 9,312 / 10,780 |
| … and a no-break space left out too | 120 / 120 | 10,044 / 10,780 |
| … and spaces left out at the end of a line only | 120 / 120 | 10,044 / 10,780 |
| … and on a line of spaces alone, spaces counted with the mark | 120 / 120 | 9,888 / 10,780 |
| … **and spaces take no part** | **120 / 120** | 10,428 / 10,780 |
| … **and a list label on its paragraph's first line only** | **120 / 120** | **10,780 / 10,780** |

`tests/test_mark.py` pins these the other way round too: each refuted rule in place of
the adopted one, the other two kept (the mark on every line 7,648 of 10,900 lines; on
the last line 8,036; spaces counted 9,784; a no-break space left out, or spaces left
out at a line's end only, 10,516; spaces counted with the mark on a line of spaces
10,360; the label on every line 10,548).

| Sweep family, four settings | The model before | The mark rule | … and spaces | … and labels |
| --- | --- | --- | --- | --- |
| `size-big` | 216 / 600 | **600 / 600** | 600 / 600 | 600 / 600 |
| `size-multibig` | 268 / 860 | **860 / 860** | 860 / 860 | 860 / 860 |
| `face-tall` | 108 / 300 | **300 / 300** | 300 / 300 | 300 / 300 |
| `face-deep` | 104 / 200 | **200 / 200** | 200 / 200 | 200 / 200 |
| `face-smallertall`, `face-smallerdeep` | 112 / 200 | **200 / 200** | 200 / 200 | 200 / 200 |
| `face-multitall` | 48 / 160 | **160 / 160** | 160 / 160 | 160 / 160 |
| `label-bigmark`, `-bigmark-label24`, `-multibigmark-label24` | 268 / 700 | **700 / 700** | 700 / 700 | 700 / 700 |
| `space-onlybig`, `-onlybig-smallmark`, `-trailingbig`, `-interiorbig`, `-onlytall`, `-interiortall` | 684 / 1,800 | 684 / 1,800 | **1,800 / 1,800** | 1,800 / 1,800 |
| `label-multibigmark`, `-multilabel30` | 248 / 600 | 248 / 600 | 248 / 600 | **600 / 600** |
| controls (`size-same`, `-small`, `-onlyspace`, `-empty`, `-multisame`, `face-short`, `face-empty`, `label-same`, `label-label30`, `space-onlysmall`, `-onlytab`, `-interiortab`, `-onlynbsp`, `-interiornbsp`) | 5,360 / 5,360 | 5,360 / 5,360 | 5,360 / 5,360 | 5,360 / 5,360 |

No rule moved any other corpus: every committed document (736 / 736),
`filesamples` (5,639 / 5,653), the local corpora, and the page-top, autospacing, label
and script probes are line for line what they were; the Phase 2, line-box and
mixed-line probes do not pass through the rule. Each was checked line by line at its
commit.

### Also found

* **A tab takes no part** (`space-onlytab`, `-interiortab`: 300 / 300 each). The parser
  keeps a tab out of a run's text, so the model already left it out.

### Not settled

* Other whitespace — ideographic (U+3000), en, em and other fixed-width spaces — is not
  probed; the model counts it as text, as it does the no-break space, which is measured.
* A paragraph that ends with a line break (`w:br`), whose last line holds only the mark,
  is not probed.

---

## Phase 3 — Line breaking (L)

**The hard problem, and the one nobody in the field has solved in the open.**

Greedy first-fit is the right starting model and is almost certainly what Word does for
Latin text; the sibling measured PowerPoint's budget as exact —
`sum of advances ≤ width − lIns − rIns − marL`, inclusive, with no slack at all — and Word
is likely the same with `w:ind` in place of the insets. **That is a hypothesis, not a
constant**, and the phase begins by measuring it the way the sibling did: a box of
`k · size + δ` and a repeated character, sweeping δ through zero.

**Must be measured before starting:** Phase 2 complete. A line breaker on top of an
advance width that is wrong by 0.1% will break in the wrong place on long lines and be
right on short ones, which is the most confusing possible failure.

**What makes this harder than the sibling's equivalent:** a document's column is fixed but
its *content* is not bounded by a box, so there is no autofit to hide an error; the
breaking must also handle tabs (a jump to a stop mid-line), `w:br`, non-breaking hyphens,
soft hyphens, and hyphenation proper, which the sibling never needed.

**Do not do:** justification, East Asian line breaking, and hyphenation in this phase. Each
is its own measurement and each can be added to a correct greedy breaker. Attempting them
together makes every residual unattributable.

**Done when:** every line of the fixture's block D, and of a wrapping probe of at least 200
lines across five column widths and five font sizes, breaks after the same word as Word's.
Counted in lines that agree, not in a percentage of pixels.

**Done for Latin text: see *Phase 3 — measured* below.** Every line of block D (4 / 4),
of the wrapping probe (1,778 / 1,778 lines in ten faces, five sizes, five columns), of
the budget sweep (7,101 / 7,101) and of the rules probe (920 / 920) breaks after the same
word as Word's; so do 7,164 of the 7,167 lines of the real documents.

---

## Phase 3 — measured

`docx2svg.linebreak` breaks a paragraph into lines; `docx2svg.measure` gives it advance
widths (`TableAdvances`, from `ooxml-common`'s tables, for the faces checked below);
`tools/breaks.py` scores it against Word's lines. A line **agrees** when it starts and
ends where Word's does — "breaks after the same word", which also requires every line
before it in the paragraph to agree. Beside it, `breaks.py` reports the *conditional*
count (from where Word started each line, does the model end it where Word did?), which
equals the other when everything agrees and otherwise counts separate mistakes.

**What the model is:** greedy first fit. Each line takes pieces while the pen position
after the last one is at or before the right edge, and breaks at the last opportunity
that fits; a word that does not fit on an empty line breaks after its last character that
does. Everything is in Word's layout unit, 1/4096 pt, exactly: an advance is
`advance × half points × 2048 / upm` units (an integer for every 2048-upm face at a
half-point size), a twip length rounds half up to the unit (Phase 2, 2.1). Nothing rounds
to the device grid. The line starts at `w:ind/@w:left` (plus `@w:firstLine`, less
`@w:hanging`, on the first line) and ends at the column's width less `@w:right`.

### 3.1 The budget: exact, inclusive, no slack — measured, as the sibling's was

`make_wrap_budget_probe.py` / `read_wrap_budget_probe.py`, `tests/test_linebreak.py`. The
sibling's method, in Word's unit: the line under test is `W1 W2`, a fixed first word, a
space and a second word composed from thirteen letters so that the line ends δ units
short of the edge. Everything authored is a multiple of 5 twips (page 11,900 twips wide,
1440 margins, indents), so the budget is exactly `1024 × twips / 5` units and δ is known
to the unit (0.001 device px). δ ∈ {±1024, ±205, ±41, ±9, ±3, ±1, 0} (even values at
even sizes, where only even differences exist; Courier New, whose letters are all 1229
units, the nearest reachable: −1032 … +1049). Eight faces (Calibri, Arial, Times New
Roman, Cambria, Georgia, Verdana, Courier New, Aptos) at 9.5, 11 and 14.5 pt, columns
1500–9020 twips, seven families, and — for four faces at two sizes, three families —
`settings.xml` at compatibility modes 12, 14 and 15.

| Family | What the line is | Word: held W2 from | Cases agreeing |
| --- | --- | --- | --- |
| `mid` | `W1 W2 Hnnn` | δ = 0 | 312 / 312 |
| `last` | `W1 W2`, then the paragraph mark | δ = 0: **the mark does not count** | 312 / 312 |
| `trail` | `W1 W2`, six spaces, `Hnnn` | δ = 0: **spaces at a line's end do not count**, however many | 312 / 312 |
| `second` | a first line of one word, then `W1 W2 Hnnn` | δ = 0: a later line has the same budget | 312 / 312 |
| `first` | `w:firstLine` 720 | δ = 0, the budget less the first-line indent | 312 / 312 |
| `hang` | `w:left` = `w:hanging` = 720 | δ = 0, the whole column | 312 / 312 |
| `single` | one word as wide as the budget less δ | δ = 0; below, split after its last character that fits | 312 / 312 |
| modes 12, 14, 15 × `mid`, `last`, `trail` | | δ = 0 in every mode | 936 / 936 |
| **All** | | not held at δ = −1, held at δ = 0, in every family and setting | **3,120 / 3,120** (7,101 lines) |

**Refuted:** a strict budget (`x < right`), 2,919 / 3,120 — it fails every δ = 0 case.
Any slack at all is refuted to one unit: no case with δ = −1 held its word. So PowerPoint's
rule transfers as a *measurement*, with `w:ind` in place of the insets and nothing else
changed.

The first run of this probe scored 2,973 / 3,120: every Aptos case held its word only
from δ ≈ +1100. The advances were Aptos **Light**'s — `face_metrics` indexed faces by
typographic family too, and `Aptos-Light.ttf` (typographic family "Aptos", subfamily
"Light") sorts before `Aptos.ttf`. Word lays Aptos out with `Aptos.ttf`, as its name
says; the index now takes a typographic family only for the four styles Word asks for.
No baseline had noticed: the cuts share their vertical metrics.

### 3.2 Real text: the wrapping probe

`make_wrap_probe.py` / `read_wrap_probe.py`. Twelve paragraphs of this project's own prose,
written to hold what real text holds (hyphenated compounds, en and em dashes, figures
with separators, parentheses, quotation marks, slashes, a path-like token, long and short
words), set in **ten faces** — Calibri, Arial, Times New Roman, Cambria, Georgia, Verdana,
Aptos, Courier New, Helvetica Neue (1000 upm) and Century Gothic — at **five sizes**
(8.5, 10, 11, 13.5, 17 pt) in **five columns** (2160, 3517, 5040, 6803, 9026 twips), the
four styles rotating: 250 paragraphs.

| | Lines agreeing |
| --- | --- |
| every face, size, column and style | **1,778 / 1,778** |
| justified, centred, right-aligned (3 faces × 5 columns each) | 257 / 257 |
| refuted: a line breaks after a hyphen-minus but not after an en or em dash | 1,758 / 1,778 |

The one rule the first run lacked was the dashes: 20 lines, all at "north–south", broke
after the en dash in Word. The rules probe then settled every character class (3.3).

**Helvetica Neue, the 1000-upm face,** agrees in all 177 of its lines with unrounded
advances; rounding each advance to the layout unit also agrees in all 177 — **not
discriminated** (a CFF face at a size where the two differ by more would be).
**Justification** (out of scope) moved no break here: the 85 justified lines break where
the left-aligned model breaks them. Whether Word ever compresses spaces to fit one more
word is not measured.

### 3.3 The rules

`make_break_rules_probe.py` / `read_break_rules_probe.py`: 457 paragraphs, each built so
one question decides where its first line ends — a filler word is chosen so that the text
under test ends a known number of units from the edge. Calibri and Times New Roman unless
said; 11.5 pt, where every width is reachable.

| Family | Lines | What Word does |
| --- | --- | --- |
| `after` / `before` | 318 / 318, 314 / 314 | **A line may break after `-`, U+2013 and U+2014** — between letters, between digits, and between the two — **and before nothing.** Not after or before `/ \ ? ! . , ; : ( ) [ ] { } " ' “ ” ‘ ’ % $ & + = * # @ _ \| ~ < > ^ … · • ¡ ¿ € £ °`, U+0060, U+2010 (drawn as `-`, never broken), U+2012 or U+2015. |
| `hang` (and the classes) | 108 / 108 | **Breaking spaces**: U+0020, U+2002 (en), U+2003 (em), U+2005 (four-per-em) and U+200B (zero-width, not drawn) — a line breaks after them, and at its end they do not count. **Not** U+00A0, U+2004, U+2006, U+2007, U+2008, U+2009, U+200A or U+202F: no break at them, and at a line's end they count like letters. |
| `shy` | 40 / 40 | `w:softHyphen` is a break when the hyphen it then draws fits (inclusive: exactly at the edge breaks, one unit over does not); mid-line it is nothing. **U+00AD in the text is a visible hyphen** with its glyph's advance, drawn mid-line, and never a break. |
| `nbh` | 12 / 12 | `w:noBreakHyphen` is drawn as `-`, as wide, and no line breaks after it. |
| `tab` | 30 / 30 | A tab may go down to the next line, where it jumps from the line's start: when the word after it does not fit, the tab goes with it; **a tab whose next stop is a default one past the edge starts the next line.** Default stops: every `w:defaultTabStop` from the column's edge, after the last custom stop only (a custom stop at 3000 clears the defaults before it); the hanging indent is a stop on the first line. **A custom left stop past the edge ends wrapping:** everything after it stays on the line, out to 6,087 px on a 2,479 px page. A right or centre stop lets the text after it move left as far as the tab; past the edge, the text goes out to it. |
| `br` | 15 / 15 | `w:br` and `w:cr` end the line; the spaces after them stay at the start of the next. |
| `kern` | 54 / 54 | With `w:kern` at or below the size (22 and 23 kern at 11.5 pt, 24 does not), **the legacy `kern` table's pairs**, across run boundaries; **a pair adjusts its left glyph's advance**, so a line's last letter is charged its pair with the space after it (Times New Roman `A ` −113, `Y ` −76). Refuted: the OpenType `kern` feature's pairs (Aptos `Bo`, in the feature and not the table), 52 / 54; charging a pair only when both its glyphs are on the line, 50 / 54. |
| `run` | 10 / 10 | No break at a run boundary inside a word, whatever changes there (colour, size, face, bold). |
| `long` | 14 / 14 | A word too long for any line goes down to a line of its own first (after a space or a hyphen), then splits. |
| `format` | 4 / 4 | `w:caps` measures the capitals; `w:spacing` adds after every glyph (Phase 2). |
| `label` | 1 / 1 | A list label and its tab start the first line. |
| **All** | **920 / 920** | |

**Also measured:** without `w:defaultTabStop` Word's stops are every **708 twips** (1.25
cm, from this machine's metric `Normal.dotm`) — an application default like the 12 pt
size, `linebreak.APPLICATION_DEFAULT_TAB`; every real document here states 720. The
probe's `w:w` (character scale) case was dropped: Word draws a scaled run as a text
object `quartz_pdf` does not place on its line, so its breaks could not be read.

### 3.4 Whose metrics: the tables against the faces Word lays out with

`tools/check_advance_tables.py` compares every advance in `ooxml-common`'s tables with the
installed face Word uses (the macOS copy first; Phase 2, 2.4), and every kern pair with its
legacy `kern` table.

| Word's face | Table | Advances, upright / bold | Kern pairs (legacy table) |
| --- | --- | --- | --- |
| Calibri | Carlito | **346 / 346, 346 / 346** | 349 / 405 pairs differ (Carlito's own pairs) |
| Times New Roman | Tinos | **349 / 349, 349 / 349** | agree |
| Courier New | Cousine | **348 / 348, 348 / 348** | none |
| Cambria | Cambria | **348 / 348, 344 / 344** | 60 differ upright |
| Aptos | Aptos | **347 / 347, 347 / 347** | 16,105 differ (the table is the OpenType feature) |
| Arial | Arimo | 348 / 349, 348 / 349: U+02C6 upright (683 against 682) and U+00B5 bold (1229 against 1180) | 9 / 7 differ (U+00A0 pairs) |
| Calibri Light | Carlito | 50 / 346 — no table of its own | — |

So `TableAdvances` answers advances for these six faces (Arial's two overridden), and
**never kerning**: the tables' `kerning` is the OpenType feature, of a clone for Calibri,
and Word applies the legacy table. (ooxml-common 0.5 added each entry's Office face's
legacy table, `legacy_kerning`, after PowerPoint too was measured charging it rather than
the feature; docx2svg's body text still kerns from the installed face itself.) No table holds an italic cut; the
tests and every score here use the advances of the installed faces, recorded
(`probe-advances.json`, `break-advances.json`). Beyond the tables, Word's widths are the
installed faces' `hmtx` for Georgia, Verdana, Helvetica Neue, Century Gothic, Symbol,
Wingdings (its symbol `cmap`, U+F0xx, for list labels) and the embedded Ubuntu — the
probes and documents above agree line for line.

### 3.5 The real documents

Their paragraph and line assignments came from the oracle; now which characters each line
holds comes from the model. `tools/read_breaks.py` (and `--record`, `--record-scratch`);
`tests/test_linebreak.py`, `tests/test_filesamples.py`, `tests/test_local_corpora.py`.
Paragraphs `baselines.blocks` puts out of scope (tables, drawings, fields, frames, note
references) are out of scope here too; lines below a table are not — breaking does not
depend on the height above.

| Document | Lines that break where Word breaks them | Paragraphs of more than one line |
| --- | --- | --- |
| `layout-sweep.docx` (block D: 4 / 4) | **91 / 91** | 1 |
| `style-document.docx` | **34 / 34** | 4 |
| `samplelib/sample-long.docx` | **470 / 470** | 73 |
| `samplelib/sample-resume.docx`, `sample-simple.docx`, `sample-blank.docx` | **20 / 20, 24 / 24, 0 / 0** | 1, 1, 0 |
| `wordto` (five documents) | **155 / 155** | 39 |
| `filesamples/sample1.docx` (mode 12) | 112 / 115 | 25 |
| `filesamples/sample2.docx`, `sample3.docx` | **10 / 10, 36 / 36** | 4, 6 |
| `filesamples/sample4.docx` (175 pages) | **6,212 / 6,212** | 600 |
| the local corpora | every scored line (their numbers stay beside them) | |
| **All committed and `filesamples`** | **7,164 / 7,167** | |

**The three that do not agree are one paragraph, and a finding:** `sample1` embeds Ubuntu
Mono, and in compatibility mode 12 Word *draws* the embedded Ubuntu Mono but **advances
it by Calibri's widths** — every glyph step of its 95 glyphs matches Calibri's `hmtx` to
0.03 px (proportional steps, from a monospaced face). In modes 14 and 15 (the same bytes
with `compatibilityMode` set, `scratch/variants/`) Word uses Ubuntu Mono's own widths and
that paragraph agrees. The same substitution explains most of **finding 7**: taking
Calibri's vertical metrics for Ubuntu Mono in mode 12 moves `sample1`'s baselines from 52
/ 66 to 63 / 66 — a diagnostic, not adopted: one document, and whether it is embedded
faces, fixed pitch or this face in mode 12 is not known. **Probe:** a generated document
embedding a proportional and a fixed-pitch OFL face, in modes 12, 14 and 15.

**The vertical model now takes each line's characters from the model:**
`baselines.predict(..., advances=...)` breaks each paragraph with `linebreak` and uses
Word's lines only where the breaker cannot measure a paragraph or makes another number of
lines (tables, which are out of scope). Every pinned baseline score stays: 495 / 495
real-world, `sample4` 5,551 / 5,551, `sample1` 52 / 66, and the rest.

### 3.6 Not measured, and left for later

* **Hyphenation** (`w:autoHyphenation`; no document here sets it), **East Asian line
  breaking** (kinsoku, `w:kinsoku`, `w:wordWrap`) and **justification** (whether Word
  compresses or only expands; 85 justified lines broke as left-aligned) — each its own
  measurement, per this phase's "do not do". *Justification: measured in 3.8 -- mode 15 compresses.
  Hyphenation: measured in 3.9.*
* Tabs: decimal and bar stops, leaders, a right or centred text wider than its room,
  a tab in a list label wider than the hanging indent past several stops; the rule
  that a *default* stop past the edge starts a new line while a *custom* one ends
  wrapping was seen once each.
* `w:w` (character scale), `w:smallCaps`, `w:fitText`; fields' drawn results; a line
  beside a floating drawing (its width is not the column's); text columns (`w:cols`) —
  read, but `sample3`'s two-column section was not matched to Word's lines. *Laid out by
  column since 4.14, every glyph Word's.*
* The advance of a 1000-upm face (per glyph rounding or not, 3.2).
* Mode 12 and embedded faces (3.5).

### 3.7 A shared paragraph interface for both breakers — the proposal

Phase 1 left `pptx2svg`'s `text/wrap.py` behind because its *method* transfers and its
*types* (`m.Paragraph`, `m.RunProperties`, `m.TextRun`) do not. With both breakers now
measured, the question can be answered from evidence rather than designed in advance.

**What is the same, measured on both sides.** Greedy first fit; a budget that is exact
and inclusive with no slack (PowerPoint: `make_cjk_wrap_probe.py`, 53 slides; Word: 3.1,
3,120 cases); a word that fits nowhere is split at the character that overflows; a
forced break ends a line; a kern pair is charged only between glyphs of one face, size
and run formatting; the widths are the same tables (3.4). That loop is the method, and
it is about forty lines in either project.

**What differs is data, and each difference is a measurement:**

| | PowerPoint (`pptx2svg`) | Word (`docx2svg`) |
| --- | --- | --- |
| break unit | a token: a Latin word, a run of spaces, one CJK character | one character |
| where a line may break | before every token; kinsoku moves the break back (pushback, and never to leave one token) | after a space, `-`, U+2013, U+2014, a soft hyphen that fits; before a tab (3.3) |
| a space at the end of a line | the token that overflows is dropped | hangs, any number (3.1) |
| kern pair | charged when both glyphs are on the line; the OpenType feature, from GPOS | charged to the left glyph, the space after it included; the legacy `kern` table (3.3) |
| budget | `available × (1 + 1e-6)` in float px (an epsilon for float sums) | exact integers of 1/4096 pt |
| a word that fits nowhere | re-measured character by character, *without* the epsilon | already characters |
| tabs | not modelled | stops, defaults, alignment, past the edge (3.3) |

**The proposal: share the loop over units that carry their own rules, and keep the
rules with each consumer.** In `ooxml_common.text.breaking`:

```python
@dataclass
class Unit:
    width: Number            # its advance: float px, or int/Fraction in 1/4096 pt
    kind: str = GLYPH        # GLYPH, SPACE, BREAK (forced), TAB
    join: Number = 0         # its kern pair with the next unit; 0 across a run change
    break_after: bool = False  # the format's break classes, decided by the consumer
    hyphen: Number = 0       # what a break after it draws (a soft hyphen) and must fit
    item: object = None      # the consumer's own (a pptx token, a docx piece)

def break_lines(units, right, *, first_start=0, start=0, tolerance=0,
                join="pair" | "left", spaces="drop" | "hang",
                tab=None, split=None, pushback=None) -> tuple[list[Unit], list[Line]]
```

Every keyword is one row of the table above, named after the measurement it encodes;
`tab(x, index, first_line) -> (x, limit)` is Word's tab stops, `split(unit, room)` is
PowerPoint's re-measuring of an over-long token (it returns the finished chunks, so the
no-epsilon check stays PowerPoint's), and `pushback(units, start, end) -> end` is
PowerPoint's kinsoku, which Word's East Asian breaking can reuse once measured. The
paragraph *protocol* is then just "a list of `Unit`s": `pptx2svg` builds it from `_Token`s
(keeping token units, so its float sums are bit for bit what they are now), `docx2svg`
from `linebreak.Piece`s; neither model type crosses the boundary, which was Phase 1's
objection.

**Evidence that it holds for both, and what it costs.** `docx2svg.linebreak.break_pieces`
is this loop with `join="left"`, `spaces="hang"`, a `tab`, and `break_after` from
`may_break`; every line in 3.1–3.5 goes through it. `pptx2svg`'s `_layout_tokens`
decomposes into the same loop with `join="pair"`, `spaces="drop"`, `split` =
`_split_token_by_chars` and `pushback` = `_kinsoku_pushback`, reading it line by line;
the two places where its current code is subtle are both kept by construction —
the fit test `(x + join) + width` against the accumulation `x + (join + width)` (the two
associate differently, and must stay so for identical floats), and kinsoku's refusal to
push back to a single token, which a static "forbidden break" flag would *not* keep
(`東「東` full and `。` next: a static rule breaks after the first `東`; `pptx2svg` breaks before `。`). That is why
`pushback` is a hook and not a flag. What it costs: six keywords whose meaning is only in
the consumers' docstrings, for a loop of forty lines — the tables are the valuable shared
part, and they are shared already.

**Not implemented here, and why.** The move needs a branch of `ooxml-common` (the
module), one of `pptx2svg` (`_layout_tokens` over it, with its suite, its VRT snapshots and
`tools/fidelity.py --json` held identical: on `main` today 2,522 passed and 43 skipped,
and the fidelity scores equal `tests/fidelity-baselines.json`) and one of `docx2svg`
(`break_pieces` over it, after `ooxml-common`'s lands). Changing the two sibling
repositories was not permitted in this session, so the proposal stops at the design, with
`main`'s numbers above as the baseline the move must reproduce. **Recommendation:** make
the move when a third breaker (Word's East Asian breaking, or a table cell's) would
otherwise copy the loop again; until then the duplication is forty lines and the
measurements, which are what must not diverge, are in one place.


### 3.8 Justification in compatibility mode 15 — measured

Phase 4.6 found that in mode 15 Word fits more words on a justified line than the
breaker (65 paragraphs of the pagination probe then), and 5.10 that it runs up to 24.8%
of a line's space width past the edge. `make_justify_probe.py` / `read_justify_probe.py`,
`tests/test_justify.py` measure it as 3.1 measured the budget: one justified paragraph a
case, its first line `k` spaced copies of a word and then a composed word `W2` whose end
lies δ units (1/4096 pt) past the edge, then more words so the line is never the last.

* **Round 1**: Calibri, Times New Roman, Arial and Georgia at 11 and 14.5 pt; `k` = 1, 2,
  3, 5, 8, 12; δ from 0 to 40% of the spaces' width by 2.5% (816 cases a document).
  Mode 15 justified, and three controls: mode 15 left-aligned, mode 14 justified, no
  `settings.xml` justified.
* **Round 2**: δ at single units about the edge round 1 found (even units at 11 pt, where
  only even widths exist); spaces of two sizes on one line; two spaces between words
  (406 cases; mode 15 justified and left-aligned).
* **Round 3**: short final words (1,500-3,000 font units) after 24 two-letter words, δ
  from 30% to 70% of the word's own width (52 cases) -- asked by the render probe's
  first justified line, which kept `is` and sent `the` down though `the` needed only
  14% of its spaces.

**The rule** (`linebreak.squeezes`), inclusive, in the layout unit -- a word may end
`over` units past the right edge when both hold:

1. **no space shrinks by more than a quarter**: `over` is at most the sum over the line's
   spaces (after its last tab) of each space's width / 4, **rounded half up to the unit
   per space** (`linebreak.squeeze`);
2. **the squeeze is at most half the stretch**: `over / k <= (slack / k') / 2`, where `k` is
   the number of spaces on the line with the word, and `slack` and `k'` are how far short
   of the edge the line without it ends and the spaces between its words -- each space
   shrinks by at most half of what each would grow were the word sent down. With no
   space on the line without it, the first condition decides.

Word then draws the line compressed: the overrun taken from its spaces in proportion to
their widths, the shift after each space rounded to the unit (`layout.line_positions`).

| Rule (cases whose first line Word and the rule agree on) | Round 1 (816) | Round 2 (406) | Round 3 (52) |
| --- | --- | --- | --- |
| **the model: both conditions** | **816** | **406** | **52** |
| the first condition alone | 814 | 400 | 14 held of 52: every short word past its half fails it |
| a quarter of the spaces' total width, unrounded | 801 | 354 | |
| each space's quarter truncated | 793 | 318 | |
| no squeeze (Phase 3.1's budget) | 343 | 176 | |
| the controls, no squeeze (mode 15 left-aligned; mode 14 and none justified) | 816 / 816 each | 406 / 406 | |

The second condition separates round 3 at 0.49 / 0.51 of the stretch (the largest ratio
held, the smallest sent down), and explains the eight round 1 and 2 cases the first
condition alone got wrong: all Georgia, whose long twelve-space lines ended in short
composed words. Where the lines agree, **every glyph is where Word drew it** in round 1
(71,586 / 71,586) and round 3 (4,309 / 4,309), and in round 2 all but 35 of 35,404, each
within a layout unit or two (how Word rounds each space's share of the overrun: equal
shares, and each share rounded down, up or to the nearest, all score lower; the
proportional share with the running shift rounded is adopted as the best tried).

**What it moved** (compared line by line with the breaker before it): **no line of any
committed document, of `filesamples`, of the local corpora, or of any document below
mode 15 or not justified** -- none of them holds justified text in mode 15, and the rule
is off everywhere else (`Geometry.squeeze`); every other recording and snapshot is
unchanged. In the probes: the render probe's mode-15 document, 35 lines of its justified
paragraphs re-broken (694 lines, one fewer; one line onto another page; its page count
unchanged), its glyphs off Word's position 1,397 -> 1,093, all of them now its
distributed paragraphs (below) but five a layout unit off; the table content probe's
mode-15 justified cell, 6 lines, its 24 glyphs off -> 0 (every glyph of every table probe
now exact). The `justification-unmeasured` warning is gone.

**Not measured**: `distribute` in mode 15 (the render probe's distributed paragraphs
break otherwise than the model there, and are not justified text in this sense); a
line whose words carry different spaces (en, em) or tabs; whether the second condition
uses the line's own last-line state (the probe's lines are never last).

### 3.9 Automatic hyphenation — measured

`make_hyphen_probe.py` / `read_hyphen_probe.py`, `tests/test_hyphenation.py`,
`docx2svg.hyphenate`, `linebreak.auto_hyphens`, `linebreak.auto_hyphen_allowed` and
`paginate.hyphen_at_page_end`. With
`w:autoHyphenation` in `settings.xml` Word breaks words at the end of a line and draws a
hyphen there. 33 documents, every one in three settings -- no compatibility mode, mode 14
and mode 15 -- since mode 15 hyphenates by other rules: 8,577 paragraphs, Calibri (Times
New Roman in some), every width composed from `ooxml-common`'s advance tables. Every
paragraph keeps its lines together and has 18 pt after it, so its lines are found in
Word's PDF by the gaps between them; what is recorded is the text of the lines that decide
-- each swept case's first line, every line of running text -- and every hyphen Word drew
at a line's end in running text.

**Where the break points come from.** Word's proofing tools hold a hyphenator per
language, and they are Microsoft's: nothing of them is in this repository. The library
breaks words by **Liang's algorithm over openly licensed patterns** (`src/docx2svg/patterns/`,
each file's source, version and licence in `SOURCES.txt` beside it): English by the
**Moby Hyphenator word list** (Grady Ward, public domain) and, for the words it does not
hold, Knuth's `hyphen.tex` with the TUGboat exception log (3.9.1; until then hyph-utf8's
American English, Kuiken and Knuth, all-permissive), and from hyph-utf8 Dutch (Tutelaers, MIT),
German, reformed spelling (the Trennmustermannschaft, MIT) and French (Flipo, Gaulle and
Reutenauer, MIT; 3.9.3). Word's own break points are observed by the probe and
kept in the recording, where they only score the patterns; no word list is taken from
them. LibreOffice's German and French patterns are LGPL and were not used; its Dutch ones
(OpenTaal, BSD) agree with Word no better than hyph-utf8's, and none of its sets agrees
better than what is shipped (3.9.5).

**The `points` sweep.** 142 words -- 48 English, 48 Dutch, 46 German, from 4 to 35
letters, compounds and diacritics among them -- at 24 pt after `Hn`, once for every `k`
from 1 to the word's length less one, in a column that ends where the first `k` letters and
a hyphen end (rounded up to 5 twips, less than any letter): Word breaks the word at its
last point at or before `k`, if anywhere. The zone is 0, and at 24 pt mode 15's own
threshold (below) is passed by any fragment. Read off the sweep, Word's points against the
patterns':

With hyph-utf8's American English, as first measured (the English row now: 3.9.1):

| Language | Words whose points are Word's: none, 14 / mode 15 | Points: in common / the patterns' / Word's, none and 14 | mode 15 |
| --- | --- | --- | --- |
| en-US | 31 / 48 (65%) / 34 / 48 (71%) | 104 / 113 / 117 | 100 / 106 / 113 |
| nl-NL | 46 / 48 (96%) / 46 / 48 | 119 / 120 / 120 | 111 / 112 / 112 |
| de-DE | 42 / 46 (91%) / 42 / 46 | 110 / 112 / 114 | 104 / 106 / 108 |

Word's English hyphenator breaks as an American dictionary does (`rep-re-sent-a-tives`,
`meas-ure-ment`, `in-de-pend-ent`, `tel-e-com-mu-ni-ca-tions`) where Liang's patterns
break by syllable rules (`rep-re-sen-ta-tives`, `mea-sure-ment`); Dutch and German agree
but for compounds the patterns do not know (`Urin-stinkt`, `Staub-ecken`, `Wach-stube`
against `Ur-instinkt`, none, `Wachstu-be`) and `kin-deren`. **Spelling changes at a
break**: Word draws Dutch `omaatje` as `oma-` / `tje`, but not `autootje` or `menuutje`
(`au-tootje`, `me-nuutje`) -- a word of its dictionary, not a rule, so not modelled; German
`Zucker` breaks `Zu-cker` (reformed spelling, no `ck` to `k-k`) and `Schiff-fahrt` keeps
its three f's.

**The rules, every one agreeing in every case:**

| Rule | Measured by | Word |
| --- | --- | --- |
| **Fragments** | `points` | At least **two letters before a break, in every mode; two after it below mode 15, three in mode 15** (`riv-er`, `com-pa-ny`, `de-ze`, `kop-je` below 15; `river`, `com-pany`, `deze`, `kopje` in 15). No other minimum word length: a word of four letters breaks below mode 15 (`de-ze`). |
| **The hyphen must fit** | `fit` (δ to the unit, 3 faces and sizes), `points` | The line through the fragment **and its hyphen** must end at or before the edge, inclusive, as a soft hyphen's (3.3). In justified text in mode 15 the hyphen may run past the edge as a word may (3.8): `infra-` held at δ = +206 units. |
| **Below mode 15, the next letter must fit too** | `points` (1,385 / 1,394 cases per setting below mode 15; the other nine a word that fit whole, or `omaatje`), `fit` | Word breaks a word **before the letter that runs past the edge, never at it**: the point is taken only if the letter after it also fits. `infra-` fitting with 1 twip to spare is refused when `s` would not fit, and `in-` taken. Mode 15 does not ask it. Soft hyphens never ask it (`shy`). |
| **The zone, below mode 15** | `zone`: δ = −1,024 … +1,024 units about the zone, 4 alignments, 3 faces and sizes, zones unset, 0, 180, 360, 720 | A word is hyphenated only if it **starts at least the zone before the right edge** (inclusive: δ = 0 hyphenates, −2 does not), in **every alignment** alike. **Unset, the zone is 425 twips** (0.75 cm, this machine's metric `Normal.dotm`; not the schema's 360) -- an application default, like the 708-twip tab (3.3). Zone 0: only the fit decides. |
| **Mode 15 does not read the zone** | `zone` (unset and 720 alike), and the exploration | Left, right and centred text: hyphenated only if the room after the line's text before the word -- the edge less the end of the previous word -- is **more than about 351.5 twips** (351.01 not, 352.00 hyphenated at 11 pt in Calibri and Times New Roman; at other sizes and faces between 350.6 and 353.2: a spread of a twip not explained). **Justified text: no threshold**, only the fit. |
| **`w:consecutiveHyphenLimit`** | `limit1`, `limit2`, `limit0` (running text) | At most that many lines in a row end in an automatic hyphen; 0 or unset, no limit. A paragraph's last line never ends in one, so the count starts again in every paragraph. **Soft hyphens are not limited** (a paragraph hyphenated by hand breaks alike under limits 0, 1 and 2); whether they count toward an automatic hyphen's limit is not measured. |
| **`w:doNotHyphenateCaps`** | `caps` (with and without it) | Leaves out a word whose letters are **all capitals as drawn** -- typed so, or by `w:caps`; `Infrastructure` and `InfraStructure` still break. Without it capitals break as lower case does. |
| **`w:suppressAutoHyphens`** | `suppress`, `limit` | Turns automatic hyphenation off for the paragraph -- direct or from its style, and `w:val="0"` turns it on again over a style; soft hyphens still break. |
| **`w:noProof`** | `noproof` | A word whose first letter is `w:noProof` is not hyphenated. Below mode 15 the word ends where a `w:noProof` run starts, and Word hyphenates it only when the line runs past the edge before that run (`infra` + `structure`: `in-` while the overflow is in `infra`, nothing after); in mode 15 only the first letter decides. |
| **Soft hyphens in the word** | `shy` | Mode 15: the word keeps all its own points, the soft hyphen's among them. Below mode 15: **only the points before its first soft hyphen**. |
| **Compounds** | `compound` (hard and non-breaking hyphens) | Mode 15 hyphenates **each part as a word of its own** (`tions-in-fra-struc-ture`, `well-es-tab-lished`, `Arbeits-Un-fä-hig-keit`); below mode 15 **only the first part** -- after a hard hyphen the line breaks at it, after a non-breaking one the rest is split as a word too long for the line. |
| **Language** | `lang` | The run's `w:lang` chooses the hyphenator: one string breaks differently in en-US, nl-NL and de-DE (`Demon-stra-tion`, `De-mon-stra-ti-on`, `De-monst-ra-ti-on`). **en-GB breaks as en-US, de-AT and de-CH as de-DE, fr-CH and fr-CA as fr-FR; nl-BE, fr-BE and `x-none` are not hyphenated** on this machine (no Belgian hyphenators installed; 3.9.3). No `w:lang` on the run: the document default's. A word whose language changes inside it breaks by neither (`In-f-ra-struktur`): not modelled, its first letter's language is used. |
| **The hyphen** | `text` (every glyph of running text, `glyphs.py`'s checks) | **`-` (U+002D) of the run's face and size, its advance**, drawn at the pen after the fragment as the line's last glyph; a justified line is spread so that the hyphen ends at the edge. 90 / 98, 90 / 98 and 95 / 105 of Word's line-end hyphens are where the library draws them; the rest are on lines whose words the patterns break elsewhere. With the Moby list (3.9.1): all of them. |

**The model** (`linebreak`): with `w:autoHyphenation` on and the paragraph not suppressing
it, `pieces` puts an automatic soft hyphen (`Piece.auto`) at every point of every word
(`auto_hyphens`), and `break_line` takes it as the line's end when its hyphen fits and
`auto_hyphen_allowed` -- the zone or mode 15's threshold, the next letter, the limit,
`w:noProof` -- allows it, else the last break before it. The layout draws it as a soft
hyphen's. **Scores** (paragraphs whose recorded lines the library's are): **8,116 / 8,577**
(none 2,782 / 2,931, mode 14 2,782 / 2,931, mode 15 2,552 / 2,715), against 3,019 /
8,577 with hyphenation off; every zone, fit, caps, suppress and noProof case agrees. The
461 that do not are the patterns' words (`telecommunications`, `responsibility`,
`measurement`...) and their lines after them, French, and the word of two languages. With
French patterns and Knuth's English (3.9.1, 3.9.3) the same documents score **8,147 /
8,577**, and every document of the probe, those of 3.9.1--3.9.3 among them, **18,039 /
18,868**; with the Moby list before Knuth's patterns (3.9.1) **18,374 / 18,868**.

**What it moved**: nothing outside the probe. No committed document, no `filesamples`
document and no local corpus document sets `w:autoHyphenation`, and without it no piece is
added; every recording, snapshot and fidelity baseline is unchanged (the harness: rc 0).

**Not measured**: other languages' hyphenators (Spanish, Italian...) and other English
variants than en-GB; apostrophes outside French; whether a line ending in a soft hyphen counts toward
`w:consecutiveHyphenLimit`; the unexplained twip of mode 15's threshold; `w:smallCaps`
under `w:doNotHyphenateCaps`; hyphenation beside a floating drawing or in a table cell
(the same breaker, not probed); digits inside a word (taken as ending it, and a word
beside digits as not hyphenated); a column's end and a soft hyphen's at a page's end
(3.9.2).

#### 3.9.1 English: which open patterns agree with Word most often — measured

`words-*` (`make_hyphen_probe.py`): **311 more English words**, ordinary dictionary words
chosen for a spread of lengths and endings before Word was asked (none holds the letters
"i" and "c" side by side), swept a letter at a time as `points` sweeps its 48, in the three
settings. Mode 14 answers as no mode does, every word. Every openly licensed English set
found, scored against Word's points over the 359 words (scratch scorer; only the shipped
set's numbers are pinned, `tests/test_hyphenation.py`):

| Set | Licence | 48 words: none / 15 | 311 words: none / 15 | **359 words: none / 15** |
| --- | --- | --- | --- | --- |
| hyph-utf8 `en-us` (Kuiken's `ushyphmax` with Knuth's; the library's until now) | all-permissive | 31 / 34 | 247 / 269 | 278 (77%) / 303 (84%) |
| hyph-utf8 `en-gb` (Wujastyk and Toal, patgen over an OUP word list) | MIT | 18 / 23 | 134 / 176 | 152 (42%) / 199 (55%) |
| LibreOffice `hyph_en_US.dic` (`hyphen.tex` and the 2007 TUGboat log, as patterns) | BSD-style | 32 / 35 | 253 / 273 | 285 (79%) / 308 (86%) |
| LibreOffice `hyph_en_GB.dic` (`ukhyphen.tex`) | BSD-style | 18 / 23 | 134 / 176 | 152 (42%) / 199 (55%) |
| Knuth's `hyphen.tex` alone | its own: copy freely, modify under another name | 31 / 34 | 249 / 269 | 280 (78%) / 303 (84%) |
| **`hyphen.tex` with TUGboat's `ushyphex.tex` (2021)** -- shipped, `hyph-en-us-knuth` | the same; TUG: "freely use, modify and/or distribute" | **32 / 35** | **256 / 275** | **288 (80%) / 310 (86%)** |
| hyph-utf8 `en-us` with `ushyphex.tex` | as above | 32 / 35 | 251 / 272 | 283 (79%) / 307 (86%) |
| patgen (`pypatgen`, 4 levels, Liang's parameters) over the Moby Hyphenator list | public domain list | 35 / 38 | 217 / 242 | 252 (70%) / 280 (78%) |
| *The Moby Hyphenator list itself, as a word list (hyph-utf8 `en-us` for words not in it)* | *public domain* | *42 / 45* | *272 / 292* | *314 (87%) / 337 (94%)* |

**Chosen: Knuth's `hyphen.tex` with the TUGboat exception log** -- the best of the
pattern sets on both probes and in both rule sets, and the best on the 311 words that
did not choose it. `ushyphex.tex` is lower-cased and loses 13 of its 1,753 entries that
hold a string this repository keeps out of its files (`SOURCES.txt`); the score above is
of the file as shipped. Nothing was fitted to Word: no word list from Word's answers, no
patgen over them. **The gain is small** -- ten words in 359 below mode 15, seven in it --
since every pattern set breaks by syllable rules and Word breaks as an American dictionary
does (`rep-re-sent-a-tives`, `in-de-pend-ent`). The one open source that breaks so is the
**Moby Hyphenator list** (Grady Ward, Project Gutenberg #3204, public domain, 187,175
entries, 2.6 MB): as a word list it agrees with Word on 87% / 94% of the words. Patterns
trained on it are cheap -- `pypatgen` (MIT, a scratch environment) trains four levels in
under three minutes -- but with Liang's parameters they come to 276,185 patterns, 3 MB,
and agree *less* (70% / 78%): not shipped. Shipping the list itself, or tuning patgen on
it, is the open option; both are openly licensed and neither touches Word's output. The
list was then shipped (below).

The model with the chosen set: `points` 1,340 / 1,394 (none, 14; 1,330 in mode 15),
`words` 2,479 / 2,579 (2,462). Running text is unchanged (`text` 16 / 18 in each setting;
one `lang` case is lost, 659 / 734 against 660).

**Then shipped: the Moby list as a lookup, Knuth's patterns after it.** The list
(`patterns/en-moby.txt.gz`, `docx2svg.hyphenate.WordList`) is Project Gutenberg's
`mhyph.txt` (SHA-256 `eeb30474...`; Mac Roman, its marks byte `0xA5`), converted as
`SOURCES.txt` says: 167,826 words, one a line, `=` at each point (never a hyphen),
gzip-compressed to **593 KB**, read and kept in memory the first time an English word is
hyphenated (about half a second). Four entries are left out for a string this repository
keeps out of its files. Every `w:lang` of `en-*`: a word the list holds breaks where the
list says, any other where the patterns say; Word's minimums and every rule of 3.9 (zone,
fit, limit, caps, `noProof`, compounds) apply on top as before. The choices, each scored
on the 359 words (none / mode 15) and on every probe paragraph (of 18,868):

| Choice | 359 words | Paragraphs |
| --- | --- | --- |
| Knuth's patterns alone (until now) | 288 / 310 | 18,039 |
| Moby, a spelling listed twice: the first entry; case folded | 314 / 336 | 18,290 |
| ... the spelling as typed first, else the lower-case one | 315 / 337 | 18,299 |
| ... without the words of Moby's compounds | 314 / 337 | 18,297 |
| ... listed twice: only the points both entries have | 314 / 336 | 18,292 |
| ... listed twice: the points either entry has | 314 / 340 | 18,344 |
| ... listed twice: the entry with more points | 317 / 342 | 18,374 |
| **... listed twice: the later entry -- shipped** | **317 (88%) / 342 (95%)** | **18,374** |

* **Case**: the list keeps proper nouns apart (`Po-lish`, `pol-ish`; `Educa-tion`,
  `ed-u-ca-tion`): the spelling as typed is taken if listed, else the lower-case one (a
  capital at a sentence's start, a word in capitals), else the first.
* **Possessives and apostrophes**: an English word ends at an apostrophe already (3.9), so
  `nation's` looks up `nation`; the list's 195 entries with an apostrophe are not kept.
* **Compounds and phrases** (13,246 with a space, 7,399 with a hyphen): not entries, since
  the library hyphenates each part as a word of its own (3.9, Compounds); a part no
  single-word entry spells is kept as a word (3,515; `circumstances` is one).
* **A spelling listed twice** (1,988, mostly two pronunciations: `pres-ident`,
  `pres-i-dent`): the later entry. The list sorts its mark after the letters, so of two
  spellings of one word the later is the one with a mark where they first differ; it and
  "the entry with more points" are the best of the rules tried, on the probe -- a choice
  among the list's own entries, not a word from Word.

Running text now breaks as Word's in every setting (`text` 18 / 18, every line-end hyphen
where Word draws it: 98 / 98, 98 / 98, 105 / 105), `points` 1,359 / 1,394 (1,351 in mode
15), `words` 2,540 / 2,579 (2,537), `rules` 675 / 734 in every setting; Dutch, German and
French are unchanged. What is left is the list's own breaks (`in-ter-na-tion-al`,
`re-spon-si-bil-i-ty`, `knowl-edge` for Word's `in-ter-na-tional`, `re-spon-si-bil-ity`,
`knowledge`) and the words it does not hold (`demonstrates`, `organizations`).

**Then: the probe's own words as exceptions** (the user's call, after the above). The 42
of the 359 words that the list and the patterns break otherwise than Word are spelt as Word
was seen breaking them, in `patterns/en-observed-exceptions.txt`, consulted before the
list (`read_hyphen_probe.py --exceptions` makes it from the recording alone, and a test
holds it to that). Eight break otherwise in mode 15 than below it, as no minimum explains
(`bene-fit`, `ben-e-fit`; `choco-late`, `choc-o-late`; `popu-la-tion`, `pop-u-la-tion`),
so the file gives those a second spelling and `hyphenate.points` takes `mode15`. These are
observations of Word's behaviour on words of probes made to score the patterns, not
Microsoft data; no new probe was made for them, and Word was not run.

| English | 359 words: none, 14 / mode 15 | Probe paragraphs (of 18,868) |
| --- | --- | --- |
| Knuth's patterns alone | 288 (80%) / 310 (86%) | 18,039 |
| Moby, then the patterns | 317 (88%) / 342 (95%) | 18,374 |
| The exceptions, then Moby, then the patterns -- shipped | 359 / 359 | 18,568 |

**Only the score without the exceptions says anything about unseen text**: the 359 words
are the exceptions' training set, so with them every one agrees by construction. Expect
about 88% / 95% of English words to break as Word's, as the Moby row says. The paragraphs
gain twice over for the same reason: `words` and every running-text and `limit` paragraph
now break as Word's in every setting, and `rules` 684 / 734 (687 in mode 15) and `points`
1,372 / 1,394 (1,365) miss only Dutch, German and French words and the word of two
languages.

#### 3.9.2 A page that would end in an automatic hyphen — measured, laid out

Found in the research (3.9.4): [MS-DOCX] 2.3.7 and 2.3.8 say Word 2013 and later does
not let a hyphenated word end a page or column. `bottom-*` (`make_hyphen_probe.py`): the
English running text, left and justified, run over a page's end after each of its
lines 1 to 16 (an empty paragraph of exact height starting each page), no `keepLines`;
in the three settings, and in mode 15 with `useWord2013TrackBottomHyphenation` 0 and with
`allowHyphenationAtTrackBottom` 1. Every line recorded, page by page:

| Setting | Word | Library before / after |
| --- | --- | --- |
| No mode, mode 14 (neither compatSetting) | A page may end in an automatic hyphen (`infra-`, `de-`, `doc-`) | 32 / 32, 32 / 32 |
| **Mode 15, neither compatSetting** (or `useWord2013TrackBottomHyphenation` true) | **The line goes to the next page** -- one line: the line before it may end in a hyphen and stay (`erable responsibility among rep-` ends a page). The keeps then act as on any line that does not fit: widow control moves a paragraph that would leave one line (`lines` 2: the whole paragraph) | 22 / 32 → **32 / 32** |
| Mode 15, `useWord2013TrackBottomHyphenation` 0 | **The word goes to the next page**, the rest of its line stays (`of telecommunications` / `infrastructure demonstrates`), and the paragraph is broken again from the word | 22 / 32 → **30 / 32**: the two others break `notwithstanding` as the patterns do (`notwith-` for Word's `not-`), the rule agrees |
| Mode 15, `allowHyphenationAtTrackBottom` 1 | As below mode 15 | 32 / 32, 32 / 32 |

Modelled (`paginate.hyphen_at_page_end`, `_scan`): where the first line that does not fit
on a page follows one ending in an automatic hyphen, that line is taken as the one that
does not fit; with `useWord2013TrackBottomHyphenation` 0 the paragraph is broken again
from the hyphenated word (`Para.move_hyphenated_word`). On pages only: a column's end,
the settings below mode 15, and a soft hyphen at a page's end are not measured. Word
writes `useWord2013TrackBottomHyphenation` 1 into a document it saves in every mode
(seen in its re-saves of the probes). No committed or local document turns automatic
hyphenation on, so nothing else moves.

#### 3.9.3 French — measured, hyphenated

`french-*` (`make_hyphen_probe.py`): the `points` sweep in fr-FR over 65 words (from
`élu` and `ami` to `constitutionnellement`; accents, `ç`, `ï`, `œ`) and 13 with an
apostrophe, typed both as `'` (U+0027) and as `’` (U+2019); five of them again in fr-BE,
fr-CH and fr-CA; and the French running text in three columns, left and justified. In
the three settings.

* **Patterns: hyph-utf8 `fr`** (Daniel Flipo, Bernard Gaulle, Arthur Reutenauer, V2.13;
  **MIT** since 2016, LPPL before; `SOURCES.txt`). LibreOffice's `hyph_fr.dic` is LGPL and
  was not used.
* **Fragments: as every language's** -- two letters before a break; two after it below
  mode 15 (`dé-jà`, `uni-ver-si-té`), three in mode 15 (`déjà`, `uni-ver-sité`). No word
  breaks after one letter (`élève`, `école`, `idée`, `élu`).
* **Accented letters are letters**: `hô-pi-tal`, `châ-teau`, `fe-nêtre`, `gar-çon`,
  `pré-fé-rence`; `cœur`, `œuvre` do not break.
* **Apostrophes.** The typewriter one is **part of the word**, and Word never breaks at it:
  `jus-qu'à`, `pres-qu'île`, `l'or-ga-ni-sa-tion`, `d'ad-mi-nis-tra-tion`,
  `qu'au-jour-d'hui`, `s'ins-crire` (the patterns, which hold the apostrophe, agree). The
  typographic one **joins the parts as a hyphen does** (3.9, Compounds), though no line
  breaks at it: below mode 15 only the part before it is hyphenated (`au-jourd’hui`,
  `l’organisation` and `jusqu’à` not at all), in mode 15 each part (`l’or-ga-ni-sa-tion`).
  Modelled for French only (`linebreak._FRENCH_JOINERS`); other languages' apostrophes are
  not measured.
* **Regional variants: fr-CH and fr-CA break as fr-FR; fr-BE is not hyphenated** on this
  machine, as nl-BE is not.
* **Agreement**: Word's points against the patterns', **67 / 78 words (86%) below mode 15,
  69 / 78 (88%) in it** (points 122 in common, 132 the patterns', 126 Word's; mode 15 121 /
  125 / 126). Word breaks `con-nais-sance`, `con-som-ma-tion`, `cons-ti-tu-tion`,
  `ex-traor-di-naire` where the patterns do not or break elsewhere. The model:
  **771 / 798** paragraphs (none, 14; 772 in mode 15), every running-text paragraph
  among them (6 / 6).

#### 3.9.4 What others have found about Word's hyphenation — researched

Read and cited only: no Microsoft dictionary, proofing data, word list or decompiled
material was downloaded or used, and none was found published.

**Nobody has published a reverse engineering of Word's hyphenator.** What exists is
behaviour: Microsoft's own specifications and help, LibreOffice's interoperability work,
one source comment in OnlyOffice, and Aspose's documentation.

*Behaviour, against ours:*

| Source | What it says | Ours |
| --- | --- | --- |
| [MS-OI29500], its note on ECMA-376 17.15.1.53 `hyphenationZone` ([link](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/678b19e3-2442-48e8-968a-61d46a51d2e0)) | The standard's default is 360 twips, but "Word assigns different default values for the hyphenation zone, based on the language of the running operating system." | **Agrees**: unset, 425 twips (0.75 cm) on this metric machine (3.9, The zone) |
| VBA `Document.HyphenationZone` ([link](https://learn.microsoft.com/en-us/office/vba/api/word.document.hyphenationzone)); Microsoft Q&A, the zone greyed out in Word 365 ([link](https://learn.microsoft.com/en-us/answers/questions/46b58323-26ba-46cc-8c0e-d46a2f3b2ed9/hyphenation-zone-setting-is-greyed-out-in-word-365)) | "Unless Word is in compatibility mode, HyphenationZone always returns 99999999"; the setting returns in Word 2010 compatibility mode | **Agrees**: mode 15 does not read the zone |
| OnlyOffice `sdkjs`, `word/Editor/Paragraph_Recalculate.js` ([link](https://github.com/ONLYOFFICE/sdkjs/blob/master/word/Editor/Paragraph_Recalculate.js)) | In mode 15 the zone is taken as 360 twips whatever the document says, "as in MSWord (checked in the 2019 version)": hyphenate when the room left on the line exceeds it | **Near**: we measure a threshold of about 351.5 twips of room after the previous word's end, in left, right and centred text, none in justified text |
| [MS-DOCX] 2.3.7 `allowHyphenationAtTrackBottom`, 2.3.8 `useWord2013TrackBottomHyphenation` ([link](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/e0c0663e-a5e7-4a44-8360-b0b5df1f43e6), [link](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/e6cd53b7-89ac-4c64-a6b2-cd1df8b08228)) | A hyphenated word at a page's or column's end: its line moves (absent or true), only it moves (false), or it stays (allow) | **Was missing: a candidate, probed and modelled** (3.9.2), pages only |
| [MS-DOC] DopBase ([link](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/f885b87a-15b3-460e-aecc-213bd17f960e)); ECMA-376 `consecutiveHyphenLimit` via [MS-OE376] ([link](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/9f086eff-3cf2-43e9-82c1-16ddaef7d515)) | `fHyphCapitals`, `dxaHotZ`, `cConsecHypLim`; a consecutive limit of 0 or none is no limit | **Agrees** (3.9, `w:consecutiveHyphenLimit`, `w:doNotHyphenateCaps`) |
| Aspose.Words, "Working with Hyphenation" ([link](https://docs.aspose.com/words/net/working-with-hyphenation/)) | Hyphenates with OpenOffice (TeX-pattern) dictionaries, "MS Word uses dictionaries other than OpenOffice dictionaries"; the dictionaries' LEFTHYPHENMIN/RIGHTHYPHENMIN "are ignored. Aspose.Words uses its own set of distance parameters depending on the document compatibility mode" (values not published) | **Agrees**: the minimums are Word's, not the patterns', and differ by mode (2 / 2 below 15, 2 / 3 in it) |
| Microsoft Q&A, minimum letters ([link](https://learn.microsoft.com/en-us/answers/questions/4804039/can-the-minimum-number-of-letters-before-or-after)) | Word 2003 would not break `party`; Word 2010 breaks `par-ty`; no setting for the minimums | **Agrees**: two letters after a break below mode 15 |
| LibreOffice tdf#149421 ([discussion](https://www.mail-archive.com/libreoffice-bugs@lists.freedesktop.org/msg1032576.html)) | László Németh: with several points in the zone "MSO doesn't choose the last one" (Hungarian `közvet-lenül`, not `közvetle-nül`) | **Candidate, refuted here for English, Dutch and German**: laid out with Word's own points from the sweeps in place of the patterns (a scratch check, nothing kept), every `text` and `limit` paragraph of every setting is Word's (54 / 54 per setting): Word takes the last point that fits. Hungarian is not measured; Word's Hungarian hyphenator may simply have no point at `közvetle-` |

*How LibreOffice's Word-compatible settings came about* (László Németh, NLnet and
Collabora; commits on [LibreOffice/core](https://github.com/LibreOffice/core)): tdf#121658
(2019, Samuel Mehrbrodt) DOCX `doNotHyphenateCaps`; tdf#149248 and tdf#149324 (7.4, 2022)
"don't hyphenate the last word" and a minimum word length, citing InDesign and CSS, not
Word; tdf#149420 and tdf#149421 (7.4) the hyphenation zone, "an OOXML interoperability
feature", and its DOCX import and export; tdf#161628 (24.8, 2024) an unset zone imported
as the standard's 360 twips; tdf#161643 `consecutiveHyphenLimit`; tdf#160518 and
tdf#165354 (24.8, 25.8) `useWord2013TrackBottomHyphenation` and
`allowHyphenationAtTrackBottom` as "hyphenation-keep", "like MSO does"; tdf#165984 (25.8,
[7d384fb](https://github.com/LibreOffice/core/commit/7d384fb1c232f7aa720436bc68dc1de334bf7297))
end zones, which names "the default OOXML hyphenation zone 18 pt (or its MS Word variants,
for example 21.25 pt [0.75 cm])" -- our 425 twips; and the justification work of
tdf#119908 and after, "interoperable with the default hyphenation algorithm of MSO 2013
and later" ([d572576](https://github.com/LibreOffice/core/commit/d572576dd6cc3b2473d9abbe8deb00006da130a8)),
"first analysis of the unknown justification algorithm" ([7d08767](https://github.com/LibreOffice/core/commit/7d08767b890e723cd502b1c61d250924f695eb98)).
The values were taken from the standard or found by comparing layouts with Word's on test
documents; no commit describes a systematic measurement, and LibreOffice's DOCX import sets
no Word fragment minimums (they come from its dictionaries). Overview:
[numbertext.org/typography](https://www.numbertext.org/typography/).

*Engine and dictionaries.* Word's hyphenators are per-language components of Microsoft's
proofing tools, files named `MSHY3xx.DLL` (later `MSHY7xx`) -- existence only, e.g.
[processlibrary.com, mshy3el.dll](https://www.processlibrary.com/en/directory/files/mshy3el/319437/).
Office's 2010-era third-party notices
([Exchange 2010 copy](https://learn.microsoft.com/en-us/previous-versions/office/exchange-server-2010/dd351225(v=exchg.141)))
name Lingsoft for the Danish, German and Norwegian hyphenators (with Xerox's two-level
compiler: morphological, not patterns), Priberam for Portuguese and Itautec Philco for
Brazilian Portuguese; no vendor is named for English, French or Dutch, and current notices
list no hyphenators. No public analysis of the binaries or their lexicons was found. Other
renderers: OnlyOffice (7.5 and later) and DevExpress hyphenate with Hunspell `.dic`
patterns, Aspose with OpenOffice dictionaries, docx4j only through FOP's patterns, and
Apache POI does no layout.

*Open English sets* (3.9.1): none claims to approximate Word. hyph-utf8 `en-us` is
Kuiken's `ushyphmax` (all-permissive), `en-gb` Wujastyk and Toal's `ukhyphen` (MIT,
trained on an OUP list that is not itself redistributable); LibreOffice's `hyph_en_US` is
Liang's `hyphen.tex` with the TUGboat exceptions, not Kuiken's, and its `hyph_en_GB` is
`ukhyphen`; neither has a conservative variant. Word's French behaviour is not described
anywhere found; hyph-utf8's `fr` forbids a break at the apostrophe (`2'2`), as Word does.

#### 3.9.5 LibreOffice's hyphenation dictionaries — measured, not used

LibreOffice's `.dic` files for every language the probe holds Word's breaks of, from the
official repository ([git.libreoffice.org/dictionaries](https://git.libreoffice.org/dictionaries),
commit `32b006a2`, 2026-08-21; byte for byte the files of
[wachin/libreoffice-dictionaries-collection](https://github.com/wachin/libreoffice-dictionaries-collection)).
Kept in a scratch directory and never committed: German and French are LGPL. Read as
libhyphen does: the charset line, the option lines (`LEFTHYPHENMIN`... -- Word's minimums
are used instead, as for every set), and the patterns, by `docx2svg.hyphenate.Patterns`.
**No file has a non-standard (`/`) entry.** German and French are two-level (`NEXTLEVEL`):
German's first level, 69,127 patterns from Karl Zeiler's list of words and compounds,
splits a word into parts, and its second, 8,718 patterns, hyphenates each part; French's
first level is empty. Scored as in 3.9.1, words whose points are Word's, (none, 14) / (15):

| Language (words) | LibreOffice set | Licence | SHA-256 | LibreOffice | Shipped |
| --- | --- | --- | --- | --- | --- |
| de-DE (46) | `hyph_de_DE.dic` (2017-01-12, from `dehyphn.tex`) | LGPL 2 or later | `2e2f5ffe...` | 42 / 42 with its compound minimums (2, 2); 41 / 41 without | hyph-utf8 `de-1996`: 42 / 42 |
| de-DE (46) | `hyph_de_CH.dic`, `hyph_de_AT.dic` | LGPL 2 or later | `82515943...`, `3e0b4687...` | the German file but for its header: 42 / 42 | 42 / 42 |
| fr-FR (78) | `hyph_fr.dic` (4.0, 2021, from `hyph-fr.tex`) | LGPL 2.1 or later | `476ca60b...` | 67 / 69 | hyph-utf8 `fr`: 67 / 69 |
| fr-FR, fr-CH, fr-CA (88) | the same | | | 75 / 77 | 75 / 77 |
| nl-NL (48) | `hyph_nl_NL.dic` (OpenTaal) | BSD-3-Clause or CC BY 3.0 | `0a97b76b...` | 46 / 46 | hyph-utf8 `nl`: 46 / 46 |
| en-US (359) | `hyph_en_US.dic` (`hyphen.tex` and the TUGboat log) | BSD-style | `486fb684...` | 285 / 308 (as in 3.9.1) | 317 / 342 without the exceptions; 359 / 359 with them |
| en-US (359) | `hyph_en_GB.dic` (`ukhyphen.tex`) | BSD-style | `9fdc97f7...` | 152 / 199 (as in 3.9.1) | as above |

**None wins**: German, French and Dutch tie with what is shipped (German only with libhyphen's
compound minimums, which Word was not measured to have), and both English files lose. The
library's behaviour is unchanged.

---

## Phase 4 — Pagination (L–XL)

**The other hard problem, and the one where a single error invalidates the rest of the
document.**

The fixture already pins the observation: 70 single-line paragraphs, and Word put the break
after L32. Nothing in the file says that.

**Must be measured before starting:** Phases 2 and 3 complete, because a page break is a
running total of line advances and a wrong advance moves it. Additionally:

1. **`keepNext`, `keepLines`, `widowControl`, `pageBreakBefore`.** The model already
   carries all four, unresolved. These are the properties that *decide* page breaks and
   that `docx-renderer` parses and discards. Each needs a probe.
2. **Footnote space reservation.** Word reserves the footnote area *before* choosing the
   break. `docx-renderer` does not, and its output can be a perfect reproduction of a break
   Word would never have made. Measure before implementing footnotes at all.
3. ~~**Where the first baseline of a page sits.**~~ **Reproduced in Phase 2.** It is not
   `top margin + ascent`: the fixture's 1427 px (page 1, after five lines of mixed
   sizes) and 344 px (page 2) both come out of `docx2svg.vertical.baseline_px` with the
   page's line tops restarting at the top margin. What Phase 4 still owes is *which line*
   starts page 2, not where it sits.

**Done when:** the page assignment of every paragraph in a multi-page probe matches Word's,
for at least three page sizes and three margin sets. This is a discrete answer and the
test should be an equality, not a tolerance.

**Done: see *Phase 4 — measured* below.** In every document of the multi-page probe —
Letter, A4 and A5 by three margin sets, the same body in three more compatibility
settings, and four sections that change size and margins — the model places every page
top where Word does (537 / 537) and every line of every paragraph on Word's page (3,380 /
3,380 paragraphs), from the file alone. On the real documents it places every page top
it reaches: `sample4` 168 / 168 (175 pages, 104 tops inside a paragraph), `sample-long`
35 / 35, `sample-10pages` 2 / 2; the tops it does not place lie past a table or a
floating drawing, where it stops rather than guess.

---

## Phase 4 — measured

`docx2svg.paginate` decides where every page ends. It lays the document out as a
**flow** (`paginate.flow`): each top-level paragraph as its lines — the Phase 3 breaker —
with each line's pitch (`docx2svg.lines`, the vertical model), the space and borders
around it, its keeps, its footnotes and its inline pictures; a table, a floating
drawing or a paragraph the breaker cannot measure is an **obstacle**, where the
paginator stops (`Unplaceable`) instead of guessing a height. Then page by page
(`paginate.page_end`, `lay_page`): the baseline model's stack from the top margin
(moved by a tall header, 4.4), each line kept while it **fits** (4.1), and at the first
that does not, the break moved back by the **keeps** (4.2); the foot of the page holds
the footnotes (4.3). Everything is exact, in device px on Word's 1/4096 pt grid.
`tools/pages.py` scores it: Word's page tops (the first drawn line of each page, as
`(paragraph, line)`) against the model's, running free from the document's start;
paragraphs whose every line is on Word's page; and, as a separate count that takes
Word's page top as given, the *conditional* page ends. Every probe below is a
committed `make_*` / `read_*` pair, exported by Word 16.106 through `tools/oracle.py`,
and held offline by `tests/test_paginate.py` (and `test_page_top.py`,
`test_border_group.py`, `test_pages.py`).

### 4.1 What fits: the line's text and its bottom border, inclusive — measured to the unit

`make_page_fit_probe.py` / `read_page_fit_probe.py`. Phase 3's method turned through
90°: every case is a page — an anchor with `pageBreakBefore`, filler composed to the
layout unit (Calibri 11 lines at 55,000 units, one Times New Roman and one Calibri line
of chosen sizes, one `exact` line of chosen twips), and a **candidate** whose box ends
`slack` units above the bottom margin; the text area is exactly 13,960 twips. Slack is
swept by δ ∈ {±1024, ±205, ±41, ±9, ±3, ±2, ±1, 0} about each candidate threshold, in
eighteen families and four settings (no `settings.xml`, modes 12, 14, 15): 1,740 cases.

| Family | Threshold Word keeps the line from, in every setting |
| --- | --- |
| plain line; `exact` 400 and 200; `atLeast` 400; with space before; the second line of a paragraph; Times New Roman 14.5 | the bottom of its **pitch**, δ = 0 (δ = −1 goes) |
| with 240 twips of space after (also under `auto` 360, and with a border) | **the space after never counts** |
| with a bottom border (`w:sz` 12, 4 pt: 22,528 units) | **the border counts**: pitch + border |
| `auto` 360, a line's first or last | **the multiple's extra does not count**: the text's bottom (pitch − extra); with a border, text + border |
| `auto` 200 (text reaching below the pitch) | the pitch: the overhang does not count |
| the *first* line of a two-line bordered paragraph | modes 12, 14 and unstated: **text + border**; mode 15: the text |

**The rule** (`paginate.fits`, `FIT = "text-border"`): a line stays when `top + pitch −
auto-multiple extra (+ bottom border)` ≤ the bottom margin — inclusive, no slack. The
border is its paragraph's last line's, and below mode 15 any line's that ends a page
(`border_ends_every_page`). The first line on a page always stays.

| Hypothesis (1,740 cases) | Agree |
| --- | --- |
| **text and border, the border on every page-ending line below mode 15** (the model) | **1,740** |
| the border on every page-ending line in every mode | 1,725 |
| the border on the last line only, in every mode | 1,695 |
| the text and border, strict (`<`) | 1,668 |
| the overhang of a multiple below one counted | 1,680 |
| the text, no border | 1,519 |
| the pitch and border | 1,440 |
| the pitch, space after and border (the whole box) | 1,288 |
| the pitch alone | 1,275 |
| the pitch, strict | 1,233 |

The model places all 636 page tops of each of the four documents. On `sample4`, the
whole-pitch rule places 100 of its 168 page tops and the whole box 75.

### 4.2 The keeps — measured

`make_keep_probe.py` / `read_keep_probe.py`: 228 cases per setting (no `settings.xml`,
modes 12, 14, 15) and 27 in a document with **no styles part at all**. Every case starts a
page (anchor, filler, the paragraphs under test), every line Calibri 11, with `room` lines
left; lines split by `w:br` (a `natural` family wraps words, and agrees). **939 / 939**
cases put every line on Word's page.

* **Widow and orphan control is on unless stated off** — with no `settings.xml`, with no
  styles part, with nothing in `docDefaults` — and keeps **two** lines at each end: a
  paragraph that would leave one line at the foot moves whole; one that would take one
  line over takes two (a three-line paragraph with room for two moves whole). Off on the
  paragraph or its style, the page is filled.
* **`keepLines`**: the paragraph moves whole — **also one taller than a page**, which then
  starts the next page and breaks there (60 lines after 10 lines of room, and after the
  anchor alone).
* **`keepNext`**: a paragraph whose next one starts the next page goes with it — **whole
  below mode 15 and when unstated; in mode 15 only its last line**, with that line's keeps
  (widow control makes it two lines, `keepLines` the whole paragraph)
  (`paginate.keep_next_moves`). A chain moves together, **even one taller than a page**
  (60 one-line `keepNext` paragraphs, with and without a last paragraph ending it, after
  10 lines of room and after the anchor alone: moved to the next page, broken there).
* **Manual page breaks**: `w:br w:type="page"` ends its line; a page break alone in its
  paragraph keeps the mark on its line (drawn on the old page); after text at a
  paragraph's end, the next paragraph starts the page with no empty line; at a
  paragraph's start, the empty line before it is not drawn (`Para.undrawn`); a column
  break in a one-column section is a page break. **In mode 15 the line a page break ends
  is a last line to widow control** (three lines and a break, with room for two, go
  together); in 12, 14 and unstated it is not (`page_break_ends_paragraph`; found by the
  pagination probe, 4.6).
* **`pageBreakBefore`** on a paragraph that would start the next page anyway adds no
  page.

| Refuted (939 cases) | Agree |
| --- | --- |
| widow control off when unstated | 848 |
| no widow control | 751 |
| three lines kept at each end | 798 |
| `keepNext` moves the last line in every mode | 921 |
| `keepNext` moves the whole paragraph in every mode | 933 |
| a page break ends a paragraph for widow control in every mode / in none | 930 / 936 |

On `sample4`, which states no widow control anywhere, the model without it places 85 of
the 168 page tops.

### 4.3 Footnotes — measured before implementing them

`make_footnote_probe.py` / `read_footnote_probe.py`, the fit probe's layout with footnote
references on the candidate (or the anchor): a coarse sweep by 10,000 units, then fine
sweeps by single units about what it showed; notes in `FootnoteText` (Calibri 10, 50,000
units a line), separators Calibri 11; five documents (no `settings.xml`, 12, 14, 15, and
15 with 20 pt separators), 506 cases each.

**Word reserves the notes' room before it breaks the page**, and in mode 15 every case is
exact (`paginate._scan`, `note_minimum`, `note_prefix`):

* a line carrying references stays when below it fit the **separator paragraph's line**
  (its mark's natural height: 55,000 units at 11 pt, 100,000 at 20 pt — Word draws the
  separator in a line of its paragraph) and **every note line**, inclusive, to the unit;
  a note referenced higher on the page counts the same;
* a note's spacing adds, in mode 15, **the larger of its space before and after**
  (120/0, 0/120, 120/120 all add 120; 120/240 and 240/120 add 240); below 15, **before
  less after, not below zero** (`note_spacing`; the 240/120 threshold is not monotone
  there);
* a note may split where its **widow control** allows — two lines of five stay (one with
  widow control off) — and the rest goes to the next page's foot under the
  **continuation separator**; the next page holds that much less (48 lines instead of
  51);
* the last note takes as many lines as fit when its reference's line is placed, and
  **keeps them**: later lines fit above it (a reference at a page top keeps its whole
  note there);
* in mode 15, when the page that could not hold a note ended at a forced break
  (`pageBreakBefore`, a page break, a section), **the rest of the note gets a page of its
  own** (`continuation_page`); below 15 it goes under the next page's text.

| Cases on Word's page | none | 12 | 14 | 15 | 15, 20 pt separator |
| --- | --- | --- | --- | --- | --- |
| the model | 498 / 506 | 498 / 506 | 498 / 506 | **506 / 506** | **506 / 506** |

The eight misses below mode 15 are one behaviour: a line with **two** references, where
Word may leave the second note to the next page, whose foot it then shares with the next
case — the thresholds oscillate by 20,000 units. Not modelled. The parser now records a
reference (`footnoteReference:<id>`) and reads `word/footnotes.xml`; a reference is still
zero-width to the line breaker (its superscript number is not measured).

### 4.4 Headers and footers take room past their margin — measured

`make_header_probe.py` / `read_header_probe.py`: twenty sections, each with its own
header and footer and 60 lines of body, in four settings. **A header that reaches below
the top margin starts the body where it ends** — its distance from the edge plus its
height, its paragraphs' space before and after included — and a footer past the bottom
margin ends it there (`paginate._geometry`, `story_height`). Every section's first
baseline and line count is Word's (`header 740`: 348, not 344; `header 1500`: 506 and 49
lines; two lines with 240 after: 356; one with 480 before: 350), in every setting, and
37 / 37 page tops (34 without the rule). A negative top margin keeps the body where it
is (its first line is drawn on top of the header, which merges the two in the export: not
counted). A first-page header is read under `w:titlePg`; even-page headers are not.
*Since* Headers, footers and fields — measured, *even pages take the `even` story under
`w:evenAndOddHeaders`, a story's height is its own flow's stack (tables and borders
counted; this probe's heights unchanged), and the stories are drawn.*

### 4.5 Sections — measured

The pagination probe's sections document (four sections, Letter / A5 / A4, three margin
sets, at `nextPage`, `oddPage` and `continuous` breaks) and `make_section_probe.py`
(thirteen sections, each break kind after each parity, three settings, 15 / 15 page tops
and 13 / 13 sections): a new section starts a new page except at a `continuous` break;
**a continuous section that changes the page size starts a new page anyway**, one that
changes only the margins stays on it; **an `oddPage` or `evenPage` section that would
start on the wrong parity gets a blank page first** (`blank_page_before`). A page takes
the size and margins of the section its first line belongs to. Before these two rules,
13 / 15 page tops and 1 / 13 sections per setting. `w:pgNumType/@w:start` (a restarted
page count, which decides parity) is not read. *Read since* Headers, footers and fields —
measured *(H.2): the parity is the page's number's, and a restarted number of the wrong
parity is raised without a blank page.*

### 4.6 Found by the pagination probe, each then measured on its own

The first run of the done-when probe (4.8) missed 11 of 537 page tops (after justified
text was set aside). Each miss was traced to a rule, the rule was probed alone, adopted
only where the probe agreed, and committed with a before/after table:

* **The page-top space collapses across the break.** Where a page top keeps its space
  before (a section's first paragraph; `pageBreakBefore` below mode 15), it keeps what
  exceeds the space after of the paragraph before the break: after 200, 240 before keeps
  40. `make_page_top_collapse_probe.py` (after 0–360 × before 1–480 × both kinds × four
  settings): **724 / 724** baselines, against 308 keeping it whole
  (`vertical.page_top_gap_px(previous_after=…)`). Whether the box holds the kept part or
  the whole space is not discriminated. It moved no baseline of any other corpus.
* **Borders that are one box.** Consecutive paragraphs whose borders — every side's style,
  width, space and colour — and left indent are the same are drawn as one box: the top
  border over the first only, the bottom border under the last only, the `w:between`
  border (its space above and below the line) between each two. Paragraph spacing does
  not part them. A top border takes its space and width above the first line (it was not
  modelled at all), and a `double` border three widths. `make_border_group_probe.py`
  (26 cases, four settings, through `baselines.predict`): **408 / 408** baselines,
  against 280 (`vertical.paragraph_borders_px`). It moved no baseline of any other
  corpus.
* **Widow control ends at a manual page break in mode 15** (4.2): the keep probe's
  `manual segment` family, 28 cases per setting.

**And one for Phase 3: in mode 15 Word fits more words on a justified line** than the
left-aligned breaker puts there — 65 paragraphs of the nine mode-15 documents broke into
fewer lines than the model's; with no `settings.xml` and in modes 12 and 14 none did
(Phase 3.2's 85 justified lines, which broke as left-aligned, were in documents with no
`settings.xml`). Whether it compresses spaces, and by how much, is Phase
3's justification measurement; the probe was made without justified text. *Measured:
Phase 3.8 -- a quarter of each space, and at most half the stretch.*

### 4.7 Inline pictures — measured

`sample4` and `sample-long` hold 39 and 18 paragraphs of inline pictures (`wp:inline`),
where the paginator first stopped. `make_picture_probe.py` / `read_picture_probe.py`: 50
pictures per rule (`auto` 240, 259, 276, 360; `exact` 300; `atLeast` 300), heights in
non-pixel steps, with and without an effect extent, under a 20 pt mark, two side by side,
in four settings; the baseline of the line after each picture — **800 / 800**
(`lines.object_line_height`):

* the line is **the picture's height** (extent plus effect extent) **in whole twips,
  truncated** (`linebreak.emu_twips`), with **nothing below it** — the mark adds no
  descent; below mode 15 (and unstated) at least the mark's natural height;
* an `auto` multiple adds its extra **over the mark's natural height**, not the
  picture's; `exact` is the line; `atLeast` at least the line;
* the breaker gives a picture its width (extent plus effect extent).

| Refuted (800 lines) | Agree |
| --- | --- |
| the mark's descent below the picture | 120 |
| the multiple scaling the whole line | 400 |
| the mark as a floor in mode 15 too / nowhere | 799 / 797 |
| the height rounded to the twip / kept exact | 676 / 712 |

Text beside a picture on one line is not measured (its extent is maxed with the
picture's). A drawing that is not a single `wp:inline` — anchored (floating), VML, an
embedded object — is still an obstacle, and the baseline scorer still leaves picture
paragraphs out.

### 4.8 Done when: every paragraph on Word's page

`make_pagination_probe.py` / `read_pagination_probe.py`: one seeded body of 260
paragraphs — Calibri, Times New Roman, Arial, Cambria and Georgia at six sizes; `auto`
240, 276 and 360, `exact` 300, `atLeast` 320; space before and after; `keepNext` headings
(chains where two meet); `keepLines`; widow control off; bottom borders; numbered items;
empty paragraphs; `pageBreakBefore` headings and manual page breaks.

| Document | Page tops | Paragraphs on Word's page |
| --- | --- | --- |
| Letter / A4 / A5 × 1 in all round (mode 15) | 34 / 34, 33 / 33, 69 / 69 | 260 / 260 each |
| Letter / A4 / A5 × 0.5 in all round | 30 / 30, 28 / 28, 52 / 52 | 260 / 260 each |
| Letter / A4 / A5 × 1800 top, 1080 bottom, 2160 left, 1440 right | 36 / 36, 35 / 35, 78 / 78 | 260 / 260 each |
| Letter, 1 in: no `settings.xml`, mode 12, mode 14 | 34 / 34 each | 260 / 260 each |
| four sections (nextPage, oddPage, continuous; three sizes, three margin sets) | 40 / 40 | 260 / 260 |
| **All** | **537 / 537** | **3,380 / 3,380** |

An equality, not a tolerance: no line of any paragraph is on another page. (The same
documents' baselines, a side result: 15,343 / 15,393 through `baselines.predict`, every
miss 1 px, most at page tops and on bold headings — recorded, not pursued here.)

### 4.9 The real documents, with nothing from the oracle

`tools/read_pages.py`, `tests/test_pages.py`. The model runs on the file alone (installed
faces' advances and metrics, recorded); Word's pages are the lines recorded for the
baseline scorer.

| Document | Page tops placed where Word placed them | Never reached | Paragraphs on Word's page |
| --- | --- | --- | --- |
| **`filesamples/sample4`** (175 pages, 104 tops inside a paragraph, 39 picture paragraphs; not committed) | **168 / 168** | 0 | 651 / 651 |
| `samplelib/sample-long` (17 manual breaks, 18 pictures) | **35 / 35** | 0 | 127 / 127 |
| `wordto/sample-10pages` (two natural breaks, a table at the end) | **2 / 2** | 0 | 36 / 36 |
| `layout-sweep.docx` (Phase 0: "Word put the break after L32") | 2 / 2 | 0 | 88 / 88 |
| `style-document.docx` | 1 / 1 | 0 | 30 / 30 |
| `samplelib/sample-simple`, `wordto/sample-5pages` | 0 / 1 each | 1 each (a table first) | all |
| `filesamples/sample1` | 2 / 7 | 5 (tables) | 16 / 16 |
| `filesamples/sample3` | 0 / 1 | 1 (a floating drawing) | 18 / 18 |
| one-page documents (`sample-blank`, `-resume`, `-1page`, `-with-images`, `-with-table`, `sample2`) | no page top | | all |
| **All** | **210 / 218** | **8, all past a table or floating drawing** | **1,034 / 1,034** |

**Every page top the model reaches, it places where Word does.** The eight it misses are
past an obstacle — tables (Phase 6) or a floating drawing — and are counted as out of
scope rather than fitted. The local corpora are scored on the machine that holds them;
their numbers stay there.

### 4.10 Constants, open questions

| Constant | Value | Evidence | Residual |
| --- | --- | --- | --- |
| Fit | text + border ≤ bottom, inclusive | 1,740 / 1,740 cases, δ to one unit | 0 |
| Widow / orphan lines | 2 | 939 / 939 keep cases | 0 |
| Picture height | EMU // 635 twips | 800 / 800 | 0 |
| Footnote separator | its paragraph's line | 1,012 / 1,012 mode-15 cases | 0 |

Not measured, or not modelled, and left:

* **Footnotes below mode 15** with two references on a line (8 cases in 506 per setting),
  a note's width to the breaker (its superscript number), notes of several paragraphs,
  endnotes, and footnote continuation notices.
* **Justification in mode 15** (4.6), **tables** (Phase 6: the paginator stops at the
  first; *since* Tables — measured *it places a table row by row, splitting rows at a
  page's foot; floating and nested tables still stop it*), **floating drawings**, ~~**multi-column sections**~~ (`w:cols`; *laid out by
  column, balanced and ruled since 4.14*), text beside an inline picture, pictures wider than the column.
* A group of bordered paragraphs split by a page (which border the fragment gets below
  mode 15), right and left borders, shading.
* ~~Even-page headers (`w:evenAndOddHeaders`), `w:pgNumType/@w:start`~~ *(H.2)*, a header that
  pushes the body and also holds footnotes' room.
* `keepNext` onto a table, a section break inside a `keepNext` chain, `pageBreakBefore`
  inside a `keepNext` chain beyond the one case measured (the kept paragraph stays).
* The conditional count's one artefact (`sample4`, a page that starts with four pictures
  Word draws nothing for).

### 4.11 A section break's empty paragraph — measured

Found by local documents (F.4): a section that ends with a table and then the empty
paragraph holding its `w:sectPr` (a `nextPage` break, mode 15) made a page more in the
model than in Word -- the paragraph did not fit, and the model gave it a page of its own.
**Probed before it was implemented**: `make_section_end_probe.py` /
`read_section_end_probe.py`, `tests/test_section_end.py`
(`tests/fixtures/section-end-observations.json`). 192 cases per document, each a section
of its own: an anchor line, filler composed to the layout unit (4.1's), and a candidate
paragraph whose fit extent ends 41 units above the bottom margin, 1 and 41 units below
it, a quarter of its height below, with its top 41 units above the margin, on it, and
12 pt and 60 pt below it (pushed there by the filler's space after). Families: an 11 pt
mark, a 16 pt mark under `auto` 276 (the documents' form), a 48 pt mark, 240 twips after,
240 twips before, a table before it, and a pair of empty paragraphs (the first swept).
Each followed by a `nextPage` break, a `continuous` one, nothing (`mid`: the paragraph
does not end its section) and with text in the candidate. Separate documents for the
document's own last paragraph (the body's `w:sectPr`). No `settings.xml`, mode 14, mode
15.

**The rule** (`paginate.section_mark_room`, `Para.section_mark`):

* **An empty paragraph that holds a section break never goes to the next page** --
  however far past the margin it would reach, every family, both breaks, every setting.
  The document's own last paragraph is not one: it goes to the next page as any line
  does (Word makes the second page); nor is a paragraph with text there, nor an empty
  paragraph that does not end its section.
* **And it takes no room**: after a `continuous` break the next section's first line
  starts where the mark would have (its space before and after not counted either) --
  a 16 pt mark 1 unit over the margin leaves the next 11 pt line on the page, in its
  place. **Except in mode 15 straight after a table**, where the next line does not
  start beside it: there the mark is placed as it stands, and still never moved.

| `section-end` | model agrees before | after | pages (Word) before | after |
| --- | --- | --- | --- | --- |
| no settings | 91 / 192 | **185 / 192** | 464 (409) | **409** |
| mode 14 | 91 / 192 | **185 / 192** | 464 (409) | **409** |
| mode 15 | 92 / 192 | **185 / 192** | 464 (410) | **410** |

~~The seven left are one other thing, recorded and not modelled: an empty 48 pt
paragraph that does not end its section, moved to the next page's top, stands 146 px
high there where its line is 244.~~ **Since: 192 / 192 in every setting; the seven were
the reader's.** Word draws an empty paragraph's mark as a space, and `baselines.merge_raised`
(which folds a superscript's smaller runs into the line they belong to) folded the next
11 pt line, 98 px below the 48 pt mark and so within half its size, into the mark's line:
the "146 px" was the mark's baseline read as the next line's. Read apart
(`read_section_end_probe.word_lines` now leaves whitespace-only lines out before the fold),
Word stands the paragraph its full 244 px line, as the model always did; the recording
changed in those 21 outcomes only, and nothing in the model.
Every document-end case is Word's (14 / 14 per setting). No committed document has such a
paragraph past a page's foot: their SVG output, VRT snapshots included, is
byte-identical (the mark is still drawn where it stands), and their fidelity with it. The local corpora's three extra pages are gone
(F.4).

### 4.12 Endnotes — measured, laid out and drawn

An endnote reference (`w:endnoteReference`) was a mark the paginator read nothing from,
and `word/endnotes.xml` was not read. `make_endnote_probe.py` / `read_endnote_probe.py`,
21 documents: `base` (references mid-line, near a line's end, adjacent, in a 24 pt
paragraph, one not in `EndnoteReference`; notes of one and three lines, of two
paragraphs, in `Normal`, with space before and after, with a number not superscript),
`overflow` (notes that go on to the next page), `format` (three sections' own
`w:endnotePr`: `decimal` from 5, `upperRoman` restarting, `lowerLetter` from 3
restarting), `normal` (a Normal style as a filesamples document has it: 12 pt, first
lines indented 432, `auto` 276), `separators` (the separators' paragraphs formatted:
20 pt Times New Roman indented 720 with 120 before, 8 pt right-aligned), each with no
`settings.xml`, in mode 14 and mode 15; and, needing the settings part, `sectend`
(`w:pos` `sectEnd`, three sections), `docpr` (the document's `w:endnotePr` format and
first number) and `unnamed` (`separators` under a `w:endnotePr` that names no separator
note), in modes 14 and 15. The notes are `EndnoteText` (10 pt), their numbers
`EndnoteReference` (superscript), as Word writes them.

**The rules** (`docx2svg.notes`, `linebreak.note_number_units`,
`paginate.continuation_line`, `layout._Placer._separator`):

* **The number** is the note's place among the references, in the format of the
  **section's** `w:endnotePr` (`lowerRoman` unstated); the first section and one that
  restarts (`eachSect`) start at its `w:numStart`, the others go on counting. The
  document's `w:endnotePr` format and first number are not taken (modes 14 and 15:
  `upperLetter` from 2 there, the notes `1`, `2`, `iii`). It is drawn in the reference
  run's properties (superscript by the style, or not), and the note's own number
  (`w:endnoteRef`) in its run's.
* **A note's number advances by whole device pixels**: each glyph's advance at the size
  it is drawn at (whole px) rounded -- Calibri `i` 7 px at 7 pt (6.66 unrounded), 6 px at
  6.5, 15 at 15.5, `v` 13 and 12 -- and each glyph is drawn at its pen position rounded
  to a whole px. The text after it follows the rounded advances (0.6 px further after
  `ii`), in the body and in the note.
* **Where**: straight after the document's last paragraph, on its page, stacked as body
  paragraphs are -- the separator's paragraph, then each referenced note's paragraphs in
  the order referenced; with `sectEnd`, each section's notes after its last paragraph,
  before its break. The notes are paragraphs of the flow (`notes.with_endnotes`), so
  breaking, widow control and pagination are the body's.
* **The separator** is the endnotes part's `separator` note **when the settings'
  `w:endnotePr` names it** (`w:endnote/@w:id`, as Word writes); with no settings part, or
  one that names none, Word draws its own -- a paragraph of the default style, no space
  after, single spacing -- and the part's formatting is not taken (`separators`,
  `unnamed`: 20 pt, the indent and the space before gone). **In mode 15 the separator's
  space before is not counted** (120 twips: the notes 25 px higher than in mode 14).
* **Its line is a strikethrough of its run** -- the face's `OS/2` strikeout position
  above the baseline and strikeout size thick, at the size the face is drawn at, each
  rounded: `[B-12, B-9]` at Calibri 11, `[B-21, B-17]` at Times New Roman 20, `[B-8, B-6]`
  at Calibri 8 -- **2,880 twips long** from where the paragraph's text starts (its indent
  counts, its face does not); the **continuation separator's** from there to the
  column's right edge, right-aligned or not.
* **A page the notes go on to** starts with the continuation separator's line, and the
  notes under it (`Para.continuation`).

| Lines within half a pixel (separator px Word only) | before | after |
| --- | --- | --- |
| `base` (47 lines, each setting) | 11 (1,800) | **47 (0)** |
| `overflow` (72) | 41 (7,527), 1 page of 2 | **72 (0)**, 2 |
| `format` (21) | 3 (1,800) | **21 (0)** |
| `sectend` (24, modes 14, 15) | 3 (5,400) | **24 (0)** |
| `docpr` (12, modes 14, 15) | 3 (1,800) | **12 (0)** |
| `separators` (72) | 41 (7,527 / 6,218) | **72 (0)** |
| `normal` (11) | 0 (1,800) | **11 (0)** |
| `unnamed` (72, modes 14, 15) | 41 (7,527) | **72 (0)** |

Not measured, and left: a custom mark (`w:customMarkFollows`: no number drawn, as the
file intends), the continuation notice, a separator paragraph that is not left-aligned
(the 2,880-twip line is drawn from the pen), a note that is a table, a section's own
`w:endnotePr/w:pos`, and whether a note's space before counts under the continuation
separator (each note here has none). Footnotes still take their room and are not drawn
(4.3); their reference's number and their separators are what this would give them.
Pinned in `tests/test_endnotes.py`. *Since 4.13 footnotes are drawn.*

### 4.13 Footnotes drawn — measured

A footnote took its room at the page foot (4.3) and nothing of it was drawn: not its
reference's number (the rest of its line stood that much left of Word's), not the note,
not the separator. `make_footnote_draw_probe.py` / `read_footnote_draw_probe.py`, 44
documents: `base` (a short page: references mid-line, near a line's end, adjacent, in a
24 pt paragraph, one not in `FootnoteReference`; notes of one and three lines, of two
paragraphs, in `Normal`, with 120 twips before and after, with a number not
superscript), `overflow` (a page nearly full, a six-line note and one of 25 lines that
goes on to the next page), `format` (four sections' own `w:footnotePr`: `decimal` from 5,
`upperRoman` restarting, `lowerLetter` from 3 restarting, `chicago` going on), `chain`
(a section stating nothing, one from 7 going on, one restarting, one stating nothing),
`eachpage` (`w:numRestart` `eachPage`, four pages of 2, 3, 9 and 2 references),
`separators` (`overflow` under separators formatted: 20 pt Times New Roman indented 720
with 120 before, the continuation separator 8 pt right-aligned), `normal` (a Normal style
as a filesamples document has it), `sectbeneath` (the section's `w:pos` `beneathText`),
each with no `settings.xml`, a settings part naming the separators but stating no mode
(as that document's does), mode 14 and mode 15; and, needing the settings part, `docpr`
(the document's `w:footnotePr`: `upperLetter` from 2), `docpage` (the document's
restarting each page), `unnamed` (`separators` under a `w:footnotePr` naming no separator)
and `beneath` (the document's `w:pos` `beneathText`). Notes in `FootnoteText` (10 pt),
numbers in `FootnoteReference` (superscript), as Word writes them.

**The rules** (`notes.with_footnotes`, `notes.footnote_numbers`,
`notes.footnote_separator_kind`, `paginate.Pagination.notes`, `layout._Placer._footnotes`),
in every setting alike but where it says:

* **The number** is drawn where the reference stands, in its run's properties, as an
  endnote's is (4.12: each glyph advancing by whole device pixels and drawn on one), and
  the note's own (`w:footnoteRef`) in its run's; the text after it follows.
* **Numbering** is the **section's** `w:footnotePr`: `w:numFmt` (`decimal` unstated),
  `w:numStart` (1), `w:numRestart`. The document's `w:footnotePr` in the settings is not
  taken -- neither its format nor its first number (`docpr`: 1, 2, 3 under `upperLetter`
  from 2), its restart (`docpage`: 1 to 16 over four pages) nor its `w:pos` (`beneath`: at
  the page foot). A section that does not restart numbers a note by **its place among all
  the document's references**, from its own start: `chain`'s fourth section, after 1, 2,
  9, 10 and 1, 2, numbers its notes 7 and 8, and `format`'s `chicago` section, after nine,
  draws `†††`, `‡‡‡`, `§§§` (10 to 12). `eachSect` counts in the section, `eachPage` on
  the page, each from `w:numStart`; a page's numbers move its lines, so the layout is
  done again with the pages found until they hold (as for page-number fields).
* **The notes are drawn at the page foot**: the separator's paragraph, then the notes'
  paragraphs (the rest of a note the page before could not hold first, then every note
  referenced on the page, the last as far as the page holds it), laid out as a story's
  paragraphs and **stacked as the body stacks them**, the **stack's foot on the text
  area's foot** -- on a short page too (`base`: the notes at the bottom margin, some 1,500 px
  below the text). The stack is not the room the paginator keeps: a note's space before
  and after are both drawn (`base`: the spaced note 25 px under the one above it and the
  next 25 px under it), where the room counts at most one of them (4.3).
* **Under the section's `w:pos` `beneathText`** the stack's top is where the body's stack
  ends on the page (its last line's pitch; its space after is not in this probe).
* **The separator** is the footnotes part's when the settings' `w:footnotePr` names it,
  Word's own otherwise (as for endnotes, 4.12); its line 2,880 twips. On a page a note goes
  on to, the **continuation separator**, from the text's start to the column's right edge
  -- **except in mode 15, where that page opens with the separator itself** (`overflow`,
  `separators`: the 2,880-twip line, the part's 20 pt one where it is named). The
  paginator keeps the same separator's room there.

| Lines within half a pixel (separator px Word only), each setting | before | after |
| --- | --- | --- |
| `base`, `beneath`, `sectbeneath` (32 lines) | 5 (1,800) | **32 (0)** |
| `overflow`, `separators`, `unnamed` (100) | 61 (3,600 to 7,527) | **100 (0)** |
| `format` (36) | 4 (7,200) | **36 (0)** |
| `chain` (28) | 4 (7,200) | **28 (0)** |
| `eachpage`, `docpage` (64) | 0 (7,200) | **64 (0)** |
| `docpr` (12) | 3 (3,600) | **12 (0)** |
| `normal` (11) | 0 (1,800) | **11 (0)** |

Every glyph Word draws is drawn (507 not drawn in `base` before, 0 after), every page count
Word's. The committed documents have no footnotes: fidelity as recorded. A filesamples
document that has one draws it now (205 more glyphs matched, every one on Word's
position); on its page 5 the line holding the reference is Word's to 0.011 px but for the
two superscript numbers, 1 px right of Word's: their pens are at 624.5004 and 965.5004 px,
and Word's -- 0.005 px left of the model's there, as on the document's other mode-12 lines
-- round down.

Not measured, and left: a footnote referenced in a table cell (it takes no room; warned
`footnotes-not-drawn`; *laid out and drawn since 4.19*), the continuation notice (`footnote-continuation-notice-not-drawn`
where it has text), a note that is a table, `beneathText` under a paragraph with space
after, `eachPage` with a `w:numStart`, and footnotes in more than one column. Pinned in
`tests/test_footnote_draw.py`.

**How the separator appears in the SVG** (documented, not changed; docx-agent's
proposal): as a group of its own, like a paragraph's, with one line group holding no
text, and the line as a `<rect data-docx-kind="separator">` whose path is the group's run
(`.../w:r[1]`). Where Word's own separator is drawn -- the settings name none -- the
group's path is `w:footnotes/separator` (`continuationSeparator`, and `w:endnotes/...` for
endnotes): it names no element (`resolve_path` gives `None`) and the group has no
`data-docx-id`. Where the settings name the part's own, the path is that note's paragraph
(`w:footnote[@w:id=-1]/w:p[1]`). A caller matching drawn paragraphs to the document's
skips the first kind.

### 4.14 Text columns — measured

A section of several columns (`w:sectPr/w:cols`) stopped the layout at its first
paragraph. `make_columns_probe.py` / `read_columns_probe.py`, 14 documents in three
settings (no `settings.xml`, mode 14, mode 15), Calibri 11 pt on A4 with the left margin
off the pixel grid: `geometry` (nine sections whose column edges fall on every fraction of
a pixel: two columns 720 and 355 apart, three 401 apart, four 113, three touching, five
250 apart, unequal ones whose widths sum to the text width and do not, one with a gutter),
`flow2`, `flow3` and `unequal` (running paragraphs with headings over two, three and
unequal columns of 2200, 4000 and the rest), `keeps` (an orphan, a widow, a `keepNext`
heading, a `keepLines` paragraph and widow control off at the first column's foot, an
orphan at the second's), `breaks` (column breaks in a line, at a paragraph's start, alone,
twice, in the last column, before space before; a page break and `pageBreakBefore` in a
column), `sep` (`w:sep` over two, three and unequal columns), `balance` (sections of two
and three columns between continuous one-column sections: lines, spaced paragraphs, a
`keepLines` paragraph, a 16 pt line, one and two lines, a column break, running words),
`balancelong` (a balanced section of 140 lines over a page), `regions` (the space around a
continuous break: space after and before either side, the empty paragraph Word writes the
break in, columns ending in space after), `table`, `picture`, `header` and `footnote`. The
settings agree but where stated below.

* **Geometry.** Equal columns share the text width less the spaces between them,
  **truncated to a whole twip** (4,404 of 4,404.5; 1,632 of 1,632.8: the last column's right
  edge is where that width puts it); unequal columns are as stated, whatever their sum.
  Each starts where the one before ends plus the space after it (`w:space`, or each
  `w:col`'s own). A column's left edge is the margin, gutter and columns before it **rounded
  to a whole pixel -- below mode 15 from the exact length, in mode 15 from the length held
  in whole layout units**: 3,324 twips is 692.5 px exactly, drawn at 693 below mode 15 and
  692 in it (680,755.2 units held as 680,755). Its right edge is its width further
  (`linebreak.column_boxes`, `layout.column_edges`).
* **Flow.** The first column, then the next, then the next page. **Widow and orphan
  control, `keepNext` and `keepLines` hold at a column's foot as at a page's**, and a
  column's first line is placed whatever its height. A paragraph in unequal columns is as
  wide as the column each line is in: it is broken again where it goes on into a column of
  another width (`Para.fit_width`).
* **Breaks.** A column break ends its column, and in the last column the page. **The mark
  of a paragraph a column break ends goes on a line of its own at the next column's top**
  (a page break's stays on its line, 4.6); a column break at a paragraph's start leaves
  an empty line before it, not drawn. A page break or `pageBreakBefore` in any column
  goes to the next page.
* **The separator** (`w:sep`): a black line 0.75 pt wide **centred on the space between
  two columns, exactly** (their unrounded edges), its edges rounded to whole pixels -- 1,253
  to 1,257 px between two columns 720 apart, 921 to 924 between three 401 apart, 809 to 812
  after a column of 2,000 and a space of 900 -- **from the region's top to the foot of its
  longest column** (313 to 3,166 on a full page, 368 to 648 over five balanced lines), and
  **only before a column that holds text** (none on a page whose second column is empty).
* **Balance.** A section of several columns that ends on a page with a **continuous**
  section after it is balanced: **the least height at which its columns, filled in turn
  under the keeps, hold it all, no column passed over empty**. Seven lines go 4 and 3; ten
  one-line paragraphs in three columns 4, 4 and 2, four 2, 2 and 0; a two- or three-line
  paragraph stays whole in the first column; paragraphs of 9 and 7 lines go 9 and 7 (8
  would leave a widow); a `keepLines` paragraph moves whole. A section followed by one that
  starts a page, and the document's last, are not balanced. A balanced section longer
  than a page fills its pages and is balanced on its last (`paginate._balance`).
* **Regions.** Continuous sections share a page one under the other. A region of
  columns under another starts **every column level with its first column's first
  paragraph**, past the space between that paragraph and the one before; a later column's
  space before at its top is dropped, as at a page's. **Under a region of columns the
  stack goes on, in mode 15, under its tallest column with that column's last space
  after; below mode 15, under the taller of its tallest column (no space after) and the
  mean of its columns each with its last space after** -- four lines and 480 twips after
  over two columns end two lines and 240 twips down; three lines and 360 after beside two
  and 120 after end 190 px down where the columns reach 168 and 112 -- and the next
  paragraph's space before collapses with the region's last space after, which the foot
  holds. The empty paragraph a break is written in takes no room and no part in either
  (`paginate.region_foot`). A column of a region under another may be left empty by its
  keeps.
* **Content.** A table in a column is laid out in it (in percent of the column), its rows
  going on into the next column, every border pixel Word's; inline pictures are lines as
  anywhere; a header and footer span the page's text width.

| Lines within half a pixel (each setting) | Before | After |
| --- | --- | --- |
| `geometry` | 0 / 27 | **27 / 27** |
| `flow2`, `flow3`, `unequal` | 0 / 840 | **840 / 840** |
| `keeps`, `breaks` | 0 / 334 | **334 / 334** |
| `sep` (separator px Word only) | 0 / 229 (39,942) | **229 / 229 (0)** |
| `balance`, `balancelong` (separator px Word only) | 3 / 150 (1,120) | **150 / 150 (0)** |
| `regions` | 1 / 71 | **71 / 71** |
| `table` (border px Word only), `picture`, `header` | 4 / 216 (37,518; 37,482 in mode 15) | **216 / 216 (0)** |
| `footnote` | 1 / 56 | 0 / 56 (stops); **56 / 56 since 4.15** |

Every page count Word's but `footnote`'s, which stops on its first page (*laid out since 4.15*). The real
document: a filesamples document whose second page ends in a continuous two-column section
broken by a column break was drawn to that section: 1,766 of Word's 2,174 glyphs; now
**2,174 / 2,174**, every one on Word's baseline, two pages as Word's. No committed
document has columns: fidelity and the local corpora as recorded.

**Footnotes in columns -- measured, not laid out by column**: below mode 15 each column's
notes stand at its own foot under its own separator (the second column's at 1,330 to 1,930
px); in mode 15 the page's notes run on through the columns at the page's foot under one
separator. A paragraph referencing a footnote in a section of columns stops the layout
(`layout-stopped:columns`), as do a floating drawing, a floating table, a drop cap or an
endnote there -- not measured. *Footnotes in columns are laid out since 4.15; drop caps,
and endnotes below mode 15, since 4.16; floating drawings positioned against their column
since 4.17.* Not measured either: a `nextColumn` section, two continuous
sections of the same columns (laid out as two regions, the first balanced), the space
before at a later column's top in a region at a page's top other than after a full
column. Pinned in `tests/test_columns.py`.

### 4.15 Footnotes in text columns — measured, laid out and drawn

4.14 measured one document of footnotes in two columns and left it stopped. Five more
documents, in the same three settings, in `make_columns_probe.py`: `notesright` (notes
referenced only from the second of two columns), `notes3` (three columns, notes from each,
two of them on one line at the first column's foot), `notesunequal` (columns of 2,200,
4,000 and the rest, a note from each), `notesover` (a note of 21 lines referenced near
the first column's foot, and the page after) and `notesregion` (a one-column region, a
balanced two-column region and a one-column region on one page, a note from each). No
settings and mode 14 agree throughout. **Word has two arrangements:**

* **Below mode 15, each column's foot is a page's** (`paginate._scan`, a column of
  several). Its notes stand at its own foot, broken in its width (`notesunequal`: three
  or four words a line in the 2,200-twip column, every line where Word breaks it), stacked
  as a page's are with the stack's foot on the text area's foot, under their own
  separator, **whose 2,880-twip line is cut at the column's right edge** (581 px in a
  column of 2,787 twips, 458 in one of 2,200, 472 in one of 2,264; 600 in a column of
  4,222). A column with no notes has no separator (`notesright`: the first column holds
  all its 51 lines). **The room is the column's own**: a line with references fits when
  it and, under it, the separator, the column's notes so far and **the least of the
  line's first new note** fit; **its later notes take what is left in turn, and the first
  that does not fit whole, with every one after it, goes on** (`notes3`: the line at the
  first column's foot keeps two lines of its first note and none of its second). **What a
  column cannot hold opens the next column's foot**, under the continuation separator,
  which runs to that column's right edge (`notesover`: 14 lines at the first column's
  foot, 7 at the second's under a line 880 px long, then the second column's own note);
  what the last column cannot hold goes to the next page, as on a page of one column.
* **In mode 15, the page's notes run on through its columns** (`paginate.note_area`,
  `paginate._lay_columns_page`): the notes referenced on the page, in order, are a flow of
  their own in the section's columns, **balanced** as a region above a continuous break
  is (4.14: the least height at which the columns, filled in turn under widow control,
  hold them all) -- 7 lines go 4 and 3, 10 lines 5 and 5 (a note of 7 split 5 and 2),
  11 lines in three columns 4, 4 and 3 -- **every column's top level, the tallest's foot
  on the text area's foot**, and a note going on into a column of another width broken
  again in it (`notesunequal`: 5, 4 and 4 lines). **One separator**, at the first
  column's left, **2,880 twips whatever the column's width** (600 px over a column of
  2,200 twips). **Every column ends above the area**, whatever it references:
  `notesright`'s first column, which references nothing, ends after 45 lines, six short
  of its 51, above the second column's notes balanced over both; `footnote`'s first
  column ends after 46 lines where its own notes alone, balanced, would leave room for 47.
  The model lays the page out again with the room the notes its columns reference take,
  until that room holds (the larger kept should it not settle).

| Lines within half a pixel (rule px Word only), each setting | Before | After |
| --- | --- | --- |
| `footnote` (56; 56 / 56 / 56) | 0 (1,800 to 3,600) | **56 (0)** |
| `notesright` (66 / 66 / 56) | 33 / 33 / 20 (1,800) | **all (0)** |
| `notes3` (56) | 0 (5,229; 1,800 in mode 15) | **56 (0)** |
| `notesunequal` (58 / 58 / 55) | 0 (4,590; 1,800) | **all (0)** |
| `notesover` (111 / 111 / 106) | 0 (4,440; 1,800), one page of Word's two | **all (0)**, two pages |

Every glyph drawn where Word draws it, no extra glyph, every page count Word's.
**`notesregion` stays a stop**: below mode 15 Word sets each of the three sections on a
page of its own (three pages where the text would fill a third of one), and in mode 15
it sets the page's three notes in one column across the text width under the
one-column sections and the balanced region; neither is modelled, so **a footnote
referenced in a section of columns that a continuous break joins to another** stops the
layout (`layout-stopped:columns`, "a footnote, in a section a continuous break joins to
another"). Stops too: in mode 15, a note going on from the page before onto a page of
columns with notes, and notes taller than the page -- neither measured. Not measured:
`beneathText` in columns (the notes are drawn at the foot), notes going on from one column
into another of a different width (the rest keeps the first column's lines), spaced notes
in mode 15's area (stacked as the body stacks them), a line with two notes below mode 15
on a page of one column (4.3, still not modelled there). The committed documents have no
columns: fidelity and the local corpora as recorded. Pinned in `tests/test_columns.py`.

### 4.16 Endnotes and drop caps in text columns — measured

Two more documents of `make_columns_probe.py`, in the same three settings: `endnotes`
(three endnotes, one of two paragraphs, referenced from both of two columns) and
`dropcap` (drop caps of three lines in columns: under two lines at the first column's
top, at the second column's top after a column break, and a `margin` drop cap under it).

* **Endnotes, below mode 15, go on in the columns after the text** as paragraphs of the
  flow (4.12), their separator's line in the column it falls in (1,330 to 1,930 px):
  **67 / 67 lines** (54 before), every glyph Word's. **In mode 15 Word balances the text
  above them**, as above a continuous break, sets the separator under it at the first
  column's left, and the endnotes under it as **a balanced region of their own** (five
  lines and five); not modelled -- an endnote in columns in mode 15 still stops the
  layout.
* **A drop cap in a column stands at that column's start**, and the text beside it
  starts from there as from a page's margin (4.14 and F.19's rules, the column's rounded
  edge for the margin's): the column-2 cap at 1,330 px, its text at 1,515. **A `margin`
  drop cap anchored to the text, in the second column, stands at the column's start as a
  `drop` one does** (its text at 1,508 px). **A drop cap is positioned where the text
  above it ends with the drop caps above it in place**: one under another in the same
  column is a line lower than the page laid out without the first puts it, so a page of
  columns with a drop cap is laid out again with its frames until they hold
  (`paginate._lay_columns_page`; a page of one column is still positioned from the page
  laid out without them, as measured in F.7). **33 / 33 lines** in every setting (1
  before).

Not measured: a `margin` drop cap anchored to the page in a column, a drop cap at a
column's foot, endnotes going on to another page in columns. Pinned in
`tests/test_columns.py`.

### 4.17 Floating pictures in text columns — measured

One more document of `make_columns_probe.py`, in the same three settings: `floating` --
in the first of two columns a picture text goes above and below (`wrapTopAndBottom`) at
the column's left; on the next page, in the first column, one text goes around
(`wrapSquare`) aligned right in the column, and one behind the text 360 twips into it.
`read_columns_probe.py` now also records the floating pictures' boxes (in `floating`)
and scores the model's (`read_anchor_probe.compare_boxes`).

* **A drawing positioned against its `column` is positioned against the column it is
  anchored in** -- its left edge and width -- not the page's text area: the picture
  aligned right ends at 1,180 px, the first column's right edge, where the text area's
  ends at 2,210 (`floating.Frames.column`; `paginate._frames`, `layout._Placer._frames`).
  `margin` stays the page's text area.
* **A drawing text goes above and below keeps lines off the columns it is across, not the
  page's**: the second column's lines stand where they would without it, the first's go
  under it (`paginate.Band.span`).
* **Text goes around a drawing in a column as beside one on a page**, the column's edges
  for the margin's (`paginate.wrap_column`).

| Each setting | Before | After |
| --- | --- | --- |
| Lines within half a pixel | 0 / 111 (stops), 26 with the stop lifted | **111 / 111** |
| Pictures where Word draws them | -- | **3 / 3** |

Every glyph Word's, both pages. **A floating drawing positioned against the margin, the
page or a character in a section of columns** (or by `simplePos`) is not measured and
still stops the layout, as does a floating table there. Not measured either: a drawing
across two columns. Pinned in `tests/test_columns.py`.

### 4.18 Line numbers (`w:lnNumType`) — measured, drawn

A section's line numbers were neither drawn nor warned (docx-agent's E4 proposal).
`make_line_number_probe.py` / `read_line_number_probe.py` (`tests/test_line_numbers.py`):
six sections -- `w:countBy` 1 restarting at each page (unstated), over two pages; `countBy`
5, `w:start` 3, `w:distance` 720, `continuous`; `countBy` 2, `distance` 200, `newSection`;
none; `start` 4, `newSection`; and the first again -- each holding wrapped and short
paragraphs, an empty one, one of 20 pt, centred, right-aligned and indented ones, one under
`w:suppressLineNumbers`, a table, a header and a footer. No `settings.xml` and mode 15,
which agree. **The rules** (`layout._Placer._line_numbers`):

* **Which lines**: every line of the body's own paragraphs, an empty paragraph's too.
  Not a table's, a header's or a footer's (nor, by extension, a note's or a text box's),
  nor a `w:suppressLineNumbers` paragraph's -- none of which is counted either.
* **The count** starts at **`w:start` + 1** (unstated, 1) and restarts at each page
  (`newPage`, unstated) or each section (`newSection`); **`continuous` goes on from the
  section before** (24 after a page that ended at 23, its own `w:start` not read). A number
  is drawn where the count is a multiple of `w:countBy`.
* **Where**: on the line's baseline, a 20 pt line's too; **ending `w:distance` left of
  the text column's edge** -- auto 360 twips -- rounded to a whole device px (200 twips:
  258.33 px drawn to 258), the paragraph's indent not moving it; each digit advances by
  whole device px and is drawn on one, as a note number's (`1` at 202, `10` at 179, the
  number ending at 225).
* **In** the `LineNumber` character style over the document's defaults (Calibri 11 here,
  as the probe's defaults).

| Glyphs, each setting (Word drew 6,255) | matched | baseline | drawn not by Word | Word's not drawn |
| --- | --- | --- | --- | --- |
| before: no line numbers | 5,885 | 5,580 | 143 | 370 |
| the count starts at `w:start` | 5,801 | 5,337 | 450 | 454 |
| `continuous` restarts with the section | 6,251 | 6,251 | 3 | 4 |
| **the model** | **6,255** | **6,255** | **0** | **0** |

A number is a span of the line it stands beside, `data-docx-path="w:sectPr[k]/w:lnNumType"`
(`kind` `line-number`), which `paths.resolve_path` resolves. Not measured: line numbers in
a section of several text columns (not drawn, warned `line-numbers-not-drawn`), beside a
drawing text wraps around, and right-to-left sections. No committed document numbers its
lines.

### 4.19 Footnotes referenced in a table cell — measured, drawn

A footnote referenced in a cell took no room at the page's foot and was not drawn
(`footnotes-not-drawn`; docx-agent's E2 proposal). `make_cell_footnote_probe.py` /
`read_cell_footnote_probe.py` (`tests/test_cell_footnotes.py`), the notes and separators as
4.13 writes them: a short page with a body reference, a reference in a cell of a table's
first row and one in each cell of its second, and a body reference after; tables of
one-line rows whose third row references a three-line note, body lines before them swept
from 42 to 48 so the row, its note, both or neither fit; tables whose second row holds a
cell of six paragraphs, the fifth with a reference, swept from 44 to 48 lines so the row
splits at the foot. No `settings.xml` and mode 15. **4.13's rules extend to a cell's
lines** (`paginate._table_notes`, `paginate.cell_notes`):

* **numbered** in document order with the body's (body 1, the cells 2, 3, 4, body 5: as
  `notes.footnote_numbers` already numbered them) and **set at the foot** of the page
  their reference's line is on, under the separator, as a body reference's are;
* **a cell line holding a reference fits as a body line does**: its end -- with the cell's
  bottom margin and the edge under its row -- above the separator, the notes already on
  the page, its new notes but the last whole and the least of the last; when it goes in,
  the last note takes what then fits, and later lines fit above the notes as placed. **A
  line that does not fit ends the table's part of the page there**: a one-line row moves to
  the next page with its note (fills 44 to 48), and the row after a placed one goes when
  the note leaves it no room (fill 43); a cell's lines split before the line holding the
  reference (the split family), the row going on at the next page's top with the note
  under it.

The layout draws the row where the paginator split it (`table.cut_piece`): placed again
without the notes' room, the row would split lower.

| Glyphs (Word drew 22,363 in each setting) | none: matched, baselines | 15: matched, baselines |
| --- | --- | --- |
| before: a cell's notes take no room and are not drawn | 4,391, 3,201 | 4,391, 3,201 |
| **the model** | **22,361, 22,337** | **22,363, 22,363** |

Left, with no `settings.xml`: in the split case filled with 44 lines Word moves the line
above the one holding the reference to the next page too, where mode 15 keeps it (not
found: 24 glyphs a page lower). 19 object starts in each setting are 0.0004 px off, as
4.13's notes are. Not measured: a cell's notes in a section of several text columns (still
warned `footnotes-not-drawn`), a cell's note in a merge going across a page, and a header
row's note.

---

## Phase 5 — SVG output (M)

Deliberately late. **Producing output before the layout is right produces a renderer whose
errors are invisible** — which is the failure mode the landscape survey found everywhere.
By this point the numbers are already verified, and emission is a transcription.

One SVG per page, `data-docx-path` identity on every element (the sibling's `data-pptx-id`
lesson: an id from the file is *not* unique, and a consumer needs the structural path as
well), fonts named rather than embedded.

**Done: see *Phase 5 — measured* below.** One SVG per page in Word's device pixels, one
`x` per glyph, `data-docx-path` and `data-docx-id` on every element, fonts named; every
glyph of every committed document -- and all 496,104 of `filesamples/sample4` -- where
Word drew it, on Word's baseline, in Word's face and size. Tables, floating drawings and
multi-column sections stop the layout, visibly and with a warning, rather than be
guessed. *Columns are laid out since 4.14.* *Since* Tables — measured, *tables are laid out and drawn; floating
and nested tables, and autofit tables Word would resize, still stop it.*

### For callers: element paths, one call, a line's column (2026-10-03)

Asked for by docx-agent after its E0, which mirrored the path rules, laid documents out
twice and could not say which column a paragraph is in. Metadata and API only: every SVG
snapshot and fidelity baseline is byte-identical.

- **Element paths, exported.** `docx2svg.paths` holds the step counter the parser writes
  every path with (`Steps`, `row_level`: local names counted XPath's way, rows and cells
  through row- and cell-level content controls), and `resolve_path(root, path)`, which
  walks a path back to its element in a standard-library or lxml tree of the part:
  `w:footnote[@w:id=N]` by id, `wp:anchor[k]` and `wp:inline[k]` among a run's drawings as
  the parser meets them, a group's members among the group's children, unindexed steps
  (`w:body`, `wps:txbx`, `w:txbxContent`) as the first below, `label` as its paragraph,
  `w:sectPr[k]/w:cols` as the k-th section's. `path_part` gives the part (`Line.part`,
  the notes part, or the main part) and `locate` both. `tests/test_paths.py` resolves
  every line, span, float and stop of the committed documents and eight probe sets to an
  element of the named kind, every line to the `w:p` whose `w14:paraId` it carries. What
  names nothing: Word's default note separators (`w:footnotes/separator`), which the
  document does not hold. **Known, and left as they are drawn** (the paths are what the
  SVG has always carried): the runs a paragraph joined across a deleted mark took from
  the paragraphs after it keep their own paragraph-relative paths under the first
  paragraph's; a stop at a table row counts the rows drawn, not the deleted ones; a
  header's rules carry no part.
- **One call.** `convert_docx(source, options) -> Conversion(layout, svgs, page_numbers)`
  lays out once (`tests/test_api.py` counts the layouts) and returns the SVGs
  `convert_docx_to_svg` writes, byte for byte, and the layout of every page.
- **`Line.column`**: the 0-based text column in a section of several (set as each column
  is walked, and for a column's footnotes below mode 15 and the note area's lines and
  separator in mode 15), `None` in a section of one column and for story and text-box
  lines. `tests/test_columns.py` holds every column probe to it: on every page each
  column's lines start right of the column before's.

## Phase 5 — measured

**Done: `docx2svg` draws.** `convert_docx_to_svg(source, options) -> list[str]` returns one
SVG per page; `convert_docx_to_png` rasterises through the optional `png` extra (resvg-py);
`docx2svg report.docx -o out/ -f both --strict` is the command line; `ConvertOptions.warnings`
collects a stable code for everything not drawn faithfully. The shape is the sibling's
(`pptx2svg/__init__.py`, `cli.py`, `png.py`), written here, not imported.

**Emission is a transcription, and that is now an assertion.** Nothing in the renderer
decides where a line breaks, a page ends or a baseline sits. `docx2svg.layout` walks the
paginator's pages with the paginator's own stack (`paginate._scan`'s gaps, borders and
page-top rules), rounds each baseline in its line box as `tools/baselines.predict` does,
places every glyph from the breaker's own advances in Word's layout unit, and records the
rules, pictures and where it stops; `docx2svg.svg` formats it. Two instruments hold it
(5.3, 5.7), and on every committed document **every glyph is where Word drew it**.

### 5.1 The SVG

* **Word's device pixels.** `viewBox` is the page's extent in whole 1/300-inch pixels
  (Phase 0.5: 2480 × 3508 for A4), `width`/`height` in points. Baselines are integers;
  a pen position is exact to four decimals (the layout unit is 0.00102 px).
* **One `x` per glyph.** Each `<text>` is one run's glyphs in one format on one line, with
  an `x` for every character, so the rasteriser's kerning, shaping and ligatures cannot
  move a glyph (SVG 1.1 10.4: each absolutely positioned character starts a new text
  chunk). Word kerns only where `w:kern` asks, and the layout has added it.
* **Identity.** Every element drawn from the document carries `data-docx-path`, its
  structural path (`w:body/w:p[3]/w:hyperlink[1]/w:r[2]`, `w:body/w:tbl[1]`); a
  paragraph's group also `data-docx-id`, its `w14:paraId` -- which Word copies with the
  paragraph, so it is not unique; hence both. Lines carry `data-docx-line` and
  `data-docx-baseline`.
* **Fonts named, not embedded**: `font-family`, `font-weight`, `font-style`.
* **Deterministic**: no set iteration, no clock; numbers formatted by one function
  (`svg.number`: four decimals, half up).

### 5.2 What draws, and the rules drawing needed

Each rule below is Word's, measured on a probe of this project's own
(`make_render_probe.py` / `read_render.py`: eight faces at five sizes, every line rule,
borders and shading, alignment, tabs; Word's filled rectangles read by
`quartz_pdf.fills`) and scored in `tests/test_render.py`.

| What | Rule | Evidence | Score |
| --- | --- | --- | --- |
| Text | the resolved face, size, weight, slant and colour (`w:color`'s value; a theme colour's `w:val` is its resolved RGB) at the model's pen position and baseline | every committed document (5.3) | every glyph |
| List labels | the label's glyphs (the mark's properties and the level's `w:rPr`) from the breaker's pieces, and the tab after them to the hanging stop | the committed documents' bullets and numbers | every glyph |
| Centred lines | half the slack **truncated to the layout unit** | two centred quotes of `style-document.docx` drawn at 754006 and 757515 units, the middles 754006.5 and 757515.5 | exact |
| Right-aligned, justified (`both`) | the slack after the line's last non-space; justified: spread over the spaces after the last tab, not on the last line or one a break ends | the render probe, no `settings.xml` | every glyph |
| Right, centre, decimal tabs | the text after the stop ends at, centres on, or puts its separator at the stop, never left of the tab; **the separator is the system's** (a comma on this machine), not the document language's | `12.345` after a decimal stop in an `en-GB` document is drawn as after a right stop | every glyph |
| Single and `words` underline | `post.underlinePosition` below the drawn baseline, `underlineThickness` thick (≥ 1), scaled by the size Word draws the glyphs at (whole px) and rounded; ends at the pen positions rounded, trailing spaces left out; `words` per word | 40 / 40 and 40 / 40 rectangles, pixel for pixel | exact |
| Strike, double strike | `OS/2.yStrikeoutPosition` above the baseline, `yStrikeoutSize` thick; double: one line above and one below it | 120 / 120 rectangles | exact |
| Highlight, run shading | the line's **text rows** (`layout.text_band`): the text line of the box -- below an `atLeast` line's extra, above a multiple's extra, the whole pitch of an `exact` line -- rounding its edges on their own, or in the baseline's three parts when there is space below the text; across the pen extent rounded | 100 / 100 (the text extent about the rounded baseline, the first rule tried: 12 / 20) | exact |
| Paragraph borders | side borders `w:space` (whole px, truncated) and **6 px** outside the text edge; a top border `w:space` above the first line's text row, a bottom border `w:space` below the last's; top, bottom and between lines run between the side borders' outer edges, or 6 px past the text; widths `w:sz` truncated to whole px | the probe's borders page pixel for pixel; bottom borders over two probes 125 / 150 exact, the rest 1 px (`round(y_after) - w`: 74, `floor(y_after - w)`: 121) | top, side, between exact; bottom 83% |
| Paragraph shading | the box inside the borders, or the text rows widened by 6 px | 3 / 3 | exact |
| Inline pictures | the extent on the baseline, left at the pen x plus the effect extent; *since 5.14 scaled to the extent truncated to whole twips, and standing where its line puts it* | Word's image `cm`: 0.02 px from ours; *5.14: 961 / 962 boxes within 0.03 px* | 0.02 px |
| `w:position` | whole half points, the device rounding symmetric in magnitude | the script probe's 384 `position` groups | exact |
| Superscript, subscript | size and offset: 5.5 | | 5.5 |
| Thick, double, dotted, dashed underlines | *approximate*: warned (`underline-style-approximate`) | top of a thick band `ceil(position/2)`: 28 / 40; its height: best of 24 formulas 23 / 40 | ±1-2 px |
| `distribute` | slack over the spaces as justified; the **last** line spread evenly over its glyph gaps -- *approximate*, warned | Word gives the gaps after a space more than the others (51.02 and 64.53 px on one line, 5.43 and 15.34 on another), and in a narrow column spreads over characters as well: not modelled | non-last lines with small slack exact |

**Word draws with its bundle's copy of a face.** It lays a face out with the macOS system
copy (2.4), but the fonts its PDFs embed are its own bundle's where it has one: Times New
Roman 7.00, Arial 6.80, Verdana 5.02 (the system's 5.01), and the system's Courier New and
Georgia 5.00 where the bundle has none. Advances and vertical metrics agree; outlines and
hinting need not. `InstalledFonts.drawing_face` finds that copy, and the PNG path draws
with it. **And a rasteriser must be able to reach it**: resvg's font database files a face
under its typographic family (name ID 16) and never then under name ID 1, so
`font-family="Calibri Light"` drew nothing -- `style-document.docx`'s headings were blank,
raster SSIM 0.8898. ~~`fonts.rasteriser_files` hands over Word's copy in place where it
answers to the document's name, and otherwise a relabelled copy
(`ooxml_common.fonts.sfnt.relabel`: name table and style bits only) in a temporary
directory that lives as long as the rasteriser needs it: 0.9631.~~ **Since 5.11, no copy:**
every installed face is handed over in place, and the SVG names a face its document name
does not reach by a fallback list, `font-family="Calibri Light, Calibri"
font-weight="300"`, which finds the original file: 0.9631, every pixel the same.

**The faces at run time.** `docx2svg.fonts` reads the installed faces with `struct` --
`head`, `hhea`, `OS/2`, `post`, `hmtx`, `cmap` (the (3, 0) symbol subtable too), the
legacy `kern` table, `name`; collections; the document's embedded, obfuscated faces --
and finds a face exactly as `tools/face_metrics.py` does. It reproduces all 16,208
advances, kern pairs and metric integers the `fontTools` tools recorded
(`tests/test_fonts.py`). `src/` stays standard-library only; PNG goes through the
`png` extra, imported inside `png.py` only (`test_src_imports...` allows that file that
import and no other).

### 5.3 The glyph-position check — every glyph where Word drew it

`tools/glyphs.py`, recorded and held offline by `tests/test_render.py`
(`tests/fixtures/render-observations.json`: Word's text objects and filled rectangles, and
the faces' numbers, for the committed documents and the render probe). Two equalities:

1. **The SVG is the layout.** Every glyph's `x` and `y` in the emitted SVG is the layout's
   exact position, formatted (to 5e-5 px): nothing in the writer moves a glyph.
2. **The layout is Word.** Matched glyph for glyph, page by page, against Word's PDF: the
   pen x of the first glyph of each text object (Quartz starts one wherever Word places a
   glyph on its own: a run, a tab, a script, the mark) within what Quartz's printing leaves
   uncertain -- a single-precision value in points, printed to seven significant digits,
   0.00018 px at x = 989 px, a sixth of the layout unit, so an error of one unit fails;
   inside an object, every **advance** from one glyph to the next within half of Quartz's
   1/1000-em grid (and half of `Tc`'s printed precision); the drawn baseline (a script's or
   `w:position`'s offset included), the face (`/BaseFont`) and the size (`Tm`).

**Found on the way, and recorded because it is where Word's raster is not its layout:**
Quartz encodes each glyph's *advance*, not its position -- integer `/Widths` per 1000 em of
the size it draws at, a `Tc` for the object, an integer `TJ` after a glyph -- so inside a
long text object the positions Word's PDF actually draws drift from the pen: 0.86 px after
100 glyphs of Calibri 11 pt, the same per glyph every time it occurs (`e` -0.0026 px, `a`
+0.0199, a space +0.0205). Word's layout is exact (every object start agrees to the
unit); its PDF draws inner glyphs up to about a pixel off it. The SVG draws the layout.

| Document | Glyphs drawn, all matched | Word drew | Object starts exact | Advances exact | Baseline, face, size |
| --- | --- | --- | --- | --- | --- |
| `layout-sweep.docx` | 679 | 679 | 94 / 94 | 585 / 585 | all |
| `style-document.docx` | 1,366 | 1,366 | 48 / 48 | 1,318 / 1,318 | all |
| `samplelib/sample-long.docx` (36 pages, 18 pictures) | 27,458 | 27,458 | 455 / 455 | 27,003 / 27,003 | all |
| `samplelib/sample-resume.docx` | 733 | 733 | 42 / 42 | 691 / 691 | all |
| `samplelib/sample-simple.docx` | 759 | 844: 85 past the table | 37 / 37 | 722 / 722 | all |
| `samplelib/sample-blank.docx` | 0 (one blank page) | 0 | | | |
| `wordto/sample-1page.docx` | 424 | 424 | 9 / 9 | 415 / 415 | all |
| `wordto/sample-5pages.docx` | 1,403 | 2,727: past a table | 36 / 36 | 1,367 / 1,367 | all |
| `wordto/sample-10pages.docx` | 3,656 | 4,211: past a table | 78 / 78 | 3,578 / 3,578 | all |
| `wordto/sample-with-images.docx` | 903 | 903 | 16 / 16 | 887 / 887 | all |
| `wordto/sample-with-table.docx` | 238 | 1,141: past a table | 7 / 7 | 231 / 231 | all |
| **Committed** | **37,619** | | **822 / 822** | **36,797 / 36,797** | **all** |
| `filesamples/sample4.docx` (175 pages; not committed) | **496,104** | 496,104 | **9,256 / 9,256** | **486,848 / 486,848** | **all** |
| `filesamples/sample2.docx` | 454 | 454 | 10 / 10 | 440 / 444: Helvetica Neue (1000 upm), 0.1-0.3 px | all |
| `filesamples/sample3.docx` | 880 | 2,174: past a floating drawing | 38 / 38 | 842 / 842 | all |
| `filesamples/sample1.docx` (mode 12) | 1,871 | 9,619: past a table | 30 / 83 | 1,665 / 1,788 | baseline 1,047: mode 12's embedded Ubuntu (finding 7) moves its lines |
| render probe, no settings | 18,817 *(19,312 since 5.19)* | 19,312: tab leaders *(drawn since 5.19)* | 1,477 / 1,477 | 16,469 / 17,340 | all; the 871 are `distribute` |
| render probe, mode 15 | 18,741 of 18,817 | 19,312 | 1,440 / 1,505 | 15,904 / 17,236 | justified lines break differently in mode 15 (4.6) |
| ink probe (5.4) | 6,864 | 6,864 | 192 / 192 | 6,672 / 6,672 | all |

The local corpora are scored on the machine that holds them, their numbers kept there.

### 5.4 Ink or advances — settled: the ink, at the device-rounded size

Word lays a glyph out at the exact size and its PDF scales the glyph's outline by the size
rounded to whole device pixels (Phase 2). **The SVG places every glyph from the model's
advances, exact to 1/4096 pt, either way**; what `ConvertOptions.glyph_size` chooses is only
the outline's scale -- `device` (Word's ink: `font-size="46"` for 11 pt) or `exact` (45.8333).
The default is `device`, on this evidence:

* **Word's own geometry says so, everywhere.** Every text object of every export compared
  -- 37,619 committed glyphs, sample4's 496,104, the probes' -- is scaled by the size
  rounded to whole pixels (the glyph check's size column, which compares `Tm` with the
  device-rounded size, agrees on all of them). That is the drawing Word specifies; the
  device size reproduces it exactly, the exact size never does except where the two
  coincide.
* **A raster cannot decide it, and saying so is the measurement.** `make_ink_probe.py` sets
  one size a page at the sizes whose rounding moves most, both ways, with 12 pt (50 px, no
  rounding) as the control; `tools/fidelity.py --ink` scores ours both ways against
  pdfium's raster of Word's PDF at 300 dpi:

  | Size | Exact px | Word draws | SSIM device | SSIM exact | O's ink centroid height, Word / device / exact |
  | --- | --- | --- | --- | --- | --- |
  | 7.5 pt | 31.25 | 31 | 0.9685 | **0.9736** | 10.03 / **10.12** / 10.20 |
  | 8.5 pt | 35.42 | 35 | 0.9657 | **0.9669** | 11.30 / **11.40** / 11.53 |
  | 9.5 pt | 39.58 | 40 | **0.9634** | 0.9243 | 12.81 / 12.92 / **12.78** |
  | 10.5 pt | 43.75 | 44 | **0.9536** | 0.9330 | 14.04 / 14.16 / **14.09** |
  | 11 pt | 45.83 | 46 | **0.9070** | 0.8876 | 14.66 / 14.79 / **14.72** |
  | 12 pt (control) | 50 | 50 | 0.9228 | 0.9228 | 16.03 / 16.01 / 16.01 |
  | 14.5 pt | 60.42 | 60 | 0.8944 | **0.9251** | 19.13 / **19.11** / 19.25 |
  | 20.5 pt | 85.42 | 85 | 0.9201 | **0.9539** | 26.87 / **26.85** / 26.98 |

  **SSIM prefers whichever of the two is larger** -- the device size where Word rounds up,
  the exact size where it rounds down -- and **the ink's centroid (which stroke weight does
  not move) prefers whichever is smaller**, in all seven sizes both. Each metric has a
  direction of its own, independent of which size Word drew at. The control says why: at
  12 pt, where the two sizes are one, the two rasterisers ink the same outline 1.4%
  differently (resvg 104,644 units of ink, pdfium 103,197): pdfium rendering Word's PDF and
  resvg rendering ours differ in hinting and anti-aliasing by more than the 0.36-1.2% of a
  glyph the question is about. On the committed documents SSIM splits likewise (5.7).
* **So:** the SVG reproduces the ink as Word specifies it -- `device` -- and the advances
  exactly, which no raster metric here can separate from `exact` more finely than the two
  rasterisers' own difference. `exact` stays an option for a consumer who wants outlines to
  match the advances.

### 5.5 The script offset — measured

Finding 5 left the offset recorded, not settled. `make_script_offset_probe.py` settles most
of it with **faces of this project's own**: 30 TrueType faces written by the script with
`fontTools` (rectangle glyphs, public domain, no outline from any existing font), each moving
one field -- `ySuperscriptYOffset` 0-1400, `ySubscriptYOffset` 0-1000, the script size, the
ascent, the descent, a negative subscript offset -- embedded obfuscated (ECMA-376 17.8.1),
which Word lays out and draws; `w:sz` 8-96, and a script beside a 36 pt run and one larger
than its text. `read_script_offset_probe.py` / `tests/test_script_offset.py`.

* **Whole half points**: all 1,380 drawn offsets are `round(k × 25/12)` px.
* **Word takes a face's script fields only if the superscript offset is at least about a
  fifth of an em and the subscript offset is not negative** (`resolve.script_fields_taken`):
  at 0, 200 and 400 units of 2048 (≤ 0.195 em) and with a negative subscript offset it draws
  the 3/5 fallback size *and* fallback offsets, although the size field (0.65) is fine. The
  bound lies in (0.195, 0.246] (Cambria's 0.246 is taken); 1/5 is a placeholder. Drawn size
  690 / 690.
* **The move is the face's offset, rounded, but no more than the difference of the two
  sizes' ascents** -- `round(A × w:sz) − round(A × s)` for the drawn size `s` and the ascent
  with the line gap `A`; descents, for a subscript. That bound is what holds Calibri, Times
  New Roman, Arial and Courier New (offsets 0.42-0.48 em) to about a third of the size. A
  refused face moves by the bound alone, at the fallback size.
* **Not bounded by the line**: beside a 36 pt run, or at 24 pt among 11 pt text, a script
  moves exactly as it does alone at its own size.

| `resolve.script_raise_half_points` | Exact | Within ½ pt | The face's offset alone |
| --- | --- | --- | --- |
| probe superscripts (690) | 545 | 661 | 309 |
| probe subscripts (690) | 445 | 670 | 131 |
| real faces' superscripts (`make_script_probe.py`, 1,140) | 786 | 1,068 | 459 |
| real faces' subscripts (1,140) | 965 | 1,135 | 505 |

What rounds the bound is not settled: `round`, `floor`, `ceil` of each term and of their
difference, in half points, px, points and twips, and the line-box rounding of each size's
extent, were tried (best as adopted). Times New Roman and Arial are the least well fitted
(15 and 19 of 95 superscripts exact, the rest mostly one half point low). **Probe that would
decide it:** the offset sweep crossed with the line gap and the `OS/2` typo and win
metrics moved one at a time, and sizes stepped by one half point.

### 5.6 What is not drawn, and how that looks

**Where the paginator stops, the output stops, and says so.** Past a table, a floating
drawing, a multi-column section (the paginator does not paginate by column, so the
renderer stops at its first paragraph rather than draw one column as if it were the page;
*since 4.14 only a footnote, floating drawing, floating table, drop cap or endnote in one;
since 4.15 a footnote only in one a continuous break joins to another; since 4.16 no drop
cap, and an endnote only in mode 15; since 4.17 a floating drawing only where it is not
positioned against its column*),
a field, a frame, a picture bullet or a paragraph it cannot measure, nothing is laid out by
guessing. The page the obstacle is on is drawn up to it, then:

* a **stop band**: a dashed, pale-grey rectangle from where the obstacle starts to the foot
  of the text area, across the column, labelled `docx2svg: not drawn: table. The layout
  stops here; N more item(s) of the document are not laid out.`, carrying
  `data-docx-unsupported="table"`, the obstacle's `data-docx-path` and
  `data-docx-remaining`;
* **no page after it** -- the page count past an unknown height is unknown;
* a warning `layout-stopped:<reason>` naming the obstacle, its page and what is not drawn;
  for an unmeasurable paragraph, what could not be measured (`no advance for 'x' in Foo`: a
  face this machine lacks).

Where the extent *is* known it is drawn as a placeholder of that extent: an inline drawing
that is not a picture (a chart, a shape; *shapes drawn since F.12, charts since F.20 and
SmartArt since F.21*), a picture in a format an SVG cannot carry (EMF,
WMF), and -- outlined inside the stop band -- a floating drawing whose anchor states its
offset against the page, margin or column and the page, margin or its paragraph
(`filesamples/sample3`'s picture: within 0.07 px of where Word drew it). A table's
width is known and its height is not, so it gets the band alone. The warning codes:

| Code | Meaning |
| --- | --- |
| `layout-stopped:table`, `:drawing`, `:columns`, `:field`, `:frame`, `:drawing (picture bullet)`, `:unmeasurable`, `:no face metrics` | the layout stops there (above) |
| ~~`table-layout-unsupported`~~, ~~`content-control-unsupported`~~, `body-element-unsupported:<name>` | the parser met them (as before); *tables are laid out since* Tables — measured*, and one that is not is `layout-stopped:table`; a block-level content control's content is drawn since* Revisions — measured |
| ~~`headers-footers-not-drawn`~~, `footnotes-not-drawn` | they take their room (Phase 4) but are not drawn; *headers and footers are drawn since* Headers, footers and fields — measured *(H.4 has their warnings); footnotes since 4.13, but one referenced in a table cell* |
| `drawing-not-drawn`, `picture-not-drawn` | a placeholder of the extent |
| `chart-unreadable`, `chart-unsupported-type`, `chart-3d-flattened`, `diagram-unreadable`, `diagram-no-cached-drawing`, `diagram-not-drawn` | a chart or a SmartArt diagram not drawn (a placeholder, with `drawing-not-drawn`), or drawn flat (F.20, F.21) |
| `chart-title-not-drawn`, `chart-axis-title-not-drawn` | a title in Word's own words (its "Chart Title" / "Axis Title", in its interface's language), the band kept; an axis title turned or placed otherwise, or in a chart of a kind not measured (F.23) |
| `underline-style-approximate`, `alignment-distribute-approximate`, ~~`justification-unmeasured` (mode 15)~~ *(gone since Phase 3.8)*, `tab-leader-not-drawn` | drawn, not to the pixel (5.2); *leaders drawn since 5.19, warned only where a face's advance is not known* |

A document that draws everything warns of nothing (`style-document.docx`); `--strict` makes
any warning exit status 2.

### 5.7 Raster fidelity

`tools/fidelity.py`, the sibling's method adapted, not imported: SSIM on grey and
colour-histogram correlation (64 bins a channel) over **foreground pixels** (grey < 245 in
either image; under 1.5% foreground a page counts 1.0), at 300 dpi -- a device pixel a
pixel -- against pdfium's raster of Word's PDF. **Its caveat, kept:** SSIM is a mean over the
foreground mask, so removing wrong ink shrinks the denominator as well as the error and a
correction can lower it while the picture improves; every row also carries the
unnormalised structural loss (the sum of `1 − SSIM` over the mask). **Both sides draw with
Word's faces**: ours with the files Word draws with, read in place (5.2, 5.11), the document's
embedded faces de-obfuscated into a temporary directory, the host's other fonts off; a
document drawing a face this machine lacks is skipped with the reason.

| Document | Pages scored (made, Word's) | SSIM device | SSIM exact | Histogram | Loss device / exact |
| --- | --- | --- | --- | --- | --- |
| `layout-sweep.docx` | 3 (3, 3) | **0.9639** | 0.9615 | 0.9914 | 5,519 / 5,887 |
| `style-document.docx` | 2 (2, 2) | **0.9631** | 0.9564 | 0.9818 | 8,817 / 10,444 |
| `samplelib/sample-long.docx` | 36 (36, 36) | 0.9482 | **0.9518** | 0.9814 | 21,528 / 20,052 |
| `samplelib/sample-resume.docx` | 1 (1, 1) | 0.9192 | **0.9213** | 0.9675 | 18,181 / 17,724 |
| `wordto/sample-10pages.docx` | 2 (3, 3) | 0.8596 | **0.8649** | 0.9678 | 65,394 / 62,826 |
| `wordto/sample-with-images.docx` | 1 (1, 1) | 0.9604 | **0.9620** | 0.9596 | 42,515 / 40,713 |
| `wordto/sample-1page.docx` | 1 (1, 1) | 1.0 (under 1.5% foreground: not scored) | 1.0 | 1.0 | |
| `samplelib/sample-simple.docx`, `wordto/sample-5pages.docx`, `wordto/sample-with-table.docx` | 0: every page made holds a stop band | *0.8091*, *0.7190*, *0.6844* over the page made | *0.8094*, *0.7207*, *0.6844* | | |
| render probe, no settings / mode 15 | 31 / 31 | 0.9242 / 0.8785 | **0.9262** / **0.8804** | 0.9970 / 0.9977 | |

Every page scored holds every glyph at Word's position (5.3), so what separates these from
1.0 is not layout. Where a page is denser the score is lower (`sample-10pages`, 5.7% of it
ink, against `sample-long`'s pages of prose and pictures): the residual is per glyph edge.
The two glyph sizes split four ways to two, as the ink probe predicts (5.4).

The uncommitted `filesamples` (device / exact, pages drawn in full): `sample4` 0.8748 /
0.8785 over its first 25 pages (text and pictures, every glyph at Word's position);
`sample2` 0.9274 / 0.9252; `sample1` (mode 12) 0.7022 / 0.7008 over the two pages before
its first table, where finding 7 moves its lines and its header is Word's alone;
`sample3` stops at its floating drawing on page 1.

The ceiling is the two rasterisers': on a page whose every glyph is at Word's position, in
Word's face and size (`sample-resume`), the difference image is a hairline around every
glyph's edge, pdfium's hinting and anti-aliasing against resvg's. Pages with a stop band
are reported separately (`SSIM over every page made`): the band is ours, Word's table is
not.

**Outliers are failures to explain** (5.12): a page or document 0.05 below the typical
one is listed, and fails the test unless explained -- `sample-with-table`'s 0.787 was its
shaded rows' text painted over.

**A page too many or too few is a regression, and is scored as one.** The harness used to
score the pages both sides made and print that count as the pages "made", so a local
template whose layout made a third, empty page where Word made two scored exactly as
before and read "2 of 2 made" (4.11 found it by eye). Now every page either side made is
scored: a page only one side made counts as wholly wrong (SSIM and histogram 0, its loss
every foreground pixel of the page that exists, an empty page included) in every mean;
each document's line reads `pages ours N, Word M`; and a count that is not Word's is
printed as a `PAGE COUNT MISMATCH`, makes the run exit 1, refuses `--record`, and fails
`test_fidelity_does_not_regress`. Such a page is kept out of the medians and the outliers,
which are about pages both sides made. Every committed document makes Word's pages, so
`tests/fidelity-baselines.json` was unchanged by this, byte for byte (`pages_made` has
always been our count wherever the two agreed). Re-run on the commit before 4.11, the
local templates that made a page too many are reported, and score a third lower.

`tests/test_fidelity.py` holds the scores (`tests/fidelity-baselines.json`) to a drop of at
most 0.02; it needs Word's exports and faces, takes minutes, and is marked `slow` (run with
`--run-slow`), so the default suite stays the five minutes it was.

### 5.8 VRT snapshots

`tests/test_vrt.py`, the sibling's arrangement: every page of the eleven committed
documents rendered with the faces' **recorded** numbers (no font file read) and compared
byte for byte with `tests/vrt/` (72 pages, 732 kB); two subprocesses under two
`PYTHONHASHSEED`s render the same bytes; the render emits no carriage return;
`.gitattributes` pins the snapshots to LF; a picture is compared by the SHA-256 of its
bytes instead of carried twice. `--update-snapshots` rewrites them and skips.

### 5.9 Proposals for `ooxml-common`, recorded rather than made

Changing the sibling repositories was not part of this phase. Two pieces now exist twice
and belong there:

* **`svg_to_png`** -- `docx2svg/png.py` and `pptx2svg/png.py` are the same wrapper (resvg
  first, cairosvg as fallback, any import failure meaning unavailable), less `pptx2svg`'s
  bundle logic. Proposed: `ooxml_common.png.svg_to_png(svg, *, width, height, scale,
  background, backend, font_dirs, font_files, skip_system_fonts, sans_serif_family)`,
  each consumer keeping its font policy.
* **Positioned text emission** -- `pptx2svg` emits one `x` per run fragment and lets the
  rasteriser advance; `docx2svg` one `x` per glyph from exact layout units. A shared
  `ooxml_common.svg.text(chars, xs, y, family, size, *, weight, style, fill, path)` with one
  number formatter (`pptx2svg`'s three decimals and `docx2svg`'s four, half up, reconciled)
  would let `pptx2svg` adopt per-glyph positions where it measures them, and keep both
  outputs byte-deterministic by one rule. And `fonts.rasteriser_files`' relabelling of a face
  the rasteriser cannot reach by name is `pptx2svg`'s `addressable_font_files` done with
  `ooxml_common.fonts.sfnt.relabel`; the two should be one function there.
  *Since 5.11 `docx2svg` relabels nothing: it names such a face by a fallback list and hands
  the original file over. `ooxml_common.fonts.sfnt.relabel` (and `write_sfnt`) are no
  longer used by `docx2svg`; `ooxml_common.fonts.embedded` still relabels `pptx2svg`'s
  embedded faces with it, so it stays. Recorded for a later clean-up of that package: if
  `pptx2svg` adopts the fallback list for installed faces, what is left of `relabel` is
  the embedded-font case alone.*

### 5.10 Not settled, or not measured

* **The script offset's rounding** (5.5), and Times New Roman and Arial within it. *5.17: not its rounding -- the bound's form.*
* **Thick, double, dotted and dashed underlines** (5.2): the band's height; the patterns. *5.15: dotted and dashed drawn to the pixel; thick and double measured further, not settled.*
* **Distributed alignment**: how Word divides the slack between character gaps and
  spaces. **Justification in mode 15**: Word fits more words on a justified line -- in the
  render probe up to 24.8% of a line's space width over the column (11 lines) -- by
  compressing spaces by an amount not measured; it moves lines, and so pages (Phase 3's
  justification measurement). *Measured and adopted: Phase 3.8.*
* **Bottom borders**: 25 of 150 one pixel off; the two rules tried each fail different cases. *5.15: probed over 770; not settled.*
* **Highlight and shading across two lines of different text rows**, and run borders
  (`w:bdr`): drawn? -- run borders take their room (finding 5) and are not drawn.
* **~~Headers, footers, footnotes~~** are laid out for their room and not drawn *(headers and footers drawn: H; footnotes: 4.13)*; ~~tab leaders are
  not drawn~~ *(5.19)*; `w:smallCaps`, `w:outline`, `w:shadow`, `w:emboss` and `w:imprint` are drawn
  plain.
* **Blank pages** Word inserts (an `oddPage` section on the wrong parity, a footnote
  continuation page) are emitted empty; their headers would be drawn there. *Measured
  (H.2): Word draws no header or footer on them either.*
* **Word's PDF draws inner glyphs up to about a pixel off its own layout** (5.3); the SVG
  draws the layout, so a raster comparison carries that too.


### 5.11 The rasteriser reaches every face in place -- no font file copied

5.2's relabelled copy wrote a modified Microsoft font file into a temporary directory for
every face a rasteriser could not reach by the document's name. It was not needed.
resvg's font database files a face under its typographic family (name ID 16) *and* its
`OS/2` weight class, width class and `fsSelection` style, so the original file answers to
that family at that weight: `calibril.ttf` and `calibri.ttf` given as `font_files`, system
fonts off, `font-family="Calibri" font-weight="300"` renders pixel-identical to the Light
file alone, and 400 renders Regular.

**The rule** (`fonts.DrawingName`, `InstalledFonts.drawing_name`, `svg._font`): where the
document's name and bold/italic already reach the face, the SVG is unchanged. Otherwise
`font-family="<document name>, <typographic family>"` -- a browser matching full names
still finds the first -- with the face's own weight, `font-stretch` from its width class
and `font-style` from its `fsSelection` bits. `fonts.rasteriser_files` returns every
installed file in place and writes nothing; only a face the document *embeds* is written
(it exists nowhere else, and resvg reads fonts from files alone): de-obfuscated, byte for
byte otherwise, while the rasteriser needs it, and never when its name table names
Microsoft (then it counts as missing).

**What had to be measured** (a scratch harness rendering each face both ways, alone and
with every other face of its family loaded):

| Face | Fallback list against the relabelled copy |
| --- | --- |
| Calibri Light (and its italic), Arial Black, Aptos Light / SemiBold / Black / ExtraBold (and italics), Microsoft YaHei (UI) Light (weight 290, written 300), DengXian Light, Malgun Gothic Semilight, Open Sans Light, Seravek Light, Charter and Iowan Old Style Black, Avenir Light, Avenir Next Condensed Medium (`font-stretch="condensed"`) | pixel-identical, alone and with every face of the family loaded |
| `font-weight="350"` (Yu Gothic UI Semilight), `"275"`, `"290"` | **resvg reads only multiples of 100**; any other weight is `normal`. So the SVG writes the multiple of 100 nearest the face's weight class that CSS font matching (CSS Fonts 3 5.2, as resvg applies it: stretch, then style, then weight) resolves to this face among the installed faces of its family (`fonts.css_match`) |
| Yu Gothic UI Semilight (350) with Yu Gothic UI Light (300) and Regular (400) loaded | **no weight reaches it**: 300 finds Light, 400 Regular |
| Avenir Book and Avenir Roman (both 400, normal, one family "Avenir" -- so `font-family="Avenir"` was already ambiguous before this change); Avenir Next Condensed's Demi Bold, Heavy and Ultra Light and their italics (the italics' `fsSelection` carries no italic bit, so each pair is one query); Hoefler Text Ornaments (filed as "Hoefler Text", 400) | two faces answer one query; the rasteriser draws the first it loaded (Avenir bold, i.e. Roman, drew Book) |

The last two rows are the only faces on this machine a name cannot reach, and no Word
document here uses them. A relabelled copy would reach them, and was **not kept**: no font
file is written. Instead `rasteriser_files` checks, among exactly the files it hands over,
that each drawn face is the one the SVG's name resolves to, and `convert_docx_to_png`
warns `face-not-addressable` where it is not (`tools/fidelity.py` skips such a document).

**No output changed.** Over 28 documents -- the eleven committed, the `filesamples` and
their compatibility variants, the local corpora, the render and ink probes -- 315 SVG
pages are byte-identical except the two of `style-document.docx`, where only the
`<text>` elements of Calibri Light changed (`font-family="Calibri Light, Calibri"
font-weight="300"`); and all 945 PNGs -- every page through `convert_docx_to_png` and
through `tools/fidelity.py` at both glyph sizes -- are pixel-identical before and after
(SHA-256 of the pixels), so the raster fidelity scores are unchanged (5.7;
`tests/test_fidelity.py --run-slow` passes on the recorded baselines). The VRT snapshots
differ in those two pages' heading lines only.

**Office's cloud-font cache, a font source too** (docx-agent's proposal; it passed the
folders itself as `font_dirs`). Office for Mac downloads faces on demand into
`~/Library/Group Containers/UBF8T346G9.Office/FontCache/4/CloudFonts/`, one folder per
family -- Aptos Display, the heading face of every new Word 365 document, and on this
machine Segoe UI, Lato, Raleway, Ubuntu, Ubuntu Mono, Anton and Merriweather Sans Light.
`fonts.cloud_font_dirs` lists the family folders and `FONT_DIRS` searches them **after**
every installed folder, Word's bundle included, so a face found there is one nothing
installed answers to; they are read in place, as the bundle is, and never written, copied
or committed (the observations hold numbers only). `make_cloud_font_probe.py` /
`read_cloud_font_probe.py` (`tests/test_cloud_fonts.py`): a page per family -- a 20 pt
line, 16 pt lines bold, italic and bold italic, and paragraphs of 11, 20 and 28 pt that
wrap -- with no `settings.xml` and in mode 15.

| Glyphs (Word drew 4,336 in each setting) | none | 15 |
| --- | --- | --- |
| before: the layout stops at the first paragraph (`layout-stopped:unmeasurable`) | 0 drawn | 0 drawn |
| **the cache searched** | **4,336 matched**; x 4,175, baseline 4,025, face 4,184 | **4,336 matched**; x 4,280, baseline 4,075, face 4,235 |

**Aptos Display is exact** in both settings: every glyph of its page at Word's pen
position, baseline, face and size. Segoe UI, Lato, Ubuntu and Ubuntu Mono are exact in mode
15. What is left is each face's own, recorded and not fitted: Raleway's every baseline
1 px above Word's, Anton's and Merriweather Sans Light's baselines (their vertical metrics
are not among the faces the vertical model was measured on), and with no `settings.xml`
the synthesised bold italic of Segoe UI and Lato, whose caches hold no such face (Word's
advances are wider). The committed documents use none of these faces: the fidelity scores
are unchanged, every document's line identical.

### 5.12 Paint order, and the glyph that is where Word drew it but not seen

**The bug.** Every shaded (banded) table row drew no text: `sample-with-table`'s
"Word Document / .docx / Yes (OOXML) / No hard limit" was an empty shaded row. The text
was in the SVG at Word's positions; the cell shading (`cell-shading`) was emitted *after*
it, with the table borders, and SVG paints in document order. Paragraph shading, run
shading and highlight were already drawn first; the table's rules were not.

**Word's order, read off its PDF** (the content stream of `sample-with-table` and
`sample-simple`): per table row, the row's border lines, then each cell's shading
followed by that cell's text; a paragraph's borders and a run's underline or strike after
its text. So `svg.render_page` now paints table borders, then cell, paragraph and run
shading and highlight (outermost first), then the text, then paragraph borders,
underlines and strikes, pictures and the stop band. The model's borders and shading
share no pixel (checked over every document below), so all borders before all shading is
Word's picture.

**The audit, every filled element against every glyph box** (the committed documents,
the render and table probes, `filesamples` and its variants): painted over text before the
fix were **cell shading** (333 glyphs of `sample-with-table`, 28 of `sample-simple`, 74 of
`sample-5pages`, 104 of `sample-10pages`, 56 of the table border probe) and **table
borders**, which overlap glyphs in the row probe (39) and the border probe (2), where
Word draws the text over them. Paragraph shading, run shading and highlight overlap
thousands of glyphs and were already under them; underlines and strikes are over the
text in Word too. List labels are text. No inline picture, placeholder or stop band
overlaps a glyph anywhere; the stop band starts below the last line laid out. Only the
order changed: every VRT snapshot that moved (5.8) holds the same lines, the table's
rectangles moved ahead of the text, everything else in its order.

**Why nothing caught it.** The glyph-position check (5.3) holds every glyph to Word's
position and cannot see what is painted over it. The VRT snapshots were recorded from the
output and asserted the bug. And the raster said so -- `sample-with-table` 0.787 against
its neighbours' 0.92-0.96, loss 400,629 -- and was recorded as a gain.

**The visibility check** (`tools/visibility.py`, `tests/test_visibility.py`): each page
rasterised as emitted, without the model's text, and as the text alone in black; a glyph
is visible when its box differs between the first two by a clear ink difference. A glyph
under an opaque fill, or in the colour it sits on, fails; one no face on the host can draw
counts apart. Offline and in the default suite, from the recorded face numbers as the VRT
(no Word, no font file), over the committed documents, the render probe (run shading,
highlight, paragraph shading and borders) and the table border probe (cell shading);
skipped without resvg. **Before the fix: 5 of 13 documents fail** -- `sample-with-table`
333 of 1,141 glyphs hidden, `sample-10pages` 104, `sample-5pages` 74, the border probe 56,
`sample-simple` 28; **after: none**.

**A low score is now a failure to explain** (5.7): `tools/fidelity.py` lists every page
more than 0.05 below the median page of its document or of the corpus, with its loss;
`tests/test_fidelity.py` fails on a document mean 0.05 below the median document's -- in
the default suite, on the recorded baselines -- and on such a page with `--run-slow`,
unless an explanation there covers that number (a score 0.02 under it is an outlier
again).

| Document (device / exact) | Before | After |
| --- | --- | --- |
| `wordto/sample-with-table.docx` | 0.7868 / 0.7876, loss 400,629 | **0.8875 / 0.8900**, loss 211,480 |
| `wordto/sample-10pages.docx` | 0.8470 / 0.8511 | 0.8766 / 0.8808 (page 3, the table: 0.8217 → 0.9106) |
| `wordto/sample-5pages.docx` | 0.8369 / 0.8415 | 0.8609 / 0.8660 (page 1: 0.7967 → 0.8446) |
| `samplelib/sample-simple.docx` | 0.9268 / 0.9275 | 0.9409 / 0.9420 (page 1: 0.8575 → 0.8857) |
| the other six | | unchanged |

What is left below the others is explained, not recorded blind: `sample-with-table`
scores 0.9443 away from its border lines, which carry 64% of its loss -- Word's
rectangles are the model's to a few hundred px, but pdfium draws a 4 px line whose edge
is off the pixel grid 5 px wide, and the left border a pixel left (Tables stage 5), and
SSIM over thin foreground punishes a one-pixel line along its length. Pages 1 of
`sample-long`, `sample-5pages` and `sample-10pages` are dense Cambria prose (5.3's drift
inside Word's text objects, per-glyph edge residual), unchanged by this fix.

### 5.13 The instrument: one rasteriser for both sides

5.7 rasterised Word's PDF with pdfium and our SVG with resvg, so part of every score was
pdfium against resvg. **That part was most of it.** On `sample-with-table`'s page, pdfium
against MuPDF on the *same PDF* scores SSIM 0.890 -- what our render scored against Word --
while MuPDF against resvg scores 0.996. pdfium is the odd engine out: it grid-fits glyph
outlines (a stem or a bar moves by up to a pixel) and widens an axis-aligned fill to whole
pixels (5.12's 4 px line drawn 5 px wide).

**The candidates, measured** on Word's own pages, the same outlines on both sides (Word's
page converted as below, so any difference is the engines'):

| Page | pdfium / resvg, 1x | 4x supersampled | 8x | MuPDF / resvg | one engine (chosen) |
| --- | --- | --- | --- | --- | --- |
| `sample-with-table` 1 | 0.8910 | 0.9923 | 0.9981 | 0.9957 | 1 by construction |
| `sample-10pages` 1 (dense prose) | 0.8554 | 0.9933 | 0.9984 | 0.9839 | 1 |
| `layout-sweep` 1 | 0.9405 | 0.9952 | 0.9988 | 0.9797 | 1 |
| ink probe, 11 pt | 0.9232 | 0.9945 | | 0.9585 | 1 |
| `sample-with-images` 1 (a picture) | 0.9605 | 0.9980 | 0.9993 | 0.9935 | 1 |

* **(b) supersampling both sides** shrinks the engines' difference but leaves 0.5-0.8% of
  SSIM at 4x -- the size of every real difference below -- and 0.1-0.2% at 8x, at 64 times
  the pixels (a 300 dpi page drawn at 2,400 dpi: 3-4 minutes a page for pdfium and resvg
  here). It would still be scoring engines.
* **(c) MuPDF instead of pdfium for Word's side** needs no conversion, but is still two
  engines (0.96-0.996).
* **(a) Word's page converted to SVG, both sides through resvg** leaves no engine
  difference to score. **Chosen.** `tools/pdf_svg.py`; `tools/fidelity.py --truth svg`
  is the default, `--truth pdfium` the old instrument, kept.

**The converter is PyMuPDF** (`page.get_svg_image(text_as_path=True)`), development only,
behind the `fidelity` extra (AGPL-3.0: fine for a local tool not distributed with the
library; nothing under `src/` imports it; CI does not install it and the tests that need it
skip). poppler's `pdftocairo -svg` was not tried: PyMuPDF passed, and is a `pip` install
where poppler is a Homebrew tree. Two faults of its own were found and are corrected in
`pdf_svg.py`, which is why neither route is trusted blind:

* **It writes glyph outlines hinted.** FreeType loads each glyph with the font's
  instructions at one pixel a unit: Word's Calibri and Cambria subsets move by at most 1-3
  units, but PowerPoint's Aptos by 30-35 (its x-height pulled up 1.5% of the em, 2.7 px at
  44 pt on `table-test`). `unhinted_outlines` redraws every glyph from the embedded
  program with fontTools, choosing the program by MuPDF's text trace and, among subsets of
  one face, by nearest outline.
* **resvg rounds a root's size in points before zooming.** PyMuPDF writes A4 as 595.2 x
  841.92 pt; resvg drew it 2,479 px wide and 1e-4 too tall, 0.3 px at the foot of the page,
  which the first recording read as every line of ours sitting 0.1 px low on the two A4
  documents. `device_grid` sizes the root in device pixels and widens the viewBox to match.

**The converter validated** (`python tools/pdf_svg.py --validate`; the committed
documents' pages in `tests/test_pdf_svg.py`, slow): every page of the ten committed
documents' exports (52) and of the probes -- the ink probe (8), `render-15` (31: eight
faces, highlights, shading, paragraph borders), `border-probe` (20), `picture-15` (200),
`table-border-15`, `table-content-15`, `table-merge-15` (20), `script-15` (384),
`style-fonts` (Hebrew, Arabic, East Asian faces), Verdana and Aptos (42), kerning (7):
**765 pages, all pass** (`picture-15` under the criteria below but one: its run with the
final bitmap rule was stopped for time; before it, its only failures were MuPDF's SVG reader
filtering its 1 x 1 px swatches, with resvg's and pdfium's rasters exact). The slow test
holds a sample of six pages (prose, the ruled and shaded table, a picture, A4 in several
faces, the resume): 3 minutes. Per page:

* **same engine, two routes**: MuPDF's raster of the PDF against MuPDF's raster of the SVG;
* **no pixel beyond anti-aliasing** (a pixel outside the other raster's 3x3 range by more
  than 48 levels): at most 18 isolated pixels a page, on the thin stems of small Times New
  Roman and Arial, where MuPDF's glyph cache puts a PDF glyph on a sub-pixel grid and the
  SVG's path is exact; a moved glyph or a missing stroke is thousands (the limit is 20);
* **no colour changed**: where both rasters are flat, no channel more than 1 level apart;
* **every glyph the embedded program's**: every font embedded, every SVG glyph addressed
  by the *subset's* glyph id, named by MuPDF's text trace, redrawn from that program; the
  SVG has no `<text>`, and resvg rasterises it with no font available at all. A font
  lookup would carry the installed face's glyph ids and could not draw.

MuPDF's own SVG reader is not a faithful reader of everything the conversion writes, and
where resvg and pdfium agree with the PDF raster it is skipped rather than believed: it
ignores `<mask>` (an image's soft mask, drawn black) and filters a stretched bitmap its own
way (`picture-15`'s 1 x 1 px swatches stretched over 118 pt). The converted pages hold
Microsoft's glyph outlines: they are cached only in `ORACLE_DIR/svg/` beside Word's PDFs,
`pdf_svg` refuses to write one inside the repository, and none is committed.

**Recalibrated.** Every score moved; this is a new instrument, not a change to the renderer
(`tests/fidelity-baselines.json` holds both truths; the pdfium numbers re-recorded equal
the old ones exactly):

| Document | pdfium: SSIM device / exact | loss | **svg: SSIM device / exact** | loss | histogram |
| --- | --- | --- | --- | --- | --- |
| `layout-sweep.docx` | 0.9639 / 0.9615 | 5,519 | **0.9881 / 0.9849** | 1,652 | 0.9914 → 1.0 |
| `style-document.docx` | 0.9631 / 0.9564 | 8,817 | **0.9878 / 0.9860** | 2,647 | 0.9818 → 1.0 |
| `samplelib/sample-long.docx` | 0.9482 / 0.9518 | 21,528 | **0.9914 / 0.9908** | 5,516 | 0.9814 → 1.0 |
| `samplelib/sample-resume.docx` | 0.9192 / 0.9213 | 18,181 | **0.9914 / 0.9862** | 1,762 | 0.9675 → 1.0 |
| `samplelib/sample-simple.docx` | 0.9409 / 0.9420 | 31,206 | **0.9971 / 0.9957** | 1,473 | 0.9908 → 1.0 |
| `wordto/sample-1page.docx` | 1.0 / 1.0 (sparse) | 0 | 1.0 / 1.0 | 0 | |
| `wordto/sample-5pages.docx` | 0.8609 / 0.8660 | 73,860 | **0.9850 / 0.9822** | 7,414 | 0.9663 → 1.0 |
| `wordto/sample-10pages.docx` | 0.8766 / 0.8808 | 67,838 | **0.9867 / 0.9832** | 6,558 | 0.9769 → 1.0 |
| `wordto/sample-with-images.docx` | 0.9604 / 0.9620 | 42,514 | **0.9983 / 0.9973** | 1,787 | 0.9596 → 0.9999 |
| `wordto/sample-with-table.docx` | 0.8875 / 0.8900 | 211,480 | **0.9919 / 0.9904** | 14,711 | 0.998 → 1.0 |
| median page | 0.9000 / 0.9072 | | **0.9896 / 0.9856** | | |

The loss falls 3-19 times; the pages that spread 0.84-0.96 now spread 0.96-1.0. The
outlier drop (5.12) is per truth: 0.05 stays pdfium's, and the `svg` truth's is 0.02, the
old drop scaled to the new spread -- 0.05 below a typical 0.99 would let a page lose five
times the whole residual unremarked. `MAX_SSIM_DROP` stays 0.02.

**The explained outliers, revisited** (`tests/test_fidelity.py`). Under pdfium all seven
stand, as explained; they are kept as `EXPLAINED_PDFIUM`. Under the new instrument:

| Outlier (pdfium) | Now | What it was |
| --- | --- | --- |
| `sample-with-table`, the document (0.8875) | 0.9919, not an outlier | Mostly pdfium: loss on the ruled lines 131,531 → 12,509. The rest is real: see the left border below. |
| `sample-5pages`, the document (0.8609) and page 1 (0.8446) | 0.9850; page 1 0.9821 | pdfium's hinting on dense prose and its lines; page 1's residual is the Quartz drift and the left border |
| `sample-10pages`, the document (0.8766) and page 1 (0.8467) | 0.9867; page 1 0.9843 | pdfium on dense prose; the drift remains |
| `sample-long` page 1 (0.8571) | 0.9857 | pdfium on dense prose; the drift remains |
| `sample-simple` page 1 (0.8857) | 0.9946 | pdfium on the table's lines; the left border remains |

**None of them is an outlier any more**, and "the glyph-edge residual of two rasterisers"
is gone from every explanation: it was the instrument. One page is an outlier at the new
drop, explained: `layout-sweep` page 1 (0.9642 / 0.9548 against a median page of 0.9896),
the drift below on a sparse page.

**What the sharper instrument shows** -- real differences, hidden in pdfium's noise until
now, recorded here and not fixed:

1. **A table's left outer border is one pixel right of Word's**, along its whole height,
   on every table page: `sample-with-table` 1, `sample-5pages` 1, `sample-10pages` 3,
   `sample-simple` 1 (Word's 4 px line covers x = 350-353, the model's 351-354). Tables
   stage 5 recorded it from the filled rectangles; 5.12 put it down partly to pdfium. It is
   not pdfium's: it carries most of what is left on the lines (85% of `sample-with-table`'s
   remaining loss, 12,509 of 14,711). *Closed: 5.14.*
2. **Word leaves a white pixel column between a border and a shaded cell** where the model
   paints the shading up to the line: beside the left border in the shaded rows, and at the
   inner lines x = 1070-1074 and 1790-1794 of `sample-with-table` over rows 1339-1760
   (Tables stage 5's "one-pixel column unshaded", now seen beside drawn borders too).
   *Closed: 5.14.*
3. **Word's PDF draws the glyphs of a long text object ahead of its own pen** -- 5.3's
   Quartz drift -- and it is now the largest residual on every text page: per word, the
   model's ink against Word's differs by 0.15-0.5 px (standard deviation) and up to 1.57
   px, growing along a line (`layout-sweep` 1, y = 1208: 0.05, 0.25, 0.43, 0.70, 1.52 px);
   vertically 0.005-0.03 px, and the ink of each word agrees to 0.1-0.4%. This is Word's
   encoding, not Word's layout (every text object's start is exact), so it is a ceiling
   on raster fidelity rather than a defect to fix; it can be separated from real layout
   error only by the glyph check (5.3), which already does.
4. **A picture sits about half a pixel low**: `sample-long`'s full-width picture on page 2
   (and every even page after it) begins 0.5 px lower than Word's and ends 0.75 px lower
   (edge rows covered 50% / 100% in ours, 100% / 25% in Word's). *Closed: 5.14.*
5. **Ink or advances is now settled by pixels too** (5.4 could not): `fidelity.py --ink`
   scores the device-rounded size above the exact size at every size, rounding up and
   down alike (8.5 pt 0.9914 / 0.9636, 9.5 pt 0.9964 / 0.9747, 11 pt 0.9753 / 0.9620,
   14.5 pt 0.9778 / 0.9538, 20.5 pt 0.9949 / 0.9606), and the 12 pt control scores 0.9992
   (0.9228 before: the 1.4% of ink that differed was pdfium's). Every committed document now
   scores `device` above `exact` too (the 4:2 split of 5.7 was the rasterisers').

**PowerPoint's PDFs** (a read-only check for `pptx2svg`, not changed here): `table-test`
and all 17 pages of `chart-gallery` pass. The 3-D chart scenes on `chart-gallery` pages
12-16 are 300 dpi bitmaps with a soft mask inside the PDF (1,812 x 970 px and alike, one a
page); no shared rasteriser un-rasterises those, but the rest of those pages is vector
and converts faithfully. The hinted-outline fault above was found here. Sharing: the
converter is instrument, not runtime, so not `ooxml-common`; it depends on neither
project, and would move as it is to a small development-tools package (or a copy per
harness, as `fidelity.py` itself is), taking its `ORACLE_DIR` as an argument.

**Cost.** Converting a page once takes about a second and is cached; a scoring run then
costs what it did (a resvg raster a page each side). The full recalibration of the
committed documents, both truths and both glyph sizes, takes about 12 minutes; the slow
fidelity tests about 8 minutes a truth (`sample-long` 357 s against pdfium's 320 s), so the
new instrument costs about a tenth more than the old and can run routinely. What is
occasional is the converter's full validation: 27 minutes for every committed page, about
an hour on four cores for all 765, and supersampling at 8x 3-4 minutes a page -- which is
part of why (b) was not chosen.

### 5.14 What the one-rasteriser instrument exposed, closed: table lines, shading, pictures

5.13 recorded three real differences (its items 1, 2 and 4). Each is now measured on a
probe of its own and drawn as Word draws it.

**Table lines** (`make_table_line_probe.py` / `read_table_line_probe.py`,
`tests/test_table_lines.py`). 288 two-row, three-column tables of empty cells, the first
row shaded: `w:tblInd` stepped by single twips over 24 of them, so the left grid line takes
every fraction of a pixel a twip reaches (5/24 px), the ties included; columns of
1500-1523 twips, so the inner and right lines do too; `w:sz` 2, 4, 6, 8, 12 and 24;
grid lines right of the text column's edge, the left one left of it, and every one left
of it; no vertical borders; no borders. Six documents: no `settings.xml` and mode 15, the
page's left margin at 1440 twips (300 px), 1442 (300.417) and 1438 (299.583).

| Rule (`layout.grid_line_px`) | Vertical lines where Word drew them (11,520) |
| --- | --- |
| **the model**: a line is centred on the text column's left edge rounded half up, plus the grid line's distance from that edge rounded **half away from zero** | **11,520** |
| before: the grid line rounded half up from the page's edge, in the layout unit | 8,260 |
| the whole position rounded half up from the page's edge, exactly | 8,296 |
| the text column's edge truncated, not rounded | 7,680 (the two 1438 documents: 0) |

So the committed documents' left outer border (`w:sz` 8, 4 px, its grid line 22.5 px
*left* of the text column's edge at 352.5 px) goes left, to 350-353, and every line right
of the edge goes right at a tie, as before. The other three sides were checked in the
same sweep: inner and right lines at a tie are right of the edge and round right, as
Word's do; a whole table left of the edge rounds every line left. **Refuted along the
way**: rounding the left border's outer edge (right for 5,864 of the first four
documents' 7,680, and wrong for a table indented 720 twips below mode 15).

**Cell shading** fills a cell between its grid lines **less the inner part of each
vertical border** there (the smaller half of an odd width in twips), each edge rounded
half up **from the page's edge** -- not from the text column's, as the lines are
(against the column's edge: 2,246 of 3,456 edges exact in the first four documents). So
where a line's width in px is less than its exact width -- `w:sz` 8 is 4.17 px, drawn 4 --
a column of white is left beside it, the column 5.13 item 2 saw at x = 1070-1074,
1790-1794 and beside the left border. With no vertical border the shading reaches the
grid line rounded, and leaves no column. Where the line above a shaded cell is short of
its exact width by half a pixel or more (`w:sz` 24: 12.5 px, drawn 12), Word fills the
row under it in the cell's colour, as wide as the line: drawn so. Cells whose shading is exact:
**4,969 / 5,184** (before: 3,915); every edge that differs is at exactly half a pixel,
where Word rounds down 216 times and up 304 by a rule not found: term-by-term rounding in
the layout unit (margin, indent, border half, each width) explains 331 of the 368 ties
of the first four documents, no simpler rule more, so it is recorded, not adopted.

Also measured there and left: in a table with no vertical borders, the horizontal lines
start 3 px left of the first grid line in Word (72 line ends a document, 144 px); in mode
15 two inner lines of a centred and a right-aligned table in the geometry probe are a
pixel left of the rule (a grid line a twip off Word's, whose snapped text starts hide it).

**Pictures** (`make_picture_place_probe.py` / `read_picture_place_probe.py`,
`tests/test_picture_place.py`; and the image boxes of Phase 4.7's `make_picture_probe.py`,
never read before). 71 pictures a document of extents that are not whole twips: at a
page's top as `sample-long`'s are, under `exact` 200-3000 and `atLeast` 3000, after space
before, under a 20 pt mark, two on a line, and after a word of text.

| Rule | Picture boxes within 0.03 px of Word's (962) |
| --- | --- |
| **the model**: the extent **scaled, both ways alike, to fit the extent truncated to whole twips** (`layout.picture_size`); its bottom, effect extent included, on the text's baseline -- under a taller mark at the bottom of the mark's height, an `auto` multiple's extra below it, an `atLeast` line's extra above it, and in an `exact` line **four fifths of the way down** (`LineHeight.object_drop`) | **961** |
| before: the exact extent standing on the rounded baseline | 0 |

`sample-long`'s full-width picture is 4,860 twips tall exactly, so it was only the
baseline: rounded to 1,313, where Word's picture ends at 1,312.5 -- half a pixel low on
every page it begins, the even ones. The one box off is not drawing: in mode 15 a picture
shorter than its run's text (49 px in an 11 pt run, under a 20 pt mark) takes a 56 px
line in Word and the model's 49; the line after it moves too (5.16).

**Scores** (the `svg` truth; device / exact):

| Document | Before (5.13) | After | Loss before / after (device) |
| --- | --- | --- | --- |
| `samplelib/sample-long.docx` | 0.9914 / 0.9908 | **0.9931 / 0.9924** | 5,516 / 2,597 |
| `samplelib/sample-simple.docx` | 0.9971 / 0.9957 | **0.9986 / 0.9971** | 1,473 / 746 |
| `wordto/sample-10pages.docx` | 0.9867 / 0.9832 | **0.9882 / 0.9847** | 6,558 / 5,410 |
| `wordto/sample-5pages.docx` | 0.985 / 0.9822 | **0.9866 / 0.9836** | 7,414 / 6,561 |
| `wordto/sample-with-images.docx` | 0.9983 / 0.9973 | **0.9983 / 0.9973** | 1,787 / 1,769 |
| `wordto/sample-with-table.docx` | 0.9919 / 0.9904 | **0.9991 / 0.9975** | 14,711 / 1,679 |
| median page | 0.9896 / 0.9856 | 0.99 / 0.9857 | |
| the other five | unchanged | | |

The pages rendered beside Word's and looked at: `sample-with-table` page 1 at the left
border and at x = 1071 in the shaded rows, and `sample-long` page 2 at the picture's top
and bottom edges -- no pixel differs in those crops.
### 5.15 The known drawing residuals, revisited

The residuals Phase 5 and Tables recorded, each against a probe. Two close, one closes by
5.14's rule, the rest are measured further and stay open.

| Residual | Now | Evidence |
| --- | --- | --- |
| A horizontal table border at a page's foot or top a pixel above Word's (Tables stage 6: 29 pages, 30 in mode 15) | **closed**: the band below the last row on a page -- the table's bottom, or a row split or ended at the foot -- is drawn **up from its bottom rounded** (`round(y + h) - w`), and the vertical lines stop where it starts; the band above a row stays drawn down from its top rounded | every horizontal edge of the recorded table probes (1,924 tops, 1,216 bottoms, before 1,063) and of the line probe (14,256 / 14,256, before 13,698); the pages probe's 25,064 displaced px a document (26,336 in mode 15) and the geometry probe's 6,127: 0; `sample-5pages`' split row: 1,789 px, 0 |
| A 4 px outer left border on a half-pixel grid line a pixel left of the model's | **closed by 5.14's rule** (the grid line's distance from the text column's edge rounds half away from zero) | 5.14 |
| Dotted and dashed underlines drawn solid | **closed**: the single underline's band cut into whole device px from the span's first pen x rounded, the last piece cut at the span's end -- dotted 6 px on, 6 off; dash 16 on, 8 off; the same at 11 and 20 pt in Calibri and Times New Roman (so fixed px, not scaled by size or thickness, at these sizes) (`layout.UNDERLINE_DASHES`) | `make_render_probe.py`'s `u-dotted` and `u-dash`: 8 / 8 underlines, every piece (was 0); the render probe's model-only px 55,731 -> 49,349; they no longer warn `underline-style-approximate` |
| Paragraph bottom borders: 25 of 150 a pixel off | **not settled** (below) | `make_para_border_probe.py` |
| Dotted and dashed borders drawn solid | **measured, not adopted**: the pattern's lengths are found, where it starts along a side is not | `make_border_pattern_probe.py` (below) |
| Double-line border joins | **not probed further** | the table border probe's residual after 5.14 and this section: 11,608 px the model's only and 1,696 Word's only a document, the dotted line and the double lines' corners and joins |
| Thick and double underlines | **measured further, not adopted** | below |
| The last line of a `distribute` paragraph | **not re-measured here** | 5.2 |

**Paragraph bottom borders** (`make_para_border_probe.py` / `read_para_border_probe.py`,
`tests/test_para_border.py`, which holds the model's count). Pairs of a one-line
bordered paragraph and a plain one, Calibri at every half point 10-22 pt, five borders
(`w:sz` / `w:space` 4/0, 6/2, 8/1, 12/4, 24/10), no space after and 120 twips after,
and at 11-15 pt under `auto` 360, `exact` 400 and `atLeast` 480; no `settings.xml` and
mode 15, which agree. Of 385 borders a document, on Word's row:

| Rule | Borders on Word's row (385) |
| --- | --- |
| the model: `w:space` (whole px, truncated) below the last line's text row | 234 |
| the border's exact position rounded: line top + pitch + space | 238 |
| `ceil(line top + pitch) + floor(space)`, the best of the two-term roundings tried (every sum of one to three of line top, text above, text below, pitch, space and width, rounded half up, down or up, plus a second such sum; and the three-part box rounding of `vertical.baseline_in_box` with the border's space as the space below the text) | 250 |

Word's border lies within about a pixel either side of the text's exact bottom plus the
space (−1.15 to +0.64 px), and for one size and border it depends on the line's top's
fraction, so it rounds against a position the model does not hold the way Word does.
**Probe that would decide it**: the same pairs with space before stepped by single twips
(moving the line's top through a pixel with everything else fixed), and a two-line
bordered paragraph (the border in the second line's box).

**Border patterns** (`make_border_pattern_probe.py` / `read_border_pattern_probe.py`):
paragraph boxes and tables in each style at `w:sz` 4, 8, 12 and 24. The runs away from
the corners, in device px (line width *w* = 2, 4, 6, 12):

| Style | On | Period |
| --- | --- | --- |
| `dotted` | *w* | 2*w* |
| `dashed` | 4*w* | 8*w* |
| `dashSmallGap` | 4*w* | 5*w* |
| `dotDash` / `dotDotDash` | runs of more than one length (the most common is half of them or fewer): not read apart | |

Every inner run of `dotted`, `dashed` and `dashSmallGap` fits, at every width (for
`dotted` at 2 px, 2,936 / 2,936).
Where the pattern starts is not settled: along a paragraph's top border it starts at the
same offset for three lengths of line (so it is anchored at the line's start, not
centred), but the offset differs by width without a rule found (a dash starting 6, 4, 34
and 28 px before the line for 2, 4, 6 and 12 px), the vertical sides have phases of
their own, and the corners are drawn as squares of both sides' colours. Drawing the
lengths with a guessed phase would move pixels without a measured gain, so the lines stay
solid and warn as before.

**Thick and double underlines** (the render probe's `u-thick` and `u-double`, eight
faces at five sizes): a double underline is two lines at the top and bottom edges of the
band a thick one fills -- its second line starts where the thick band ends, in all 40 --
so the two are one question: the band. Its top is half the underline position rounded
(30 / 40; the model's `ceil` 28); its bottom fits no linear rule of the `post` table's
position and thickness at the drawn or exact size (best 18 / 40). Not adopted beyond what
5.2 drew.

**Scores** (the `svg` truth, device / exact): `wordto/sample-5pages.docx` 0.9866 / 0.9836 ->
**0.9904 / 0.9874** (loss 6,561 -> 4,160: the border under the row split at page 1's foot);
median page 0.9900 / 0.9857 -> 0.9917 / 0.9879; every other document unchanged (no other
committed document has a border at a page's foot, a dotted or dashed underline). Page 1 of
`sample-5pages` rendered beside Word's at that border: no pixel differs.

### 5.16 A picture's line in mode 15 is at least its run's text height -- measured, adopted

Found by 5.14's picture probe (case 54, mode 15: a 49 px picture under a 20 pt mark took
a 56 px line, Calibri 11 pt's natural height). `make_picture_run_probe.py` /
`read_picture_run_probe.py`, `tests/test_picture_run.py`: pictures of 33, 49 and 66 px
alone in a paragraph, in a run of `w:sz` 16, 22, 40 or 60, under a mark of 22 or 40;
the baseline of the line after each. **Below mode 15 the mark's natural height floors the
line and the run's size does not count** (4.7, unchanged: 24 / 24); **in mode 15 the mark
does not count and the natural height of the run holding the picture does**
(`lines.object_line_height(..., runs=)`): 24 / 24, before 6 / 24. The picture stands at
the bottom of that height (5.14's `object_drop`), so case 54's box is now Word's too
(962 / 962). No committed document has a picture shorter than its run's text, so no
committed page moved.

### 5.17 The script offset's bound, revisited (stage D) -- not settled

5.5 left what rounds the bound (`round(A x w:sz) - round(A x s)`) open, and Times New
Roman and Arial poorly fitted. Re-scored on the recorded data (the offset probe's 1,380
offsets and the size sweep's 2,280, both kinds, 3,660 in all; `resolve.script_raise_half_points`
exact on 2,741):

* **Rounding**: each term and the difference rounded half up, down or up, in half
  points, device px, points, twips or font units, then to half points: the best, the
  terms in twips (the first rounded up, the second half up), 2,764 -- 23 more, no family
  of them close to exact. Not adopted: not a rule found, a fit.
* **The ascent the bound uses**, fitted per face over 1,200-2,600 font units: Times New
  Roman at best 62 / 95 superscripts (its ascent with line gap, 1,912, gives 15), Arial
  72 / 95, Calibri 66, Charter 67, Galvji 63, Helvetica Neue 4. So no choice of `A` makes
  the bound exact for these faces: **it is the bound's form, not its rounding or its
  metric, that is wrong** where it binds, and the probe 5.5 proposed (the offset crossed
  with the line gap and the `OS/2` typo and win metrics moved one at a time, sizes by
  half points) is still the one that would decide it. Not built in this pass: stages A-C
  were taken first, as ordered, and a wrong rule here would move glyphs Word draws
  exactly today.

### 5.18 Picture bullets (`w:lvlPicBulletId`) — measured, adopted

A list level may show a picture as its label: `w:lvlPicBulletId` names a
`w:numPicBullet` of the numbering part, holding a VML `w:pict` (what every Word writes;
without the frame shape type's `v:formulas` Word stops at the document with a dialog) or
a DrawingML `w:drawing`, whose picture is related from the numbering part. The layout
stopped at such a paragraph. `make_picture_bullet_probe.py` / `read_picture_bullet_probe.py`, 80
cases a page each, a PNG of our own as the bullet, no `settings.xml`, mode 14 and mode 15
(which drew alike to the pixel): label sizes 8 to 72 pt by whole and half points, the
picture's stated size (4.5 to 72 pt, 18 x 9, 9 x 18), its pixels (8 to 128 px square,
16 x 32, 64 x 16, a `pHYs` of 72 and 300 dpi), the level's `w:sz`, bold, face and
`w:position` and `w:vertAlign`, the mark's size and the text's, indents, `w:lvlJc`,
`w:suff`, line rules, space before, DrawingML. Three rounds: the first showed the size
did not follow the picture's stated size, the second that its pixels count, the third
that `w:lvlJc` -- which the model did not read, for any label -- moves it.

**The rules** (`linebreak.picture_bullet_units`, `label_shift`, `lines.PictureLabel`,
`layout._Placer._picture_bullet`):

* **A picture bullet is a square of `9 (s - 4) / h` pt** for a label of `s` half points
  (the mark's size under the level's `w:sz`; not the face, not the text's size) and a
  picture `h` pixels high -- whatever its stated size or its aspect, which is stretched
  to the square, and whatever density it states. A 16 px picture is `9/8 (pt - 2)` pt at
  every size from 8 to 72 pt; 8, 12, 24 and 48 px pictures twice, four thirds, half and
  a third of that, exactly. Its advance is its size.
* **It stands on the baseline** and takes part in the line's height by its size above
  it, with no descent, as a label does. Drawn, its left edge and its side are each
  rounded to a whole device pixel.
* **`w:position` moves it twice** (±6 half points, 13 px, draw it 26 px up or down; the
  line grows by the raise once); a `w:vertAlign` superscript by the script's raise once,
  at its full size.
* **`w:lvlJc` center or right** stands any label -- text or picture -- centred on or ending
  at its start, and the tab or space after it from its end.
* **A space after it (`w:suff` space) is an Arial space at the label's size** (0.2778
  em at 8, 11 and 24 pt; a level in Symbol or in Calibri alike).

| Lines within half a pixel / bullets' boxes (246 / 76 each) | before | after |
| --- | --- | --- |
| `picture-bullet-none`, `-14`, `-15` | 1 / 0 (the layout stopped at the first bullet) | **246 / 76** |

Not settled: pictures of 32, 64 and 128 px advance up to 18 layout units (0.02 px) off
the rule, and one wider than it is high about 460 units (0.47 px) short; no probe line
shows it. A picture in a format not read (EMF, WMF) is still a stop (`unmeasurable`).
Pinned in `tests/test_picture_bullet.py`. A filesamples document's one picture bullet (a
12 px GIF, page 8) is drawn on Word's box to the pixel.

### 5.19 Tab leaders — measured, adopted

A tab stop's `w:leader` was warned (`tab-leader-not-drawn`) and nothing of it drawn: the
render probe's one dot leader left 495 of Word's glyphs undrawn, and a filesamples
document's table of contents some 2,000. `make_tab_leader_probe.py` /
`read_tab_leader_probe.py`, one family a page, each line a paragraph (Calibri 11 pt, A4,
the left margin 1,442 twips: off the pixel grid and off every leader's grid): `phase` (a
left stop with a dot leader after `Lead` and 0 to 39 `i`: the text's end through a
dot), `end` (a right stop before `H` and 0 to 39 `i`), `kinds` (every leader before
left, right, centre and decimal stops), `faces` (every leader in Times New Roman 12,
Arial 9, Georgia 14, Cambria 26 and Calibri 7), `runs` (the tab in a run of its own at
20 pt, bold, italic, red, in Times New Roman, at 8 pt, underlined, raised; the text
around it formatted otherwise; an underlined tab with no leader; a leader tab ending its
line), `indent` (357 and 1,000 twips, a first-line indent, hanging), `short` (stops so
near that no glyph or one fits; two leader tabs on a line) and `align` (right-aligned,
centred, justified). No `settings.xml`, mode 14 and mode 15.

**The rules** (`layout._Placer._leader`, `linebreak.tab_leader`), every setting alike:

* **A leader is glyphs**: `dot` `.`, `hyphen` `-`, `underscore` `_`, **`heavy` `_` as
  well**, `middleDot` `·` (`LEADER_GLYPHS`) -- in the **tab's own run's format**: its
  face, size, weight, slant and colour, and a script's size and offset (`raised`: dots at
  7 pt on the superscript's baseline), whatever the text around it.
* **On a grid from the page's left edge**: every glyph at a whole multiple of its
  advance, the advance at the size drawn rounded to the layout unit (Calibri `.` at
  11 pt: 11,373.75 units, 11,374; a filesamples document's Ubuntu `.` at 12 pt: 12,091.39, 12,091) -- not from the
  column, the indent or the text's end (`indent`: the same dots under every indent).
* **From the first at or after the pen at the tab to the last that ends at or before the
  pen after it** (the stop, or where a right, centre or decimal stop puts the text):
  `phase` and `end`, every one of 41 and 41 lines; `short`, none or one where one fits.
  Nothing moves: the text after the tab is where it was.
* **An underlined tab is underlined across its room**, from the pen at the tab to the
  pen after it -- its stop where it ends its line -- as a run under a space is, leader or
  none.

| Lines within half a pixel (leader glyphs; rule px Word only), each setting | before | after |
| --- | --- | --- |
| `leader-none`, `-14`, `-15` (176 lines) | 8 (0 / 14,478; 17,160) | **176 (14,478 / 14,478; 0)** |

The render probe's leader draws its 495 glyphs where Word does (`render-none` 19,312 of
19,312 glyphs; `render-15` 19,274, the 38 left its distributed lines'). That filesamples
document's table of contents is drawn: every glyph Word draws there is drawn, and the
document's 9,631 of 9,631. The committed documents have no leader: fidelity as recorded.
A leader in a face whose advance is not known is still warned (`tab-leader-not-drawn`).
Not measured, and left: a bar tab (`w:val="bar"`), a leader beside a floating drawing, a
leader in a right-to-left paragraph. Pinned in `tests/test_tab_leaders.py`.

### 5.20 Without Word's faces: open substitutes, recorded symbol faces, coverage -- measured

Production feedback (2026-10): on a machine without Office's faces, a real document's
layout stopped on page 1 and its PNG was drawn in fallback faces. Every Calibri, Arial,
Times New Roman or Courier New paragraph is unmeasurable without its face. On the
committed fixtures, with only open faces visible, the commonest stop after the text faces
was **a symbol-font bullet** (Symbol `U+F0B7`, Wingdings `U+F0A7`/`U+F0D8`...):
`sample-resume.docx`, `style-document.docx` and a filesamples document all stopped at
their first bulleted paragraph.

**Open substitutes, measured** (`docx2svg.fonts.SUBSTITUTES`). Each pair was read with
docx2svg's own reader, against the copy Word lays out with, in all four styles:

| Word's face | Open face | Advances compared / differing | Line metrics, script sizes and offsets |
| --- | --- | --- | --- |
| Calibri | Carlito 1.1 | 2,094 / 3 (U+0192, U+026A, U+0299) | equal |
| Arial | Liberation Sans 2.1.5 | 2,209 / 38 to 40, none Latin, Greek or modern Cyrillic | equal; sampled `kern` pairs equal |
| Times New Roman | Liberation Serif 2.1.5 | 2,209 / 40 to 50, likewise | equal; sampled `kern` pairs equal |
| Courier New | Liberation Mono 2.1.5 | 2,118 / 179, the combining marks | equal |
| Cambria | Caladea | -1.8 to -5.7% over a sentence | 1.150 em against 1.172: **not compatible** |
| Calibri Light | Carlito | +1.2 to +3.4% | **not compatible** |

Only the first four are substituted, and only where Word's face is neither installed nor
embedded. Each substitution is reported (`font-substituted`). What still differs is the
drawing: outlines, and the underline and strikeout geometry (Carlito's underline is at
-103 units where Calibri's is at -232). Also, Carlito has no legacy `kern` table, so a
Calibri run Word kerns (`w:kern`) is laid out unkerned. With only the substitutes visible,
`layout-sweep.docx` (Calibri) lays out glyph for glyph as from Word's recorded Calibri
numbers, which `tests/test_render.py` holds to Word's PDF. A filesamples document
(Courier New, Symbol, Wingdings) and two probes lay out identically to their layout in
Word's own faces, line by line. A caller-named substitute (`font_substitutes`) is used
and reported as approximate. With Caladea for Cambria, `sample-long.docx` lays out to 19
pages where Word draws 36, because the break differences cascade through its keeps. That
is why such a substitute is never a default.

**Symbol and Wingdings, recorded** (`docx2svg.recorded`, written by
`tools/record_symbol_faces.py`). These are Microsoft's faces, with no open clone. A bullet
needs only these numbers from them: the vertical metrics (SymbolMT's 2,059 + 450, the
1.2251 em line of 2.4), the script sizes, the decoration integers and an advance per
character (Symbol's 188 private-use codes and the same codes below U+0100; Wingdings'
223). They are facts, recorded from the copies Word lays out with: Word's bundle's
SymbolMT and macOS's Wingdings. No outline is recorded, and no file. Where the face is
absent, the layout answers from them. The glyph is drawn as its Unicode equivalent (•, ▪,
➢, ❖, ✓, ☑...: `recorded.UNICODE`, each checked against the face's own glyph by eye) in a
generic face, at the position Word's face gives it. `tests/test_substitutes.py` holds the
module to the installed faces where they are present.

**Coverage** (`docx2svg.coverage`). Every conversion now sets `ConvertOptions.coverage`
(also `layout.coverage`), a summary of the warnings: the pages laid out, the blocks laid
out and skipped, the first stop (reason, page, element path), every header, footer and
text-box stop, and the faces substituted or missing. Where the layout stopped, the
estimated page count is Word's own count from `docProps/app.xml`, and otherwise none is
invented. `complete` is true only when nothing was skipped, so a caller (docx-agent's
`check`, `render` and `save_document`) can tell "laid out, and nothing was wrong" from
"could not lay it all out".

**Word verification.** Word exported the symbol-bullet probe: Calibri body text, then six
bullets in Symbol `U+F0B7`, Wingdings `U+F0A7`, `U+F0D8`, `U+F076` and `U+F0FC`, and Courier
New `o`, in one list. Under the oracle's lock, Word closed it without saving. In
docx2svg's layout, with only Carlito, Liberation and the recorded Symbol and Wingdings
faces visible, Word's PDF matches glyph for glyph (`tools/glyphs.py`'s comparison, layout
against PDF): **420 of 420 glyphs matched; x 420/420 (object starts 16/16 exact, advances
404/404); baseline, face and size 420/420; nothing extra or undrawn.** The worst advance
was 0.517/1000 em, which is within the instrument. This is the same score as the layout in
Word's own faces. The bullets' pen positions, the text after them at the hanging indent
and the lines' baselines (SymbolMT's taller line among them) are Word's. The drawn glyph is
the only thing that differs: a Unicode equivalent instead of the symbol face's outline. The
probe and the scripts that build and score it are kept locally, not committed.

---

## Tables — measured (Phase 6, Tables row)

`docx2svg.table` lays a top-level table out: it resolves the table style's chain (the
default table style where the table names none) and the table's own `w:tblPr`, a
row's `w:trPr` and a cell's `w:tcPr` over them with the style's conditional formats
between; places the grid; gives each cell a text start and a line budget; breaks each
cell paragraph with the Phase 3 breaker at the cell's width and stacks its lines with the
vertical model; and sizes each row. The paginator places a table row by row
(`paginate.TableItem`, `table.place_table`), and the page walk draws exactly what it
placed (`layout._table`). What is not modelled yet raises `table.Unsupported` and the
table is an obstacle, as every table was before: the layout stops there with the stop
band and a `layout-stopped:table` warning that says why. `table.STAGES` lists what is
modelled; each stage below switched its part on with its probe.

Each stage below is a committed `make_table_*_probe.py`, exported by Word 16.106 through
`tools/oracle.py`, scored by `tools/read_table_probes.py` -- the glyph-position check
(5.3) over every cell: a cell's text start is a text object's start, held exactly; a
right-aligned word ends where a cell's line may end, so the line budget is held too;
every baseline on Word's device row -- and recorded into
`tests/fixtures/table-observations.json`, from which `tests/test_tables.py` holds the
model to all of it offline. Nothing in the model is taken from the export.

### Stage 1 -- the grid: fixed layout, widths, margins, borders, indent, alignment

`make_table_geometry_probe.py`: 45 one-row tables, each varying one thing from a base
(fixed layout, three 2000-twip columns, 108-twip margins, `w:sz` 4 borders): margins
0-200 and none stated; border widths 2-27, none, one thick side; `w:tblInd` 0, 100,
333, -200, 720, stated by the style only and overridden; centred and right-aligned
tables with and without an indent and with thick borders; uneven, narrow, odd-twip and
over-wide columns; a cell's own `w:tcMar`; cells whose `w:tcW` disagree with the grid;
autofit tables whose every width is stated. Each cell holds a left-aligned word and a
right-aligned one. Four documents: no `settings.xml`, modes 12, 14, 15.

**Every cell's text start and line end, in every setting: 552 / 552**; through the
whole layout, **1,583 / 1,583 glyphs per document** at Word's pen position and baseline.

| Rule | Evidence |
| --- | --- |
| **Widths**: `w:tblGrid` as authored -- unless every cell states a `dxa` width (`w:tcW`) and those put the grid lines elsewhere: then the cells' widths (`table.column_widths`) | a 1000/3000/2000 grid under cells of 2000 each is laid out on 2000s, fixed and autofit alike |
| **Where the grid starts** (`table.grid_origin`): **mode 15** -- the left border's outer edge at `w:tblInd`; **below 15 with `w:tblInd` stated** (by the table or its style) -- the first cell's *text* at the indent; **below 15 with none stated** -- the first grid line at the margin; **centred / right-aligned** tables ignore the indent: below 15 the grid is centred, or the last cell's text ends at the right margin; in mode 15 the outer border edges are centred or end there | every indent and alignment case, in all four settings |
| **A cell's text starts** at its left grid line plus the larger of its left margin and the inner part of its own left border, **snapped to a whole device pixel** -- held in Word's layout unit (740 px is 740.0004: 727,450 units, as Word's PDF prints every cell start) | unsnapped: 21 / 552 |
| **A cell's line may run** from that snapped start the grid distance less the larger of each margin and the inner part of the border on each side -- so its right edge moves with the snapping | every right-aligned word's end, to 0.001 px |
| **Borders** take their width in **whole twips, truncated** (`w:sz` 5 is 12 twips, 27 is 67), a `double` three widths; the inner part of a border is the smaller half, the part beyond a table's outer edge the larger | exact widths, halves rounded: 551 / 552 (the mode-15 `w:sz` 5 middle cell a pixel right) |
| **Default cell margins**: with no `w:tblCellMar` anywhere -- no table style either -- 10 twips left and right (the text 2 px in, the line 20 twips short) | ECMA-376's 0: 528 / 552 |
| **A tie below mode 15** (a text start of exactly *n*.5 px -- 1440 + 108 twips is 322.5 px -- in a table with no `w:tblInd`): Word holds the table's edge (the margin less the first cell's inner border part) and the text's distance from it as two lengths, each truncated to the unit, and the text goes a pixel right exactly where the border part's unit fraction is 0.6 or more (`w:sz` 1, 2, 5, 6, 9, 10, 18) | plain rounding: 540 / 552; over a scratch sweep of 17 widths x 16 margins about the tie, 439 / 442 against 436 |

Refuted, with scores (cells whose start and end are both exact, of 552):

| Hypothesis | Score |
| --- | --- |
| **the model** | **552** |
| below 15, no `w:tblInd` read as 0 (the text at the margin) | 291 |
| mode 15 placed as below 15 (the text at the indent) | 441 |
| the text start not snapped | 21 |
| `w:tblGrid` always, the cells' widths ignored | 536 |
| ECMA-376's default margin of 0 | 528 |
| exact border widths | 551 |
| no tie rule | 540 |

Not settled: the tie rule in mode 15 (a scratch sweep about the tie leaves 2 of 442
cases a pixel off by any rule of the families tried) and below 15 (3 of 442).

**What stage 1 left as obstacles** (`table.STAGES` was `{"grid"}`): a stated row height
or vertical alignment (stage 3), merged cells (4), a cell's own or a conditional format's
borders or shading, a table border that is not a single line (5), a row that does not fit
on its page (6), a floating or nested table, an autofit table whose content would widen a
column (7). Every real document here still stops at its first table: all of them have a
table style with conditional borders. The parser no longer warns
`table-layout-unsupported`: where a table is not laid out, the layout says so.

### Stage 2 -- a cell's content, and finding 6

`make_table_content_probe.py`: 60 one-row tables (fixed layout, 108-twip margins,
`w:sz` 4 borders), each a family's case: a line budget composed to δ layout units of
the text (δ = -1024, -2, 0, 2, 1024; and with six trailing spaces), three paragraphs with
space before/after 120/0, 0/120, 120/240, 240/120, 100/100, contextual spacing, Word's
`Normal.dotm` spacing; two-line paragraphs under `exact` 300, `atLeast` 400, `auto` 360
and 200; indents (left 200, right 300, first line 360, hanging 360); centred,
right-aligned and justified text; default and custom tab stops; a numbered and a
bulleted list; 9, 14 and 20 pt cells and a cell of mixed sizes; and one contextual
paragraph alone in a cell between every pairing of its style and the styles around the
table (36 cases). Four documents (none, 12, 14, 15).

**Every glyph of every cell where Word drew it: 3,558 / 3,558 in each of none, 12 and
14**; in mode 15, 3,554 matched, every one but the 24 glyphs of the justified case exact
-- mode 15 fits more words on a justified line than the breaker does, the known Phase 3
gap (4.6, 5.10), no table's.

| Rule | Evidence |
| --- | --- |
| **A cell paragraph is laid out as a body paragraph at the cell's line budget**: the Phase 3 breaker (inclusive, no slack: the word stays at δ = 0 and 2, goes at -2), the vertical model, indents from the cell's text start, tab stops (default and custom) measured from it, alignment against its line end, list labels | every family but the edge rules below |
| **A cell stacks its paragraphs as the body does**: the first one's space before counts, neighbours collapse to the larger, the last one's space after counts (`table.cell_stack`); **no spacing collapses across the table's edge**: the paragraph before the table keeps its whole space after, the one after it its whole space before | the spacing family; `make_table_row_probe.py` |
| **Contextual spacing at a cell's edges**: a contextual first paragraph loses its space before when it is of the style of the paragraph *before the table*; a contextual last paragraph loses its space after when it is of the *default* paragraph style -- whatever follows the table; and a contextual paragraph before a table loses its space after when the first cell starts with a paragraph of its style (`paginate.table_gap_px`) | the 36 edge cases; inside the cell only: 8,640 / 10,674 glyphs |

**Finding 6, settled** (`make_table_style_size_probe.py`: Normal 12 pt over
`w:docDefaults` 11 pt, six table styles -- none, only table properties, a size of 9 pt, a
right alignment, a line spacing, a colour -- three paragraph styles -- Normal, one based on
it stating nothing, one stating 16 pt -- in eight settings -- none, 12, 14, 15, 14 and 15
with `overrideTableStyleFontSizeAndJustification`, 14 and 15 with a centred Normal -- and
eleven more pairings of `w:docDefaults` 8-14 pt and Normal 10-14 pt): **19 documents,
9,462 / 9,462 glyphs** at Word's size and position. With no mode stated and in modes 12
and 14, unless the `overrideTableStyleFontSizeAndJustification` compatibility setting is
on (every document Word 2013 and later writes carries it beside mode 15), **in a table
whose style states any paragraph or run property, a paragraph whose styles give it 12 pt
takes the table style's size instead** -- or, where the table style states none,
`w:docDefaults`' -- unless that is 10 pt, which counts as unset too, and the 12 pt stays
(`resolve.cascade.legacy_table_size`). Twelve points was Word's default before 2007, and
this is Word treating it as "no size". Any other size stays: Normal at 10, 11 or 14 pt
beats a table style's 9 or 15 pt -- which is the style probe's p06-p08 (Normal 11 pt, a
table style's 14 pt: Normal wins), the other side of the same rule -- and a paragraph
style stating 16 pt keeps it. Alignment is not affected by a centred Normal.
`filesamples/sample1`'s Normal is 12 pt and its `TableGrid` states a line spacing: its
cells are 11 pt, as Word drew them -- its glyphs in the resolved face and size go from
2,020 to **2,080 / 2,095**; the other 15 are the baseline checker losing its place in the
line under that table.

| Refuted (glyphs of the 19 size documents) | Score |
| --- | --- |
| **the model** | **9,462** |
| ECMA-376 (Normal above the table style, whatever the mode) | 7,147 |
| `w:docDefaults` of 10 pt taken like any other | 9,221 |
| the Normal style's own size given way wherever the table style states something (the first hypothesis: fails every document whose Normal is not 12 pt) | refuted by the p06-p08 style cases and the `sizes-*` documents |

Not measured: whether a Normal stating *left* alignment gives way to a table style's (the
setting's name says justification too; a centred Normal does not give way).

Stage 2 changes no real document yet: each of them still stops at its first table, whose
style has conditional borders (stage 5).

### Stage 3 -- row heights

`make_table_row_probe.py`: 24 three-row, two-column tables, a one-line cell beside a
three-line one (so the row is the second cell's height, and the first has room to move):
`w:trHeight` `atLeast` 200 (under the content), 1000, 1001, 1003, 1203, and 1000 with no
`w:hRule`; `exact` 400 (under the content), 800, 801, 1203; top, inside and bottom
borders of 24/4/12 with and without `atLeast` 1000 and `exact` 800, and 4/24/4 under
`exact` 800; `w:vAlign` centre and bottom, alone, with `atLeast` 1000, `exact` 800, and
with cell margins of 100 above and 50 below; margins with `atLeast`. Four documents
(none, 12, 14, 15), which agree.

**Every glyph where Word drew it: 2,116 / 2,116 per document.**

| Rule | Evidence |
| --- | --- |
| A row is as tall as its tallest cell: its paragraphs (the cell stack, stage 2) and its top and bottom margins (`table.row_height`) | every case |
| Each **horizontal border** takes its width once, between the rows it parts (the table's top border above the first, its bottom border under the last), as wide as the widest border stated on that edge; margins and borders are **outside** the lines' boxes -- a cell's first line is a line on its own at the row's top plus border plus margin | the borders family; the scratch vertical probe (120 cases): this 467 of 480 (the 13 are rows Word moved to the next page), the margins or borders in the first line's box 420-436 |
| **`atLeast`** (and a height with no `w:hRule`): the row is at least the stated height **plus** the cells' top and bottom margins -- the height is compared with the paragraphs alone | margins 100/50 under 1000 twips: 240 px, not 208 |
| **`exact`**: the stated height **less the border on the row's top edge** (the table's top border for the first row, `insideH` below it), whatever the content, which is cut | `w:sz` 24 on top, 4 inside: every line exact |
| **`w:vAlign`** centre and bottom move a cell's lines down in the room its row leaves between its margins -- never up: content taller than an `exact` row stays at the top | the vAlign family |
| **Row heights do not round**: a row is an exact length, and each line's baseline rounds in its own line box as the body's does | heights of 1001, 1003, 801 and 1203 twips, off the pixel grid |

| Refuted (glyphs of the four documents) | Score |
| --- | --- |
| **the model** | **8,464** |
| `atLeast` holds the margins (ECMA-376: the height is the row's) | 6,632 |
| `exact` is the height below the top border | 5,068 |

### Stage 3a -- an empty row, and `w:hideMark`

Found on local documents: their empty spacer rows (a small `w:trHeight` `atLeast`, one
empty paragraph per cell) stood shorter in Word than in the model -- by the paragraph's
line less the stated height -- and the error carried to everything below. The first reading -- empty paragraphs standing shorter -- was
wrong: every empty paragraph *between* tables was already exact. The rows' cells carry
**`w:hideMark`**, which the model parsed and ignored. **Probed before it was implemented**:
`make_hide_mark_probe.py` / `read_hide_mark_probe.py`, `tests/test_hide_mark.py`
(`tests/fixtures/hide-mark-observations.json`): 48 cases, one per page, each a three-row
table whose middle row is the case, read by the distance from the top row's baseline to
the bottom row's and to the line after the table; no settings, modes 12, 14 and 15, which
agree case for case.

**The rule** (`table.hides_mark`, `CellParagraph.hidden`, `cell_stack`): **in a cell with
`w:hideMark`, a last paragraph that is empty takes no room at all** -- not its line,
whatever its mark's size (6, 10, 11, 48 pt), and not its space before or after. A row of
such cells is its `w:trHeight` (`atLeast`, or with no `w:hRule`) plus the cells' top and
bottom margins; with no `w:trHeight`, its margins alone -- 0 px without, the row only its
borders. What the probe settled around it:

* **Only the last paragraph.** Two empty paragraphs are one line; a no-break space or text
  before an empty paragraph is that one line (and its space after still counts, 240 twips
  after the no-break space: 50 px more); an empty paragraph with space before after it
  adds nothing.
* **"Empty" is `paginate.is_empty`'s**: an empty run is nothing; a space, a no-break
  space or a tab is something (the row stays a line tall).
* **Only its own cell**: an empty cell beside it without the element keeps the row a line
  tall; a cell with text beside it, likewise.
* The same **first, last and alone** in its table, under `w:vAlign` centre and bottom,
  cell margins and `insideH` of `w:sz` 24; `exact` rows were already exact.
* **Body paragraphs are not affected**: an empty 12 pt paragraph between two tables, or
  after one, with or without `w:hideMark` on their cells, stands its full line (exact
  before and after).

| `hide-mark` (48 cases per setting) | model before | after |
| --- | --- | --- |
| `control` (no `w:hideMark`, or cells with text) | 4 / 4 | 4 / 4 |
| `height` | 4 / 7 | **7 / 7** |
| `content` | 5 / 18 | **18 / 18** |
| `cells` | 5 / 15 | **15 / 15** |
| `body` | 4 / 4 | 4 / 4 |
| lines within half a pixel | 141 / 190 | **190 / 190** |

The same in every setting. Refuted and pinned: `w:hideMark` ignored (the model before,
above); every empty paragraph of such a cell hidden, not just its last (off in `empty2`,
with and without a stated height).

The local documents (`tools/fidelity.py`, svg truth, device glyphs; scores in their local
notes): every table line now on Word's row, and their SSIM from the corpus's lowest to
within 0.0001 of 1. No committed document has a `w:hideMark`: their output is
byte-identical (every raster a cache hit). The other local documents that carry it: unchanged.

**Not measured:** a hidden mark in a vertically merged cell, a nested table's cell, a cell
whose last paragraph is contextual or has borders; where Word draws the hidden mark (it is
a space, and the model draws no mark).

### Stage 4 -- merged cells

`make_table_merge_probe.py`: nine tables of three 2000-twip columns (four for
`w:gridBefore`), every cell a left-aligned and a right-aligned word, cells beside a merge
of one or three lines: a cell over two of three columns (`w:gridSpan`) beside a single
one; a cell over all three; spans in both rows, offset; the first column merged down
three rows (`w:vMerge`) holding one line -- top, centred and bottom -- and holding seven
lines, taller than its rows; merged down two of three rows; a cell over two columns
merged down two rows; a row starting one column in (`w:gridBefore`, `w:wBefore`) and one
ending a column short (`w:gridAfter`). Four documents (none, 12, 14, 15), which agree.

**Every glyph where Word drew it: 674 / 674 per document (2,696).**

| Rule | Evidence |
| --- | --- |
| A spanning cell's box runs from its first grid column's left line to its last one's right line; its text start and line budget are a single cell's there (stage 1) | the span family |
| A `w:gridBefore` row's first cell starts at its grid column's line; nothing is drawn in the columns it skips, before or after | grid before / after |
| A cell merged down counts over all its rows, not in its first: each row is sized by its other cells (`row_height`), and the merged cell's lines sit in the height of all of them and the edges between (`w:vAlign` centre and bottom in that height) | vMerge one line, top / centred / bottom |
| **A merged cell taller than its rows makes the *last* of them taller** (`table._merge_groups`, `flow_table`) | the seven-line case |

| Refuted (glyphs of the four documents) | Score |
| --- | --- |
| **the model** | **2,696** |
| the first row of the merge grows | 2,368 |

Across pages (stage 6), a row holding a merged cell that does not fit on its page is not
split: it moves whole to the next page, and one that fits on no page raises
`table.Overflow` (the layout stops).
Not measured: splitting a merged cell across pages. *Measured and split since stage 6a.*

### Stage 5 -- borders and shading, where two disagree

`make_table_border_probe.py`: two-row, two-column tables, each cell a left-aligned and
a right-aligned word, a different colour per border (so Word's fills say which border it
drew): the first cell's right border against the second's left -- 4/12, 12/4, one width
in two colours, `single` against `double`, a border against `nil`, against none; and
with no cell margins, where each text starts at the inner half of the border between;
the same for the edge between rows; a thin cell border under a thick `insideV` /
`insideH`, a `nil` cell border under a table border; `double` and `dotted` lines, `w:sz`
2 to 24; a cell's `w:shd` fill and the table's, with and without borders; and a table
style's conditional borders (a `firstRow` bottom of `w:sz` 18 against `insideH` 8, as
the real documents' Light Grid has). Four documents (none, 12, 14, 15).

**Every glyph where Word drew it: 9,812 / 9,812** over the four documents, and the
filled rectangles scored beside them (`read_render.rect_score`).

| Rule | Evidence |
| --- | --- |
| **The wider border wins** an edge two cells (or a cell and the table) state -- whichever level states it; a `double` line counts as three widths; `nil` or no border loses to any border (a `nil` cell border under an `insideV` of 24 leaves the 24); at equal widths the left or upper cell's (`table.edge_winner`) | the vertical, horizontal and cell-over-table families; the colours drawn |
| **A cell's text starts from its *own* left border**, not the one Word draws there: a cell whose left border is 4 beside a neighbour's 24 starts its text as if the 4 were drawn (`left_offset`); the line budget's end is set by the border that wins (`budget_left`, `right_offset`) | the no-margin cases |
| A border's inner part is the smaller half of an odd width in whole twips (stage 1); the edge between rows is as tall as the widest border stated on it (`_edge_width`) | `w:sz` 18 between cells with no margins: 22 twips each side |
| **Drawing** (`layout._row_borders`): a cell's shading fills it between its borders; the edge above a row is a band as tall as its widest border, each cell's winner drawn from the band's top at its own width; each vertical edge's winner is centred on its grid line (the larger half to the left) and fills the band where it crosses it; a line is its width in whole twips truncated to whole device px, a `double` a line, a gap and a line of that width | the shading and double cases, Word's filled rectangles |
| A conditional format's borders are the cell's own, over the table's (`_cell_borders`), and wear the same rules | the style case |

| Refuted (glyphs of the four documents) | Score |
| --- | --- |
| **the model** | **9,812** |
| each cell's text from the border drawn on its edge | 9,792 |
| a cell's own border beats the table's, whatever its width | 9,796 |
| a double border counted one width | 7,376 |

Filled rectangles, per document (mode none; 15 within a few px of it): the model's
851,381 px against Word's 841,469, 832,525 in common. The residuals: the `dotted` line,
drawn solid of its width (8,844 px the model's only -- dotted and dashed patterns are not
transcribed yet); the `double` lines' corners and joins (9,528 px the model's only,
8,388 Word's); a pixel column or row at single lines' joins elsewhere (under 500 px).

**The real documents' tables are laid out from here on.** Every committed document's
table has a table style with conditional borders and shading (Light Grid, Light
Shading); with stage 5 they are drawn, and every glyph of them is where Word drew it:
`wordto/sample-with-table` 1,141 / 1,141 (was 238 drawn, the layout stopping at the
table), `samplelib/sample-simple` 844 / 844 over two pages (was 759 on one),
`wordto/sample-10pages` 4,211 / 4,211 (was 3,656). `wordto/sample-5pages` lays out its
table's first four rows and stops at the fifth, which does not fit on the page (stage
6). Their borders and shading (`tests/fixtures/render-observations.json`,
`tests/test_render.py`) are Word's filled rectangles but for two residuals, not probed
yet: Word leaves a one-pixel column unshaded on each inner grid line of a shaded row
where no vertical border is drawn (`sample-simple`: 3 columns of 112 px), and draws a
table's left outer border (`w:sz` 8, 4 px, on a grid line at a half pixel) a pixel left
of the model's, while the inner lines at half pixels are exact. *Both closed in Phase 5.14
(`make_table_line_probe.py`): the column is left beside a drawn border, not where none
is -- the shading stops at the border's exact inner half, the line is drawn a whole px
narrower -- and the grid lines round against the text column's edge, half away from it.*

**Found later: the shaded rows' text was painted over** (Phase 5.12). The cell shading
was emitted after the text, so every banded row drew as an empty shaded row, every
glyph still at Word's position; raster SSIM 0.787, recorded as a gain. Word paints a
row's borders, then each cell's shading and its text; so does the SVG now (0.8875), and
`tests/test_visibility.py` fails any glyph that draws no ink.

### Stage 6 -- tables across pages

`make_table_pages_probe.py`, every case starting a page: **fit** -- an anchor line, a
filler paragraph of one `exact` line of 12,000 twips whose space after is swept by
single twips, and a table of one-line rows whose second row ends at the page's foot;
borders of `w:sz` 4, 24 and none, a cell bottom margin of 100, the row the table's last
or a middle one -- ten variants, each bracketing to a twip where Word first moves the row;
**split** -- body lines filling the page to a few lines from its foot, then a table whose
second row does not fit: cells of one-line paragraphs (5 and 3 lines) with room for 0-3
of them; a six-line paragraph with room for 1-3 lines; `w:sz` 24 borders with margins of
100; paragraphs with 240 twips before; one and two header rows (`w:tblHeader`); a
`w:cantSplit` row, and one taller than a page; rows of one line. Four documents.

**Every glyph where Word drew it, on Word's page: 27,136 / 27,136** (6,782 in each of
none, 12 and 14; 6,790 in 15, whose tall `cantSplit` row runs on).

| Rule | Evidence |
| --- | --- |
| **A row stays on the page when it ends, with the border under it, at or above the page's foot** -- inclusive; the thresholds agree to one twip in all ten fit variants | the fit sweep |
| **A row that does not fit is split at the foot**: each cell keeps the lines that end above it with its bottom margin and that border; the rest goes on at the next page's top under the row's top border and the cells' top margins again (`table.place_table`) | the split family |
| A line starting a cell on the next page **keeps its paragraph's space before** | 240 before every line |
| **Widow control in a cell only in mode 15**: a split paragraph keeps two lines on each side there; unstated, 12 and 14 may leave one | the six-line cases |
| Where a cell with lines left would keep none, the whole row moves | room for 0 lines |
| A `w:cantSplit` row moves whole; one taller than a page, at a page's top, is split in mode 15 and runs off the page below it in the others | the cantSplit cases |
| On each page the table continues on, its header rows (`w:tblHeader` from the first row) come first | one and two header rows |

| Refuted (glyphs of the four documents) | Score |
| --- | --- |
| **the model** | **27,136** |
| the border under a row not counted | 26,540 |
| a continued row's first paragraph loses its space before | 27,080 |
| widow control in cells in every mode | 26,698 |
| no widow control in cells | 26,990 |
| header rows not repeated | 26,804 |

The probe's borders are Word's filled rectangles but for one residual, not probed yet:
on 29 pages of each document (30 in mode 15) a horizontal border line at a page's foot
or top (under the last row placed, or over a row going on) is drawn a pixel above Word's
-- the same area, 25,064 px of 1,792,490 displaced (26,336 in mode 15). *Closed in Phase
5.15: the band below the last row on a page is drawn up from its bottom rounded.*

`wordto/sample-5pages` is laid out to its end now: its table's fifth row, which does not
fit, is split at the page's foot and goes on at the next page's top, **2,727 / 2,727
glyphs** where Word drew them over two pages (was 1,403 on one, the layout stopping at
the table), SSIM 0.837 (none before: no page was drawn in full).

**What stays an obstacle** (`table.STAGES` is every stage above): a floating table
positioned as no probe measured (a floating table is laid out since F.18; a table nested in a cell since "A table nested in a cell"), a floating
drawing text wraps around (drawn since F.16 and F.17 but below mode 15 against the page) or other unmodelled
content in a cell, a line taller than a page (a merged cell split across pages
since stage 6a), and what of autofit stage 7b leaves unsettled (a cell
across columns wider than they are, below). Each stops the layout with the stop band and
says which.

### Stage 6a -- rows at a page's foot: merged cells and header rows

docx-agent saw Word split a row holding a vertically merged cell at a page's foot, the
merge's content going on at the next page's top, where docx2svg moved the row whole; and
once saw Word keep a two-line header row whole where docx2svg split it.
`make_table_foot_probe.py` / `read_table_foot_probe.py` (`tests/test_table_foot.py`):
every case starts a page, body lines filling it to a few lines from its foot -- the fill
swept a line at a time -- then a table of two 3000-twip columns. **merge**: a one-line row,
then three rows whose first cell is merged down over all three beside two-line cells,
the merged cell holding one line or seven (more than its rows); **merge row**: a row whose
first cell starts a two-row merge of five lines beside three, so the row itself splits;
**header**: a header row (`w:tblHeader`) of two lines -- two one-line paragraphs, one
wrapped paragraph, one under `w:keepLines` -- and the same row not a header; **keep**: an
ordinary row of `w:keepLines` paragraphs. 37 cases; no `settings.xml` and mode 15.

**The rules** (`table.place_table`, `RowPiece.carry`, `layout._Placer._table`):

* **A row holding a cell merged down splits as any row does** (stage 6): its own cells
  keep the lines that end above the foot, the rest going on at the next page's top. **The
  merged cell's lines run on through the merge's rows**, ignoring the edges between them,
  as on one page (stage 4): those that end above the page's break stay -- within the
  split row's part, or down to the foot where the row is the merge's last -- and the rest
  go on at the next page's top, from the top of the merge's first row there, with the
  cell's top margin again. A break between two rows of the merge is a break like any.
* **The merge's last row is as tall as its own cells or the merged cell's lines still to
  place, whichever is taller, on each page** (stage 4's rule, per page): the merge's last
  row splits at the foot for the merged cell's lines alone, its own cells done.
* **The merged column has its top border at the page's top** where the merge goes on, as
  a row going on has its own.
* **A header row is never split**: it moves whole, as a `w:cantSplit` row does -- two
  one-line paragraphs with room for one go to the next page, where an ordinary row's
  split.
* **In mode 15 the table's header rows are not left at the foot with no row after them**:
  the table starts the next page. With no `settings.xml` they stay there alone, and are
  repeated over the next page's rows.
* `w:keepLines` keeps nothing whole in a cell: without `settings.xml` a row of
  `keepLines` paragraphs splits as any; in mode 15 widow control moves it, as stage 6 had
  it.

| Glyphs (Word drew 7,531 / 7,411) | none: matched, baselines, misplaced | 15: matched, baselines, misplaced |
| --- | --- | --- |
| before: a row holding a merged cell moved whole, header rows split | 7,373, 7,000, 158 | 7,331, 7,050, 200 |
| a header row split like any row | 7,447, 7,318, 84 | 7,405, 7,368, 6 |
| mode 15: header rows left at the foot with no row after them | -- | 7,411, 7,411, 120 drawn twice |
| **the model** | **7,531, 7,531, 0** | **7,411, 7,411, 0** |

Every rectangle is Word's too (653,090 and 628,514 px; before, 124,026 and 110,448 px drawn
where Word drew none). The stage 4 and stage 6 probes are unchanged (every glyph, as
recorded). Not measured: a merge going across three pages (its rows above a page's first
taken as laid out one after another), a merged cell's `w:vAlign` across pages (drawn from
the top), and a header row taller than a page.

### Stage 7 -- autofit: recorded (modelled since stage 7b, below)

Every committed document's table is autofit (`w:tblW` auto, no `w:tblLayout`), and each
is laid out exactly on the widths its file states: Word writes the widths its autofit
arrived at into `w:tblGrid` and every cell's `w:tcW`, and keeps them while the content
fits. The model does not fit columns to content. It lays an autofit table out on its
stated widths and stops the layout (`table.autofit_unmodelled`, `flow_table`) at the three
cases the probe below shows Word resizing, or does not measure: a word wider than its
column, a cell with no width in twips, and a table width in percent.

`make_table_autofit_probe.py` / `read_table_autofit_probe.py`: 20 one-row, two-column
autofit tables, `w:tblGrid` 2000 / 2000 twips (417 px each), margins 108; the first
cell one unbreakable word of *n* `H`s, or *n* words that may wrap; the second `Kx`, whose
text start says how wide Word made the first column. Modes none and 15.

| Family | What Word did (first column, device px) |
| --- | --- |
| no `w:tcW`, *n* = 1-40 | **as wide as its content plus its two margins** (45 px), whatever `w:tblGrid` says: content 76 → column 121, 1,190 → 1,235 (the residual within ±1 px, the grid line's rounding) |
| `w:tcW` 2000, *n* = 1, 5, 10 | **417 px, the stated width**, the content 76-333 px |
| `w:tcW` 2000, *n* = 15-40 | **the content plus its margins** once that exceeds the width: 499 → 544, 1,213 → 1,258 |
| no `w:tcW`, 5 words | one line, 693 px: content plus margins |
| no `w:tcW`, 20 and 60 words | wrapped, the column filling what the second column's content and margins leave of the text column (1,768 px in mode none, 1,743 in 15; the 25 px between them not explained) |
| `w:tblW` 5000 pct, *n* = 1, 10, 30 | 993, 1,441 and 1,691 px (mode none; 968, 1,404, 1,648 in 15): the width shared by a rule these three cases do not settle |

Not settled, so not modelled: how Word shares a table's width stated in percent over
columns with no width of their own (a width in twips wider than its columns is
unprobed); the minimum width a wrapping column keeps beside another's content; and the
first column's width when several columns want more than the text column holds. When
those are measured, an autofit table whose stated widths are stale can be laid out; until
then it stops, and what the file states is used only where Word keeps it. A table whose
cells state their widths too -- in percent, or in twips -- is stage 7a.

### Stage 7a -- a width in percent: measured

A table "stretched to 100%" is written with `w:tblW` in percent and every cell's `w:tcW`
in percent; a filesamples document holds two, the layout stopping at the first.
`tools/make_pct_table_probe.py` / `read_pct_table_probe.py`: 52 two-row autofit tables of
short labelled cells, borders of `w:sz` 4, a `w:tblGrid` wrong on purpose (every column
1,000 twips), in three settings (none, 14, 15): `w:tblW` of 5000, 3500, 2500, 2250, 2000,
1750, 1500 and 1250 pct over two to four even shares; six unequal shares as Word writes
them; shares adding up to less and to more than the table; the grid Word would write,
twice it, and unequal grids of the table's width; margins 0, 300 and 300 / 0; centred,
right-aligned and indented; a 16-letter word, and 15 words, in one cell and not the other;
and `w:tcW` in `dxa` under a width in percent, equal and 1:3, with the same contents.

**The rules** (`table.percent_widths`):

* **The grid in the file does not matter**: the grid Word would write, twice it and
  unequal grids of the table's width are drawn alike.
* **100 percent** is, below mode 15, the text column **and** the first cell's left and
  the last cell's right offset (the margin, or the border's inner half where wider):
  9,380 twips of a 9,164-twip column with margins of 108, 9,174 with none, 9,764 with
  300; in mode 15 the column less the outer halves of the outer borders (9,154),
  whatever the margins. The table is `w:tblW` of it, and is placed as any table is
  (indent, `w:jc`; below mode 15 with no `w:tblInd`, its grid from the margin).
* **The columns share it in proportion to the cells' `w:tcW` in percent, whatever is in
  them** (1,000 / 2,000 pct under 5,000: a third and two thirds; a 16-letter word beside
  a label, or 15 words beside none, change nothing), each grid line at the floor of its
  cumulative share -- the grid Word writes for 818 / 849 / 849 / 850 / 851 / 783 pct of
  9,576 twips is 1,566, 1,626, 1,626, 1,628, 1,630, 1,500, exactly.
* **Shares in percent adding up to more than the table**: the columns before the last
  take what they say, the last what is left (1,000 / 4,500 pct: 1,876 and 7,504 twips,
  within the pixel Word drew them at).
* A share narrower than a word in its column is widened to hold it (782-twip shares of
  three and four columns labelled with a 671-twip word are drawn 888 twips wide; 763-twip
  ones with margins of 0 as they say): autofit, which stops the layout as a word wider
  than its column does.

**Cells in `dxa` under a width in percent are shared by their widths and their content**
(`table.dxa_shares`, `table.column_minimums`): where the cells' widths add up to no more
than the table, in proportion to them (1,000 / 3,000 twips under 3,500 pct: a quarter and
three quarters, whatever is in them); where they add up to more, **each gives up the
excess in proportion to its width less its narrowest content** -- the widest word of its
cells' paragraphs (and their indents) or the narrowest of a nested table in it (its
columns' and the outer halves of its outer borders), with the cell's margins. Equal
`w:tcW` of 4,788 twips under 3,500 pct (6,566 twips) with a 16-letter word (1,524 twips)
in one and a 671-twip label in the other are drawn 3,470 / 3,096 (the rule: 3,467 /
3,099; 3,398 / 3,010 in mode 15, the rule 3,398 / 3,009); 15 words beside none, whose
widest is the label, are drawn halves; with a nested table whose cells' words are 842
twips in one, 3,624 / 2,942 (the rule: 3,628 / 2,938) -- a nested table of 3,000 twips and
one of 1,000 over the same words weigh alike. Each within the pixel Word drew it at.

| Cells' text within half a pixel (232 lines each) | before | after |
| --- | --- | --- |
| `pct-table-none` / `-14` | 0 (the layout stopped at the first table) | **203** |
| `pct-table-15` | 0 | **201** |

What is left: 6 lines below mode 15 and 8 in mode 15, one cell's text a pixel from Word's
in a table (the rounding of a grid line at the pixel, not settled; ties at 1,277.5 px
among them), and the last family, which the layout stops at -- a column narrower than its
word, a row of a cell in percent beside one in `dxa`, `w:tcW` in percent under `w:tblW`
auto. Pinned in `tests/test_pct_table.py`.

### A table nested in a cell: measured

A table in a cell stopped the layout. `tools/make_nested_table_probe.py` /
`read_nested_table_probe.py` nest a two-by-two table in the first cell of a 4,500 /
4,500-twip table -- 30 cases, three settings (none, 14, 15): first in its cell or after
paragraphs, an empty paragraph after it, indents and alignments, margins and borders,
spacing around it and in it, widths in percent, taller and shorter than the other cell,
`w:trHeight`, the second of two rows, a merged cell, and tables in percent over cells in
`dxa`, outer and nested, one as a filesamples document nests it.

**The rules** (`table.nested_entry`, `ResolvedTable.nested`):

* **Across, a nested table is placed as a mode-15 table is in the column, in every
  mode**, the column being the outer cell's text area: its left border's outer edge at
  the cell's text start (grid lines at 324 / 637 / 949 px from a text start of 322.9,
  whatever the nested margins; with no borders its grid at the text start), `w:tblInd` from
  there (a negative one as none), centred in the text area, or its last grid line on the
  text's end; **100 percent is the text area less the outer halves of its outer
  borders**, below mode 15 too (4,274 of 4,284 twips).
* **Down, it is stacked in the cell as a paragraph's lines are** (one "line" as tall as
  the table): at the cell's content top, or right below the paragraph before it (whose
  space after counts), its rows as a table's, and the paragraph after it right below its
  bottom border, with its own space before. **An empty last paragraph after it takes no
  room** (the cell ends at the nested table's bottom border, with 240 twips of space
  before on that paragraph or not). The row is as tall as the cell's stack, or its
  `w:trHeight`.

| Lines within half a pixel (184 each) | before | after |
| --- | --- | --- |
| `nested-table-none` / `-14` | 0 (the layout stopped at the first table) | **178** |
| `nested-table-15` | 0 | **182** |

What is left: nested tables of 4000 pct centred, their text a pixel right of Word's below
mode 15 (6 lines), and in mode 15 one aligned right (2 lines). A floating table nested in
a cell is not measured and stops the layout. Pinned in `tests/test_nested_table.py`.

With F.18, stage 7a and this, a filesamples document Word sets on nine pages is laid out
past its floating table, its two tables in percent and its nested table, every glyph
it draws matched to one of Word's on the same page (3,374): to its fourth page, where a calendar table
stops the layout -- an autofit table of `dxa` cells (`w:tblW` auto, cells with their own
margins, `w:noWrap`, spans and `w:gridAfter`) whose columns Word resizes from their
content, narrowing the day columns to break "Sun" inside the word and widening the
spacer columns between them. Past it the document holds a drop cap (`w:framePr`), a
picture bullet and endnotes, none modelled.

Stage 7b shows that calendar was not resized at all: its Normal style indents every
first line by 432 twips, which Word's autofit does not count below mode 15, so every
column keeps the width its cells state (the grid drawn as the file says), and a day name
or a date that fits its cell but not its first line is broken inside the word -- the
empty spacer cells' marks pushed right by the same indent are what looked like wider
spacers. It is laid out now, and the document with it to its fifth page, every glyph it
draws matched to one of Word's on the same page (3,724, was 3,374; the calendar's
right-aligned text within a hundredth of a pixel of Word's, its last row a pixel low),
where the drop cap (`w:framePr`, "a frame") stops the layout. Past it are a picture
bullet and endnotes.

### Stage 7b -- autofit sized from its content: measured

An autofit table (`w:tblW` auto, or in `dxa`, or in percent over cells with no width)
whose content Word resizes stopped the layout. `tools/make_autofit_width_probe.py` /
`read_autofit_width_probe.py`: 38 autofit tables of labels and words of `x` (95.3 twips
each), borders of `w:sz` 4, a `w:tblGrid` wrong on purpose, three settings (none, 14,
15): first-line, hanging and left indents in a cell whose width holds the word but not
the word and the indent; cells with no width whose content fits (one to three columns,
two rows, two paragraphs, `w:noWrap`, margins 43 and 300, empty cells, indents); a word
wider than a width in `dxa`; content wider than the text column, in cells with no width
and in `dxa`; `w:tblW` in `dxa` and in percent over cells in `dxa` and with none; and,
last, a cell across two columns wider than they are.

**The rules** (`table.paragraph_extent`, `cell_extent`, `autofit_room`, `autofit_widths`):

* **A cell's narrowest content is its widest word, its widest content its longest line
  unbroken** (a break starts a new one), each with the indent in front of it, the right
  indent, and the cell's left and right offsets (the margin, or the border's inner half).
  An empty cell counts 6 twips of content: empty cells, and one stated 150 twips wide,
  are drawn 222 twips wide with margins of 108.
* **Below mode 15 a first-line indent counts for nothing, and a hanging one takes its
  width off the left indent**: every word counts from `left + min(firstLine, 0)`. A cell
  whose width holds a word but not the word after the first line's indent keeps its
  width, and the word is broken inside, as the line breaker breaks a word wider than a
  line: a 953-twip word after a hanging indent of 720 in a 1,440-twip cell is drawn
  `xxxxx` / `xxxxx`; a first-line indent of 1,440 there puts one character on the first
  line. With a left indent of 720 on every line the cell is widened to hold both (1,891
  twips). A cell with no width is as wide as its content without the first-line indent.
  **In mode 15 every line counts from where it starts**: the first word and first line
  from `left + firstLine`, the rest from `left` -- the hanging case is widened to 1,891,
  the first line of 1,440 to 2,040, and the cell with no width holds its first line.
* **A column is as wide as its cells' width in `dxa`, or, where none states one, as its
  widest content** (one to three columns, the wider of two rows, two paragraphs, margins
  43 and 300, indents: each within half a pixel), **and never narrower than its narrowest
  content**: a word needing 2,122 twips with its margins, in the middle of three cells of
  1,440, widens that column to 2,122 and the table with it. `w:noWrap` in a cell in `dxa` changes nothing. A column
  sized by its content is as wide as it to the twip above (at the twip below, a line of
  exactly that width wraps).
* **The room** is, below mode 15, the text column and the last cell's right offset (its
  grid from the margin, its text ending at the right margin: 9,272 twips of 9,164); in
  mode 15 (and nested in a cell) the column less the outer halves of the outer borders
  (9,154) -- as a width in percent's whole is, less its first cell's left offset.
* **Content wider than the room: the columns with no width give way first**, each keeping
  its narrowest content and taking a share of what is left **in proportion to its widest
  content less its narrowest**: 60 words beside 10 are drawn 7,430 twips (the rule
  7,430), beside 20 6,591 (6,590), beside a 2,598-twip word 6,548 (6,548), three columns
  6,365 / 2,136 (6,365 / 2,135); a column in `dxa` keeps its width while they can (3,000
  beside 60 words: 3,000). Where every column is in `dxa` they give way the same way, by
  their width less their narrowest: three of 4,000 are drawn 3,092 (3,091); 4,000 / 2,000
  / 3,000 with a 5,933-twip word in the second 1,848 / 5,933 / 1,491 (the rule 1,846 /
  5,933 / 1,493). Each grid line at the floor of its cumulative width.
* **A table width in `dxa` or in percent wider than the columns** widens the columns
  with no width in proportion to their widest content (6,000 twips over 952 and 1,534:
  2,300; the rule 2,298), or, where every column is in `dxa`, all of them in proportion
  to their widths (1,000 / 3,000 under 6,000: 1,500 / 4,500); a column in `dxa` beside
  columns with none keeps its width. Its grid lines at the nearest twip (2,297 draws the
  text a pixel left of Word's). **Narrower**, it is shared as the room is (2,000 / 2,000
  under 3,000: 1,500 each). In percent, the whole is stage 7a's (9,380 twips below mode
  15, 9,154 in 15): 5,000 and 2,500 pct over content of 952 and 1,534 are drawn 3,591 and
  1,796 (the rule 3,592 and 1,796).
* **A table whose cells state their widths in `dxa` and hold their content in the room
  keeps the grid it states** (stage 1, stage 7); where its widths cannot be measured (a
  tab, a drawing or a nested table in a cell) it still is laid out on them, and a word
  wider than its column stops it.

| Cells' text within half a pixel | before | after |
| --- | --- | --- |
| `autofit-width-none` / `-14` (169 lines each) | 13 (the layout stopped at the third table) | **161** |
| `autofit-width-15` (165) | 13 | **157** |

What is left: the last family, a cell across two columns wider than they are, which the
layout stops at: Word shares the extra nearly evenly but not exactly (753 / 763 twips over
two columns with no width, 532 / 546 over two of 1,000 in `dxa`), not settled. **A second
round** (seven more tables, the same three settings; widths from Word's grid lines, each
within 2.4 twips) rules out the simple rules and leaves it unsettled: borders of `w:sz`
24 and none draw the same grid as `w:sz` 4; two columns of equal content take 753 / 762
of the same excess as 952- and 1,153-twip columns (752 / 762), so the share does not
follow the content; a much wider cell (2,949 twips over) gives 1,424 / 1,521, so the
difference grows with the excess, the final widths keeping one ratio (0.4707) where the
content's is 0.452; across three columns 397 / 364 / 388 of 1,151, not in the columns'
order of content; across the second and third 786 / 755, the wider content taking more;
and a cell of 3,000 in `dxa` across 1,000 and 2,000 widens the first to 1,186 and the
second to 2,002. Still a stop. In mode 15,
where content wider than the room is shared, a grid line of two tables is drawn a pixel
from Word's (672 border pixels), their text exact. Not measured, and so stops: `w:gridBefore`
/ `w:gridAfter` or a column only spanned in a table sized by its content, cells in
percent in a table not in percent, a column in `dxa` giving way beside columns with no
width at their narrowest, a cell with `w:noWrap` and no width giving way, a table width
in `dxa` wider than the room, and a table narrower than its narrowest content (since
measured: stage 7c). Pinned in `tests/test_autofit_width.py`.

Production feedback (2026-10) met this stop again: real templates stop at their first such
table. The layout reports the stop and never guesses, which is right. But until the
layout had a coverage summary (5.20), a caller saw only a short layout, with no sign that
anything had been skipped. The coverage summary now says how many blocks were skipped and
names this table (`coverage.stop`: `table`, with its path). The rule is still not
settled, so the layout still stops here.

### Stage 7c -- autofit narrower than its content: measured

A table whose columns' narrowest content -- each its widest word with the cell's margins
-- adds up to more than the room stopped the layout ("a table narrower than its content";
a review table of nine narrow columns headed by two-word labels on a portrait page is
one). `tools/make_autofit_over_probe.py` / `read_autofit_over_probe.py`: eleven tables,
three settings (none, 14, 15): three, five and nine columns in `dxa`, three with no width
(one word each, and long words beside many words), `dxa` beside no width, one column whose
word is wider than the room (with no width and in `dxa`), and `w:tblW` in `dxa` and in
percent over such content.

**The rule** (`table.autofit_widths`): **the table keeps to the room** (`autofit_room`:
9,272 twips below mode 15, 9,154 in it), whatever width it states in `dxa` (6,000: the
room), and **each column keeps its margins and takes a share of the rest in proportion
to its narrowest content** -- its widest word, not its widest content: three columns of
equal words share it evenly whether or not one also holds thirty more words, and a column
in `dxa` beside columns with no width the same. The words wider than their column are
broken inside, as the line breaker breaks a word wider than a line. Nine columns of
1,334 ... 1,239-twip words: Word 1,157 / 955 / 1,022 / 955 / 888 / 1,022 / 960 / 1,224 /
1,090, the rule 1,157 / 955 / 1,023 / 955 / 888 / 1,023 / 955 / 1,225 / 1,090. A grid line
is at the nearest twip below mode 15 and at the floor in it (five columns of 1,854.4: the
third's text a pixel off Word's either way round).

| Cells' text within half a pixel | before | after |
| --- | --- | --- |
| `autofit-over-none` / `-14` (62 lines each) | 0 (the layout stopped at the first table) | **59** |
| `autofit-over-15` (63) | 0 | **63** |

What is left: below mode 15 a width in percent narrower than its content is drawn **past
the margin**, about as wide as its narrowest content (5,000 pct over three columns of
3,552: 3,552 / 3,552 / 3,547), its words still broken by a rule not settled -- a stop (the
three lines missing above). In mode 15 it keeps to the room like the rest. In mode 15 a
grid line of six tables is drawn a pixel from Word's (1,288 border pixels), their text
exact. Pinned in `tests/test_autofit_over.py`.

**Open in the tools, not the model:** the line matcher behind the baseline and page-top
scorers (`tools/baselines.py`, `tools/pages.py`) takes each baseline Word drew as one
line and does not split a table row's baseline into its cells, so it cannot score cell
paragraphs. Run live on the committed documents, it counts `sample-5pages`' page top
(inside the split fifth row) as a miss and some cell paragraphs' line counts as
different, while the glyph-position check puts every glyph of those pages on Word's page
at Word's position. `tests/test_pages.py` keeps its recorded advances, which predate
tables (lacking the cells' faces, its paginator still stops at the first table), until
the scorers are table-aware; the tables are held by `tests/test_tables.py`,
`tests/test_render.py` and the VRT snapshots.

## Headers, footers and fields — measured (Phase 6, Headers and footers and Fields rows)

**Done: headers and footers are drawn, and `PAGE`, `NUMPAGES` and `SECTIONPAGES` are
computed from the layout's own pages.** Phase 4.4 had measured the room a story takes;
Phase 5 warned `headers-footers-not-drawn` on every document that has one. Now each page
draws the story Word shows on it, laid out by the machinery the body uses -- the style
cascade, the line breaker, the vertical model, tables, pictures, borders and paint
order -- and every field's instruction is gone from the text. Three probes, each a
committed `make_*` / `read_*` pair exported by Word 16.106 through `tools/oracle.py`,
recorded in `tests/fixtures/story-observations.json` and held offline by
`tests/test_stories.py`; the glyph-position check (5.3) counts a story's glyphs apart
(`Score.story_*`), and the VRT snapshots and the visibility check cover three of the
probe documents.

### H.1 Where a story's glyphs go (`make_story_probe.py`)

34 sections, each its own case with its own header and footer part (Word 16's `Header`
and `Footer` styles unless said), in four settings (no `settings.xml`, modes 12, 14,
15): distances 0, 360, 725, 731 (not whole pixels), 1000, 1417 and 1800 / 1600; Times New
Roman 14.5, Cambria 9.5, Arial 20, Georgia 7.5; `auto` 276 and 360, `exact` 300, `atLeast`
400, space 120 before / 240 after; two paragraphs, a wrapped paragraph, a `w:br`; tabs on
Word's stops, centred, right-aligned; a bottom border under the header and a top border
over the footer; a table in each; an inline picture in each; five lines (past both
margins); a gutter; indents of -720 into both margins; A4 and A5; `w:vAlign` center and
bottom; the stories in `Normal`.

**The rules** (`layout._draw_story`, `paginate.story_document`, `paginate.stack_height`):

* **A story is laid out as a document of its own**: its blocks the body of one section
  with the page's size and margins and one text column, flowed and stacked exactly as a
  page's body is from its top -- the first paragraph starts a section, so it keeps its
  space before; lines round in their line boxes as the body's do (`baseline_in_box`), in
  every setting; the text width is the page's between the margins.
* **A header's stack starts at `w:pgMar/@w:header`** below the page's top edge.
* **A footer's stack ends at `@w:footer`** above the page's foot, its last paragraph's
  space after included: it grows upward -- its top is the foot less the footer distance
  less the stack's height ([`stack_height`], the same height the paginator reserves:
  one function for both).
* **Nothing else moves a story**: the section's `w:vAlign` centres or bottom-aligns the
  body, not the header or footer (both exact on those pages).
* A stated `w:header="0"` is 0: the parser read any margin of 0 as its default (`or
  720`), which put the `distance 0` case's header and footer 150 px from their edges.
  Every `w:pgMar` value is now read as stated (no committed or local document moved).

| Setting | Glyphs matched of drawn (Word drew) | Header and footer glyphs exact | Rectangles |
| --- | --- | --- | --- |
| none, 12, 14, 15 (each) | 3,422 of 3,422 (3,422) | **1,786 / 1,788**; *1,788 / 1,788 with H.5* | 11,772 px, 0 / 0 |

The two story glyphs off, and one body object start, are the A5 case's: its left margin
is 1134 twips (236.25 px) and Word starts every line of the page -- body, header,
footer -- at 236.0 (H.5). The body glyphs of the two `vAlign` pages are Word's centred
and bottom-aligned body, which is not laid out (warned `vertical-alignment-not-drawn`);
their stories are exact. The Phase 4.4 header probe, drawn now, puts 274 / 274 of its
matched story glyphs where Word does (a negative top margin's header shares a baseline
with the body, which Quartz merges: two glyphs unmatched, as 4.4 recorded).

| Refuted (mode 15; story glyphs exact of 1,788) | Exact |
| --- | --- |
| **the model** | **1,786** |
| a footer stacked down from its distance above the foot (not grown upward) | 537 (116 glyphs not even matched) |
| a footer's height without its last paragraph's space after | 1,760 (the `space 120/240` footer) |

### H.2 Which story, on which page, numbered what (`make_story_select_probe.py`)

Six documents, 20 sections, 71 pages; every story names itself and every footer carries
`PAGE`, `NUMPAGES` and `SECTIONPAGES`: inheritance, `w:titlePg` switched on and off,
`first` and `even` stated where unused, `w:evenAndOddHeaders` on and off, `oddPage` /
`evenPage` breaks after each parity, `w:pgNumType/@w:start` restarts on each kind of
break, continuous sections starting mid-page. **1,752 / 1,752 glyphs on Word's page at
Word's position, every blank page blank.** The rules (`paginate.story_references`,
`story_kind`, `section_start_number`):

* **Inheritance is per type**: a section stating no `default`, `first` or `even` takes
  the last one stated before it -- also one the stating section did not use (a `first`
  stated under no `w:titlePg` shows on the next `w:titlePg` section's first page).
  `w:titlePg` itself is not inherited.
* A page shows **`first`** when its section starts on it under `w:titlePg`; else
  **`even`** when its **number** is even (restarts counted) under `w:evenAndOddHeaders`;
  else **`default`**. A kind no section has stated is drawn **empty** -- not the
  `default` (`undefined`: a `w:titlePg` first page and every even page blank).
* **A page's section is its first line's**: a continuous section that starts mid-page
  does not change that page's stories, and its "first page" never comes (`blank`: the
  page after shows its `default`).
* **Numbers**: `@w:start`, or the page before's plus one. An `oddPage` / `evenPage`
  section wants its **number**'s parity, not its place's: a restarted number of the
  wrong parity is raised by one with **no blank page** (`start 2` on an `oddPage`
  section is drawn 3; `start 7` on an `evenPage` one, 8); an unrestarted one gets a blank
  page first, which takes the number. **Under `w:evenAndOddHeaders` a page's number and
  its place keep one parity**: a restart that would break it gets a blank page first
  (`start 1` on the eighth page puts page 1 on the ninth); without the setting it does
  not (`start 5` on the twelfth page).
* **Blank pages Word inserts show nothing** -- no header, no footer -- in both settings.
* `NUMPAGES` counts every page, blank ones included; `SECTIONPAGES` the pages the
  section has content on (a blank page before it is not its; one shared with two
  continuous sections counts 1 for each).

| Refuted (story glyphs exact of 1,380 Word drew) | Exact |
| --- | --- |
| **the model** | **1,380** |
| blank pages by the page's place (Phase 4.5's rule, which never read a restart) | 1,108 |
| no blank page for a restart under `w:evenAndOddHeaders` | 1,166 |
| a restarted wrong-parity number getting a blank page too | 1,228 |
| an undefined `first` / `even` falling back to `default` | 1,380, and 96 glyphs drawn that Word does not draw |

Phase 4.5's section probe (no restarts, so place and number agree) is unchanged.

### H.3 Fields (`make_field_probe.py`)

**The parser no longer keeps an instruction as text** (Phase 3's recorded gap closed):
`w:fldChar` `begin` / `separate` / `end` and `w:instrText` are followed across runs,
paragraphs and nesting; an instruction is never drawn or measured; a field's cached
result -- between `separate` and `end`, or a `w:fldSimple`'s runs -- is ordinary runs,
drawn as Word cached it; **`PAGE`, `NUMPAGES` and `SECTIONPAGES` are one run each**
(`model.Field`), whose text the layout computes page by page (`docx2svg.fields`). A field
in the body no longer stops the layout (`layout-stopped:field` is gone).

Two documents. **`computed`** -- 14 sections of page-number formats and 15 switches,
header fields whose runs differ in format, the fields Word draws as cached, a
300-character instruction, a table of contents spanning paragraphs, and body paragraphs
whose `PAGE` / `NUMPAGES` (cached `1`, computed `1000000` / `34`) end a line: **6,437 /
6,437 glyphs at Word's position, 3,673 / 3,673 of them the stories'.** **`word`** --
what the file cannot say (below): pinned, 115 glyphs off and warned.

* **Formats** (`w:pgNumType/@w:fmt`, `PAGE` only): decimal; upper/lower Roman (4000
  `MMMM`, 4001 `MMMMI`); upper/lower letter (27 `AA`, 52 `zz`, 53 `aaa`, 255 ten `u`s);
  `numberInDash` `- 9 -`; `decimalZero` `07`; `hex` `FF`; `chicago` `*`, `†`. Roman and
  letters of 0 are drawn as a space. `NUMPAGES` and `SECTIONPAGES` are arabic in any
  section's format.
* **Switches** override the section's format: `\* roman` / `Roman` (the switch's first
  letter's case is the result's), `alphabetic` / `ALPHABETIC`, `Arabic`, `ArabicDash`,
  `Hex` (upper case), `\# "000"`; then `Upper`, `Lower`, `FirstCap`. `MERGEFORMAT` and
  `CHARFORMAT` do not change the text.
* **The format a computed result is drawn in** is the **instruction's first run's**
  (`begin` plain, instruction red, cached result blue: red), with `\* MERGEFORMAT` the
  **cached result's first run's**, with `CHARFORMAT` the instruction's; a `w:fldSimple`'s
  is the paragraph's (its bold result run is drawn plain).
* **The breaker measures the computed value**: a body `PAGE` or `NUMPAGES` changes what
  it is measured in, so the body is laid out again with the values found until they are
  the ones it was laid out with (`layout._Placer.run`, at most four passes; a paragraph
  that spans pages takes the page of the field's own line). Word's breaks in the
  million-numbered section need it: with the cached `1` measured, lines break elsewhere
  (`test_the_breaker_measures_the_computed_page_number`).
* **Drawn as cached, and so is Word's**: `AUTHOR`, `FILENAME`, `TITLE`, `NUMWORDS`, `=`,
  `HYPERLINK`, `DOCPROPERTY`, `QUOTE`, `TOC`.
* **Computed again by Word on export**: `DATE`, `TIME` (today's date and time),
  `REF`, `PAGEREF`, `SEQ`, `IF` (also an `IF` whose instruction holds a `PAGE`). Drawn as
  cached, warned `field-cached:<name>` -- in a real document the cached result is usually
  the current one; `DATE` and `TIME` cannot be, by construction.
* **The application's language, not the document's**: `ordinal`, `cardinalText`,
  `ordinalText` and the switches `Ordinal`, `CardText`, `OrdText` are drawn by this
  Word as `1e`, `één`, `eerste` in an `en-GB` document -- the host's language. The file
  cannot say it, so they are drawn as cached, warned `field-not-computed:PAGE`.
* A `NUMPAGES` or `SECTIONPAGES` past where the layout stops is not known: drawn as
  cached, warned.

Seen and not measured further: a footer paragraph holding a `\* alphabetic` result of
154 letters (3,999), or any switch on 1,000,000, is not drawn by Word at all -- the probe
keeps those sections' footers short.

### H.4 Warnings

`headers-footers-not-drawn` is gone. A story is drawn in full or its gap is named:
`story-drawing-not-drawn` for a floating drawing text does not wrap around (a logo
behind or in front of the text: its mark is taken out, the text is drawn where Word
draws it, the drawing is not; *gone since* Floating drawings — measured, *F.6: drawn*); `story-stopped:<reason>` where a story holds what the
body would stop at (a floating drawing text wraps around, an unmeasurable paragraph):
drawn up to it. `field-cached:<name>`, `field-not-computed:<name>`,
`vertical-alignment-not-drawn`, and `field-layout-unsettled` if the body's page numbers
do not settle in four passes.

### H.5 Found by the story probe: a left margin off the pixel grid

The A5 case's left margin is 1134 twips, 236.25 px; Word starts every line of that page
-- body, header and footer -- at 236.0, where the model put 236.25. Not a story rule:
the body was off by the same. **Measured on a probe of its own and adopted (an extra
rule)**: `make_margin_probe.py` / `read_margin_probe.py`, `tests/test_margin.py` -- left
and right margins of 1440-1449 twips (every fraction of a pixel a twip reaches), gutters
of 1, 3 and 7 twips, a 5-twip indent on such a margin, and the A5 page; a header, a
footer and left-, centre- and right-aligned body lines; no `settings.xml` and mode 15.

**The text column starts at the left margin plus gutter rounded to a whole device pixel,
held in whole layout units** (`layout._Placer._new_page`: 301 px is 295,895.04 units,
drawn at 295,895 -- 300.9999 px in Word's PDF); indents, tabs and alignment are measured
from there as before, and the column is as wide as before (its right edge moves with
it). Integer twips never put the margin on a half pixel, so the rounding's tie is not
reached. A table's grid lines already hung from the rounded edge (5.14) and its cell text
from whole pixels; a floating drawing's anchor offset is left from the exact margin (not
measured).

| Rule (object starts exact of 85 a document) | none | 15 |
| --- | --- | --- |
| before: the exact margin | 5 | 5 |
| the margin rounded to a whole pixel | 82 | 79 |
| **rounded, held in whole layout units** | **85** | **82** |

The three left in mode 15 are right-aligned footer lines a layout unit (0.001 px) off,
by a rule not found (the same lines are exact with no `settings.xml`, and so are the
body's right-aligned lines). The story probe's A5 page is now exact: **1,788 / 1,788**
story glyphs in every setting, and no object start off. Checked line by line against the
commit before over the same 279 documents, only pages whose left margin is off the pixel
grid moved: **Tables' line probe's `m1442` and `m1438` documents** (four; their 289 text
lines a document, 0 / 289 object starts exact before, **289 / 289** now -- its table
lines were already exact), and the local corpora,
whose pages draw no line before their stop and whose stop band moved with the
column. No committed document has such a margin.

### H.6 The real documents

**None of the real documents here has a header or footer glyph.** `samplelib`, `word.to`
and `filesamples` hold no header or footer part at all -- including `filesamples/sample4`,
whose 175 pages carry **no page numbers** (no header or footer part, no field): the
premise that it would exercise them is not borne out -- and none has a `PAGE` field. The
local corpora hold no header or footer text. So on every real document: 0 story glyphs
to draw, and **no line, baseline, page top or glyph moved** -- checked line by line
against the previous commit over 279 documents -- every committed, `filesamples`,
variant and local document and 251 generated probe documents, 428,177 lines (every
line's baseline, every glyph's x, face and size, every rule, picture, page geometry and
stop): only warnings changed
(`headers-footers-not-drawn` gone; `sample1` and its variants now warn
`field-cached:PAGEREF` for their table of contents). The VRT snapshots of the eleven
committed documents are byte-identical, so their fidelity is unchanged by construction
(`tests/fidelity-baselines.json` untouched).

Fidelity (the `svg` truth, device / exact) of the probe documents, the previous commit
against this one -- their pages are sparse, so what they show is the story and the page
count, not glyph edges:

| Document | Before | After |
| --- | --- | --- |
| `story-15` (34 pages) | 0.9725 / 0.9725, loss 4,953 | **0.9992 / 0.9991**, loss 156 |
| `select-even` | 18 pages made of Word's 20 (blank pages wrong); 1.0 over them, under 1.5% ink | **20 of 20**, 1.0 |
| `fields-computed` | 1 page of 28 made (stopped at the first field): 0.9542 | **28 of 28, 0.9995 / 0.9993**, loss 223 |

The pages were rendered beside Word's and looked at (the header and footer bands of the
bordered, table and picture cases, the wrapped story, the field pages): no difference
beyond anti-aliasing.

### H.7 Open

* **A floating drawing in a story** (the common header logo) is not drawn: floating
  drawings are the next task. One text wraps around makes the story stop there.
  *Drawn since* Floating drawings — measured *(F.6), where text does not wrap around it.*
* **`PAGEREF`, `REF`, `SEQ`, `IF`, `DATE`, `TIME`** are drawn as cached where Word
  recomputes them; `PAGEREF` is computable from this layout (a table of contents' page
  numbers) once bookmarks are read. *`PAGEREF`, `REF` and `SEQ` computed since H.8.*
* **Number words** need the host's language; not computed.
* A story whose height depends on its page's field values (a page number that wraps
  its line) reserves the height of its cached values; not measured.
* A `w:pgNumType/@w:start` on a continuous section that starts mid-page; `w:vAlign`
  of the body; footnotes under a story that pushes the body; a story taller than half
  the page.
* **A header's or footer's VML drawing** (`w:pict`: a `v:rect` band, a `v:line` rule, a
  watermark's `v:shape` with a `v:textpath`, a VML picture) is not laid out. The story is
  drawn up to its paragraph (`story-stopped:drawing`, and in `coverage.story_stops`); the
  rest of that header is not drawn. The body is unaffected: its pages and positions are
  laid out in full. Templates use these often (production feedback, 2026-10). They are
  not a small gap: VML's geometry, its positioning (`mso-position-*` against the page,
  margin or text), its wrap (`w10:wrap`) and its paint order against the story's text and
  the body's drawings each need probes. Next step: a probe of absolutely positioned VML
  shapes in a header with no wrap, which should leave the header's text where it is, so
  the text after them could be laid out before their outlines are drawn.

### H.8 `PAGEREF`, `REF` and `SEQ` computed (`make_xref_field_probe.py`)

docx-agent writes these fields' results from docx2svg's layout (its E4, 143 of 143 cases
Word's) and proposed docx2svg compute them as it computes `PAGE`. Word computes all three
again when it exports (H.3), so a cached result says nothing; every field of the probe
caches one Word could not have computed (`999`), and its runs differ in format -- the
`begin` / `separate` / `end` runs Calibri 11, the instruction Arial 9, the cached result
Georgia 12 italic, a bookmark's text Times New Roman 14 bold -- so the export shows whose
format the result takes. Bookmarks: a run on page 1; a paragraph on page 3; one on the
second page of a `lowerRoman` section numbered from 10; one over two paragraphs; a caption
(`Figure`, a `SEQ`, text); an empty one; one in a table cell; one starting in the last line
of a paragraph that runs from page 4 to page 5; a run in a Cambria heading style; a run of
a Courier New character style; and one never defined. 15 `REF`, 11 `PAGEREF` and 15 `SEQ`
cases and a table of contents whose entries hold `PAGEREF`s; no `settings.xml` and mode 15,
which agree. **The rules** (`fields.cross_references`, `fields.sequence_values`,
`fields.reference_runs`, `layout._Placer._bookmark_pages`):

* **`PAGEREF name`** is the page the bookmark's *start* is on -- the line it stands in: the
  paragraph from page 4 to 5 gives 5 -- shown as `PAGE` would show it *there*: **in the
  bookmark's section's format** (`xi`, from a decimal section), switches applied (`\*
  Arabic` 11, `\# "00"` 03, `\* roman` of page 4 `iv`); formatted as `PAGE` is (the
  instruction's first run; `\* MERGEFORMAT` the cached result's). Inside a table of
  contents' result too: Word recomputes the entries' page numbers on export, the TOC's
  own text as cached. `fields.field_text` takes `PAGEREF`.
* **`SEQ id`** counts the body's `SEQ` fields of its identifier in document order, table
  cells in their place, **the identifier's case ignored** (`figure` continues `Figure`):
  nothing or `\n` adds one, `\c` repeats, `\r 5` sets 5, `\h` adds one and draws nothing,
  `\*` formats (`vii`, `h`). A `w:fldSimple`'s result is in the paragraph's format, a
  complex one's as `PAGE`'s.
* **`REF name`** is the bookmark's text **in its own runs' formatting**: direct and
  character-style formatting carried (Times New Roman 14 bold; the character style's
  Courier New), **the source paragraph's style not** (a plain run of a Cambria heading is
  drawn in the destination's Calibri); a `SEQ` in it as counted (`Figure 1 caption
  text`); an empty bookmark nothing. `\* CHARFORMAT`: all of it in the instruction's
  first run's format. **`\* MERGEFORMAT`: the *k*-th word in the *k*-th word's format of
  the cached result** -- `Alpha` in Georgia italic, the space and `target` after it in the
  source's, the cached `999` having one word. `Upper`, `Lower`, `FirstCap`, `Caps` apply;
  `\h` changes nothing drawn.
* **Not computed, drawn as cached and warned `field-not-computed:<name>`**: a missing
  bookmark (Word draws `Fout! Verwijzingsbron niet gevonden.` -- its application's
  language), `PAGEREF \p` (`above`), `REF` of a bookmark over two paragraphs (Word draws
  both, its paragraph mark and all), `SEQ \s` (a restart at heading levels: here, with no
  heading, Word drew 1), a `REF` switch other than those above, and a `REF` whose bookmark
  holds a tab, an object or a page-number field. In a header, a footer, a note or a text
  box all three stay as cached (`field-cached:<name>`), as before -- not measured there.

| Glyphs (Word drew 2,147 in each setting) | matched | x | baseline | face |
| --- | --- | --- | --- | --- |
| before: drawn as cached | 1,853 | 1,809 | 1,656 | 1,852 |
| `REF` in its instruction's format, as `PAGE` is | 2,033 | 1,901 | 1,825 | 1,907 |
| `PAGEREF` in the format of the field's own section | 2,029 | 2,022 | 2,028 | 2,029 |
| **the model** | **2,033** | **2,028** | **2,032** | **2,033** |

The five glyphs off and the 114 not drawn are the five cases not computed (four closing
brackets after a cached result of another width, and the line after the two-paragraph
`REF`, which Word draws a line lower). H.3's `fields-word` document draws its `REF`,
`PAGEREF` and `SEQ` as Word does now (glyphs off Word's x 115 -> 110; only `DATE`, `TIME`
and `IF` left). A filesamples document's table of contents holds 17 `PAGEREF`s: each is
computed from this layout to the number Word cached. No committed document has one of
these fields: fidelity as recorded.

---

## Floating drawings — measured (Phase 6, Floating objects and `w:drawing` rows)

A floating drawing (`wp:anchor`) stopped the layout (`layout-stopped:drawing`): every
local template stopped at its first paragraph, which anchors one, and a header's logo
was warned of, not drawn. The work is staged from the least layout impact to the most;
each stage has its own probes, each a committed `make_*` / `read_*` pair exported by
Word 16.106 through `tools/oracle.py`, and its own commit.

### F.1 Stage 1 — anchors that do not move text: where they go

`make_anchor_probe.py` / `read_anchor_probe.py`, `tests/test_anchors.py`
(`tests/fixtures/anchor-observations.json`). 138 pages, each a heading line, the
paragraph that anchors a picture and a line after it; the picture is a 16 x 16 PNG of
one colour stretched, so Word's image box is the drawing's box, and its paint order in
the PDF (before the page's first glyph or after its last) is its layer. The section is
off the pixel grid (A4, margins 1500 / 1300 / 1600 / 1442 twips, header 700, footer
650); extents step by 3,701 EMU across and 37,003 down. Families: every `positionH` and
`positionV` `relativeFrom` with an offset and with every alignment (the two axes paired,
read apart); the anchor after space before, second on its page, on its paragraph's
second line, on a centred and a right-aligned line, first on its page (the templates'
title); after a letter, a tab, a break, beside larger text; `simplePos`; effect extents
with offsets and alignments; past every page edge; stacking (`relativeHeight`,
`behindDoc`, a shaded paragraph, a shaded cell); `allowOverlap="0"`, `locked`; a second
section with other header and footer distances; anchors in table cells. Two documents:
no `settings.xml` and mode 15.

**The rules** (`docx2svg.floating`, `layout._anchors`):

* **Twips, truncated.** An offset is truncated to whole twips towards zero (400,003 EMU
  is 629 twips, -250,001 is -393); the extent an alignment uses is truncated to whole
  twips on each axis; a centred drawing may sit on a half twip. A picture is drawn at
  `layout.picture_size` (Phase 5.14's inline rule, the same to 0.001 px), a shape or
  group at its extent in whole twips on each axis.
* **Effect extent.** An alignment places the box of the extent *and* its effect extent
  (`left`/`top`: the effect's left/top edge on the frame's; `right`/`bottom`: its
  right/bottom edge; `center`: its middle); an offset places the drawing itself.
* **Horizontal frames.** `page`; `margin` and `column` from the **exact** left margin
  (300.417 px, where the text column is rounded to 300); `leftMargin`, `rightMargin`;
  `insideMargin` / `outsideMargin` by the page's parity -- in mode 15 inside is the left
  margin on an odd page, below mode 15 the right one; `character` -- the pen position of
  the piece before the anchor's mark on its line in document order (a character, a
  space, the start of a tab), or the line's start, measured from the exact margin: a
  frame of no width (`right` puts the drawing's right edge on it).
* **Vertical frames.** `page`; `margin`; `topMargin`, `bottomMargin`; `insideMargin` is
  the top margin on an odd page in every mode, and `outsideMargin` the bottom one in
  mode 15 -- **below mode 15 outside is inside** (the top margin on an odd page).
  `paragraph`: the paragraph's top is **where the paragraph before it ends after its own
  space after** (the rest of this paragraph's space before is inside it: 137 twips
  before after 120 after puts it 120 below), or the page's text top for the page's first
  paragraph whatever its space before; **any alignment against it is its top**. `line`:
  the anchor's line's box, whose top is the paragraph's top on its first line and the
  line's top on the others; in mode 15 it reaches down to the line's pitch (`center`
  puts the drawing's middle between the two, `bottom` its bottom on the pitch's), and
  **below mode 15 it has no height** (`center` and `bottom` put the drawing's middle and
  bottom on its top).
* **Inside / outside against the page or margin**: inside is left and top on an odd
  page, in both settings. **Against the page vertically, the top is half the header
  distance below the page's top edge and the bottom half the footer distance above its
  foot** (72.917 px with a 700-twip header, 93.75 with 900; the foot 325 twips up with a
  650-twip footer) -- whatever the drawing's size.
* **`simplePos`**: the point from the page's corner in whole twips -- vertically through
  the paragraph, its distance from the paragraph's top rounded to whole twips (1,111.445
  twips drawn 1,111: 599.907 px for a stated 600).
* **The text does not move.** The anchor's mark is no piece of the line breaker's: every
  line breaks and stands as without it -- 23,507 / 23,507 body glyphs at Word's pen x
  and baseline in both documents, the anchor's own lines included.

| | Pictures where Word drew them, in its order and layer | Body glyphs exact |
| --- | --- | --- |
| `anchor-none` | **143 / 155** | 23,507 / 23,507 |
| `anchor-15` | **143 / 155** | 23,507 / 23,507 |

The twelve: `relativeHeight` values of 1-3 (below the range Word writes, which starts at
251,658,240) are not stacked by value (two pages, five pictures: heights 3, 1, 2 drawn
2, 3, 1); `allowOverlap="0"` makes Word move the second of two overlapping pictures
clear of the first, to its right edge (0.42 px past it in mode 15, 0.21 below it) --
drawn where positioned and warned `drawing-overlap-not-modelled`; and the five pages
whose anchor is in a table cell, where the layout stops (a floating drawing in a cell is
still `layout-stopped:table`: with `layoutInCell` Word positions it against the cell --
in mode 15 whatever `layoutInCell` says -- which is not modelled). *Modelled since, F.16.*

**Refuted on the way**: the paragraph's top as its first line's top (the space before
counted: 25 px low after 120 twips after); the character as the next glyph's pen x
(9.94 px right: the space before the anchor); a picture's position from its drawn size
(0.09 px off at the right margin: the truncated extent is used); alignment ignoring the
effect extent (8.33 px, the effect's asymmetry).

**Found on the way, an export artifact**: Word's PDF export drops a *one-pixel* picture
that the page's edge crops (a 1 x 1 PNG past the bottom by 36 twips of a 1,102-twip
picture), and draws a larger picture cropped the same way; the probes' pictures are
16 x 16 for that reason. A picture wholly off the page is not drawn; a shape is.

### F.2 Stacking — paint order

Read off the PDF's content stream (`page.get_bboxlog`): **`behindDoc` drawings are painted
first on the page -- under everything, a paragraph's shading and a table cell's shading
included**; the others after everything, the text, borders and inline pictures; in each
layer by ascending `relativeHeight`, ties in the order met
(`layout.Page.layers`, `svg.render_page`). `locked` changes nothing.

**The visibility check** (5.12) counts a glyph that shows no ink under a drawing painted
after it in front of the text as **covered**, apart from hidden: Word paints that
drawing over it too (`tools/visibility.py`, `Glyph.covered`). A glyph under a drawing
behind the text must still show. `anchor-15` and `drawing-15` are in the default suite's
check: 0 hidden; 3,211 and 22 covered.

### F.3 What a drawing shows: shapes, colours, groups, text boxes

`make_drawing_probe.py` (read by `read_anchor_probe.py`): 62 pages -- rectangles filled in
`srgbClr` and in `schemeClr` with `lumMod` / `lumOff`, outlines, the `line` preset,
`roundRect`, `ellipse`, `triangle`, `rtTriangle` (flipped), custom geometries (lines, a
cubic, an arc), eleven text boxes, six groups, a rotated rectangle, a gradient, a shape
behind the text, and 144 text boxes swept down by single twips; half the drawings in
`mc:AlternateContent` as Word writes them (the parser takes the `mc:Choice` holding a
`w:drawing`: `parse.drawing.drawing_element`).

* **Shapes** are drawn at their extent in whole twips on each axis: every one of the 22
  filled rectangles Word drew (fills, groups' members, filled text boxes, the one behind
  the text) is the model's to 0.03 px, colour included.
* **Colours**: a theme colour through `w:clrSchemeMapping` (the default map where the
  settings state none), `lumMod` / `lumOff` on the HSL lightness, each channel rounded to
  the nearest level **a half down** (black at `lumMod 50000 lumOff 50000`, 127.5, is
  `7F7F7F`); every probe colour exact. (Every other transform since measured and drawn,
  by `ooxml_common.drawingml.color.WORD`: F.12.)
* **Groups**: the group's box is its extent in whole twips; each member's `a:xfrm` in the
  group's child coordinates, mapped from `a:chOff` / `a:chExt` onto that box (a member
  of 700,001 EMU in a 3,149-twip group of 1,700,000 child units is 270.14 px, Word's
  270.13), untruncated; flips and rotation about each box's centre.
* **Geometry**: the guides and the preset table are `ooxml-common`'s
  (`drawingml.guides`, `drawingml.preset_specs`); `docx2svg.drawing` only turns their
  commands into path data. That table leaves out the presets `pptx2svg` draws by hand,
  among them `rect`, `roundRect`, `ellipse`, `line` and `rtTriangle`, which the
  documents here use: they are written in `drawing.EXTRA_PRESETS` from ECMA-376's
  definitions, in the table's form, until the table carries them (F.7). The rendered
  pages were looked at beside Word's: the presets, flips, the arc, the cubic, the group
  of flipped members and the rotated rectangle are Word's to the eye. (Since F.12 the
  table is complete in `ooxml-common` and `EXTRA_PRESETS` is gone.)
* **Not drawn to the pixel, and said so**: a gradient, pattern or picture fill (a
  placeholder of the extent, `drawing-not-drawn:gradFill`), effects, a theme style
  reference with nothing stated, dashed outlines (drawn solid,
  `drawing-outline-approximate`), a chart or anything else in `graphicData` (a
  placeholder). (Since F.12 all of these but a picture fill and a chart are drawn; a chart
  since F.20, and SmartArt since F.21.)

**Text boxes** (`layout._text_box`): the content (`w:txbxContent`) is laid out as a
document of its own by the body's machinery at the inner width -- the extent less
`wps:bodyPr`'s insets (default 0.1 in left and right, 0.05 in top and bottom) -- and
stacked from the inner box's top exactly, as a page from its top margin; `anchor="ctr"`
and `"b"` centre it or put it at the bottom **when it fits, and put it at the top when
it does not** (`taller than its box`). Its lines start at the inner box's left edge
**rounded to a whole device pixel, a half down**, held in the **nearest** layout unit
(567 px is drawn at 567.0003): 144 / 144 swept boxes' first glyphs, where half up puts
the tie at 567.5 a pixel right of Word's.

| Text box glyphs and the probe's body (7,180) | pen x | baseline |
| --- | --- | --- |
| `drawing-none` | 7,169 | 7,121 |
| `drawing-15` | 7,169 | 6,661 |

*(Since F.13 the insets count from the geometry's text rectangle, inside half the
outline's width; the sweep's boxes have neither, and keep every score.)*

**Not settled -- where a text box's line stands vertically** (*settled since, F.14: the
box's last line keeps its room below it*). In mode 15 a text box's
last line (its paragraph mark's) of 14 pt and larger stands a pixel lower than the
model's, and a two-line paragraph's first line does not: 519 of the 7,180 (the sweep's
16 pt boxes, every fraction); no settings has 59 (a first line whose top is exactly on a
half pixel, and a text box inside a group, which Word draws with its own spacing and
partly as a picture). A scratch sweep of four faces at 7-30 pt tried: the inner top
rounded half up, half down, down and up; laid out from 0 and moved; the top inset as
space before in the first line's box; the text box laid out as below mode 15 -- none
better than the exact top (the one adopted). **Probe that would decide it**: the sweep
crossed with the paragraph mark's size and the number of lines, per face, both settings.

### F.4 The real documents

`filesamples/sample3` and `sample1` anchor pictures text wraps around (`wrapSquare`,
`wrapTopAndBottom`): stage 1 leaves them where they were. The **local corpora** are the
documents this stage was for: each of their templates' anchors is `wrapNone`, positioned
against the column and the paragraph. **Every one of them is now laid out to its end** --
no stop band, every page Word made -- with its title's picture, its shapes, text boxes
and groups drawn; their lines and glyphs are scored on the machine that holds them. What
they show that is general is recorded here in general terms: dotted outlines
(`prstDash sysDot`) are drawn solid and warned; some text boxes' first lines stand 2-3 px
off Word's (F.3's open question; since measured and adopted, F.13); and in three, Word
keeps a section's last, empty paragraph (the one holding its `w:sectPr`) on a page it
does not fit on, where the model moves it to a page of its own and so makes a page more
-- pagination, not drawings (since probed and adopted, 4.11: every local document now
has Word's pages).

**Nothing that was exact moved.** Checked line by line against the commit before over
134 documents -- the committed ones, `filesamples` and their variants, the local corpora
and the generated probes, 331,625 lines and 2,423,565 glyphs (every baseline, every
glyph's x, face and size, every rule, picture, page geometry and stop): only the local
corpora changed, and only by drawing what was not drawn (their stop band gone, lines and
pages added). The committed documents hold no floating drawing: their VRT snapshots and
fidelity are unchanged.

### F.5 Extra rule: the text column starts on the nearest layout unit

Found by the local corpora, whose left margin is 576 twips (120 px): Word starts their
lines at 120.0002 px, the model at 119.9992 -- H.5's rule held the rounded pixel in the
layout unit below it (`floor`), which every earlier margin could not tell from the
nearest (1440 twips, 300 px, is 294,912 units exactly; 301 px 295,895.04). **Measured on
its own probe cases and adopted**: `make_margin_probe.py` now also has left margins of
576, 1152 and 1296 twips (120, 240, 270 px: 117,964.8, 235,929.6, 265,420.8 units), and
Word starts every line of them on the nearest unit: object starts 100 / 100 with no
settings and 97 / 100 in mode 15 (before: 85 / 100 and 82 / 100; the three left are H.5's
right-aligned footer lines). A text box's lines start the same way (F.3). Checked line by
line against the commit before over 138 documents: only the margin probe's new pages and
the local corpora's lines moved, each by 0.001 px, onto Word's.

### F.6 Stage 2 — a header's and a footer's drawings

`make_story_anchor_probe.py` (read by `read_anchor_probe.py`): twelve sections, each a
page with its own header and footer -- a header's picture against its paragraph (in
front of the body's text it overlaps, and behind it), the page, the margin, the top
margin, its line and character; a footer's against its paragraph, the page's foot and
the bottom margin; a header's text box; and a body's picture behind and in front of the
text over the header's area. Two documents: no `settings.xml` and mode 15.

* **Positioned by F.1's rules unchanged**, from the story as it is laid out (H.1): a
  header's first paragraph's top is its distance from the page's top, a footer's is
  where it grows up to; the frames are the page's margins. **11 / 11 pictures where Word
  drew them, in its order and layer, in both settings; 2,266 / 2,266 glyphs exact; the
  text box's fill exact.**
* **Stacking** (`layout.Page.paint`): Word paints a story before the body -- the story's
  drawings behind its text, its text, **its drawings in front of its text, and only then
  the body's drawings behind the body's text**, the body and its drawings in front. So a
  header's logo "in front of the text" is under the body's text it overlaps, and a body's
  drawing "behind the text" is over the header's text. A page that has no drawing either
  way is painted as before.
* `story-drawing-not-drawn` is gone: a story's drawing text does not wrap around is
  drawn. One text wraps around still makes the story stop there (`story-stopped:drawing`).

The visibility check covers `story-anchor-15` (0 hidden). Checked line by line against
the commit before over 138 documents: only the new probe's pages; no real document here
has a drawing in a header or footer.

### F.7 Stage 3 — `wrapTopAndBottom`: text continues below

`make_wrap_anchor_probe.py` (read by `read_anchor_probe.py`): 25 pages -- a picture
against its paragraph at offset 0, below the first line, above the paragraph and inside
its lines, the anchor first and in the fourth line; against the page and the margin over
paragraphs before and after the anchor's; `distT` / `distB` of 0, a twip, 0.125 in and
0.25 in; anchored near the page's foot so the text below goes to the next page; heights
stepped by 2 twips. Paragraphs of `auto` 259 with 120 twips after. Two documents.

**The rules** (`paginate.page_bands`, `paginate.clear_of_bands`, `layout._walk`):

* **The drawing is positioned first, from the page as laid out without it** (F.1's
  rules), and does not move when the text does: a drawing 65 px above its paragraph
  stays there while its paragraph goes below it.
* **Its band** is its box, as tall as its extent **in whole twips** (not the picture's
  drawn height, up to 0.06 px less), widened by `distT` above and `distB` below (whole
  twips), across the column.
* **A line that would reach into a band starts at the band's foot** -- every line, the
  paragraphs before the anchor's on the page included -- and everything after it follows;
  a line that only touches the band's edge stays. A line reaches down its pitch, and
  **below mode 15 a paragraph's last line also its space after** (a line whose space
  after ends 0.04 px inside the band moves; in mode 15 one whose pitch ends 5 px short of
  a band 20 px into its space after stays).
* **Below mode 15, the paragraph's and the line's tops a wrapping drawing is positioned
  from are taken in whole twips, the nearest** (514.2049 px drawn at 514.1667, 2,468
  twips; 2,554.58 at 2,554.79): every picture; in mode 15 exactly.
* The paginator lays each page out twice -- once to position its drawings, once with
  their bands -- and a line pushed past the foot goes to the next page as any line does;
  the page walk pushes the same lines by the bands the paginator found
  (`Pagination.bands`).

| | Pictures where Word drew them | Glyphs at Word's position and baseline |
| --- | --- | --- |
| `wrap-anchor-none` | **25 / 25** | **26,741 / 26,741** |
| `wrap-anchor-15` | **25 / 25** | **26,741 / 26,741** |

**Refuted**: the band as the picture's drawn height (2 lines a pixel high in mode 15); the
space after counted in mode 15 (130 glyphs moved that Word keeps); the paragraph's top
exact below mode 15 (20 of 25 pictures 0.04 px low, and a line pushed that Word keeps);
truncated to twips rather than the nearest (a picture a twip high near the foot).

Not modelled: a band that pushes the anchor's own line to the next page (the drawing
stays where it was positioned); a table's rows (not pushed); a `wrapTopAndBottom` drawing
in a header or footer or a table cell (the story stops there, the table is
`layout-stopped:table`). `filesamples/sample1`'s one is past a table the layout stops at.

### F.8 Stage 4 — text beside a drawing (`wrapSquare`, `wrapTight`, `wrapThrough`)

`make_wrap_side_probe.py` / `read_wrap_side_probe.py`, `tests/test_wrap_side.py`
(`tests/fixtures/wrap-side-observations.json`). Six documents, 944 pages. `wrap-side-none`
/ `-15` (147 pages each, left margin 1,442 twips: off the pixel grid): families `square`
(gaps of 100-1,440 twips either side, long words beside a narrow gap, `wrapText` `left` /
`right` / `largest`, indents beside a drawing, centred / right / justified text, two
drawings side by side, stacked and overlapping, the drawing's top stepped through a line,
each `dist`, a drawing above its paragraph, taller than it, against the margin and the
page, wider than the column), `min` (gaps of 340-400 twips at 8, 11 and 20 pt), `vertical`
(top and bottom stepped by 2 twips across a line's edges under `auto` 259 and 240, `exact`
320, `atLeast` 400), `polygon` (tight and through around a triangle, a diamond, a C and a
polygon past the extent, with and without `distL` / `distR`) and `more` (nothing fits
beside it, narrow `left` / `right` wraps, first-line and hanging indents beside a drawing
inside the indent, tabs, an effect extent, `distT` on a tight wrap, empty, centred and
justified paragraphs, a drawing near the page's foot). `wrap-edge-<setting>-<margin>`
(90 pages, margins 1,442 and 1,440): the right edge's offset, extent, `distR` and a
page-relative offset each swept by single twips (`right`); words of known width with the
drawing's left edge swept by single twips about the end of the third, fifth and seventh
(`left`); edges on exactly half a pixel (`tie`). Scratch sweeps before it (a dozen
exports, not committed) found the rules; the committed probe holds them.

**The rules** (`docx2svg.wrap`, `paginate.page_wraps`, `paginate.fit_beside`):

* **The drawing is positioned first**, from the page laid out without it (F.7's rule: a
  drawing 131 px above its paragraph wraps the paragraph before it, and stays where that
  layout put it when those lines grow). Below mode 15 a drawing aligned in the column is
  aligned **between the paragraph's indents**, and on a whole twip (a centred drawing's
  half twip up): 450.42 px with a 720-twip indent where mode 15 has 300.42.
* **Which lines are beside it**: a line whose top is above the box's foot and which
  reaches below its top -- the box its extent in whole twips with its effect extent, widened
  by `distT` / `distB` -- touching not counting. A line reaches down its pitch **less an
  `auto` multiple's extra in mode 15** (a drawing 268 twips below an `auto` 259 line's top
  is beside it, 270 is not: the text is 268.55 twips) and its **whole pitch below mode 15,
  with its space after on a paragraph's last line** (289.8 twips; a paragraph's last line
  is beside a drawing 4 px into its space after). `exact` and `atLeast` lines reach their
  pitch in both.
* **What it keeps text off**: its box with its effect extent (63,500 EMU of effect moves the
  text 20.83 px), widened by `distL` / `distR`. `wrapTight` and `wrapThrough` behave alike
  (every case): the polygon, each point in whole twips from the drawing's corner,
  truncated, scaled from 21,600 to the extent in whole twips, read across the line's
  *text* in both settings -- every line of the triangle, the diamond and the C, 303 / 303 in
  both; the C's opening lets text in; its height is the polygon's, widened by `distT` /
  `distB` in mode 15 and **not below mode 15**.
* **`wrapText`**: `left` / `right` keep only that side; `largest` the side with more room
  between the drawing and the column's edges, **indents aside** (a 1,440-twip indent that
  makes the left side the narrower does not change it), the left one on a tie.
* **Segments**: what the drawings leave of the column, left to right; **one narrower than
  360 twips takes no text** (351-359 take none, 360 does, left and right, at 8, 11 and 20
  pt alike: a width, not a word). Text fills the segments in order, each as a line by
  Phase 3's breaker, **up to the drawing's left edge less `distL` exactly** -- measured from
  the left margin's exact edge, to the layout unit: every one of 294 lines of words of
  known width, the edge swept by single twips, in both settings and on both margins (the
  third, fifth and seventh word fits at an edge of 1487, 2511, 3535 twips for ends of
  1486.504, 2510.664, 3534.824). **A segment whose next word does not fit is passed over**:
  no word is broken in it (a 146-px segment beside `Anextraordinarily...` stays empty, the
  word goes right of the drawing). **A line no segment takes text on goes down to the foot
  of the drawings it is beside** and is tried again (wider than the column, `wrapText`
  `left` with no room left of it, long words beside two 450-twip gaps).
* **Where a segment's text starts**: at the drawing's right edge plus `distR` from the left
  margin's exact edge, `w`, or at the left indent where that is further right; the first
  line's indent moves it from there, a hanging one no further left than the edge (720 /
  first line 360 beside a drawing 787 px wide: the edge plus 75 px; hanging 540 beside one
  inside the indent: the edge). Then **the segment moves by a rounding on the device-pixel
  grid**: to the edge's pixel on the page, `round(E)`, from `C + w` (`C` the text column's
  start, a whole pixel) below mode 15 -- text right of a drawing starts on a whole pixel --
  and from `round(C + w)` in mode 15 -- on the edge where the margin is on the grid, up to a
  pixel right of it where it is not. Offset, extent, `distR` and page-offset sweeps: 816 /
  816 lines in all four documents (the old probe's one-pixel mystery). Tabs right of a
  drawing: stops from the column's left edge in mode 15, **from the segment's start below
  it**.
* **Alignment** is each segment's: centred, right-aligned and justified text is set in its
  segment as a line in the column (each segment of a justified line reaches its edge).
* **Pagination**: a line beside a drawing is broken, and its paragraph's rest, where the
  paginator places it (`Para.source`, `Para.set_lines`): the page is laid out once without
  its drawings, which are positioned from that layout, then again with them, as F.7's
  bands; lines the layout moves down and lines it adds move the page's end (a drawing near
  the foot sends the paragraph's rest to the next page, as Word does). The page the layout
  stops on is laid out with its drawings too.

| Lines within ½ px of Word's (every glyph at Word's pen x and baseline) | before | after |
| --- | --- | --- |
| `wrap-side-15` (square / min / vertical / polygon / more) | 3 / 2,477 | **2,451 / 2,477** (974/974, 410/410, 432/432, 277/303, 356/356) |
| `wrap-side-none` | 3 / 2,498 | **2,347 / 2,498** (966/979, 405/410, 436/436, 303/303, 235/368) |
| `wrap-edge-15-1442` (left / right / tie) | 1 / 786 | **766 / 786** (294/294, 384/384, 88/108) |
| `wrap-edge-15-1440` | 1 / 786 | **786 / 786** |
| `wrap-edge-none-1442` | 1 / 786 | **771 / 786** (294/294, 384/384, 93/108) |
| `wrap-edge-none-1440` | 1 / 786 | **771 / 786** |
| Pictures where Word drew them, in its order | 0 / 658 | **658 / 658** |

Before, every probe stopped at its first drawing (`layout-stopped:drawing`).

**Not settled, pinned in the test**: (1) **an edge on exactly half a pixel** (the column's
in mode 15, the page's below it: one twip in 24 on an off-grid margin) goes to the pixel
up or down by a rule not found -- 1,507.5, 1,512.5, 1,532.5, 1,822.5 up; 1,397.5, 1,422.5,
1,427.5, 1,447.5, 1,517.5-1,527.5 down; no float rounding in points, twips, inches or EMU
reproduces it. The model takes the column's half down in mode 15 and the page's half up
(below mode 15, the page's half down): the tie family's 20 (mode 15) and 15 (no settings)
lines, and the polygon past its extent with `distR` (26 lines). (2) **Below mode 15 only**:
a line no segment takes text on goes further down than the drawing's foot (by 3-54 px,
not a pitch, not a fixed step); `wrapText="left"` with less than 360 twips left of the
drawing puts the text right of it (`right` on the right edge goes down, as in mode 15),
and with 400 twips breaks words into it letter by letter; a hanging first line beside a
drawing inside the indent starts at the edge's pixel plus the indent less the hanging -- 133 lines of
`more`, 13 of `square`, 5 of `min` (ties). Recorded, not modelled.

**Not modelled, and warned**: a table beside such a drawing is laid out across the whole
column (`wrap-table-not-modelled`: Word moves it clear; *modelled since, F.15*); a drawing text wraps around in a
header or footer (Word wraps the *body's* text around it; the story still stops there,
`story-stopped:drawing`); one in a table cell (the table stops, F.1; *since F.16 a
`wrapNone` one is drawn*). Two drawings' paint
order is the order met, as F.2's (three pages of two drawings: all six where Word drew
them).

**The real documents**: `filesamples/sample3` wraps its first page's text around a picture
and its second page's around a chart: page 1 is now drawn in full (fidelity 0.9884) and
page 2 to a table in its two-column section, every paragraph on Word's page (`test_pages`
26 / 26, and the page top). The committed documents hold no floating drawing: their
output, VRT snapshots and fidelity are unchanged; the local corpora's drawings are all
`wrapNone`, and none of their scores moved.

### F.9 Extra rule: a line break that ends a paragraph ends a line of its own

Found by the local corpora's text boxes, whose list paragraphs end in `<w:br/>`: Word draws
the paragraph's mark on a line of its own after the break, and the model drew none -- so
every line after it stood a line high. It is the line breaker's (Phase 3), not the text
box's: a body paragraph does the same. **Measured on its own probe and adopted**
(`make_break_end_probe.py` / `read_break_end_probe.py`, `tests/test_break_end.py`): one,
two and three trailing breaks under `auto` 240 and 360, `exact` 400 and `atLeast` 480,
with space after, under a 20 pt mark, before a space, in a table cell -- every one of
1,034 glyphs at Word's baseline in both settings (**820** before). A page or column break
that ends a paragraph is unchanged (the paginator's, Phase 4). Checked line by line over
142 documents: no body line of any document moved -- no committed or `filesamples`
document ends a paragraph in a line break -- and in the local corpora only text boxes'
lines, onto Word's.

### F.10 Proposal for `ooxml-common`: DrawingML fill, outline, colour and geometry rendering

**Done**: the renderers moved to `ooxml-common` with pptx2svg's DrawingML value types,
and the reader after them; docx2svg draws with them (F.12). The proposal as it was made:

Not made here: `pptx2svg` and `ooxml-common` are another agent's to change. `pptx2svg`'s
`render/fill.py` and `render/geometry.py` were left out of the extraction (Phase 1)
because they take `pptx2svg`'s model types; `docx2svg` now needs the same capabilities
and has written only the glue it could not avoid (`docx2svg.drawing`), so the two now
overlap. **Proposed**: lift the DrawingML value types out of `pptx2svg.model` into
`ooxml_common.drawingml.model` -- colour choice with its transforms, `SolidFill`,
`GradientFill`, `PatternFill`, `BlipFill`, `NoFill`, `Outline` (width, fill, dash, cap,
join, head and tail ends), preset and custom geometry, `Transform2D` (offset, extent,
rotation, flips, a group's child offset and extent) -- and move beside them:

1. **Geometry**: the preset table *complete* (its generator today leaves out the presets
   `pptx2svg` draws by hand -- `rect`, `roundRect`, `ellipse`, `line`, `rtTriangle`,
   `straightConnector1`, the flowchart shapes, callouts -- which `docx2svg` has had to
   write from ECMA-376 in `drawing.EXTRA_PRESETS`), and one function from a geometry,
   its adjustments and a box to SVG path data with each path's fill mode and stroke flag
   (what `docx2svg.drawing.geometry_paths` does and `pptx2svg`'s `render/geometry.py`
   does with its hand-written presets).
2. **Colour**: resolving a colour choice against a theme's colour scheme and a colour
   map, with every transform (`lumMod`, `lumOff`, `tint`, `shade`, `satMod`, `alpha`...)
   and one rounding (`docx2svg` measured Word's: HSL, each channel rounded a half down;
   `pptx2svg`'s `parse/drawing.py` has `_hsl_to_hex`). `docx2svg` implements only
   `lumMod` / `lumOff` and warns of the rest.
3. **Fills and outlines as SVG**: `render_fill_attrs` (solid, gradient, pattern -- with
   `ooxml_common.drawingml.pattern` -- and picture fills) and `render_outline_attrs` with
   dashes, caps, joins and markers. `docx2svg` draws solid fills and solid outlines and
   warns of gradients, patterns, picture fills, dashes and effects (a placeholder where
   nothing else could be drawn).
4. **Effects** (`render/effect.py`: shadows, glow, soft edges), at least as a list of
   what a shape asks for, so both can warn alike.

**What `docx2svg` needs from it, in order of the local corpora's use**: dashed outlines
(`prstDash sysDot`), the complete preset table, colour transforms beyond `lumMod` /
`lumOff`, gradient fills, effects. Each consumer keeps its own layout (a text box's
content, a slide's placeholders) and its own measured positioning; only the value types
and their drawing move.

### F.11 Fidelity and what is left

The committed documents hold no floating drawing, and no committed document's output
changed in any stage (their VRT snapshots are byte-identical), so
`tests/fidelity-baselines.json` is untouched: scores as before. The probes' pages were
rendered and looked at beside Word's (presets, custom paths, flips, the arc and the cubic,
groups, the rotated rectangle and the gradient's placeholder, headers' pictures, the
pushed text of `wrapTopAndBottom`): Word's to the eye but the gradient. The local corpora
were scored and looked at on the machine that holds them (their numbers are kept there).

Open, in the order they would help: text beside a drawing's ties and its behaviour below
mode 15 (F.8), a table beside a drawing (*since measured and modelled: F.15*) and a header's
drawing that body text wraps around (F.8); anchors in table cells (F.1; *since F.16, F.17*); `allowOverlap="0"`; a text box's line in mode 15 at larger sizes (F.3, since settled: F.14); a text
box's vertical or rotated text (`bodyPr@vert`, `@rot`: warned, not laid out); mirror
margins (`w:mirrorMargins`, which also moves the body's text, not read). Dashed and dotted
outlines, gradients and effects: F.12.

### F.12 DrawingML drawn by the shared renderers — measured

F.10's proposal is done, in `ooxml-common` 0.3: pptx2svg's DrawingML value types,
renderers and XML reader moved there with their history, and docx2svg now reads a
shape's `spPr`, `wps:style` and a theme's format scheme with the shared reader
(`ooxml_common.drawingml.read`) and draws geometry, fills, outlines, arrowheads and effects
with the shared renderers. What stays here is Word's own (`docx2svg.drawing`): a theme
colour through `w:clrSchemeMapping` (Word writes `t1="dark1"` where a slide master writes
`tx1="dk1"`: both spelled as DrawingML's slots now), the style references, the group
fill, and where a shape goes and in which order it paints (F.1-F.7, unchanged).

**The renderers were measured against PowerPoint, so they were measured again against
Word** before docx2svg drew with them: `make_dml_probe.py` / `read_dml_probe.py`,
`tests/test_dml.py` (`tests/fixtures/dml-observations.json`). 81 pages, one case each,
read off Word 16.106's PDF (fills and strokes by PyMuPDF; each gradient's shading --
type, coordinates in the shape's frame, colour samples, profile -- from the PDF's
resources): colour transforms (61 swatches); linear gradients at 0-300 degrees, `scaled`
on and off, three, unsorted and inset stops, `circle` / `rect` / `shape` path gradients
with centred and cornered `fillToRect`, `rotWithShape` on rotated and flipped shapes, a
gradient outline; 12 pattern presets, one off the lattice, one rotated; every
`prstDash` at each cap and two widths, a `custDash`, dashed closed shapes; joins; every
arrowhead type at five sizes and two widths; `cmpd` lines; theme style references
(Office's own format scheme); `a:grpFill` under solid, gradient, pattern and no group
fill, nested; outer and inner shadows, glow, soft edge; WordArt text with
`w14:textFill` / `w14:textOutline`; automatic text in text boxes on 32 fills and under
6 font references; a shape and a dotted ellipse drawn inline.

Where Word's drawing differs from what the renderers drew for pptx2svg, the difference
is a field of `ooxml_common.drawingml.rules.DrawingRules` (and `color.ColorRules`), and
docx2svg draws with `WORD`; pptx2svg's `POWERPOINT`, the default, is unchanged (its
suite, VRT snapshots and fidelity baselines byte-identical). **The rules:**

* **Colour transforms compose in document order on unrounded channels.** `lumOff 40000`
  before `lumMod 60000` is `517CC8` (the paired pass gives `8FAADC`); Office's theme
  gradient stops (three transforms each) came out a level off when every step rounded;
  the saturation is not clamped at 1 (`satMod 200000` on `4472C4` is `0460FF`, the HLS
  formula with the channels clamped; clamping the saturation gives `0961FF`); `satOff`,
  `hueMod`, `hueOff`, `comp` (the hue turned half way), `gray` (Rec. 601 luma of the sRGB
  channels) and `inv` (in linear light) are applied; `a:scrgbClr` is linear light
  (`r=50000 g=20000` is `BC7C00`); every `a:prstClr` name (`darkSeaGreen` `8FBC8F`).
  `tint` and `shade` in linear light, `lumMod` / `lumOff` / `satMod` in HLS and F.3's half
  down are PowerPoint's, to the level. **61 / 61** swatches (9 before).
* **A linear gradient runs through the box's centre, corner to corner** -- from where the
  first corner meets its direction to where the last does, `w |cos| + h |sin|` long
  (1,908,750 EMU at 45 degrees on a 1,799,590 x 899,795 box, Word's to the EMU); with
  `scaled="1"` the unit square's gradient stretched onto the box (45 degrees runs at
  63.4); `rotWithShape="0"` holds the angle to the page against the shape's rotation and
  flips. A `circle` path gradient runs from its `fillToRect` point to a circle round the
  box's centre through the corners (a little wider when the point is a corner itself); a
  `rect` path gradient is rectangular rings, drawn as four trapezoids (Word draws it as a
  picture); a `shape` one on an ellipse is its inscribed ellipses. **A gradient of exactly
  two stops, at 0 and 100%, eases in linear light**: Word writes it in a linear profile
  with the cosine ease `(1 - cos pi t) / 2` between them (`890000` to `002A89` for `C00000`
  to `0070C0`, `730614` a quarter of the way where a straight blend gives `670B22`); every
  other gradient is sRGB and straight. A stop's `alpha` is not drawn in Word's export.
  **30 / 30** gradients where Word drew them (4 before: the ones Word draws as pictures
  matched the old placeholder's "not a shading").
* **Dashes**: flat caps draw the preset times the width (`sysDashDot` `3 1 1 1`,
  `sysDashDotDot` `3 1 1 1 1 1` added); **a round cap shortens each dash by the width and
  lengthens each gap by it**, so a rounded dash covers the stated length (a round `sysDot`
  is dots a width apart); a square cap draws the flat
  pattern and squares only the line's two ends (drawn butt here: the ends' half width
  short). **69 / 72** lines (6 before); the three left are the reader's (lines it cannot
  take apart in Word's PDF: two dotted ellipses and a merged dash), listed in the test.
* **An outline that states no join is round** (SVG's default is miter).
* **Arrowheads** are 2, 3 and 5 times the line's width for `sm`, `med` and `lg`, 2 pt at
  least (a 1 pt line's `lg` head is 10 pt), the line cut back under a triangle to half a
  unit behind its base and under a stealth to its middle; a stealth's notch is 0.6 of
  its length from the tip; a diamond and an oval are centred on the end
  (`fill.render_arrowheads`, drawn as shapes: SVG markers cannot cut the line). Looked at
  beside Word's, not scored.
* **A pattern is registered to the page's corner and square to the page** on a rotated
  shape (Word's tile at 88 pt for a shape at 95.7). **14 / 14** (0 before).
* **Theme style references** (`wps:style`): `fillRef`, `lnRef` and `effectRef` index the
  theme's lists **from 1** (`effectRef idx="3"` is the third, the shadow), `fillRef` 1001
  and up the background fills, `phClr` the reference's colour; a shape's own `a:ln`
  overrides the line style's parts it states. **11 / 11** fills and outlines (0 before).
* **A text box's automatic text** takes its style's `fontRef` colour where it names one,
  whatever the fill (`tx1` on navy stays black, `lt1` on cream is white); without one, it
  is **white on a solid fill whose Rec. 601 luma is below 74** and black otherwise: grey
  `48` (72) and darker, `E00000` (67), `C00000`, `404000` turn it white, grey `4C` (76),
  `008000` (75.1), `7030A0`, `4472C4` leave it black -- the threshold lies between 72 and
  75.1 (`drawing.automatic_text_colour`). **48 / 48** (27 before).
* **`a:grpFill`** is the group's fill spread over the group's box (a gradient runs across
  the whole group), passed down through a group that states none; with no group fill it
  draws nothing -- which is how a heading converted to letter outlines is filled.
* **WordArt**: a run's `w14:textFill` (solid, gradient -- across the run's text -- or none)
  and `w14:textOutline` are drawn on its glyphs, in a text box or in the body. A
  `prstTxWarp` other than `textNoShape` is not (warned).
* **Effects** (shadows, glow, soft edge) are the shared SVG filters, looked at beside
  Word's: the shadows are Word's to the eye; Word's glow is a tighter, stronger halo than
  the filter's and its soft edge eats further in -- not measured further.

**Not drawn, and warned**: a picture fill on a shape (`blipFill`), compound lines (`dbl`,
`thickThin`, `thinThick`, `tri`: one stroke of the full width, as pptx2svg draws them),
3-D. **Not measured**: a `custDash` with a square cap (Word squared every dash of the one
case), a `rect` path gradient's exact corner, effects' sizes.

**An instrument defect, found and fixed.** A local template scored *worse* once its
shapes were drawn: the fidelity truth had lost them. Word's export wraps every drawing in
Quartz's "no clip" -- a rectangle half a billion points across -- which MuPDF writes into
the converted page's clip path as it is, and resvg drops a clip that far out, and the
drawing inside it with it. The converter's validation did not see it, because it compares
MuPDF's raster of the SVG, which draws it. `pdf_svg.rasterise` now brings such
coordinates in to a million points (`bounded_clips`). No committed document has a
drawing under it: their truth is unchanged, and so are their baselines.

**Scores.** The committed documents hold no DrawingML shape (F.11): their VRT snapshots
and fidelity are byte-identical. The probe, before and after:

| `dml-15` | before | after |
| --- | --- | --- |
| colours | 9 / 61 | **61 / 61** |
| gradients | 4 / 30 | **30 / 30** |
| dashes | 6 / 72 | **69 / 72** |
| patterns | 0 / 14 | **14 / 14** |
| style references | 0 / 11 | **11 / 11** |
| automatic text colour | 27 / 48 | **48 / 48** |

`drawing-none` and `drawing-15` (F.3) keep every score; their gradient is drawn, so the
`drawing-not-drawn` warning is gone from their recording.

**Inline shapes** (`wp:inline` holding a `wps:wsp` or a `wpg:wgp`) were placeholders; they
are drawn now, in their line's box as an inline picture is placed (4.7), through the same
path as a floating drawing, in front of the text: the probe's two, where Word drew them.

The local corpora (F.4) were scored before and after on the machine that holds them,
with the fixed truth for both: every document with a DrawingML shape scores better by
SSIM, by the colour histogram and by the unnormalised loss; every other one, and every
`filesamples` document, is unchanged. Their numbers are kept with them.

**Text colour in dark boxes, checked on the local corpora.** With the automatic text rule
above, every glyph Word drew in a local template's text boxes and body is drawn in the
colour Word drew it in (each page's glyphs counted per fill colour, Word's PDF against
the layout's spans: no colour where the counts differ but by the spaces Word writes as
glyphs). Their run colours state a theme colour with a tint or shade whose cached
`w:val` is the theme's own, so the model's `w:val` is Word's colour there; a document
whose `w:val` disagrees with its `w:themeColor` was not found, and Word's choice between
them is not measured.

### F.13 A text box's text area — measured

The local corpora's text boxes put their text 2 px right and down of Word's at 300 dpi,
and 8-11 px in some rounded rectangles. Neither was the insets: the boxes state
`a:ln` with `a:noFill` under a style whose line reference names the theme's 1 pt line,
and some are `roundRect`s. **Measured before it was implemented**:
`make_text_box_probe.py` / `read_text_box_probe.py`, `tests/test_text_box.py`
(`tests/fixtures/text-box-observations.json`), 115 boxes per document, each's first
glyph's pen x and baseline read from Word's PDF by PyMuPDF (a character's origin: the
pen on the baseline, the convention of the model's `Span.xs` / `Span.y`; the instrument
was first held to the body lines of the local documents, every one within 0.01 px).
Families: the outline (none, drawn at 0.5-8 pt, `a:noFill` with a width, the style's
line reference 1-3 under `a:noFill` and drawn, reference 0, a drawn line stating no
width), each at `anchor` `t` / `ctr` / `b`; insets (zero, the default, others) with and
without an outline, centred and right-aligned; geometry (`roundRect` at four
adjustments, `ellipse`, `triangle`, `rtTriangle`, `octagon`, a custom geometry with its
own `a:rect`) at `t` and `ctr`; `wrap="none"`; sizes 8-28 pt, two paragraphs under
`auto` 276, `exact`, spacing; recorded only: `spAutoFit`, `normAutofit`, `bodyPr@rot`,
`vert`, `vert270`, a shape turned 90 degrees. No `settings.xml` and mode 15.

**The rule** (`drawing.text_area_edges`, `layout._text_box`):

* **The insets count from the geometry's text rectangle** -- the preset's `a:rect` from
  ECMA-376's definitions (`ooxml_common.drawingml.geometry.text_rect`, over
  `preset_text_rects`), a custom geometry's own: a `roundRect`'s text keeps clear of its
  corners by 0.29 of the corner's radius (11.2 px on a 700,000 EMU-high box at the
  default adjustment, Word's 11), an ellipse's is its inscribed rectangle (96 / 34 px,
  Word's 96 / 33-34), a triangle's, a right triangle's and an octagon's the
  specification's, to the pixel.
* **Inside half the outline's width, on every side**, whether the outline is drawn or
  `a:noFill`, its width the shape's own `a:ln@w` or its style's line reference's: a
  1 pt style line under `a:noFill` moves the text 2.08 px right and down (Word: 2), 8 pt
  16.7 (16-17); `anchor="ctr"` moves it neither up nor down, `b` up. A line that states
  no width and no style moves nothing (a drawn default line included).
* The left edge still rounds to a whole device pixel a half down (F.3), and the top is
  exact.

| `text-box` first glyphs within half a pixel | before | after |
| --- | --- | --- |
| no settings | 34 / 115 | **100 / 115** |
| mode 15 | 34 / 115 | **101 / 115** |

**Not settled, pinned in the test** (*(1) and (2) since measured and settled: F.14*): (1) a left edge exactly on a half pixel went **up**
in five of the probe's boxes (both settings) and **down** in `make_drawing_probe.py`'s
`default insets` box and in one local text box. Summing the margin, the offset and the
inset each rounded to a whole layout unit (1/4096 pt) reproduced every probe box and
failed the local one; rounding each a different way reproduced all nine ties, which is
fitting four choices to nine cases -- neither adopted. A probe that would decide it
sweeps the margin, the offset and the inset each by a fraction of a layout unit. (2) A pixel of baseline (F.3's open
question): 14 and 18 pt boxes, `tIns` 0.1 in, 8 pt outlines at the bottom. (3) Centred
text under `wrap="none"` stands 146 px left of the model's: Word does not centre it on
the box's inner width; not modelled. (4) What `spAutoFit`, `normAutofit` and vertical
or rotated text do is recorded (Word's positions in the fixture), not laid out.

**The local corpora**: their text box lines started 2 px (the style's 1 pt line) to
9 px (a `roundRect` besides) right of Word's; now every one starts on Word's pixel but
one (a tie, (1)), and a quarter of them stand a pixel off Word's baseline (2); the templates
with text boxes score better by SSIM, the colour histogram and the loss, the others are
unchanged, and the visibility check finds no hidden glyph. `make_drawing_probe.py`'s
boxes state no outline and a `rect`: its recording is unchanged. The committed documents
hold no text box: their output and fidelity are unchanged.

`ooxml-common` gained the text rectangles (a branch there: the generated
`drawingml/preset_text_rects.py` from the same ECMA-376 source file, its generator
`tools/derive_preset_text_rects.py`, `geometry.text_rect` and `read.parse_text_rect`);
nothing pptx2svg calls changed.

### F.14 A text box's first baseline and its left edge — measured

F.13 left a quarter of the local templates' text box lines a pixel off Word's baseline --
most of what was left of their loss -- and a left edge on exactly half a pixel that Word
rounded up five times and down twice. **Measured before it was implemented**:
`make_text_box_top_probe.py` / `read_text_box_top_probe.py`, `tests/test_text_box_top.py`
(`tests/fixtures/text-box-top-observations.json`). 1,412 boxes on 33 pages, one to three
labelled lines each, every first glyph's pen x and baseline read from Word's PDF as F.13's
reader reads them (PyMuPDF's character origin, the model's `Span.xs` / `Span.y`). Each
family is a **sweep**: one quantity moved by a fraction of a pixel from box to box, on both
axes at once, so that where the baseline or the pen x steps to the next pixel says where
Word starts the text. Families: the offset by whole twips (six to a pixel) at 8-24 pt,
anchors `t` / `ctr` / `b`, one line, two lines (a break) and one line over an 11 pt mark;
the offset by single EMU about a twip's edge and by fifths of a twip; `lIns` / `tIns` /
`bIns` and the outline's width by fifths of a twip, and an inset under an outline half a
fraction of a twip wide; the box's height and its bottom inset by fifths of a twip below
one line; boxes against the paragraph (below a heading line) and against margins off the
pixel grid; two paragraphs with space before; `auto` 276 and 200, `exact` and `atLeast`
400; Times New Roman and Arial; recorded only, boxes flipped and turned. Three documents:
no `settings.xml`, mode 14 and mode 15.

**The rules** (`layout._text_box`, `layout._walk`'s `story_bottom`, `drawing.text_area_edges`):

* **Word holds a text box in whole twips.** Its corner is the nearest twip (the offset is
  already truncated, F.1; a box against its paragraph starts a fraction of one off it: a
  paragraph top of 1,708.55 twips is taken as 1,709); each side's inset **plus** its
  distance in from the shape's edge (F.13's text rectangle and half outline) is rounded to
  the nearest twip, a half up -- `tIns` swept by fifths of a twip steps the baseline where
  74.6 twips becomes 75, not where 74.5 would. **The outline's half is whole EMU,
  truncated**: a 635 EMU outline over 72 twips of inset keeps the text 72 twips in (317 EMU,
  not 317.5), and an 889 EMU one likewise; a custom text rectangle on exactly half a twip
  goes up (`make_text_box_probe.py`'s `custom rect`).
* **The text starts at that point in the nearest layout unit** (1/4096 pt), and its lines on
  that point's device pixel, **a half up**. This is the half-pixel tie: an integer number of
  twips `n` on half a pixel (`n` = 12 mod 24) is `n` x 204.8 units, whose fraction is .2,
  .4, .6, .8 or 0 -- so 2,724 twips (567.5 px, `make_drawing_probe.py`'s `default insets`
  box) is held at 567.4998 and goes down, 1,836 (382.5) at 382.5002 and goes up, and 6,180
  (1,287.5, exactly a unit) goes up. Every tie F.13 recorded (the probe's and the local
  one), and every tie of the sweeps (four on the off-grid margin), is Word's. The half-down rule is gone.
  The top is the same point's unit, and a baseline on a half pixel is decided the same way
  (a box against its paragraph at 632.5 px goes down at 1,272.5 px: .8 and .4 of a unit).
* **Below mode 15 the room between the insets is a twip shorter**, truncated to the layout
  unit (204 of 204.8 units): a bottom-anchored or centred box stands that much higher, and
  the last line keeps that much less below it. The sweeps admit 203.1-204.3 units (204.8
  misses two ties); mode 14 is as no settings.
* **`ctr` and `b`** move the content by the room less the content (half of it for `ctr`),
  exactly, when it fits, and **the moved top is held in the nearest unit** again.
* **The box's last line keeps the rest of the room below its text.** It is the difference
  from the body, and it is the pixel of baseline: the line box of the last line of the box's
  last paragraph reaches down to the room's foot (after a `ctr` or `b` move, the room's foot
  moves with it), so it rounds as a line with space below it does (the line box, "space
  below the text"): the text on its own rounded height, never above its rounded ascent. A
  one-line 14 pt box stands a pixel lower than the model had it at every position (the
  body's rule rounds a 71.2 px line against its top by half its fraction; with room below,
  its ascent of 55.54 rounds whole), 10.5 and 15 pt likewise; a 16 pt line is a pixel lower
  when the room's fraction of a pixel is in [.5, .88) and 20 pt a pixel higher at .417 --
  both swept by the box's height and its bottom inset, fifth of a twip by fifth. A first
  paragraph's last line, and every line but the last, are the body's (two lines by a break:
  the first on the model's pixel at every size, the second as the last).
* Unchanged, and confirmed: an offset is truncated to whole twips (single-EMU sweep about a
  twip's edge); a box's height is its extent in whole twips.

| First glyphs within half a pixel (pen x and baseline) | before | after |
| --- | --- | --- |
| `text-box-top-none` (1,710 lines) | x 1,706, baseline 1,178 | **1,710 / 1,710** |
| `text-box-top-14` | x 1,706, baseline 1,178 | **1,710 / 1,710** |
| `text-box-top-15` | x 1,706, baseline 1,211 | **1,710 / 1,710** |
| `text-box-none` (F.13, 115) | 100 | **113** |
| `text-box-15` | 101 | **113** |
| `drawing-none` / `-15` baselines (7,180 glyphs, F.3) | 7,121 / 6,661 | **7,180 / 7,180** |

*(The "before" of the new probe is the commit before, on the probe as committed.)*

**Recorded, not modelled** (pinned in the tests): centred text under `wrap="none"` (F.13's
(3), the two left of `make_text_box_probe.py`); a box flipped vertically or turned (180 and
10 degrees), whose text Word turns with it (a horizontal flip leaves it where it was, as
the model has it); `spAutoFit`, `normAutofit`, `vert` / `vert270` and `bodyPr@rot` (F.13's
(4)).

**The local corpora**: every text box line of their templates is on Word's pen x and
baseline (a quarter were a pixel off, and one on a half-pixel tie); the templates with text
boxes score better by SSIM, the colour histogram and the loss, every other local and
`filesamples` document is unchanged, and the visibility check finds no hidden glyph (their
numbers are kept with them).
The committed documents hold no text box: their output and fidelity are unchanged, and
every other probe keeps its recording.

### F.15 A table beside a drawing text wraps around — measured

F.8 laid a table beside a `wrapSquare`, `wrapTight` or `wrapThrough` drawing across the
whole column and warned (`wrap-table-not-modelled`). `tools/make_wrap_table_probe.py`
puts 150 tables after a paragraph that anchors a picture at the column's left or right
edge or in its middle (A4, the left margin off the pixel grid), in three settings (none,
14, 15), and `tools/read_wrap_table_probe.py` reads every cell's first glyph and the
picture from Word's PDF: tables of 3,000-10,000 twips; `jc` and `tblInd`; drawings that
end inside the first row, above it, or start inside the table; autofit, long cell text,
one column; `wrapTight`, `wrapThrough`, `wrapTopAndBottom`, `wrapText`, `distR`, `distB`;
the width about the room by 10 twips and then by single twips (borders of `w:sz` 4, 24
and none, an indent); the drawing's foot by single twips about the table's top.

**The rules** (`paginate.table_beside`, `table_fits`, `place_item`):

* **Which rows.** A row is beside a drawing when it reaches into the drawing's box
  widened by `distT` and `distB` (touching is not), as a line is (F.8). **In mode 15
  the whole table goes beside it** when any row is -- a drawing that starts in the
  table's third row moves the first two as well; **below mode 15 the rows from the first
  one beside it**, the rows above staying where they are (they end with the table's
  bottom edge, and the rest start a new edge).
* **Where.** In what the drawings leave of the column, the spans a line would have
  (`wrap.free_spans`: `distL` / `distR`, `wrapText`, a span under 360 twips not
  counted). The table is laid out as in a column that is that span: its indent, `jc`
  and the device-pixel rounding all from the span's left edge. In mode 15 the first span
  it fits in; below mode 15 the first span only (a 400-800 twip room left of the
  drawing holds a 3,000-twip table back, where mode 15 puts it right of the drawing).
* **What fits.** In mode 15 the table's right border's *outer* edge within the span:
  beside a room of 5,205 twips a table of 5,195 twips with `w:sz` 4 borders fits and
  5,196 does not; 5,145 with `w:sz` 24; 5,205 with none; 4,195 indented 1,000. Below
  mode 15 its last cell's *text* (the last grid line less the cell's right margin):
  5,313 fits and 5,314 does not, 108 twips of right margin past the room.
* **Otherwise it goes down** to the foot of the drawings it is beside and is tried there
  again: a table too wide for the room starts below the drawing, at its own place in the
  column.
* An autofit table is not narrowed to fit; `wrapTight` and `wrapThrough` keep a table
  off their box, as `wrapSquare` does (the probe's polygons are rectangles).

| Lines within half a pixel (every cell's first glyph) | before | after |
| --- | --- | --- |
| `wrap-table-none` (928 lines) | 534 | **909** |
| `wrap-table-14` | 534 | **909** |
| `wrap-table-15` | 479 | **909** |

The 19 left in each: a floating table (`w:tblpPr`, 16 lines) still stops the layout, and
a 2,999-twip table centred in the room draws its second cell's text a pixel right of
Word's (3 lines, a tie of the centring). Pictures 148 / 150 (the two past the stop).
`wrap-table-not-modelled` now warns only of what is left: a table that reaches into the
band of a `wrapTopAndBottom` drawing (one anchored below the table's top) is laid out where
it is, not below the drawing.
Pinned in `tests/test_wrap_table.py`.

### F.16 A drawing anchored in a table cell — measured

A floating drawing in a cell stopped the table (F.1, F.8). `tools/make_cell_anchor_probe.py`
anchors a picture in one cell of a two-by-two table of 4,000-twip columns, 102 cases in
three settings (none, 14, 15), `layoutInCell` on and off: every `positionH` frame that
means something in a cell (`column`, `character`, `margin`, `page`, `leftMargin`) by an
offset, `column` and `margin` by each alignment, the second column's cell; `positionV`
against `paragraph` (a cell's first and second paragraph, each alignment), `line`,
`margin` (an offset and each alignment), `page`, `topMargin`, the second row; the cell's
margins, a table indent, a centred table, space before; drawings taller than their row,
wider than their cell, above it; and, last, `wrapSquare`, `wrapTight` and
`wrapTopAndBottom` beside and above a cell's lines (with and without `distB`). Read by
`read_wrap_table_probe.py`.

**The rules** (`floating.in_cell`, `Frames.cell`, `layout._Placer._cell_anchors`):

* **In the cell** -- with `layoutInCell`, **in mode 15 whatever it says**, and below
  mode 15 without it where either axis is against `character` or `line` (then both: a
  `character` offset takes `paragraph` from the cell, a `line` offset `column`).
  Horizontally `column` and `margin` are the cell's text area -- its grid lines less its
  margins, exact -- and `page` and `leftMargin` start inside its left border; vertically
  `paragraph` is the cell paragraph's top (the body's rule: the first below the cell's
  top margin, above its own space before; the others where the one before ends after its
  space after), `line` the line's, and `margin`, `page` and `topMargin` start below the
  row's top border. **`margin` has no height in a cell**: `top`, `center` and `bottom`
  all put the drawing's top on the row's; against `paragraph` every alignment is its top,
  as in the body.
* **Against the page** -- below mode 15 without `layoutInCell`: every frame is the page's,
  but `paragraph` is the row's top (above its border), whichever of the cell's paragraphs
  anchors it.
* **A drawing text does not wrap around** (`wrapNone`) moves nothing: not the cell's
  text, not its row, however far past the cell it reaches.

| Lines within half a pixel | before | after |
| --- | --- | --- |
| `cell-anchor-none` (739 lines) | 1 | **449** |
| `cell-anchor-14` | 1 | **449** |
| `cell-anchor-15` (754) | 1 | **441** |

Pictures 74 / 102 in each, 0 before: every `wrapNone` case; the rest are past the stop.
**Recorded, not modelled**: text that wraps around a drawing in a cell stops the table (*since
F.17 only below mode 15 against the page*), as does a
`character` alignment or a frame the probe did not measure. In mode 15 a centred table's
second cell draws its text a pixel right of Word's (8 lines: the table's own centring, not
the drawing). The anchor probe's five pages with a picture in a cell (F.1) are now drawn,
every picture where Word drew it (`anchor-none` / `-15` pictures 143 -> 148 / 155, every
glyph on its baseline). Pinned in `tests/test_wrap_table.py`.

### F.17 Text around a drawing in a table cell — measured

F.16's probe ends with the drawings a cell's text wraps around: `wrapSquare`, `wrapTight`
and `wrapTopAndBottom` at the cell's left (and `wrapSquare` at its right) in a cell of
several lines, anchored in its first or second paragraph, 100,000 EMU down, taller than
the cell's text, and with a `distB` of 0.25 in -- each with `layoutInCell` on and off.

**The rules** (`table.wrap_cell`), for a drawing positioned in the cell (F.16):

* It is **positioned from the cell laid out without it**, as F.16 positions one text does
  not wrap around, and stays there.
* The cell's lines **go beside it or below it as the body's go beside and below a drawing
  on the page** (F.7, F.8) -- which lines, the segments and their pixel rounding, the
  narrowest segment, a line no segment takes going down to the drawing's foot -- with
  the cell's text area as their column. Every paragraph of the cell, the ones before the
  anchor's too.
* **The row reaches down to the drawing's foot** where the drawing is taller than the
  cell's text: in mode 15 to its `distB` below it (75 px lower with 0.25 in), below mode 15
  not (the same row with `distB` 0 or 0.25 in).
* Mode 15 does this whatever `layoutInCell` says (F.16). **Below mode 15 without
  `layoutInCell`** the drawing is the page's and moves the table: a `wrapSquare` or
  `wrapTopAndBottom` one moves the whole row below it, as a table too wide for the room
  beside a drawing goes below it (F.15); a `wrapTight` one leaves its own row where it is,
  the text over it, and moves a row below that it reaches into below it. Recorded, not
  modelled: it stops the table.

| Lines within half a pixel (`wrap` family) | before | after |
| --- | --- | --- |
| `cell-anchor-none` (291 lines) | 1 | **144** |
| `cell-anchor-14` | 1 | **144** |
| `cell-anchor-15` (306) | 1 | **306** |

Pictures 102 / 102 in mode 15 and 85 / 102 below it (74 before). What is left below mode
15: the `layoutInCell="0"` pages, where the table stops (136 lines, 14 pictures), and a
drawing at the top of a cell's *second* paragraph, which Word puts 0.05 px higher than
four stacked lines of 55.95 px do (594.282 against 594.328 px, as if each line were 268.5
twips): the first paragraph's last line, which ends where the drawing starts, is then
beside it (11 lines, 3 pictures) -- not settled. Pinned in `tests/test_wrap_table.py`.

### F.18 A floating table (`w:tblpPr`) — measured

A floating table stopped the layout (Tables, "What stays an obstacle"): the first table of
a filesamples document is one, so the layout of that document ended on its third page of
nine. `tools/make_float_table_probe.py` puts a floating table of two 1,500-twip columns
and three rows on each of 104 pages, a paragraph of text long enough to run beside it and
past it after it, in three settings (none, 14, 15), positioned as that document positions
its one (`vertAnchor="text"`, `tblpY="1"`, `rightFromText="187"`, `bottomFromText="72"`)
unless the case says otherwise: `horzAnchor` absent, `text`, `margin`, `page` with
`tblpX` 0, 1,000, 1,003, -500 and `tblpXSpec` left, centre, right; `vertAnchor` absent,
`text`, `margin`, `page` with `tblpY` 0, 1, 1,000, -300 and `tblpYSpec`; space before and
after around it; `rightFromText` 0-1,001, `leftFromText`, `topFromText`,
`bottomFromText`; borders of `w:sz` 0-24, cell margins 0 and 300, widths off the pixel
grid; rooms of 340-380 twips beside it; a table in the flow after it, indented, centred and
first-line-indented text. `tools/read_float_table_probe.py` reads every cell's first
glyph, every line of the text and the borders from Word's PDF.

**The rules** (`table.floating_of`, `table.floating_top`, `table.floating_edges`,
`paginate.table_wrap`):

* **Across, the table is laid out as a table with that indent or alignment is in the
  column** -- the frame being the column (`horzAnchor` absent, `text`, `margin`) or the
  page (`page`): `tblpX` is where the indent would be (below mode 15 the first cell's
  text, in mode 15 the left border's outer edge), **less one twip where it is positive**
  (tables at 567, 1,000 and 5,705-5,730 twips put their cells' text and the text beside
  them where 566, 999 and 5,704-5,729 would: 27 lines one pixel off without it, none with
  it); `tblpXSpec` `left`, `center` and `right` place it as `w:jc` would.
* **Down**, `vertAnchor="text"` is the top of the paragraph after the table -- where the
  paragraph before ends after its space after, above the paragraph's own space before,
  as a drawing's `paragraph` frame (F.1) -- plus `tblpY`; `margin` (and `vertAnchor`
  absent) is the top margin, `page` the page's top, plus `tblpY` or aligned by
  `tblpYSpec` `top`, `center` or `bottom` in them. Below mode 15 a table against the margin
  or the page at `tblpY` 0 goes where `text` would put it (floating, because of its
  `tblpX` or `tblpXSpec`); **with every position 0 and `vertAnchor` not `text` the table is
  not floating at all**: Word lays it out in the flow, in every mode.
* **It takes no room in the flow**, and the paragraphs on either side of it do not
  collapse their spacing: the one after it keeps its space before whole, after the one
  before's space after (a 240-twip after and a 120-twip before are 360 twips apart).
* **Text goes beside it as beside a `wrapSquare` drawing on both sides** (F.8) --
  which lines, the 360-twip narrowest segment, a line with no room going down past it,
  the lines before it in the document that it reaches up to, a table in the flow beside
  it (F.15) -- keeping off its box: its outer grid lines **and the exact half of its outer
  borders** (7.5 twips for `w:sz` 6), **or 15 twips where it has no border**, from its top
  to below its bottom border, widened by `topFromText` / `bottomFromText` and, across, by
  `leftFromText` / `rightFromText`, **each at least 10 twips** (0, 5 and 10 keep text 10
  twips off, 15 and more what they say).
* **The pixel rounding of the text beside it takes a half up** where a drawing's takes it
  down: below mode 15 edges at 642.5 and 1,112.5 px put the text at 643 and 1,113; in
  mode 15 `half_up(page) - half_up(text)` (an edge at 937.917 px leaves text starting at
  937.5 from the column's pixel where it is).

| Lines within half a pixel | before | after |
| --- | --- | --- |
| `float-table-none` (1,230 lines) | 0 (the layout stopped at the first table) | **1,226** |
| `float-table-14` | 0 | **1,226** |
| `float-table-15` (1,231) | 0 | **1,199** |
| `wrap-table-*`, `floating` family (F.15's three floating tables beside a drawing) | 2 / 18 | **18 / 18** |

**Recorded, not modelled** (pinned in `tests/test_float_table.py`): three ties of a cell's
text on half a pixel (722.5 px; in mode 15 a centred table's second cell at 1,277.5, 9
lines); a table centred in the margin, and in mode 15 two tables whose top is on a half
pixel, a pixel below Word's (8 lines); in mode 15 a table with no borders aligned right,
which Word draws 15 twips further left, as if a border were there (6 lines); and in mode
15 a table that reaches below the bottom margin, which Word splits across two pages: the
layout stops there. The borders are drawn a pixel from Word's on some vertical grid lines
of tables not at the column's edge (7,638 px of 1.3 million; Word seems to hang them from
the first cell's text, not from the column's edge). A floating table aligned `inside` or
`outside`, aligned against the text, or with its own `w:jc` or `w:tblInd` still stops the
layout.

The filesamples document is now laid out past its floating table -- every cell line and
every line beside it on Word's baseline, each text object starting where Word's does
(three cell lines drift up to 0.1 px inside a run: advances, not the table) -- to its next
table, whose width is stated in percent (stage 7): 17 of its paragraphs on Word's pages,
16 before.
The committed documents hold no floating table: their output and fidelity are unchanged.

### F.19 A drop cap (`w:framePr` `w:dropCap`) — measured

A drop cap is a paragraph in a frame -- `w:framePr` with `w:dropCap` `drop` or `margin`,
`w:lines`, `w:hSpace` -- holding the letter, before the paragraph it drops into; Word
writes the letter's size, an exact line and a lowering (`w:position`) on it when it makes
one. A frame stopped the layout (`layout-stopped:frame`). `make_drop_cap_probe.py` /
`read_drop_cap_probe.py`, 52 cases a page each, no `settings.xml`, mode 14 and mode 15,
three rounds: `drop` over 2 to 5 lines (sizes, exact lines and lowerings of the kind Word
writes), `D` and `W`; `w:lines` 3 with nothing else stated and with only the size; frame
lines far shorter and taller than three lines; `w:hSpace` 144 and 432; `margin`
anchored to the text, the margin and the page; text at 12 pt `auto` 276, a first-line
indent, right-aligned, justified; two letters; a paragraph shorter than the drop, then
another; space before the text and before the drop cap; `I M A O j T` at 30 and 70 pt with
the left margin off the pixel grid (1,442 twips) and on it (1,440); the drop cap's own
paragraph indented (a first line of 432, as a filesamples document's Normal style has
it; a left indent of 360).

**The rules** (`paginate.DropCap`, `drop_cap_of`, `page_wraps`,
`layout._Placer._drop_cap`):

* **The frame takes no room**: the paragraph it drops into starts where it would have
  without it. Its top is **that paragraph's top** -- where the paragraph before ends after
  its own space after, not after this one's space before -- and it is as tall as its own
  paragraph's space before and lines, **as written**: Word does not size the letter from
  `w:lines` (with nothing else stated the letter is the text's size, and drops one line).
* **Its letter is set in its frame as a paragraph's line** -- at the column's start (its
  own indent counts), its baseline from the foot of its line: the line's bottom rounded,
  less what is below the text rounded (exact lines of 400 to 1,600 twips, single lines of
  11 to 50 pt, whatever their fraction of a pixel), then lowered by `w:position` as a
  run is.
* **A line is beside the frame when its top is above the frame's foot** (4 lines beside a
  three-line cap whose frame reaches 0.07 px into the fourth), in this paragraph and the
  ones after it; `w:lines` counts for nothing more.
* **The text beside it starts after it**: from the column's start (a whole pixel), the
  frame's width -- its widest line's reach, its indent included -- `w:hSpace`, and 103
  layout units (half a twip, rounded up), `S`. In mode 15 the text is set from `S` and moved
  a pixel right where the left margin's exact edge is further into its pixel than the
  width is into its own (36 of 36 widths, both margins). Below mode 15 `S` is on the
  page's nearest whole twip and the text is moved from there to that pixel, rounded half
  down -- so a right-aligned line beside it ends two thirds of a pixel past the column's
  edge, as Word's does. Alignment, justification and a first-line indent are the
  segment's, as beside a drawing (F.8).
* **`margin`** anchored to the text or the margin is drawn as `drop` is; anchored to the
  page, the letter stands its width left of the column, rounded, and the text is not
  moved.

| Lines within half a pixel (433 each) | before | after |
| --- | --- | --- |
| `drop-cap-none`, `-14`, `-15` | 1 (the layout stopped at the first frame) | **433** |

Not measured, and left: a frame that is not a drop cap (still a stop), a drop cap before a
table or at a document's end (a stop), `w:vSpace`, a drop cap whose frame reaches past the
page's foot, a `margin` cap's position against the page when the margin is not whole, and
tab stops beside a frame. Pinned in `tests/test_drop_cap.py`.

With this, 5.18 and 4.12, the filesamples document Word sets on nine pages is laid out on
nine, to its end: every glyph the model draws matched to one of Word's on the same page
(7,450 of 7,450; Word draws 9,631), its drop cap's letter and the three lines beside it
where Word draws them, its picture bullet on Word's box, and its endnote after the last
paragraph under the separator. What the model does not draw there: the footnote (its
reference's number is not drawn either, and the rest of its line is 19 px left of Word's)
and the tab leaders. *Since 4.13 the footnote is drawn (7,655 of 7,655 glyphs matched), and
since 5.19 the leaders: every glyph Word draws is drawn, 9,631 of 9,631, on Word's page,
baseline, face and size.*

### F.20 Charts (`c:chart`) — measured

A chart in a `w:drawing` -- inline or floating, in the body, a header, a footer, a text box
or a group's `wpg:graphicFrame` -- was a placeholder of its extent (`drawing-not-drawn`).
It is drawn now by the code pptx2svg draws a slide's charts with, which moved to
`ooxml-common` (0.4: `chart.read`, `chart.layout`, the scene renderers): `docx2svg.chart`
relates the chart part from the part the drawing sits in, reads it and its cached values
(`c:numCache` / `c:strCache`, which is what Word draws; the embedded workbook is not
read), and supplies what is Word's -- colours through the theme and
`w:clrSchemeMapping` (or the chart's `c:clrMapOvr`) under Word's colour rules (F.12),
`+mn-lt` / `+mj-lt` as the document theme's faces, a title's text, and Word's line box.
The chart is laid out in its frame -- the room an inline picture takes in its line (4.7),
the place a floating anchor gives a shape (F.1-F.8) -- and drawn there.

**Measured before it was drawn**, because every constant of that layout was measured on
PowerPoint: `make_chart_probe.py` / `read_chart_probe.py`, `tests/test_chart.py`
(`tests/fixtures/chart-observations.json`). 58 charts written here, a page each (no
embedded workbook, whole-number values: Word writes a fraction with the host's decimal
separator, `0,5` on this Mac), in three documents -- no `settings.xml`, mode 14 and mode
15, which **draw every chart alike** -- read off Word's PDF by PyMuPDF: every span of the
chart's text (string, face, size, advance box, baseline), every filled shape and every
stroked segment, against the same read off the chart's SVG fragment. The families:
nothing stated; `c:txPr` at 6 to 18 pt; titles of 8 to 36 pt in Arial, Aptos and
`+mj-lt`, with and without an `a:defRPr`; the legend at every position, with long names;
clustered, stacked and 100% columns and bars, lines, an area, a pie, a doughnut, a
scatter, a radar; the chart's and the plot area's fill and line stated, alone and as
none; series fills and lines; data labels; inline, floating in front, wrapped, and off
the twip grid.

**What is Word's, where it is not pptx2svg's** -- each a field of `ooxml-common`'s
`chart.rules.ChartRules` (or `drawingml.rules.DrawingRules`), `WORD` beside the
`POWERPOINT` default, which keeps pptx2svg's output byte for byte:

* **The chart space** that states no fill is **white**; one that states no line is
  outlined in **`898989` at 0.5 pt** -- each on its own (a fill alone keeps the line, a
  line alone the white). An explicit `a:noFill` draws nothing. pptx2svg: transparent.
* **The plot area** that states no fill is **white** -- a radar's is the square round its
  web, a pie has none -- and its own `a:ln` is drawn (1 and 0.75 pt, where stated).
* **A title** whose paragraph states **no `a:defRPr`** is **Arial 18 pt, not bold**,
  whatever the theme, the chart's `c:txPr` or the document's faces (Arial, Times New
  Roman and Verdana documents under Georgia and Aptos themes); one that states an
  `a:defRPr`, even an empty one, takes the chart's text -- the theme's minor face at
  `c:txPr`'s size or 10 pt -- **bold** (`docx2svg.chart.title_text`). Its **band** is its
  line pitch **plus 9.0 pt** (to 0.003 pt, Arial and Aptos at 8 to 36 pt) and its
  **baseline 7.5 pt + 0.9412 em** below the frame's top, whatever the face (every Aptos
  title on Word's device pixel; Arial at 28 pt a pixel higher).
* **Axis and legend text** is the theme's minor face at `c:txPr`'s size or 10 pt
  (Georgia, Aptos and Calibri themes), drawn at its size rounded to the device pixel
  (10 pt drawn 10.08).
* **A side legend** keeps **13.25 pt and half its key** from the plot and **10.13 pt**
  from the frame's edge, at 6 to 18 pt (pptx2svg's 1.6 and 1.01 em, measured at 10 pt,
  where the two agree to 0.03 pt), and **stands against the frame**: where the plot gives
  up room to a label centred on its edge -- an area's, a scatter's, a horizontal bar
  chart's last value -- the plot shrinks and the legend stays. A left legend's key is
  8.25 pt and half the key in.
* **A top legend** stands **under the title**; a **top right** one is a column at the
  right, from where a top legend's first row stands, the plot giving up its rows at the
  top and 6 pt at the bottom.
* **Legend order** is **reversed** for clustered horizontal bars (beside or below) and
  for stacked and 100% columns beside a side legend -- the order the series stand in,
  bottom to top -- and series order otherwise (stacked bars, stacked columns over a legend
  below). **Clustered bars put the first series lowest** in its group.
* **Lines**: a custom path's outline is drawn at its stated width (a line series 1.5 pt,
  2.25 pt where stated without a colour, radar rings 0.5 pt -- pptx2svg's path `scale()`
  draws them a third wider); a legend key is outlined with its series' line.

| `chart-none`, `-14`, `-15` (each) | text | fills | strokes |
| --- | --- | --- | --- |
| before (a placeholder) | 0 / 804 | 0 / 665 | 0 / 1,576 |
| the shared code, PowerPoint's rules | 569 / 804 | 331 / 890 | 961 / 1,955 |
| Word's rules | 790 / 804 | **664 / 666** | **1,574 / 1,578** |
| ooxml-common 0.7 (the radar's category labels) | **794 / 804** | 664 / 666 | 1,574 / 1,578 |

A family's count is what agreed within 0.5 pt (face and device size too, for text), out
of Word's items and the model's left over. **What is left**: the radar's five value labels
(0.51 to 0.65 pt: their offset is not measured beyond PowerPoint's; its four category
labels, up to 2.4 pt out until 0.7 placed them 4% of the radius off the rim, now agree),
five labels of charts whose text is 6, 14 or
18 pt (two value labels and three legend names, 0.55 to 0.6 pt), and one scatter marker
Word draws 0.6 pt off its point. Pinned in `tests/test_chart.py`.

**Not measured, and warned or left as pptx2svg draws it**: a chart's embedded workbook
(only the caches are read: `c:externalData` is recorded, not opened), 3-D charts (drawn
flat, `chart-3d-flattened`), `c:roundedCorners` (every chart here states `0`, as Word
writes it), a chart style part (`c14:style`, `cs:chartStyle`), number formats in the host's
locale, and the chart types the layout does not draw (`chart-unsupported-type`, a
placeholder). The committed documents and the local corpora hold no chart (the one
`filesamples` "chart" is a picture): their output and fidelity are unchanged.

**Chart labels, measured later** (ooxml-common 0.7, its `tests/test_chart_labels.py`): on
the charts docx-agent's `insert_chart` writes, Word draws tick, category and legend text in
the `c:txPr` fill they state (`tx1` at 65%, `595959` on Office's theme), each element's own
fill over the chart space's, and plain `tx1` where no `c:txPr` states one -- as PowerPoint
does; data labels in their own number format (`"€"#,##0.0"m"` reads `€41,2m`) unless it is
source-linked; and a radar's category labels 4% of the radius off their vertex, centred on
the anchor on a sloping spoke. docx-agent's w14 radar had looked "slightly larger" than
Word's: its labels were the same size (9.12 pt on Word's device pixel, widths within
0.3 pt) round the same web (its value labels 0 to 5 span 94.99 pt, Word's 95.04), but drawn black where Word drew
them `595959`, and the four on sloping spokes hung up to 5.5 pt further out. Every label
of eight radars of 3 to 8 categories in Aptos 9 and 14 pt now lands within 0.13 pt across
and 0.22 pt down of Word's PDF.

### F.21 SmartArt (`dgm:relIds`) — measured

A SmartArt diagram is data and a layout definition, and Word lays it out whenever it opens
a document and writes the result as a drawing part (`dsp:drawing`) on every save -- so a
document Word saved holds Word's own layout. `docx2svg.diagram` draws that cache: found from
the data model's `dsp:dataModelExt` through the part the drawing sits in
(`ooxml_common.drawingml.diagram`), read as a shape tree, each shape's fill, line and
effects through the theme and its `dsp:style` references, its text in the face its
`a:fontRef` names (the theme's minor or major) and the reference's colour, drawn by the
scene renderers in the drawing's box. A diagram with no cache, or an empty one, is a
placeholder and `diagram-no-cached-drawing`.

**Measured**: `make_smartart_probe.py` / `read_smartart_probe.py`, `tests/test_smartart.py`
(`tests/fixtures/smartart-observations.json`). Word lays out only the Basic Block List from
a definition it is given empty -- every other layout name, built in or not, it replaces
with that one -- so the probe's diagrams are block lists: three short names, three shorter
words (Word sets them at 47 pt), five with a long name (shrunk to 19 pt and wrapped on
three lines), two items with two bulleted items each, and one floating; with a colour
definition of the probe's own. The cache each carries is Word's: `--caches` has Word save
the document again (`tools/oracle.py`'s `resave`, `word_save_docx.applescript`) and reads
each shape's place, size and text back into the recording as numbers, and every run checks
that Word still writes it (it does, but for the frames' effect extents, which Word sets
apart in each setting and which move only what follows a diagram on its line).

**What is Word's**: **DrawingML text sits in the face's own line box** -- Aptos's 1.2207 em,
where PowerPoint's (and the shared measurer's) is 1.2 em for every face -- and **its first
baseline is that box, spaced, less the face's descent, which the spacing does not scale**;
the next line steps on by that descent, the gap, and its own box less its descent
(`DrawingRules.first_baseline`, and `docx2svg.chart.WordTextMeasurer`). Three 19 pt lines
at 90% step 20.88 pt (0.9 x 1.2207 em; 1.2 em gives 20.52), and every baseline of the
probe lands on Word's device pixel. A chart's one-line labels are placed for the same rule,
so they do not move.

| `smartart-none`, `-14`, `-15` (each) | text | fills | strokes |
| --- | --- | --- | --- |
| before (a placeholder) | 0 / 26 | 0 / 16 | 0 / 64 |
| the shared code, PowerPoint's rules | 3 / 26 | 16 / 16 | 64 / 64 |
| Word's rules | **26 / 26** | **16 / 16** | **64 / 64** |

**Not measured**: other layouts' shapes (arrows, connectors, cycles' arcs: drawn by the same
renderers, not compared with Word), a diagram's text above 100% spacing, pictures in a
diagram (warned, `diagram-not-drawn:image`), and Word laying out a diagram whose cache an
older version wrote (the cache is drawn as it is).

### F.22 A text box wrapped square beside several tables — isolated, not reproduced

docx-agent saw Word fit three lines more on a page than docx2svg where a square-wrapped
text box was anchored in the paragraph before three tables (one with spans, one holding a
nested table), each alone agreeing (its E5). `make_wrap_tables_probe.py` /
`read_wrap_tables_probe.py` (`tests/test_wrap_tables.py`) isolate it: the anchoring
paragraph (a text box at the column's right, 2,000,000 x 3,000,000 EMU), then a plain
table, one with a cell over two columns and one merged down, and one with a table nested
in a cell -- an empty paragraph between each, or nothing between them -- each alone, and
the three too wide to go beside the box; then one-line paragraphs on to the next page. No
`settings.xml`, mode 12 and mode 15. **Every page breaks where Word breaks it**, in every
case and setting: 2,333 / 2,333 glyphs on Word's page at Word's x. docx-agent's own
operations replayed on `layout-sweep.docx` (outside this repository; the box in the first
paragraph, its tables after it, with and without its content controls, shape and picture)
break every page where Word does too, before this round of changes and after it. So the
three lines were not reproduced, and nothing was changed for them.

What the probe found instead: **two tables with nothing between them** (`adjacent`) --
Word draws the second, and everything after it, 2 px higher than docx2svg drew it (243
glyphs, every one at Word's x and on Word's page): the edge between them is one border, not
two. *Drawn so since F.24: 2,333 / 2,333 on Word's baseline.* In the replayed documents a first line holding
content controls (a check box's symbol among them) is drawn 1 px high, the page's lines
with it.

### F.23 Chart titles, axis titles and short charts — measured

docx-agent's chart work (its roadmap, "Charts and SmartArt", proposals 1 to 6) found what
F.20 had not reached: a chart in a footnote related from the main part
(`chart-unreadable`), a chart inline in a text box neither drawn nor warned, no axis
titles, no title for a `c:title` with no text, an empty `a:defRPr` drawn at 10 pt where
Word drew 18, and a chart 63 pt high whose labels stood apart from Word's.

**A chart in a note, or in a text box.** A note's drawing is related from the notes part
(`chart.Frames.owner`, from its path, as `paths.path_part` reads it), and a text box's
inline drawing that is not a picture -- a chart, a shape -- is drawn with the box, in its
line's place (it was drawn on a scratch page and dropped). Against Word's PDF of
docx-agent's `chart-places.docx` every text of its five charts -- the header's, the note's,
the text box's, the group's and the floating one -- is Word's within half a point, 45 / 45
(with F.24's pagination).

**Measured**: `make_chart_text_probe.py` / `read_chart_text_probe.py`,
`tests/test_chart_text.py` (`tests/fixtures/chart-text-observations.json`): 80 charts, a
page each, in three documents (no settings part, mode 14, mode 15, which draw every one
alike). The reader is `read_chart_probe.py`'s, with turned text compared by the midpoint of
its baseline and its angle. Families: titles over every `c:txPr` size and with Word 365's
chart style and without; titles with no text over one series, two, a pie, and none at all;
axis titles in the empty form and Office's chart-style form, unturned, at 10 to 18 pt, on
columns, bars and lines, beside a legend at every side, long, over wide labels; charts as
Word 365 writes them 30 to 126 pt high at 8, 10 and 14 pt.

**What Word does** -- the layout's are `ooxml-common`'s `ChartRules.WORD` fields
(`auto_title`, `axis_titles`, `short_plot`, 0.4.4); the faces and sizes are
`docx2svg.chart`'s:

* **A title that takes the chart's text** -- an empty `a:defRPr`, or Word's automatic
  title -- is the minor face, bold, at **1.2 times `c:txPr`'s size** (8 gives 9.6, 10 gives
  12, 12 gives 14.4, 18 gives 21.6); where `c:txPr` states none, **18 pt in a chart that
  states `c14:style`** (Word 365's chart style, in `mc:AlternateContent`; `c:style` alone
  does not count, and 101 counts as 102) and 10 pt in one that does not. F.20's
  "`c:txPr`'s size or 10 pt" was this rule's one case its probe had: no `c:txPr`, no style.
  docx-agent's 18 pt was the style's.
* **A `c:title` with no text** shows the **sole series' name** (a pie's too); over several
  series, or under `c:autoTitleDeleted` `1`, Word's own "Chart Title" in its interface's
  language (`Grafiektitel` on this Mac), which a document does not say: its band is kept
  and nothing drawn in it (`chart-title-not-drawn`). No `c:title` but `c:autoTitleDeleted`
  stated `0`: the sole series' name; not stated, nothing. A title with text is drawn under
  `c:autoTitleDeleted` `1` too.
* **An axis title** is the chart's text, bold, at `c:txPr`'s size or 10 pt (the empty form);
  Office's form as it states (10 pt, `595959`, not bold); no `a:defRPr`, Arial 18 pt -- the
  title's rules, but for the 1.2 and the style. At the left it reads **upwards** whatever
  its `a:bodyPr` states but an explicit `rot="0"` (not laid out: Word sets it across, and
  the plot gives it its width; `chart-axis-title-not-drawn`); at the bottom it reads across.
  Each takes its **line pitch plus 9.0 pt** off the plot (21.21 pt for 10 pt Aptos, 29.70
  for 18 pt Arial, to the device pixel), its line box **12.5 pt** in from the frame -- or
  from a legend on that side -- the gap away from the plot, centred on the plot. A bottom
  title's band comes off what the value axis divides (its labels stay Word's).
* **A plot too short for its axis**: where a vertical chart's plot would be shorter than
  **1.1 em of its value labels**, the band under it -- the gap above a bottom legend, or the
  frame's inset -- gives up **half the shortfall** (slopes 0.48 to 0.51 at every height; the
  em fitted at 8, 10 and 14 pt to half a point), down to no plot at all. A 216 x 63 pt
  chart's labels came 3.9 pt above Word's. Word does draw the `0` docx-agent missed there:
  its chart had a plot of no height, its `0` and top label on one line, which its reader
  took for one text.
* **Lines under Word 365's chart style**: every axis, tick mark and gridline that states no
  colour is `898989` (0.5 pt), where a chart without `c14:style` draws them black
  (`ChartStyle.line_color`).

| `chart-text-none`, `-14`, `-15` (each) | text | fills | strokes |
| --- | --- | --- | --- |
| before | 343 / 877 | 250 / 1,234 | 522 / 2,396 |
| after | **862 / 877** | **723 / 761** | **1,437 / 1,478** |

What is left is pinned by case: Word's "Chart Title" (three texts), the unturned value-axis
title (twelve texts, eighteen fills, 38 strokes: the whole plot), a `c14:style` of 101's
series colours (twenty fills, not modelled), and a plot of no height, where Word draws a
gridline over its axis (one stroke, three charts). docx-agent's `charts.docx` scores SSIM
0.9704 against 0.9438 before (the style's grey lines), `chart-places.docx` 0.9871 on Word's
two pages (one page before, 0.3198), `smartart.docx` as before; the committed documents
hold no chart, and their fidelity is as recorded.

**Not measured**: a right or top axis' title, an axis title in a scatter, radar, pie of a
pie, stock or 3-D chart (`chart-axis-title-not-drawn`), a title over several lines, and a
short horizontal bar chart.

### F.24 Drawings stacked down a page, a note's spacing, two tables touching — measured

docx-agent's `chart-places.docx` stacks floating drawings wrapped top and bottom, each at
offset 0 in its own paragraph, over a footnote holding a chart: Word draws them on two
pages, docx2svg drew them on one, overlapping (its proposal 7). Four rules, each probed on
its own.

**Each drawing positioned below what the drawings before it pushed down.**
`make_wrap_stack_probe.py` (read by `read_anchor_probe.py`, `tests/test_anchors.py`):
pictures in consecutive paragraphs, with a paragraph between, at offsets, three where the
third's paragraph cannot stay on the page, the same anchored after a line, one positioned
against the margin before one against its paragraph, two in one paragraph; no settings and
mode 15. F.7's rule holds for the first -- positioned from the page laid out without it --
and **each after it is positioned from the page laid out with the bands of those before
it**, not its own (`paginate._stacked_bands`); one whose anchor those bands push off the
page goes on with its paragraph, positioned there. 16 / 16 pictures and 2,561 / 2,561
glyphs on Word's baseline in both documents, where 3 / 20 and 379 were.

**A note's spacing** (`make_footnote_draw_probe.py`, new `after` family, four settings):
**the last note's space after is in the note stack** (160 and 240 twips put the notes 33
and 50 px higher; 4.13 had measured spacing only on a note followed by another), and **the
first note's space before stands over the separator**, not under it (120 and 240 twips put
the separator 25 and 50 px lower, the note where it was). 22 / 22 lines and every
separator, where 15 and none were. And a note line holding a picture or a drawing takes
its height in the room the paginator keeps, as the drawn note does (it took a text line's).

**A line's space after above the notes, in mode 15** (`make_note_after_probe.py` /
`read_note_after_probe.py`, `tests/test_footnote_draw.py`): a candidate line with 240
twips after, a spacer swept by 20 twips under a line carrying a reference -- in mode 15 the
candidate leaves the page **240 twips sooner** than one with none (to the step); with no
settings and in mode 14 they leave together. So in mode 15 a paragraph's last line keeps
its space after above the page's footnotes, where it need not above the bottom margin
(4.1). 40 / 40 cases in each of six documents.

**Two tables touching** (F.22's `adjacent`): the second starts the narrower of the first's
bottom border and its own top border higher -- one edge (`paginate.table_gap_px`): 2,333 /
2,333 glyphs on Word's baseline, where 2,090 were.

`chart-places.docx` now breaks where Word breaks it, every drawing where Word drew it. The
committed documents hold none of these cases: their fidelity is as recorded.

**Not measured**: the stacked drawings in a section of several columns (positioned as
before), a band pushing its anchor's own line to the next page where the drawing is
positioned against the page or the margin, a note's space before over the continuation
separator, and two tables touching with borders of different widths (the narrower taken).

---

## Phase 6 — Everything else, sized

| Item | Effort | Notes |
| --- | --- | --- |
| ~~Style inheritance (`w:style` `basedOn` chain, `docDefaults`, direct formatting)~~ | M | **Done** — see *Style inheritance — measured*. "Solved in the field" turned out to mean "specified"; Word departs from ECMA-376 on toggles, numbering precedence, complex-script classification and East Asian / complex-script theme fonts. |
| ~~Numbering and lists~~ | M | **Done for counting** — see *Numbering and lists — measured*: instances share their definition's count (one `w:nsid`, one definition), `w:startOverride` restarts at the instance's first item, a whole `w:lvl` override, `w:lvlRestart`, shallower levels taken at their start, `w:isLgl`; text boxes, notes and headers counted apart from the body. |
| Tables | L | Word's auto-fit algorithm is a genuine measurement problem; fixed layout is not. *Stages 1 (the grid), 2 (a cell's content; finding 6), 3 (row heights), 4 (merged cells), 5 (borders and shading) and 6 (across pages) measured; the committed documents' tables are laid out, every glyph exact. Autofit is measured and modelled (stages 7, 7a and 7b): a column from its cells' widths or content, the room shared by widest less narrowest content, a table width in dxa or percent; a cell across columns wider than they are still stops the layout: see* Tables — measured. |
| ~~Headers and footers~~ | S–M | **Done** — see *Headers, footers and fields — measured*: drawn where Word draws them, the story each page shows (first, even, inherited, blank pages), every probe glyph exact; a floating drawing in one is warned, not drawn. |
| ~~Fields~~ | M | **Done** — `PAGE`, `NUMPAGES`, `SECTIONPAGES` computed from the layout's pages (formats, switches, restarts), the rest drawn as cached; instructions never text. Word recomputes `DATE`, `TIME`, `REF`, `PAGEREF`, `SEQ`, `IF` on export and writes number words in the host's language: warned. |
| Footnotes and endnotes | L | Reserve the space before paginating — see Phase 4. *Footnotes' room measured (4.3); endnotes numbered, placed and drawn (4.12); footnotes numbered and drawn at the page foot (4.13).* |
| ~~Text columns~~ | M–L | **Done** — see Phase 4, 4.14: geometry, flow, breaks, the separator and balancing measured and laid out; footnotes in columns measured and laid out (4.15) but where a continuous break joins the section to another; drop caps, and endnotes below mode 15, in columns (4.16); floating drawings positioned against their column (4.17). |
| `w:drawing` (DrawingML) | M | Mostly free via Phase 1's shared package; the *anchoring* is not. *Shapes, theme colours, groups and text boxes: see* Floating drawings — measured *(F.3); fills, gradients, patterns, outlines, dashes, arrowheads, effects, theme style references and WordArt's text fill and outline drawn by `ooxml-common`'s renderers under Word's measured rules (F.12); charts by its chart layout under Word's measured chart rules (F.20) and SmartArt from the drawing Word caches (F.21); picture fills and compound lines warned.* |
| Floating objects and text wrap | XL | Word's wrap geometry is not CSS floats. The field approximates this universally. *Stages 1 (anchors that do not move text), 2 (in headers and footers), 3 (`wrapTopAndBottom`) and 4 (text beside a drawing: `wrapSquare`, `wrapTight`, `wrapThrough`), a table beside a drawing (F.15), a drawing in a table cell (F.16, F.17), a floating table (F.18) and a drop cap (F.19) done; other frames stop the layout; a header's drawing the body wraps around is not: see* Floating drawings — measured. |
| `w:compat` | M–L | Not a feature; a bundle of switches. Measure which ones move this oracle. |
| East Asian typography, justification, hyphenation | L each | Each its own measurement; do not combine. *Justification in mode 15 measured (3.8); automatic hyphenation measured for English, Dutch, German and French, and at a page's end (3.9).* |
| ~~Tracked changes, content controls, comments~~ | M | **Done for the final view** — see *Revisions — measured*: moves, deleted paragraph marks, content controls at every level, property changes, rows and cells, as Word's *No Markup* view draws them; comments take no room. No markup view is drawn, and none is planned. |

### Revisions — measured

docx-agent's roadmap found three gaps in the final view of tracked changes the renderer
drew: text in `w:moveTo` was dropped, a paragraph whose mark is deleted was not joined to
the next, and an inline content control's text was not drawn. `tools/make_revision_probe.py`
measures what Word draws instead: 70 cases, one a page, in seven families (moves; deleted
marks; content controls; property-change records; rows and cells; comments; section
ends), with no compatibility mode, in mode 14 and in mode 15; `tools/read_revision_probe.py`
scores them (`tests/fixtures/revision-observations.json`, `tests/test_revisions.py`). A
case agrees when every line of its pages is Word's: the same glyphs, page, baseline below
the case's label and left edge, to the device pixel.

**R.0 The oracle exports the final view.** See Phase 0, 0.1: Word's default export prints
the markup; the probe is exported in the *No Markup* view, and once more with every
revision accepted first (the script's `accept`, also with markup hidden, since a
comment's pane is printed otherwise). **The final view is not always what accepting
makes**: the two differ in 8 cases in mode 15 and 17 below it (R.3, R.4, R.6), and the
renderer draws the final view.

**R.1 Moves.** The final view draws `w:moveTo`'s content where it stands and nothing of
`w:moveFrom`'s, with range markers in paragraphs or between them, or none; a paragraph
whose mark is `w:moveFrom` is a paragraph whose mark is deleted (R.2).

**R.2 A deleted paragraph mark joins its paragraph to the next.** A paragraph whose mark
is deleted (`w:pPr/w:rPr/w:del`, or moved away) is drawn as one paragraph with the next
paragraph of its container — the body, a cell, a block-level control's content, which is
the container's — **under the next paragraph's properties**: its style (the first
paragraph's text loses the first's paragraph style's run properties), numbering,
indents, alignment, spacing and mark. A chain of deleted marks joins into its last
paragraph, under the last's properties, in every mode. Measured both ways: the joined
paragraph under the *first* paragraph's properties scores 11 / 32 of the family (and 4 /
6 moves, 2 / 4 section ends). The one paragraph is kept by its first paragraph's path.

**R.3 With no paragraph to join.** Before a table, at the end of a cell and at the end of
the body, the paragraph keeps its own text where it stands, but under the paragraph
properties of **the next paragraph in the story** — the table's first cell's first
paragraph, the next cell's, none at the body's end (then no properties of its own: the
default paragraph style). One whose text is all deleted, before a table, is not drawn at
all. Accepting differs here: Word moves the text into the table's first cell, and with a
cell's last mark deleted moves the cell's first paragraph out of the table.

**R.4 Below mode 15, a join at a line's start draws the line lower.** In the final view
below mode 15 (none and 14 alike), a line other than the paragraph's first that begins a
joined paragraph's text — because the text before it filled its line or ended in a
`w:br` — is drawn lower by the paragraph's space before, and every later line of the
paragraph with it, cumulatively (25 px for 120 twips, 100 px for 480). What follows the
paragraph moves down by **one fewer** of them: by nothing for one such join, so the last
line overlaps what follows; by one space before for two, by two for three. A join
mid-line moves nothing. Mode 15, and accepting, draw no line lower. Without the rule the
family scores 23 / 32 below mode 15.

**R.5 Content controls.** Run-level `w:sdt` draws its `w:sdtContent` as ordinary runs —
nested, inside a hyperlink, holding revisions; block-level `w:sdt` draws its paragraphs
and tables as the container's; cell-level and row-level `w:sdt` are drawn as the cell's
content and the row. A placeholder (`w:showingPlcHdr`) is drawn as its runs say, as any
text: here in the `PlaceholderText` character style, at its place in the line.

**R.6 Rows, cells, property changes, comments.** A row whose `w:trPr` holds `w:del` is not
drawn, whether or not its content is deleted, and a table all of whose rows are deleted
is not drawn at all; an inserted row is. `w:cellIns` and `w:cellDel` change nothing in the
final view (a deleted cell is drawn, empty); nor does `w:cellMerge`, whose `w:vMerge` is
ignored: the cell's own `w:vMerge` decides (accepting applies it). `w:rPrChange` (on a
run and on a mark), `w:pPrChange` (a list's too), `w:tblPrChange`, `w:trPrChange`,
`w:tcPrChange` and `w:sectPrChange` change nothing: the current properties are drawn. A
deleted mark on a section's last paragraph removes the section break; the two sections
are one, under the following section's properties. Comments are not drawn: a commented
range, its reference at 8 pt or 48 pt, between words or alone in a paragraph, takes no
room.

**Scores.** Before: 15 / 70 cases in each setting (moves 0 / 6, marks 1 / 32, controls 1 /
11, properties 4 / 5, rows 5 / 9, comments 3 / 3, section ends 1 / 4), 137 / 291 lines,
72 pages for Word's 70. After: **70 / 70 cases, 291 / 291 lines, every glyph of 6,073 at
Word's position (two object starts in mode 15 0.001 px off, as elsewhere), and every rule
pixel**, in all three settings. The committed documents have no revisions, controls or
comments: the fidelity baselines do not move.

**Not handled.** What accepting makes, and any markup view: not drawn, by decision. The
paths of a deleted row's neighbours' cells keep their place among the rows, but a
table's border rules are named by row index after deletion. Not measured: a deleted mark
before a table whose paragraph also ends a section; a run-level control's `w:rPr` (its
runs' own are drawn); a dropped line falling past the page's foot; `w:numberingChange`,
`w:tblGridChange` and `w:tblPrExChange` (read as nothing, as the others measured); a
deleted mark in a header, a footer, a note or a text box (handled as in the body,
unmeasured).

### Numbering and lists — measured

docx-agent found docx2svg drawing `1.` where Word draws `7.` (a list restarted at 7, and
one at 5): `w:startOverride` was ignored, and items were counted per `w:num`, where
docx-agent's model counts per `w:abstractNum`. `tools/make_numbering_probe.py` measures
how Word counts: 43 cases, one a page, each with abstract definitions of its own, in five
families (overrides; instances; restarts; legal numbering; stories), with no
compatibility mode, in mode 14 and in mode 15; `tools/read_numbering_probe.py` scores
them (`tests/fixtures/numbering-observations.json`, `tests/test_numbering.py`). A list
item's text names its instance and level, so a line of Word's PDF reads `7.B0`. A case
agrees when every line of its pages is Word's: the same glyphs, labels included, at the
same baseline below the case's label and the same left edge, to the device pixel.
**Word counts identically in all three settings.**

**N.1 Instances share their definition's counts.** Two `w:num` over one `w:abstractNum`
are one list: `A A B B A` draws `1 2 3 4 5`, interleaved too, and a level-0 item of one
restarts the other's level 1. Two definitions alike are two lists (`1 2 1 2 3`). **Two
definitions with one `w:nsid` are one**: the first in the part, its levels and its counts
(`1 2 3 4 5` in decimal, though the second's level 0 says upperRoman). A
`w:numStyleLink` definition takes the levels of the definition under the numbering
style's `w:numId`, but counts as a list of its own (`1 2 1 2 3`). `numId` 0 between items
interrupts nothing.

**N.2 `w:lvlOverride`.** A `w:startOverride` restarts the shared count **at the
instance's first item at that level, and only there**: `B B B` with 7 draws `7 8 9`;
`A A B B A A` with 5 on B draws `1 2 5 6 7 8` (the other instance continues from the
restart); `B A B A B` draws `5 6 7 8 9`; two overridden instances interleaved, `5 10 11
12`. It is also the level's start in that instance: restarted by a shallower item, an
override of 3 at level 1 draws `c` again; another instance's item there restarts at its
own start (`a`). An item first used deeper than its override (`B1 B0`) still restarts
at `B0` (`a 5`), so the trigger is per level. A whole `w:lvl` in the override replaces
the level (format, text, start) **and restarts nothing**: `A A B B A` draws `1 2 III IV
5`. With a `w:startOverride` too, the restart is at **the `w:lvl`'s `w:start`**, not the
`w:startOverride` (4 drawn, 9 stated: `1 2 IV V 6`).

**N.3 Restarts and labels.** An item restarts each deeper level its `w:lvlRestart`
allows: unstated, every deeper one; `0`, none (`1 a b 2 c`); `1` on level 2, only after
level 0 (`b iii`, then `2 a i`). An item **sets every shallower level that has no count
since it restarted to its start, as if it had an item there**: `1 2`, then two level-2
items, draw `2.1.1 2.1.2`, and a level-1 item after them `2.2`; two first items at level 1
then one at level 0 draw `a b 2` (with `%1.%2.`: `1.1 1.2 2`). A `%n` of a deeper level
than any item has reached shows its start. **`w:isLgl` puts every number of that
level's label in decimal**, its own too: upperRoman over `%1.%2.` draws `1.1`, an
upperLetter level `1.2`; on level 0 its own `%1.` draws `1.` while level 1's `%1.%2.`,
without it, keeps `II.a`.

**N.4 Stories.** A table's cells count in the body's sequence, row by row, cell by cell.
A **text box** does not: the text boxes are a story of their own, counted on from box to
box in anchor order (`1 2`, then `3`) and neither continuing the body's count nor
continued by it. A **footnote** likewise: the notes are one story, counted on from note to
note, apart from the body. A **header**'s list is apart from the body's and draws the same
numbers on every page.

**The model.** `parse_numbering` reads each instance's list (the `w:nsid`'s first
definition), its levels (overrides applied, `w:numStyleLink` followed through the
numbering style) and the levels a `w:startOverride` restarts
(`Document.numbering_lists`, `numbering_restarts`); `linebreak.ListCounters` counts by
N.1–N.3; `linebreak.text_box_labels` counts a story's text boxes as one story in anchor
order. Measured against, refuted: counting per `w:num` (cases 28 / 43 in mode 15, lines
229 / 266), a `w:startOverride` as a start only (37 / 43, 249 / 266), each text box
counted alone (42 / 43, 265 / 266).

**Scores.** Before: **14 / 43 cases in each setting** (override 1 / 15, instance 2 / 8,
restart 3 / 8, legal 2 / 5, story 6 / 7), 192 / 266 lines, 1,133 of Word's 1,227 glyphs
matched. After: **43 / 43 cases, 266 / 266 lines, 1,227 / 1,227 glyphs** (three object
starts 0.0005 px off, as elsewhere), in all three settings. No committed or local document
counts otherwise (every body label of 28 documents is unchanged) and `tools/fidelity.py`
draws every page as before: the fidelity baselines do not move.

**A public entry point** (docx-agent's proposal, which passed a stand-in document and
paragraph): `ListCounters.item(numbering, numId, ilvl)` counts the next item from a
`parse_numbering` result alone and returns its label and its number at that level
(`ListItem`); `label`, the layout's, is the same count for a parsed paragraph. Every body
paragraph of the `numbering-15` probe gets the same label and number from either
(`test_the_public_entry_point_counts_as_the_layout_does`).

**docx-agent's model** (`markdown/read.py`, `Reader.number`) agrees on N.1's sharing,
the first-use restart, the notes and the header, and the link. It differs from Word in:
definitions sharing a `w:nsid` (counted apart, each with its own levels); the restart
value of a `w:lvlOverride` holding a `w:lvl` and a `w:startOverride` (it takes the
`w:startOverride`); a level restarted by a shallower item in an overridden instance (it
restarts at the level's `w:start`, not the override); `w:lvlRestart` (ignored: every
deeper level restarts); a shallower level with no count (shown at its start but not
counted, so `2.1` where Word draws `2.2`, and `1` where Word draws `2`); `w:isLgl`
(ignored); and text boxes (counted in their part's sequence after their anchor paragraph,
where Word counts them apart).

**Not measured.** Endnotes and comments; two header or footer parts holding one list;
text boxes in headers (counted per box, as before); a level with no `w:start` (read as 1;
ECMA-376 says 0); numbering from a paragraph style's `w:numPr`; a list item whose
paragraph mark is deleted or whose numbering changed (`w:numberingChange`); bullets
between numbered levels.

---

## Non-goals

- **Editing.** This reads and renders.
- **Round-tripping to DOCX.** A different project.
- **Matching LibreOffice.** It is an approximation of Word; the oracle is Word.
- **A number that scores the whiteness of paper.** See the landscape survey. If a fidelity
  metric is ever added here, it measures glyph positions. *Phase 5 added two, in that order: the
  glyph-position check (every glyph against Word's pen position, baseline, face and size)
  is the metric; the raster SSIM is taken over foreground pixels only and reported beside
  it, never instead (5.3, 5.7).*

---

## Suggested order

1. ~~**Phase 2's two one-hour probes first**~~ **Done** — x is not quantised to the
   device grid (Word's unit is 1/4096 pt); the accumulator is reproduced for every rule
   but `auto` multiples above one. See *Phase 2 — measured*.
2. ~~Style inheritance (Phase 6, first row)~~ **Done** — every glyph of every probe and
   document resolves to the face and size Word drew; 488 / 495 real-world baselines and
   24 / 34 of the generated inherited-style document reproduce to 0 px, every miss 1 px
   and in the vertical model. See *Style inheritance — measured*.
3. ~~**The vertical model's remaining 1 px cases**~~ **Done** — one cause for all three:
   the baseline rounds in a line box that includes the space around the text. Every
   in-scope baseline of every probe (36,632) and document (495 real-world, 34 + 91
   generated) reproduces to 0 px. See *The line box — measured*.
   ~~Nine more real documents~~ **Scored** — `wordto` (committed) and `filesamples`
   (not: no licence), Word 12–15, compatibility modes 12/14/15, natural pagination. The
   model reproduced 5,636 / 5,769 of their baselines unchanged; every miss is one of
   seven recorded findings. See *Nine more real documents*.
   ~~Findings 1–4~~ **Adopted, each on its own probe** — the first paragraph of a
   section keeps its space before (and a dropped one stays in its box);
   `pageBreakBefore` keeps it below mode 15; autospacing is 14 pt, but nothing between
   list items or before the document's first paragraph; a list label's descent does
   not count. The nine documents score 5,749 / 5,769 and every probe line is exact but
   the 15 that record `doNotUseHTMLParagraphAutoSpacing`. See *Findings 1–4, probed and
   adopted*.
   ~~Finding 5~~ **Probed; the line rule for scripts refuted, two others adopted** — a
   superscript or subscript takes part in its line at its own size, on the baseline
   (the model already did); a run border grows its run's extent, and an `auto` multiple
   scales it; `w:position` moves the run's extent; the script's glyphs are drawn at the
   face's OS/2 script size. Every one of the probe's 43,200 lines is exact; the nine
   documents score 5,755 / 5,769, every miss mode 12's (`sample1`) or a table cell. See
   *Superscripts, subscripts, position and run borders — measured*.
   **Local corpora** — documents that may be measured but not named are scored from
   `scratch/` by a committed, generic mechanism. One exposed that a paragraph mark
   larger than the text of its line does not count on that line. See *Local corpora*.
   ~~The paragraph mark~~ **Adopted** — on its own probe, swept over size, face,
   three-line paragraphs, labels, five line rules and four compatibility settings: the
   mark takes part only in a line of nothing but spaces; and, found by the same probe, a
   space takes no part in its line at all, and a list label counts on its paragraph's
   first line only. Every one of the probe's 10,900 lines is exact. See *The paragraph
   mark in the line height — measured*.
   **Next: Phase 3.** Nothing recorded is left on the line height outside mode 12.
   Findings 6 (table-cell size by mode) can wait for tables *(settled: Tables —
   measured, stage 2)*; 7 (mode 12 and embedded
   Ubuntu, now with two more 1 px lines) for mode 12 work; the script *offset* for
   Phase 5, which draws it; Baskerville Old Face's metrics for the first document that needs
   them (its probe is written down); and `doNotUseHTMLParagraphAutoSpacing`'s adding for the first
   document that sets it.
4. ~~Phase 1's decision executes when Phase 3 starts, not before.~~ **Executed before
   Phase 3, by decision: the shared half of `pptx2svg` is now `ooxml-common`.** Phase 3
   decides the paragraph protocol that `text/wrap.py` would need to follow it.
5. Phases 3 → 4 in order. There is no shortcut: each is measured on top of the last.
   ~~Phase 3~~ **Done for Latin text** — the budget is exact, inclusive and without
   slack, and every line of every probe and 7,164 / 7,167 real-document lines break
   where Word breaks them. See *Phase 3 — measured*. **Next: Phase 4.**
   ~~Phase 4~~ **Done** — what fits (the text and its bottom border, inclusive, to the
   unit), widow control, `keepLines`, `keepNext` chains, manual and section breaks,
   footnotes' room, tall headers and inline pictures, each on its own probe; every page
   top and every paragraph of the multi-page probe on Word's page (537 / 537), and every
   page top it reaches in the real documents (`sample4` 168 / 168). See *Phase 4 —
   measured*. **Next: Phase 5**, and tables (Phase 6), past which the paginator does not
   go.
6. ~~Phase 5 only when 4 is done.~~ **Done** -- one SVG per page, a transcription of the
   layout: every glyph of every committed document at Word's pen position, baseline, face
   and size (37,619 / 37,619; `sample4` 496,104 / 496,104); decorations, borders and
   shading measured on a probe of their own; the ink drawn at the device-rounded size; the
   script offset measured with faces of our own; tables, floating drawings and columns
   stop the layout with a marked band and a warning. See *Phase 5 — measured*. **Next:**
   tables (Phase 6) -- the one obstacle every real document here meets -- then headers and
   footers (drawn where their room already is), floating drawings, and Phase 3's
   justification in mode 15.
   ~~Tables~~ **Done, stages 1-6** -- the grid, a cell's content (and finding 6), row
   heights, merged cells, borders and shading, and tables across pages, each on a
   committed probe Word exported, every glyph of every committed document's tables where
   Word drew it (`sample-with-table` 1,141 / 1,141, `sample-simple` 844 / 844,
   `sample-10pages` 4,211 / 4,211, `sample-5pages` 2,727 / 2,727). Autofit is recorded,
   not modelled (stage 7). See *Tables — measured*. **Next:** floating drawings (every
   local template stops at one), floating tables (`sample1`), headers and footers,
   autofit's shares, and Phase 3's justification in mode 15.
   ~~Headers and footers~~ **Done, with fields** -- drawn from where Word draws them, the
   story each page shows and its number, `PAGE` / `NUMPAGES` / `SECTIONPAGES` computed
   from the layout's own pages; every glyph of the three probes' 11 documents exact but
   what Word alone can know (recomputed fields, number words) and two margin-grid glyphs.
   No real document here has a header or footer glyph (`sample4` has no page numbers).
   See *Headers, footers and fields — measured*. **Next:** floating drawings (a header
   logo is one), floating tables, autofit's shares, and Phase 3's justification in mode 15.
   **Floating drawings, stages 1-3** -- anchors that do not move text are positioned
   (every `relativeFrom`, offsets and alignments, in twips), stacked (behind and in front
   of the text, by `relativeHeight`, a header's before the body's) and drawn (pictures,
   shapes, theme colours, groups, text boxes), in the body and in headers and footers;
   `wrapTopAndBottom` pushes the text below; every local template is laid out to its end.
   ~~Text beside a drawing (stage 4)~~ **Done** -- lines beside a drawing are broken
   into segments where the paginator places them (the minimum segment, the exact left
   edge, the right edge's pixel rounding, polygons, `wrapText`, indents, alignment), every
   line of the probe's mode-15 families within half a pixel but ties on a half pixel. See
   *Floating drawings — measured*, F.8. ~~A table beside a drawing~~ **Done** (F.15).
   ~~Anchors in table cells~~ **Done** (F.16, and the cell's text around them: F.17).
   ~~Floating tables~~ **Done** (F.18).
   ~~A width in percent~~ **Done** where the cells state theirs in percent (Tables, stage
   7a). ~~A table nested in a cell~~ **Done** (Tables, "A table nested in a cell").
   Cells in `dxa` under a width in percent: shared by their widths and their narrowest
   content (stage 7a). ~~Autofit's shares~~ **Done** where the probe settles them (stage
   7b): columns from their content, the room shared by widest less narrowest content;
   a table narrower than its content kept to the room, shared by its widest words (7c).
   **Next:** a cell across columns wider than they are (stage 7b), a header's drawing the
   body wraps around.
   ~~Charts and SmartArt~~ **Drawn** -- a chart by `ooxml-common`'s chart layout under
   Word's measured rules (white frame and plot, a title's band and baseline, side legends
   against the frame, legend order, stated line widths), 794 / 804 texts, 664 / 666 fills
   and 1,574 / 1,578 strokes of 58 probe charts where Word drew them, in every setting;
   SmartArt from the drawing Word caches, its text in the face's own line box, every
   probe glyph and shape exact. See F.20 and F.21. **Next:** the radar's value labels, and
   SmartArt layouts other than the block list, which need a layout definition written out.
