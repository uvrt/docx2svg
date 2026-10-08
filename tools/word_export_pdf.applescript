-- Export a .docx to PDF using the locally installed Microsoft Word.
--
-- Used as the ground-truth oracle for fidelity testing: Word is the reference
-- implementation, so its own output is the only fully authoritative answer to "where
-- does this line break, and on which page does this paragraph land?".
--
--   osascript tools/word_export_pdf.applescript <input.docx> <output.pdf> [final|accept]
--
-- **Revisions.**  With no third argument Word exports what its window shows on opening,
-- and for a document with tracked changes that is *All Markup*: measured on Word 16.106,
-- a deletion is drawn in a balloon ("heeft verwijderd: ...", the interface's language)
-- in a markup pane beside the text, and the text is moved to make room for it (on an A4
-- page, 73 px left and 363 px down at 300 dpi, the glyphs at their size).
-- Neither `w:revisionView w:markup="0"` in settings.xml nor the document's `print
-- revisions` changes that export.  `final` sets the window's view to *No Markup* --
-- revisions and comments hidden, the final view, no markup pane -- before exporting;
-- the setting belongs to the window and does not outlive it (the next export without it
-- shows markup again).  `accept` accepts every revision in the open copy, which is
-- closed unsaved, then exports in the same view (comments are not revisions: without it
-- their pane is printed).  A document without revisions or comments exports
-- identically all three ways, and `tools/oracle.py` exports one that has them in the
-- final view unless asked otherwise: the final view is what the renderer draws.
--
-- Then measure per page with pypdfium2.  Word's PDF carries text as *vector* objects --
-- no raster anywhere -- so `page.get_textpage().get_charbox(i)` returns the exact ink
-- box of every glyph, in points, to better than a thousandth of a point.  That makes
-- this a materially better oracle than the PowerPoint one in the sibling `pptx2svg`
-- project, where fidelity is scored by rasterising and computing SSIM: here the thing
-- being measured -- line breaking and pagination -- is read out as *numbers*, not as
-- pixel overlap.  See ROADMAP.md, "Phase 0".
--
-- Constraints discovered on Word 16.106 / macOS.  Each of these cost the sibling
-- `pptx2svg` project real time in its PowerPoint form; they are written down here so
-- they cost this one none.
--
--   * **`open` does not return a document reference.**  PowerPoint's dictionary has
--     `open` yield the presentation, so `set p to open inPath` works there.  Word's does
--     not: the same line binds `p` to something unusable and every later reference fails
--     with a type error that names neither `open` nor the document.  The fix is one
--     line -- `set d to active document` *after* opening -- and the failure is worth
--     recognising on sight, because it reads as a broken input file.
--
--   * **An AppleEvent timeout (-1712) is the default failure mode.**  AppleScript's
--     own deadline is two minutes, and Word can exceed it on a first launch (font cache
--     rebuild), on a document with many pages, or whenever it decides to show a dialog.
--     Without an explicit `with timeout of N seconds` the script fails with -1712 while
--     Word is working normally, which reads as "Word cannot open this" and is not.  The
--     180 seconds below is deliberately generous; a caller that wants to fail fast
--     should impose its own wall-clock timeout on the `osascript` process, which is
--     where recovery has to live anyway -- this script is blocked inside `open` while a
--     modal dialog is up and cannot dismiss it.
--
--   * **A `~$` lock file left by a failed attempt wedges the next attempt.**  Word writes
--     `~$<name>.docx` beside the document it has open and removes it on a clean close.
--     A run that died -- a timeout, a kill, a crash -- leaves the lock behind, and the
--     next `open` of that path then either silently opens a *read-only* copy (whose
--     `save as` writes nothing while the script still exits 0) or raises a "file is in
--     use" dialog that blocks AppleEvents entirely.  Both look like a broken document.
--     The recovery is two commands:
--
--         pkill -x "Microsoft Word"
--         rm -f "$(dirname "$INPUT")/~\$$(basename "$INPUT")"
--
--     Run them *before* concluding anything about the input.  `tools/oracle.py` does
--     this automatically; if you are driving this script by hand, do it by hand.
--
--   * **Word is sandboxed, and its file grants are per app.**  An Office app may read
--     and write outside its own container only where the user has granted *that app*
--     access.  PowerPoint was granted `~/pptx2svg-oracle/` long ago, which is why the
--     sibling project never saw a prompt; Word had never been granted anything, so
--     every export into `~/docx2svg-oracle/` raised a "grant file access" dialog for the
--     new probe file.  From this script's side a pending dialog is indistinguishable
--     from a hang: `open` blocks, the AppleEvent times out with -1712 after the full
--     180 s, no window is visible to System Events, and no `~$` lock is written (the
--     file was never opened).  Measured 2026-09-24, Phase 2: the same known-good probe
--     that exported in 4 s one minute timed out the next, for this reason alone.
--
--     The fix needs no grant at all: **the Office group container**,
--     `~/Library/Group Containers/UBF8T346G9.Office/`, is shared by every Office app
--     and each may read and write there without asking.  The oracle lives in
--     `docx2svg-oracle/` inside it (`ORACLE_DIR` in `tools/oracle.py` is the one place
--     the path is written).  Verified: a document staged there opened and exported
--     through Word's own `open` event in 3.9 s with no dialog.  The path contains a
--     space -- callers pass it as one argv element, never through a shell.  `/tmp`
--     and arbitrary directories are refused, as before.
--
-- One further note carried over from the PowerPoint oracle, which applies unchanged:
--
--   * Always check the output PDF exists rather than trusting the exit status.  Several
--     of the failure modes above return success and write nothing.

on run argv
    if (count of argv) is not 2 and (count of argv) is not 3 then
        error "Usage: osascript word_export_pdf.applescript <input-docx> <output-pdf> [final|accept]"
    end if

    set inputPosixPath to item 1 of argv
    set outputPosixPath to item 2 of argv
    set revisionView to ""
    if (count of argv) is 3 then set revisionView to item 3 of argv

    -- See the second note above: without this, a Word that is merely slow fails with
    -- -1712 and is indistinguishable from a Word that will not open the file at all.
    with timeout of 180 seconds
        tell application "Microsoft Word"
            activate
            open (POSIX file inputPosixPath)
            -- See the first note above: `open` yields nothing usable in Word's
            -- dictionary, unlike PowerPoint's.  The document just opened is the active
            -- one; this script opens exactly one document per run, so that is
            -- unambiguous here.
            set d to active document
            -- See "Revisions" above.
            if revisionView is "accept" then accept all revisions d
            if revisionView is not "" then
                set show revisions and comments of view of active window to false
                set revisions view of view of active window to revisions view final
            end if
            save as d file name outputPosixPath file format format PDF
            close d saving no
            -- Quit rather than leave Word running: a document left open from an earlier
            -- run is, in the PowerPoint equivalent, a documented cause of every later
            -- export failing with nothing visible to blame.
            quit saving no
        end tell
    end timeout
end run
