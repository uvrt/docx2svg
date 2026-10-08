-- Save a .docx through the locally installed Microsoft Word: open it and save it again as
-- a .docx, so that what Word computes and caches on saving is in the file -- the laid-out
-- drawing of a SmartArt diagram (``dsp:drawing``) first of all, which Word lays out again
-- from the diagram's data whenever it opens one and writes on every save.
--
--   osascript tools/word_save_docx.applescript <input.docx> <output.docx>
--
-- Every constraint in ``word_export_pdf.applescript``'s header applies here unchanged --
-- `open` returns nothing usable, the 180-second timeout, the `~$` lock, the sandbox --
-- and ``tools/oracle.py``'s ``resave`` is the one caller, with the recovery in place.

on run argv
    if (count of argv) is not 2 then
        error "Usage: osascript word_save_docx.applescript <input-docx> <output-docx>"
    end if

    set inputPosixPath to item 1 of argv
    set outputPosixPath to item 2 of argv

    with timeout of 180 seconds
        tell application "Microsoft Word"
            activate
            open (POSIX file inputPosixPath)
            set d to active document
            save as d file name outputPosixPath file format format document
            close d saving no
            quit saving no
        end tell
    end timeout
end run
