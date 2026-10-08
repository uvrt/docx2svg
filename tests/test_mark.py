"""The paragraph mark's part in its line's height, without Word.

``tests/fixtures/mark-observations.json`` holds the lines Word drew for every document of
``tools/make_mark_probe.py`` (recorded by ``tools/read_mark_probe.py --record``); the
documents are regenerated from the generator.  Every baseline goes through
``baselines.predict``, and so through the model's rule (``vertical.line_items``): a
space takes no part in its line's height, the mark takes part only in a line that holds
nothing but spaces (``vertical.line_items``), and a list label only in its paragraph's
first line (``vertical.label_items``).  ROADMAP.md, "The paragraph mark in the line height", has the
tables.
"""

from __future__ import annotations

import json
from collections import defaultdict
from pathlib import Path

import pytest

import baselines
import make_mark_probe as probe
import read_mark_probe as reader
from docx2svg import vertical

DATA = json.loads((Path(__file__).parent / "fixtures" / "mark-observations.json").read_text(encoding="utf-8"))

#: The sweep's families under the model, [exact, scored], the same in each of the four
#: settings: every line.
SWEEP = {
    "size-same": [150, 150], "size-big": [150, 150], "size-small": [150, 150],
    "size-onlyspace": [150, 150], "size-empty": [150, 150],
    "size-multisame": [215, 215], "size-multibig": [215, 215],
    "face-tall": [75, 75], "face-deep": [50, 50], "face-smallertall": [25, 25],
    "face-smallerdeep": [25, 25], "face-short": [25, 25], "face-multitall": [40, 40],
    "face-empty": [25, 25],
    "space-onlybig": [75, 75], "space-onlybig-smallmark": [75, 75], "space-onlysmall": [75, 75],
    "space-trailingbig": [75, 75], "space-interiorbig": [75, 75], "space-onlytab": [75, 75],
    "space-interiortab": [75, 75], "space-onlytall": [75, 75], "space-interiortall": [75, 75],
    "space-onlynbsp": [75, 75], "space-interiornbsp": [75, 75],
    "label-same": [50, 50], "label-bigmark": [50, 50], "label-bigmark-label24": [50, 50],
    "label-label30": [50, 50], "label-multibigmark": [75, 75],
    "label-multibigmark-label24": [75, 75], "label-multilabel30": [75, 75],
}


def _families(which=lambda p: True) -> dict:
    """``{(setting, family): [exact, scored]}``; the first two documents are family
    ``first``."""
    out: dict = defaultdict(lambda: [0, 0])
    for (setting, case), (exact, count, _) in reader.tabulate(reader.offline(DATA, which)).items():
        cell = out[(setting, case.split("/")[0] if setting.startswith("sweep") else "first")]
        cell[0] += exact
        cell[1] += count
    return out


def _total(table: dict) -> list[int]:
    return [sum(v[0] for v in table.values()), sum(v[1] for v in table.values())]


def test_the_recording_is_of_these_documents():
    assert set(DATA["documents"]) == {p.name for p in probe.PROBES}
    for p, results in reader.offline(DATA):
        assert all(r.status in ("exact", "miss") for r in results), p.name


def test_a_large_mark_takes_no_part_in_a_line_with_text():
    """The first two documents: 26 to 30 pt marks over 24 pt text, plain, bold, after a
    trailing space; an 11 pt mark; empty paragraphs.  Every line (the mark on the line,
    as the model had it: 56 / 120)."""
    assert _total(_families(lambda p: not p.sweep)) == [120, 120]


@pytest.mark.parametrize("setting", sorted(probe.SWEEP_SETTINGS))
def test_the_sweep(setting):
    """Size, face, three-line paragraphs, labels and spaces, five line rules; each
    compatibility setting draws every line alike."""
    families = _families(lambda p: p.sweep and p.setting == setting)
    assert {family: cell for (_, family), cell in families.items()} == SWEEP


class _Share(list):
    last = False


def _every_line(chars, mark):
    return [item for item, _ in chars] + [mark]


def _with_spaces(chars, mark):
    """The mark rule, with spaces counted like text."""
    if any(char != " " for _, char in chars):
        return [item for item, _ in chars]
    return [item for item, _ in chars] + [mark]


def _without_no_break_spaces(chars, mark):
    """U+00A0 left out as a space is."""
    blank = (" ", "\u00a0")
    if any(char not in blank for _, char in chars):
        return [item for item, char in chars if char not in blank]
    return [mark]


def _trailing_spaces_only(chars, mark):
    """Spaces left out at the end of the line only."""
    if all(char == " " for _, char in chars):
        return [mark]
    last = max(k for k, (_, char) in enumerate(chars) if char != " ")
    return [item for k, (item, _) in enumerate(chars) if k <= last]


def _spaces_and_mark_on_a_blank_line(chars, mark):
    if any(char != " " for _, char in chars):
        return [item for item, char in chars if char != " "]
    return [item for item, _ in chars] + [mark]


#: Each replaces one rule of the model and keeps the others.
@pytest.mark.parametrize(("rule", "score"), [
    # The mark on every line of its paragraph, spaces counted: the model before, but
    # for the label rule.
    (_every_line, 7648),
    # The mark on the paragraph's last line, where it sits (and on a line of spaces).
    ("last line", 8036),
    # Spaces: counted like text; left out only at the end of a line; a no-break space
    # left out too; on a line of spaces alone, counted with the mark.
    (_with_spaces, 9784),
    (_trailing_spaces_only, 10516),
    (_without_no_break_spaces, 10516),
    (_spaces_and_mark_on_a_blank_line, 10360),
    # A list label on every line of its paragraph.
    ("label on every line", 10548),
])
def test_the_refuted_rules_really_are_refuted(monkeypatch, rule, score):
    shares, line_items = baselines.line_shares, baselines.line_items

    def tagged(document, paragraph, texts):
        out = [_Share(share) for share in shares(document, paragraph, texts)]
        out[-1].last = True
        return out

    def last_line(document, paragraph, pp, chars=None, **options):
        items, label = line_items(document, paragraph, pp, chars, **options)
        if chars is not None and not chars.last and any(char != " " for _, char in chars):
            items = items[:-1]
        return items, label

    if rule == "label on every line":
        monkeypatch.setattr(vertical, "label_items", lambda label, first_line: list(label))
    elif rule == "last line":
        monkeypatch.setattr(vertical, "line_items", _every_line)
        monkeypatch.setattr(baselines, "line_shares", tagged)
        monkeypatch.setattr(baselines, "line_items", last_line)
    else:
        monkeypatch.setattr(vertical, "line_items", rule)
    assert _total(_families()) == [score, 10900]
