"""Measurements kept per process (round 37): :class:`docx2svg.relayout.MeasureStore`.

A layout handed between processes without its measurements (they are most of its pickle)
is laid out again from with the measurements the receiving process made of the same
document before, by lineage.  **The property** is :mod:`test_relayout`'s: page for page,
byte for byte, the layout of the new bytes alone -- with a store, without one, and from a
store filled by layouts of other versions.  Every document is made up (``tools/report_docs.py``).
"""

from __future__ import annotations

import dataclasses
import pickle
import random
import threading

import report_docs
from docx2svg import ConvertOptions, lay_out_docx, page_svgs
from docx2svg.fonts import InstalledFonts
from docx2svg.relayout import MeasureStore


def _light(layout):
    """``layout`` as a worker hands it on: pickled, without its measurements."""
    state = dataclasses.replace(layout.state, memo={})
    return pickle.loads(pickle.dumps(dataclasses.replace(layout, state=state)))


def _full_svgs(data: bytes) -> list:
    layout, data, fonts = lay_out_docx(data, keep=False)
    return page_svgs(layout, data, fonts=fonts)


def test_random_edits_from_layouts_without_their_measurements_lay_out_as_a_full_layout():
    # Two "processes", each its own store, taking turns: each layout is made from the one
    # before, handed over without its measurements, as a worker pool hands them on.
    rng = random.Random(37)
    stores = [MeasureStore(), MeasureStore()]
    blocks = report_docs.report(6, seed=37)
    layout = lay_out_docx(report_docs.package(blocks), measures=stores[0])[0]
    hits = 0
    for step in range(10):
        blocks, what = report_docs.edit(blocks, rng)
        data = report_docs.package(blocks)
        store = stores[step % 2]
        options = ConvertOptions()
        again, again_data, fonts = lay_out_docx(data, options, previous=_light(layout), measures=store)
        assert page_svgs(again, again_data, fonts=fonts) == _full_svgs(data), f"step {step}: {what}"
        hits += len(again.state.memo)
        layout = again
    assert hits


def test_a_lineage_goes_on_while_measurements_are_reused_and_starts_again_when_not():
    blocks = report_docs.report(4, seed=3)
    first = lay_out_docx(report_docs.package(blocks))[0]
    assert first.state.lineage
    edited = list(blocks)
    edited[3] = dataclasses.replace(edited[3], text=edited[3].text + " More.")
    again = lay_out_docx(report_docs.package(edited), previous=first)[0]
    assert again.state.lineage == first.state.lineage
    # A header changed: the measurements are not reused, a new lineage starts.
    other = lay_out_docx(report_docs.package(edited, header="Another made-up header"), previous=again)[0]
    assert other.relaid == "headers-footers"
    assert other.state.lineage != again.state.lineage
    # A whole layout of its own starts one too.
    assert lay_out_docx(report_docs.package(blocks))[0].state.lineage != first.state.lineage


def test_a_store_measures_the_unchanged_blocks_no_more():
    blocks = report_docs.report(6, seed=5)
    store = MeasureStore()
    first = lay_out_docx(report_docs.package(blocks), measures=store)[0]
    edited = list(blocks)
    k = next(k for k in range(len(blocks) // 2, len(blocks)) if blocks[k].kind == "para")
    edited[k] = dataclasses.replace(edited[k], text=edited[k].text + " Changed.")
    data = report_docs.package(edited)
    calls = []
    import docx2svg.lines as lines

    original = lines.broken

    def counted(*args, **kwargs):
        calls.append(1)
        return original(*args, **kwargs)

    lines.broken = counted
    try:
        with_store = lay_out_docx(data, previous=_light(first), measures=store)
        measured_with = len(calls)
        calls.clear()
        without = lay_out_docx(data, previous=_light(first))
        measured_without = len(calls)
    finally:
        lines.broken = original
    assert page_svgs(with_store[0], with_store[1], fonts=with_store[2]) == \
        page_svgs(without[0], without[1], fonts=without[2])
    assert measured_with < measured_without / 4, (measured_with, measured_without)


def test_a_store_is_bounded_and_safe_to_share():
    store = MeasureStore(lineages=2, spare=0)
    store.put(b"a", {k: k for k in range(10)})
    store.put(b"a", {k: k for k in range(10, 13)})
    assert set(store.get(b"a")) == {7, 8, 9, 10, 11, 12}    # twice the newest, the newest last
    store.put(b"b", {1: 1})
    store.put(b"c", {1: 1})
    assert store.get(b"a") == {} and store.get(b"c") == {1: 1}
    assert store.get(b"") == {}
    copy = store.get(b"c")
    copy[2] = 2
    assert store.get(b"c") == {1: 1}

    def work(n: int) -> None:
        for k in range(200):
            store.put(bytes([n % 3]), {(n, k): k})
            store.get(bytes([n % 3]))

    threads = [threading.Thread(target=work, args=(n,)) for n in range(6)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()
    assert len(store) <= 2 * 2 + 0


def test_changed_pages_left_out_lays_out_the_same():
    blocks = report_docs.report(4, seed=9)
    first = lay_out_docx(report_docs.package(blocks))[0]
    edited = list(blocks)
    edited[2] = dataclasses.replace(edited[2], text=edited[2].text + " Changed.")
    data = report_docs.package(edited)
    told = lay_out_docx(data, previous=first)
    quiet = lay_out_docx(data, previous=first, changed_pages=False)
    assert quiet[0].changed_pages is None and told[0].changed_pages is not None
    assert page_svgs(quiet[0], quiet[1], fonts=quiet[2]) == page_svgs(told[0], told[1], fonts=told[2])


def test_installed_fonts_pickle_without_their_index():
    fonts = InstalledFonts()
    face = fonts.face("Calibri")
    assert "_indexes" in fonts.__dict__
    again = pickle.loads(pickle.dumps(fonts))
    assert "_indexes" not in again.__dict__
    assert again == fonts
    found = again.face("Calibri")
    assert (found is None) == (face is None)
    if face is not None:
        assert (found.source, found.offset) == (face.source, face.offset)
