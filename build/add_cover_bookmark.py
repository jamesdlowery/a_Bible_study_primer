#!/usr/bin/env python3
"""Prepend a "Front Cover" entry to a PDF's own native bookmarks/outline
panel, pointing at page 1.

Why this exists rather than a heading in the source markdown: a
"HiddenHeading" (invisible, white 2pt text with outlineLvl=0 -- the same
mechanism the Title Page already uses successfully) was tried first, since
LibreOffice's PDF export is supposed to turn any outline-level heading into
a PDF bookmark automatically. In practice it didn't work for the cover on
the real GitHub Actions build -- Title Page appeared in the bookmarks
panel, Front Cover didn't, despite using the identical style -- most
likely a LibreOffice-version-specific quirk in that conversion that this
sandbox's own LibreOffice install doesn't reproduce, so it couldn't be
directly debugged here. Rather than keep guessing at that heuristic, this
script edits the compiled PDF's own outline data structure directly, via
PyMuPDF (already a pipeline dependency for detect_pages.py's page-counting)
-- a well-defined, documented operation with no dependency on how any
particular LibreOffice version chooses to interpret heading styles.

Usage: python3 add_cover_bookmark.py <pdf_path>
Edits the PDF in place.
"""
import sys
import os
import fitz  # PyMuPDF


def add_cover_bookmark(pdf_path):
    doc = fitz.open(pdf_path)
    toc = doc.get_toc(simple=True)  # [[level, title, page], ...], 1-indexed pages

    # Idempotent: if a top-level "Front Cover" entry already exists at
    # page 1 (e.g. this script somehow ran twice, or a future LibreOffice
    # version starts picking up the heading on its own after all), don't
    # add a second one.
    already_present = any(
        level == 1 and title.strip() == "Front Cover" and page == 1
        for level, title, page in toc
    )
    if not already_present:
        toc.insert(0, [1, "Front Cover", 1])
        doc.set_toc(toc)
        # A full rewrite (not an incremental save) -- incremental saves on
        # a 1000+ page PDF were observed to add several hundred KB for
        # this one small outline change, apparently re-serializing more
        # than just the edited structure. A full save with garbage
        # collection is both more compact and a cleaner file overall.
        tmp_path = pdf_path + ".tmp"
        doc.save(tmp_path, garbage=4, deflate=True)
        doc.close()
        os.replace(tmp_path, pdf_path)
        print(f"Added 'Front Cover' bookmark to {pdf_path}")
    else:
        doc.close()
        print(f"'Front Cover' bookmark already present in {pdf_path}; left unchanged")


if __name__ == "__main__":
    add_cover_bookmark(sys.argv[1])
