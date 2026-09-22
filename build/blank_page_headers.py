#!/usr/bin/env python3
"""Remove misleading running-header text from pages in the compiled PDF
that shouldn't show one, leaving the watermark, footer, and (where only
part of the header is wrong) the rest of the header text untouched.
Covers two distinct cases, both confirmed directly against the actual
compiled PDF:

1. Blank pages. They carry a running header the same way every other
   page does (via a STYLEREF field that shows the nearest heading's
   text), which is misleading on a page with no real content of its own
   -- it just repeats whatever section was active on the page before it.
   The cover's own three post-cover blank pages already get a genuinely
   blank header at the DOCX level (see build_cover_md()'s header3.xml in
   custom-reference.docx), but the same approach for the ~100
   mid-document blank pages inserted throughout the book to enforce
   odd-page starts was tried and reverted: each would need its own pair
   of OOXML section breaks, and confirmed directly, adding the roughly
   200 extra section boundaries that requires made LibreOffice's
   field-recalculation step in update_fields_and_export.py hang for
   minutes with no output.

2. The book-name half of the header (the text after the em dash) on
   every page before the very first "Heading 4"-styled paragraph in the
   whole document. The running header is built from two STYLEREF fields
   -- STYLEREF 1 for the current top-level section (e.g. "Preface",
   "Table of Contents"), STYLEREF 4 for the current "book" -- and a
   STYLEREF field with nothing valid yet preceding it falls forward to
   the *next* matching paragraph instead of showing nothing. The book's
   very first Heading-4-styled paragraph is "Before 1611: The Road to the
   King James Version" (a subheading inside the King James Tradition
   history chapter, well into the book -- see 080 Histories of Various
   Bible Translations/Protestant/010 History of the King James
   Tradition.md), so every single page before that chapter incorrectly
   shows "-- Before 1611: The Road to the King James Version" tacked onto
   its header, even though the STYLEREF 1 half (the actual section name)
   is completely correct. Confirmed directly: this was previously found
   and "fixed" by removing the title page's entire header instead of
   diagnosing the actual cause, which was wrong -- the title page's own
   header (once corrected to not show a stray section name at all) is
   fine and should not be blanked; only this specific trailing text,
   wherever it wrongly appears, needs to go.

Both are handled the same way: find the affected page or text (differently
for each case), then redact just that -- either the whole header-region
text (blank pages) or just the specific trailing substring (case 2),
never anything else. This works safely because the watermark is confirmed
to not be extractable text at all (a VML shape, invisible to get_text()),
the header text is confirmed to sit in its own narrow band near the very
top of the page (roughly y 36-45), well above where any real body content
starts (confirmed at y >= ~115 even on a heading-only page) -- so a
conservative y-position cutoff can only ever catch the header, never real
content or the footer (near the bottom of the page, y >= ~750) -- and
page.search_for() gives the exact bounding box of just the trailing
substring in case 2, confirmed directly to already exclude the preceding
section-name text sharing the same text span.

Usage: python3 blank_page_headers.py <pdf_path>
Edits the PDF in place.
"""
import sys
import os
import fitz  # PyMuPDF

HEADER_REGION_Y_MAX = 60  # points from the top; header text confirmed ~36-45

# The book's own first Heading-4-styled paragraph; see module docstring
# case 2. If this chapter's own subheading structure ever changes, this
# should be updated to whatever becomes the new first Heading-4 text.
FIRST_HEADING4_TEXT = "Before 1611: The Road to the King James Version"


def _body_text(page):
    """This page's text, excluding whatever sits in the header or footer
    band by position (content varies -- section name, page number, URL,
    version -- but position doesn't)."""
    d = page.get_text("dict")
    leftover = []
    for block in d.get("blocks", []):
        for line in block.get("lines", []):
            for span in line["spans"]:
                y0 = span["bbox"][1]
                if y0 < HEADER_REGION_Y_MAX or y0 > 750:
                    continue
                leftover.append(span["text"])
    return "".join(leftover)


def is_blank_page(page):
    """A blank page's only real content is a bare non-breaking space (the
    placeholder every blank page in this book is built from)."""
    return _body_text(page).replace("\xa0", "").strip() == ""


def find_first_heading4_page(doc):
    """The 0-indexed page where FIRST_HEADING4_TEXT first appears as real
    body content (not just header text) -- i.e. where it's actually
    correct for STYLEREF 4 to start showing it. Every page before this
    one showing it in the header is the forward-looking artifact."""
    for page in doc:
        d = page.get_text("dict")
        for block in d.get("blocks", []):
            for line in block.get("lines", []):
                for span in line["spans"]:
                    y0 = span["bbox"][1]
                    if HEADER_REGION_Y_MAX <= y0 <= 750 and span["text"].strip() == FIRST_HEADING4_TEXT:
                        return page.number
    return None


def redact_full_header(page):
    d = page.get_text("dict")
    for block in d.get("blocks", []):
        for line in block.get("lines", []):
            for span in line["spans"]:
                if span["bbox"][1] < HEADER_REGION_Y_MAX:
                    page.add_redact_annot(span["bbox"])
    page.apply_redactions()


def redact_forward_looking_book_name(page):
    """Removes only the "-- Before 1611..." trailing text from this
    page's header, leaving the section-name text before it untouched --
    confirmed directly that search_for() returns just that substring's
    own bounding box, not the whole (single, undivided) text span it's
    part of."""
    for rect in page.search_for(f"\u2014 {FIRST_HEADING4_TEXT}"):
        if rect.y0 < HEADER_REGION_Y_MAX:
            page.add_redact_annot(rect)
    page.apply_redactions()


def strip_misleading_headers(pdf_path):
    doc = fitz.open(pdf_path)
    blank_pages = []
    forward_looking_pages = []

    first_h4_page = find_first_heading4_page(doc)

    for page in doc:
        if is_blank_page(page):
            blank_pages.append(page.number + 1)  # 1-indexed for the log
            redact_full_header(page)
        elif first_h4_page is not None and page.number < first_h4_page:
            if FIRST_HEADING4_TEXT in page.get_text():
                forward_looking_pages.append(page.number + 1)
                redact_forward_looking_book_name(page)

    if blank_pages or forward_looking_pages:
        tmp_path = pdf_path + ".tmp"
        doc.save(tmp_path, garbage=4, deflate=True)
        doc.close()
        os.replace(tmp_path, pdf_path)
        print(f"Removed header text from {len(blank_pages)} blank page(s) "
              f"and fixed {len(forward_looking_pages)} page(s) with a "
              f"forward-looking book name in {pdf_path}")
    else:
        doc.close()
        print(f"No blank pages or forward-looking book names found in {pdf_path}")


if __name__ == "__main__":
    strip_misleading_headers(sys.argv[1])
