#!/usr/bin/env python3
"""Remove misleading running-header text from pages in the compiled PDF
that shouldn't show one, leaving the watermark, footer, and (where only
part of the header is wrong) the rest of the header text untouched.
Covers three distinct cases, all confirmed directly against the actual
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
   is completely correct.

3. The section-name half of the header (STYLEREF 1 itself) on every page
   before the very first "Heading 1"-styled paragraph in the whole
   document. Same mechanism as case 2, one level up: the title page's own
   heading uses a different, custom style ("HiddenHeading", used so it
   registers as a PDF bookmark without being visible on the page -- see
   build_title_md()), which STYLEREF doesn't match at all, and nothing
   before the title page uses "Heading 1" either -- so it falls forward
   to the book's first real Heading-1 paragraph, "Table of Contents".
   Confirmed directly: with case 2 already fixed, this left the title
   page's header reading exactly "Table of Contents" and nothing else,
   which is what this case removes.

   (An earlier version of this script tried to fix case 3 by blanking the
   title page's entire header, without having diagnosed that the problem
   was specifically an unmatched STYLEREF style, not something wrong with
   the title page needing a header at all. That was reverted; the title
   page's header is fine in general and shouldn't be blanked outright --
   only this specific forward-looking text needs to go, the same
   surgical way as case 2.)

All three are handled the same way: find the affected page or text
(differently for each case), then redact just that -- either the whole
header-region text (blank pages) or just the specific wrong substring
(cases 2 and 3), never anything else. This works safely because the
watermark is confirmed to not be extractable text at all (a VML shape,
invisible to get_text()), the header text is confirmed to sit in its own
narrow band near the very top of the page (roughly y 36-45), well above
where any real body content starts (confirmed at y >= ~115 even on a
heading-only page) -- so a conservative y-position cutoff can only ever
catch the header, never real content or the footer (near the bottom of
the page, y >= ~750) -- and page.search_for() gives the exact bounding
box of just the wrong substring in cases 2 and 3, confirmed directly to
already exclude any correct text sharing the same text span.

Usage: python3 blank_page_headers.py <pdf_path>
Edits the PDF in place.
"""
import sys
import os
import fitz  # PyMuPDF

HEADER_REGION_Y_MAX = 60  # points from the top; header text confirmed ~36-45

# The book's own first Heading-4- and Heading-1-styled paragraphs; see
# module docstring cases 2 and 3. If the document's own heading structure
# ever changes, these should be updated to whatever becomes the new
# first Heading-4/Heading-1 text.
FIRST_HEADING4_TEXT = "Before 1611: The Road to the King James Version"
FIRST_HEADING1_TEXT = "Table of Contents"


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


def find_first_occurrence_page(doc, text):
    """The 0-indexed page where `text` first appears as real body content
    (not just header text) -- i.e. where it's actually correct for a
    STYLEREF field to start showing it. Every page before this one
    showing it in the header is the forward-looking artifact. Takes the
    *first* match in document order, since some heading text (e.g. "Table
    of Contents") can also legitimately recur later as body content, such
    as a "Back to Table of Contents" link at the end of each chapter --
    only the first occurrence is the paragraph a STYLEREF field is
    actually referencing."""
    for page in doc:
        d = page.get_text("dict")
        for block in d.get("blocks", []):
            for line in block.get("lines", []):
                for span in line["spans"]:
                    y0 = span["bbox"][1]
                    if HEADER_REGION_Y_MAX <= y0 <= 750 and span["text"].strip() == text:
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


def redact_header_substring(page, text):
    """Removes only the given substring from this page's header, leaving
    the rest of the header text untouched -- confirmed directly that
    search_for() returns just that substring's own bounding box, not the
    whole (single, undivided) text span it's part of."""
    for rect in page.search_for(text):
        if rect.y0 < HEADER_REGION_Y_MAX:
            page.add_redact_annot(rect)
    page.apply_redactions()


def strip_misleading_headers(pdf_path):
    doc = fitz.open(pdf_path)
    blank_pages = []
    fixed_pages = []

    first_h4_page = find_first_occurrence_page(doc, FIRST_HEADING4_TEXT)
    first_h1_page = find_first_occurrence_page(doc, FIRST_HEADING1_TEXT)

    for page in doc:
        if is_blank_page(page):
            blank_pages.append(page.number + 1)  # 1-indexed for the log
            redact_full_header(page)
            continue

        fixed_this_page = False
        if first_h4_page is not None and page.number < first_h4_page:
            if FIRST_HEADING4_TEXT in page.get_text():
                redact_header_substring(page, f"\u2014 {FIRST_HEADING4_TEXT}")
                fixed_this_page = True
        if first_h1_page is not None and page.number < first_h1_page:
            if FIRST_HEADING1_TEXT in page.get_text():
                redact_header_substring(page, FIRST_HEADING1_TEXT)
                fixed_this_page = True
        if fixed_this_page:
            fixed_pages.append(page.number + 1)

    if blank_pages or fixed_pages:
        tmp_path = pdf_path + ".tmp"
        doc.save(tmp_path, garbage=4, deflate=True)
        doc.close()
        os.replace(tmp_path, pdf_path)
        print(f"Removed header text from {len(blank_pages)} blank page(s) "
              f"and fixed {len(fixed_pages)} page(s) with a forward-looking "
              f"header in {pdf_path}")
    else:
        doc.close()
        print(f"No blank pages or forward-looking headers found in {pdf_path}")


if __name__ == "__main__":
    strip_misleading_headers(sys.argv[1])
