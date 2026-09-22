#!/usr/bin/env python3
"""Remove misleading running-header text from pages in the compiled PDF
that shouldn't show one, leaving the watermark and footer untouched.
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

2. The title page specifically. Confirmed directly (a screenshot of the
   actual compiled PDF) that it shows "Table of Contents" as its running
   header, despite having real content of its own and not being part of
   that section at all. Root cause: STYLEREF 1/4 look specifically for
   the nearest preceding paragraph in the built-in "Heading 1"/"Heading
   4" styles. The title page's own heading uses a different, custom
   style ("HiddenHeading", used so it registers as a PDF bookmark without
   being visible on the page -- see build_title_md()), which STYLEREF
   doesn't match at all; since nothing before the title page uses
   "Heading 1" either, the field falls forward to the next paragraph that
   does, which is Table of Contents' own heading. No real book puts a
   running header on its title page anyway, so the fix is the same as
   for blank pages: remove it.

Both are handled the same way: find the affected page (differently for
each case), then redact just the header-region text there, leaving
everything else untouched. This works safely because the watermark is
confirmed to not be extractable text at all (a VML shape, invisible to
get_text()), and the header text is confirmed to sit in its own narrow
band near the very top of the page (roughly y 36-45), well above where
any real body content starts (confirmed at y >= ~115 even on a
heading-only page) -- so a conservative y-position cutoff can only ever
catch the header, never real content or the footer (near the bottom of
the page, y >= ~750).

Usage: python3 blank_page_headers.py <pdf_path>
Edits the PDF in place.
"""
import sys
import os
import fitz  # PyMuPDF

HEADER_REGION_Y_MAX = 60  # points from the top; header text confirmed ~36-45


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


def is_title_page(page):
    """The title page's own heading text ("Title Page") is present but
    invisible (white, 2pt -- see build_title_md()'s HiddenHeading), so
    it's still found by get_text() even though nothing shows on screen."""
    return "Title Page" in _body_text(page)


def redact_header(page):
    d = page.get_text("dict")
    for block in d.get("blocks", []):
        for line in block.get("lines", []):
            for span in line["spans"]:
                if span["bbox"][1] < HEADER_REGION_Y_MAX:
                    page.add_redact_annot(span["bbox"])
    page.apply_redactions()


def strip_misleading_headers(pdf_path):
    doc = fitz.open(pdf_path)
    blank_pages = []
    title_pages = []
    for page in doc:
        if is_blank_page(page):
            blank_pages.append(page.number + 1)  # 1-indexed for the log
            redact_header(page)
        elif is_title_page(page):
            title_pages.append(page.number + 1)
            redact_header(page)

    if blank_pages or title_pages:
        tmp_path = pdf_path + ".tmp"
        doc.save(tmp_path, garbage=4, deflate=True)
        doc.close()
        os.replace(tmp_path, pdf_path)
        print(f"Removed header text from {len(blank_pages)} blank page(s) "
              f"and {len(title_pages)} title page(s) in {pdf_path}")
    else:
        doc.close()
        print(f"No blank or title pages with header text found in {pdf_path}")


if __name__ == "__main__":
    strip_misleading_headers(sys.argv[1])
