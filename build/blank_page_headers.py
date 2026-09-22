#!/usr/bin/env python3
"""Remove the running-header text from every blank page in the compiled
PDF, leaving the watermark and footer untouched.

Blank pages carry a running header the same way every other page does
(via a STYLEREF field that shows the nearest heading's text), which is
misleading on a page with no real content of its own -- it just repeats
whatever section was active on the page before it. The cover's own three
post-cover blank pages already get a genuinely blank header at the DOCX
level (see build_cover_md()'s header3.xml in custom-reference.docx), but
the same approach for the ~100 mid-document blank pages inserted
throughout the book to enforce odd-page starts was tried and reverted:
each would need its own pair of OOXML section breaks, and confirmed
directly, adding the roughly 200 extra section boundaries that requires
made LibreOffice's field-recalculation step in update_fields_and_export.py
hang for minutes with no output.

This takes a completely different, much cheaper approach on the compiled
PDF itself instead of the DOCX: find every page whose only real content
is the blank-page placeholder (a bare non-breaking space), and redact
just the header-region text on those pages. This works safely because
the watermark is confirmed to not be extractable text at all (a VML
shape, invisible to get_text()), and the header text is confirmed to sit
in its own narrow band near the very top of the page (roughly y 36-45),
well above where any real body content starts (confirmed at y >= ~115
even on a heading-only page) -- so a conservative y-position cutoff can
only ever catch the header, never real content or the footer (near the
bottom of the page, y >= ~750).

Usage: python3 blank_page_headers.py <pdf_path>
Edits the PDF in place.
"""
import sys
import fitz  # PyMuPDF

HEADER_REGION_Y_MAX = 60  # points from the top; header text confirmed ~36-45


def is_blank_page(page):
    """A blank page's only real content is a bare non-breaking space (the
    placeholder every blank page in this book is built from); the header
    and footer are the only other text on it."""
    text = page.get_text().strip()
    # Strip out anything that looks like header/footer text by position
    # instead of by content (content varies -- section name, page number,
    # URL, version -- but position doesn't), then see what's left.
    d = page.get_text("dict")
    leftover = []
    for block in d.get("blocks", []):
        for line in block.get("lines", []):
            for span in line["spans"]:
                y0 = span["bbox"][1]
                if y0 < HEADER_REGION_Y_MAX or y0 > 750:
                    continue  # header or footer region, not body
                leftover.append(span["text"])
    body = "".join(leftover).replace("\xa0", "").strip()
    return body == ""


def strip_blank_page_headers(pdf_path):
    doc = fitz.open(pdf_path)
    redacted_pages = []
    for page in doc:
        if not is_blank_page(page):
            continue
        d = page.get_text("dict")
        for block in d.get("blocks", []):
            for line in block.get("lines", []):
                for span in line["spans"]:
                    if span["bbox"][1] < HEADER_REGION_Y_MAX:
                        page.add_redact_annot(span["bbox"])
        page.apply_redactions()
        redacted_pages.append(page.number + 1)  # 1-indexed for the log

    if redacted_pages:
        tmp_path = pdf_path + ".tmp"
        doc.save(tmp_path, garbage=4, deflate=True)
        doc.close()
        import os
        os.replace(tmp_path, pdf_path)
        print(f"Removed header text from {len(redacted_pages)} blank page(s) in {pdf_path}")
    else:
        doc.close()
        print(f"No blank pages with header text found in {pdf_path}")


if __name__ == "__main__":
    strip_blank_page_headers(sys.argv[1])
