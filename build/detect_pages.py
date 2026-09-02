#!/usr/bin/env python3
"""
Detect the actual rendered page number of each target (from
targets_meta_<fmt>.json) in a rendered PDF, then compute which targets need
an extra blank page inserted before them so every target lands on an odd
page number. Writes blanks_<fmt>.json: a JSON array of target ids.

Usage: python3 detect_pages.py <fmt> <pdf_path> <meta_json_path> <out_blanks_json_path>
"""
import sys
import re
import json
import fitz  # PyMuPDF


def normalize_quotes(s):
    """Pandoc's smart-typography converts straight quotes to curly ones in
    output, so normalize both sides before substring matching."""
    return (s.replace("\u2019", "'").replace("\u2018", "'")
             .replace("\u201c", '"').replace("\u201d", '"'))


def normalize_whitespace(s):
    """Longer headings can wrap across two lines in the rendered PDF, which
    PyMuPDF represents as an embedded newline right where the line wraps.
    Collapse all whitespace runs (including those newlines) to single spaces
    so a heading that happens to wrap doesn't break the substring match."""
    return re.sub(r'\s+', ' ', s).strip()


def find_page(doc, search_text, start_page):
    """Forward-only search for search_text starting at start_page (0-indexed).
    Returns the 0-indexed page number where it's first found at or after
    start_page, or None if not found."""
    target = normalize_whitespace(normalize_quotes(search_text))
    for i in range(start_page, len(doc)):
        page_text = normalize_whitespace(normalize_quotes(doc[i].get_text()))
        if target in page_text:
            return i
    return None


def main():
    fmt, pdf_path, meta_path, out_path = sys.argv[1:5]

    with open(meta_path, encoding="utf-8") as f:
        meta = json.load(f)

    doc = fitz.open(pdf_path)

    cursor = 0
    raw_pages = {}
    for m in meta:
        page = find_page(doc, m["search_text"], cursor)
        if page is None:
            print(f"WARNING: could not locate target '{m['id']}' "
                  f"(search text: {m['search_text']!r}) anywhere from page "
                  f"{cursor + 1} onward. Leaving its position unadjusted.")
            continue
        raw_pages[m["id"]] = page + 1  # 1-indexed page number
        cursor = page

    # Walk targets in order, accumulating shift from earlier insertions.
    needs_blank = []
    shift = 0
    for m in meta:
        if m["id"] not in raw_pages:
            continue
        effective = raw_pages[m["id"]] + shift
        if effective % 2 == 0:
            needs_blank.append(m["id"])
            shift += 1

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(needs_blank, f, indent=2)

    print(f"[{fmt}] Checked {len(meta)} targets, "
          f"{len(raw_pages)} located, {len(needs_blank)} need a blank page.")
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    main()
