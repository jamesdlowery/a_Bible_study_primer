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


HEADING_WINDOW_CHARS = 300

# A much larger budget, used only for the exact-standalone-line check
# (short/bare search text). Folded-in divider content or a long section
# intro ahead of a target can push a heading well past 300 characters
# into its own page; confirmed directly for "Genesis" (~900 characters of
# divider text ahead of it), "Tobit" (similarly, the Apocrypha block's
# own canon-status note), and the first numbered entry in Major U.S.
# Christian Denominations ("1. The Catholic Church", pushed to character
# offset 3186 by that section's own multi-paragraph intro -- the reason
# this was raised from 2000 to 4000). A wide window carries much less
# false-positive risk here than it would for the substring check below,
# since ordinary prose only very rarely puts a single short word entirely
# alone on its own line.
LINE_MATCH_WINDOW_CHARS = 4000


def find_page(doc, search_text, start_page, line_match_max_len=40):
    """Forward-only search for search_text starting at start_page (0-indexed).
    Returns the 0-indexed page number where it's first found at or after
    start_page, or None if not found.

    A plain "does this substring appear anywhere on the page" search was
    found to produce real false positives, each confirmed directly against
    actual rendered pages and fixed here in three layers:
      1. Restrict matching to near the top of the page (HEADING_WINDOW_CHARS),
         since every target starts at the top of a fresh page (via
         PAGEBREAK) -- "How to Use This Book"'s own Folder Structure table
         lists every other section's name as plain reference text, which an
         unrestricted search matched hundreds of characters into that
         table's page, long before those sections' real headings.
      2. Drop the page's own running header line before matching -- a bare,
         short search text (Reportedly Contradicting Passages' book targets
         use just the plain book name, e.g. "Exodus") can match that name
         inside an unrelated *earlier* chapter's own running header.
      3. For search text short enough for this to matter, require an exact
         standalone-line match rather than a substring match at all --
         even with the header dropped, a common word like "Genesis" (almost
         certainly this book's single most cross-referenced word) still
         turns up constantly in ordinary prose near the top of unrelated
         pages. A heading, uniquely, renders as its own complete line with
         nothing else on it. This check uses the wider
         LINE_MATCH_WINDOW_CHARS budget, not HEADING_WINDOW_CHARS, since
         folded-in divider content immediately ahead of a target (a
         testament/category label, or a longer bridge paragraph) can push
         its heading well past 300 characters into the page.
    """
    target = normalize_whitespace(normalize_quotes(search_text))
    for i in range(start_page, len(doc)):
        raw_text = doc[i].get_text()
        # Drop the running header -- its own first line, e.g. "Manuscript
        # and Translation Differences — Exodus". Short, generic search
        # texts (the Reportedly Contradicting Passages book targets use
        # just the bare book name, e.g. "Exodus" or "Genesis") can match
        # that header on pages that aren't the target's own start at all --
        # confirmed directly for "Exodus" (present in the running header
        # of every page of the earlier, unrelated Manuscript and
        # Translation Differences Exodus chapter).
        body_lines = raw_text.split("\n")[1:] if "\n" in raw_text else [raw_text]
        body_text = "\n".join(body_lines)
        page_text = normalize_whitespace(normalize_quotes(body_text))
        window = page_text[:HEADING_WINDOW_CHARS]

        # For a short, generic search text (a bare book name like
        # "Genesis" is exactly this project's worst case -- almost
        # certainly the single most cross-referenced word in the entire
        # book), a plain substring check isn't reliable even after
        # dropping the header: ordinary prose mentions it constantly
        # ("...as seen in Genesis 3:15...") within the first 300
        # characters of plenty of pages that aren't its own heading.
        # A heading, uniquely, renders as its own complete line with
        # nothing else on it -- so for search text short enough for that
        # distinction to be meaningful, require an exact line match
        # first, only falling back to the softer substring check (needed
        # for longer titles that might legitimately wrap across two
        # rendered lines) when no such standalone line exists on this
        # page at all.
        if len(target) <= line_match_max_len:
            # Use the same HEADING_WINDOW_CHARS budget as the substring
            # check below, not a fixed line count -- a divider's own
            # content (a testament/category label, or a longer bridge
            # paragraph like Reportedly Contradicting Passages' own
            # reminder note) can fold onto the same page as the very next
            # target and push its heading down several lines; confirmed
            # directly for "Genesis" (pushed to line 12 of the page by a
            # 9-line divider paragraph ahead of it) and "Tobit" (pushed
            # similarly by the Apocrypha block's own canon-status note).
            consumed = 0
            body_window_lines = []
            for l in body_lines:
                norm = normalize_whitespace(normalize_quotes(l))
                if consumed >= LINE_MATCH_WINDOW_CHARS and body_window_lines:
                    break
                body_window_lines.append(norm)
                consumed += len(norm) + 1
            if target in body_window_lines:
                return i
            continue

        if target in window:
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
