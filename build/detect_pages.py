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
        page = doc[i]
        raw_text = page.get_text()
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
            # An exact standalone-line match alone isn't enough: "How to
            # Use This Book"'s own Folder Structure table lists nearly
            # every other main section's name as its own isolated table
            # cell -- confirmed directly as a real, severe bug (not
            # theoretical): "Reading Paths for Different Readers" and
            # eight other main section targets all falsely matched a
            # cell in that table instead of their real, much later
            # heading, and because this function's own cursor advances
            # to wherever the *previous* target was found, one false
            # match let the next search start from the same wrong page
            # and often falsely match again, cascading through most of
            # the table. Every heading in this book, at every level, is
            # rendered bold; the table's own cell text never is
            # (confirmed directly: LiberationSans-Bold 16pt for a real
            # H1 vs. plain LiberationSans 11pt for the same words as a
            # table cell) -- so this uses get_text("dict") to also
            # require the matched line's own font name to contain
            # "Bold", which a table cell's text never does regardless of
            # which page it appears on.
            d = page.get_text("dict")
            page_lines = []  # [(text, is_bold), ...], in reading order
            for block in d.get("blocks", []):
                for line in block.get("lines", []):
                    spans = line.get("spans", [])
                    if not spans:
                        continue
                    line_text = "".join(s["text"] for s in spans)
                    is_bold = any("bold" in s["font"].lower() for s in spans)
                    page_lines.append((line_text, is_bold))
            # Drop the running header the same way as above -- it's
            # always the page's own first line.
            body_page_lines = page_lines[1:] if page_lines else []

            consumed = 0
            window_lines = []
            for line_text, is_bold in body_page_lines:
                norm = normalize_whitespace(normalize_quotes(line_text))
                if consumed >= LINE_MATCH_WINDOW_CHARS and window_lines:
                    break
                window_lines.append((norm, is_bold))
                consumed += len(norm) + 1

            for norm, is_bold in window_lines:
                if norm == target and is_bold:
                    return i
            # A heading long enough to need a raised line_match_max_len
            # is also long enough to legitimately wrap across two
            # rendered lines (confirmed directly: a 91-character
            # References for Further Reading category heading wraps
            # after "Character," onto its own second line) -- so also
            # try each consecutive pair of lines joined by a space,
            # still requiring the *pair* to equal the target exactly
            # (not a substring match) and both lines to be bold, before
            # falling through to the softer substring check below.
            for j in range(len(window_lines) - 1):
                text_a, bold_a = window_lines[j]
                text_b, bold_b = window_lines[j + 1]
                joined = (text_a + " " + text_b).strip()
                if target == joined and bold_a and bold_b:
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
