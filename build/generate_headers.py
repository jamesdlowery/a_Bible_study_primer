#!/usr/bin/env python3
"""Build per-page running headers for the compiled PDF, replacing the
book's blanket STYLEREF-based header entirely. Rules (as specified):

1. Blank pages: no header (already handled by the existing blank-page
   detection).
2. A main section that begins and ends on the same (content) page: no
   header at all.
3. A main section spanning multiple content pages:
   a. No header on its own first page.
   b. On every subsequent page: the main section's own title, left-
      aligned on odd pages, right-aligned on even pages.
   c. If the section has sub-sections: also show a "guide word" --the
      last sub-section (or sub-section -- sub-sub-section pair) that
      starts on that page, on odd pages, right-aligned; the first such
      pair that starts on that page, on even pages, left-aligned. (I.e.
      the guide word is aligned opposite to the main title.)

Phase boundary, stated plainly: this implements 1-3 for every main
section (using each section's own page range, from the PDF's bookmarks);
3c for the three sections with existing per-entry target data
(Manuscript and Translation Differences, Reportedly Contradicting
Passages, Histories of English Bible Translations); and 3c for Major
U.S. Christian Denominations and Prominent English Study Bibles, whose
27 numbered entries each (e.g. "## 1. The Catholic Church") are parsed
directly from their own source markdown file rather than from target
metadata (they're each a single target with no per-entry data of their
own) -- together covering the large majority of the book's content
pages. It does NOT yet implement 3c for the front-matter sections with
their own internal sub-headings (Purpose & Scope, How to Use This Book,
etc.), including the specified recursive em-dash pairing for a
sub-section that itself has sub-sub-headings -- that needs font-based
heading-level detection that hasn't been built/calibrated yet.

Usage: python3 generate_headers.py <pdf_path> <targets_meta_json_path> <repo_root>
Edits the PDF in place.
"""
import sys
import os
import re
import json
import fitz  # PyMuPDF

sys.path.insert(0, os.path.dirname(__file__))
from detect_pages import find_page as _dp_find_page

HEADER_Y_MAX = 60
FOOTER_Y_MIN = 750
HEADER_BASELINE_Y = 43.5
LEFT_X = 54.1
RIGHT_X = 558.0
HEADER_SIZE = 8.0
HEADER_COLOR = (0, 0, 0)
# The exact font the book's existing header already uses, extracted from
# the compiled document itself (word/embedded, not a PDF base-14 name) --
# confirmed necessary: a base-14 standard font ("Helvetica-Oblique")
# doesn't have proper glyphs for the curly quotes several section titles
# use (e.g. "What Is Meant by the "Word of God"?"), rendering them as a
# fallback middle-dot character instead.
FONT_FILE = os.path.join(os.path.dirname(__file__), "LiberationSans-Italic.ttf")
FONT_ALIAS = "hdrfont"


def is_blank_page(page):
    d = page.get_text("dict")
    body = []
    for block in d.get("blocks", []):
        for line in block.get("lines", []):
            for span in line["spans"]:
                y0 = span["bbox"][1]
                if y0 < HEADER_Y_MAX or y0 > FOOTER_Y_MIN:
                    continue
                body.append(span["text"])
    return "".join(body).replace("\xa0", "").strip() == ""


def clear_existing_header(page):
    d = page.get_text("dict")
    for block in d.get("blocks", []):
        for line in block.get("lines", []):
            for span in line["spans"]:
                if span["bbox"][1] < HEADER_Y_MAX:
                    page.add_redact_annot(span["bbox"])
    page.apply_redactions()


_HEADER_FONT = fitz.Font(fontfile=FONT_FILE)


def smarten_quotes(s):
    """Convert straight quotes/apostrophes to curly, matching the rest of
    the book's own typography. Needed because targets_meta_docx.json's
    search_text values are copied from source markdown before pandoc's
    own smart-quote pass runs, so they still have straight ones (e.g.
    "Young's" from the King James history chapter) -- confirmed directly:
    without this, inserted header/guide-word text was the only straight-
    quoted text left in an otherwise fully curly-quoted document."""
    out = []
    double_open = True
    for i, ch in enumerate(s):
        if ch == '"':
            out.append('\u201c' if double_open else '\u201d')
            double_open = not double_open
        elif ch == "'":
            prev_alpha = i > 0 and s[i - 1].isalpha()
            next_alpha = i + 1 < len(s) and s[i + 1].isalpha()
            if prev_alpha:
                out.append('\u2019')
            elif not prev_alpha and next_alpha:
                out.append('\u2018')
            else:
                out.append('\u2019')
        else:
            out.append(ch)
    return "".join(out)


def insert_header_text(page, text, align):
    """align: 'left' or 'right'."""
    text = smarten_quotes(text)
    page.insert_font(fontname=FONT_ALIAS, fontfile=FONT_FILE)
    width = _HEADER_FONT.text_length(text, fontsize=HEADER_SIZE)
    x = LEFT_X if align == "left" else (RIGHT_X - width)
    page.insert_text((x, HEADER_BASELINE_Y), text, fontname=FONT_ALIAS, fontfile=FONT_FILE,
                      fontsize=HEADER_SIZE, color=HEADER_COLOR)


def build_main_section_ranges(doc):
    """[(title, start_page_0idx, end_content_page_0idx_inclusive), ...],
    trimming trailing blank pages from each section's raw bookmark-to-
    next-bookmark span so the range reflects real content only."""
    toc = doc.get_toc()
    mains = [(title, page - 1) for level, title, page in toc
             if level == 1 and title not in ("Front Cover",)]
    ranges = []
    for i, (title, start) in enumerate(mains):
        raw_end = (mains[i + 1][1] - 1) if i + 1 < len(mains) else len(doc) - 1
        end = raw_end
        while end > start and is_blank_page(doc[end]):
            end -= 1
        ranges.append((title, start, end))
    return ranges


def build_sub_entries(doc, meta, prefixes, start_0idx, end_0idx):
    """For a section with per-entry target ids (identified by any of
    `prefixes`), find each entry's own 0-indexed page within
    [start_0idx, end_0idx], using the same robust search as
    detect_pages.py. Returns [(page_0idx, title), ...] in page order."""
    entries = [m for m in meta if m["id"].startswith(prefixes)]
    cursor = start_0idx
    found = []
    for m in entries:
        page = _dp_find_page(doc, m["search_text"], cursor)
        if page is None or page > end_0idx:
            continue
        found.append((page, m["search_text"]))
        cursor = page
    return found


def parse_numbered_headings(source_path):
    """Extract "## N. Name" headings, in order, from a source markdown
    file -- used for sections (Denominations, Study Bibles) that are a
    single target with no per-entry target data of their own. Keeps the
    number prefix, since that's how the heading actually renders in the
    compiled document (confirmed directly: "1. The Catholic Church", not
    just "The Catholic Church")."""
    headings = []
    with open(source_path, encoding="utf-8") as f:
        for line in f:
            m = re.match(r"^## (\d+\.\s.+)$", line.rstrip("\n"))
            if m:
                headings.append(m.group(1))
    return headings


def build_sub_entries_from_headings(doc, headings, start_0idx, end_0idx):
    """Same idea as build_sub_entries(), but for a list of plain heading
    strings (from parse_numbered_headings) rather than target metadata."""
    cursor = start_0idx
    found = []
    for h in headings:
        # A higher line_match_max_len than the default 40: confirmed
        # directly that several of these numbered headings (up to 59
        # characters, e.g. "8. Church of Jesus Christ of Latter-day
        # Saints (LDS/Mormon)") would otherwise fall into the plainer,
        # narrower-window substring-match path instead of the more
        # robust standalone-line match, and be just as vulnerable to
        # this section's own long multi-paragraph intro pushing them
        # far down the page as the shorter ones already were.
        page = _dp_find_page(doc, h, cursor, line_match_max_len=80)
        if page is None or page > end_0idx:
            continue
        found.append((page, h))
        cursor = page
    return found


def guide_word_for_page(sub_entries, page_0idx, want):
    """want: 'first' or 'last'. Among sub-entries whose own page is
    <= page_0idx (i.e. already started by this page), returns the
    first-in-document-order or last-in-document-order title that starts
    ON this exact page, if any; otherwise the most recent one before it
    (still in effect on this page, just not itself starting here)."""
    on_this_page = [t for p, t in sub_entries if p == page_0idx]
    if on_this_page:
        return on_this_page[0] if want == "first" else on_this_page[-1]
    prior = [t for p, t in sub_entries if p < page_0idx]
    return prior[-1] if prior else None


def generate_headers(pdf_path, meta_path, repo_root):
    doc = fitz.open(pdf_path)
    meta = json.load(open(meta_path, encoding="utf-8"))

    ranges = build_main_section_ranges(doc)

    # Sections with known per-entry target data (id prefix match).
    STRUCTURED = {
        "Manuscript and Translation Differences": ("ot_", "apoc_", "nt_"),
        "Reportedly Contradicting Passages": ("rcpbook_",),
        "Histories of English Bible Translations": ("history_",),
    }
    # Sections that are a single target with no per-entry target data of
    # their own, but whose entries can be parsed directly out of their
    # one source file as "## N. Name" headings.
    NUMBERED_HEADING_SOURCES = {
        "Major U.S. Christian Denominations":
            os.path.join(repo_root, "110 Top Christian Denominations", "010 Top Christian Denominations.md"),
        "Prominent English Study Bibles":
            os.path.join(repo_root, "120 Top Study Bibles", "010 Top Study Bibles.md"),
    }

    touched = 0
    for title, start, end in ranges:
        if start == end:
            # Rule 2: single content page, no header at all (nothing to do,
            # header already cleared or never had one -- but ensure clean).
            if not is_blank_page(doc[start]):
                clear_existing_header(doc[start])
                touched += 1
            continue

        sub_entries = []
        if title in STRUCTURED:
            sub_entries = build_sub_entries(doc, meta, STRUCTURED[title], start, end)
        elif title in NUMBERED_HEADING_SOURCES:
            headings = parse_numbered_headings(NUMBERED_HEADING_SOURCES[title])
            sub_entries = build_sub_entries_from_headings(doc, headings, start, end)

        for pno in range(start, end + 1):
            page = doc[pno]
            if is_blank_page(page):
                continue  # handled separately by blank_page_headers.py
            clear_existing_header(page)
            if pno == start:
                touched += 1
                continue  # Rule 3a: no header on the section's own first page
            odd = (pno + 1) % 2 == 1  # 1-indexed page number odd?
            main_align = "left" if odd else "right"
            guide_align = "right" if odd else "left"
            insert_header_text(page, title, main_align)
            if sub_entries:
                want = "last" if odd else "first"
                gw = guide_word_for_page(sub_entries, pno, want)
                if gw:
                    insert_header_text(page, gw, guide_align)
            touched += 1

    tmp_path = pdf_path + ".tmp"
    doc.save(tmp_path, garbage=4, deflate=True)
    doc.close()
    os.replace(tmp_path, pdf_path)
    print(f"Regenerated headers on {touched} page(s) in {pdf_path}")


if __name__ == "__main__":
    generate_headers(sys.argv[1], sys.argv[2], sys.argv[3])
