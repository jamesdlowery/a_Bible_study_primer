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

Implements 1-3 for every main section (using each section's own page
range, from the PDF's bookmarks), including 3c for every section that
has sub-sections, via three different sources for the sub-entries
(whichever fits how that section is actually built):
  - Three sections with existing per-entry target data (Manuscript and
    Translation Differences, Reportedly Contradicting Passages,
    Histories of English Bible Translations).
  - Two sections that are a single target with a numbered-list
    structure of their own (Major U.S. Christian Denominations,
    Prominent English Study Bibles) -- their entries (e.g. "1. The
    Catholic Church") are parsed directly out of their one source file.
  - The remaining sections with their own "##" sub-headings (Purpose &
    Scope, How to Use This Book, Biblical Source Manuscripts, etc.),
    also parsed directly out of source. One of these (Manuscript
    Traditions: Character, Relationships, and Weighing) has a "###"
    sub-sub-heading nested under one of its "##" sub-headings; per the
    specified rule, pages within that sub-sub-heading's own range show
    the pair joined by an em dash ("Principles of textual criticism
    that govern how scholars weigh them \u2014 A worked example: weighing
    the evidence for Mark's ending (Mark 16:9-20)") as the guide word,
    rather than either title alone.
  - "Bible Translations and Their Source Manuscripts" has no "##"
    headings of its own at all (a single continuous table), so it
    correctly gets 3a/3b only, no guide word -- there's nothing to be
    one.

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


HEADER_GAP = 18  # minimum points required between the two header pieces


def insert_header_text(page, text, align):
    """align: 'left' or 'right'."""
    text = smarten_quotes(text)
    page.insert_font(fontname=FONT_ALIAS, fontfile=FONT_FILE)
    width = _HEADER_FONT.text_length(text, fontsize=HEADER_SIZE)
    x = LEFT_X if align == "left" else (RIGHT_X - width)
    page.insert_text((x, HEADER_BASELINE_Y), text, fontname=FONT_ALIAS, fontfile=FONT_FILE,
                      fontsize=HEADER_SIZE, color=HEADER_COLOR)


def _truncate_to_width(text, max_width):
    """Shortens text with a trailing ellipsis until it fits max_width at
    the header font/size, cutting at the last space before the limit
    where possible (rather than mid-word)."""
    if _HEADER_FONT.text_length(smarten_quotes(text), fontsize=HEADER_SIZE) <= max_width:
        return text
    ellipsis = "\u2026"
    ellipsis_width = _HEADER_FONT.text_length(ellipsis, fontsize=HEADER_SIZE)
    lo, hi = 0, len(text)
    best = ""
    while lo <= hi:
        mid = (lo + hi) // 2
        candidate = text[:mid].rstrip()
        w = _HEADER_FONT.text_length(smarten_quotes(candidate), fontsize=HEADER_SIZE) + ellipsis_width
        if w <= max_width:
            best = candidate
            lo = mid + 1
        else:
            hi = mid - 1
    # Prefer cutting at the last full word if that doesn't lose much.
    last_space = best.rfind(" ")
    if last_space > len(best) * 0.6:
        best = best[:last_space]
    return best + ellipsis


def insert_header_pair(page, main_text, main_align, guide_text, guide_align):
    """Inserts the main title and (if present) the guide word together,
    truncating the guide word with an ellipsis first if the two would
    otherwise overlap -- confirmed directly as a real problem: the one
    combined "sub-section — sub-sub-section" guide word in the book
    (Manuscript Traditions' worked-example pairing) is long enough that,
    even right-aligned, it starts well to the left of where the main
    title ends on the same page, overlapping it outright. The main
    title is never touched -- it's always short enough to fit alone (the
    longest is confirmed under half the available width) -- only the
    guide word, which can run arbitrarily long once two titles are
    joined, ever needs shortening."""
    main_text = smarten_quotes(main_text)
    main_width = _HEADER_FONT.text_length(main_text, fontsize=HEADER_SIZE)
    insert_header_text(page, main_text, main_align)
    if not guide_text:
        return
    available = (RIGHT_X - LEFT_X) - main_width - HEADER_GAP
    guide_text = _truncate_to_width(guide_text, available)
    insert_header_text(page, guide_text, guide_align)


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


def parse_h2_h3_headings(source_path):
    """Extract "## " (sub-section) and "### " (sub-sub-section) headings,
    in order, from a source markdown file. Returns a list of
    (search_text, display_text) pairs:
      - a plain "## Name" heading with no "###" children of its own:
        search_text == display_text == "Name"
      - a "### Sub-name" nested under a "## Name": search_text is just
        "Sub-name" (that's the only text that actually appears verbatim
        in the compiled document -- the pair is never literally written
        out anywhere), but display_text is "Name — Sub-name", the joined
        pair specified for a sub-section that itself has sub-sub-
        headings. The parent "## Name" heading's own entry -- covering
        the pages between the H2 itself and its first H3 child -- is
        still included as its own plain, unpaired entry beforehand.
    "###" headings not nested under any "##" (shouldn't happen in this
    book's own structure, but handled safely) are skipped."""
    entries = []
    current_h2 = None
    with open(source_path, encoding="utf-8") as f:
        for line in f:
            line = line.rstrip("\n")
            m2 = re.match(r"^## (.+)$", line)
            if m2:
                current_h2 = m2.group(1)
                entries.append((current_h2, current_h2))
                continue
            m3 = re.match(r"^### (.+)$", line)
            if m3 and current_h2:
                h3 = m3.group(1)
                entries.append((h3, f"{current_h2} \u2014 {h3}"))
    return entries


def parse_section_headings(source_path):
    """Extract "## Title" sub-section headings, in document order, from a
    source markdown file -- for sections with plain (non-numbered) H2
    sub-headings rather than the "## N. Name" list pattern
    parse_numbered_headings() handles. If a "##" sub-section itself has
    "###" sub-sub-headings nested under it (confirmed directly: exactly
    one case in the whole book, "Principles of textual criticism..." in
    Manuscript Traditions, with "### A worked example: weighing the
    evidence for Mark's ending" under it), each "###" becomes its own
    entry pairing both titles with an em dash for display ("Principles...
    -- A worked example..."), per the specified recursive pairing rule --
    but the page is located by searching for the "###" text alone, since
    that's what actually appears in the rendered document; the combined
    pair never does. Any part of the "##" section's own content before
    its first "###" child still uses the plain, unpaired "##" title
    (searched and displayed the same way) on its own.
    Returns [(search_text, display_text), ...] in document order."""
    headings = []
    current_h2 = None
    with open(source_path, encoding="utf-8") as f:
        for line in f:
            line = line.rstrip("\n")
            m2 = re.match(r"^## (.+)$", line)
            m3 = re.match(r"^### (.+)$", line)
            if m2:
                current_h2 = m2.group(1)
                headings.append((current_h2, current_h2))
            elif m3 and current_h2:
                headings.append((m3.group(1), f"{current_h2} \u2014 {m3.group(1)}"))
    return headings


def build_sub_entries_from_headings(doc, headings, start_0idx, end_0idx):
    """Same idea as build_sub_entries(), but for a list of (search_text,
    display_text) pairs (from parse_numbered_headings/parse_section_headings)
    rather than target metadata. parse_numbered_headings's entries are
    plain strings, used as both search and display text; normalize those
    here too."""
    cursor = start_0idx
    found = []
    for h in headings:
        search_text, display_text = h if isinstance(h, tuple) else (h, h)
        # A higher line_match_max_len than the default 40: confirmed
        # directly that several of these headings (numbered-list ones up
        # to 59 characters, e.g. "8. Church of Jesus Christ of
        # Latter-day Saints (LDS/Mormon)"; plain "##"/"###" search text
        # up to 124, e.g. the Reading Paths chapter's own question-style
        # sub-headings -- the em-dash pairing is for display only and
        # never itself searched, so it doesn't factor into this) would
        # otherwise fall into the plainer, narrower-window substring-
        # match path instead of the more robust standalone-line match,
        # and be just as vulnerable to a section's own long intro
        # pushing them far down the page as the shorter ones already
        # were.
        page = _dp_find_page(doc, search_text, cursor, line_match_max_len=150)
        if page is None or page > end_0idx:
            continue
        found.append((page, display_text))
        cursor = page
    return found


def build_sub_entries_from_pairs(doc, entries, start_0idx, end_0idx):
    """Like build_sub_entries_from_headings(), but for the (search_text,
    display_text) pairs parse_h2_h3_headings() returns -- searches for
    search_text (what's actually written in the document) but records
    display_text (the possibly-joined pair) as the guide word."""
    cursor = start_0idx
    found = []
    for search_text, display_text in entries:
        # A still-higher line_match_max_len than the numbered-heading
        # case: confirmed directly that some of this book's own "##"
        # sub-headings run well past 80 characters (up to 118, e.g. a
        # References for Further Reading category heading spanning three
        # section names), and without this they were silently dropped
        # entirely (fell into the substring-match path and weren't
        # found), not just mis-positioned.
        page = _dp_find_page(doc, search_text, cursor, line_match_max_len=150)
        if page is None or page > end_0idx:
            continue
        found.append((page, display_text))
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
    # Sections with plain "## Title" sub-headings (not the numbered-list
    # pattern above) -- parsed with parse_section_headings(), which also
    # handles the one case (Manuscript Traditions) where a "##"
    # sub-section has its own "###" sub-sub-heading, per the specified
    # recursive em-dash pairing rule. "Bible Translations and Their
    # Source Manuscripts" is deliberately not listed here -- confirmed
    # directly it has no "##"/"###" headings of its own at all (a single
    # flat table), so it gets no guide word, same as before.
    PLAIN_HEADING_SOURCES = {
        "Purpose & Scope":
            os.path.join(repo_root, "030 Introduction", "020 Purpose and Scope.md"),
        "How to Use This Book":
            os.path.join(repo_root, "030 Introduction", "050 How to Use This Book.md"),
        "Reading Paths for Different Readers":
            os.path.join(repo_root, "030 Introduction", "060 Reading Paths for Different Readers.md"),
        "What Is Meant by the \u201cWord of God\u201d?":
            os.path.join(repo_root, "030 Introduction", "030 What Is Meant by the Word of God.md"),
        "What Is Meant by an \u201cInerrant\u201d Word of God?":
            os.path.join(repo_root, "030 Introduction", "040 What Is Meant by an Inerrant Word of God.md"),
        "Background on Textual Transmission":
            os.path.join(repo_root, "030 Introduction", "070 Background on Textual Transmission.md"),
        "A Note on Method and Verification":
            os.path.join(repo_root, "030 Introduction", "080 A Note on Method and Verification.md"),
        "Biblical Source Manuscripts":
            os.path.join(repo_root, "040 Biblical Source Manuscripts", "010 Biblical Source Manuscripts.md"),
        "Manuscript Traditions: Character, Relationships, and Weighing":
            os.path.join(repo_root, "050 Character Of Each Source Manuscript Tradition", "010 Character of Each Tradition.md"),
        "Popular Bible Translations":
            os.path.join(repo_root, "060 Popular Bible Translations", "010 Popular Bible Translations.md"),
        "References for Further Reading":
            os.path.join(repo_root, "130 References for Further Reading", "010 References for Further Reading.md"),
    }
    # The remaining front-matter (and similar) sections: no per-entry
    # target data and no numbered-list structure, but each has its own
    # "##" sub-headings (some with "###" sub-sub-headings beneath them,
    # handled via the search/display pairing in parse_h2_h3_headings()).
    # "Bible Translations and Their Source Manuscripts" is deliberately
    # not listed here: confirmed it has no "##" headings of its own at
    # all (a single continuous table), so it correctly gets no guide
    # word, matching rule 3c's "if the section has sub-sections" --
    # this one doesn't.
    H2_H3_SOURCES = {
        "Purpose & Scope":
            os.path.join(repo_root, "030 Introduction", "020 Purpose and Scope.md"),
        "How to Use This Book":
            os.path.join(repo_root, "030 Introduction", "050 How to Use This Book.md"),
        "Reading Paths for Different Readers":
            os.path.join(repo_root, "030 Introduction", "060 Reading Paths for Different Readers.md"),
        "What Is Meant by the \u201cWord of God\u201d?":
            os.path.join(repo_root, "030 Introduction", "030 What Is Meant by the Word of God.md"),
        "What Is Meant by an \u201cInerrant\u201d Word of God?":
            os.path.join(repo_root, "030 Introduction", "040 What Is Meant by an Inerrant Word of God.md"),
        "Background on Textual Transmission":
            os.path.join(repo_root, "030 Introduction", "070 Background on Textual Transmission.md"),
        "A Note on Method and Verification":
            os.path.join(repo_root, "030 Introduction", "080 A Note on Method and Verification.md"),
        "Biblical Source Manuscripts":
            os.path.join(repo_root, "040 Biblical Source Manuscripts", "010 Biblical Source Manuscripts.md"),
        "Manuscript Traditions: Character, Relationships, and Weighing":
            os.path.join(repo_root, "050 Character Of Each Source Manuscript Tradition", "010 Character of Each Tradition.md"),
        "Popular Bible Translations":
            os.path.join(repo_root, "060 Popular Bible Translations", "010 Popular Bible Translations.md"),
        "References for Further Reading":
            os.path.join(repo_root, "130 References for Further Reading", "010 References for Further Reading.md"),
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
        elif title in PLAIN_HEADING_SOURCES:
            headings = parse_section_headings(PLAIN_HEADING_SOURCES[title])
            sub_entries = build_sub_entries_from_headings(doc, headings, start, end)
        elif title in H2_H3_SOURCES:
            pairs = parse_h2_h3_headings(H2_H3_SOURCES[title])
            sub_entries = build_sub_entries_from_pairs(doc, pairs, start, end)

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
            gw = None
            if sub_entries:
                want = "last" if odd else "first"
                gw = guide_word_for_page(sub_entries, pno, want)
            insert_header_pair(page, title, main_align, gw, guide_align)
            touched += 1

    tmp_path = pdf_path + ".tmp"
    doc.save(tmp_path, garbage=4, deflate=True)
    doc.close()
    os.replace(tmp_path, pdf_path)
    print(f"Regenerated headers on {touched} page(s) in {pdf_path}")


if __name__ == "__main__":
    generate_headers(sys.argv[1], sys.argv[2], sys.argv[3])
