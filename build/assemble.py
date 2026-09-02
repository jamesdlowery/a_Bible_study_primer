#!/usr/bin/env python3
import re
import os
import sys
import html as html_entities

REPO = "/home/claude/repo"
FORMAT = sys.argv[1] if len(sys.argv) > 1 else "docx"  # "docx" or "odt"
if len(sys.argv) > 2:
    REPO = sys.argv[2]  # path to the checked-out repo root (for CI use)

if FORMAT == "docx":
    PAGEBREAK = '\n\n```{=openxml}\n<w:p><w:r><w:br w:type="page"/></w:r></w:p>\n```\n\n'
elif FORMAT == "odt":
    PAGEBREAK = '\n\n::: {custom-style="PageBreak"}\n\u200B\n:::\n\n'
else:  # html -- no real pagination; just a print-only page break hint
    PAGEBREAK = '\n\n<div class="print-page-break"></div>\n\n'

def fix_malformed_tables(text):
    """Fix tables authored with a blank header row followed by the
    separator row and the real header labels as the first body row:

        |     |     |
        |-----|-----|
        | **A** | **B** |
        | data  | data  |

    Both pandoc's docx and odt writers need the true header row directly
    above the separator to correctly apply header-row styling; only the
    docx writer happens to compensate for this pattern by dropping the
    blank row, so we normalize the source instead of relying on that.
    """
    lines = text.split("\n")
    out = []
    i = 0
    while i < len(lines):
        line = lines[i]
        is_blank_row = (
            line.strip().startswith("|")
            and line.count("|") >= 2
            and set(line.replace("|", "").strip()) <= {""}
        )
        if is_blank_row and i + 2 < len(lines):
            sep = lines[i + 1].strip()
            is_sep = sep.startswith("|") and set(sep.replace("|", "").strip()) <= set("-: ")
            if is_sep:
                header_row = lines[i + 2]
                # emit: real header row, then separator, skipping the blank row
                out.append(header_row)
                out.append(lines[i + 1])
                i += 3
                continue
        out.append(line)
        i += 1
    return "\n".join(out)


def fix_column_widths(text):
    """Recompute pipe-table separator dash counts so column width (which
    pandoc derives from relative dash-count in the separator row) reflects
    actual content, instead of the source's arbitrary dash lengths which
    often left label columns too narrow (e.g. 'Samaritan Pentateuch (SP)'
    wrapping across 3 lines). Longest-content columns are capped so one
    very long paragraph column doesn't squeeze a short label column down
    to unreadable width.
    """
    CAP = 60
    FLOOR = 15
    lines = text.split("\n")
    out = []
    i = 0

    def is_sep(s):
        s = s.strip()
        return s.startswith("|") and s.count("|") >= 2 and set(s.replace("|", "").strip()) <= set("-: ")

    def split_row(row):
        row = row.strip()
        if row.startswith("|"):
            row = row[1:]
        if row.endswith("|"):
            row = row[:-1]
        return row.split("|")

    def clean_len(cell):
        c = cell.strip()
        c = c.replace("**", "").replace("*", "").replace("`", "")
        return len(c)

    while i < len(lines):
        if i + 1 < len(lines) and lines[i].strip().startswith("|") and is_sep(lines[i + 1]):
            header_row = lines[i]
            ncols = len(split_row(header_row))
            # gather body rows until a non-table line
            j = i + 2
            body_rows = []
            while j < len(lines) and lines[j].strip().startswith("|"):
                body_rows.append(lines[j])
                j += 1
            all_rows = [header_row] + body_rows
            max_lens = [FLOOR] * ncols
            for row in all_rows:
                cells = split_row(row)
                for c_idx in range(min(ncols, len(cells))):
                    l = clean_len(cells[c_idx])
                    if l > max_lens[c_idx]:
                        max_lens[c_idx] = l
            capped = [min(l, CAP) for l in max_lens]
            new_sep = "|" + "|".join("-" * n for n in capped) + "|"
            out.append(header_row)
            out.append(new_sep)
            out.extend(body_rows)
            i = j
            continue
        out.append(lines[i])
        i += 1
    return "\n".join(out)


def read(path):
    with open(os.path.join(REPO, path), encoding="utf-8") as f:
        text = fix_malformed_tables(f.read().strip())
        text = fix_column_widths(text)
        return text + "\n"

def demote(text, levels):
    """Demote all ATX headings in text by `levels` (add that many '#')."""
    def repl(m):
        return ("#" * (len(m.group(1)) + levels)) + " " + m.group(2)
    return re.sub(r'^(#{1,6}) (.*)$', repl, text, flags=re.MULTILINE)

# ---------------------------------------------------------------------------
# TARGETS: every item in document order that must (a) start on a fresh page
# and (b) be pushed to the next ODD page (with a blank page inserted before
# it) if it would otherwise land on an even page. This covers: the 9 main
# "Part" sections, the 3 Testament dividers (Old Testament/Apocrypha/New
# Testament), each of the 25 translation histories, and each book in the
# Variants section. Category labels (Pentateuch/Law/Torah, Gospels, etc.)
# are NOT separate targets -- they're prepended to whichever book follows.
# ---------------------------------------------------------------------------

PROTESTANT_HISTORIES = [
    ("060 History of the King James Version.md", "King James Version"),
    ("070 History of Webster's Bible.md", "Webster's Bible"),
    ("080 History of Young's Literal Translation.md", "Young's Literal Translation"),
    ("090 History of Smith's Literal Translation.md", "Smith's Literal Translation"),
    ("100 History of the Darby Bible.md", "Darby Bible"),
    ("110 History of the American Standard Version.md", "American Standard Version"),
    ("120 History of the Revised Standard Version.md", "Revised Standard Version"),
    ("130 History of the Amplified Bible.md", "Amplified Bible"),
    ("140 History of the New American Standard Bible.md", "New American Standard Bible"),
    ("150 History of the New International Version.md", "New International Version"),
    ("160 History of the New King James Version.md", "New King James Version"),
    ("170 History of the Easy to Read Version.md", "Easy-to-Read Version"),
    ("180 History of the New Living Translation.md", "New Living Translation"),
    ("190 History of the American King James Version.md", "American King James Version"),
    ("200 History of the World English Bible.md", "World English Bible"),
    ("210 History of the English Standard Version.md", "English Standard Version"),
    ("220 History of the New English Translation.md", "New English Translation"),
    ("230 History of the Christian Standard Bible.md", "Christian Standard Bible"),
    ("240 History of the Berean Standard Bible.md", "Berean Standard Bible"),
    ("250 History of the Legacy Standard Bible.md", "Legacy Standard Bible"),
]

CATHOLIC_HISTORIES = [
    ("010 History of the Douay Rheims Bible.md", "Douay-Rheims Bible"),
    ("020 History of the New Revised Standard Version Catholic Edition.md", "New Revised Standard Version Catholic Edition"),
    ("030 History of the Revised Standard Version Second Catholic Edition.md", "Revised Standard Version Second Catholic Edition"),
    ("040 History of the Catholic Public Domain Version.md", "Catholic Public Domain Version"),
    ("050 History of the New American Bible Revised Edition.md", "New American Bible Revised Edition"),
]

# Neither Protestant nor Catholic (a Jewish translation of the Tanakh, no New
# Testament) -- kept as its own group rather than shoehorned into either.
OTHER_HISTORIES = [
    ("260 History of the Jewish Publication Society Bible.md", "Jewish Publication Society Bible"),
]

VARIANTS_DIR = "100 Manuscript and Translation Differences"

# Headings in the variants files that introduce a secondary side-note,
# closing summary, or background-context paragraph rather than a genuine,
# distinct variant entry. Excluded from the book_stats() count below so
# that number reflects actual variant write-ups, not every "## " heading.
_VARIANTS_NON_ENTRY_HEADING = re.compile(r'^(side notes?:?|summary observations?|background)', re.I)

def book_stats():
    """Compute the three headline counts referenced in Purpose & Scope and
    How to Use This Book, live from the actual current content, so those
    numbers can never drift out of sync the way the hardcoded "25
    translations" / "866 claims" text has in the past. Returns a dict with
    keys: translation_count, variant_count, rcp_count."""
    translation_count = len(PROTESTANT_HISTORIES) + len(CATHOLIC_HISTORIES) + len(OTHER_HISTORIES)

    variant_count = 0
    variants_dir = os.path.join(REPO, VARIANTS_DIR)
    for fn in os.listdir(variants_dir):
        if not fn.endswith(".md"):
            continue
        with open(os.path.join(variants_dir, fn), encoding="utf-8") as f:
            content = f.read()
        for heading in re.findall(r'^## (.*)$', content, re.M):
            if not _VARIANTS_NON_ENTRY_HEADING.match(heading.strip()):
                variant_count += 1

    # Each book's by-claim file has the full, deduplicated set of numbered
    # entries -- one row per alleged contradiction, not one per translation.
    # Walked recursively (not a flat os.listdir()) as a defensive measure:
    # if a future change ever nests these files in subdirectories again, a
    # flat listdir would silently count zero files rather than raising an
    # error -- exactly the kind of quiet failure this project has been
    # bitten by before.
    rcp_dir = os.path.join(REPO, "110 Reportedly Contradicting Passages",
                           "015 Reportedly Contradicting Passages By Claim")
    rcp_count = 0
    rcp_files_found = 0
    for _root, _dirs, _files in os.walk(rcp_dir):
        for fn in _files:
            if not fn.endswith(".md"):
                continue
            rcp_files_found += 1
            with open(os.path.join(_root, fn), encoding="utf-8") as f:
                content = f.read()
            rcp_count += len(re.findall(r'^### \d+\.', content, re.M))

    # Fail loudly rather than silently produce a wrong (too-low) count if a
    # future path or naming change causes files to go unfound again.
    if rcp_files_found < 70:
        raise RuntimeError(
            f"book_stats() only found {rcp_files_found} RCP by-claim files "
            f"under {rcp_dir!r} -- expected around 73. A path or filename "
            f"mismatch is likely silently hiding books from the build."
        )

    return {
        "translation_count": translation_count,
        "variant_count": variant_count,
        "rcp_count": rcp_count,
    }

OT_GROUPS = [
    ("Pentateuch/Law/Torah", ["010 Genesis.md", "020 Exodus.md", "040 Leviticus.md", "060 Numbers.md", "080 Deuteronomy.md"]),
    ("Historical Books", ["100 Joshua.md", "120 Judges.md", "140 Ruth.md", "160 1 Samuel.md", "180 2 Samuel.md",
                           "200 1 Kings.md", "220 2 Kings.md", "240 1 Chronicles.md", "250 2 Chronicles.md",
                           "260 Ezra.md", "270 Nehemiah.md", "280 Esther.md"]),
    ("Wisdom/Poetic Books", ["290 Job.md", "300 Psalms.md", "310 Proverbs.md", "320 Ecclesiastes.md", "330 Song of Solomon.md"]),
    ("Major Prophets", ["340 Isaiah.md", "360 Jeremiah.md", "370 Lamentations.md", "380 Ezekiel.md", "390 Daniel.md"]),
    ("Minor Prophets", ["400 Hosea.md", "410 Joel.md", "420 Amos.md", "430 Obadiah.md", "440 Jonah.md", "450 Micah.md",
                         "460 Nahum.md", "480 Habakkuk.md", "490 Zephaniah.md", "500 Haggai.md", "510 Zechariah.md",
                         "520 Malachi.md"]),
]

APOCRYPHA = ["530 Tobit.md", "540 Judith.md", "550 Wisdom of Solomon.md", "560 Sirach.md", "570 Baruch.md",
             "580 1 Maccabees.md", "590 2 Maccabees.md"]

NT_GROUPS = [
    ("Gospels", ["600 Matthew.md", "610 Mark.md", "620 Luke.md", "630 John.md"]),
    ("History", ["640 Acts.md"]),
    ("Pauline Epistles", ["650 Romans.md", "660 1 Corinthians.md", "670 2 Corinthians.md", "680 Galatians.md",
                           "690 Ephesians.md", "700 Philippians.md", "710 Colossians.md", "720 1 Thessalonians.md",
                           "730 2 Thessalonians.md", "740 1 Timothy.md", "750 2 Timothy.md", "760 Titus.md",
                           "770 Philemon.md"]),
    ("General Epistles", ["780 Hebrews.md", "790 James.md", "800 1 Peter.md", "810 2 Peter.md", "820 1 John.md",
                           "830 2 John.md", "840 3 John.md", "850 Jude.md"]),
    ("Apocalyptic", ["860 Revelation.md"]),
]

# ---------------------------------------------------------------------------
# Reportedly Contradicting Passages (110). Each book has up to 10
# translation-specific files (by-book arrangement); each translation has all
# 73 book files (by-translation arrangement). Both arrangements are included
# in the compiled document as separate sections. Reuses the OT_GROUPS /
# APOCRYPHA / NT_GROUPS filename lists above (same book set/order as the
# 100 Variants section) to drive both loops.
# ---------------------------------------------------------------------------

RCP_DIR = "110 Reportedly Contradicting Passages"
RCP_BY_TRANS_DIR = os.path.join(RCP_DIR, "010 Arranged By Bible Translation-Version")
RCP_BY_BOOK_DIR = os.path.join(RCP_DIR, "020 Arranged By Books of the Bible")

# Populated after resolve_anchors() runs (see bottom of file). The by-book
# index's render functions are built before that point, but Python closures
# look up module-level names at *call* time, and render() is called again
# after this dict is filled in -- so as long as this is mutated in place
# (.update(), never reassigned), the render functions will see the real,
# resolved anchors on their second (real) invocation.
RCP_ANCHORS = {}


def extract_claims(path):
    """Return [(number, title), ...] straight from a RCP source file's own
    '### N. Title' headings (undemoted) -- used to build a book's claim
    list for its index page, independent of how that content is nested
    when embedded in the compiled document."""
    text = read(path)
    return re.findall(r'^### (\d+)\. (.+)$', text, flags=re.MULTILINE)

# (translation folder name, filename code suffix, display name)
RCP_TRANSLATIONS = [
    ("020 Amplified Bible", "AMP", "Amplified Bible (AMP)"),
    ("060 Christian Standard Bible", "CSB", "Christian Standard Bible (CSB)"),
    ("100 English Standard Version", "ESV", "English Standard Version (ESV)"),
    ("120 King James Version", "KJV", "King James Version (KJV)"),
    ("140 New King James Version", "NKJV", "New King James Version (NKJV)"),
    ("160 New American Standard Bible", "NASB", "New American Standard Bible (NASB)"),
    ("180 Legacy Standard Bible", "LSB", "Legacy Standard Bible (LSB)"),
    ("200 New International Version", "NIV", "New International Version (NIV)"),
    ("220 New Living Translation", "NLT", "New Living Translation (NLT)"),
    ("240 Revised Standard Version, Second Catholic Edition", "RSV2CE", "Revised Standard Version, Second Catholic Edition (RSV2CE)"),
]


def rcp_book_folder(fn):
    """'010 Genesis.md' -> '010 Genesis' (matches the by-book subfolder name)."""
    return fn[:-3].strip()


def rcp_book_display(fn):
    """'160 1 Samuel.md' -> '1 Samuel'."""
    name = re.sub(r'^\d+ ', '', fn)
    return re.sub(r'\.md$', '', name)


def rcp_book_token(fn):
    """'160 1 Samuel.md' -> '1Samuel' (matches the RCP filename token used in
    both arrangements; 'Song of Solomon' / 'Wisdom of Solomon' keep their
    internal 'Of' capitalized per the source files' own naming)."""
    token = rcp_book_display(fn).replace(' ', '')
    token = token.replace('Songof', 'SongOf').replace('Wisdomof', 'WisdomOf')
    return token


def read_rcp_body(path, demote_levels=3):
    """RCP files open with their own '# Reportedly Contradictory Passages in
    ... ' H1, which is redundant once wrapped in our own book/translation
    heading -- so this drops that line entirely and demotes the rest (H2/H3
    in the source) by `demote_levels`, keeping everything within Markdown's
    6-heading-level ceiling once nested under a book or translation divider."""
    text = read(path)
    lines = text.split("\n")
    idx = 0
    while idx < len(lines) and lines[idx].strip() == "":
        idx += 1
    if idx >= len(lines) or not lines[idx].startswith("# "):
        raise ValueError(f"Expected an H1 heading on line 1 of {path!r}")
    idx += 1
    while idx < len(lines) and lines[idx].strip() == "":
        idx += 1
    body = "\n".join(lines[idx:])
    return demote(body, demote_levels)


def rcp_all_books():
    """Flat ordered list of (group_label, fn) covering OT groups, Apocrypha,
    and NT groups, for looping book-by-book regardless of arrangement."""
    out = []
    out.append(("__OT__", None))
    for group_title, files in OT_GROUPS:
        for fn in files:
            out.append((group_title, fn))
    out.append(("__APOCRYPHA__", None))
    for fn in APOCRYPHA:
        out.append(("Apocrypha", fn))
    out.append(("__NT__", None))
    for group_title, files in NT_GROUPS:
        for fn in files:
            out.append((group_title, fn))
    return out


def get_heading_text(path):
    """Read the exact H1 heading text directly from a source file, rather
    than assuming it can be derived from the filename. Book heading text has
    changed format before (e.g. 'Textual Variants in the Book of Genesis' ->
    'Genesis: Significant Textual Variants Across Ten Translations') without
    the filename changing, which silently broke anything that assumed a
    fixed pattern -- so we always read it live from the file instead."""
    with open(os.path.join(REPO, path), encoding="utf-8") as f:
        first_line = f.readline().strip()
    m = re.match(r'^#+\s+(.*)$', first_line)
    if not m:
        raise ValueError(f"Expected an H1 heading on line 1 of {path!r}, got: {first_line!r}")
    return m.group(1).strip()


def book_title_text(fn):
    """Look up the book's actual heading text live from its source file."""
    return get_heading_text(os.path.join(VARIANTS_DIR, fn))


def build_targets():
    """Return an ordered list of dicts: id, search_text, render() -> markdown.
    Every target gets a page break before it and is a candidate for an extra
    blank page to push it onto an odd page number -- except targets marked
    is_divider=True (pure section/subsection title headings with no body
    content of their own, e.g. "## Protestant Bibles"), which instead have
    their heading folded onto the same page as the next non-divider target
    that follows, rather than being given a page of their own. See the
    final-assembly loop below for how this folding is applied."""
    targets = []

    def add(id_, search_text, render_fn, is_divider=False):
        targets.append({"id": id_, "search_text": search_text, "render": render_fn,
                         "is_divider": is_divider})

    add("preface", get_heading_text("030 Introduction/010 Preface.md"),
        lambda: read("030 Introduction/010 Preface.md"))

    def render_with_book_stats(path):
        text = read(path)
        stats = book_stats()
        text = text.replace("{{TRANSLATION_COUNT}}", str(stats["translation_count"]))
        text = text.replace("{{VARIANT_COUNT}}", str(stats["variant_count"]))
        text = text.replace("{{RCP_COUNT}}", str(stats["rcp_count"]))
        return text

    add("purpose_and_scope", get_heading_text("030 Introduction/020 Purpose and Scope.md"),
        lambda: render_with_book_stats("030 Introduction/020 Purpose and Scope.md"))
    add("how_to_use_this_book", get_heading_text("030 Introduction/030 How to Use This Book.md"),
        lambda: render_with_book_stats("030 Introduction/030 How to Use This Book.md"))
    add("reading_paths", get_heading_text("030 Introduction/040 Reading Paths for Different Readers.md"),
        lambda: read("030 Introduction/040 Reading Paths for Different Readers.md"))
    add("background_on_textual_transmission", get_heading_text("030 Introduction/050 Background on Textual Transmission.md"),
        lambda: read("030 Introduction/050 Background on Textual Transmission.md"))
    add("note_on_method_and_verification", get_heading_text("030 Introduction/060 A Note on Method and Verification.md"),
        lambda: read("030 Introduction/060 A Note on Method and Verification.md"))
    add("biblical_source_manuscripts", "Biblical Source Manuscripts",
        lambda: read("040 Biblical Source Manuscripts/010 Biblical Source Manuscripts.md"))
    add("character_of_each_tradition", get_heading_text("050 Character Of Each Source Manuscript Tradition/010 Character of Each Tradition.md"),
        lambda: read("050 Character Of Each Source Manuscript Tradition/010 Character of Each Tradition.md"))
    add("popular_bible_translations", "Popular Bible Translations",
        lambda: read("060 Popular Bible Translations/010 Popular Bible Translations.md"))
    add("bible_translations_and_sources", "Bible Translations and Their Source Manuscripts",
        lambda: read("070 Bible Translations and Their Source Manuscripts/010 Bible Translations and Their Source Manuscripts.md"))

    add("histories_title", "Histories of Various Bible Translations",
        lambda: "# Histories of Various Bible Translations\n", is_divider=True)

    def add_history_group(divider_id, divider_label, group):
        add(divider_id, divider_label, lambda: f"## {divider_label}\n", is_divider=True)
        for fn, title in group:
            def render(fn=fn):
                body = read(os.path.join("080 Histories of Various Bible Translations", fn))
                return demote(body, 2)
            heading_text = get_heading_text(os.path.join("080 Histories of Various Bible Translations", fn))
            add(f"history_{fn}", heading_text, render)

    add_history_group("histories_protestant_divider", "Protestant Bibles", PROTESTANT_HISTORIES)
    add_history_group("histories_catholic_divider", "Catholic Bibles", CATHOLIC_HISTORIES)
    add_history_group("histories_jewish_divider", "Jewish Bibles", OTHER_HISTORIES)

    add("variants_title", "Manuscript and Translation Differences",
        lambda: (
            "# Manuscript and Translation Differences\n\n"
            "**A note on recurring patterns:** the variants documented book by book below fall "
            "into a few recurring types, worth naming once rather than repeating at every "
            "entry. Some are secondary expansions or liturgical/theological additions to an "
            "earlier, shorter text (the longer ending of Mark, the Pericope Adulterae, the "
            "Comma Johanneum, and the Samaritan Pentateuch's Gerizim commandment in Deuteronomy "
            "are the clearest examples). Others reflect genuine textual plurality already "
            "present in the Second Temple period, before any Christian or sectarian editorial "
            "interest existed to explain it away -- Jeremiah's two literary editions, Samuel's "
            "unusually extensive Dead Sea Scrolls corrections, and Deuteronomy 32:8's \"sons of "
            "God\" are the clearest examples of this. And in a few places, the Dead Sea Scrolls "
            "or the Septuagint preserve an older reading that the later Masoretic or Byzantine "
            "tradition appears to have altered, whether for theological, scribal, or unknown "
            "reasons. None of this is unique to any one book; it recurs enough across the "
            "biblical text that it's worth having in view from the start.\n"
        ))

    add("old_testament_divider", "Old Testament", lambda: "## Old Testament\n", is_divider=True)
    for group_title, files in OT_GROUPS:
        for idx, fn in enumerate(files):
            prefix = f"### {group_title}\n\n" if idx == 0 else ""
            def render(fn=fn, prefix=prefix):
                body = read(os.path.join(VARIANTS_DIR, fn))
                return prefix + demote(body, 3)
            add(f"ot_{fn}", book_title_text(fn), render)

    add("apocrypha_divider", "Apocrypha", lambda: "## Apocrypha\n", is_divider=True)
    for fn in APOCRYPHA:
        def render(fn=fn):
            body = read(os.path.join(VARIANTS_DIR, fn))
            return demote(body, 3)
        add(f"apoc_{fn}", book_title_text(fn), render)

    add("new_testament_divider", "New Testament", lambda: "## New Testament\n", is_divider=True)
    for group_title, files in NT_GROUPS:
        for idx, fn in enumerate(files):
            prefix = f"### {group_title}\n\n" if idx == 0 else ""
            def render(fn=fn, prefix=prefix):
                body = read(os.path.join(VARIANTS_DIR, fn))
                return prefix + demote(body, 3)
            add(f"nt_{fn}", book_title_text(fn), render)

    # ---- Reportedly Contradicting Passages (arranged by claim, one entry
    # per alleged contradiction, with translation differences noted inline
    # only where a genuine difference exists -- see 015 Reportedly
    # Contradicting Passages By Claim/) ----
    add("rcp_title", "Reportedly Contradicting Passages", lambda: "# Reportedly Contradicting Passages\n", is_divider=True)

    testament_labels = {"__OT__": "Old Testament", "__APOCRYPHA__": "Apocrypha", "__NT__": "New Testament"}
    RCP_BY_CLAIM_DIR = "110 Reportedly Contradicting Passages/015 Reportedly Contradicting Passages By Claim"

    prev_group = None
    for group_label, fn in rcp_all_books():
        if fn is None:
            # A real heading (not bold text) so this appears as its own
            # entry in the PDF bookmarks/ToC navigation, matching how the
            # Manuscript and Translation Differences section's testament
            # dividers behave. Folded onto the same page as the next
            # non-divider target, like that section's dividers are.
            label = testament_labels[group_label]
            def render_testament_divider(label=label):
                return f"## {label}\n"
            add(f"rcp_{group_label}_divider", label, render_testament_divider, is_divider=True)
            continue
        book_name = rcp_book_display(fn)
        book_folder = rcp_book_folder(fn)
        rel_path = os.path.join(RCP_BY_CLAIM_DIR, f"{book_folder}.md")
        if not os.path.exists(os.path.join(REPO, rel_path)):
            continue

        header_md = ""
        if group_label != prev_group:
            # Also a real heading, one level deeper than the testament
            # divider above, matching the category-level "### {group}"
            # headings used the same way in Manuscript and Translation
            # Differences.
            header_md += f"### {group_label}\n\n"
        prev_group = group_label

        def render_book(rel_path=rel_path, book_name=book_name, header_md=header_md):
            return header_md + f"#### {book_name}\n" + read_rcp_body(rel_path, demote_levels=3)
        add(f"rcpbook_{fn}", book_name, render_book)

    # ---- Back matter ----
    add("references", get_heading_text("120 References for Further Reading/010 References for Further Reading.md"),
        lambda: read("120 References for Further Reading/010 References for Further Reading.md"))

    return targets


def normalize_quotes(s):
    """Pandoc's smart-typography converts straight quotes to curly ones in
    rendered HTML, so normalize both sides before comparing heading text --
    otherwise anything with an apostrophe in its heading (e.g. "Smith's
    Literal Translation") never matches and silently loses its ToC link."""
    return (s.replace("\u2019", "'").replace("\u2018", "'")
             .replace("\u201c", '"').replace("\u201d", '"'))


def resolve_anchors(targets):
    """Determine the actual pandoc-generated heading anchor id for each
    target by rendering all target content to HTML (cheap: no LibreOffice
    needed) and matching headings back to targets by their known heading
    text, in document order. This avoids hand-computing pandoc's slug
    algorithm (which would silently drift out of sync if heading text in
    the source repo changes -- exactly what happened when book headings
    were rewritten from 'Textual Variants in the Book of X' to
    'X: Significant Textual Variants...' and broke every hardcoded anchor
    in the Table of Contents)."""
    content_md = "\n\n".join(t["render"]() for t in targets)
    html = pypandoc_convert(content_md)

    heading_pattern = re.compile(r'<h[1-6][^>]*\bid="([^"]+)"[^>]*>(.*?)</h[1-6]>', re.DOTALL)
    found = []
    for m in heading_pattern.finditer(html):
        anchor, raw_text = m.group(1), m.group(2)
        text = re.sub(r'<[^>]+>', '', raw_text)  # strip inline tags e.g. <strong>
        text = html_entities.unescape(text)  # &amp; -> & etc., so e.g. "Purpose & Scope" matches
        text = text.replace("\n", " ").strip()
        found.append((anchor, normalize_quotes(text)))

    anchors = {}
    cursor = 0
    for t in targets:
        target_text = normalize_quotes(t["search_text"])
        match_idx = None
        for i in range(cursor, len(found)):
            if found[i][1] == target_text:
                match_idx = i
                break
        if match_idx is None:
            print(f"WARNING: could not resolve an anchor for target "
                  f"'{t['id']}' (heading text: {target_text!r}). "
                  f"Its Table of Contents entry will not be a working link.")
            continue
        anchors[t["id"]] = found[match_idx][0]
        cursor = match_idx
    return anchors


def pypandoc_convert(markdown_text):
    """Convert markdown to HTML via a pandoc subprocess (fast, no reference
    doc or LibreOffice needed -- this is purely to read back the heading ids
    pandoc assigns)."""
    import subprocess
    result = subprocess.run(
        ["pandoc", "-f", "markdown+raw_attribute", "-t", "html"],
        input=markdown_text, capture_output=True, text=True, check=True,
    )
    return result.stdout


def build_title_md():
    """Read the title live from the source file. The version is normally
    read from the same file too, but a CI build can override it via the
    BUILD_VERSION environment variable (set by the workflow's version-
    computation step) so the rendered title page always shows the
    auto-incrementing vYYYYMMDDx build version rather than whatever's
    manually typed in the repo's title file. Returns (markdown, version)
    so callers needing just the resolved version string (e.g. the HTML
    closing footer, which has no reference-doc template to bake a
    {{BUILD_VERSION}} placeholder into) don't have to re-derive it.

    Anything in the source file after the title and version lines (e.g. a
    personal note to readers) is preserved and rendered as ordinary body
    text beneath the title block, rather than silently dropped -- this
    file previously only ever contained exactly those two content lines,
    so nothing read past them; a later addition of extra text below the
    version line surfaced that gap."""
    with open(os.path.join(REPO, "010 Title/010 Title Page.md"), encoding="utf-8") as f:
        raw_lines = f.read().split("\n")
    lines = [l.strip() for l in raw_lines if l.strip()]
    # Expected: ["# Title Page", "**A Bible Study Primer**", "v20260823a"]
    title_line = next((l for l in lines if l.startswith("**") and l.endswith("**")), None)
    version_line = next((l for l in lines[1:] if l != title_line), None)
    if title_line is None or version_line is None:
        raise ValueError(f"Could not parse title/version out of 010 Title Page.md; got lines: {lines!r}")
    title_text = title_line.strip("*")
    version_line_resolved = os.environ.get("BUILD_VERSION", version_line)

    # Everything after the version line, if anything, is extra body content
    # (e.g. a note to readers) -- include it as plain paragraphs. Each
    # source line is de-indented so an indented line in the source file
    # (easy to introduce by accident when hand-editing) doesn't get
    # misread as a Markdown code block instead of an ordinary paragraph.
    version_idx = next(i for i, l in enumerate(lines) if l == version_line)
    extra_lines = lines[version_idx + 1:]
    extra_md = ""
    if extra_lines:
        extra_md = "\n\n" + "\n\n".join(l.strip() for l in extra_lines)

    md = f"""::: {{custom-style="HiddenHeading"}}
Title Page
:::

::: {{custom-style="Title"}}
{title_text}
:::

::: {{custom-style="Subtitle"}}
{version_line_resolved}
:::
{extra_md}
"""
    return md, version_line_resolved

title_md, BUILD_VERSION_RESOLVED = build_title_md()

def canonical_display_name(heading_text, filename):
    """Derive a short display name for a Table-of-Contents entry from the
    book's actual current heading text (e.g. 'Genesis: Significant Textual
    Variants Across Ten Translations' -> 'Genesis'), so the ToC label always
    matches what the heading currently says instead of a hand-curated label
    that can silently drift out of sync. Falls back to a name derived from
    the filename if the heading doesn't follow the 'Name: description'
    pattern."""
    if ":" in heading_text:
        return heading_text.split(":", 1)[0].strip()
    name = re.sub(r'^\d+ ', '', filename)
    name = re.sub(r'\.md$', '', name)
    return name


def build_toc_md(anchors):
    def link(target_id, label):
        anchor = anchors.get(target_id)
        if anchor is None:
            return f"{label} (link unavailable)"
        return f"[{label}](#{anchor})"

    lines = ["# Table of Contents", ""]
    lines.append(f"- {link('preface', 'Preface')}")
    lines.append(f"- {link('purpose_and_scope', 'Purpose & Scope')}")
    lines.append(f"- {link('how_to_use_this_book', 'How to Use This Book')}")
    lines.append(f"- {link('reading_paths', 'Reading Paths for Different Readers')}")
    lines.append(f"- {link('background_on_textual_transmission', 'Background on Textual Transmission')}")
    lines.append(f"- {link('note_on_method_and_verification', 'A Note on Method and Verification')}")
    lines.append(f"- {link('biblical_source_manuscripts', 'Biblical Source Manuscripts')}")
    lines.append(f"- {link('character_of_each_tradition', 'Character of Each Manuscript Tradition, Relationships, and Principles of Weighing')}")
    lines.append(f"- {link('popular_bible_translations', 'Popular Bible Translations')}")
    lines.append(f"- {link('bible_translations_and_sources', 'Bible Translations and Their Source Manuscripts')}")

    lines.append("- **Histories of Various Translations/Versions of the Bible:**")
    lines.append("  - Protestant Bibles")
    for fn, label in PROTESTANT_HISTORIES:
        lines.append(f"    - {link(f'history_{fn}', label)}")
    lines.append("  - Catholic Bibles")
    for fn, label in CATHOLIC_HISTORIES:
        lines.append(f"    - {link(f'history_{fn}', label)}")
    lines.append("  - Jewish Bibles")
    for fn, label in OTHER_HISTORIES:
        lines.append(f"    - {link(f'history_{fn}', label)}")

    lines.append("- **Manuscript and Translation Differences:**")
    lines.append("  - Old Testament")
    for group_title, files in OT_GROUPS:
        lines.append(f"    - {group_title}")
        for fn in files:
            heading_text = book_title_text(fn)
            name = canonical_display_name(heading_text, fn)
            lines.append(f"      - {link(f'ot_{fn}', name)}")
    lines.append("    - Apocrypha")
    for fn in APOCRYPHA:
        heading_text = book_title_text(fn)
        name = canonical_display_name(heading_text, fn)
        lines.append(f"      - {link(f'apoc_{fn}', name)}")
    lines.append("  - New Testament")
    for group_title, files in NT_GROUPS:
        lines.append(f"    - {group_title}")
        for fn in files:
            heading_text = book_title_text(fn)
            name = canonical_display_name(heading_text, fn)
            lines.append(f"      - {link(f'nt_{fn}', name)}")

    lines.append(f"- **{link('rcp_title', 'Reportedly Contradicting Passages')}:**")
    lines.append("    - Old Testament")
    for group_title, files in OT_GROUPS:
        lines.append(f"      - {group_title}")
        for fn in files:
            lines.append(f"        - {link(f'rcpbook_{fn}', rcp_book_display(fn))}")
    lines.append("      - Apocrypha")
    for fn in APOCRYPHA:
        lines.append(f"        - {link(f'rcpbook_{fn}', rcp_book_display(fn))}")
    lines.append("    - New Testament")
    for group_title, files in NT_GROUPS:
        lines.append(f"      - {group_title}")
        for fn in files:
            lines.append(f"        - {link(f'rcpbook_{fn}', rcp_book_display(fn))}")

    lines.append(f"- {link('references', 'References for Further Reading')}")

    return "\n".join(lines) + "\n"

# ---------------------------------------------------------------------------
# Assembly: title (page 1, no break) + TOC + all TARGETS in order. Each
# target gets a page break before it; targets listed in BLANKS_FILE (a JSON
# array of target ids, produced by detect_pages.py after a pass-1 render)
# additionally get one extra page break (= one inserted blank page) before
# them, so they land on an odd page number.
# ---------------------------------------------------------------------------
import json

BLANKS_FILE = sys.argv[3] if len(sys.argv) > 3 else None
needs_blank = set()
if BLANKS_FILE and os.path.exists(BLANKS_FILE):
    with open(BLANKS_FILE, encoding="utf-8") as f:
        needs_blank = set(json.load(f))

targets = build_targets()
anchors = resolve_anchors(targets)
RCP_ANCHORS.update(anchors)
TOC_MD = build_toc_md(anchors)

out = [title_md]

# Also write out the ordered id/search_text metadata for the page-detection
# pass, regardless of whether this run itself uses blanks yet.
#
# The ToC lists ~200 entries and can span several pages on its own. If the
# very next target's search started right where the ToC's *heading* was
# found (its first page), that search could accidentally match its own
# short title/label text still sitting further down within the ToC's own
# multi-page body -- e.g. "Preface" is both the literal ToC link label
# *and* the real section heading, and detect_pages.py's forward-only search
# has no way to tell "a mention inside the ToC" from "the actual heading"
# apart from position. This bit the six new front-matter sections
# specifically because they're the first real content right after the ToC
# -- anything later in the document is naturally safe, since by then the
# cursor has already moved well past every page the ToC itself occupies.
# Inserting an explicit marker right after the ToC's own content, and
# resuming all subsequent searches from there, closes that gap generally
# rather than special-casing any one section.
TOC_END_MARKER = "End of Table of Contents"
meta = [
    {"id": "toc", "search_text": "Table of Contents"},
    {"id": "toc_end_marker", "search_text": TOC_END_MARKER},
]
for t in targets:
    # Divider targets no longer start a page of their own (see the
    # folding logic below), so there's nothing useful for detect_pages.py
    # to locate or odd-page-enforce for them individually -- they always
    # land wherever the next non-divider target lands.
    if t["is_divider"]:
        continue
    meta.append({"id": t["id"], "search_text": t["search_text"]})
meta_path = f"build/targets_meta_{FORMAT}.json"
with open(meta_path, "w", encoding="utf-8") as f:
    json.dump(meta, f, indent=2)

# TOC (first real target, right after the title page)
if "toc" in needs_blank:
    out.append(PAGEBREAK)
out.append(PAGEBREAK)
out.append(TOC_MD)
if "toc_end_marker" in needs_blank:
    out.append(PAGEBREAK)
out.append(f"\n\n*{TOC_END_MARKER}*\n\n")

# Pure divider/title-only targets (e.g. "## Protestant Bibles", or a single
# translation's "### King James Version" divider within the by-translation
# RCP index) have no body content of their own. Giving each one a page
# break of its own, as every other target gets, would strand it alone on
# a near-blank page immediately before the real content it introduces --
# exactly the "wasted page" pattern this buffer avoids. Instead, their
# rendered heading(s) are accumulated here and folded onto the same page
# as the next non-divider target: only that following target's own id is
# checked against needs_blank and given the page break, with any pending
# divider heading(s) prepended directly above its content on that one page.
pending_divider_md = []

for t in targets:
    if t["is_divider"]:
        pending_divider_md.append(t["render"]())
        continue
    if t["id"] in needs_blank:
        out.append(PAGEBREAK)
    out.append(PAGEBREAK)
    if pending_divider_md:
        out.append("\n\n".join(pending_divider_md))
        pending_divider_md = []
    out.append(t["render"]())
    if FORMAT == "html":
        out.append('\n\n[↑ Back to Table of Contents](#table-of-contents){.back-link}\n\n')

# Safety net: flush any divider(s) left pending if the target list somehow
# ended on a divider (not expected -- the document always ends with the
# "references" target -- but content should never be silently dropped).
if pending_divider_md:
    out.append(PAGEBREAK)
    out.append("\n\n".join(pending_divider_md))

# HTML has no per-page footer (it's a single continuous document, not
# paginated), so the version + GitHub link that DOCX/ODT/PDF carry on every
# page's footer are instead appended once, here, as a small closing footer
# at the very end of the document.
if FORMAT == "html":
    out.append(
        '\n\n<hr class="closing-footer-rule"/>\n\n'
        f'<p class="closing-footer">{BUILD_VERSION_RESOLVED} &middot; '
        '<a href="https://github.com/jamesdlowery/a_Bible_study_primer">'
        'https://github.com/jamesdlowery/a_Bible_study_primer</a></p>\n\n'
    )

final = "\n".join(out)
outpath = f"build/full_document_{FORMAT}.md"
with open(outpath, "w", encoding="utf-8") as f:
    f.write(final)

print("Wrote", len(final), "characters to", outpath)
print("Metadata for", len(meta), "targets written to", meta_path)
if needs_blank:
    print("Applied", len(needs_blank), "blank-page insertions from", BLANKS_FILE)
