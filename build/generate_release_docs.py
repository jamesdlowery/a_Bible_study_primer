#!/usr/bin/env python3
"""Generate the two per-release PDFs:

    a_Bible_study_primer_feature_list_<version>.pdf
        What the book contains -- rendered from
        build/release_docs/feature_list.md with its {{PLACEHOLDERS}} filled
        from the live tree (tree_stats.py), so the counts can never go stale.

    a_Bible_study_primer_change_log_<version>.pdf
        What changed since the previous release. Two parts:
          1. Editor's notes: the "## Unreleased" section of
             build/release_docs/changelog.md (Added / Corrected / Expanded /
             Restructured / Also), written by hand as work is done.
          2. Measured changes: computed by comparing this tree with the
             previous release's tree -- section word counts, tracked
             translations, variants, entries, denominations, study Bibles,
             and the titles of entries added or removed.

With --promote, the "## Unreleased" heading in changelog.md is renamed to
this version (with today's date) and a fresh empty "## Unreleased" is
inserted above it, so the file stays the single running history. The
workflow commits that edit back alongside README and the title page.

Usage:
  python3 build/generate_release_docs.py <version> <repo_root> <out_dir>
          [--prev <previous_tree_dir> --prev-version <tag>] [--promote]
"""
import argparse
import datetime
import os
import re
import sys

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (ListFlowable, ListItem, Paragraph,
                                SimpleDocTemplate, Spacer, Table, TableStyle)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import tree_stats  # noqa: E402

FEATURE_MD = os.path.join("build", "release_docs", "feature_list.md")
CHANGELOG_MD = os.path.join("build", "release_docs", "changelog.md")
REPO_URL = "github.com/jamesdlowery/a_Bible_study_primer"

# ---------------------------------------------------------------- styles
_ss = getSampleStyleSheet()
ST = {
    "title": ParagraphStyle("t", parent=_ss["Title"], fontSize=17, leading=21, spaceAfter=2, alignment=0),
    "sub": ParagraphStyle("s", parent=_ss["Normal"], fontSize=9.5, textColor=colors.HexColor("#555555"), spaceAfter=8),
    "h1": ParagraphStyle("h1", parent=_ss["Heading2"], fontSize=11.5, leading=14, spaceBefore=8, spaceAfter=3, textColor=colors.HexColor("#1F3864")),
    "h2": ParagraphStyle("h2", parent=_ss["Heading3"], fontSize=10, leading=12, spaceBefore=5, spaceAfter=1),
    "b": ParagraphStyle("b", parent=_ss["Normal"], fontSize=9, leading=11.2),
    "p": ParagraphStyle("p", parent=_ss["Normal"], fontSize=9, leading=11.2, spaceAfter=4),
}


def inline(md):
    """Minimal inline markdown -> reportlab markup: **bold**, *italic*, `code`."""
    t = md.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
    t = re.sub(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", r"<i>\1</i>", t)
    t = re.sub(r"`(.+?)`", r"<font face='Courier'>\1</font>", t)
    return t


def md_to_flowables(md):
    """Render a markdown-lite document (#, ##, ###, paragraphs, '- ' bullets,
    pipe tables). The first '# ' heading becomes the title and the paragraph
    right after it the subtitle."""
    out, lines, i = [], md.split("\n"), 0
    title_done = False
    while i < len(lines):
        line = lines[i].rstrip()
        if not line.strip():
            i += 1
            continue
        if line.startswith("<!--"):
            while i < len(lines) and "-->" not in lines[i]:
                i += 1
            i += 1
            continue
        if line.startswith("# ") and not title_done:
            out.append(Paragraph(inline(line[2:]), ST["title"]))
            title_done = True
            i += 1
            if i < len(lines) and lines[i].strip() and not lines[i].startswith(("#", "- ", "|")):
                out.append(Paragraph(inline(lines[i].strip()), ST["sub"]))
                i += 1
            continue
        if line.startswith("### "):
            out.append(Paragraph(inline(line[4:]), ST["h2"]))
            i += 1
            continue
        if line.startswith("## ") or line.startswith("# "):
            out.append(Paragraph(inline(line.lstrip("# ")), ST["h1"]))
            i += 1
            continue
        if line.startswith("- "):
            items = []
            while i < len(lines) and lines[i].startswith("- "):
                items.append(ListItem(Paragraph(inline(lines[i][2:].strip()), ST["b"]), leftIndent=10))
                i += 1
            out.append(ListFlowable(items, bulletType="bullet", start="•", leftIndent=10, bulletFontSize=7, spaceAfter=2))
            continue
        if line.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(re.match(r"^:?-+:?$", c) for c in cells):
                    rows.append(cells)
                i += 1
            ncol = max(len(r) for r in rows)
            widths = [2.3 * inch] + [(6.9 * inch - 2.3 * inch) / (ncol - 1)] * (ncol - 1) if ncol > 1 else [6.9 * inch]
            data = [[Paragraph(inline(c), ST["b"]) for c in r + [""] * (ncol - len(r))] for r in rows]
            t = Table(data, colWidths=widths, repeatRows=1)
            t.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DCE3F0")),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#999999")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ]))
            out.append(t)
            out.append(Spacer(1, 4))
            continue
        # paragraph: gather until blank
        para = []
        while i < len(lines) and lines[i].strip() and not lines[i].startswith(("#", "- ", "|")):
            para.append(lines[i].strip())
            i += 1
        out.append(Paragraph(inline(" ".join(para)), ST["p"]))
    return out


def write_pdf(path, flowables, title):
    doc = SimpleDocTemplate(path, pagesize=letter, leftMargin=0.8 * inch, rightMargin=0.8 * inch,
                            topMargin=0.7 * inch, bottomMargin=0.7 * inch, title=title,
                            author="A Bible Study Primer")
    flowables = list(flowables) + [Spacer(1, 8), Paragraph(f"A Bible Study Primer -- {REPO_URL}", ST["sub"])]
    doc.build(flowables)


# ---------------------------------------------------------------- feature list
def fill_feature_list(md, s, version, date):
    intro = s["intro_chapters"]
    rep = {
        "{{VERSION}}": version, "{{DATE}}": date,
        "{{WORDS}}": f"{round(s['words'], -3):,}", "{{MD_FILES}}": str(s["md_files"]),
        "{{TRANSLATIONS}}": str(s["tracked_translations"]), "{{VARIANTS}}": str(s["variants"]), "{{SIDE_NOTES}}": str(s.get("side_notes", 0)),
        "{{RCP}}": str(s["rcp_entries"]), "{{RCP_BOOKS}}": str(s["rcp_books"]),
        "{{HISTORIES}}": str(s["histories"]), "{{DENOMINATIONS}}": str(s["denominations"]),
        "{{STUDY_BIBLES}}": str(s["study_bibles"]),
        "{{INTRO_COUNT}}": str(len(intro)), "{{INTRO_LIST}}": "; ".join(intro),
    }
    for k, v in rep.items():
        md = md.replace(k, v)
    left = re.findall(r"\{\{[A-Z_]+\}\}", md)
    if left:
        raise SystemExit(f"feature_list.md has unknown placeholders: {sorted(set(left))}")
    return md


# ---------------------------------------------------------------- change log
def split_changelog(md):
    """Return (preamble, {heading: body}) preserving order of headings."""
    parts = re.split(r"(?m)^## ", md)
    pre = parts[0]
    sections = []
    for p in parts[1:]:
        head, _, body = p.partition("\n")
        sections.append((head.strip(), body))
    return pre, sections


def editor_notes(repo_root, version):
    """Notes for this release: the version's own section if it already
    exists (a re-run after promotion), else whatever is under Unreleased."""
    with open(os.path.join(repo_root, CHANGELOG_MD), encoding="utf-8") as fh:
        md = fh.read()
    _, sections = split_changelog(md)
    for head, body in sections:
        if head.split()[0] == version and body.strip():
            return body.strip()
    for head, body in sections:
        if head.lower() == "unreleased":
            return body.strip()
    return ""


def promote(repo_root, version, date):
    path = os.path.join(repo_root, CHANGELOG_MD)
    with open(path, encoding="utf-8") as fh:
        md = fh.read()
    if re.search(rf"(?m)^## {re.escape(version)}\b", md):
        return False  # already promoted (re-run)
    _, sections = split_changelog(md)
    if not any(h.lower() == "unreleased" and b.strip() for h, b in sections):
        return False  # nothing under Unreleased: leave the file alone
    new = re.sub(r"(?m)^## Unreleased[ \t]*$", f"## Unreleased\n\n## {version} -- {date}", md, count=1)
    if new == md:
        return False
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(new)
    return True


def _diff_list(old, new, label, cap=40):
    added = [x for x in new if x not in old]
    removed = [x for x in old if x not in new]
    lines = []
    if added:
        lines.append(f"- **{label} added ({len(added)}):** " + "; ".join(added[:cap]) + (" ..." if len(added) > cap else ""))
    if removed:
        lines.append(f"- **{label} removed ({len(removed)}):** " + "; ".join(removed[:cap]) + (" ..." if len(removed) > cap else ""))
    return lines


def _norm_title(t):
    return re.sub(r"[^a-z0-9 ]", "", t.lower().replace("document", "chapter")).strip()


def _titles_diff(old, new, label, cap=30):
    """Compare {book: [titles]} dicts per book. A title that closely matches
    one in the other release (same text after normalisation, or difflib
    ratio >= 0.75) is treated as retitled rather than added + removed."""
    import difflib
    lines, detail = [], []
    a_total = r_total = t_total = 0
    for book in sorted(set(old) | set(new)):
        o, nn = old.get(book, []), new.get(book, [])
        on, nnn = [_norm_title(x) for x in o], [_norm_title(x) for x in nn]
        added, removed, retitled = [], [], 0
        used = set()
        for i, t_ in enumerate(nn):
            if nnn[i] in on:
                used.add(on.index(nnn[i]))
                continue
            best = max(((difflib.SequenceMatcher(None, nnn[i], x).ratio(), j) for j, x in enumerate(on) if j not in used), default=(0, -1))
            if best[0] >= 0.75:
                used.add(best[1])
                retitled += 1
            else:
                added.append(t_)
        removed = [o[j] for j in range(len(o)) if j not in used]
        a_total += len(added)
        r_total += len(removed)
        t_total += retitled
        if added or removed:
            detail.append((book, added, removed))
    summary = f"- **{label}:** {a_total} added, {r_total} removed" + (f", {t_total} retitled." if t_total else ".")
    lines.append(summary)
    for shown, (book, added, removed) in enumerate(detail):
        if shown >= cap:
            lines.append(f"- ... and changes in {len(detail) - shown} more books.")
            break
        bits = []
        if added:
            bits.append("added: " + "; ".join(x[:90] for x in added[:5]) + (" ..." if len(added) > 5 else ""))
        if removed:
            bits.append("removed: " + "; ".join(x[:90] for x in removed[:5]) + (" ..." if len(removed) > 5 else ""))
        lines.append(f"- *{book}* -- " + " | ".join(bits))
    return lines


def measured(cur, prev, prev_version):
    md = [f"## Measured changes since {prev_version}", ""]
    rows = [("Measure", prev_version, "This release"),
            ("Words (all sections)", f"{prev['words']:,}", f"{cur['words']:,}"),
            ("Source files", prev["md_files"], cur["md_files"]),
            ("Tracked translations", prev["tracked_translations"], cur["tracked_translations"]),
            ("Translation histories", prev["histories"], cur["histories"]),
            ("Catalogued variants (main entries)", prev["variants"], cur["variants"]),
            ("Side-note variants", prev.get("side_notes", 0), cur.get("side_notes", 0)),
            ("Alleged-contradiction entries", prev["rcp_entries"], cur["rcp_entries"]),
            ("Denominations profiled", prev["denominations"], cur["denominations"]),
            ("Study Bibles profiled", prev["study_bibles"], cur["study_bibles"]),
            ("'Not independently confirmed' notes", prev["unconfirmed"], cur["unconfirmed"])]
    md.append("| " + " | ".join(str(c) for c in rows[0]) + " |")
    md.append("|---|---|---|")
    for r in rows[1:]:
        md.append("| " + " | ".join(str(c) for c in r) + " |")
    md.append("")
    md.append("### Sections")
    ps = {n: (f, w) for n, f, w in prev["sections"]}
    cs = {n: (f, w) for n, f, w in cur["sections"]}
    for n in cs:
        if n not in ps:
            md.append(f"- **New section:** {n} ({cs[n][0]} files, {cs[n][1]:,} words).")
    for n in ps:
        if n not in cs:
            md.append(f"- **Section removed:** {n}.")
    grown = [(n, cs[n][1] - ps[n][1]) for n in cs if n in ps and abs(cs[n][1] - ps[n][1]) >= 200]
    for n, d in sorted(grown, key=lambda x: -abs(x[1])):
        md.append(f"- {n}: {'+' if d > 0 else ''}{d:,} words ({ps[n][1]:,} to {cs[n][1]:,}).")
    if not any(l.startswith("- ") for l in md[md.index("### Sections") + 1:]):
        md.append("- No section changed by more than 200 words.")
    md.append("")
    md.append("### Contents")
    md += _diff_list(prev["intro_chapters"], cur["intro_chapters"], "Introduction chapters")
    md += _diff_list(prev["history_titles"], cur["history_titles"], "Translation histories")
    md += _diff_list(prev["denomination_names"], cur["denomination_names"], "Denominations")
    md += _diff_list(prev["study_bible_names"], cur["study_bible_names"], "Study Bibles")
    md += _titles_diff(prev["variant_titles"], cur["variant_titles"], "Manuscript and Translation Differences")
    md += _titles_diff(prev["rcp_titles"], cur["rcp_titles"], "Reportedly Contradicting Passages")
    return "\n".join(md)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("version")
    ap.add_argument("repo_root")
    ap.add_argument("out_dir")
    ap.add_argument("--prev", help="directory holding the previous release's tree")
    ap.add_argument("--prev-version", default="the previous release")
    ap.add_argument("--promote", action="store_true")
    ap.add_argument("--date", default=datetime.date.today().strftime("%-d %B %Y"))
    a = ap.parse_args()

    cur = tree_stats.stats(a.repo_root)

    # 1. feature list
    with open(os.path.join(a.repo_root, FEATURE_MD), encoding="utf-8") as fh:
        fmd = fh.read()
    fmd = fill_feature_list(fmd, cur, a.version, a.date)
    fpath = os.path.join(a.out_dir, f"a_Bible_study_primer_feature_list_{a.version}.pdf")
    write_pdf(fpath, md_to_flowables(fmd), f"A Bible Study Primer -- Feature List ({a.version})")

    # 2. change log
    notes = editor_notes(a.repo_root, a.version)
    parts = [f"# A Bible Study Primer -- Change Log ({a.version})",
             f"Release of {a.date}. Changes since {a.prev_version}.", ""]
    parts.append("## Editor's notes")
    parts.append(notes if notes else "- No editor's notes were recorded for this release; see the measured changes below.")
    parts.append("")
    if a.prev and os.path.isdir(a.prev):
        prev = tree_stats.stats(a.prev)
        parts.append(measured(cur, prev, a.prev_version))
    else:
        parts.append("## Measured changes")
        parts.append("- No previous release was available for comparison.")
    cpath = os.path.join(a.out_dir, f"a_Bible_study_primer_change_log_{a.version}.pdf")
    write_pdf(cpath, md_to_flowables("\n".join(parts)), f"A Bible Study Primer -- Change Log ({a.version})")

    promoted = promote(a.repo_root, a.version, a.date) if a.promote else False
    print(f"Wrote {fpath} and {cpath}" + ("; changelog.md: Unreleased promoted to " + a.version if promoted else ""))


if __name__ == "__main__":
    main()
