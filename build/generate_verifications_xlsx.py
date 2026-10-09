#!/usr/bin/env python3
"""Generate the remaining-verifications workbook from its source data.

The spreadsheet itself is a build artifact, never hand-edited. Its source of
truth is two plain-text files kept in the repo so they can be diffed and
edited like everything else:

    build/verifications/remaining_verifications.csv        one row per item
    build/verifications/remaining_verifications_notes.md   one paragraph per
                                                           Summary note

Every build regenerates the workbook as
a_Bible_study_primer_remaining_verifications_<version>.xlsx (plus a stable
unversioned copy), attaches both to the release, and update_readme_links.py
links the versioned one from the README beneath the PDF.

What "update" means here, beyond re-stamping the version: each CSV row may
carry a "Marker file" and "Marker phrase" -- the exact wording in the book
that records the gap (e.g. "would be expected to ..."). On every build the
generator looks for that phrase in the live markdown and writes the result
to the "Text check (auto)" column: "phrase present" means the gap is still
in the text as described; "PHRASE NOT FOUND -- review" means someone has
edited that passage since the row was written, so the row's Status is
probably stale and should be revisited. Rows with no marker are left as
written. This never changes Status by itself -- the CSV stays the single
place a row is closed -- it only flags drift.

Usage: python3 build/generate_verifications_xlsx.py <version> [repo_root] [out_path]
"""
import csv
import os
import re
import sys
from datetime import datetime, timezone

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import positions_matrix as pm
import dividing_issues as di  # noqa: E402

DATA_CSV = os.path.join("build", "verifications", "remaining_verifications.csv")
NOTES_MD = os.path.join("build", "verifications", "remaining_verifications_notes.md")
POS_CSV = os.path.join("build", "verifications", "position_verifications.csv")

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF")
WRAP = Alignment(wrap_text=True, vertical="top")

# Column widths for the Residual verifications sheet, by header name.
WIDTHS = {
    "ID": 5, "Category": 34, "Section": 8, "File": 30, "Entry / Verse": 22,
    "Translation(s) or Subject": 28, "What the text currently says": 60,
    "Status": 16, "How to close": 48, "Priority": 10, "Est. lookups": 10,
    "Cells verified": 10, "Cells total": 10,
    "Change since 1 Oct inventory": 60, "Text check (auto)": 24,
    "Marker file": 30, "Marker phrase": 40,
}

CATEGORY_ORDER = ["A", "B", "C", "D", "E", "F", "G", "H"]

# Ledger rows for the 31 position categories are identified by their
# Section (110 = Denominations, 120 = Study Bibles) plus the category
# name in "Entry / Verse"; the build refills their status text and lookup
# count from the live tables so they can never go stale by hand.
SECTION_BY_CODE = {"110": "Denominations", "120": "Study Bibles"}
VERIFIED_FILL = PatternFill("solid", fgColor="C6EFCE")
NOSRC_FILL = PatternFill("solid", fgColor="FFEB9C")
NOPOS_FILL = PatternFill("solid", fgColor="EDEDED")


def read_position_checks(repo_root):
    """{(section, entity, category): row} from position_verifications.csv."""
    path = os.path.join(repo_root, POS_CSV)
    checks = {}
    if not os.path.exists(path):
        return checks
    with open(path, encoding="utf-8", newline="") as fh:
        for r in csv.DictReader(fh):
            key = (r["Section"].strip(), r["Entity"].strip(), pm._norm(r["Category"]))
            checks[key] = r
    return checks


def source_class(position, source):
    """'none' (nothing to verify), 'weak' (claim whose source the text itself
    marks as not identified / not independently documented), or 'cited'."""
    if not pm.has_position(position, source):
        return "none"
    src = source or ""
    # A label that opens with "Official" and names a document stays 'cited'
    # even if a clause of it is marked not independently documented; one that
    # admits no document was identified is 'weak' like any other such label.
    if re.match(r"^\s*official", src, re.I) and not re.search(r"no (document|text) identified|none identified", src, re.I):
        return "cited"
    # "not independently read/checked" names a place still to be read and stays
    # 'cited' (it feeds the manual checklist); "not independently documented"
    # admits there is no document, and is 'weak'.
    if re.search(r"not independently (documented|confirmed)|not identified|none identified", src, re.I):
        return "weak"
    return "cited"


def refresh_position_row(row, data, checks):
    """Rewrite the status text and lookup count of an F/G ledger row."""
    section = SECTION_BY_CODE.get(row.get("Section", "").strip())
    cat = pm._norm(row.get("Entry / Verse", ""))
    if not section or cat not in pm.CATEGORIES:
        return row
    entities = data[section]
    counts = {"none": 0, "weak": 0, "cited": 0}
    open_cells, verified = [], []
    for ent, cats in entities.items():
        if cat not in cats:
            counts["none"] += 1
            continue
        cls = source_class(*cats[cat])
        counts[cls] += 1
        if cls == "none":
            continue
        chk = checks.get((section, ent, cat))
        if chk and chk.get("Status", "").strip().lower().startswith("verified"):
            verified.append(ent)
        else:
            open_cells.append(ent)
    row = dict(row)
    row["What the text currently says"] = (
        f"{len(entities)} entries: {counts['cited']} "
        + ("state a position" if cat in pm.DOCTRINAL else "state a position with a cited source")
        + f", {counts['weak']} make a claim the text itself marks as not independently documented, "
        f"{counts['none']} take no position. "
        f"Verified cell-by-cell so far: {len(verified)} of {counts['cited'] + counts['weak']}."
        + (f" Still open: {', '.join(open_cells)}." if open_cells else " Nothing open.")
    )
    row["Est. lookups"] = str(len(open_cells))
    row["Cells verified"] = str(len(verified))
    row["Cells total"] = str(counts["cited"] + counts["weak"])
    if not open_cells and (counts["cited"] + counts["weak"]) > 0:
        row["Status"] = "Closed"
    elif verified:
        row["Status"] = f"Open – {len(verified)} of {len(verified) + len(open_cells)} verified"
    return row


def refresh_dividing_row(row, chapter_checks, checks):
    """Rewrite the status text and lookup count of a category H ledger row
    (Section 130, one per chapter of Major Dividing Issues Among Believers)
    from the chapter's own checkable parts and position_verifications.csv
    (Section 130, Entity = component name, Category = chapter)."""
    if row.get("Section", "").strip() != "130":
        return row
    cat = pm._norm(row.get("Entry / Verse", ""))
    comps = chapter_checks.get(cat)
    if not comps:
        return row
    bits, open_lookups, verified, manual_open = [], 0, [], []
    for name, kind in di.COMPONENTS:
        auto, n = comps[name]
        chk = checks.get(("130", name, cat))
        done = bool(chk and chk.get("Status", "").strip().lower().startswith("verified"))
        if kind == "auto":
            bits.append(f"passages: {'all catalog references resolve' if auto == 'ok' else auto}")
            open_lookups += n
        elif kind == "generated":
            bits.append("where-they-fall tables: generated at build, nothing to check")
        else:
            short = name.split(" (")[0].lower()
            if done:
                verified.append(short)
                bits.append(f"{short}: verified ({chk.get('Checked against', '').strip() or 'source not named'})")
            else:
                manual_open.append(short)
                open_lookups += n
                bits.append(f"{short}: {n} to check")
    row = dict(row)
    text = "; ".join(bits)
    row["What the text currently says"] = text[:1].upper() + text[1:] + "."
    row["Est. lookups"] = str(open_lookups)
    row["Cells verified"] = str(len(verified))
    row["Cells total"] = str(len(verified) + len(manual_open))
    if not manual_open and open_lookups == 0:
        row["Status"] = "Closed"
    elif verified:
        row["Status"] = f"Open – {len(verified)} of {len(verified) + len(manual_open)} verified"
    else:
        row["Status"] = "Open"
    return row


def write_divisive_sheet(wb, title, repo_root, chapter_checks, checks):
    """Chapter x checkable-component sheet for Major Dividing Issues Among
    Believers: plain = open manual check with its count; amber = a catalog
    cross-reference that does not resolve; grey = generated at build;
    green ✔ = recorded as verified in position_verifications.csv
    (Section 130, Entity = component, Category = chapter) or passed the
    automatic cross-reference check."""
    ws = wb.create_sheet(title)
    ws.append(["Chapter"] + [name for name, _ in di.COMPONENTS])
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for ch in di.chapters(repo_root):
        cat = ch["category"]
        row, fills = [f"{ch['number']}. {ch['title']}"], [None]
        for name, kind in di.COMPONENTS:
            auto, n = chapter_checks[cat][name]
            chk = checks.get(("130", name, cat))
            if chk and chk.get("Status", "").strip().lower().startswith("verified"):
                against = chk.get("Checked against", "").strip()
                row.append("✔ verified" + (f": {against}" if against else ""))
                fills.append(VERIFIED_FILL)
            elif kind == "auto":
                refs = len(ch["ids"]) + len(ch["variant_refs"])
                if auto == "ok":
                    row.append(f"✔ auto: all {refs} catalog references resolve")
                    fills.append(VERIFIED_FILL)
                else:
                    row.append(auto)
                    fills.append(NOSRC_FILL)
            elif kind == "generated":
                row.append("— generated at build from the position tables")
                fills.append(NOPOS_FILL)
            else:
                row.append(f"Open: {n} to check")
                fills.append(None)
        ws.append(row)
        r = ws.max_row
        for i, f in enumerate(fills, start=1):
            c = ws.cell(r, i)
            c.alignment = WRAP
            if f:
                c.fill = f
    ws.column_dimensions["A"].width = 44
    for i in range(2, len(di.COMPONENTS) + 2):
        ws.column_dimensions[get_column_letter(i)].width = 34
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(di.COMPONENTS) + 1)}{ws.max_row}"
    legend = ws.max_row + 2
    ws.cell(legend, 1, "Legend: plain = open manual check, with the number of separate attributions or dates to confirm; "
                       "amber = a catalog cross-reference (entry ID or variant §) that does not resolve; "
                       "grey = generated at build, nothing to verify; green ✔ = passed the automatic cross-reference "
                       "check, or recorded as verified in build/verifications/position_verifications.csv "
                       "(Section 130, Entity = column name, Category = chapter name).")
    ws.cell(legend, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=legend, start_column=1, end_row=legend, end_column=5)


def write_matrix(wb, title, section, entities, checks):
    """Entity x category sheet: each cell shows the source label the text
    gives for that position (or '—' where no position is taken), colored
    grey for no position, amber for a claim the text marks as not
    independently documented, green once position_verifications.csv
    records it as verified (the cell then also names what it was checked
    against)."""
    ws = wb.create_sheet(title)
    ws.append(["Entity"] + pm.CATEGORIES)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for ent, cats in entities.items():
        row = [ent]
        fills = [None]
        for cat in pm.CATEGORIES:
            if cat not in cats:
                row.append("—")
                fills.append(NOPOS_FILL)
                continue
            position, source = cats[cat]
            cls = source_class(position, source)
            if cls == "none":
                row.append("—")
                fills.append(NOPOS_FILL)
                continue
            label = (source or "stated").split(";")[0].strip()
            chk = checks.get((section, ent, cat))
            if chk and chk.get("Status", "").strip().lower().startswith("verified"):
                against = chk.get("Checked against", "").strip()
                row.append("✔ verified" + (f": {against}" if against else ""))
                fills.append(VERIFIED_FILL)
            else:
                row.append(label)
                fills.append(NOSRC_FILL if cls == "weak" else None)
        ws.append(row)
        r = ws.max_row
        for i, f in enumerate(fills, start=1):
            c = ws.cell(r, i)
            c.alignment = WRAP
            if f:
                c.fill = f
    ws.column_dimensions["A"].width = 34
    for i in range(2, len(pm.CATEGORIES) + 2):
        ws.column_dimensions[get_column_letter(i)].width = 22
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(pm.CATEGORIES) + 1)}{ws.max_row}"
    legend = ws.max_row + 2
    ws.cell(legend, 1, "Legend: grey '—' = no position taken (nothing to verify); plain = position with a cited source, "
                       "not yet checked; amber = claim the text itself marks as not independently documented; "
                       "green ✔ = recorded as verified in build/verifications/position_verifications.csv.")
    ws.cell(legend, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=legend, start_column=1, end_row=legend, end_column=8)


def read_rows(repo_root):
    with open(os.path.join(repo_root, DATA_CSV), encoding="utf-8", newline="") as fh:
        reader = csv.DictReader(fh)
        rows = list(reader)
        return reader.fieldnames, rows


def read_notes(repo_root):
    path = os.path.join(repo_root, NOTES_MD)
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8") as fh:
        text = fh.read()
    return [p.strip() for p in text.split("\n\n") if p.strip()]


def text_check(repo_root, row):
    """Confirm the row's marker phrase is still in the named file."""
    mf, mp = row.get("Marker file", "").strip(), row.get("Marker phrase", "").strip()
    if not mf or not mp:
        return ""
    path = os.path.join(repo_root, mf)
    if not os.path.exists(path):
        return "MARKER FILE NOT FOUND -- review"
    with open(path, encoding="utf-8") as fh:
        return "phrase present" if mp in fh.read() else "PHRASE NOT FOUND -- review"


def categories_in_order(rows):
    seen = {}
    for r in rows:
        seen.setdefault(r["Category"], None)
    return sorted(seen, key=lambda c: (CATEGORY_ORDER.index(c[0]) if c[:1] in CATEGORY_ORDER else 99, c))


def build(version, repo_root, out_path):
    fieldnames, rows = read_rows(repo_root)
    notes = read_notes(repo_root)
    data = pm.parse(repo_root)
    checks = read_position_checks(repo_root)
    rows = [refresh_position_row(r, data, checks) for r in rows]
    chapter_checks = di.chapter_checks(repo_root)
    rows = [refresh_dividing_row(r, chapter_checks, checks) for r in rows]

    # Output columns: everything from the CSV except the two marker columns,
    # then the auto text-check result, then the marker columns at the far
    # right (kept visible so the check is auditable).
    marker_cols = ["Marker file", "Marker phrase"]
    main_cols = [c for c in fieldnames if c not in marker_cols]
    # Two build-computed columns for the position categories (F, G) and the
    # Major Dividing Issues chapters (H): how many of the row's checkable
    # cells (or chapter components) are recorded as verified, out of how
    # many there are. Blank for the translation rows (A-E), whose unit of
    # work is the row itself.
    cell_cols = ["Cells verified", "Cells total"]
    if "Est. lookups" in main_cols:
        k = main_cols.index("Est. lookups") + 1
        main_cols = main_cols[:k] + cell_cols + main_cols[k:]
    else:
        main_cols = main_cols + cell_cols
    out_cols = main_cols + ["Text check (auto)"] + marker_cols

    wb = Workbook()
    ws = wb.active
    ws.title = "Residual verifications"
    ws.append(out_cols)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    for r in rows:
        values = []
        for c in out_cols:
            if c == "Text check (auto)":
                values.append(text_check(repo_root, r))
            else:
                v = r.get(c, "")
                if c in ("ID", "Est. lookups", "Cells verified", "Cells total") and v.strip().lstrip("-").isdigit():
                    v = int(v)
                values.append(v)
        ws.append(values)

    for cell_row in ws.iter_rows(min_row=2):
        for cell in cell_row:
            cell.alignment = WRAP
    for i, c in enumerate(out_cols, start=1):
        ws.column_dimensions[get_column_letter(i)].width = WIDTHS.get(c, 18)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(out_cols))}{ws.max_row}"

    # --- Summary sheet: live formulas over the data sheet ---
    n = ws.max_row
    col = {c: get_column_letter(i) for i, c in enumerate(out_cols, start=1)}
    cat_rng = f"'Residual verifications'!${col['Category']}$2:${col['Category']}${n}"
    st_rng = f"'Residual verifications'!${col['Status']}$2:${col['Status']}${n}"
    lk_rng = f"'Residual verifications'!${col['Est. lookups']}$2:${col['Est. lookups']}${n}"
    cv_rng = f"'Residual verifications'!${col['Cells verified']}$2:${col['Cells verified']}${n}"
    ct_rng = f"'Residual verifications'!${col['Cells total']}$2:${col['Cells total']}${n}"

    sm = wb.create_sheet("Summary")
    sm.append(["Category", "Items", "Est. lookups",
               "Open (needs a check; incl. print-only and partly verified)",
               "Closed (checked against the exact source)",
               "Unclassified (should be 0)",
               "Cells verified (F, G: position cells; H: chapter components)",
               "Cells total",
               "Cells verified %"])
    for cell in sm[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    cats = categories_in_order(rows)
    for i, cat in enumerate(cats, start=2):
        sm.append([
            cat,
            f"=COUNTIF({cat_rng},A{i})",
            f"=SUMIF({cat_rng},A{i},{lk_rng})",
            f'=COUNTIFS({cat_rng},A{i},{st_rng},"Open*")',
            f'=COUNTIFS({cat_rng},A{i},{st_rng},"Closed*")',
            f"=B{i}-D{i}-E{i}",
            f"=SUMIF({cat_rng},A{i},{cv_rng})",
            f"=SUMIF({cat_rng},A{i},{ct_rng})",
            f'=IF(H{i}=0,"",G{i}/H{i})',
        ])
    first, last = 2, 1 + len(cats)
    tot = last + 1
    sm.append(["Total"] + [f"=SUM({c}{first}:{c}{last})" for c in "BCDEFGH"]
              + [f'=IF(H{tot}=0,"",G{tot}/H{tot})'])
    for cell in sm[tot]:
        cell.font = Font(bold=True)
    for rr in range(first, tot + 1):
        sm.cell(rr, 9).number_format = "0%"

    r = tot + 2
    sm.cell(r, 1, "Generated").font = Font(bold=True)
    stamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    sm.cell(r + 1, 1,
            f"This workbook is regenerated by every build of the repository. "
            f"This copy was generated for build {version} on {stamp} from "
            f"{DATA_CSV} and {NOTES_MD}. To change a row, edit the CSV, not "
            f"the spreadsheet. The 'Text check (auto)' column reports whether "
            f"each row's marker phrase is still present in the live text.")
    r += 3
    sm.cell(r, 1, "Notes").font = Font(bold=True)
    for k, note in enumerate(notes, start=1):
        sm.cell(r + k, 1, note)
    for cell_row in sm.iter_rows(min_row=tot + 2, max_row=sm.max_row, max_col=1):
        for cell in cell_row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    sm.column_dimensions["A"].width = 58
    sm.column_dimensions["B"].width = 20
    for c in "CDEF":
        sm.column_dimensions[c].width = 22
    for c in "GHI":
        sm.column_dimensions[c].width = 16

    write_matrix(wb, "Denominations x categories", "Denominations", data["Denominations"], checks)
    write_matrix(wb, "Study Bibles x categories", "Study Bibles", data["Study Bibles"], checks)
    write_divisive_sheet(wb, "Divisive Issues x categories", repo_root, chapter_checks, checks)

    wb.move_sheet("Summary", offset=-1)  # Summary first, data second
    wb.save(out_path)
    flagged = [row[out_cols.index("ID")] for row in ws.iter_rows(min_row=2, values_only=True)
               if str(row[out_cols.index("Text check (auto)")]).startswith(("PHRASE", "MARKER"))]
    print(f"Wrote {out_path}: {n - 1} rows, {len(cats)} categories"
          + (f"; TEXT-CHECK FLAGS on rows {flagged}" if flagged else "; all marker phrases present"))


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    version = sys.argv[1]
    repo_root = sys.argv[2] if len(sys.argv) > 2 else "."
    out_path = sys.argv[3] if len(sys.argv) > 3 else \
        f"a_Bible_study_primer_remaining_verifications_{version}.xlsx"
    build(version, repo_root, out_path)


if __name__ == "__main__":
    main()
