#!/usr/bin/env python3
"""Generate the remaining-verifications workbook from its source data.

The spreadsheet itself is a build artifact, never hand-edited. Its source of
truth is two plain-text files kept in the repo so they can be diffed and
edited like everything else:

    build/verifications/remaining_verifications.csv        one row per item
    build/verifications/remaining_verifications_notes.md   one paragraph per
                                                           Summary note

Every build regenerates the workbook as
A_Bible_Study_Primer_Remaining_Verifications_<version>.xlsx (plus a stable
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
import sys
from datetime import datetime, timezone

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

DATA_CSV = os.path.join("build", "verifications", "remaining_verifications.csv")
NOTES_MD = os.path.join("build", "verifications", "remaining_verifications_notes.md")

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF")
WRAP = Alignment(wrap_text=True, vertical="top")

# Column widths for the Residual verifications sheet, by header name.
WIDTHS = {
    "ID": 5, "Category": 34, "Section": 8, "File": 30, "Entry / Verse": 22,
    "Translation(s) or Subject": 28, "What the text currently says": 60,
    "Status": 16, "How to close": 48, "Priority": 10, "Est. lookups": 10,
    "Change since 1 Oct inventory": 60, "Text check (auto)": 24,
    "Marker file": 30, "Marker phrase": 40,
}

CATEGORY_ORDER = ["A", "B", "C", "D", "E"]


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

    # Output columns: everything from the CSV except the two marker columns,
    # then the auto text-check result, then the marker columns at the far
    # right (kept visible so the check is auditable).
    marker_cols = ["Marker file", "Marker phrase"]
    main_cols = [c for c in fieldnames if c not in marker_cols]
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
                if c in ("ID", "Est. lookups") and v.strip().lstrip("-").isdigit():
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

    sm = wb.create_sheet("Summary")
    sm.append(["Category", "Items", "Est. lookups",
               "Open (needs a check; incl. print-only)",
               "Closed (incl. closed by proxy / via NRSV)",
               "Deliberate / honest label / convention"])
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
        ])
    first, last = 2, 1 + len(cats)
    tot = last + 1
    sm.append(["Total"] + [f"=SUM({c}{first}:{c}{last})" for c in "BCDEF"])
    for cell in sm[tot]:
        cell.font = Font(bold=True)

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
        f"A_Bible_Study_Primer_Remaining_Verifications_{version}.xlsx"
    build(version, repo_root, out_path)


if __name__ == "__main__":
    main()
