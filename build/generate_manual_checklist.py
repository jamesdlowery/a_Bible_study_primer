#!/usr/bin/env python3
"""Build the manual-verification checklist workbook for a release:
a_Bible_study_primer_manual_checklist_<version>.xlsx.

It lists the checks that cannot be done from online sources and so need
a reader with the book in hand:

  * "Study Bibles - notes to check": one row per study-Bible position cell
    whose entry cites a source and is not yet recorded as verified in
    build/verifications/position_verifications.csv, for the volumes whose
    notes are not freely available online (ONLINE_VOLUMES are left out).
    Each row names the verses to open, what the entry says the notes say,
    and leaves columns for the finding, the edition/page and the date.
  * "RSV2CE - verses to read": the open ledger rows that need the Ignatius
    Press RSV Second Catholic Edition itself.

Rows disappear from the checklist as the checks are recorded, so the
workbook is always the current to-do list. Usage:
    python3 build/generate_manual_checklist.py <version> <repo_root> [out_path]
"""
import csv
import os
import re
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import positions_matrix as pm  # noqa: E402
import generate_verifications_xlsx as gv  # noqa: E402

# Volumes whose notes can be read online (public domain or freely
# published), so their cells are checked from here rather than by hand.
ONLINE_VOLUMES = {
    "NET Bible, Full-Notes Edition",
    "Haydock's Catholic Bible and Commentary",
    "Scofield Reference Bible / Scofield Study Bible",
    "Jonathan Edwards Study Bible",
}

# Where to open a volume for a category when the entry itself names no verse.
DEFAULT_PLACES = {
    "View of Scripture": "Preface / introduction; 2 Timothy 3:16",
    "Salvation": "Romans 3:28; Ephesians 2:8-9; James 2:24",
    "Baptism": "Acts 2:38; Romans 6:3-4; 1 Peter 3:21",
    "The Lord's Supper": "Matthew 26:26-28; 1 Corinthians 11:23-29; John 6:53",
    "Church government": "1 Timothy 3:1-13; Titus 1:5-7; Acts 20:17, 28",
    "Eschatology": "Revelation 20:1-6; 1 Thessalonians 4:17; Daniel 9:24-27",
    "Spiritual gifts": "1 Corinthians 12-14 (esp. 13:8-10, 14:39); Acts 2:4",
    "Women's ordination": "1 Timothy 2:11-15; 1 Corinthians 14:34-35; Romans 16:1-7",
    "Abortion": "Psalm 139:13-16; Exodus 21:22-25; Jeremiah 1:5",
    "Homosexuality": "Romans 1:26-27; 1 Corinthians 6:9-10; Leviticus 18:22",
    "Alcohol": "Ephesians 5:18; John 2:1-11; Proverbs 20:1; 1 Timothy 5:23",
    "Divorce and remarriage": "Matthew 19:3-9; Mark 10:2-12; 1 Corinthians 7:10-16",
    "Contraception": "Genesis 1:28; Genesis 38:8-10; 1 Corinthians 7:3-5",
    "Gambling": "Proverbs 16:33; 1 Timothy 6:9-10; Acts 1:26",
    "Dancing": "Exodus 15:20; 2 Samuel 6:14; Ecclesiastes 3:4",
    "Premarital sex/cohabitation": "1 Corinthians 6:12-20; Hebrews 13:4; 1 Thessalonians 4:3-8",
    "War/pacifism": "Matthew 5:38-48; Romans 13:1-4; Luke 3:14",
    "Death penalty": "Genesis 9:6; Romans 13:4; John 8:1-11",
    "Scripture and Tradition": "2 Thessalonians 2:15; Mark 7:8-13; 2 Timothy 3:16-17",
    "The Trinity and the person of Christ": "John 1:1, 18; Philippians 2:6-11; Colossians 1:15; 1 John 5:7-8",
    "Predestination and free will": "Romans 9; Ephesians 1:4-11; 2 Peter 3:9",
    "Sanctification and holiness": "Romans 7:14-25; Philippians 3:12; 1 Thessalonians 5:23",
    "Hell, judgment and the fate of the lost": "Matthew 25:46; Revelation 20:10-15; Luke 16:19-31",
    "Mary, the saints and prayer for the dead": "Luke 1:28, 48; Matthew 12:46-50; 2 Maccabees 12:44-45; Revelation 5:8",
    "Worship: day, style and liturgy": "Colossians 2:16-17; Acts 20:7; Revelation 1:10; Exodus 20:8-11",
    "Church and state / political engagement": "Romans 13:1-7; Acts 5:29; Matthew 22:21",
    "Race, ethnicity and the church": "Acts 17:26; Galatians 3:28; Genesis 9:20-27; Acts 10",
    "Creation and the age of the earth": "Genesis 1-2 (notes and any introductory essay); Exodus 20:11",
    "Tithing and the prosperity gospel": "Malachi 3:8-12; 2 Corinthians 8-9; 1 Timothy 6:5-10",
    "Bible translations and KJV-onlyism": "Preface / introduction to the volume",
    "Marriage roles: complementarian and egalitarian": "Ephesians 5:21-33; 1 Peter 3:1-7; 1 Corinthians 11:3; Genesis 3:16",
}

_BOOKS = (r"(?:[1-3] )?(?:Genesis|Exodus|Leviticus|Numbers|Deuteronomy|Joshua|Judges|Ruth|Samuel|Kings|"
          r"Chronicles|Ezra|Nehemiah|Esther|Job|Psalms?|Proverbs|Ecclesiastes|Isaiah|Jeremiah|Lamentations|"
          r"Ezekiel|Daniel|Hosea|Joel|Amos|Jonah|Micah|Malachi|Maccabees|Sirach|Wisdom|Tobit|Matthew|Mark|"
          r"Luke|John|Acts|Romans|Corinthians|Galatians|Ephesians|Philippians|Colossians|Thessalonians|"
          r"Timothy|Titus|Philemon|Hebrews|James|Peter|Jude|Revelation)")
_REF = re.compile(_BOOKS + r" \d+(?::\d+)?(?:[–\-]\d+(?::\d+)?)?(?:,\s?\d+)*")

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
WRAP = Alignment(wrap_text=True, vertical="top")


def _group(cat):
    if cat in pm.DOCTRINAL:
        return "Doctrinal positions"
    if cat in pm.ETHICAL:
        return "Ethical and lifestyle positions"
    return "Other divisive positions"


def _style(ws, widths):
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = WRAP
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{chr(64 + len(widths))}{max(ws.max_row, 2)}"


def build(version, repo_root, out_path):
    data = pm.parse(repo_root)["Study Bibles"]
    checks = gv.read_position_checks(repo_root)

    wb = Workbook()
    ws = wb.active
    ws.title = "Study Bibles - notes to check"
    ws.append(["Volume", "Table", "Category", "Verses / place to open",
               "What the entry says the notes say", "Source label in the entry",
               "Your finding (confirmed / differs - quote the note)", "Edition and page checked", "Date"])
    n_rows, per_volume = 0, {}
    for vol, cats in data.items():
        if vol in ONLINE_VOLUMES:
            continue
        for cat in pm.CATEGORIES:
            if cat not in cats:
                continue
            position, source = cats[cat]
            if gv.source_class(position, source) != "cited":
                continue
            chk = checks.get(("Study Bibles", vol, cat))
            if chk and chk.get("Status", "").strip().lower().startswith("verified"):
                continue
            refs = sorted(set(_REF.findall(position)))
            where = "; ".join(refs) if refs else DEFAULT_PLACES[cat]
            ws.append([vol, _group(cat), cat, where, position, source, "", "", ""])
            n_rows += 1
            per_volume[vol] = per_volume.get(vol, 0) + 1
    _style(ws, [30, 22, 26, 34, 70, 34, 40, 24, 12])

    ws2 = wb.create_sheet("RSV2CE - verses to read")
    ws2.append(["Ledger ID", "Section", "Verse", "What the entry currently says", "How to close",
                "Your reading (quote the RSV2CE text and any footnote)",
                "Edition (Ignatius Press, year/printing)", "Date"])
    n_rsv = 0
    with open(os.path.join(repo_root, gv.DATA_CSV), encoding="utf-8", newline="") as fh:
        for r in csv.DictReader(fh):
            if r["Status"].startswith("Open") and r["Translation(s) or Subject"].strip().startswith("RSV2CE"):
                ws2.append([int(r["ID"]), r["Section"], r["Entry / Verse"], r["What the text currently says"],
                            r["How to close"], "", "", ""])
                n_rsv += 1
    _style(ws2, [10, 9, 26, 60, 50, 44, 24, 12])

    ws3 = wb.create_sheet("How to use")
    ws3.column_dimensions["A"].width = 120
    online = ", ".join(sorted(ONLINE_VOLUMES))
    per = "; ".join(f"{v} {n}" for v, n in sorted(per_volume.items(), key=lambda x: -x[1]))
    for line in [
        f"Manual verification checklist for a Bible study primer, generated with build {version} from the live "
        f"position tables, the ledger and build/verifications/position_verifications.csv. It is regenerated on "
        f"every build; rows drop off as checks are recorded, so this copy is the current to-do list.",
        "",
        f"Sheet 'Study Bibles - notes to check': {n_rows} rows across {len(per_volume)} volumes -- every study-Bible "
        f"position cell whose entry cites a source and is not yet recorded as verified, for the volumes whose notes "
        f"are not freely available online (left out, to be checked online: {online}).",
        "For each row: open the named volume at the verses in column D (where the entry itself names no verse, the "
        "column gives the standard passages for that question), read the note, and fill column G with 'Confirmed' "
        "or 'Differs:' followed by the note's actual wording; put the edition and page in column H.",
        "Returned rows become lines in build/verifications/position_verifications.csv (Section 120, Entity = volume, "
        "Category = category, Status Verified, Checked against = column H), which turns the cell green on the "
        "'Study Bibles x categories' sheet of the remaining-verifications workbook and advances the category G rows.",
        "",
        f"Sheet 'RSV2CE - verses to read': the {n_rsv} open ledger rows that need the Ignatius Press RSV Second "
        f"Catholic Edition itself (print or ebook). Quote the verse and any footnote in column F; those close the "
        f"rows directly.",
        "",
        "Filter column A on either sheet to work one volume at a time. Rows per volume: " + (per or "none"),
    ]:
        ws3.append([line])
    for row in ws3.iter_rows():
        for cell in row:
            cell.alignment = WRAP
    wb.move_sheet("How to use", offset=-2)
    wb.save(out_path)
    print(f"Wrote {out_path}: {n_rows} study-Bible rows across {len(per_volume)} volumes, {n_rsv} RSV2CE rows")


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    version = sys.argv[1]
    repo_root = sys.argv[2] if len(sys.argv) > 2 else "."
    out_path = sys.argv[3] if len(sys.argv) > 3 else f"a_Bible_study_primer_manual_checklist_{version}.xlsx"
    build(version, repo_root, out_path)


if __name__ == "__main__":
    main()
