#!/usr/bin/env python3
"""Rewrite the AUTO-GENERATED-DOWNLOAD-LINKS block in README.md to point at
the current build's versioned release assets, each annotated with its own
current file size in MB. Safe to run repeatedly -- only replaces content
between the two marker comments, leaving the rest of the README untouched.

Sizes are read directly from the actual versioned files on disk (e.g.
a_Bible_study_primer_v20260924z.docx or
a_Bible_study_primer_remaining_verifications_v20260924z.xlsx in the
current working directory),
not hand-maintained, so they can never silently drift out of sync with
the real file -- recomputed fresh every time this script runs. Depends on
running after those files already exist in the working directory, which
the workflow's own step order already guarantees (this step runs after
"Prepare versioned copies for release", not before).

Usage: python3 update_readme_links.py <version> [readme_path] [files_dir]
"""
import os
import re
import sys

START = "<!-- AUTO-GENERATED-DOWNLOAD-LINKS:START -->"
END = "<!-- AUTO-GENERATED-DOWNLOAD-LINKS:END -->"

# A second marked block, the paragraph beneath the download list that
# describes the remaining-verifications workbook. The workbook's download
# link lives inside that paragraph rather than in the list above, so this
# block is rewritten with the versioned filename on every build too.
VSTART = "<!-- AUTO-GENERATED-VERIFICATIONS-LINK:START -->"
VEND = "<!-- AUTO-GENERATED-VERIFICATIONS-LINK:END -->"


def format_size(path):
    """'2.4 MB', or None if the file isn't there to measure -- e.g. a
    manual/local run of this script done without having actually built
    every format first. Callers fall back to no size annotation rather
    than a fake or stale one in that case."""
    try:
        size_bytes = os.path.getsize(path)
    except OSError:
        return None
    return f"{size_bytes / (1024 * 1024):.1f} MB"


def format_size_kb(path):
    """Same as format_size, but in whole KB -- for the remaining-verifications
    workbook, which is a few tens of KB and would otherwise show as '0.0 MB'."""
    try:
        size_bytes = os.path.getsize(path)
    except OSError:
        return None
    return f"{max(1, round(size_bytes / 1024))} KB"


def build_block(version, files_dir="."):
    book = f"a_Bible_study_primer_{version}"
    # (label, versioned filename) -- in README display order. The
    # remaining-verifications workbook is linked from its own paragraph
    # below the list (see build_verifications_block), not from here.
    entries = [
        ("📄 Word (.docx)", f"{book}.docx"),
        ("📄 OpenDocument (.odt)", f"{book}.odt"),
        ("📄 PDF", f"{book}.pdf"),
        ("🌐 HTML", f"{book}.html"),
    ]
    lines = [START]
    for label, filename in entries:
        size = format_size(os.path.join(files_dir, filename))
        suffix = f" — {size}" if size else ""
        lines.append(f"- [{label}](../../releases/download/{version}/{filename}){suffix}")
    lines.append(END)
    return "\n".join(lines)


def build_verifications_block(version, files_dir="."):
    filename = f"a_Bible_study_primer_remaining_verifications_{version}.xlsx"
    size = format_size_kb(os.path.join(files_dir, filename))
    suffix = f" ({size})" if size else ""
    link = f"[📊 **Remaining verifications (.xlsx)**](../../releases/download/{version}/{filename}){suffix}"
    paragraph = (
        f"{link} is the book's open-items ledger: every place where a claim "
        f"about a translation's wording still rests on an expectation or a "
        f"proxy rather than a direct check of that translation's own text, "
        f"with what was checked, what remains, and how to close it. It is "
        f"regenerated on every build from "
        f"[`build/verifications/remaining_verifications.csv`]"
        f"(build/verifications/remaining_verifications.csv) (edit that file, "
        f"not the spreadsheet), carries the same version stamp as the book it "
        f"was built with, and re-checks each row's marker phrase against the "
        f"current text so a row whose passage has since been edited is "
        f"flagged for review."
    )
    return "\n".join([VSTART, paragraph, VEND])


def main():
    version = sys.argv[1]
    readme_path = sys.argv[2] if len(sys.argv) > 2 else "README.md"
    files_dir = sys.argv[3] if len(sys.argv) > 3 else "."

    with open(readme_path, encoding="utf-8") as f:
        content = f.read()

    pattern = re.compile(re.escape(START) + r".*?" + re.escape(END), re.DOTALL)
    if not pattern.search(content):
        print(f"WARNING: could not find {START} / {END} markers in "
              f"{readme_path}; leaving it unchanged. (Have you pasted in "
              f"README_snippet.md yet?)")
        return

    new_content = pattern.sub(build_block(version, files_dir), content)

    vpattern = re.compile(re.escape(VSTART) + r".*?" + re.escape(VEND), re.DOTALL)
    if vpattern.search(new_content):
        new_content = vpattern.sub(build_verifications_block(version, files_dir), new_content)
    else:
        print(f"WARNING: could not find {VSTART} / {VEND} markers in "
              f"{readme_path}; the remaining-verifications paragraph was "
              f"left unchanged.")
    if new_content == content:
        print("README already up to date for this version; no changes made.")
        return

    with open(readme_path, "w", encoding="utf-8") as f:
        f.write(new_content)
    print(f"README.md download links updated to {version}")


if __name__ == "__main__":
    main()
