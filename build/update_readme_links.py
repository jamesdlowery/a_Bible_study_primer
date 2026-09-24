#!/usr/bin/env python3
"""Rewrite the AUTO-GENERATED-DOWNLOAD-LINKS block in README.md to point at
the current build's versioned release assets, each annotated with its own
current file size in MB. Safe to run repeatedly -- only replaces content
between the two marker comments, leaving the rest of the README untouched.

Sizes are read directly from the actual versioned files on disk (e.g.
A_Bible_Study_Primer_v20260924z.docx in the current working directory),
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


def build_block(version, files_dir="."):
    base_name = f"A_Bible_Study_Primer_{version}"
    base_url = f"../../releases/download/{version}/{base_name}"
    entries = [
        ("📄 Word (.docx)", "docx"),
        ("📄 OpenDocument (.odt)", "odt"),
        ("📄 PDF", "pdf"),
        ("🌐 HTML", "html"),
    ]
    lines = [START]
    for label, ext in entries:
        size = format_size(os.path.join(files_dir, f"{base_name}.{ext}"))
        suffix = f" — {size}" if size else ""
        lines.append(f"- [{label}]({base_url}.{ext}){suffix}")
    lines.append(END)
    return "\n".join(lines)


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
    if new_content == content:
        print("README already up to date for this version; no changes made.")
        return

    with open(readme_path, "w", encoding="utf-8") as f:
        f.write(new_content)
    print(f"README.md download links updated to {version}")


if __name__ == "__main__":
    main()
