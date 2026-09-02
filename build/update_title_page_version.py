#!/usr/bin/env python3
"""Rewrite the version line in 010 Title/010 Title Page.md to the current
build's version, so the source file in the repo never sits stale between
releases. This is a belt-and-suspenders companion to assemble.py's own
BUILD_VERSION override (which already makes every compiled DOCX/ODT/PDF/HTML
title page show the correct version regardless of what's in this file) --
this script keeps the *source* markdown itself in sync too, exactly the way
update_readme_links.py already keeps README.md's download links in sync.

Usage: python3 update_title_page_version.py <version> [title_page_path]
"""
import re
import sys

VERSION_PATTERN = re.compile(r'^v\d{8}[a-z]$', re.MULTILINE)


def main():
    version = sys.argv[1]
    path = sys.argv[2] if len(sys.argv) > 2 else "010 Title/010 Title Page.md"

    with open(path, encoding="utf-8") as f:
        content = f.read()

    if not VERSION_PATTERN.search(content):
        print(f"WARNING: could not find a vYYYYMMDDx version line in "
              f"{path}; leaving it unchanged.")
        return

    new_content = VERSION_PATTERN.sub(version, content)
    if new_content == content:
        print("Title page already up to date for this version; no changes made.")
        return

    with open(path, "w", encoding="utf-8") as f:
        f.write(new_content)
    print(f"{path} version updated to {version}")


if __name__ == "__main__":
    main()
