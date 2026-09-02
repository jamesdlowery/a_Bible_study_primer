#!/usr/bin/env python3
"""Rewrite the AUTO-GENERATED-DOWNLOAD-LINKS block in README.md to point at
the current build's versioned release assets. Safe to run repeatedly --
only replaces content between the two marker comments, leaving the rest of
the README untouched.

Usage: python3 update_readme_links.py <version> [readme_path]
"""
import re
import sys

START = "<!-- AUTO-GENERATED-DOWNLOAD-LINKS:START -->"
END = "<!-- AUTO-GENERATED-DOWNLOAD-LINKS:END -->"


def build_block(version):
    base = f"../../releases/download/{version}/A_Bible_Study_Primer_{version}"
    lines = [
        START,
        f"- [📄 Word (.docx)]({base}.docx)",
        f"- [📄 OpenDocument (.odt)]({base}.odt)",
        f"- [📄 PDF]({base}.pdf)",
        f"- [🌐 HTML]({base}.html)",
        END,
    ]
    return "\n".join(lines)


def main():
    version = sys.argv[1]
    readme_path = sys.argv[2] if len(sys.argv) > 2 else "README.md"

    with open(readme_path, encoding="utf-8") as f:
        content = f.read()

    pattern = re.compile(re.escape(START) + r".*?" + re.escape(END), re.DOTALL)
    if not pattern.search(content):
        print(f"WARNING: could not find {START} / {END} markers in "
              f"{readme_path}; leaving it unchanged. (Have you pasted in "
              f"README_snippet.md yet?)")
        return

    new_content = pattern.sub(build_block(version), content)
    if new_content == content:
        print("README already up to date for this version; no changes made.")
        return

    with open(readme_path, "w", encoding="utf-8") as f:
        f.write(new_content)
    print(f"README.md download links updated to {version}")


if __name__ == "__main__":
    main()
