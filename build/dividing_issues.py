#!/usr/bin/env python3
"""Fill the "where the bodies in this book fall" tables in the Major
Dividing Issues Among Believers section at build time.

The section's source file carries one marker per chapter:

    <!-- POSITIONS: <category> -->

where <category> is one of positions_matrix.CATEGORIES. The build replaces
each marker with two tables -- the 27 denominations and the 27 study
Bibles -- showing the opening sentence of that body's position cell and
the source label the text gives for it, read live from the position
tables in the Denominations and Study Bibles sections so the summaries
can never drift from the entries they summarise.

Usage (standalone, for a look at the output):
    python3 build/dividing_issues.py <repo_root> [category]
"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import positions_matrix as pm  # noqa: E402

MARKER = re.compile(r"^<!-- POSITIONS: (.+?) -->\s*$", re.M)
MAX_CHARS = 240

SECTION_LABEL = {
    "Denominations": ("Denominations", "Body"),
    "Study Bibles": ("Study Bibles", "Volume"),
}


def first_sentence(text, max_chars=MAX_CHARS):
    """The opening sentence of a position cell, trimmed to max_chars. A
    sentence ends at '. ' outside parentheses and not after a common
    abbreviation or a digit; a sentence longer than max_chars is cut at
    its last '; ' before the limit, else at a word boundary. Pipes are
    impossible (cells cannot contain them)."""
    t = text.strip()
    depth = 0
    end = None
    for i, ch in enumerate(t):
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth = max(0, depth - 1)
        elif ch == "." and depth == 0 and i + 1 < len(t) and t[i + 1] == " ":
            prev = t[max(0, i - 5):i + 1]
            if re.search(r"\b(e\.g|i\.e|vs|St|Mt|Dr|No|ch|cf|Art|Arts|rev|ed)\.$", prev):
                continue
            if re.search(r"\d\.$", prev):
                continue
            end = i + 1
            break
    out = t[:end] if end else t
    if len(out) > max_chars:
        head = out[:max_chars]
        k = head.rfind("; ")
        cut = head[:k] if k > max_chars // 2 else head.rsplit(" ", 1)[0]
        out = cut.rstrip(",;:") + " ..."
    # an unbalanced bold marker from a trim would spill over the cell
    if out.count("**") % 2:
        out += "**"
    return out


def source_label(source):
    s = (source or "").strip()
    return s if s else "—"


def table_for(section, entities, category):
    title, head = SECTION_LABEL[section]
    lines = [f"*{title}*", "", f"| {head} | Position (opening sentence) | Source |", "|---|---|---|"]
    for ent, cats in entities.items():
        if category not in cats:
            lines.append(f"| {ent} | — (no row for this category) | — |")
            continue
        position, source = cats[category]
        if not pm.has_position(position, source):
            pos = first_sentence(position) if position else "No position taken."
            lines.append(f"| {ent} | {pos} | {source_label(source)} |")
            continue
        lines.append(f"| {ent} | {first_sentence(position)} | {source_label(source)} |")
    return "\n".join(lines)


def render_category(data, category):
    if category not in pm.CATEGORIES:
        raise ValueError(f"Unknown position category in marker: {category!r}")
    parts = []
    for section in ("Denominations", "Study Bibles"):
        parts.append(table_for(section, data[section], category))
    note = ("*Opening sentence of each position cell; the full entries in \"Major U.S. Christian "
            "Denominations\" and \"Prominent English Study Bibles\" are the text to quote. "
            "A dash in the Source column means the row comes from a doctrinal table, which has no "
            "Source column of its own.*")
    return "\n\n".join(parts) + "\n\n" + note


def fill(md, repo_root):
    """Replace every POSITIONS marker in md with its generated tables."""
    data = pm.parse(repo_root)
    seen = []

    def repl(m):
        cat = pm._norm(m.group(1))
        seen.append(cat)
        return render_category(data, cat)

    out = MARKER.sub(repl, md)
    missing = [c for c in pm.CATEGORIES if c not in seen]
    if missing:
        raise ValueError("Major Dividing Issues section has no POSITIONS marker for: " + ", ".join(missing))
    extra = [c for c in seen if seen.count(c) > 1]
    if extra:
        raise ValueError("Duplicate POSITIONS marker(s): " + ", ".join(sorted(set(extra))))
    return out


if __name__ == "__main__":
    root = sys.argv[1] if len(sys.argv) > 1 else "."
    cat = sys.argv[2] if len(sys.argv) > 2 else pm.CATEGORIES[0]
    print(render_category(pm.parse(root), pm._norm(cat)))
