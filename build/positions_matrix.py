#!/usr/bin/env python3
"""Extract the doctrinal / ethical-and-lifestyle position tables from the
Denominations and Study Bibles sections into an entity x category matrix.

Shared by generate_verifications_xlsx.py (which writes the matrices into
the workbook) and usable standalone for a quick console summary.
"""
import os
import re

CATEGORIES = [
    "View of Scripture", "Salvation", "Baptism", "The Lord's Supper",
    "Church government", "Eschatology", "Spiritual gifts", "Women's ordination",
    "Abortion", "Homosexuality", "Alcohol", "Divorce and remarriage",
    "Contraception", "Gambling", "Dancing", "Premarital sex/cohabitation",
    "War/pacifism", "Death penalty",
]
DOCTRINAL = CATEGORIES[:8]
ETHICAL = CATEGORIES[8:]

SECTIONS = {
    "Denominations": os.path.join("110 Top Christian Denominations", "010 Top Christian Denominations.md"),
    "Study Bibles": os.path.join("120 Top Study Bibles", "010 Top Study Bibles.md"),
}

# Position text beginning this way records that the entity takes no
# position, so there is nothing to verify against a primary source.
NO_POSITION = re.compile(
    r"^\s*(not addressed|no (single|doctrinal|official|confessional|formal|stated|denominational) "
    r"(position|statement|stance|doctrine)|no position|not applicable|not treated|varies)",
    re.I)


def _norm(s):
    return re.sub(r"[’‘]", "'", s).strip()


def parse(repo_root):
    """Return {section: {entity: {category: (position, source)}}} ordered as in the text."""
    out = {}
    for section, rel in SECTIONS.items():
        entities = {}
        with open(os.path.join(repo_root, rel), encoding="utf-8") as fh:
            lines = fh.read().split("\n")
        entity = None
        for line in lines:
            m = re.match(r"^## \d+\.\s+(.*)$", line)
            if m:
                entity = m.group(1).strip()
                entities[entity] = {}
                continue
            if entity and line.startswith("|") and not line.startswith("|---") \
                    and not line.startswith("| Category"):
                cells = [c.strip() for c in line.strip().strip("|").split("|")]
                if len(cells) < 2:
                    continue
                cat = _norm(cells[0].strip("*"))
                if cat in CATEGORIES and cat not in entities[entity]:
                    position = cells[1]
                    source = cells[2] if len(cells) > 2 else ""
                    entities[entity][cat] = (position, source)
        out[section] = entities
    return out


def has_position(position, source=""):
    """True when the cell makes a claim that can be checked against a primary
    source. Where the table has a Source column (ethical/lifestyle tables),
    that column decides: 'Not addressed' / 'Not applicable' means nothing to
    verify, anything else (an official statement, a general pattern of the
    notes) is a checkable claim. Doctrinal tables have no Source column, so
    the wording of the position itself decides."""
    if source:
        return not re.match(r"^\s*(not addressed|not applicable)", source, re.I)
    return not NO_POSITION.match(position or "")


def summary(repo_root):
    """{(section, category): (n_entities, n_with_position, [entities with position])}"""
    data = parse(repo_root)
    res = {}
    for section, entities in data.items():
        for cat in CATEGORIES:
            with_pos = [e for e, cats in entities.items() if cat in cats and has_position(*cats[cat])]
            res[(section, cat)] = (len(entities), len(with_pos), with_pos)
    return res


if __name__ == "__main__":
    import sys
    root = sys.argv[1] if len(sys.argv) > 1 else "."
    for (section, cat), (n, k, _) in summary(root).items():
        print(f"{section:13} {cat:28} {k:2}/{n} take a position")
