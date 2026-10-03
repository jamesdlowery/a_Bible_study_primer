#!/usr/bin/env python3
"""Headline statistics for any checkout of this book -- the current tree or
an earlier release's tree -- so the release documents can compare two.

Folders are matched by name with their numeric prefix stripped (so the
Manuscript and Translation Differences section is found whether it is
numbered 090 or 100), and files likewise, which keeps old releases
comparable after renumbering.

Returned dict keys (all ints unless noted):
  words, md_files, sections (list of (name, files, words)),
  tracked_translations (from Popular Bible Translations' numbered list),
  variants (side-note entries in Manuscript and Translation Differences),
  side_notes, rcp_entries, rcp_books, histories, denominations, study_bibles,
  intro_chapters (list of names), history_titles (list), denomination_names
  (list), study_bible_names (list), rcp_titles ({book: [titles]}),
  unconfirmed (count of "not independently confirmed" notes),
  variant_titles ({book: [titles]})
"""
import os
import re

_NUM = re.compile(r"^\d+ ")


def strip_num(name):
    return _NUM.sub("", name)


def _section_dir(tree, name):
    for d in os.listdir(tree):
        p = os.path.join(tree, d)
        if os.path.isdir(p) and strip_num(d) == name:
            return p
    return None


def _md_files(d):
    out = []
    for root, _, files in os.walk(d):
        for f in sorted(files):
            if f.endswith(".md") and f != "README.md":
                out.append(os.path.join(root, f))
    return out


def _read(p):
    with open(p, encoding="utf-8", errors="ignore") as fh:
        return fh.read()


def _wc(text):
    return len(text.split())


def stats(tree):
    s = {"sections": [], "words": 0, "md_files": 0}
    for d in sorted(os.listdir(tree)):
        p = os.path.join(tree, d)
        if not os.path.isdir(p) or not _NUM.match(d):
            continue
        files = _md_files(p)
        w = sum(_wc(_read(f)) for f in files)
        s["sections"].append((strip_num(d), len(files), w))
        s["words"] += w
        s["md_files"] += len(files)

    def names_from(section, pattern, flags=re.M):
        d = _section_dir(tree, section)
        if not d:
            return []
        out = []
        for f in _md_files(d):
            out += re.findall(pattern, _read(f), flags)
        return out

    intro = _section_dir(tree, "Introduction")
    s["intro_chapters"] = [strip_num(os.path.basename(f))[:-3] for f in _md_files(intro)] if intro else []
    hist = _section_dir(tree, "Histories of Various Bible Translations")
    s["history_titles"] = [strip_num(os.path.basename(f))[:-3] for f in _md_files(hist)] if hist else []
    s["histories"] = len(s["history_titles"])

    s["tracked_translations"] = len(names_from("Popular Bible Translations", r"^\d+\. \*\*"))

    var = _section_dir(tree, "Manuscript and Translation Differences")
    s["variant_titles"] = {}
    s["variants"] = 0
    if var:
        for f in _md_files(var):
            book = strip_num(os.path.basename(f))[:-3]
            # Main entries are "## <ref> -- <question>" headings; the
            # "Side notes" / "Summary" / "Background" headings are not
            # entries. Side-note variants (bold-led bullets under Side
            # notes) are counted separately as side_notes.
            text = _read(f)
            titles = [h.strip() for h in re.findall(r"^## (.*)$", text, re.M)
                      if not re.match(r"^(side notes?:?|summary observations?|background)", h.strip(), re.I)]
            s.setdefault("side_notes", 0)
            s["side_notes"] += len(re.findall(r"^- \*\*[1-3]?\s?[A-Z][a-z]+[^*]*\d+[^*]*\*\*", text, re.M))
            s["variant_titles"][book] = titles
            s["variants"] += len(titles)

    rcp = _section_dir(tree, "Reportedly Contradicting Passages")
    s["rcp_titles"] = {}
    s["rcp_entries"] = 0
    if rcp:
        for f in _md_files(rcp):
            if "By Claim" not in f:
                continue
            book = strip_num(os.path.basename(f))[:-3]
            titles = [re.sub(r"^\d+\.\s*", "", t) for t in re.findall(r"^### (.*)$", _read(f), re.M)]
            s["rcp_titles"][book] = titles
            s["rcp_entries"] += len(titles)
    s["rcp_books"] = len(s["rcp_titles"])

    s["denomination_names"] = [t.strip() for t in names_from("Top Christian Denominations", r"^## \d+\.\s+(.*)$")]
    s["denominations"] = len(s["denomination_names"])
    s["study_bible_names"] = [t.strip() for t in names_from("Top Study Bibles", r"^## \d+\.\s+(.*)$")]
    s["study_bibles"] = len(s["study_bible_names"])

    s["unconfirmed"] = 0
    for sec in ("Manuscript and Translation Differences", "Reportedly Contradicting Passages"):
        d = _section_dir(tree, sec)
        if d:
            for f in _md_files(d):
                s["unconfirmed"] += len(re.findall(r"not independently confirmed|unable to independently confirm", _read(f), re.I))
    return s


if __name__ == "__main__":
    import sys
    import json
    r = stats(sys.argv[1] if len(sys.argv) > 1 else ".")
    print(json.dumps({k: v for k, v in r.items() if not isinstance(v, (dict, list))}, indent=2))
    for name, n, w in r["sections"]:
        print(f"{name:55} {n:3} files {w:7} words")
