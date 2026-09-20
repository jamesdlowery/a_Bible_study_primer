#!/usr/bin/env python3
"""Post-process a pandoc-generated docx:
1. Fix tblLook so the header-row shading defined in the Table style's
   firstRow conditional formatting actually renders (pandoc emits
   firstRow="0" by default, which suppresses it).
2. Replace the {{BUILD_VERSION}} placeholder baked into the reference
   template's footers with the actual build version, since the footers
   live in the static reference doc and never pass through the markdown
   pipeline that already handles this substitution for the title page.
3. Mark every table row "cannot split across pages" (w:cantSplit).
   Word's default lets a tall row break mid-row at a page boundary,
   which is exactly what produced entries like a table row's label on
   one page and the rest of its content stranded alone at the top of
   the next -- confusing to read and easy to mistake for a missing or
   duplicated row. Pandoc doesn't expose a markdown-level way to set
   this, so it's applied uniformly here to every row in every table
   rather than requiring it be set table-by-table in source.

Usage: python3 postprocess_docx.py <docx_path> [build_version]
If build_version is omitted, falls back to the BUILD_VERSION environment
variable; if neither is available, the placeholder is left as-is (with a
warning) rather than silently shipping a blank.
"""
import sys
import os
import shutil
import zipfile
import re

def fix_docx(path, build_version=None):
    tmp = path + ".tmp.zip"
    with zipfile.ZipFile(path, "r") as zin:
        names = zin.namelist()
        doc_xml = zin.read("word/document.xml").decode("utf-8")
        footer_names = [n for n in names if re.fullmatch(r"word/footer\d+\.xml", n)]
        footers = {n: zin.read(n).decode("utf-8") for n in footer_names}

    old = '<w:tblLook w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0" w:val="0000" />'
    new = '<w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1" w:val="04A0" />'
    count = doc_xml.count(old)
    doc_xml = doc_xml.replace(old, new)

    # Prevent any table row from splitting across a page break. A row
    # already carrying <w:trPr>...</w:trPr> (e.g. the header row's
    # tblHeader marker) gets <w:cantSplit/> inserted inside it; a bare
    # <w:tr> with no trPr at all gets one added.
    def add_cant_split_to_trpr(m):
        return m.group(0)[:-len("</w:trPr>")] + "<w:cantSplit/></w:trPr>"
    doc_xml, trpr_count = re.subn(r"<w:trPr>.*?</w:trPr>", add_cant_split_to_trpr, doc_xml, flags=re.DOTALL)
    doc_xml, bare_count = re.subn(r"<w:tr>(?!<w:trPr>)", '<w:tr><w:trPr><w:cantSplit/></w:trPr>', doc_xml)
    row_count = trpr_count + bare_count

    version = build_version or os.environ.get("BUILD_VERSION")
    version_count = 0
    if version:
        for n in footer_names:
            version_count += footers[n].count("{{BUILD_VERSION}}")
            footers[n] = footers[n].replace("{{BUILD_VERSION}}", version)
    elif any("{{BUILD_VERSION}}" in footers[n] for n in footer_names):
        print("WARNING: no build_version provided and {{BUILD_VERSION}} "
              "placeholder found in footer(s); leaving placeholder as-is.")

    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "word/document.xml":
                zout.writestr(item, doc_xml)
            elif item.filename in footers:
                zout.writestr(item, footers[item.filename])
            else:
                zout.writestr(item, zin.read(item.filename))
    shutil.move(tmp, path)
    print(f"Fixed {count} table(s) in {path}")
    print(f"Marked {row_count} table row(s) as cantSplit in {path}")
    if version:
        print(f"Replaced {version_count} {{{{BUILD_VERSION}}}} placeholder(s) with {version}")

if __name__ == "__main__":
    docx_path = sys.argv[1]
    version_arg = sys.argv[2] if len(sys.argv) > 2 else None
    fix_docx(docx_path, version_arg)
