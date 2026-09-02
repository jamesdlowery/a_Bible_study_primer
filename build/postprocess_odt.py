#!/usr/bin/env python3
"""Post-process a pandoc-generated odt:
1. Add header-row shading and thin borders to tables. Pandoc's ODT writer
   generates a fixed-name 'TableHeaderRowCell' / 'TableRowCell' table-cell
   style once per document in content.xml's automatic-styles (not
   controllable via reference-doc), so we patch it directly here.
2. Replace the {{BUILD_VERSION}} placeholder baked into the reference
   template's footer/footer-left (in styles.xml) with the actual build
   version, since that content lives in the static reference doc and never
   passes through the markdown pipeline that already handles this
   substitution for the title page.

Usage: python3 postprocess_odt.py <odt_path> [build_version]
If build_version is omitted, falls back to the BUILD_VERSION environment
variable; if neither is available, the placeholder is left as-is (with a
warning) rather than silently shipping a blank.
"""
import sys
import os
import shutil
import zipfile

def fix_odt(path, build_version=None):
    tmp = path + ".tmp.zip"
    with zipfile.ZipFile(path, "r") as zin:
        content_xml = zin.read("content.xml").decode("utf-8")
        names = zin.namelist()
        styles_xml = zin.read("styles.xml").decode("utf-8") if "styles.xml" in names else None

    old_header = '<style:style style:name="TableHeaderRowCell" style:family="table-cell">\n      <style:table-cell-properties fo:border="none" />\n    </style:style>'
    new_header = ('<style:style style:name="TableHeaderRowCell" style:family="table-cell">\n'
                  '      <style:table-cell-properties fo:background-color="#1F4E79" '
                  'fo:border="0.5pt solid #BFBFBF" fo:padding="0.04in" style:vertical-align="middle" />\n'
                  '    </style:style>')
    old_row = '<style:style style:name="TableRowCell" style:family="table-cell">\n      <style:table-cell-properties fo:border="none" />\n    </style:style>'
    new_row = ('<style:style style:name="TableRowCell" style:family="table-cell">\n'
               '      <style:table-cell-properties fo:border="0.5pt solid #BFBFBF" fo:padding="0.04in" />\n'
               '    </style:style>')

    n1 = content_xml.count(old_header)
    n2 = content_xml.count(old_row)
    content_xml = content_xml.replace(old_header, new_header).replace(old_row, new_row)

    version = build_version or os.environ.get("BUILD_VERSION")
    version_count = 0
    if styles_xml is not None:
        if version:
            version_count = styles_xml.count("{{BUILD_VERSION}}")
            styles_xml = styles_xml.replace("{{BUILD_VERSION}}", version)
        elif "{{BUILD_VERSION}}" in styles_xml:
            print("WARNING: no build_version provided and {{BUILD_VERSION}} "
                  "placeholder found in styles.xml; leaving placeholder as-is.")

    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "content.xml":
                zout.writestr(item, content_xml)
            elif item.filename == "styles.xml" and styles_xml is not None:
                zout.writestr(item, styles_xml)
            else:
                zout.writestr(item, zin.read(item.filename))
    shutil.move(tmp, path)
    print(f"Fixed header style ({n1}) and row style ({n2}) in {path}")
    if version:
        print(f"Replaced {version_count} {{{{BUILD_VERSION}}}} placeholder(s) with {version}")

if __name__ == "__main__":
    odt_path = sys.argv[1]
    version_arg = sys.argv[2] if len(sys.argv) > 2 else None
    fix_odt(odt_path, version_arg)
