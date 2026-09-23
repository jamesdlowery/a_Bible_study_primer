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
4. Convert the front-cover image, and every full-bleed illustration
   (assemble.py's ILLUSTRATIONS, one per main section plus the Table of
   Contents), from an inline drawing to a full-page anchored (floating)
   one. Pandoc silently constrains an inline image's width to the page's
   text-area width regardless of any explicit width/height markdown
   attributes -- a 9in-wide request was actually emitted as ~7in wide
   (this document's text width) at the image's own aspect ratio, leaving
   a visible white margin on every side instead of the intended
   full-bleed image. Anchoring the drawing to the page itself (rather
   than the text flow), positioned at (0,0) and sized to the exact page
   dimensions, bypasses that text-width constraint entirely, since it's
   no longer part of the inline content flow pandoc applies that limit
   to. Identified by matching each inline drawing's own descr attribute
   against "front-cover.jpg" or any filename containing "Illustration"
   (not by position or a fixed count), so this keeps working as more
   illustrations are added or moved between folders.
5. Retarget the Table of Contents' own list paragraphs from the shared
   "Compact" style to a dedicated "TOCCompact" style (identical except
   for one added right-aligned, dot-leader tab stop), so the ToC's page
   numbers line up consistently regardless of entry length or nesting
   depth. This can't be done from markdown: pandoc's docx writer always
   assigns "Compact" to tight list items, ignoring any custom-style div
   wrapped around them, and "Compact" is shared by roughly 2,600 other
   tight-list paragraphs throughout the rest of the book (plus whatever
   style LibreOffice's own INDEX field generator happens to reuse for
   the compiled back-of-book index) -- giving the shared style its own
   tab stop broke the Index's two-column layout the first time this was
   tried. Scoping the rename to only the text between the "Table of
   Contents" heading's own bookmark and the literal "End of Table of
   Contents" marker text (both emitted by assemble.py specifically to
   bound this section) keeps the change confined to the ToC itself.

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

    # Convert the front-cover image, and every full-bleed illustration
    # (see assemble.py's ILLUSTRATIONS/illustration_md() and
    # build_cover_md()), from a text-width-constrained inline drawing to
    # a full-page anchored one (see module docstring point 4). Matches
    # any inline drawing whose own descr attribute names a file called
    # exactly "front-cover.jpg" or containing "Illustration" -- covers
    # both without depending on a specific folder, so this keeps working
    # if illustrations move between folders later. Processes every match
    # in the document (not just the first), since there are now up to 20
    # of these rather than the original single cover image.
    cover_count = 0
    image_pattern = re.compile(
        r'<w:drawing><wp:inline>.*?descr="([^"]*(?:front-cover\.jpg|Illustration[^"]*))"'
        r'.*?<a:blip r:embed="([^"]+)"\s*/>.*?</wp:inline></w:drawing>',
        flags=re.DOTALL,
    )

    def anchor_image(m):
        nonlocal cover_count
        descr, embed_id = m.group(1), m.group(2)
        # US Letter, full bleed, in EMU (914400 EMU = 1 inch): 8.5in x 11in.
        page_cx, page_cy = 7772400, 9906000
        cover_count += 1
        # Unique per image (there are now up to 21 of these in the same
        # document, not just the original single cover image) -- reusing
        # the same docPr/cNvPr id for all of them was confirmed directly
        # to cause several of them to silently fail to render at all in
        # the full document, despite rendering correctly in isolation
        # (where only one such image exists, so no id collision arises).
        # Spaced 10 apart per image as a safety margin against any other
        # id this document might already assign in the 20-29 range.
        docpr_id = 20 + cover_count * 10
        cnvpr_id = docpr_id + 1
        return (
            '<w:drawing>'
            '<wp:anchor behindDoc="1" distT="0" distB="0" distL="0" distR="0" '
            'simplePos="0" locked="0" layoutInCell="0" allowOverlap="1" relativeHeight="1">'
            '<wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>'
            '<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{page_cx}" cy="{page_cy}"/>'
            '<wp:effectExtent b="0" l="0" r="0" t="0"/>'
            '<wp:wrapNone/>'
            f'<wp:docPr id="{docpr_id}" name="Picture{cover_count}"/>'
            '<wp:cNvGraphicFramePr/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{cnvpr_id}" name="Picture{cover_count}" descr="{descr}"/>'
            '<pic:cNvPicPr><a:picLocks noChangeArrowheads="1" noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{embed_id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr bwMode="auto">'
            f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{page_cx}" cy="{page_cy}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:noFill/><a:ln w="9525"><a:noFill/><a:headEnd/><a:tailEnd/></a:ln>'
            '</pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>'
        )

    doc_xml = image_pattern.sub(anchor_image, doc_xml)

    # Retarget "Compact" -> "TOCCompact" only within the Table of Contents
    # section, bounded by its own heading bookmark and its closing marker
    # text. Both are emitted once, in that exact form, by assemble.py.
    toc_start_marker = '<w:bookmarkStart w:id="20" w:name="table-of-contents" />'
    toc_end_marker = 'End of Table of Contents'
    toc_count = 0
    start_idx = doc_xml.find(toc_start_marker)
    if start_idx == -1:
        # Bookmark ids aren't guaranteed stable across builds; fall back to
        # searching for the bookmark name alone, ignoring its numeric id.
        m = re.search(r'<w:bookmarkStart w:id="\d+" w:name="table-of-contents" />', doc_xml)
        start_idx = m.start() if m else -1
    end_idx = doc_xml.find(toc_end_marker, start_idx if start_idx != -1 else 0)
    if start_idx != -1 and end_idx != -1 and end_idx > start_idx:
        before, toc_section, after = doc_xml[:start_idx], doc_xml[start_idx:end_idx], doc_xml[end_idx:]
        toc_section, toc_count = re.subn(r'w:pStyle w:val="Compact"', 'w:pStyle w:val="TOCCompact"', toc_section)
        doc_xml = before + toc_section + after
    else:
        print("WARNING: could not find Table of Contents start/end markers; "
              "ToC paragraph style left unchanged.")

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
    print(f"Retargeted {toc_count} ToC paragraph(s) to TOCCompact in {path}")
    print(f"Converted {cover_count} full-bleed image(s) (cover + illustrations) to full-page anchored placement in {path}")
    if version:
        print(f"Replaced {version_count} {{{{BUILD_VERSION}}}} placeholder(s) with {version}")

if __name__ == "__main__":
    docx_path = sys.argv[1]
    version_arg = sys.argv[2] if len(sys.argv) > 2 else None
    fix_docx(docx_path, version_arg)
