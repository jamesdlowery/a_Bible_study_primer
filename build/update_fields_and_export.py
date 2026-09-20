#!/usr/bin/env python3
"""Open a docx/odt in a headless LibreOffice instance, force every field
and index (Table of Contents, PAGEREF, STYLEREF, NUMPAGES, etc.) to
recompute against the document's actual pagination, then export to PDF.

Why this exists: `soffice --headless --convert-to pdf file.docx` does NOT
resolve fields on its own -- it renders whatever cached field text is
already baked into the file (usually a placeholder like "Right-click and
select Update Field to generate table of contents"), because LibreOffice
normally only recalculates fields interactively, on open, when a human is
driving. Setting <w:updateFields w:val="true"/> in the docx's own
settings.xml does not change this for a headless conversion either --
confirmed by direct testing. The reliable fix is to drive LibreOffice
through its UNO scripting API, which can force a genuine field/index
recalculation before the PDF is written.

Usage: python3 update_fields_and_export.py <input.docx|odt> <output.pdf>
"""
import subprocess
import sys
import os
import time


def main():
    if len(sys.argv) != 3:
        print("Usage: update_fields_and_export.py <input> <output.pdf>", file=sys.stderr)
        sys.exit(1)

    in_path = os.path.abspath(sys.argv[1])
    out_path = os.path.abspath(sys.argv[2])

    proc = subprocess.Popen([
        "soffice", "--headless", "--invisible", "--nocrashreport", "--nodefault",
        "--norestore", "--nologo", "--nofirststartwizard",
        "--accept=socket,host=localhost,port=2002;urp;",
    ])
    try:
        # Give the LibreOffice process time to start listening. A fixed
        # sleep is crude but simple; retried connection below covers the
        # rest of the startup variance.
        time.sleep(6)

        import uno
        from com.sun.star.beans import PropertyValue

        def make_prop(name, value):
            p = PropertyValue()
            p.Name = name
            p.Value = value
            return p

        local_ctx = uno.getComponentContext()
        resolver = local_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx)

        ctx = None
        last_err = None
        for _ in range(10):
            try:
                ctx = resolver.resolve(
                    "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
                break
            except Exception as e:  # noqa: BLE001 - resolver raises a UNO-specific type
                last_err = e
                time.sleep(2)
        if ctx is None:
            raise RuntimeError(f"Could not connect to LibreOffice: {last_err}")

        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

        in_url = "file://" + in_path
        out_url = "file://" + out_path

        doc = desktop.loadComponentFromURL(
            in_url, "_blank", 0, (make_prop("Hidden", True),))

        try:
            # Update every index (Table of Contents, any other TOX) --
            # this is what actually resolves TOC/PAGEREF-style fields.
            if hasattr(doc, "getDocumentIndexes"):
                indexes = doc.getDocumentIndexes()
                for i in range(indexes.getCount()):
                    indexes.getByIndex(i).update()

            # Update remaining simple fields (STYLEREF, PAGE, NUMPAGES,
            # etc.) throughout the document body and in headers/footers.
            if hasattr(doc, "getTextFields"):
                doc.getTextFields().refresh()

            export_props = (make_prop("FilterName", "writer_pdf_Export"),)
            doc.storeToURL(out_url, export_props)
        finally:
            doc.close(False)

        print(f"Wrote {out_path}")
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            proc.kill()


if __name__ == "__main__":
    main()
