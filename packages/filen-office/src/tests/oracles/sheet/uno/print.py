# Print oracle over UNO: loads a workbook in a running headless LibreOffice and reports, per sheet, the page breaks
# LibreOffice lays out (XSheetPageBreak: automatic and manual), its print areas, and the page count of the whole
# workbook's PDF export. Started with LibreOffice's bundled Python in isolated mode (-I):
#   print.py <LibreOffice Resources directory> <pipe name>
# A first line {"ready": true} once connected, then one JSON job per stdin line: {"id", "input", "pdf", "alone"?}; one JSON
# result per stdout line:
#   {"id", "ok", "sheets": [{"name", "visible", "rowBreaks": [[position, manual]], "colBreaks": [...],
#    "areas": [[startRow, startCol, endRow, endCol]]}], "pdfPages", "alonePages": {name: pages}, "error"?}
# "alone" names sheets also exported on their own (the others removed from a fresh load): the page count of such a PDF
# is what LibreOffice prints for the sheet, fit-to scaling included, which its page break API does not reflect.
import json
import os
import re
import sys
import time

RESOURCES = sys.argv[1]
sys.path[0:0] = [RESOURCES, os.path.join(os.path.dirname(RESOURCES), "Frameworks")]

import uno  # noqa: E402
from com.sun.star.beans import PropertyValue  # noqa: E402

PAGE = re.compile(rb"/Type\s*/Page(?![a-zA-Z])")


def prop(name, value):
    out = PropertyValue()
    out.Name = name
    out.Value = value
    return out


def connect(pipe):
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(240):
        try:
            return resolver.resolve("uno:pipe,name=%s;urp;StarOffice.ComponentContext" % pipe)
        except Exception:
            time.sleep(0.25)
    raise RuntimeError("no office on pipe " + pipe)


def breaks(items):
    return [[item.Position, bool(item.ManualBreak)] for item in items]


def run(desktop, job):
    url = uno.systemPathToFileUrl(job["input"])
    doc = desktop.loadComponentFromURL(url, "_blank", 0, (prop("Hidden", True), prop("ReadOnly", True)))
    try:
        # The export lays every page out: the breaks are read after it.
        doc.storeToURL(uno.systemPathToFileUrl(job["pdf"]), (prop("FilterName", "calc_pdf_Export"),))
        with open(job["pdf"], "rb") as handle:
            pages = len(PAGE.findall(handle.read()))
        sheets = []
        for sheet in doc.getSheets():
            areas = [[a.StartRow, a.StartColumn, a.EndRow, a.EndColumn] for a in sheet.getPrintAreas()]
            sheets.append(
                {
                    "name": sheet.getName(),
                    "visible": bool(sheet.IsVisible),
                    "rowBreaks": breaks(sheet.getRowPageBreaks()),
                    "colBreaks": breaks(sheet.getColumnPageBreaks()),
                    "areas": areas,
                }
            )
    finally:
        doc.close(True)
    alone = {}
    for name in job.get("alone", []):
        alone[name] = alone_pages(desktop, url, name, job["pdf"] + ".alone.pdf")
    return {"id": job["id"], "ok": True, "sheets": sheets, "pdfPages": pages, "alonePages": alone}


def alone_pages(desktop, url, name, pdf):
    doc = desktop.loadComponentFromURL(url, "_blank", 0, (prop("Hidden", True),))
    try:
        sheets = doc.getSheets()
        if not sheets.hasByName(name):
            return None
        for other in list(sheets.getElementNames()):
            if other != name:
                sheets.removeByName(other)
        doc.storeToURL(uno.systemPathToFileUrl(pdf), (prop("FilterName", "calc_pdf_Export"),))
        with open(pdf, "rb") as handle:
            return len(PAGE.findall(handle.read()))
    finally:
        doc.close(True)
        if os.path.exists(pdf):
            os.remove(pdf)


def main():
    context = connect(sys.argv[2])
    desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
    sys.stdout.write(json.dumps({"ready": True}) + "\n")
    sys.stdout.flush()
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        job = json.loads(line)
        try:
            result = run(desktop, job)
        except Exception as error:  # noqa: BLE001
            result = {"id": job.get("id"), "ok": False, "error": str(error)[:500]}
        sys.stdout.write(json.dumps(result) + "\n")
        sys.stdout.flush()


main()
