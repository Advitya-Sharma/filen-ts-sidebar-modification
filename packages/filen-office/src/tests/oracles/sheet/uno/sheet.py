# Spreadsheet oracle over UNO: loads a workbook in a running headless LibreOffice, hard-recalculates it
# (calculateAll), dumps every non-empty cell, optionally replays structure operations, dumps again and stores
# the result as xlsx. Started by uno.ts with LibreOffice's bundled Python in isolated mode (-I):
#   sheet.py <LibreOffice Resources directory> <pipe name>
# -I leaves PYTHONPATH out, so the UNO modules are put on sys.path here, from the install's own directories.
#
# One JSON job per stdin line:
#   {"id", "input", "out", "ops"?: [...], "saveAs"?: path, "maxCells"?: n, "password"?}
# one JSON result per stdout line:
#   {"id", "ok", "ms", "cells", "truncated", "calculated", "skipCalc"?, "ops"?: [{"ok", "error"?}], "saved"?, "error"?}
# The dump goes to "out": {"before": Workbook, "after"?: Workbook}, Workbook = {"sheets": [{"name", "visible",
# "cells": [[row, column, type, value, shown, formula?]]}], "names": [{"name", "content"}]}; rows and columns
# 0-based; type "n" number, "s" text, "b" boolean, "e" error (value = LibreOffice's error text). Operations:
#   {"op": "insertRows"|"deleteRows"|"insertColumns"|"deleteColumns", "sheet", "at", "count"}
#   {"op": "sort", "sheet", "range": "A1:D9", "keys": [{"column", "ascending"?, "caseSensitive"?}], "header"?}
#   {"op": "deleteSheet", "sheet"}  {"op": "moveSheet", "sheet", "to"}  ("to": the final 0-based position)
#   {"op": "mergeCells"|"unmergeCells", "sheet", "range"}  {"op": "hideRows"|"hideColumns", "sheet", "at", "count", "hidden"?}
#   {"op": "freeze", "sheet", "rows", "cols"}  {"op": "copySheet", "sheet", "name", "to"}  {"op": "renameSheet", "sheet", "name"}
#   {"op": "hideSheet", "sheet", "hidden"?}  {"op": "tabColor", "sheet", "color"}  {"op": "protectSheet", "sheet", "password"?}
# "sheet" is a name or a 0-based index; a sort key's column is 0-based within the range.
# "views": true adds each sheet's view as loaded: {"name", "frozen", "splitRow", "splitColumn", "panes": [[firstRow,
# firstColumn], ...]} (the view's panes in LibreOffice's order: XViewPane.getFirstVisibleRow/Column).
import json
import os
import re
import sys
import time

RESOURCES = sys.argv[1]
sys.path[0:0] = [RESOURCES, os.path.join(os.path.dirname(RESOURCES), "Frameworks")]

import uno  # noqa: E402
from com.sun.star.beans import PropertyValue  # noqa: E402
from com.sun.star.table import TableSortField  # noqa: E402

# com.sun.star.sheet.CellFlags: VALUE | DATETIME | STRING | FORMULA.
CONTENT = 1 | 2 | 4 | 16
# com.sun.star.util.NumberFormat.LOGICAL.
LOGICAL = 1024
# Functions that reach the network or another program when calculated: such a workbook is not recalculated.
OUTSIDE = re.compile(r"\b(WEBSERVICE|STOCKHISTORY|RTD|DDE)\s*\(", re.I)
ROWS_PER_READ = 4096


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


def sheet_of(doc, ref):
    sheets = doc.getSheets()
    return sheets.getByIndex(ref) if isinstance(ref, int) else sheets.getByName(ref)


def content_ranges(sheet):
    return sheet.queryContentCells(CONTENT).getRangeAddresses()


def formulas(doc):
    for sheet in doc.getSheets():
        for address in content_ranges(sheet):
            for start in range(address.StartRow, address.EndRow + 1, ROWS_PER_READ):
                end = min(address.EndRow, start + ROWS_PER_READ - 1)
                block = sheet.getCellRangeByPosition(address.StartColumn, start, address.EndColumn, end)
                for row in block.getFormulaArray():
                    for text in row:
                        if text.startswith("="):
                            yield text


def dump(doc, budget):
    formats = doc.getNumberFormats()
    logical = {}
    sheets = []
    count = 0
    truncated = False
    for sheet in doc.getSheets():
        cells = []
        for address in content_ranges(sheet):
            for start in range(address.StartRow, address.EndRow + 1, ROWS_PER_READ):
                end = min(address.EndRow, start + ROWS_PER_READ - 1)
                block = sheet.getCellRangeByPosition(address.StartColumn, start, address.EndColumn, end)
                values = block.getDataArray()
                texts = block.getFormulaArray()
                for r, row in enumerate(values):
                    if truncated:
                        break
                    for c, value in enumerate(row):
                        formula = texts[r][c]
                        if value == "" and formula == "":
                            continue
                        if count >= budget:
                            truncated = True
                            break
                        count += 1
                        cell = block.getCellByPosition(c, r)
                        shown = cell.getString()
                        if value is None or cell.getError() != 0:
                            kind, value = "e", shown
                        elif isinstance(value, str):
                            kind = "s"
                        else:
                            kind = "n"
                            if shown in ("TRUE", "FALSE") and value in (0.0, 1.0):
                                key = cell.getPropertyValue("NumberFormat")
                                if key not in logical:
                                    logical[key] = (formats.getByKey(key).Type & LOGICAL) != 0
                                if logical[key]:
                                    kind, value = "b", value == 1.0
                        entry = [start + r, address.StartColumn + c, kind, value, shown]
                        if formula.startswith("="):
                            entry.append(formula)
                        cells.append(entry)
        cells.sort(key=lambda cell: (cell[0], cell[1]))
        sheets.append({"name": sheet.getName(), "visible": bool(sheet.getPropertyValue("IsVisible")), "cells": cells})
    names = []
    ranges = doc.getPropertyValue("NamedRanges")
    for name in ranges.getElementNames():
        names.append({"name": name, "content": ranges.getByName(name).getContent()})
    return {"sheets": sheets, "names": names}, count, truncated


def apply(doc, op):
    kind = op["op"]
    if kind in ("insertRows", "deleteRows", "insertColumns", "deleteColumns"):
        sheet = sheet_of(doc, op["sheet"])
        axis = sheet.getRows() if kind.endswith("Rows") else sheet.getColumns()
        if kind.startswith("insert"):
            axis.insertByIndex(op["at"], op["count"])
        else:
            axis.removeByIndex(op["at"], op["count"])
    elif kind == "sort":
        block = sheet_of(doc, op["sheet"]).getCellRangeByName(op["range"])
        fields = []
        for key in op["keys"]:
            field = TableSortField()
            field.Field = key["column"]
            field.IsAscending = key.get("ascending", True)
            field.IsCaseSensitive = key.get("caseSensitive", False)
            fields.append(field)
        descriptor = list(block.createSortDescriptor())
        for index, item in enumerate(descriptor):
            if item.Name == "SortFields":
                descriptor[index] = prop("SortFields", uno.Any("[]com.sun.star.table.TableSortField", tuple(fields)))
            elif item.Name == "ContainsHeader":
                descriptor[index] = prop("ContainsHeader", bool(op.get("header", False)))
            elif item.Name == "IsSortColumns":
                descriptor[index] = prop("IsSortColumns", False)
        uno.invoke(block, "sort", (tuple(descriptor),))
    elif kind == "deleteSheet":
        doc.getSheets().removeByName(sheet_of(doc, op["sheet"]).getName())
    elif kind == "moveSheet":
        sheets = doc.getSheets()
        name = sheet_of(doc, op["sheet"]).getName()
        current = list(sheets.getElementNames()).index(name)
        # moveByName inserts before the sheet at the given position of the current order.
        sheets.moveByName(name, op["to"] + 1 if op["to"] > current else op["to"])
    elif kind == "mergeCells":
        sheet_of(doc, op["sheet"]).getCellRangeByName(op["range"]).merge(True)
    elif kind == "unmergeCells":
        sheet_of(doc, op["sheet"]).getCellRangeByName(op["range"]).merge(False)
    elif kind in ("hideRows", "hideColumns"):
        sheet = sheet_of(doc, op["sheet"])
        axis = sheet.getRows() if kind == "hideRows" else sheet.getColumns()
        for index in range(op["at"], op["at"] + op["count"]):
            axis.getByIndex(index).setPropertyValue("IsVisible", not op.get("hidden", True))
    elif kind == "freeze":
        controller = doc.getCurrentController()
        controller.setActiveSheet(sheet_of(doc, op["sheet"]))
        controller.freezeAtPosition(op["cols"], op["rows"])
    elif kind == "copySheet":
        sheets = doc.getSheets()
        sheets.copyByName(sheet_of(doc, op["sheet"]).getName(), op["name"], op["to"])
    elif kind == "renameSheet":
        sheet_of(doc, op["sheet"]).setName(op["name"])
    elif kind == "hideSheet":
        sheet_of(doc, op["sheet"]).setPropertyValue("IsVisible", not op.get("hidden", True))
    elif kind == "tabColor":
        sheet_of(doc, op["sheet"]).setPropertyValue("TabColor", op["color"])
    elif kind == "protectSheet":
        sheet_of(doc, op["sheet"]).protect(op.get("password", ""))
    else:
        raise ValueError("unknown operation " + kind)


def views(doc):
    controller = doc.getCurrentController()
    out = []
    sheets = doc.getSheets()
    for index in range(sheets.getCount()):
        sheet = sheets.getByIndex(index)
        controller.setActiveSheet(sheet)
        panes = []
        for k in range(controller.getCount()):
            pane = controller.getByIndex(k)
            panes.append([pane.getFirstVisibleRow(), pane.getFirstVisibleColumn()])
        out.append({
            "name": sheet.getName(),
            "frozen": bool(controller.hasFrozenPanes()),
            "splitRow": controller.getSplitRow(),
            "splitColumn": controller.getSplitColumn(),
            "panes": panes,
        })
    return out


def run(desktop, job):
    props = [prop("Hidden", True), prop("ReadOnly", False), prop("MacroExecutionMode", 0), prop("UpdateDocMode", 0)]
    if job.get("password"):
        props.append(prop("Password", job["password"]))
    # Real paths: LibreOffice loads nothing through a symbolic link such as macOS's /tmp.
    url = uno.systemPathToFileUrl(os.path.realpath(job["input"]))
    doc = desktop.loadComponentFromURL(url, "_blank", 0, tuple(props))
    if doc is None:
        return {"ok": False, "error": "not loaded"}
    try:
        result = {"ok": True}
        outside = next((text for text in formulas(doc) if OUTSIDE.search(text)), None)
        result["calculated"] = outside is None
        if outside is None:
            doc.calculateAll()
        else:
            result["skipCalc"] = outside[:200]
        if job.get("views"):
            result["views"] = views(doc)
        budget = int(job.get("maxCells", 250000))
        before, count, truncated = dump(doc, budget)
        out = {"before": before}
        if job.get("ops"):
            applied = []
            for op in job["ops"]:
                try:
                    apply(doc, op)
                    applied.append({"ok": True})
                except Exception as error:
                    applied.append({"ok": False, "error": ("%s: %s" % (type(error).__name__, error))[:300]})
            if outside is None:
                doc.calculateAll()
            out["after"], after_count, after_truncated = dump(doc, budget)
            count += after_count
            truncated = truncated or after_truncated
            result["ops"] = applied
        if job.get("saveAs"):
            target = uno.systemPathToFileUrl(os.path.realpath(os.path.dirname(job["saveAs"])) + "/" + os.path.basename(job["saveAs"]))
            doc.storeToURL(target, (prop("FilterName", "Calc MS Excel 2007 XML"), prop("Overwrite", True)))
            result["saved"] = job["saveAs"]
        with open(job["out"], "w", encoding="utf-8") as handle:
            json.dump(out, handle, ensure_ascii=False, separators=(",", ":"))
        result["cells"] = count
        result["truncated"] = truncated
        return result
    finally:
        doc.close(True)


def main():
    ctx = connect(sys.argv[2])
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    print(json.dumps({"ready": True}), flush=True)
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        job = json.loads(line)
        started = time.time()
        try:
            result = run(desktop, job)
        except Exception as error:
            result = {"ok": False, "error": ("%s: %s" % (type(error).__name__, error))[:500]}
        result["id"] = job["id"]
        result["ms"] = int((time.time() - started) * 1000)
        print(json.dumps(result), flush=True)


main()
