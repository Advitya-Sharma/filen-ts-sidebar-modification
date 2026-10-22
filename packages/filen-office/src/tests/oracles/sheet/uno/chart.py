# Chart oracle over UNO: loads a workbook in a running headless LibreOffice and reports, per sheet, each chart as
# LibreOffice reads it: its chart types, each series' data (the values LibreOffice draws, from the cells its ranges
# name) with those ranges, its title, whether a legend shows and where, and each axis's scale; and exports the
# workbook as PDF. Started with LibreOffice's bundled Python in isolated mode (-I):
#   chart.py <LibreOffice Resources directory> <pipe name>
# A first line {"ready": true} once connected, then one JSON job per stdin line: {"id", "input", "pdf"}; one JSON result
# per stdout line: {"id", "ok", "charts": [{"sheet", "name", "types", "title", "legend", "series": [{"label",
# "values", "categories", "xValues", "ranges"}], "axes": [{"dimension", "index", "min", "max", "major"}]}], "pdf"}
import json
import os
import sys
import time

RESOURCES = sys.argv[1]
sys.path[0:0] = [RESOURCES, os.path.join(os.path.dirname(RESOURCES), "Frameworks")]

import uno  # noqa: E402
from com.sun.star.beans import PropertyValue  # noqa: E402


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


def number(value):
    try:
        result = float(value)
    except (TypeError, ValueError):
        return None
    return result if result == result else None


def values(sequence):
    if sequence is None:
        return None
    out = []
    for value in sequence.getData():
        if isinstance(value, (int, float)):
            out.append(number(value))
        elif value is None:
            out.append(None)
        else:
            out.append(str(value))
    return out


def text_of(title):
    if title is None:
        return None
    return "".join(part.getString() for part in title.getText())


def axis_scale(axis):
    scale = axis.getScaleData()
    major = None
    try:
        major = number(scale.IncrementData.Distance)
    except Exception:
        major = None
    return number(scale.Minimum), number(scale.Maximum), major


def chart_report(sheet_name, name, document):
    diagram = document.getFirstDiagram()
    types = []
    series_out = []
    axes = []
    for system in diagram.getCoordinateSystems() if diagram is not None else []:
        for dimension in range(system.getDimension()):
            for index in range(system.getMaximumAxisIndexByDimension(dimension) + 1):
                axis = system.getAxisByDimension(dimension, index)
                if axis is None:
                    continue
                low, high, major = axis_scale(axis)
                axes.append({"dimension": dimension, "index": index, "min": low, "max": high, "major": major})
        for chart_type in system.getChartTypes():
            types.append(chart_type.getChartType())
            for series in chart_type.getDataSeries():
                entry = {"label": None, "values": None, "categories": None, "xValues": None, "ranges": []}
                for sequence in series.getDataSequences():
                    data = sequence.getValues()
                    label = sequence.getLabel()
                    role = data.Role if data is not None else ""
                    if data is not None:
                        entry["ranges"].append(data.getSourceRangeRepresentation())
                    if role == "values-y":
                        entry["values"] = values(data)
                    elif role == "values-x":
                        entry["xValues"] = values(data)
                    if label is not None and entry["label"] is None:
                        texts = label.getData()
                        entry["label"] = " ".join(str(text) for text in texts)
                series_out.append(entry)
    categories = None
    try:
        scale = diagram.getCoordinateSystems()[0].getAxisByDimension(0, 0).getScaleData()
        if scale.Categories is not None:
            categories = values(scale.Categories.getValues())
    except Exception:
        categories = None
    for entry in series_out:
        entry["categories"] = categories
    legend = diagram.getLegend() if diagram is not None else None
    position = None
    if legend is not None and legend.getPropertyValue("Show"):
        position = str(legend.getPropertyValue("AnchorPosition").value)
    title = document.getTitleObject()
    return {
        "sheet": sheet_name,
        "name": name,
        "types": types,
        "title": text_of(title),
        "legend": position,
        "series": series_out,
        "axes": axes,
    }


def run(desktop, job):
    url = uno.systemPathToFileUrl(job["input"])
    document = desktop.loadComponentFromURL(url, "_blank", 0, (prop("Hidden", True), prop("ReadOnly", True)))
    try:
        charts = []
        sheets = document.getSheets()
        for index in range(sheets.getCount()):
            sheet = sheets.getByIndex(index)
            table_charts = sheet.getCharts()
            for name in table_charts.getElementNames():
                embedded = table_charts.getByName(name).getEmbeddedObject()
                charts.append(chart_report(sheet.getName(), name, embedded))
        pdf = None
        if job.get("pdf"):
            document.storeToURL(uno.systemPathToFileUrl(job["pdf"]), (prop("FilterName", "calc_pdf_Export"),))
            pdf = os.path.getsize(job["pdf"])
        return {"id": job["id"], "ok": True, "charts": charts, "pdf": pdf}
    finally:
        document.close(True)


def main():
    context = connect(sys.argv[2])
    desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
    print(json.dumps({"ready": True}), flush=True)
    for line in sys.stdin:
        if not line.strip():
            continue
        job = json.loads(line)
        try:
            result = run(desktop, job)
        except Exception as error:  # noqa: BLE001
            result = {"id": job.get("id"), "ok": False, "error": str(error)}
        print(json.dumps(result), flush=True)


main()
