# Converts documents in a running headless LibreOffice over UNO, one load per document. Started by
# soffice.ts with LibreOffice's bundled Python: worker.py <pipe name>. Reads one JSON job per line on stdin
#   {"id", "input", "outdir", "formats": [...], "password"?}
# and prints one JSON result per line: {"id", "ok", "outputs": {format: path}, "pages", "redlines", "error"?}.
# Formats: pdf, odt, fodt, docx (LibreOffice's own save), txt (as loaded: tracked changes shown), txtFinal
# (all changes accepted), txtOriginal (all rejected). Accepting and rejecting run on the loaded copy, never
# stored over the input.
import json
import os
import sys
import time

import uno
from com.sun.star.beans import PropertyValue

FILTERS = {
    "pdf": ("writer_pdf_Export", None),
    "odt": ("writer8", None),
    "fodt": ("OpenDocument Text Flat XML", None),
    "docx": ("MS Word 2007 XML", None),
    "txt": ("Text (encoded)", "UTF8"),
    "txtFinal": ("Text (encoded)", "UTF8"),
    "txtOriginal": ("Text (encoded)", "UTF8"),
}
EXTENSIONS = {"txtFinal": "final.txt", "txtOriginal": "original.txt"}


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


def store(doc, path, name):
    flt, options = FILTERS[name]
    props = [prop("FilterName", flt), prop("Overwrite", True)]
    if options is not None:
        props.append(prop("FilterOptions", options))
    doc.storeToURL(uno.systemPathToFileUrl(os.path.realpath(path)), tuple(props))


def dispatch(ctx, doc, command):
    helper = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
    helper.executeDispatch(doc.getCurrentController().getFrame(), command, "", 0, ())


def pages(doc):
    try:
        cursor = doc.getCurrentController().getViewCursor()
        cursor.jumpToLastPage()
        return cursor.getPage()
    except Exception:
        return -1


def convert(ctx, desktop, job):
    source = job["input"]
    stem = os.path.splitext(os.path.basename(source))[0]
    props = [prop("Hidden", True), prop("ReadOnly", False)]
    if job.get("password"):
        props.append(prop("Password", job["password"]))
    # Real paths: LibreOffice loads nothing through a symbolic link such as macOS's /tmp.
    doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.realpath(source)), "_blank", 0, tuple(props))
    if doc is None:
        return {"ok": False, "error": "not loaded"}
    outputs = {}
    try:
        redlines = doc.getRedlines().getCount() if hasattr(doc, "getRedlines") else 0
        count = pages(doc)
        formats = job["formats"]
        for name in formats:
            if name in ("txtFinal", "txtOriginal"):
                continue
            path = os.path.join(job["outdir"], stem + "." + EXTENSIONS.get(name, name))
            store(doc, path, name)
            outputs[name] = path
        for name, command in (("txtFinal", ".uno:AcceptAllTrackedChanges"), ("txtOriginal", ".uno:RejectAllTrackedChanges")):
            if name not in formats:
                continue
            path = os.path.join(job["outdir"], stem + "." + EXTENSIONS[name])
            if redlines > 0:
                dispatch(ctx, doc, command)
                store(doc, path, name)
                dispatch(ctx, doc, ".uno:Undo")
            else:
                store(doc, path, name)
            outputs[name] = path
        return {"ok": True, "outputs": outputs, "pages": count, "redlines": redlines}
    finally:
        doc.close(True)


def main():
    ctx = connect(sys.argv[1])
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    print(json.dumps({"ready": True}), flush=True)
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        job = json.loads(line)
        started = time.time()
        try:
            result = convert(ctx, desktop, job)
        except Exception as error:
            result = {"ok": False, "error": ("%s: %s" % (type(error).__name__, error))[:500]}
        result["id"] = job["id"]
        result["ms"] = int((time.time() - started) * 1000)
        print(json.dumps(result), flush=True)


main()
