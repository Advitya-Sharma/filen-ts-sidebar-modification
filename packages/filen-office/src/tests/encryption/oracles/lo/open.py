# LibreOffice UNO oracle for encrypted OOXML. Connects to a headless soffice over a pipe, opens one
# password-protected file (a copy, never the original) and writes its plain text out, proving LibreOffice
# accepts the password and the package. Started by lo.ts with LibreOffice's bundled Python:
#   open.py <pipe name> <input copy> <out.txt> <password>
# Prints one JSON line: {"ok", "textLength", "pages"} or {"ok": false, "error": ...}.
import json
import os
import sys
import time

import uno
from com.sun.star.beans import PropertyValue


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


def pages(doc):
    try:
        cursor = doc.getCurrentController().getViewCursor()
        cursor.jumpToLastPage()
        return cursor.getPage()
    except Exception:
        return -1


def main():
    pipe, source, out_txt, password = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
    ctx = connect(pipe)
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    props = [prop("Hidden", True), prop("ReadOnly", True), prop("Password", password)]
    # Real path: LibreOffice loads nothing through a symbolic link such as macOS's /tmp.
    url = uno.systemPathToFileUrl(os.path.realpath(source))
    try:
        doc = desktop.loadComponentFromURL(url, "_blank", 0, tuple(props))
    except Exception as error:
        # A wrong password (or an unsupported scheme) makes type detection abort before a document exists.
        print(json.dumps({"ok": False, "error": "%s: %s" % (type(error).__name__, error)}), flush=True)
        return
    if doc is None:
        print(json.dumps({"ok": False, "error": "not loaded (wrong password or unsupported)"}), flush=True)
        return
    try:
        text = doc.getText().getString() if hasattr(doc, "getText") else ""
        count = pages(doc)
        store = [prop("FilterName", "Text (encoded)"), prop("FilterOptions", "UTF8"), prop("Overwrite", True)]
        try:
            doc.storeToURL(uno.systemPathToFileUrl(os.path.realpath(out_txt)), tuple(store))
        except Exception:
            pass
        print(json.dumps({"ok": True, "textLength": len(text), "pages": count}), flush=True)
    finally:
        doc.close(True)


main()
