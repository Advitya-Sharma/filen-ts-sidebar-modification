# msoffcrypto-tool oracle for encrypted Word 97-2003 files: decrypts <file> with <password> into <out> (a compound
# file with the WordDocument, Table and Data streams decrypted). Run under a Python with msoffcrypto-tool installed
# (FILEN_OFFICE_MSOFFCRYPTO), always with -I:
#   docDecryptOracle.py <file> <password> <out>
# Prints one JSON line: {"ok": true} or {"ok": false, "error": ...}. The password is never printed.
import json
import sys

import msoffcrypto


def main():
    path, password, out = sys.argv[1], sys.argv[2], sys.argv[3]
    with open(path, "rb") as handle:
        office = msoffcrypto.OfficeFile(handle)
        try:
            office.load_key(password=password)
        except Exception as error:
            print(json.dumps({"ok": False, "error": type(error).__name__}))
            return
        with open(out, "wb") as target:
            office.decrypt(target)
    print(json.dumps({"ok": True}))


main()
