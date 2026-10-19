# msoffcrypto-tool oracle for encrypted OOXML. Decrypts one CFB-wrapped package with a password and reports
# whether the plaintext is a zip and, optionally, whether it equals an expected plaintext. Run under a Python
# that has msoffcrypto-tool installed (FILEN_OFFICE_MSOFFCRYPTO), always with -I:
#   decrypt.py <file> <password> [expectedPlaintextZip]
# Prints one JSON line: {"ok", "startsPK", "plaintextSize", "integrityVerified", "matchedPlaintext"?} or an error.
import io
import json
import sys

import msoffcrypto


def main():
    path, password = sys.argv[1], sys.argv[2]
    expected = sys.argv[3] if len(sys.argv) > 3 else None
    with open(path, "rb") as handle:
        office = msoffcrypto.OfficeFile(handle)
        try:
            office.load_key(password=password, verify_password=True)
        except Exception as error:
            print(json.dumps({"ok": False, "error": "%s: %s" % (type(error).__name__, error)}))
            return
        out = io.BytesIO()
        # Agile files carry a data-integrity HMAC; ask msoffcrypto to check it.
        integrity = getattr(office, "type", None) == 1 or office.__class__.__name__.lower().find("agile") >= 0
        try:
            office.decrypt(out, verify_integrity=True)
            integrity_verified = True
        except Exception:
            # Older/standard files have no integrity container; decrypt without the check.
            out = io.BytesIO()
            office.decrypt(out)
            integrity_verified = False
    data = out.getvalue()
    result = {
        "ok": True,
        "startsPK": data[:2] == b"PK",
        "plaintextSize": len(data),
        "integrityVerified": integrity_verified,
    }
    if expected is not None:
        with open(expected, "rb") as handle:
            result["matchedPlaintext"] = handle.read() == data
    print(json.dumps(result))


main()
