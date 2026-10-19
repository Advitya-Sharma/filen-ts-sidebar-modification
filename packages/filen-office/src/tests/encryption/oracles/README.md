# Encryption oracles

Independent checks that our encrypted OOXML output (always sealed as Agile AES-256 / SHA-512 / spin 100,000) is
real: a third-party tool decrypts it with the password and, for Agile, verifies the data-integrity HMAC. Every
oracle is opt-in and skipped when its tool is not configured, so `pnpm test` stays green without any of them.

They are reusable: point an oracle at an **"our encrypted output" directory** — a `manifest.json` of the form

```json
{ "version": 1, "entries": [{ "file": "sealed.docx", "password": "…", "plaintext": "expected.zip" }] }
```

(`plaintext` optional, paths relative to the directory) plus the sealed CFB files it names. `docxEncryptedCorpus.test.ts` writes one with
`FILEN_OFFICE_SEAL_OUT=<dir>` (point `FILEN_OFFICE_SEALED` at `<dir>/all`). With no such directory, `oracles.test.ts` falls back to a fixed set of
corpus samples of the same encryption schemes, so the oracles run wherever the corpus is fetched.

LibreOffice opens copies only; the originals are never handed to it. No password or key is ever logged.

## Environment variables

| Oracle           | Variable                   | Value                                                                      |
| ---------------- | -------------------------- | -------------------------------------------------------------------------- |
| msoffcrypto-tool | `FILEN_OFFICE_MSOFFCRYPTO` | a Python interpreter with `msoffcrypto-tool` installed                     |
| Apache POI       | `FILEN_OFFICE_POI`         | a directory of POI jars (Apache-2.0); `FILEN_OFFICE_JAVA` overrides `java` |
| LibreOffice UNO  | `FILEN_OFFICE_SOFFICE`     | a LibreOffice `soffice` binary (its bundled Python is used)                |
| sealed output    | `FILEN_OFFICE_SEALED`      | the directory described above (else the corpus fallback)                   |

## One command each

All three over the sealed directory (or the corpus fallback), from `packages/filen-office`:

```sh
FILEN_OFFICE_MSOFFCRYPTO=<python> \
FILEN_OFFICE_POI=<poi jars> FILEN_OFFICE_JAVA=<java> \
FILEN_OFFICE_SOFFICE=<soffice> \
FILEN_OFFICE_SEALED=<sealed output> \
npx vitest run src/tests/encryption/oracles/oracles.test.ts
```

Set only the variables for the oracles you have; the rest skip. Behind the test, each oracle is one call:

- **msoffcrypto-tool** — `python -I src/tests/encryption/oracles/msoffcrypto/decrypt.py <file> <password> [expected.zip]`
- **Apache POI** — compile `src/tests/encryption/oracles/poi/OfficeOracle.java` against the jars, then
  `java -cp <jars>:<classes> OfficeOracle <file> <password> [expected.zip]`
- **LibreOffice UNO** — start `soffice --headless --accept=pipe,name=<pipe>;urp;`, then its bundled Python runs
  `src/tests/encryption/oracles/lo/open.py <pipe> <copy> <out.txt> <password>`

Each prints one JSON line (`{ok, startsPK, plaintextSize, hmacVerified, matchedPlaintext?}` or `{ok:false,error}`).

## Tool versions these oracles were checked with

- msoffcrypto-tool 6.0.0 in a venv (PyPI, verified against the published wheel SHA-256).
- Apache POI 5.3.0 + runtime deps from Maven Central (jar SHA-1 verified against the published `.sha1`).
- LibreOffice 7.5.0.3 (its bundled Python).
