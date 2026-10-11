# Polish (Poland) (pl_PL)

Hunspell spelling dictionary for Filen's on-device spellchecker, built from the [LibreOffice dictionaries](https://github.com/LibreOffice/dictionaries) repository.

- Upstream commit: `32b006a2c22a4ac7e8ed3f03346f7b3d85a970a4`
- Affix file: `pl_PL/pl_PL.aff` → `pl_PL.aff`
- Word list: `pl_PL/pl_PL.dic` → `pl_PL.dic`
- Upstream version: 2026-05-11 (as stated in `pl_PL/README_pl_PL.txt`)
- Licence: `GPL-2.0-only OR LGPL-2.1-only OR MPL-1.1 OR Apache-2.0 OR CC-BY-4.0`

README_pl_PL.txt: GPL 2, LGPL 2.1, MPL 1.1, Apache 2.0 and Creative Commons Attribution 4.0 International.

Copyright holders, as the upstream notices state them:

- sjp.pl (authors of the Polish dictionary)

## Modifications

- Converted from ISO-8859-2 to UTF-8 (SET line rewritten to "SET UTF-8")
- Prepended this modification notice to the .aff as comment lines

## Upstream notices

Copied verbatim from the same commit:

- `README_pl_PL.txt` from `pl_PL/README_pl_PL.txt`

`packages/filen-spellcheck/scripts/build-dictionaries.ts` builds `pl_PL.aff` and `pl_PL.dic` into `packages/filen-spellcheck/.dictionaries/pl_PL/` (not committed) and verifies them byte for byte against `manifest.json`.
