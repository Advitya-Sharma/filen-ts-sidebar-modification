# English (United Kingdom) (en_GB)

Hunspell spelling dictionary for Filen's on-device spellchecker, built from the [LibreOffice dictionaries](https://github.com/LibreOffice/dictionaries) repository.

- Upstream commit: `32b006a2c22a4ac7e8ed3f03346f7b3d85a970a4`
- Affix file: `en/en_GB.aff` → `en_GB.aff`
- Word list: `en/en_GB.dic` → `en_GB.dic`
- Upstream version: 4.0.1 (as stated in `en/en_GB.aff`)
- Licence: `LGPL-2.1-only`

README_en_GB.txt: word list and affix file under the LGPL without a version number, so any version may be chosen; the phonetic table in the .aff is LGPL 2.1 only, so the whole is used under LGPL 2.1. en/ ships no LGPL text: the LGPL 2.1 text comes from de/ of the same commit.

Copyright holders, as the upstream notices state them:

- Kevin Atkinson (author of the original English wordlist)
- David Bartlett, Brian Kelk, Andrew Brown and Marco A.G.Pinto (authors of the British English dictionary)
- © 2000 Björn Jacke
- © OpenStreetMap contributors
- © The Clergy of the Church of England Database Project, 2005.

## Modifications

- Removed the UTF-8 byte order mark from the .aff and .dic
- Prepended this modification notice to the .aff as comment lines

## Upstream notices

Copied verbatim from the same commit:

- `README_en_GB.txt` from `en/README_en_GB.txt`
- `COPYING_LGPL_v2.1.txt` from `de/COPYING_LGPL_v2.1.txt`

`packages/filen-spellcheck/scripts/build-dictionaries.ts` builds `en_GB.aff` and `en_GB.dic` into `packages/filen-spellcheck/.dictionaries/en_GB/` (not committed) and verifies them byte for byte against `manifest.json`.
