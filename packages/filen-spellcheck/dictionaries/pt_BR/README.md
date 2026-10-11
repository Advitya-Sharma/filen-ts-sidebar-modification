# Portuguese (Brazil) (pt_BR)

Hunspell spelling dictionary for Filen's on-device spellchecker, built from the [LibreOffice dictionaries](https://github.com/LibreOffice/dictionaries) repository.

- Upstream commit: `32b006a2c22a4ac7e8ed3f03346f7b3d85a970a4`
- Affix file: `pt_BR/pt_BR.aff` → `pt_BR.aff`
- Word list: `pt_BR/pt_BR.dic` → `pt_BR.dic`
- Upstream version: 3.2 (as stated in `pt_BR/pt_BR.aff`)
- Licence: `LGPL-3.0-only`

README_pt_BR.txt, README_en.txt and the .aff header: GNU LGPL version 3 and the Mozilla Public License (no MPL version stated). pt_BR/ ships no licence text: LGPL 3 and the GPL 3 it builds on come from es/ of the same commit.

Copyright holders, as the upstream notices state them:

- Copyright (C) 2006 - 2013 by Raimundo Santos Moura

## Modifications

- Removed the UTF-8 byte order mark from the .aff and .dic
- Prepended this modification notice to the .aff as comment lines

## Upstream notices

Copied verbatim from the same commit:

- `README_pt_BR.txt` from `pt_BR/README_pt_BR.txt`
- `README_en.txt` from `pt_BR/README_en.txt`
- `LGPLv3.txt` from `es/LGPLv3.txt`
- `GPLv3.txt` from `es/GPLv3.txt`

`packages/filen-spellcheck/scripts/build-dictionaries.ts` builds `pt_BR.aff` and `pt_BR.dic` into `packages/filen-spellcheck/.dictionaries/pt_BR/` (not committed) and verifies them byte for byte against `manifest.json`.
