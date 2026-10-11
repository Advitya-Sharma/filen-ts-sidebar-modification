# Portuguese (Portugal) (pt_PT)

Hunspell spelling dictionary for Filen's on-device spellchecker, built from the [LibreOffice dictionaries](https://github.com/LibreOffice/dictionaries) repository.

- Upstream commit: `32b006a2c22a4ac7e8ed3f03346f7b3d85a970a4`
- Affix file: `pt_PT/pt_PT.aff` → `pt_PT.aff`
- Word list: `pt_PT/pt_PT.dic` → `pt_PT.dic`
- Upstream version: not stated upstream
- Licence: `GPL-2.0-only OR LGPL-2.1-only OR MPL-1.1`

README_pt_PT.txt: "GPL/LGPL/MPL", versions GPL 2, LGPL 2.1 and MPL 1.1 (LICENSES.txt summarises the spellchecker as "GPL and BSD" without detail). The licence texts come from de/ and es/ of the same commit.

Copyright holders, as the upstream notices state them:

- Copyright (C) 2006-2012 José João de Almeida, Rui Vilela, Alberto Simões

## Modifications

- Prepended this modification notice to the .aff as comment lines

## Upstream notices

Copied verbatim from the same commit:

- `README_pt_PT.txt` from `pt_PT/README_pt_PT.txt`
- `LICENSES.txt` from `pt_PT/LICENSES.txt`
- `COPYING_GPLv2` from `de/COPYING_GPLv2`
- `COPYING_LGPL_v2.1.txt` from `de/COPYING_LGPL_v2.1.txt`
- `MPL-1.1.txt` from `es/MPL-1.1.txt`

`packages/filen-spellcheck/scripts/build-dictionaries.ts` builds `pt_PT.aff` and `pt_PT.dic` into `packages/filen-spellcheck/.dictionaries/pt_PT/` (not committed) and verifies them byte for byte against `manifest.json`.
