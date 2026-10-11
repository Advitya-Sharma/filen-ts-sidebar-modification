# Spanish (Spain) (es_ES)

Hunspell spelling dictionary for Filen's on-device spellchecker, built from the [LibreOffice dictionaries](https://github.com/LibreOffice/dictionaries) repository.

- Upstream commit: `32b006a2c22a4ac7e8ed3f03346f7b3d85a970a4`
- Affix file: `es/es_ES.aff` → `es_ES.aff`
- Word list: `es/es_ES.dic` → `es_ES.dic`
- Upstream version: not stated upstream
- Licence: `GPL-3.0-or-later OR LGPL-3.0-or-later OR MPL-1.1+`

README_hunspell_es.txt and LICENSE.md: GNU GPL 3 or later, GNU LGPL 3 or later or MPL 1.1 or later, at the user's choice; LICENSE.md names the three licence texts.

Copyright holders, as the upstream notices state them:

- Santiago Bosio (author of the RLA-ES Spanish dictionary)

## Modifications

- Prepended this modification notice to the .aff as comment lines

## Upstream notices

Copied verbatim from the same commit:

- `README_hunspell_es.txt` from `es/README_hunspell_es.txt`
- `LICENSE.md` from `es/LICENSE.md`
- `GPLv3.txt` from `es/GPLv3.txt`
- `LGPLv3.txt` from `es/LGPLv3.txt`
- `MPL-1.1.txt` from `es/MPL-1.1.txt`

`packages/filen-spellcheck/scripts/build-dictionaries.ts` builds `es_ES.aff` and `es_ES.dic` into `packages/filen-spellcheck/.dictionaries/es_ES/` (not committed) and verifies them byte for byte against `manifest.json`.
