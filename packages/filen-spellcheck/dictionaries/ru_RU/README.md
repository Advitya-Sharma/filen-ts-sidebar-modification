# Russian (Russia) (ru_RU)

Hunspell spelling dictionary for Filen's on-device spellchecker, built from the [LibreOffice dictionaries](https://github.com/LibreOffice/dictionaries) repository.

- Upstream commit: `32b006a2c22a4ac7e8ed3f03346f7b3d85a970a4`
- Affix file: `ru_RU/ru_RU.aff` → `ru_RU.aff`
- Word list: `ru_RU/ru_RU.dic` → `ru_RU.dic`
- Upstream version: not stated upstream
- Licence: `BSD-style (Lebedev)`

README_ru_RU.txt: BSD-style licence by Alexander I. Lebedev with the extra condition "Modified versions must be clearly marked as such".

Copyright holders, as the upstream notices state them:

- Copyright (c) 1997-2008, Alexander I. Lebedev

## Modifications

- Prepended this modification notice to the .aff as comment lines

## Upstream notices

Copied verbatim from the same commit:

- `README_ru_RU.txt` from `ru_RU/README_ru_RU.txt`

`packages/filen-spellcheck/scripts/build-dictionaries.ts` builds `ru_RU.aff` and `ru_RU.dic` into `packages/filen-spellcheck/.dictionaries/ru_RU/` (not committed) and verifies them byte for byte against `manifest.json`.
