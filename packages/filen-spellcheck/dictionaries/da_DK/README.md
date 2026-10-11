# Danish (Denmark) (da_DK)

Hunspell spelling dictionary for Filen's on-device spellchecker, built from the [LibreOffice dictionaries](https://github.com/LibreOffice/dictionaries) repository.

- Upstream commit: `32b006a2c22a4ac7e8ed3f03346f7b3d85a970a4`
- Affix file: `da_DK/da_DK.aff` → `da_DK.aff`
- Word list: `da_DK/da_DK.dic` → `da_DK.dic`
- Upstream version: 2.9.101 (as stated in `da_DK/README_da_DK.txt`)
- Licence: `GPL-2.0-only OR LGPL-2.1-only OR MPL-1.1`

README_da_DK.txt and the .aff header: GPL 2.0, LGPL 2.1 or MPL 1.1 (tri-license). The licence texts come from de/ and es/ of the same commit.

Copyright holders, as the upstream notices state them:

- © 2020 Foreningen for frit tilgængelige sprogværktøjer

## Modifications

- Added a final newline to the .aff
- Rewrote the .dic lines `"A/S"`, `"c/o"`, `"I/S"` as `A\/S`, `c\/o`, `I\/S`: the unescaped "/" is read as a flag separator, which FLAG num rejects, and the quotes would keep the entries from ever matching a word
- Prepended this modification notice to the .aff as comment lines

## Upstream notices

Copied verbatim from the same commit:

- `README_da_DK.txt` from `da_DK/README_da_DK.txt`
- `COPYING_GPLv2` from `de/COPYING_GPLv2`
- `COPYING_LGPL_v2.1.txt` from `de/COPYING_LGPL_v2.1.txt`
- `MPL-1.1.txt` from `es/MPL-1.1.txt`

`packages/filen-spellcheck/scripts/build-dictionaries.ts` builds `da_DK.aff` and `da_DK.dic` into `packages/filen-spellcheck/.dictionaries/da_DK/` (not committed) and verifies them byte for byte against `manifest.json`.
