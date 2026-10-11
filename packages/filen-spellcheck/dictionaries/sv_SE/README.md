# Swedish (Sweden) (sv_SE)

Hunspell spelling dictionary for Filen's on-device spellchecker, built from the [LibreOffice dictionaries](https://github.com/LibreOffice/dictionaries) repository.

- Upstream commit: `32b006a2c22a4ac7e8ed3f03346f7b3d85a970a4`
- Affix file: `sv_SE/dictionaries/sv_SE.aff` → `sv_SE.aff`
- Word list: `sv_SE/dictionaries/sv_SE.dic` → `sv_SE.dic`
- Upstream version: not stated upstream
- Licence: `LGPL-3.0-only`

LICENSE_en_US.txt: GNU Lesser General Public License Version 3. sv_SE/ ships no licence text: LGPL 3 and the GPL 3 it builds on come from es/ of the same commit.

Copyright holders, as the upstream notices state them:

- Copyright © 2003-19 Göran Andersson

## Modifications

- Prepended this modification notice to the .aff as comment lines

## Upstream notices

Copied verbatim from the same commit:

- `LICENSE_en_US.txt` from `sv_SE/LICENSE_en_US.txt`
- `LICENSE_sv_SE.txt` from `sv_SE/LICENSE_sv_SE.txt`
- `LGPLv3.txt` from `es/LGPLv3.txt`
- `GPLv3.txt` from `es/GPLv3.txt`

`packages/filen-spellcheck/scripts/build-dictionaries.ts` builds `sv_SE.aff` and `sv_SE.dic` into `packages/filen-spellcheck/.dictionaries/sv_SE/` (not committed) and verifies them byte for byte against `manifest.json`.
