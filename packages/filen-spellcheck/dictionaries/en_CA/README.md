# English (Canada) (en_CA)

Hunspell spelling dictionary for Filen's on-device spellchecker, built from the [LibreOffice dictionaries](https://github.com/LibreOffice/dictionaries) repository.

- Upstream commit: `32b006a2c22a4ac7e8ed3f03346f7b3d85a970a4`
- Affix file: `en/en_CA.aff` → `en_CA.aff`
- Word list: `en/en_CA.dic` → `en_CA.dic`
- Upstream version: 2020.12.07 (as stated in `en/README_en_CA.txt`)
- Licence: `SCOWL`

The README: word lists under the SCOWL and VarCon permission notices (Kevin Atkinson, Benjamin Titze), the affix file under Geoff Kuenning's BSD-style Ispell licence (modified versions must be marked), plus the WordNet and UKACD notices; UKACD requires its notice to ship verbatim, hence the README.

Copyright holders, as the upstream notices state them:

- Copyright 2000-2018 by Kevin Atkinson
- Copyright 2000-2016 by Kevin Atkinson
- Copyright 2016 by Benjamin Titze
- Copyright 1993, Geoff Kuenning, Granada Hills, CA
- Copyright (c) J Ross Beresford 1993-1999. All Rights Reserved.
- WordNet 1.6 Copyright 1997 by Princeton University. All rights reserved.

## Modifications

- Prepended this modification notice to the .aff as comment lines

## Upstream notices

Copied verbatim from the same commit:

- `README_en_CA.txt` from `en/README_en_CA.txt`

`packages/filen-spellcheck/scripts/build-dictionaries.ts` builds `en_CA.aff` and `en_CA.dic` into `packages/filen-spellcheck/.dictionaries/en_CA/` (not committed) and verifies them byte for byte against `manifest.json`.
