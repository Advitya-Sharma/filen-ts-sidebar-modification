# Italian (Italy) (it_IT)

Hunspell spelling dictionary for Filen's on-device spellchecker, built from the [LibreOffice dictionaries](https://github.com/LibreOffice/dictionaries) repository.

- Upstream commit: `32b006a2c22a4ac7e8ed3f03346f7b3d85a970a4`
- Affix file: `it_IT/it_IT.aff` → `it_IT.aff`
- Word list: `it_IT/it_IT.dic` → `it_IT.dic`
- Upstream version: 5.1.1 (as stated in `it_IT/README_it_IT.txt`)
- Licence: `GPL-3.0-only`

README_it_IT.txt and the .aff header: GNU GPL version 3 (the licence text closes README_it_IT.txt); the copyright notice and authors of the original "Dizionario italiano" 2.4 come from legacy/.

Copyright holders, as the upstream notices state them:

- Copyright (C) 2020-2022, LibreItalia - Marina Latini
- Copyright (C) 2001, 2002 Gianluca Turconi
- Copyright (C) 2002, 2003, 2004 Gianluca Turconi and Davide Prina
- Copyright (C) 2004, 2005, 2006, 2007 Davide Prina
- Andrea Pescetti (previous maintainer of the Italian Writing Aids extension)

## Modifications

- Prepended this modification notice to the .aff as comment lines

## Upstream notices

Copied verbatim from the same commit:

- `README_it_IT.txt` from `it_IT/README_it_IT.txt`
- `it_IT_license.txt` from `it_IT/legacy/it_IT_license.txt`
- `it_IT_AUTHORS` from `it_IT/legacy/it_IT_AUTHORS`

`packages/filen-spellcheck/scripts/build-dictionaries.ts` builds `it_IT.aff` and `it_IT.dic` into `packages/filen-spellcheck/.dictionaries/it_IT/` (not committed) and verifies them byte for byte against `manifest.json`.
