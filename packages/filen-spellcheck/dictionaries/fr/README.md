# French (fr)

Hunspell spelling dictionary for Filen's on-device spellchecker, built from the [LibreOffice dictionaries](https://github.com/LibreOffice/dictionaries) repository.

- Upstream commit: `32b006a2c22a4ac7e8ed3f03346f7b3d85a970a4`
- Affix file: `fr_FR/dictionaries/fr.aff` → `fr.aff`
- Word list: `fr_FR/dictionaries/fr.dic` → `fr.dic`
- Upstream version: 7.7 (as stated in `fr_FR/dictionaries/README_dict_fr.txt`)
- Licence: `MPL-2.0`

README_dict_fr.txt: Mozilla Public License 2.0; the .aff header carries the MPL 2.0 Exhibit A notice with the licence URL.

Copyright holders, as the upstream notices state them:

- Olivier R. (author of the Grammalecte French dictionaries)

## Modifications

- Prepended this modification notice to the .aff as comment lines

## Upstream notices

Copied verbatim from the same commit:

- `README_dict_fr.txt` from `fr_FR/dictionaries/README_dict_fr.txt`

`packages/filen-spellcheck/scripts/build-dictionaries.ts` builds `fr.aff` and `fr.dic` into `packages/filen-spellcheck/.dictionaries/fr/` (not committed) and verifies them byte for byte against `manifest.json`.
