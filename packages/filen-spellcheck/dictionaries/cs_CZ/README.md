# Czech (Czechia) (cs_CZ)

Hunspell spelling dictionary for Filen's on-device spellchecker, built from the [LibreOffice dictionaries](https://github.com/LibreOffice/dictionaries) repository.

- Upstream commit: `32b006a2c22a4ac7e8ed3f03346f7b3d85a970a4`
- Affix file: `cs_CZ/cs_CZ.aff` → `cs_CZ.aff`
- Word list: `cs_CZ/cs_CZ.dic` → `cs_CZ.dic`
- Upstream version: not stated upstream
- Licence: `GPL-1.0-or-later`

README_en.txt and README_cs.txt: "licensed under the GNU/GPL license" without a version number, so any version may be chosen; README_cs.txt carries the GPL 2 text.

Copyright holders, as the upstream notices state them:

- Petr Kolář and numerous contributors (authors of the original Czech ispell dictionary)
- Miroslav Pošta (author of the 2021 update of the Czech dictionary)

## Modifications

- Prepended this modification notice to the .aff as comment lines

## Upstream notices

Copied verbatim from the same commit:

- `README_en.txt` from `cs_CZ/README_en.txt`
- `README_cs.txt` from `cs_CZ/README_cs.txt`

`packages/filen-spellcheck/scripts/build-dictionaries.ts` builds `cs_CZ.aff` and `cs_CZ.dic` into `packages/filen-spellcheck/.dictionaries/cs_CZ/` (not committed) and verifies them byte for byte against `manifest.json`.
