# German (Switzerland) (de_CH)

Hunspell spelling dictionary for Filen's on-device spellchecker, built from the [LibreOffice dictionaries](https://github.com/LibreOffice/dictionaries) repository.

- Upstream commit: `32b006a2c22a4ac7e8ed3f03346f7b3d85a970a4`
- Affix file: `de/de_CH_frami.aff` → `de_CH.aff`
- Word list: `de/de_CH_frami.dic` → `de_CH.dic`
- Upstream version: 20161207+frami20170109 (as stated in `de/de_CH_frami.aff`)
- Licence: `GPL-2.0-only OR GPL-3.0-only`

README_de_DE_frami.txt (covers every xxx_frami dictionary): "GNU GPL, Version 2 oder 3"; the .aff header: "License: GPLv2, GPLv3", with a copy of both licences in every distribution.

Copyright holders, as the upstream notices state them:

- Bjoern Jacke (author of the igerman98 base dictionary)
- Franz Michael Baumann (author of the frami extension)

## Modifications

- Converted from ISO-8859-1 to UTF-8 (SET line rewritten to "SET UTF-8")
- Prepended this modification notice to the .aff as comment lines

## Upstream notices

Copied verbatim from the same commit:

- `README_de_DE_frami.txt` from `de/README_de_DE_frami.txt`
- `COPYING_GPLv2` from `de/COPYING_GPLv2`
- `COPYING_GPLv3` from `de/COPYING_GPLv3`

`packages/filen-spellcheck/scripts/build-dictionaries.ts` builds `de_CH.aff` and `de_CH.dic` into `packages/filen-spellcheck/.dictionaries/de_CH/` (not committed) and verifies them byte for byte against `manifest.json`.
