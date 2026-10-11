# Dutch (Netherlands) (nl_NL)

Hunspell spelling dictionary for Filen's on-device spellchecker, built from the [LibreOffice dictionaries](https://github.com/LibreOffice/dictionaries) repository.

- Upstream commit: `32b006a2c22a4ac7e8ed3f03346f7b3d85a970a4`
- Affix file: `nl_NL/nl_NL.aff` → `nl_NL.aff`
- Word list: `nl_NL/nl_NL.dic` → `nl_NL.dic`
- Upstream version: 2.20.21 (as stated in `nl_NL/nl_NL.aff`)
- Licence: `BSD-3-Clause OR CC-BY-3.0`

LICENSE.txt and the .aff header: "Revised BSD License and/or CC BY 3.0", at the user's discretion; LICENSE.txt asks to be kept with every copy and contains both licence texts.

Copyright holders, as the upstream notices state them:

- © 2020 OpenTaal (Simon Brouwer, Sander van Geloven)
- © 2006 - 2011 OpenTaal (Ruud Baars, Simon Brouwer)
- © 2001 - 2005 Simon Brouwer and others
- © 1996 Nederlandstalige TeX Gebruikersgroep

## Modifications

- Prepended this modification notice to the .aff as comment lines

## Upstream notices

Copied verbatim from the same commit:

- `LICENSE.txt` from `nl_NL/LICENSE.txt`

`packages/filen-spellcheck/scripts/build-dictionaries.ts` builds `nl_NL.aff` and `nl_NL.dic` into `packages/filen-spellcheck/.dictionaries/nl_NL/` (not committed) and verifies them byte for byte against `manifest.json`.
