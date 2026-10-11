# Dictionaries

Hunspell dictionaries for Filen's on-device spellchecker, built from the [LibreOffice dictionaries](https://github.com/LibreOffice/dictionaries) repository at commit `32b006a2c22a4ac7e8ed3f03346f7b3d85a970a4`.

Committed here: `manifest.json` (every source URL and SHA-256, and the size and SHA-256 every built file must have) and per dictionary the upstream licence notices (verbatim) with a README listing the source, licence, copyright holders and every modification.

The built `.aff` and `.dic` files are not committed. `pnpm run build:dictionaries` downloads the pinned sources (cached under `node_modules/.cache`), builds them into `../.dictionaries/<id>/` and verifies every file against the manifest; it returns at once when `.dictionaries/.complete` already matches the manifest. `--offline` uses the download cache only.

Changing the sources or the transformation in `scripts/build-dictionaries.ts` (bump `TRANSFORM` when the output changes) needs `pnpm run build:dictionaries --update`, which rewrites this directory. `src/tests/dictionaries.test.ts` checks the manifest and the committed notices; `FILEN_SPELLCHECK_DICTIONARIES=1 pnpm test` also rebuilds everything from scratch and compares it to the manifest.

| Id | Language | Licence | Source |
| --- | --- | --- | --- |
| `en_US` | English (United States) | `SCOWL` | `en/en_US.aff`, `en/en_US.dic` |
| `en_GB` | English (United Kingdom) | `LGPL-2.1-only` | `en/en_GB.aff`, `en/en_GB.dic` |
| `en_CA` | English (Canada) | `SCOWL` | `en/en_CA.aff`, `en/en_CA.dic` |
| `en_AU` | English (Australia) | `SCOWL` | `en/en_AU.aff`, `en/en_AU.dic` |
| `de_DE` | German (Germany) | `GPL-2.0-only OR GPL-3.0-only` | `de/de_DE_frami.aff`, `de/de_DE_frami.dic` |
| `de_AT` | German (Austria) | `GPL-2.0-only OR GPL-3.0-only` | `de/de_AT_frami.aff`, `de/de_AT_frami.dic` |
| `de_CH` | German (Switzerland) | `GPL-2.0-only OR GPL-3.0-only` | `de/de_CH_frami.aff`, `de/de_CH_frami.dic` |
| `fr` | French | `MPL-2.0` | `fr_FR/dictionaries/fr.aff`, `fr_FR/dictionaries/fr.dic` |
| `es_ES` | Spanish (Spain) | `GPL-3.0-or-later OR LGPL-3.0-or-later OR MPL-1.1+` | `es/es_ES.aff`, `es/es_ES.dic` |
| `it_IT` | Italian (Italy) | `GPL-3.0-only` | `it_IT/it_IT.aff`, `it_IT/it_IT.dic` |
| `pt_BR` | Portuguese (Brazil) | `LGPL-3.0-only` | `pt_BR/pt_BR.aff`, `pt_BR/pt_BR.dic` |
| `pt_PT` | Portuguese (Portugal) | `GPL-2.0-only OR LGPL-2.1-only OR MPL-1.1` | `pt_PT/pt_PT.aff`, `pt_PT/pt_PT.dic` |
| `nl_NL` | Dutch (Netherlands) | `BSD-3-Clause OR CC-BY-3.0` | `nl_NL/nl_NL.aff`, `nl_NL/nl_NL.dic` |
| `pl_PL` | Polish (Poland) | `GPL-2.0-only OR LGPL-2.1-only OR MPL-1.1 OR Apache-2.0 OR CC-BY-4.0` | `pl_PL/pl_PL.aff`, `pl_PL/pl_PL.dic` |
| `ru_RU` | Russian (Russia) | `BSD-style (Lebedev)` | `ru_RU/ru_RU.aff`, `ru_RU/ru_RU.dic` |
| `sv_SE` | Swedish (Sweden) | `LGPL-3.0-only` | `sv_SE/dictionaries/sv_SE.aff`, `sv_SE/dictionaries/sv_SE.dic` |
| `da_DK` | Danish (Denmark) | `GPL-2.0-only OR LGPL-2.1-only OR MPL-1.1` | `da_DK/da_DK.aff`, `da_DK/da_DK.dic` |
| `cs_CZ` | Czech (Czechia) | `GPL-1.0-or-later` | `cs_CZ/cs_CZ.aff`, `cs_CZ/cs_CZ.dic` |
