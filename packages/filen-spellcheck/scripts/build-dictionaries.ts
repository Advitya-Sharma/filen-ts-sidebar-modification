import { createHash } from "node:crypto"
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { dirname, join, posix } from "node:path"
import { fileURLToPath } from "node:url"

// Builds the Hunspell dictionaries from the LibreOffice dictionaries repository at the commit pinned below: every source
// file is pinned by SHA-256, transcoded to UTF-8, given its documented fixes and a modification notice. The built .aff/.dic
// go to .dictionaries/ (not committed); dictionaries/ holds what is committed: manifest.json with the expected output
// hashes, and per dictionary its upstream licence notices and a README. Output is byte-for-byte deterministic.
//
//   node scripts/build-dictionaries.ts [--offline]            build .dictionaries/ and verify it against the manifest
//   node scripts/build-dictionaries.ts --update [--offline]   maintainers: also rewrite dictionaries/

export const UPSTREAM = {
	repository: "https://github.com/LibreOffice/dictionaries",
	commit: "32b006a2c22a4ac7e8ed3f03346f7b3d85a970a4"
} as const

// Bump when the output for unchanged sources changes.
export const TRANSFORM = 1

const RAW = `https://raw.githubusercontent.com/LibreOffice/dictionaries/${UPSTREAM.commit}/`

export const SOURCES = {
	"en/en_US.aff": "e746c882dd6f303c2c46e7452804b9201115a6942cfeb15f18f8edf774d2e24e",
	"en/en_US.dic": "f0b1a234bd178bdd01875b2a392a9647f888b8fe879f79c52aae62c2759b3647",
	"en/README_en_US.txt": "168b4c01cb841f766a72f310b065c04d09cff46376c37a81d799593ce751c371",
	"en/en_GB.aff": "0fd6ed120ef28957847d98ba5149b117e27116cf81b5aa36208453f6755a36fd",
	"en/en_GB.dic": "04e90f34f5263bf26780e9c4a442e9ad16584e227af49ddd1b3b21b01df5b29c",
	"en/README_en_GB.txt": "18b5833de60a52ffcb6e5b6b5048bb5a5d910acfecd5b096985b77058223880c",
	"en/en_CA.aff": "e746c882dd6f303c2c46e7452804b9201115a6942cfeb15f18f8edf774d2e24e",
	"en/en_CA.dic": "53096689870246a78fa408b1e7a8c49bc140bd466a84a384cfb509445fff7173",
	"en/README_en_CA.txt": "86700e39b2307dda4c087ad3505a1e634a5b4d518ab808c398a16d21ebe957e2",
	"en/en_AU.aff": "e746c882dd6f303c2c46e7452804b9201115a6942cfeb15f18f8edf774d2e24e",
	"en/en_AU.dic": "aa07c46571f306b79fc1bc534357ed357af15687381b26f891ba66e8a2caed89",
	"en/README_en_AU.txt": "f0897dcfdf058bcfc93fbc8e7e1b5da4c27ae9e175ee3a9242e887f02eb84a8b",
	"de/de_DE_frami.aff": "646bf3333ac69c23e9d794533ee5241d6f755c359e8fe10a648f87613743d594",
	"de/de_DE_frami.dic": "4ca3c958b0e5545910999bc246f668840bf8ede3df8e5e6790d05edd5a586c38",
	"de/de_AT_frami.aff": "8b428f6215e36e145626e3d150cd0ffe273e1ad9a9ec0a1bbd31ee468ee39edc",
	"de/de_AT_frami.dic": "d821e0a5a029b66df696c982a4cd55eb792139c33716ccbde364941db2448cbb",
	"de/de_CH_frami.aff": "d86884be4c3f07e5e08057884856feabf2872b4e2f90a587a11376082bc4126b",
	"de/de_CH_frami.dic": "80a4df38cdc3e16d2f042f8eba4a99668ab9c075662e4262c62f8bba46c7bd16",
	"de/README_de_DE_frami.txt": "c141f4f79c428b7348b5012836f4ad3db4d124f288f15effc22696dc876512ae",
	"de/COPYING_GPLv2": "e7431ac20e815b6797369e802f9e4bb1b0082485060d8b096235f035d741bd89",
	"de/COPYING_GPLv3": "8ceb4b9ee5adedde47b31e975c1d90c73ad27b6b165a1dcd80c7c545eb65b903",
	"de/COPYING_LGPL_v2.1.txt": "dc626520dcd53a22f727af3ee42c770e56c97a64fe3adb063799d8ab032fe551",
	"fr_FR/dictionaries/fr.aff": "c176610cd5dc4846806a65ddd029f422d87978bf58f224aa44222662a16a2de5",
	"fr_FR/dictionaries/fr.dic": "b78a868e31dd6e373b6c3217969afb898a9acde828a5e7ef97308da42218c88c",
	"fr_FR/dictionaries/README_dict_fr.txt": "4a25abf740ad6ba5e8d0cfc0d2ffadc3c13c303f8fe389f323b061a389350f83",
	"es/es_ES.aff": "e73a9bf8e1383f4986a5dc9e2fbed49371c0c61f511c626d15586bd433c1cad9",
	"es/es_ES.dic": "6975dddec3d5d2c676069537bc67b4b5f786c65c5d4cf6703a82acf779ac9ec1",
	"es/README_hunspell_es.txt": "0b5c39a863ecb61f5c5631ec1cd977667ef9bda6da6f90e80e9fc0edcbfadd97",
	"es/LICENSE.md": "2b70fd5cb8d3b34cc6ae322c969557a2270b99017594addb2531ef5ffd9e1b17",
	"es/GPLv3.txt": "8ceb4b9ee5adedde47b31e975c1d90c73ad27b6b165a1dcd80c7c545eb65b903",
	"es/LGPLv3.txt": "a853c2ffec17057872340eee242ae4d96cbf2b520ae27d903e1b2fef1a5f9d1c",
	"es/MPL-1.1.txt": "53692a2ed6c6a2c6ec9b32dd0b820dfae91e0a1fcdf625ca9ed0bdf8705fcc4f",
	"it_IT/it_IT.aff": "951afaa19272f13555b8823e8bcf9ccf78f8fe1a07835bdfb912ab3e4d537c2b",
	"it_IT/it_IT.dic": "bae1e3501dcd2a923669592493b3fde6c02aae7c7aab83bf5e5b49077e73dd64",
	"it_IT/README_it_IT.txt": "34c3e93595cf3cf5f3afc9bc0d98eea12593750383f1e82ed1d3ca29f9681283",
	"it_IT/legacy/it_IT_license.txt": "4de624abc748506298e7be68f2213e422673b43e1674da4b376804096b87dcec",
	"it_IT/legacy/it_IT_AUTHORS": "67cc91ed0bff69cb9b387e5a28e37503a2a4afadc282b6a29cf08be7440b38ef",
	"pt_BR/pt_BR.aff": "21d8ad2a769a60e17e2b5ea4ef11d4d593a58b9e2a82d642ef82d6a4c5523865",
	"pt_BR/pt_BR.dic": "a38bfb26b68ece2834e79fe83e48d5792652970ace12db89d1b9674bf9933183",
	"pt_BR/README_pt_BR.txt": "9974ce691fdc1fe731717d7a2dc668244405fdc2bf9bf3367eb9b29e85177c88",
	"pt_BR/README_en.txt": "bb064f3db7ba5f5536da7d079773f692c5f928822e8aebc3d78259ff6cceb9e5",
	"pt_PT/pt_PT.aff": "975a209fcc892cb382fa5f34a28c391a39668661ce373ae071287809c5fcae24",
	"pt_PT/pt_PT.dic": "e29ba2d7aa8a2ad43e9cb46ac6473064b661545c87002aea90e18899d98d3cc9",
	"pt_PT/README_pt_PT.txt": "36de7d88a406a4947bf646a64145f00827566808988632e3df969aa95776060c",
	"pt_PT/LICENSES.txt": "d2c1cfe2e2dd81c651aec3fda5d1b4b4e7679b9e04f8cdc37586e521837384d1",
	"nl_NL/nl_NL.aff": "f0233d4f721f4661cf5f4d05ed2739549322bf3b6b66764b55a38257e1e16e6f",
	"nl_NL/nl_NL.dic": "bc28af45307700a9927ad5719184da44dfd7eed4f707b8c1477f6d8a21b586a6",
	"nl_NL/LICENSE.txt": "2189c248cbc054445aeac6cd65b87eebd0a115efeb55ed498b0f2a73441805dd",
	"pl_PL/pl_PL.aff": "82973651651aa930335c865b339b98db376ca3dbf3a661b70b9eeb71fdf41dca",
	"pl_PL/pl_PL.dic": "c0848440599eb88e5aca500418d5f389e562ec2c157b63dbe39d354658ffba49",
	"pl_PL/README_pl_PL.txt": "e691ed1e79934293b0cd6996a14af1520ff893c02f920c09ed0c90aa5f0a8abc",
	"ru_RU/ru_RU.aff": "38ce7d4af78e211e9bafe4bf7e3d6a2c420591136cb738ec6648f8fdf6524cd7",
	"ru_RU/ru_RU.dic": "f6047416a0204adbecf3a451b874ec8a97ee37e2cbc714466ef04d8dbcc0d6fc",
	"ru_RU/README_ru_RU.txt": "262af2f6ad70a61e5ee1332ff44fa8ee50edca819cf33207d8ad6ba6a0c9be52",
	"sv_SE/dictionaries/sv_SE.aff": "b721c9d44bee912feb182b601a1bc2ae3e7dffef660f4130cf2751867488a9dd",
	"sv_SE/dictionaries/sv_SE.dic": "384a2126eff333f5f6f9790ae892554546f53948d2988c600397cb5ad6ce66e8",
	"sv_SE/LICENSE_en_US.txt": "1314b6004bff32053ca25aca5fc419e905d75a6d49f9f9825baf4f02298af6d9",
	"sv_SE/LICENSE_sv_SE.txt": "841736f210169ce2c9c6907c0d88495e6d41ff7d067a3033c56ae1179ce62127",
	"da_DK/da_DK.aff": "ee4fe193d19b72aa64bc0d2b9bcdf44ee44f11d9b4e7dcf6da36fd0e91528561",
	"da_DK/da_DK.dic": "298aabeca362ee2c71836c5c7748d22080a460bd7193e6fbd2471f27a70d0fd3",
	"da_DK/README_da_DK.txt": "9508f4ec9e14a3d2a3360dd694881440d0d73fe1fe076e547639e5fe33c7b42d",
	"cs_CZ/cs_CZ.aff": "7ecb20620ecd46ebd9c36f3f33e69dd4eda385cba5b2bb4e6bc396d910e297f7",
	"cs_CZ/cs_CZ.dic": "d8e8c88c006fdae72dac8c85df11b0c99a773e05a4ab0fcbe92244876668ca74",
	"cs_CZ/README_en.txt": "0fe6d017aa91ffb58146d19160f8207900cc0c49d5fffef0b1a7d3a364cb29bd",
	"cs_CZ/README_cs.txt": "24d1d07409b62e8e6f0ee114991d4749d3e97b05ea19feca835916af67312720"
} as const

export type SourcePath = keyof typeof SOURCES

export interface DictionarySource {
	readonly id: string
	readonly name: string
	readonly aff: SourcePath
	readonly dic: SourcePath
	// Copied verbatim. Licence texts from another directory of the same commit fill in where the dictionary's own
	// directory names a licence without shipping its text.
	readonly notices: readonly SourcePath[]
	readonly license: string
	readonly licenseEvidence: string
	// Copyright holders as the shipped notices state them (contact addresses left out); where a notice names authors
	// without a copyright line, the authors with their role.
	readonly copyright: readonly string[]
	readonly upstreamVersion: string | null
	readonly versionSource: SourcePath | null
}

const SCOWL_EVIDENCE =
	"The README: word lists under the SCOWL and VarCon permission notices (Kevin Atkinson, Benjamin Titze), the affix file under Geoff Kuenning's BSD-style Ispell licence (modified versions must be marked), plus the WordNet and UKACD notices; UKACD requires its notice to ship verbatim, hence the README."

const SCOWL_COPYRIGHT = [
	"Copyright 2000-2018 by Kevin Atkinson",
	"Copyright 2000-2016 by Kevin Atkinson",
	"Copyright 2016 by Benjamin Titze",
	"Copyright 1993, Geoff Kuenning, Granada Hills, CA",
	"Copyright (c) J Ross Beresford 1993-1999. All Rights Reserved.",
	"WordNet 1.6 Copyright 1997 by Princeton University. All rights reserved."
]

export const DICTIONARIES: readonly DictionarySource[] = [
	{
		id: "en_US",
		name: "English (United States)",
		aff: "en/en_US.aff",
		dic: "en/en_US.dic",
		notices: ["en/README_en_US.txt"],
		license: "SCOWL",
		licenseEvidence: SCOWL_EVIDENCE,
		upstreamVersion: "2020.12.07",
		copyright: SCOWL_COPYRIGHT,
		versionSource: "en/README_en_US.txt"
	},
	{
		id: "en_GB",
		name: "English (United Kingdom)",
		aff: "en/en_GB.aff",
		dic: "en/en_GB.dic",
		notices: ["en/README_en_GB.txt", "de/COPYING_LGPL_v2.1.txt"],
		license: "LGPL-2.1-only",
		licenseEvidence:
			"README_en_GB.txt: word list and affix file under the LGPL without a version number, so any version may be chosen; the phonetic table in the .aff is LGPL 2.1 only, so the whole is used under LGPL 2.1. en/ ships no LGPL text: the LGPL 2.1 text comes from de/ of the same commit.",
		copyright: [
			"Kevin Atkinson (author of the original English wordlist)",
			"David Bartlett, Brian Kelk, Andrew Brown and Marco A.G.Pinto (authors of the British English dictionary)",
			"© 2000 Björn Jacke",
			"© OpenStreetMap contributors",
			"© The Clergy of the Church of England Database Project, 2005."
		],
		upstreamVersion: "4.0.1",
		versionSource: "en/en_GB.aff"
	},
	{
		id: "en_CA",
		name: "English (Canada)",
		aff: "en/en_CA.aff",
		dic: "en/en_CA.dic",
		notices: ["en/README_en_CA.txt"],
		license: "SCOWL",
		licenseEvidence: SCOWL_EVIDENCE,
		upstreamVersion: "2020.12.07",
		copyright: SCOWL_COPYRIGHT,
		versionSource: "en/README_en_CA.txt"
	},
	{
		id: "en_AU",
		name: "English (Australia)",
		aff: "en/en_AU.aff",
		dic: "en/en_AU.dic",
		notices: ["en/README_en_AU.txt"],
		license: "SCOWL",
		licenseEvidence: SCOWL_EVIDENCE,
		upstreamVersion: "2020.12.07",
		copyright: SCOWL_COPYRIGHT,
		versionSource: "en/README_en_AU.txt"
	},
	...(["de_DE", "de_AT", "de_CH"] as const).map((id): DictionarySource => ({
		id,
		name: { de_DE: "German (Germany)", de_AT: "German (Austria)", de_CH: "German (Switzerland)" }[id],
		aff: `de/${id}_frami.aff`,
		dic: `de/${id}_frami.dic`,
		notices: ["de/README_de_DE_frami.txt", "de/COPYING_GPLv2", "de/COPYING_GPLv3"],
		license: "GPL-2.0-only OR GPL-3.0-only",
		licenseEvidence:
			'README_de_DE_frami.txt (covers every xxx_frami dictionary): "GNU GPL, Version 2 oder 3"; the .aff header: "License: GPLv2, GPLv3", with a copy of both licences in every distribution.',
		copyright: ["Bjoern Jacke (author of the igerman98 base dictionary)", "Franz Michael Baumann (author of the frami extension)"],
		upstreamVersion: "20161207+frami20170109",
		versionSource: `de/${id}_frami.aff`
	})),
	{
		id: "fr",
		name: "French",
		aff: "fr_FR/dictionaries/fr.aff",
		dic: "fr_FR/dictionaries/fr.dic",
		notices: ["fr_FR/dictionaries/README_dict_fr.txt"],
		license: "MPL-2.0",
		licenseEvidence:
			"README_dict_fr.txt: Mozilla Public License 2.0; the .aff header carries the MPL 2.0 Exhibit A notice with the licence URL.",
		copyright: ["Olivier R. (author of the Grammalecte French dictionaries)"],
		upstreamVersion: "7.7",
		versionSource: "fr_FR/dictionaries/README_dict_fr.txt"
	},
	{
		id: "es_ES",
		name: "Spanish (Spain)",
		aff: "es/es_ES.aff",
		dic: "es/es_ES.dic",
		notices: ["es/README_hunspell_es.txt", "es/LICENSE.md", "es/GPLv3.txt", "es/LGPLv3.txt", "es/MPL-1.1.txt"],
		license: "GPL-3.0-or-later OR LGPL-3.0-or-later OR MPL-1.1+",
		licenseEvidence:
			"README_hunspell_es.txt and LICENSE.md: GNU GPL 3 or later, GNU LGPL 3 or later or MPL 1.1 or later, at the user's choice; LICENSE.md names the three licence texts.",
		copyright: ["Santiago Bosio (author of the RLA-ES Spanish dictionary)"],
		upstreamVersion: null,
		versionSource: null
	},
	{
		id: "it_IT",
		name: "Italian (Italy)",
		aff: "it_IT/it_IT.aff",
		dic: "it_IT/it_IT.dic",
		notices: ["it_IT/README_it_IT.txt", "it_IT/legacy/it_IT_license.txt", "it_IT/legacy/it_IT_AUTHORS"],
		license: "GPL-3.0-only",
		licenseEvidence:
			'README_it_IT.txt and the .aff header: GNU GPL version 3 (the licence text closes README_it_IT.txt); the copyright notice and authors of the original "Dizionario italiano" 2.4 come from legacy/.',
		copyright: [
			"Copyright (C) 2020-2022, LibreItalia - Marina Latini",
			"Copyright (C) 2001, 2002 Gianluca Turconi",
			"Copyright (C) 2002, 2003, 2004 Gianluca Turconi and Davide Prina",
			"Copyright (C) 2004, 2005, 2006, 2007 Davide Prina",
			"Andrea Pescetti (previous maintainer of the Italian Writing Aids extension)"
		],
		upstreamVersion: "5.1.1",
		versionSource: "it_IT/README_it_IT.txt"
	},
	{
		id: "pt_BR",
		name: "Portuguese (Brazil)",
		aff: "pt_BR/pt_BR.aff",
		dic: "pt_BR/pt_BR.dic",
		notices: ["pt_BR/README_pt_BR.txt", "pt_BR/README_en.txt", "es/LGPLv3.txt", "es/GPLv3.txt"],
		license: "LGPL-3.0-only",
		licenseEvidence:
			"README_pt_BR.txt, README_en.txt and the .aff header: GNU LGPL version 3 and the Mozilla Public License (no MPL version stated). pt_BR/ ships no licence text: LGPL 3 and the GPL 3 it builds on come from es/ of the same commit.",
		copyright: ["Copyright (C) 2006 - 2013 by Raimundo Santos Moura"],
		upstreamVersion: "3.2",
		versionSource: "pt_BR/pt_BR.aff"
	},
	{
		id: "pt_PT",
		name: "Portuguese (Portugal)",
		aff: "pt_PT/pt_PT.aff",
		dic: "pt_PT/pt_PT.dic",
		notices: ["pt_PT/README_pt_PT.txt", "pt_PT/LICENSES.txt", "de/COPYING_GPLv2", "de/COPYING_LGPL_v2.1.txt", "es/MPL-1.1.txt"],
		license: "GPL-2.0-only OR LGPL-2.1-only OR MPL-1.1",
		licenseEvidence:
			'README_pt_PT.txt: "GPL/LGPL/MPL", versions GPL 2, LGPL 2.1 and MPL 1.1 (LICENSES.txt summarises the spellchecker as "GPL and BSD" without detail). The licence texts come from de/ and es/ of the same commit.',
		copyright: ["Copyright (C) 2006-2012 José João de Almeida, Rui Vilela, Alberto Simões"],
		upstreamVersion: null,
		versionSource: null
	},
	{
		id: "nl_NL",
		name: "Dutch (Netherlands)",
		aff: "nl_NL/nl_NL.aff",
		dic: "nl_NL/nl_NL.dic",
		notices: ["nl_NL/LICENSE.txt"],
		license: "BSD-3-Clause OR CC-BY-3.0",
		licenseEvidence:
			'LICENSE.txt and the .aff header: "Revised BSD License and/or CC BY 3.0", at the user\'s discretion; LICENSE.txt asks to be kept with every copy and contains both licence texts.',
		copyright: [
			"© 2020 OpenTaal (Simon Brouwer, Sander van Geloven)",
			"© 2006 - 2011 OpenTaal (Ruud Baars, Simon Brouwer)",
			"© 2001 - 2005 Simon Brouwer and others",
			"© 1996 Nederlandstalige TeX Gebruikersgroep"
		],
		upstreamVersion: "2.20.21",
		versionSource: "nl_NL/nl_NL.aff"
	},
	{
		id: "pl_PL",
		name: "Polish (Poland)",
		aff: "pl_PL/pl_PL.aff",
		dic: "pl_PL/pl_PL.dic",
		notices: ["pl_PL/README_pl_PL.txt"],
		license: "GPL-2.0-only OR LGPL-2.1-only OR MPL-1.1 OR Apache-2.0 OR CC-BY-4.0",
		licenseEvidence: "README_pl_PL.txt: GPL 2, LGPL 2.1, MPL 1.1, Apache 2.0 and Creative Commons Attribution 4.0 International.",
		upstreamVersion: "2026-05-11",
		copyright: ["sjp.pl (authors of the Polish dictionary)"],
		versionSource: "pl_PL/README_pl_PL.txt"
	},
	{
		id: "ru_RU",
		name: "Russian (Russia)",
		aff: "ru_RU/ru_RU.aff",
		dic: "ru_RU/ru_RU.dic",
		notices: ["ru_RU/README_ru_RU.txt"],
		license: "BSD-style (Lebedev)",
		licenseEvidence:
			'README_ru_RU.txt: BSD-style licence by Alexander I. Lebedev with the extra condition "Modified versions must be clearly marked as such".',
		copyright: ["Copyright (c) 1997-2008, Alexander I. Lebedev"],
		upstreamVersion: null,
		versionSource: null
	},
	{
		id: "sv_SE",
		name: "Swedish (Sweden)",
		aff: "sv_SE/dictionaries/sv_SE.aff",
		dic: "sv_SE/dictionaries/sv_SE.dic",
		notices: ["sv_SE/LICENSE_en_US.txt", "sv_SE/LICENSE_sv_SE.txt", "es/LGPLv3.txt", "es/GPLv3.txt"],
		license: "LGPL-3.0-only",
		licenseEvidence:
			"LICENSE_en_US.txt: GNU Lesser General Public License Version 3. sv_SE/ ships no licence text: LGPL 3 and the GPL 3 it builds on come from es/ of the same commit.",
		copyright: ["Copyright © 2003-19 Göran Andersson"],
		upstreamVersion: null,
		versionSource: null
	},
	{
		id: "da_DK",
		name: "Danish (Denmark)",
		aff: "da_DK/da_DK.aff",
		dic: "da_DK/da_DK.dic",
		notices: ["da_DK/README_da_DK.txt", "de/COPYING_GPLv2", "de/COPYING_LGPL_v2.1.txt", "es/MPL-1.1.txt"],
		license: "GPL-2.0-only OR LGPL-2.1-only OR MPL-1.1",
		licenseEvidence:
			"README_da_DK.txt and the .aff header: GPL 2.0, LGPL 2.1 or MPL 1.1 (tri-license). The licence texts come from de/ and es/ of the same commit.",
		copyright: ["© 2020 Foreningen for frit tilgængelige sprogværktøjer"],
		upstreamVersion: "2.9.101",
		versionSource: "da_DK/README_da_DK.txt"
	},
	{
		id: "cs_CZ",
		name: "Czech (Czechia)",
		aff: "cs_CZ/cs_CZ.aff",
		dic: "cs_CZ/cs_CZ.dic",
		notices: ["cs_CZ/README_en.txt", "cs_CZ/README_cs.txt"],
		license: "GPL-1.0-or-later",
		licenseEvidence:
			'README_en.txt and README_cs.txt: "licensed under the GNU/GPL license" without a version number, so any version may be chosen; README_cs.txt carries the GPL 2 text.',
		copyright: [
			"Petr Kolář and numerous contributors (authors of the original Czech ispell dictionary)",
			"Miroslav Pošta (author of the 2021 update of the Czech dictionary)"
		],
		upstreamVersion: null,
		versionSource: null
	}
]

// .dic lines a strict parser rejects: an unescaped "/" inside the word makes the rest a flag field that is not valid for
// the dictionary's FLAG type. Found by validating every line; the build fails when the pinned input yields another set.
export interface DicLineFix {
	// Each upstream line and its replacement.
	readonly lines: readonly (readonly [string, string])[]
	readonly reason: string
}

export const DIC_LINE_FIXES: Readonly<Record<string, DicLineFix>> = {
	da_DK: {
		lines: [
			['"A/S"', "A\\/S"],
			['"c/o"', "c\\/o"],
			['"I/S"', "I\\/S"]
		],
		reason: 'the unescaped "/" is read as a flag separator, which FLAG num rejects, and the quotes would keep the entries from ever matching a word'
	}
}

interface Encoding {
	readonly name: string
	readonly label: string
	// WHATWG decodes these ISO names as their Windows superset, which matches only while no byte in 0x80-0x9F occurs.
	readonly c1?: true
}

// Hunspell SET names (lower-cased) to WHATWG labels.
const ENCODINGS: Readonly<Record<string, Encoding>> = {
	"utf-8": { name: "UTF-8", label: "utf-8" },
	"iso8859-1": { name: "ISO-8859-1", label: "windows-1252", c1: true },
	"iso8859-2": { name: "ISO-8859-2", label: "iso-8859-2" },
	"iso8859-3": { name: "ISO-8859-3", label: "iso-8859-3" },
	"iso8859-4": { name: "ISO-8859-4", label: "iso-8859-4" },
	"iso8859-5": { name: "ISO-8859-5", label: "iso-8859-5" },
	"iso8859-6": { name: "ISO-8859-6", label: "iso-8859-6" },
	"iso8859-7": { name: "ISO-8859-7", label: "iso-8859-7" },
	"iso8859-8": { name: "ISO-8859-8", label: "iso-8859-8" },
	"iso8859-9": { name: "ISO-8859-9", label: "windows-1254", c1: true },
	"iso8859-10": { name: "ISO-8859-10", label: "iso-8859-10" },
	"iso8859-13": { name: "ISO-8859-13", label: "iso-8859-13" },
	"iso8859-14": { name: "ISO-8859-14", label: "iso-8859-14" },
	"iso8859-15": { name: "ISO-8859-15", label: "iso-8859-15" },
	"koi8-r": { name: "KOI8-R", label: "koi8-r" },
	"koi8-u": { name: "KOI8-U", label: "koi8-u" },
	"microsoft-cp1251": { name: "windows-1251", label: "windows-1251" },
	"windows-1250": { name: "windows-1250", label: "windows-1250" },
	"windows-1251": { name: "windows-1251", label: "windows-1251" },
	"tis620-2533": { name: "TIS-620", label: "windows-874", c1: true }
}

export const MARK = "# Modified for Filen by packages/filen-spellcheck/scripts/build-dictionaries.ts (filen-ts repository)."

export const CACHE = fileURLToPath(new URL(`../node_modules/.cache/filen-spellcheck-dictionaries/${UPSTREAM.commit}/`, import.meta.url))
export const DICTIONARIES_DIR = fileURLToPath(new URL("../dictionaries/", import.meta.url))
export const DICTIONARIES_OUT = fileURLToPath(new URL("../.dictionaries/", import.meta.url))
// In an output directory: the SHA-256 of the manifest.json its files were verified against.
export const STAMP = ".complete"

export interface ManifestFile {
	readonly file: string
	readonly bytes: number
	readonly sha256: string
}

export interface ManifestSource {
	readonly url: string
	readonly sha256: string
}

export interface ManifestDictionary {
	readonly id: string
	readonly name: string
	readonly upstreamVersion: string | null
	readonly license: string
	readonly copyright: readonly string[]
	readonly sources: readonly ManifestSource[]
	// Relative to the output directory (.dictionaries/).
	readonly aff: ManifestFile
	readonly dic: ManifestFile
	// Relative to dictionaries/.
	readonly notices: readonly ManifestFile[]
	readonly modifications: readonly string[]
}

export interface Manifest {
	readonly version: 1
	readonly transform: number
	readonly upstream: { readonly repository: string; readonly commit: string }
	readonly dictionaries: readonly ManifestDictionary[]
}

export function sourceUrl(path: SourcePath): string {
	return `${RAW}${path}`
}

function sha256(bytes: Uint8Array): string {
	return createHash("sha256").update(bytes).digest("hex")
}

function cachePath(path: SourcePath): string {
	return join(CACHE, path)
}

// A pinned source from the cache, verified. Throws when it is missing or does not match its pin.
function readSource(path: SourcePath): Uint8Array {
	const file = cachePath(path)

	if (!existsSync(file)) throw new Error(`${path} is not cached; run without --offline`)

	const bytes = new Uint8Array(readFileSync(file))

	if (sha256(bytes) !== SOURCES[path]) throw new Error(`${path} does not match its pinned SHA-256`)

	return bytes
}

async function download(path: SourcePath): Promise<void> {
	const file = cachePath(path)

	if (existsSync(file) && sha256(new Uint8Array(readFileSync(file))) === SOURCES[path]) return

	const response = await fetch(sourceUrl(path))

	if (!response.ok) throw new Error(`${path}: HTTP ${response.status}`)

	const bytes = new Uint8Array(await response.arrayBuffer())

	if (sha256(bytes) !== SOURCES[path]) throw new Error(`${path} does not match its pinned SHA-256`)

	mkdirSync(dirname(file), { recursive: true })
	writeFileSync(file, bytes)
}

async function downloadAll(): Promise<void> {
	const queue = Object.keys(SOURCES) as SourcePath[]

	await Promise.all(
		Array.from({ length: 6 }, async () => {
			for (let path = queue.shift(); path !== undefined; path = queue.shift()) await download(path)
		})
	)
}

const BOM = [0xef, 0xbb, 0xbf] as const

function hasBom(bytes: Uint8Array): boolean {
	return bytes[0] === BOM[0] && bytes[1] === BOM[1] && bytes[2] === BOM[2]
}

// The encoding the .aff's SET line declares; the line is ASCII, so it is read before decoding.
function declaredEncoding(path: SourcePath, bytes: Uint8Array): Encoding {
	const text = new TextDecoder("windows-1252").decode(hasBom(bytes) ? bytes.subarray(3) : bytes)
	const sets = [...text.matchAll(/^SET[ \t]+(\S+)[ \t]*\r?$/gm)]
	const name = sets[0]?.[1]

	if (sets.length !== 1 || name === undefined) throw new Error(`${path}: expected exactly one SET line, found ${sets.length}`)

	const encoding = ENCODINGS[name.toLowerCase()]

	if (encoding === undefined) throw new Error(`${path}: unknown encoding ${name}`)

	return encoding
}

interface Decoded {
	readonly text: string
	readonly bom: boolean
	readonly lineEndings: boolean
	readonly finalNewline: boolean
}

// UTF-8 text with LF line endings and a final newline, plus which of those changed the input.
function decode(path: SourcePath, bytes: Uint8Array, encoding: Encoding): Decoded {
	if (encoding.c1 === true && bytes.some(byte => byte >= 0x80 && byte <= 0x9f)) {
		throw new Error(`${path}: bytes 0x80-0x9F are ambiguous in ${encoding.name}`)
	}

	const bom = encoding.label === "utf-8" && hasBom(bytes)
	let text = new TextDecoder(encoding.label, { fatal: true, ignoreBOM: true }).decode(bom ? bytes.subarray(3) : bytes)
	const lineEndings = text.includes("\r")

	if (lineEndings) text = text.replace(/\r\n?/g, "\n")

	const finalNewline = !text.endsWith("\n")

	if (finalNewline) text += "\n"

	return { text, bom, lineEndings, finalNewline }
}

type FlagType = "short" | "long" | "num" | "UTF-8"

function flagType(path: SourcePath, aff: string): FlagType {
	if (/^AF[ \t]/m.test(aff)) throw new Error(`${path}: flag aliases (AF) are not supported by the .dic validation`)

	const value = /^FLAG[ \t]+(\S+)/m.exec(aff)?.[1]

	if (value === undefined) return "short"
	if (value === "long" || value === "num" || value === "UTF-8") return value

	throw new Error(`${path}: unknown FLAG type ${value}`)
}

// The flag field of a .dic entry: after the first "/" that is neither escaped ("\/") nor the entry's first character, up to
// the morphological fields (a tab, or whitespace before an "xx:" field), as Hunspell splits it.
function flagField(line: string): string | undefined {
	let end = line.indexOf("\t")

	if (end < 0) end = line.length

	const morph = /[ \t]+\S\S:/.exec(line.slice(0, end))
	const entry = line.slice(0, morph === null ? end : morph.index)
	let slash = entry.indexOf("/", 1)

	while (slash > 0 && entry[slash - 1] === "\\") slash = entry.indexOf("/", slash + 1)

	return slash < 0 ? undefined : entry.slice(slash + 1).trimEnd()
}

function validFlags(field: string, type: FlagType): boolean {
	switch (type) {
		case "num":
			return /^\d+(?:,\d+)*$/.test(field) && field.split(",").every(flag => Number(flag) <= 65535)
		case "long":
			return [...field].length % 2 === 0
		default:
			return true
	}
}

function invalidDicLines(dic: string, type: FlagType): string[] {
	const invalid: string[] = []
	const lines = dic.split("\n")

	for (let index = 1; index < lines.length; index++) {
		const line = lines[index] ?? ""

		if (!line.includes("/")) continue

		const field = flagField(line)

		if (field !== undefined && !validFlags(field, type)) invalid.push(line)
	}

	return invalid
}

const FLAG_DIRECTIVES = new Set([
	"CIRCUMFIX",
	"COMPOUNDBEGIN",
	"COMPOUNDEND",
	"COMPOUNDFLAG",
	"COMPOUNDFORBIDFLAG",
	"COMPOUNDLAST",
	"COMPOUNDMIDDLE",
	"COMPOUNDPERMITFLAG",
	"COMPOUNDROOT",
	"COMPOUNDRULE",
	"FORBIDDENWORD",
	"FORCEUCASE",
	"KEEPCASE",
	"LEMMA_PRESENT",
	"NEEDAFFIX",
	"NOSUGGEST",
	"ONLYINCOMPOUND",
	"PSEUDOROOT",
	"SUBSTANDARD",
	"WARN"
])

const ASCII = /^\p{ASCII}*$/u

// Short flags are single bytes, so transcoding preserves them only when every flag is ASCII.
function assertAsciiFlags(id: string, aff: string, dic: string): void {
	for (const line of aff.split("\n")) {
		const fields = line.trim().split(/[ \t]+/)
		const directive = fields[0] ?? ""
		// PFX/SFX: the flag, then continuation flags after "/" in the affix field.
		const flags =
			directive === "PFX" || directive === "SFX"
				? `${fields[1] ?? ""}${fields[3]?.split("/")[1] ?? ""}`
				: FLAG_DIRECTIVES.has(directive)
					? (fields[1] ?? "")
					: ""

		if (!ASCII.test(flags)) throw new Error(`${id}: non-ASCII flag in .aff line ${JSON.stringify(line)}`)
	}

	for (const line of dic.split("\n").slice(1)) {
		const field = flagField(line)

		if (field !== undefined && !ASCII.test(field)) throw new Error(`${id}: non-ASCII flag in .dic line ${JSON.stringify(line)}`)
	}
}

function fixDic(source: DictionarySource, dic: string, type: FlagType): { dic: string; modification: string | null } {
	const fix = DIC_LINE_FIXES[source.id]
	const expected = fix?.lines.map(([line]) => line) ?? []
	const found = invalidDicLines(dic, type)

	if (found.length !== expected.length || found.some((line, index) => line !== expected[index])) {
		throw new Error(`${source.id}: invalid .dic lines changed; expected ${JSON.stringify(expected)}, found ${JSON.stringify(found)}`)
	}

	if (fix === undefined) return { dic, modification: null }

	const replacements = new Map(fix.lines)
	const fixed = dic
		.split("\n")
		.map((line, index) => (index === 0 ? line : (replacements.get(line) ?? line)))
		.join("\n")

	if (invalidDicLines(fixed, type).length !== 0) throw new Error(`${source.id}: .dic line fixes did not take`)

	const list = (index: 0 | 1): string => fix.lines.map(pair => `\`${pair[index]}\``).join(", ")

	return { dic: fixed, modification: `Rewrote the .dic lines ${list(0)} as ${list(1)}: ${fix.reason}` }
}

interface Built {
	readonly aff: Uint8Array
	readonly dic: Uint8Array
	readonly modifications: readonly string[]
}

function buildDictionary(source: DictionarySource): Built {
	const affBytes = readSource(source.aff)
	const encoding = declaredEncoding(source.aff, affBytes)
	const aff = decode(source.aff, affBytes, encoding)
	const dic = decode(source.dic, readSource(source.dic), encoding)
	const type = flagType(source.aff, aff.text)

	if (encoding.label !== "utf-8" && type === "short") assertAsciiFlags(source.id, aff.text, dic.text)

	const modifications: string[] = []
	const changed = (what: string, inAff: boolean, inDic: boolean): void => {
		if (inAff || inDic) modifications.push(`${what} ${[inAff ? ".aff" : "", inDic ? ".dic" : ""].filter(Boolean).join(" and ")}`)
	}
	let affText = aff.text

	if (encoding.label !== "utf-8") {
		affText = affText.replace(/^SET[ \t]+\S+[ \t]*$/m, "SET UTF-8")
		modifications.push(`Converted from ${encoding.name} to UTF-8 (SET line rewritten to "SET UTF-8")`)
	}

	changed("Removed the UTF-8 byte order mark from the", aff.bom, dic.bom)
	changed("Normalized line endings to LF in the", aff.lineEndings, dic.lineEndings)
	changed("Added a final newline to the", aff.finalNewline, dic.finalNewline)

	const fixed = fixDic(source, dic.text, type)

	if (fixed.modification !== null) modifications.push(fixed.modification)

	modifications.push("Prepended this modification notice to the .aff as comment lines")

	const notice = [
		MARK,
		`# Upstream: LibreOffice dictionaries ${UPSTREAM.commit} ${source.aff}, ${source.dic}`,
		...modifications.map(line => `# - ${line}`)
	]
	const encoder = new TextEncoder()

	return {
		aff: encoder.encode(`${notice.join("\n")}\n${affText}`),
		dic: encoder.encode(fixed.dic),
		modifications
	}
}

function basename(path: SourcePath): string {
	return posix.basename(path)
}

function sourcesOf(source: DictionarySource): ManifestSource[] {
	return [source.aff, source.dic, ...source.notices].map(path => ({ url: sourceUrl(path), sha256: SOURCES[path] }))
}

function dictionaryReadme(source: DictionarySource, modifications: readonly string[]): string {
	const version =
		source.upstreamVersion === null || source.versionSource === null
			? "not stated upstream"
			: `${source.upstreamVersion} (as stated in \`${source.versionSource}\`)`

	return [
		`# ${source.name} (${source.id})`,
		"",
		`Hunspell spelling dictionary for Filen's on-device spellchecker, built from the [LibreOffice dictionaries](${UPSTREAM.repository}) repository.`,
		"",
		`- Upstream commit: \`${UPSTREAM.commit}\``,
		`- Affix file: \`${source.aff}\` → \`${source.id}.aff\``,
		`- Word list: \`${source.dic}\` → \`${source.id}.dic\``,
		`- Upstream version: ${version}`,
		`- Licence: \`${source.license}\``,
		"",
		source.licenseEvidence,
		"",
		"Copyright holders, as the upstream notices state them:",
		"",
		...source.copyright.map(line => `- ${line}`),
		"",
		"## Modifications",
		"",
		...modifications.map(line => `- ${line}`),
		"",
		"## Upstream notices",
		"",
		"Copied verbatim from the same commit:",
		"",
		...source.notices.map(path => `- \`${basename(path)}\` from \`${path}\``),
		"",
		`\`packages/filen-spellcheck/scripts/build-dictionaries.ts\` builds \`${source.id}.aff\` and \`${source.id}.dic\` into \`packages/filen-spellcheck/.dictionaries/${source.id}/\` (not committed) and verifies them byte for byte against \`manifest.json\`.`,
		""
	].join("\n")
}

function directoryReadme(): string {
	return [
		"# Dictionaries",
		"",
		`Hunspell dictionaries for Filen's on-device spellchecker, built from the [LibreOffice dictionaries](${UPSTREAM.repository}) repository at commit \`${UPSTREAM.commit}\`.`,
		"",
		"Committed here: `manifest.json` (every source URL and SHA-256, and the size and SHA-256 every built file must have) and per dictionary the upstream licence notices (verbatim) with a README listing the source, licence, copyright holders and every modification.",
		"",
		"The built `.aff` and `.dic` files are not committed. `pnpm run build:dictionaries` downloads the pinned sources (cached under `node_modules/.cache`), builds them into `../.dictionaries/<id>/` and verifies every file against the manifest; it returns at once when `.dictionaries/.complete` already matches the manifest. `--offline` uses the download cache only.",
		"",
		"Changing the sources or the transformation in `scripts/build-dictionaries.ts` (bump `TRANSFORM` when the output changes) needs `pnpm run build:dictionaries --update`, which rewrites this directory. `src/tests/dictionaries.test.ts` checks the manifest and the committed notices; `FILEN_SPELLCHECK_DICTIONARIES=1 pnpm test` also rebuilds everything from scratch and compares it to the manifest.",
		"",
		"| Id | Language | Licence | Source |",
		"| --- | --- | --- | --- |",
		...DICTIONARIES.map(
			source => `| \`${source.id}\` | ${source.name} | \`${source.license}\` | \`${source.aff}\`, \`${source.dic}\` |`
		),
		""
	].join("\n")
}

function write(dir: string, file: string, bytes: Uint8Array): ManifestFile {
	const path = join(dir, file)

	mkdirSync(dirname(path), { recursive: true })
	writeFileSync(path, bytes)

	return { file, bytes: bytes.length, sha256: sha256(bytes) }
}

function fileEntry(file: string, bytes: Uint8Array): ManifestFile {
	return { file, bytes: bytes.length, sha256: sha256(bytes) }
}

function manifestBytes(manifest: Manifest): Uint8Array {
	return new TextEncoder().encode(`${JSON.stringify(manifest, null, "\t")}\n`)
}

// Builds every dictionary into outDir/<id>/ (one at a time, so only one is held in memory) and returns the manifest
// describing the result. `notices` receives each dictionary's notice files, README.md first.
async function generate(
	outDir: string,
	offline: boolean,
	notices?: (id: string, files: readonly [string, Uint8Array][]) => void
): Promise<Manifest> {
	if (!offline) await downloadAll()

	const encoder = new TextEncoder()
	const dictionaries: ManifestDictionary[] = []

	for (const source of DICTIONARIES) {
		const built = buildDictionary(source)
		const files: [string, Uint8Array][] = [["README.md", encoder.encode(dictionaryReadme(source, built.modifications))]]
		const names = new Set([`${source.id}.aff`, `${source.id}.dic`, "README.md"])

		for (const path of source.notices) {
			const name = basename(path)

			if (names.has(name)) throw new Error(`${source.id}: notice ${path} collides with ${name}`)

			names.add(name)
			files.push([name, readSource(path)])
		}

		rmSync(join(outDir, source.id), { recursive: true, force: true })

		const aff = write(outDir, `${source.id}/${source.id}.aff`, built.aff)
		const dic = write(outDir, `${source.id}/${source.id}.dic`, built.dic)

		notices?.(source.id, files)
		dictionaries.push({
			id: source.id,
			name: source.name,
			upstreamVersion: source.upstreamVersion,
			license: source.license,
			copyright: source.copyright,
			sources: sourcesOf(source),
			aff,
			dic,
			notices: files.map(([name, bytes]) => fileEntry(`${source.id}/${name}`, bytes)),
			modifications: built.modifications
		})
	}

	return { version: 1, transform: TRANSFORM, upstream: { repository: UPSTREAM.repository, commit: UPSTREAM.commit }, dictionaries }
}

const STALE = "dictionaries/manifest.json is stale; run `pnpm run build:dictionaries --update`"

// The committed manifest must describe what the pinned table builds; checked before any download or build.
function assertCurrent(manifest: Manifest): void {
	const expected = DICTIONARIES.map(source => [
		source.id,
		source.name,
		source.upstreamVersion,
		source.license,
		source.copyright,
		sourcesOf(source)
	])
	const actual = manifest.dictionaries.map(entry => [
		entry.id,
		entry.name,
		entry.upstreamVersion,
		entry.license,
		entry.copyright,
		entry.sources
	])
	const same =
		manifest.version === 1 &&
		manifest.transform === TRANSFORM &&
		manifest.upstream.repository === UPSTREAM.repository &&
		manifest.upstream.commit === UPSTREAM.commit &&
		JSON.stringify(actual) === JSON.stringify(expected)

	if (!same) throw new Error(`${STALE} (sources, transform or metadata differ from scripts/build-dictionaries.ts)`)
}

function readStamp(outDir: string): string | undefined {
	const path = join(outDir, STAMP)

	return existsSync(path) ? readFileSync(path, "utf8") : undefined
}

// Builds the dictionaries into outDir unless its stamp already matches the committed manifest, and verifies every built
// file against the manifest. Resolves to whether it had to build.
export async function ensureDictionaries(options: { readonly offline: boolean; readonly outDir?: string }): Promise<boolean> {
	const outDir = options.outDir ?? DICTIONARIES_OUT
	const committedBytes = new Uint8Array(readFileSync(join(DICTIONARIES_DIR, "manifest.json")))
	const committed = JSON.parse(new TextDecoder().decode(committedBytes)) as Manifest
	const hash = sha256(committedBytes)

	assertCurrent(committed)

	if (readStamp(outDir) === hash) return false

	rmSync(join(outDir, STAMP), { force: true })

	const built = await generate(outDir, options.offline)

	for (const [index, entry] of built.dictionaries.entries()) {
		const expected = committed.dictionaries[index]

		for (const key of ["aff", "dic"] as const) {
			const file = entry[key]
			const want = expected?.[key]

			if (want === undefined || file.bytes !== want.bytes || file.sha256 !== want.sha256) {
				throw new Error(
					`.dictionaries/${file.file} is not reproducible: ${file.bytes} bytes, SHA-256 ${file.sha256}; the manifest expects ${want?.bytes} bytes, SHA-256 ${want?.sha256}`
				)
			}
		}
	}

	if (sha256(manifestBytes(built)) !== hash) throw new Error(`${STALE} (modifications or notices differ)`)

	writeFileSync(join(outDir, STAMP), hash)

	return true
}

// Maintainers: rebuilds .dictionaries/ and rewrites dictionaries/ (manifest, notices, READMEs) from the pinned table.
export async function updateDictionaries(options: { readonly offline: boolean }): Promise<Manifest> {
	rmSync(join(DICTIONARIES_OUT, STAMP), { force: true })

	const manifest = await generate(DICTIONARIES_OUT, options.offline, (id, files) => {
		rmSync(join(DICTIONARIES_DIR, id), { recursive: true, force: true })

		for (const [name, bytes] of files) write(DICTIONARIES_DIR, `${id}/${name}`, bytes)
	})
	const bytes = manifestBytes(manifest)

	write(DICTIONARIES_DIR, "manifest.json", bytes)
	write(DICTIONARIES_DIR, "README.md", new TextEncoder().encode(directoryReadme()))
	writeFileSync(join(DICTIONARIES_OUT, STAMP), sha256(bytes))

	return manifest
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
	const args = process.argv.slice(2)
	const unknown = args.find(arg => arg !== "--offline" && arg !== "--update")

	if (unknown !== undefined) throw new Error(`Unknown argument ${unknown}; usage: build-dictionaries.ts [--update] [--offline]`)

	const offline = args.includes("--offline")

	if (args.includes("--update")) {
		const manifest = await updateDictionaries({ offline })

		console.log(`Rewrote dictionaries/ and built ${manifest.dictionaries.length} dictionaries into ${DICTIONARIES_OUT}`)
	} else if (await ensureDictionaries({ offline })) {
		console.log(`Built and verified ${DICTIONARIES.length} dictionaries in ${DICTIONARIES_OUT}`)
	}
}
