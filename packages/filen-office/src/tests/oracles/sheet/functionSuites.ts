import { createHash } from "node:crypto"
import { existsSync, mkdirSync, readdirSync, readFileSync, renameSync, rmSync, statSync, writeFileSync } from "node:fs"
import { dirname, join, relative } from "node:path"
import { fileURLToPath } from "node:url"

// Third-party spreadsheet function test suites, fetched on request into node_modules/.cache (never committed)
// and checked against pins here; the cases are generated from them at test time (functionCases.ts).
//
//   pnpm exec tsx src/tests/oracles/sheet/functionSuites.ts --fetch [libreoffice] [formulajs] [poi]
//
// - libreoffice: LibreOffice's sc/qa/unit/data/functions/**/*.fods at the commit of the 26.8.1.1 release,
//   MPL-2.0 (https://www.mozilla.org/MPL/2.0/). The directory is listed through GitHub's API at its git tree
//   id; each file is checked against its git blob id, and the whole set against SET_SHA256.
// - formulajs: formulajs/formulajs 4.6.1 test/*.js, MIT (Copyright (c) 2014 Sutoiku, Inc.); each file pinned.
// - poi: Apache POI 5.5.1 test-data/spreadsheet function workbooks, Apache-2.0
//   (https://www.apache.org/licenses/LICENSE-2.0); each file pinned.
//
// FILEN_OFFICE_SHEET_SUITES names another cache directory.

export const SUITES_DIR =
	process.env["FILEN_OFFICE_SHEET_SUITES"] ||
	fileURLToPath(new URL("../../../../node_modules/.cache/filen-office-sheet-functions", import.meta.url))

export type SuiteName = "libreoffice" | "formulajs" | "poi"

export const LIBREOFFICE = {
	commit: "bf6ef2662401c7d6dce671ab4deb26ee0af3a7c1",
	directory: "sc/qa/unit/data/functions",
	tree: "563220ac167b17071faa71f59a528b72428fb38b",
	files: 507,
	// SHA-256 over the sorted lines "<path>\0<SHA-256 of the file>\n".
	setSha256: "4c4ced1b093887482002d438f81c28b5ee34d670cd96e2de3a02dfc434b8441e"
} as const

export const FORMULAJS = {
	commit: "af0f0b41ed9786b795909b93ace268b139cbdb08",
	files: {
		"test/database.js": "d702515a6e952f9dd3cf98fe628613b98452950181dc287231dfc393204d22cb",
		"test/date-time.js": "a36f77ed6def93d489930a97ee5f1222dc66a7d4ccab4d784981dcda843babf1",
		"test/engineering.js": "9846a469a614f8bbd16527395ab46e981cb10447e5bb4740e6e3df9bb8510059",
		"test/financial.js": "84c581f422e99ddaf5dd3d3869bcda2b3db90eff18dfc7abf797e39e21019e7a",
		"test/information.js": "dead9e393d9567d837d32f307a4fb7a3832f7fc9134840cbe68c3bf165165d19",
		"test/logical.js": "8a813d0da873bfedf99c03445b751c2bada4c5f26d3f50f2e3b5291717552730",
		"test/lookup-reference.js": "e00ca01803f84145ad5445601d08457fb8709c433cf6224dea7dce1b2f71fbcd",
		"test/math-trig.js": "6ee9df96c25d955774667144b05e80739feb7e37edde72b00d374a1556061803",
		"test/statistical.js": "4bb7738f7d8ce0aad209aab147312521d73aac2369c85b6018123ea95507b4a8",
		"test/text.js": "7f2a02b0d61b023ddcbb5c97057c904942c235ec7359816a82524e98e86efea0"
	}
} as const

// The workbooks POI's BaseTestFunctionsFromSpreadsheet subclasses read (one function group per sheet).
export const POI = {
	commit: "094968cfc3d48224db08f0b7f0a6fc341b035114",
	directory: "test-data/spreadsheet",
	files: {
		"57003-FixedFunctionTestCaseData.xls": "9b9a2f462e918bd9f89d98d46176b9c59b8a58973cef1c32f280a1a82729c7ea",
		"BooleanFunctionsTestCaseData.xls": "703dc9a126ba17dd9fb05e45fe1e43061e76e6a464a15944b69283cbd9bc3d42",
		"CodeFunctionTestCaseData.xls": "b815357fd486912de6450c2e07d13afa591483c9b19908e131f6851564fa59de",
		"ComplexFunctionTestCaseData.xls": "e64097dffec5a2c7ce9226efc204aa1187430a2e6bea57d0aa40deee60006bcd",
		"DGet.xls": "9a3875f019a0a09b924ebca55da87a09060614fa352bcc49b30ca8ca76681ebe",
		"DStar.xls": "62f3ede92cecbfdcf07fc16077281473c4351ed2df450441c57149d0f7ee5c1b",
		"DateTimeToNumberTestCases.xls": "424f0d0c9a183458bfa3b7c09a8cbf4edcc9cec3d72b00e68e4bb0bf27292b6c",
		"DeltaFunctionTestCaseData.xls": "62635268793016fbba351ea0d671ad451a3497806977618ef6e4f0e16eb82240",
		"FactDoubleFunctionTestCaseData.xls": "09ff9ed74434f63953d0c1a2b447066182edad0ac53d529d05a370c1e1f9f6e2",
		"IfFunctionTestCaseData.xls": "cc757102a22d88304bc71dd43a5bc7dfbc34788cc6531875cbc580d68fcaf579",
		"IfNaTestCaseData.xls": "809f57c11771f5735af0b4ab371c9527eff52af3f3a83b1ad842466f682555bc",
		"ImRealFunctionTestCaseData.xls": "e53fa8d40e58737c3dcf99b5224add4ab1d3247404567895da7b813a8d92fad4",
		"ImaginaryFunctionTestCaseData.xls": "d75ca8896f8407543c1065110ae1a170a7e496624cbaaed941f54369ed19a1d6",
		"IndexFunctionTestCaseData.xls": "69417abbe4a48207ba408d6cd4e803afbef7bf765813462760ee3bfc7a3b04c7",
		"IndirectFunctionTestCaseData.xls": "42d538bafa0b44fec5fc03856010e34977fbf04fb254e1ab0d7e09f44239e2aa",
		"LogicalFunctionsTestCaseData.xls": "385120ae42a70a225315393b053ce56d11af48d4c39686262b4aa048b623dcb2",
		"LookupFunctionsTestCaseData.xls": "db71048f1034cf9b435469d7c8ab685501e62e5aa01bafe9209cd5717926f4ea",
		"MatchFunctionTestCaseData.xls": "1ebc7f281811efc6f6d09b1bb80f6d9e336cdaee19840c6c12ea5c2ac650afd0",
		"QuotientFunctionTestCaseData.xls": "3770fe8f5a3f910501d9ba6772860fb604b5e75449e3943e894a58e536fe2a3c",
		"ReptFunctionTestCaseData.xls": "f2bfb09f805bb19bb4aee2e022b7219f8d99f5b9016a14d347986cbbcce06dd6",
		"RomanFunctionTestCaseData.xls": "4ecbd643ff495d24bca0f0a90228db52de00a62fcc1b2964b44d728a0e5dfab6",
		"RowFunctionTestCaseData.xls": "35ed9e564ddea3f9e57fe8639d921654005478f29692a71a1de54209411c11a7",
		"Trend.xls": "959a7dcc64520544680595de30fbd155db88df5ad9f5e75704e4d846f3271f42",
		"TwoOperandNumericFunctionTestCaseData.xls": "1a3f2d8d9abaafe646e6735273eb73789cfb85d40a5513181bb174b07cb03188",
		"WeekNumFunctionTestCaseData.xls": "412ac3c961af66c897102b92260f5ba6b5c51045e1fd92dcadfc92c648291bdd",
		"WeekNumFunctionTestCaseData2013.xls": "e31ef3a53462d9e365cbc79478bc05762c47055e5f626ecd8fe97cd4dd5b44c2"
	}
} as const

export function suiteDir(suite: SuiteName): string {
	return join(SUITES_DIR, suite)
}

function sha256(bytes: Uint8Array): string {
	return createHash("sha256").update(bytes).digest("hex")
}

function gitBlobId(bytes: Uint8Array): string {
	return createHash("sha1").update(`blob ${bytes.length}\0`).update(bytes).digest("hex")
}

function filesUnder(directory: string): string[] {
	const out: string[] = []
	const walk = (path: string): void => {
		for (const name of readdirSync(path).sort()) {
			const full = join(path, name)

			if (statSync(full).isDirectory()) walk(full)
			else out.push(relative(directory, full))
		}
	}

	if (existsSync(directory)) walk(directory)

	return out.sort()
}

// The SHA-256 over a directory's files: sorted lines "<path>\0<SHA-256>\n".
export function setDigest(directory: string): string {
	const hash = createHash("sha256")

	for (const path of filesUnder(directory)) hash.update(`${path}\0${sha256(new Uint8Array(readFileSync(join(directory, path))))}\n`)

	return hash.digest("hex")
}

// The suite's files, verified: LibreOffice's against the set pin, the others each against its own.
export function suiteFiles(suite: SuiteName): string[] {
	const directory = join(suiteDir(suite), "files")

	if (suite === "libreoffice") {
		const digest = setDigest(directory)

		if (digest !== LIBREOFFICE.setSha256) throw new Error(`The LibreOffice function suite does not match its pin (${digest})`)

		return filesUnder(directory).map(path => join(directory, path))
	}

	const pins: Readonly<Record<string, string>> = suite === "formulajs" ? FORMULAJS.files : POI.files

	return Object.entries(pins).map(([name, digest]) => {
		const path = join(directory, name)

		if (sha256(new Uint8Array(readFileSync(path))) !== digest) throw new Error(`${suite} ${name} does not match its pin`)

		return path
	})
}

// Why the suite is not there to generate cases from, or undefined when it is.
export function suiteMissing(suite: SuiteName): string | undefined {
	return existsSync(join(suiteDir(suite), "complete"))
		? undefined
		: `the ${suite} function suite is not fetched (pnpm exec tsx src/tests/oracles/sheet/functionSuites.ts --fetch ${suite})`
}

async function download(url: string): Promise<Uint8Array> {
	const response = await fetch(url, { headers: { "User-Agent": "filen-office-oracles" } })

	if (!response.ok) throw new Error(`${url}: HTTP ${response.status}`)

	return new Uint8Array(await response.arrayBuffer())
}

// Downloads into a staging directory, which replaces the suite's only once every file checks out.
async function fetchSuite(suite: SuiteName): Promise<void> {
	const staging = join(SUITES_DIR, `${suite}.staging`)
	const files = join(staging, "files")

	rmSync(staging, { recursive: true, force: true })
	mkdirSync(files, { recursive: true })

	const write = (path: string, bytes: Uint8Array): void => {
		mkdirSync(dirname(join(files, path)), { recursive: true })
		writeFileSync(join(files, path), bytes)
	}

	if (suite === "libreoffice") {
		const listing = JSON.parse(
			new TextDecoder().decode(
				await download(`https://api.github.com/repos/LibreOffice/core/git/trees/${LIBREOFFICE.tree}?recursive=1`)
			)
		) as { truncated: boolean; tree: { path: string; type: string; sha: string }[] }
		const blobs = listing.tree.filter(entry => entry.type === "blob")

		if (listing.truncated || blobs.length !== LIBREOFFICE.files)
			throw new Error(`Unexpected LibreOffice tree listing (${blobs.length} files)`)

		// A few at a time: raw.githubusercontent.com throttles bursts.
		for (let start = 0; start < blobs.length; start += 8) {
			await Promise.all(
				blobs.slice(start, start + 8).map(async entry => {
					if (!/^[\w./-]+$/.test(entry.path) || entry.path.split("/").includes(".."))
						throw new Error(`Unexpected path ${entry.path}`)

					const bytes = await download(
						`https://raw.githubusercontent.com/LibreOffice/core/${LIBREOFFICE.commit}/${LIBREOFFICE.directory}/${entry.path}`
					)

					if (gitBlobId(bytes) !== entry.sha) throw new Error(`${entry.path} does not match its git blob id`)

					write(entry.path, bytes)
				})
			)
		}

		const digest = setDigest(files)

		if (digest !== LIBREOFFICE.setSha256) throw new Error(`The LibreOffice function suite does not match its pin (${digest})`)
	} else {
		const [base, pins] =
			suite === "formulajs"
				? [`https://raw.githubusercontent.com/formulajs/formulajs/${FORMULAJS.commit}/`, FORMULAJS.files]
				: [`https://raw.githubusercontent.com/apache/poi/${POI.commit}/${POI.directory}/`, POI.files]

		for (const [name, digest] of Object.entries(pins as Readonly<Record<string, string>>)) {
			const bytes = await download(base + name)

			if (sha256(bytes) !== digest) throw new Error(`${suite} ${name} does not match its pin`)

			write(name, bytes)
		}
	}

	writeFileSync(join(staging, "complete"), "")
	rmSync(suiteDir(suite), { recursive: true, force: true })
	renameSync(staging, suiteDir(suite))
}

export async function fetchFunctionSuites(suites: readonly SuiteName[]): Promise<void> {
	mkdirSync(SUITES_DIR, { recursive: true })

	for (const suite of suites) await fetchSuite(suite)
}

if (process.argv[1] === fileURLToPath(import.meta.url) && process.argv.includes("--fetch")) {
	const named = process.argv.slice(2).filter((arg): arg is SuiteName => arg === "libreoffice" || arg === "formulajs" || arg === "poi")

	await fetchFunctionSuites(named.length > 0 ? named : ["libreoffice", "formulajs", "poi"])
}
