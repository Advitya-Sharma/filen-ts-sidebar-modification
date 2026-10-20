import { spawnSync } from "node:child_process"
import {
	copyFileSync,
	existsSync,
	mkdirSync,
	mkdtempSync,
	readdirSync,
	readFileSync,
	renameSync,
	rmSync,
	statSync,
	writeFileSync
} from "node:fs"
import { tmpdir } from "node:os"
import { dirname, extname, join } from "node:path"
import { convertWith, offices, officeEnv, type Office } from "../../interop/soffice"

// Headless LibreOffice for the spreadsheet oracles: the installs FILEN_OFFICE_SOFFICE names (the same variable
// the document interop suite reads), each with whether its bundled Python runs (the UNO runner needs it), a
// profile pinned to en-US that never updates links or runs macros, and the "as shown" exports.
//
// LibreOffice's bundled Python on macOS is a shell wrapper around LibreOfficePython; on some installs that binary
// is killed at exec (SIGKILL, exit 137) although its signature verifies. Such an install still converts with
// --convert-to, and the UNO runner uses another install.

export interface SheetOffice extends Office {
	// The install's Resources directory (uno.py, fundamentalrc), or undefined when not found.
	readonly resources: string | undefined
	// Why its bundled Python cannot run the UNO runner, or undefined when it can.
	readonly unoMissing: string | undefined
}

const PROBE_TIMEOUT_MS = 30_000

function resourcesOf(soffice: string): string | undefined {
	const resources = join(dirname(soffice), "..", "Resources")

	return existsSync(join(resources, "python")) && existsSync(join(resources, "uno.py")) ? resources : undefined
}

// The environment the bundled Python needs for UNO, the parts its wrapper would set and -I keeps.
export function unoEnv(resources: string): NodeJS.ProcessEnv {
	return {
		...officeEnv(),
		URE_BOOTSTRAP: `vnd.sun.star.pathname:${join(resources, "fundamentalrc")}`,
		UNO_PATH: resources
	}
}

// Verdicts per install, kept across runs: a killed bundled Python writes a crash report (and, on macOS, a dialog)
// each time it is exec'd, and no static check (codesign, spctl) predicts the kill, so each install is exec-probed
// once. The key is the install plus the real Python binary's size and mtime, so a reinstall is probed again.
const PROBE_CACHE = join(tmpdir(), "filen-office-lo-probe.json")

function pythonKey(resources: string): string {
	const frameworks = join(resources, "..", "Frameworks", "LibreOfficePython.framework", "Versions")
	let binary = join(resources, "python")

	try {
		for (const version of readdirSync(frameworks)) {
			const candidate = join(frameworks, version, "Resources", "Python.app", "Contents", "MacOS", "LibreOfficePython")

			if (existsSync(candidate)) {
				binary = candidate

				break
			}
		}

		const stat = statSync(binary)

		return `${resources}|${stat.size}|${Math.floor(stat.mtimeMs)}`
	} catch {
		return `${resources}|unknown`
	}
}

function readProbeCache(): Record<string, string | null> {
	try {
		const parsed: unknown = JSON.parse(readFileSync(PROBE_CACHE, "utf8"))

		return typeof parsed === "object" && parsed !== null ? (parsed as Record<string, string | null>) : {}
	} catch {
		return {}
	}
}

function execProbe(resources: string): string | undefined {
	const result = spawnSync(join(resources, "python"), ["-I", "-c", "import sys; sys.stdout.write('ok')"], {
		encoding: "utf8",
		timeout: PROBE_TIMEOUT_MS,
		env: unoEnv(resources)
	})

	if (result.stdout === "ok" && result.status === 0) return undefined
	if (result.signal !== null) return `its bundled Python was killed at exec (${result.signal})`

	return `its bundled Python failed (exit ${result.status}): ${(result.stderr ?? "").slice(0, 200)}`
}

// Why the install's bundled Python cannot run, or undefined when it can. Exec'd at most once per install.
export function probePython(resources: string): string | undefined {
	const key = pythonKey(resources)
	const cache = readProbeCache()
	const cached = cache[key]

	if (cached !== undefined) return cached ?? undefined

	const verdict = execProbe(resources)

	try {
		const tmp = `${PROBE_CACHE}.${process.pid}.tmp`

		writeFileSync(tmp, JSON.stringify({ ...readProbeCache(), [key]: verdict ?? null }))
		renameSync(tmp, PROBE_CACHE)
	} catch {
		// An unwritable cache only costs a re-probe.
	}

	return verdict
}

export function sheetOffices(env: NodeJS.ProcessEnv = process.env): SheetOffice[] {
	return offices(env).map(office => {
		const resources = resourcesOf(office.path)
		const unoMissing = resources === undefined ? "no bundled Python next to soffice" : probePython(resources)

		return { ...office, resources, unoMissing }
	})
}

// The first install whose bundled Python runs, for the UNO runner.
export function unoOffice(installs: readonly SheetOffice[]): SheetOffice | undefined {
	return installs.find(office => office.unoMissing === undefined)
}

// Configuration every profile starts with: en-US formatting (shown text and CSV separators do not follow the
// host's locale), links never updated on load, macros never run, and an explicit recalculation-on-load mode for
// OOXML (left to itself, 26.8 recalculates some files on load and 7.5 does not).
const REGISTRY =
	'<?xml version="1.0" encoding="UTF-8"?>\n' +
	'<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n' +
	'<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooSetupSystemLocale" oor:op="fuse"><value>en-US</value></prop></item>\n' +
	'<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooLocale" oor:op="fuse"><value>en-US</value></prop></item>\n' +
	'<item oor:path="/org.openoffice.Office.Calc/Content/Update"><prop oor:name="Link" oor:op="fuse"><value>1</value></prop></item>\n' +
	'<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop></item>\n' +
	'<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="DisableMacrosExecution" oor:op="fuse"><value>true</value></prop></item>\n' +
	'<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>MODE</value></prop></item>\n' +
	'<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>MODE</value></prop></item>\n' +
	"</oor:items>\n"

// "never": the file's cached results are shown (what a viewer shows on open); "always": LibreOffice's own.
export type Recalculation = "never" | "always"

// A new profile under `root` (one per office and worker), configured before LibreOffice first starts on it.
export function sheetProfile(root: string, name: string, recalculation: Recalculation = "never"): string {
	const directory = join(root, `lo-${name}`)

	rmSync(directory, { recursive: true, force: true })
	mkdirSync(join(directory, "user"), { recursive: true })
	writeFileSync(join(directory, "user", "registrymodifications.xcu"), REGISTRY.replaceAll("MODE", recalculation === "never" ? "1" : "0"))

	return directory
}

// CSV options (LibreOffice's CSV filter tokens): comma, double quote, UTF-8, from line 1, standard cell format,
// language 0, quoted fields not text, special numbers detected, cell content AS SHOWN, results not formulas,
// spaces kept, every sheet to a file of its own ("<stem>-<sheet name>.csv").
const SHOWN_CSV = "csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,true,false,false,-1"

export const SHEET_TARGETS = {
	csv: SHOWN_CSV,
	html: "html:HTML (StarCalc)",
	xlsx: "xlsx:Calc MS Excel 2007 XML",
	ods: "ods:calc8",
	pdf: "pdf:calc_pdf_Export"
} as const

export type SheetTarget = keyof typeof SHEET_TARGETS

export interface SheetExport {
	readonly input: string
	readonly ok: boolean
	// The output files: one per sheet for csv ("<stem>-<sheet>.csv"), one otherwise.
	readonly outputs: readonly string[]
	readonly error?: string
}

// Converts copies of `inputs` (named by index, so outputs never collide and the originals are never handed to
// LibreOffice) to `target` in one soffice run, `timeoutMs` per input.
export async function exportSheets(
	office: Office,
	profile: string,
	target: SheetTarget,
	inputs: readonly string[],
	outdir: string,
	timeoutMs: number
): Promise<SheetExport[]> {
	mkdirSync(outdir, { recursive: true })

	const copies = mkdtempSync(join(outdir, "in-"))

	try {
		const named = inputs.map((input, index) => {
			const copy = join(copies, `${index}${extname(input).toLowerCase()}`)

			copyFileSync(input, copy)

			return copy
		})
		const batch = await convertWith(office, profile, SHEET_TARGETS[target], named, outdir, timeoutMs * Math.max(1, named.length))
		const files = readdirSync(outdir)

		return inputs.map((input, index) => {
			const outputs = files
				.filter(name =>
					target === "csv" ? name.startsWith(`${index}-`) || name === `${index}.csv` : name === `${index}.${target}`
				)
				.map(name => join(outdir, name))

			if (outputs.length > 0) return { input, ok: true, outputs }

			return { input, ok: false, outputs, error: batch.timedOut ? "timed out" : batch.stderr.slice(-300) || "no output" }
		})
	} finally {
		rmSync(copies, { recursive: true, force: true })
	}
}

// The sheet name of a per-sheet CSV output ("<index>-<sheet>.csv"); undefined for a single-sheet "<index>.csv".
export function csvSheetName(path: string): string | undefined {
	const name = path.slice(path.lastIndexOf("/") + 1, -".csv".length)
	const dash = name.indexOf("-")

	return dash < 0 ? undefined : name.slice(dash + 1)
}

// RFC 4180 fields of a CSV file, row by row (quoted fields may hold separators and line breaks).
export function parseCsv(text: string): string[][] {
	const rows: string[][] = []
	let row: string[] = []
	let field = ""
	let quoted = false
	let index = text.charCodeAt(0) === 0xfeff ? 1 : 0

	for (; index < text.length; index++) {
		const char = text[index]

		if (quoted) {
			if (char === '"') {
				if (text[index + 1] === '"') {
					field += '"'
					index++
				} else quoted = false
			} else field += char

			continue
		}

		if (char === '"') quoted = true
		else if (char === ",") {
			row.push(field)
			field = ""
		} else if (char === "\n" || char === "\r") {
			if (char === "\r" && text[index + 1] === "\n") index++

			row.push(field)
			rows.push(row)
			row = []
			field = ""
		} else field += char
	}

	if (field !== "" || row.length > 0) {
		row.push(field)
		rows.push(row)
	}

	return rows
}

export function readCsv(path: string): string[][] {
	return parseCsv(readFileSync(path, "utf8"))
}

export interface HtmlCell {
	// 0-based within the exported table, which starts at the sheet's data start (the first row and the first
	// column holding a value or a visible format), not at A1: alignHtml finds that offset.
	readonly row: number
	readonly column: number
	readonly rows: number
	readonly columns: number
	readonly text: string
	// The cell's background as #RRGGBB, conditional formats applied (see ShownWorkbook.conditionalFills).
	readonly fill: string | undefined
}

export interface HtmlSheet {
	readonly name: string
	readonly cells: readonly HtmlCell[]
}

const ENTITIES: Readonly<Record<string, string>> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: " " }

function decodeEntities(text: string): string {
	return text.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (whole, name: string) => {
		if (name.startsWith("#x") || name.startsWith("#X")) return String.fromCodePoint(Number.parseInt(name.slice(2), 16))
		if (name.startsWith("#")) return String.fromCodePoint(Number.parseInt(name.slice(1), 10))

		return ENTITIES[name.toLowerCase()] ?? whole
	})
}

function htmlAttribute(tag: string, name: string): string | undefined {
	return new RegExp(`\\b${name}=(?:"([^"]*)"|([^\\s>]+))`, "i")
		.exec(tag)
		?.slice(1)
		.find(value => value !== undefined)
}

// The tables of LibreOffice's "HTML (StarCalc)" export, one per sheet: every cell with its span, its text and
// its fill. Rows and columns count spans, as the grid does.
export function parseCalcHtml(html: string): HtmlSheet[] {
	const sheets: HtmlSheet[] = []
	const names = [...html.matchAll(/<h1>[^<]*<em>([\s\S]*?)<\/em><\/h1>/gi)].map(match => decodeEntities(match[1] ?? ""))
	let tableIndex = 0

	for (const table of html.matchAll(/<table\b[^>]*>([\s\S]*?)<\/table>/gi)) {
		const taken = new Set<string>()
		const cells: HtmlCell[] = []
		let row = 0

		for (const tr of (table[1] ?? "").matchAll(/<tr\b[^>]*>([\s\S]*?)<\/tr>/gi)) {
			let column = 0

			for (const td of (tr[1] ?? "").matchAll(/<t[dh]\b([^>]*)>([\s\S]*?)<\/t[dh]>/gi)) {
				while (taken.has(`${row},${column}`)) column++

				const tag = td[1] ?? ""
				const rows = Number(htmlAttribute(tag, "rowspan") ?? 1)
				const columns = Number(htmlAttribute(tag, "colspan") ?? 1)
				const fill = htmlAttribute(tag, "bgcolor")?.toUpperCase()
				const text = decodeEntities((td[2] ?? "").replace(/<br\s*\/?>/gi, "\n").replace(/<[^>]*>/g, "")).trim()

				for (let r = 0; r < rows; r++) for (let c = 0; c < columns; c++) taken.add(`${row + r},${column + c}`)

				cells.push({ row, column, rows, columns, text, fill })
				column += columns
			}

			row++
		}

		sheets.push({ name: names[tableIndex] ?? `Sheet${tableIndex + 1}`, cells })
		tableIndex++
	}

	return sheets
}

function normalize(text: string): string {
	return text.replace(/\s+/g, " ").trim()
}

// The 0-based A1 position of `html`'s first row and column, found by matching its texts against the sheet's "as
// shown" CSV (which starts at A1): the offset at which at least 90 % of them match. 7.5's HTML writes booleans as
// 1 and 0 where its CSV writes TRUE and FALSE: those count as matching. Undefined when the sheet has no text or
// no offset matches.
export function alignHtml(html: HtmlSheet, csv: readonly (readonly string[])[]): { row: number; column: number } | undefined {
	const texts = html.cells.map(cell => ({ cell, text: normalize(cell.text) })).filter(item => item.text !== "")

	if (texts.length === 0) return undefined

	// The longest text is the likeliest to be unique, so the fewest offsets are tried.
	const anchor = texts.reduce((best, item) => (item.text.length > best.text.length ? item : best))
	const allowed = Math.floor(texts.length / 10)

	for (let row = anchor.cell.row; row < csv.length; row++) {
		const line = csv[row] ?? []

		for (let column = anchor.cell.column; column < line.length; column++) {
			if (normalize(line[column] ?? "") !== anchor.text) continue

			const offset = { row: row - anchor.cell.row, column: column - anchor.cell.column }
			let misses = 0

			for (const item of texts) {
				const shown = normalize(csv[item.cell.row + offset.row]?.[item.cell.column + offset.column] ?? "")
				const same = shown === item.text || (shown === "TRUE" && item.text === "1") || (shown === "FALSE" && item.text === "0")

				if (!same && ++misses > allowed) break
			}

			if (misses <= allowed) return offset
		}
	}

	return undefined
}

export interface ShownSheet {
	readonly name: string
	// The sheet's cells as shown (number formats applied), from A1: rows of fields, trailing empties as exported.
	readonly text: readonly (readonly string[])[]
	// "<row>,<column>" (0-based, A1-aligned) -> #RRGGBB background; undefined when the HTML table could not be
	// aligned or was not exported.
	readonly fills: ReadonlyMap<string, string> | undefined
}

export interface ShownWorkbook {
	readonly input: string
	readonly ok: boolean
	readonly error?: string
	readonly sheets: readonly ShownSheet[]
	// Whether `fills` include every conditional format's fill: 26.8's HTML export applies color scales too, 7.5's
	// only rule fills (dxf); versions between were not measured, so they count as 7.5.
	readonly conditionalFills: boolean
}

// The "as shown" text of every sheet (CSV export) and, with `fills`, every cell's background (HTML export),
// A1-aligned. The values shown are what `profile`'s recalculation mode leaves (sheetProfile).
export async function exportShown(
	office: Office,
	profile: string,
	inputs: readonly string[],
	outdir: string,
	timeoutMs: number,
	fills = true
): Promise<ShownWorkbook[]> {
	const csv = await exportSheets(office, profile, "csv", inputs, join(outdir, "csv"), timeoutMs)
	const html = fills ? await exportSheets(office, profile, "html", inputs, join(outdir, "html"), timeoutMs) : undefined
	const [major = 0, minor = 0] = office.version.split(".").map(Number)
	const conditionalFills = major > 26 || (major === 26 && minor >= 8)

	return csv.map((result, index) => {
		if (!result.ok) return { input: result.input, ok: false, error: result.error ?? "no output", sheets: [], conditionalFills }

		const tables = html?.[index]?.ok === true ? parseCalcHtml(readFileSync(html[index].outputs[0] ?? "", "utf8")) : []
		const sheets = result.outputs.map((path): ShownSheet => {
			const name = csvSheetName(path) ?? ""
			const text = readCsv(path)
			const table = tables.length === 1 && result.outputs.length === 1 ? tables[0] : tables.find(item => item.name === name)
			const origin = table === undefined ? undefined : alignHtml(table, text)

			if (table === undefined || origin === undefined) return { name, text, fills: undefined }

			const map = new Map<string, string>()

			for (const cell of table.cells)
				if (cell.fill !== undefined)
					for (let r = 0; r < cell.rows; r++)
						for (let c = 0; c < cell.columns; c++)
							map.set(`${origin.row + cell.row + r},${origin.column + cell.column + c}`, cell.fill)

			return { name, text, fills: map }
		})

		return { input: result.input, ok: true, sheets, conditionalFills }
	})
}
