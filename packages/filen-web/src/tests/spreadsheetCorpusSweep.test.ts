import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs"
import { dirname, join, resolve } from "node:path"
import { describe, expect, it, vi } from "vitest"
import type {
	SpreadsheetFileKind,
	SpreadsheetOpen,
	SpreadsheetOpened,
	SpreadsheetUnlock,
	SpreadsheetWorkerApi
} from "@/features/spreadsheet/workers/spreadsheet.worker"
import type { CsvWorkerApi } from "@/features/spreadsheet/workers/csv.worker"

// How much of the office corpus the spreadsheet worker can edit, opened as the page opens a file: each
// spreadsheet of corpus/manifest.json goes through the worker's open (a protected one unlocked with its
// manifest password, as the password prompt would), and then through the save proof the page asks for before
// the first edit. Each file lands in one outcome: editable, view-only (with why), the proof failed, or it did
// not open. Tallied per extension and for the GOV.UK publications, so a later engine is compared on the same
// files.
//
// Opt-in: FILEN_SHEET_CORPUS_SWEEP=1 (the corpus is never committed: `pnpm run corpus:fetch` in
// packages/filen-office, or FILEN_OFFICE_CORPUS). FILEN_SHEET_CORPUS_SWEEP_OUT=<file.json> also writes every
// file's outcome and the table.

const { exposed } = vi.hoisted(() => {
	const api: { api: unknown } = { api: undefined }

	return { exposed: api }
})

vi.mock(import("comlink"), async importOriginal => ({
	...(await importOriginal()),
	expose: (api: unknown) => {
		exposed.api = api
	}
}))

const ENABLED = process.env["FILEN_SHEET_CORPUS_SWEEP"] === "1"
const OUT = process.env["FILEN_SHEET_CORPUS_SWEEP_OUT"]
const OFFICE = join(process.cwd(), "..", "filen-office")
const configured = process.env["FILEN_OFFICE_CORPUS"]
const CORPUS = configured !== undefined && configured !== "" ? resolve(configured) : join(OFFICE, ".corpus")

// The worker's kind for each extension it opens (an .xlt is not routed to the spreadsheet viewer).
const KINDS: Record<string, SpreadsheetFileKind> = {
	xlsx: "xlsx",
	xlsm: "xlsx",
	xltx: "xlsx",
	xltm: "xlsx",
	xls: "xls",
	csv: "csv",
	tsv: "tsv"
}
const EXTENSIONS = Object.keys(KINDS)
const GOVUK = "govuk"

interface Sample {
	path: string
	source: string
	password?: string
}

type Outcome = { kind: "editable" } | { kind: "viewOnly"; reason: string } | { kind: "proofFail" } | { kind: "openError"; reason: string }

interface Swept {
	path: string
	extension: string
	source: string
	encrypted: boolean
	outcome: Outcome
	openMs: number
	proofMs: number | null
}

const manifest = JSON.parse(readFileSync(join(OFFICE, "corpus", "manifest.json"), "utf8")) as { files: Sample[] }
const SAMPLES = manifest.files
	.filter(file => EXTENSIONS.includes(file.path.split("/")[0] ?? ""))
	.filter(file => existsSync(join(CORPUS, file.path)))
const swept: Swept[] = []

// Workbooks open in the workbook worker, CSV and TSV in the text worker, as on the page.
const workbookApi: SpreadsheetWorkerApi | null = ENABLED
	? await import("@/features/spreadsheet/workers/spreadsheet.worker").then(() => exposed.api as SpreadsheetWorkerApi)
	: null
const textApi: CsvWorkerApi | null = ENABLED
	? await import("@/features/spreadsheet/workers/csv.worker").then(() => exposed.api as CsvWorkerApi)
	: null

function worker(): SpreadsheetWorkerApi {
	if (workbookApi === null) {
		throw new Error("the sweep is not enabled")
	}

	return workbookApi
}

function textWorker(): CsvWorkerApi {
	if (textApi === null) {
		throw new Error("the sweep is not enabled")
	}

	return textApi
}

function message(error: unknown): string {
	return (error instanceof Error ? error.message : String(error)).split("\n")[0]?.slice(0, 160) ?? ""
}

// An open's answer: the document, or why there is none. A locked workbook is unlocked with the manifest's
// password; without one it stays unopened.
async function opened(result: SpreadsheetOpen, password: string | undefined): Promise<SpreadsheetOpened | { reason: string }> {
	if ("refused" in result) {
		return { reason: `refused: ${result.refused}` }
	}

	if (!("locked" in result)) {
		return result
	}

	if (password === undefined) {
		worker().discardLocked(result.locked)

		return { reason: "locked: no password" }
	}

	const unlocked: SpreadsheetUnlock = await worker().unlock(result.locked, password)

	if ("id" in unlocked) {
		return unlocked
	}

	worker().discardLocked(result.locked)

	return { reason: "refused" in unlocked ? `refused: ${unlocked.refused}` : "wrongPassword" in unlocked ? "wrong password" : "aborted" }
}

async function sweep(sample: Sample): Promise<Swept> {
	const extension = sample.path.split("/")[0] ?? ""
	const kind = KINDS[extension]

	if (kind === undefined) {
		throw new Error(`not a spreadsheet: ${sample.path}`)
	}

	const bytes = new Uint8Array(readFileSync(join(CORPUS, sample.path)))
	const text = kind === "csv" || kind === "tsv"
	const base = { path: sample.path, extension, source: sample.source }
	let started = performance.now()
	let document: SpreadsheetOpened | { reason: string }

	try {
		document =
			kind === "csv" || kind === "tsv"
				? textWorker().open(bytes.slice(), kind)
				: await opened(await worker().open(bytes.slice(), kind), sample.password)
	} catch (error) {
		return {
			...base,
			encrypted: false,
			outcome: { kind: "openError", reason: message(error) },
			openMs: performance.now() - started,
			proofMs: null
		}
	}

	const openMs = performance.now() - started

	if (!("id" in document)) {
		return { ...base, encrypted: false, outcome: { kind: "openError", reason: document.reason }, openMs, proofMs: null }
	}

	const encrypted = document.encrypted !== undefined
	const { doc } = document

	try {
		if (doc.kind === "xls") {
			return { ...base, encrypted, outcome: { kind: "viewOnly", reason: "xls" }, openMs, proofMs: null }
		}

		if (doc.unprovable !== undefined) {
			return { ...base, encrypted, outcome: { kind: "viewOnly", reason: doc.unprovable }, openMs, proofMs: null }
		}

		started = performance.now()

		// The page sends the file as opened; a protected workbook's plaintext is the worker's own.
		const writable = text
			? textWorker().writability(document.id)
			: await worker().writability(document.id, encrypted ? null : bytes.slice())
		const proofMs = performance.now() - started

		if (doc.kind === "csv") {
			return {
				...base,
				encrypted,
				outcome: writable ? { kind: "editable" } : { kind: "viewOnly", reason: "format" },
				openMs,
				proofMs
			}
		}

		return { ...base, encrypted, outcome: writable ? { kind: "editable" } : { kind: "proofFail" }, openMs, proofMs }
	} finally {
		if (text) {
			textWorker().close(document.id)
		} else {
			worker().close(document.id)
		}
	}
}

const COLUMNS = ["editable", "viewOnly", "proofFail", "openError"] as const

interface Row {
	group: string
	files: number
	counts: Record<(typeof COLUMNS)[number], number>
	reasons: Record<string, number>
}

function tally(group: string, items: readonly Swept[]): Row {
	const counts = { editable: 0, viewOnly: 0, proofFail: 0, openError: 0 }
	const reasons: Record<string, number> = {}

	for (const item of items) {
		counts[item.outcome.kind]++

		if ("reason" in item.outcome) {
			const key = `${item.outcome.kind}: ${item.outcome.reason}`

			reasons[key] = (reasons[key] ?? 0) + 1
		}
	}

	return { group, files: items.length, counts, reasons }
}

function table(rows: readonly Row[]): string {
	const lines = ["| | files | editable | view-only | proof fail | open error |", "|---|---|---|---|---|---|"]

	for (const row of rows) {
		lines.push(`| ${row.group} | ${String(row.files)} | ${COLUMNS.map(column => String(row.counts[column])).join(" | ")} |`)
	}

	return lines.join("\n")
}

describe.skipIf(!ENABLED || SAMPLES.length === 0)("spreadsheet corpus sweep", () => {
	it.each(SAMPLES.map(sample => [sample.path, sample] as const))("%s", { timeout: 300_000 }, async (_path, sample) => {
		const result = await sweep(sample)

		swept.push(result)

		// A protected workbook the manifest has the password for always opens past its prompt.
		if (sample.password !== undefined && result.outcome.kind === "openError") {
			expect(result.outcome.reason).not.toMatch(/^(locked|wrong password)/u)
		}
	})

	it("tallies every file once, per extension and for the GOV.UK publications", () => {
		expect(swept.map(item => item.path).sort()).toEqual(SAMPLES.map(sample => sample.path).sort())

		const rows = [
			...EXTENSIONS.map(extension =>
				tally(
					extension,
					swept.filter(item => item.extension === extension)
				)
			).filter(row => row.files > 0),
			tally(
				"GOV.UK (OOXML)",
				swept.filter(item => item.source === GOVUK && KINDS[item.extension] === "xlsx")
			)
		]
		const shown = table(rows)

		for (const row of rows) {
			expect(COLUMNS.reduce((sum, column) => sum + row.counts[column], 0)).toBe(row.files)
		}

		console.log(`${shown}\n\n${rows.map(row => `${row.group}: ${JSON.stringify(row.reasons)}`).join("\n")}`)

		if (OUT !== undefined && OUT !== "") {
			mkdirSync(dirname(OUT), { recursive: true })
			writeFileSync(OUT, JSON.stringify({ table: shown, rows, files: swept }, null, "\t"))
		}
	})
})
