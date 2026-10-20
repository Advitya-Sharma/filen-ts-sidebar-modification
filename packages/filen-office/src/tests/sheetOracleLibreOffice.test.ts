import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { afterAll, beforeAll, describe, expect, it } from "vitest"
import { bytesSource } from "../core/opc/index"
import type { SheetFileKind } from "../core/sheet/protocol"
import { RenderStyleTable } from "../core/sheet/session"
import { XlsxDocument } from "../core/xlsx/edit/document"
import { openXlsx } from "../core/xlsx/read/index"
import { fixtureBytes, FIXTURES } from "./fuzz/sheet/fixtures"
import { DEFAULT_SEQUENCE, isFinding, runSequence } from "./fuzz/sheet/harness"
import { calculates } from "./fuzz/sheet/recalc"
import { sheetFuzzTargets } from "./fuzz/sheet/run"
import { compareCached, formulaFunctions, VOLATILE } from "./oracles/sheet/excelCached"
import { exportSheets, sheetOffices, sheetProfile, unoOffice } from "./oracles/sheet/libreOffice"
import { UnoSheets, unoSheetByName, type UnoCell } from "./oracles/sheet/unoSheets"
import { formatA1, readXlsxCells, type CellScalar } from "./oracles/sheet/xlsxCells"

// LibreOffice on what the sessions save: workbooks edited by fuzz sequences (the fixtures, and corpus workbooks with
// FILEN_OFFICE_SHEET_LO_CORPUS=<count>) must open and convert to ods, csv and pdf in every install (no error, no
// timeout), and LibreOffice's hard recalculation of each (the UNO runner) must give our results for the formulas
// whose functions we calculate. Opt-in: FILEN_OFFICE_SOFFICE=<soffice>[,<soffice>]. Profiles and copies go under
// FILEN_OFFICE_SHEET_LO_SCRATCH (default: the system temp directory). FILEN_OFFICE_SHEET_LO_MIN (default 0.99) is
// the share of compared formula cells that must agree.

const INSTALLS = sheetOffices()
const UNO = unoOffice(INSTALLS)
const SCRATCH = process.env["FILEN_OFFICE_SHEET_LO_SCRATCH"] || tmpdir()
const CORPUS_COUNT = Number(process.env["FILEN_OFFICE_SHEET_LO_CORPUS"] ?? "0")
const SEEDS = 4
const CONVERT_TIMEOUT_MS = 120_000

interface EditedSave {
	readonly name: string
	readonly path: string
}

// Fuzz findings met while making the saves (the fuzz itself fails on them; here their workbooks are left out).
const findings: string[] = []

// The last save of one fuzz sequence (seed `seed`, sync) over `bytes`, or undefined when it saved nothing.
async function editedSave(bytes: Uint8Array, kind: SheetFileKind, label: string, seed: number): Promise<Uint8Array | undefined> {
	let last: Uint8Array | undefined
	const result = await runSequence(bytes, kind, label, seed, "sync", {
		...DEFAULT_SEQUENCE,
		edges: false,
		ranges: false,
		undoAll: false,
		reopen: false,
		keep: saved => {
			last = saved
		}
	})

	if (isFinding(result)) {
		findings.push(JSON.stringify(result).slice(0, 1500))

		return undefined
	}

	return last
}

// Our recalculation of a saved workbook: each formula cell's engine result, by "sheet!A1".
async function ourResults(bytes: Uint8Array): Promise<Map<string, CellScalar | undefined>> {
	const opened = await openXlsx(bytesSource(bytes))

	if (!opened.ok) throw new Error(opened.reason)

	const document = new XlsxDocument(opened.workbook, new RenderStyleTable(), { kind: "xlsx" })
	const out = new Map<string, CellScalar | undefined>()

	try {
		await document.readyAllSheets()

		const engine = document.buildEngine()
		const read = await readXlsxCells(bytes, cell => cell.formula !== undefined)

		document.engineSheets.forEach((sheet, index) => {
			const name = document.entry(sheet)?.name ?? ""

			for (const cell of read.sheets.find(candidate => candidate.name === name)?.cells ?? []) {
				const result = engine.result(index, cell.row, cell.column)
				const value = result.value

				out.set(
					`${name}!${formatA1(cell.row, cell.column)}`,
					value === null
						? undefined
						: result.error === true
							? { type: "e", value: String(value) }
							: typeof value === "number"
								? { type: "n", value }
								: typeof value === "boolean"
									? { type: "b", value }
									: { type: "s", value }
				)
			}
		})
	} finally {
		document.close()
	}

	return out
}

function unoScalar(cell: UnoCell): CellScalar {
	const [, , type, value] = cell

	return type === "n"
		? { type: "n", value: Number(value) }
		: type === "b"
			? { type: "b", value: value === true }
			: type === "e"
				? { type: "e", value: String(value) }
				: { type: "s", value: String(value) }
}

describe.skipIf(INSTALLS.length === 0)("LibreOffice on edited saves", () => {
	let work: string
	const saves: EditedSave[] = []

	beforeAll(async () => {
		mkdirSync(SCRATCH, { recursive: true })
		work = mkdtempSync(join(SCRATCH, "lo-c7-"))

		const targets = [
			...Object.keys(FIXTURES).map(name => ({ label: `fixture:${name}`, bytes: fixtureBytes(name), kind: "xlsx" as SheetFileKind })),
			...sheetFuzzTargets(undefined, 2_000_000)
				.filter(target => !target.startsWith("fixture:"))
				.slice(0, CORPUS_COUNT)
				.map(path => ({
					label: path,
					bytes: new Uint8Array(readFileSync(path)),
					kind: path.slice(path.lastIndexOf(".") + 1) as SheetFileKind
				}))
		]

		for (const target of targets)
			for (let seed = 1; seed <= (target.label.startsWith("fixture:") ? SEEDS : 1); seed++) {
				const bytes = await editedSave(target.bytes, target.kind, target.label, seed)

				if (bytes === undefined) continue

				const path = join(work, `${saves.length}.xlsx`)

				writeFileSync(path, bytes)
				saves.push({ name: `${target.label.slice(target.label.lastIndexOf("/") + 1)} s${seed}`, path })
			}
	}, 1_800_000)

	afterAll(() => {
		if (findings.length > 0) console.log(`sheet-lo: ${findings.length} fuzz findings left their saves out\n${findings.join("\n")}`)

		rmSync(work, { recursive: true, force: true })
	})

	for (const office of INSTALLS)
		it(`opens and converts every save (${office.name})`, { timeout: 3_600_000 }, async () => {
			expect(saves.length).toBeGreaterThan(0)

			const profile = sheetProfile(work, `c7-${office.name}`)
			const failures: string[] = []

			for (const target of ["ods", "csv", "pdf"] as const) {
				const results = await exportSheets(
					office,
					profile,
					target,
					saves.map(save => save.path),
					join(work, `${office.name}-${target}`),
					CONVERT_TIMEOUT_MS
				)

				results.forEach((result, index) => {
					if (!result.ok) failures.push(`${saves[index]?.name ?? result.input} → ${target}: ${result.error ?? "failed"}`)
				})
			}

			console.log(`sheet-lo ${office.name}: ${saves.length} edited saves to ods, csv and pdf, ${failures.length} failures`)
			expect(failures).toEqual([])
		})

	it.skipIf(UNO === undefined)("recalculates every save as we do", { timeout: 3_600_000 }, async () => {
		const uno = new UnoSheets(UNO as NonNullable<typeof UNO>, work)
		const mismatches: string[] = []
		const byFunction = new Map<string, number>()
		let compared = 0
		let matched = 0

		try {
			for (const save of saves) {
				const bytes = new Uint8Array(readFileSync(save.path))
				const ours = await ourResults(bytes)
				const result = await uno.run({ input: save.path, maxCells: 2_000_000 })

				expect(result.ok, `${save.name}: ${result.error ?? ""}`).toBe(true)

				if (result.calculated !== true) continue

				const read = await readXlsxCells(bytes, cell => cell.formula !== undefined)

				for (const sheet of read.sheets) {
					const theirs = unoSheetByName(result.before ?? { sheets: [], names: [] }, sheet.name)
					const byCell = new Map((theirs?.cells ?? []).map(cell => [`${cell[0]},${cell[1]}`, cell]))
					const masters = new Map<string, string>()

					for (const cell of sheet.cells)
						if (cell.formula?.si !== undefined && cell.formula.text !== "") masters.set(cell.formula.si, cell.formula.text)

					for (const cell of sheet.cells) {
						const formula = cell.formula

						if (formula === undefined || formula.kind === "array" || formula.kind === "dataTable") continue

						const text = formula.text !== "" || formula.si === undefined ? formula.text : (masters.get(formula.si) ?? "")
						const functions = formulaFunctions(text)

						if (!calculates(functions) || functions.some(name => VOLATILE.has(name)) || /\[\d+\]/.test(text)) continue

						const key = `${sheet.name}!${formatA1(cell.row, cell.column)}`
						const lo = byCell.get(`${cell.row},${cell.column}`)
						// LibreOffice dumps no empty cell: a formula it left empty reads as an empty text result.
						const expected: CellScalar = lo === undefined ? { type: "s", value: "" } : unoScalar(lo)
						const ourValue = ours.get(key)
						const actual = ourValue ?? { type: "n" as const, value: 0 }

						compared++

						const reason = compareCached(expected, actual)

						if (reason === undefined || (expected.type === "s" && expected.value === "" && ourValue === undefined)) {
							matched++
							continue
						}

						mismatches.push(`${save.name} ${key} =${text.slice(0, 80)}: ${reason.replace(", Excel ", ", LibreOffice ")}`)

						for (const name of functions.length === 0 ? ["(operators)"] : functions)
							byFunction.set(name, (byFunction.get(name) ?? 0) + 1)
					}
				}
			}
		} finally {
			await uno.close()
		}

		const rate = compared === 0 ? 1 : matched / compared

		console.log(
			`sheet-uno ${UNO?.name}: ${saves.length} saves, ${compared} formula cells compared, ${matched} agree (${(rate * 100).toFixed(2)} %); ` +
				`by function ${JSON.stringify([...byFunction].sort((a, z) => z[1] - a[1]).slice(0, 20))}\n${mismatches.slice(0, 40).join("\n")}`
		)
		expect(compared).toBeGreaterThan(0)
		expect(rate).toBeGreaterThanOrEqual(Number(process.env["FILEN_OFFICE_SHEET_LO_MIN"] ?? "0.99"))
	})
})

// Keeps the corpus helper honest when the corpus is absent: the fixtures alone still give saves.
describe("edited saves for LibreOffice", () => {
	it("come from fuzz sequences without findings", async () => {
		const bytes = await editedSave(fixtureBytes("unnamed"), "xlsx", "fixture:unnamed", 1)

		expect(findings).toEqual([])
		expect(bytes).toBeDefined()
		expect(existsSync(SCRATCH)).toBe(true)
	})
})
