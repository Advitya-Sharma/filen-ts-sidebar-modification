import { readFileSync } from "node:fs"
import { dirname, join } from "node:path"
import { describe, expect, it, vi } from "vitest"
import { writeXlsx } from "hucre/xlsx"
import { applyEditResult, gridDoc, type GridDoc, type GridSheet } from "@/features/spreadsheet/lib/cellStore.logic"
import type { EditResult } from "@/features/spreadsheet/lib/edits"
import { cellKey, type SheetView } from "@/features/spreadsheet/lib/model"
import type { XlsxDocument } from "@/features/spreadsheet/lib/xlsxDocument"
import type { CsvWorkerApi } from "@/features/spreadsheet/workers/csv.worker"
import type { SpreadsheetOpened, SpreadsheetWorkerApi } from "@/features/spreadsheet/workers/spreadsheet.worker"
import { formula, openSheets } from "@/tests/spreadsheetProven"

// A workbook opens with its active sheet's view alone, the others deferred until the page first shows them;
// edits send nothing for a deferred sheet, whose view, once asked for, is built from the workbook as it then
// stands. Text opens in a worker without the workbook reader and formula engine, and a proven workbook's
// formula engine is built once it is proven, not at its first edit.

const { exposed, enginesBuilt } = vi.hoisted(() => ({ exposed: [] as unknown[], enginesBuilt: { count: 0 } }))

vi.mock(import("comlink"), async importOriginal => ({
	...(await importOriginal()),
	expose: (api: unknown) => {
		exposed.push(api)
	}
}))
vi.mock(import("@/features/spreadsheet/lib/formulaEngine"), async importOriginal => {
	const actual = await importOriginal()

	class CountedEngine extends actual.FormulaEngine {
		constructor(...args: ConstructorParameters<typeof actual.FormulaEngine>) {
			super(...args)
			enginesBuilt.count++
		}
	}

	return { ...actual, FormulaEngine: CountedEngine }
})

await import("@/features/spreadsheet/workers/spreadsheet.worker")
await import("@/features/spreadsheet/workers/csv.worker")

const workbookWorker = exposed[0] as SpreadsheetWorkerApi
const csvWorker = exposed[1] as CsvWorkerApi

function nextTask(): Promise<void> {
	return new Promise(resolve => {
		setTimeout(resolve, 0)
	})
}

function workbook(): Promise<XlsxDocument> {
	return openSheets([
		{ name: "Data", rows: [[1], [2], [3]] },
		{ name: "Summary", rows: [[6, "x"]], cells: new Map([["0,0", formula("SUM(Data!A1:A3)", 6)]]) },
		{ name: "Styled", rows: [[1]], cells: new Map([["0,0", { value: 1, type: "number", style: { font: { bold: true } } }]]) }
	])
}

// The one sheet a result sends whole, and the style table it sends.
function sent(result: EditResult): { sheets: (SheetView | null)[]; styles: unknown[] } {
	if (result.type !== "sheets") {
		throw new Error(`expected a sheets result, not ${result.type}`)
	}

	return { sheets: result.sheets, styles: result.styles }
}

function comparable(sheet: GridSheet | SheetView | undefined): object | undefined {
	if (sheet === undefined) {
		return undefined
	}

	const cells = sheet.cells

	return { ...sheet, cells: [...cells.keys()].sort((a, b) => a - b).map(key => [key, cells.get(key)]) }
}

// The page's sheets that are built equal fresh views of the workbook.
function expectInSync(held: GridDoc, document: XlsxDocument): void {
	const fresh = gridDoc(document.doc())

	// doc() builds every sheet: take the page's deferred ones as deferred again.
	document.doc(true)
	held.sheets.forEach((sheet, index) => {
		if (sheet.deferred !== true) {
			document.sheet(index)
			expect(comparable(sheet)).toEqual(comparable(fresh.sheets[index]))
		}
	})
	expect(held.styles).toEqual(fresh.styles.slice(0, held.styles.length))
}

describe("deferred sheets", () => {
	it("opens with the active sheet's view alone, and builds another once asked, keeping the rest", async () => {
		const document = await workbook()
		const full = document.doc()
		const opened = document.doc(true)

		expect(opened.sheets.map(sheet => sheet.deferred)).toEqual([undefined, true, true])
		expect(opened.sheets[1]).toMatchObject({ name: "Summary", rowCount: 1, colCount: 2, structureLocked: true })
		expect(opened.sheets[1]?.cells.size).toBe(0)

		const answer = document.sheet(1)
		const shown = sent(answer)

		expect(shown.sheets).toEqual([null, full.sheets[1], null])
		expect(shown.styles).toEqual(full.styles)

		const held = applyEditResult(gridDoc(opened), answer)

		// Asked again, nothing is sent anew: the first answer is on its way.
		expect(sent(document.sheet(1)).sheets).toEqual([null, null, null])
		expect(held.sheets.map(sheet => sheet.deferred)).toEqual([undefined, undefined, true])
	})

	it("sends the style table as building a sheet grew it", async () => {
		const document = await workbook()
		const styles = document.doc(true).styles.length
		const shown = sent(document.sheet(2))

		expect(shown.styles.length).toBeGreaterThan(styles)
		expect(shown.sheets[2]?.cells.get(cellKey(0, 0))?.style).toBe(shown.styles.length - 1)
	})

	it("sends nothing for a deferred sheet an edit recalculates, which shows the result once built", async () => {
		const document = await workbook()
		let held = gridDoc(document.doc(true))
		const edited = document.apply({ type: "setCells", sheet: 0, cells: [{ row: 0, col: 0, input: "10" }] })

		expect(edited.type).toBe("cells")
		expect(edited.type === "cells" ? edited.patches.map(patch => patch.sheet) : null).toEqual([0])

		held = applyEditResult(held, edited)
		held = applyEditResult(held, document.sheet(1))
		expect(held.sheets[1]?.cells.get(cellKey(0, 0))).toMatchObject({ text: "15" })

		// Once built, it is patched.
		const again = document.apply({ type: "setCells", sheet: 0, cells: [{ row: 1, col: 0, input: "20" }] })

		expect(again.type === "cells" ? again.patches.map(patch => patch.sheet).sort() : null).toEqual([0, 1])
		held = applyEditResult(held, again)
		expect(held.sheets[1]?.cells.get(cellKey(0, 0))).toMatchObject({ text: "33" })
		expectInSync(held, document)
	})

	it("keeps a deferred sheet deferred through structural edits, undo and redo, renamed as the workbook is", async () => {
		const document = await workbook()
		let held = gridDoc(document.doc(true))
		const fold = (result: ReturnType<XlsxDocument["apply"]>): void => {
			held = applyEditResult(held, result)
			expectInSync(held, document)
		}

		const inserted = document.apply({ type: "insert", sheet: 0, axis: "rows", at: 0, count: 1 })

		expect(inserted).toMatchObject({ type: "sheets", sheets: [expect.any(Object), null, null], patches: [], extents: [] })
		fold(inserted)
		expect(held.sheets[1]?.deferred).toBe(true)

		fold(document.apply({ type: "renameSheet", sheet: 1, name: "Totals" }))
		expect(held.sheets[1]).toMatchObject({ name: "Totals", deferred: true })

		fold(document.apply({ type: "addSheet", name: "Later" }))
		expect(held.sheets.map(sheet => sheet.deferred)).toEqual([undefined, true, true, undefined])

		fold(document.undo())
		fold(document.undo())
		expect(held.sheets.map(sheet => sheet.name)).toEqual(["Data", "Summary", "Styled"])
		fold(document.redo())

		held = applyEditResult(held, document.sheet(1))
		expect(held.sheets[1]).toMatchObject({ name: "Totals" })
		// The insert moved Data!A1:A3 down a row; the formula follows.
		expect(held.sheets[1]?.cells.get(cellKey(0, 0))).toMatchObject({ text: "6", input: "=SUM(Data!A2:A4)" })
		expectInSync(held, document)
	})
})

describe("the workbook worker", () => {
	async function open(bytes: Uint8Array): Promise<SpreadsheetOpened> {
		const result = await workbookWorker.open(bytes, "xlsx")

		if (!("id" in result)) {
			throw new Error("expected an opened workbook")
		}

		return result
	}

	async function formulaBook(): Promise<Uint8Array> {
		return await writeXlsx({
			sheets: [
				{ name: "S", rows: [[2, 4]], cells: new Map([["0,1", formula("A1*2", 4)]]) },
				{ name: "T", rows: [[1]] }
			]
		})
	}

	it("opens a workbook with its other sheets deferred, and builds one once shown", async () => {
		const bytes = await formulaBook()
		const opened = await open(bytes)

		expect(opened.doc.sheets.map(sheet => sheet.deferred)).toEqual([undefined, true])
		expect(sent(workbookWorker.sheet(opened.id, 1)).sheets[1]).toMatchObject({
			name: "T",
			cells: new Map([[cellKey(0, 0), { text: "1", numeric: true }]])
		})
		workbookWorker.close(opened.id)
	})

	it("builds a proven workbook's formula engine after answering that it is writable, before its first edit", async () => {
		const bytes = await formulaBook()
		const opened = await open(bytes)
		const before = enginesBuilt.count

		await expect(workbookWorker.writability(opened.id, bytes.slice())).resolves.toBe(true)
		// The answer goes first.
		expect(enginesBuilt.count).toBe(before)

		await nextTask()
		expect(enginesBuilt.count).toBe(before + 1)

		await expect(
			workbookWorker.apply(opened.id, { type: "setCells", sheet: 0, cells: [{ row: 0, col: 0, input: "3" }] })
		).resolves.toMatchObject({
			type: "cells",
			patches: [
				{
					cells: [
						[cellKey(0, 0), { text: "3" }],
						[cellKey(0, 1), { text: "6" }]
					]
				}
			]
		})
		expect(enginesBuilt.count).toBe(before + 1)
		workbookWorker.close(opened.id)
	})

	it("builds no engine for a workbook without formulas, or one closed before it would", async () => {
		const plain = await writeXlsx({ sheets: [{ name: "S", rows: [[1]] }] })
		const before = enginesBuilt.count
		const first = await open(plain)

		await expect(workbookWorker.writability(first.id, plain.slice())).resolves.toBe(true)
		await nextTask()

		const bytes = await formulaBook()
		const second = await open(bytes)

		await workbookWorker.writability(second.id, bytes.slice())
		workbookWorker.close(second.id)
		await nextTask()

		expect(enginesBuilt.count).toBe(before)
		workbookWorker.close(first.id)
	})
})

describe("the text worker", () => {
	it("opens, edits and saves CSV and TSV", () => {
		const opened = csvWorker.open(new TextEncoder().encode("a\tb\n1\t2\n"), "tsv")

		expect(opened.doc).toMatchObject({ kind: "csv", writable: true })
		expect(csvWorker.apply(opened.id, { type: "setCells", sheet: 0, cells: [{ row: 1, col: 1, input: "3" }] })).toMatchObject({
			type: "cells"
		})

		const { bytes, version } = csvWorker.serialize(opened.id)

		expect(new TextDecoder().decode(bytes)).toBe("a\tb\n1\t3\n")
		expect(csvWorker.markSaved(opened.id, version)).toMatchObject({ dirty: false })
		csvWorker.close(opened.id)
		expect(() => csvWorker.undo(opened.id)).toThrow("no open document")
	})

	// Every module the text worker reaches, through the app's own imports: none is the workbook reader or the
	// formula engine (the bundle check counts their bytes).
	it("reaches neither the workbook reader nor the formula engine", () => {
		const SRC = join(dirname(new URL(import.meta.url).pathname), "..")
		const seen = new Set<string>()
		const packages = new Set<string>()
		const visit = (file: string): void => {
			if (seen.has(file)) return

			seen.add(file)

			for (const [, specifier] of readFileSync(file, "utf8").matchAll(/^import\s(?!type\b)[^"]*"([^"]+)"/gm)) {
				if (specifier === undefined) continue

				if (specifier.startsWith("@/")) {
					visit(join(SRC, `${specifier.slice(2).replace(/\.ts$/, "")}.ts`))
				} else {
					packages.add(specifier)
				}
			}
		}

		visit(join(SRC, "features/spreadsheet/workers/csv.worker.ts"))
		expect(seen.size).toBeGreaterThan(3)
		expect([...packages].sort()).toEqual(["comlink", "hucre"])
	})
})
