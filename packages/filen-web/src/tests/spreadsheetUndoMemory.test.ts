import { describe, expect, it } from "vitest"
import { openXlsx, writeXlsx } from "hucre/xlsx"
import type { Cell } from "hucre"
import { proven, savedWorkbook, shownCell } from "@/tests/spreadsheetProven"

const RANGE = { startRow: 0, startCol: 0, endRow: 99, endCol: 99 }

// 200 rows by 100 columns of numbers; the first row italic, each cell with a format of its own as the
// reader gives them.
async function grid(historyBudget?: number) {
	const rows = Array.from({ length: 200 }, (_, row) => Array.from({ length: 100 }, (_, col) => row * 100 + col))
	const cells = new Map<string, Cell>()

	for (let col = 0; col < 100; col++) {
		cells.set(`0,${String(col)}`, { value: col, type: "number", style: { font: { italic: true } } })
	}

	return await proven(await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows, cells }] }), { readStyles: true }), historyBudget)
}

function styleOf(document: Awaited<ReturnType<typeof grid>>, row: number, col: number) {
	const view = shownCell(document, row, col)

	return view?.style === undefined ? undefined : document.doc().styles[view.style]
}

describe("XlsxDocument undo of formats", () => {
	it("keeps a format step to the formats it replaced, so many toggles over a large range stay undoable", async () => {
		// Each step holds 4 bytes a cell and its few distinct formats: thirty of them over 10,000 cells fit
		// in 2 MiB, where a copy of every cell would not hold two.
		const document = await grid(2 * 1024 * 1024)

		for (let step = 0; step < 30; step++) {
			document.apply({ type: "format", sheet: 0, range: RANGE, patch: { bold: step % 2 === 0 } })
		}

		let undone = 0

		while (document.undo().type !== "none") undone++

		expect(undone).toBe(30)
		expect(styleOf(document, 0, 5)).toMatchObject({ italic: true })
		expect(styleOf(document, 0, 5)?.bold).toBeUndefined()
		expect(styleOf(document, 5, 5)).toBeUndefined()
	})

	it("puts each cell's own format back exactly, and drops the details a format gave a plain cell", async () => {
		const document = await grid()

		document.apply({ type: "format", sheet: 0, range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 }, patch: { fill: "#ff0000" } })
		document.apply({ type: "format", sheet: 0, range: { startRow: 0, startCol: 0, endRow: 2, endCol: 2 }, patch: { bold: true } })

		expect(styleOf(document, 0, 0)).toMatchObject({ italic: true, bold: true, fill: "#ff0000" })
		expect(styleOf(document, 2, 2)).toMatchObject({ bold: true })

		document.undo()

		expect(styleOf(document, 0, 0)).toMatchObject({ italic: true, fill: "#ff0000" })
		expect(styleOf(document, 0, 0)?.bold).toBeUndefined()
		expect(styleOf(document, 2, 2)).toBeUndefined()

		document.undo()

		const saved = (await savedWorkbook(document)).sheets[0]

		expect(saved?.cells?.get("0,0")?.style).toEqual({ font: { italic: true } })
		expect(saved?.cells?.get("1,1")).toBeUndefined()
		expect(document.redo().type).toBe("cells")
		expect(styleOf(document, 1, 1)).toMatchObject({ fill: "#ff0000" })
	})

	it("shrinks a sheet a format grew back on undo", async () => {
		const document = await grid()

		document.apply({ type: "format", sheet: 0, range: { startRow: 300, startCol: 0, endRow: 301, endCol: 1 }, patch: { bold: true } })

		expect(document.doc().sheets[0]?.rowCount).toBe(302)

		document.undo()

		expect(document.doc().sheets[0]?.rowCount).toBe(200)
	})
})

describe("XlsxDocument undo weighs what replaced cells held", () => {
	const long = "x".repeat(100_000)
	// A cell holding 100,000 characters beside its value outweighs a 64 KiB history on its own.
	const holding: [string, Cell | undefined][] = [
		["nothing", undefined],
		["a formula", { value: null, type: "formula", formula: `LEN("${long}")` }],
		// Its runs are held beside the text they join into: 40,000 characters twice.
		["rich text", { value: long.slice(0, 40_000), type: "richText", richText: [{ text: long.slice(0, 40_000) }] }],
		["a note", { value: 1, type: "number", comment: { text: long } }],
		["a note's author", { value: 1, type: "number", comment: { text: "n", author: long } }],
		["a link", { value: 1, type: "number", hyperlink: { target: `https://example.com/${long}` } }]
	]

	for (const [what, cell] of holding) {
		it(`counts ${what}`, async () => {
			const cells = new Map<string, Cell>(cell === undefined ? [] : [["0,0", cell]])
			const workbook = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows: [[cell?.value ?? 1]], cells }] }), {
				readStyles: true
			})
			const document = await proven(workbook, 64 * 1024)

			document.apply({ type: "setCells", sheet: 0, cells: [{ row: 0, col: 0, input: "plain" }] })
			document.apply({ type: "setCells", sheet: 0, cells: [{ row: 1, col: 0, input: "2" }] })

			let undone = 0

			while (document.undo().type !== "none") undone++

			expect(undone).toBe(cell === undefined ? 2 : 1)
		})
	}
})
