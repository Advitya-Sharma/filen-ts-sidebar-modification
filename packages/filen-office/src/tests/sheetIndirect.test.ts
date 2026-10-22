import { strFromU8, unzipSync } from "fflate"
import { describe, expect, it } from "vitest"
import type { SheetCommandResult, SheetSessionApi, SheetTile } from "../core/sheet/protocol"
import { inThreadSheetHost } from "../dom/sheet/host"
import { buildXlsx } from "./oracles/sheet/synthetic"
import { FormulaEngine, type EngineName, type EngineSheet, type RecalculatedCell } from "../core/sheet/formula/index"
import { engineCellFormula, storedFormula } from "../core/sheet/refs/formulaRefs"

// INDIRECT (support.microsoft.com "INDIRECT function"): the reference a text names, A1 or R1C1, a defined name's
// included; calculated again whenever a cell it read changes, though the engine never linked it to that cell.

type Typed = string | number | boolean | null

function engineRows(rows: Typed[][]): EngineSheet["rows"] {
	return rows.map(row =>
		row.map(cell => (typeof cell === "string" && cell.startsWith("=") ? (engineCellFormula(storedFormula(cell.slice(1))) ?? "") : cell))
	)
}

function build(rows: Typed[][], other: Typed[][] = [], names: EngineName[] = []): FormulaEngine {
	return new FormulaEngine(
		[
			{ name: "S", rows: engineRows(rows), formulas: [] },
			{ name: "My Sheet", rows: engineRows(other), formulas: [] }
		],
		names
	)
}

function value(engine: FormulaEngine, a1: string, sheet = 0): unknown {
	const col = a1.charCodeAt(0) - 65
	const row = Number(a1.slice(1)) - 1
	const result = engine.result(sheet, row, col)

	return result.error === true ? `!${String(result.value)}` : result.value
}

function set(engine: FormulaEngine, a1: string, content: Typed): RecalculatedCell[] {
	const [cell] = engineRows([[content]])

	return engine.set(0, [{ row: Number(a1.slice(1)) - 1, col: a1.charCodeAt(0) - 65, content: cell?.[0] ?? null }])
}

describe("INDIRECT", () => {
	it("reads the reference a text names, in A1 or R1C1 style", () => {
		// "INDIRECT function": =INDIRECT(A2) gives the value of the cell A2 names; a1 FALSE reads R1C1.
		const engine = build(
			[
				[1, "B2", '=INDIRECT("A1")', "=INDIRECT(B1)", '=INDIRECT("R1C1",FALSE)', '=INDIRECT("R[1]C[-5]",FALSE)'],
				[
					2,
					7,
					'=SUM(INDIRECT("A1:A3"))',
					"=INDIRECT(\"'My Sheet'!B1\")",
					'=SUM(INDIRECT("A:A"))',
					'=COUNTIF(INDIRECT("A1:A3"),">1")'
				],
				[3, null, '=INDIRECT("nonsense")', '=INDIRECT("[1]S!A1")', '=INDIRECT("Nowhere!A1")', '=INDIRECT("F3")']
			],
			[[null, "x"]]
		)

		try {
			expect([value(engine, "C1"), value(engine, "D1"), value(engine, "E1"), value(engine, "F1")]).toEqual([1, 7, 1, 2])
			expect([value(engine, "C2"), value(engine, "D2"), value(engine, "E2"), value(engine, "F2")]).toEqual([6, "x", 6, 2])
			expect([value(engine, "C3"), value(engine, "D3"), value(engine, "E3")]).toEqual(["!#REF!", "!#REF!", "!#REF!"])
			// Its own cell: a circle.
			expect(value(engine, "F3")).toBe("!#CYCLE!")
		} finally {
			engine.destroy()
		}
	})

	it("reads a defined name, the sheet's own scope first", () => {
		const engine = build(
			[[1, 2, 3, '=SUM(INDIRECT("Data"))', '=INDIRECT("Total")']],
			[],
			[
				{ name: "Data", expression: "S!$A$1:$C$1" },
				{ name: "Total", expression: "S!$D$1" },
				{ name: "Data", expression: "S!$A$1", scope: 0 }
			]
		)

		try {
			expect([value(engine, "D1"), value(engine, "E1")]).toEqual([1, 1])
		} finally {
			engine.destroy()
		}
	})

	it("calculates again when a cell it read changes, a value or a formula, and reports it", () => {
		const engine = build([[1, '=INDIRECT("A1")', "=B1*10", "=E1*2", 5, '=INDIRECT("D1")+1']])

		try {
			expect([value(engine, "B1"), value(engine, "C1"), value(engine, "F1")]).toEqual([1, 10, 11])

			const changes = set(engine, "A1", 4)

			expect(changes.map(cell => `${String(cell.col)}=${String(cell.value)}`).sort()).toEqual(["0=4", "1=4", "2=40"])

			// E1 changes D1, a formula, which F1 reads through INDIRECT.
			set(engine, "E1", 50)
			expect(value(engine, "F1")).toBe(101)
		} finally {
			engine.destroy()
		}
	})

	it("is right at the first calculation whatever order it comes in", () => {
		const engine = build([['=INDIRECT("B1")*2', "=C1+1", "=D1*3", 2, '=INDIRECT("A1")+INDIRECT("B1")']])

		try {
			expect([value(engine, "A1"), value(engine, "E1")]).toEqual([14, 21])
		} finally {
			engine.destroy()
		}
	})

	it("follows the text, not the cells, when rows are inserted", () => {
		const engine = build([['=INDIRECT("A3")'], [null], [9]])

		try {
			engine.insert(0, "rows", 0, 1)
			// A3 now holds nothing: the 9 moved to A4.
			expect(value(engine, "A2")).toBe(0)
		} finally {
			engine.destroy()
		}
	})

	it("reads a new reference when its text changes", () => {
		const engine = build([[1, 2, "A1", "=INDIRECT(C1)"]])

		try {
			expect(value(engine, "D1")).toBe(1)
			set(engine, "C1", "B1")
			expect(value(engine, "D1")).toBe(2)
			set(engine, "B1", 3)
			expect(value(engine, "D1")).toBe(3)
			// A1 is no longer read.
			expect(set(engine, "A1", 8).map(cell => cell.col)).toEqual([0])
		} finally {
			engine.destroy()
		}
	})
})

describe("INDIRECT in a workbook", () => {
	const IMMEDIATE = { sliceMs: 1_000_000, now: () => 0, pause: () => Promise.resolve() }

	async function texts(session: SheetSessionApi, sheet: number, cells: readonly string[]): Promise<string[]> {
		const { tiles } = await session.tiles(sheet, Uint32Array.of(0, 0))
		const tile = tiles[0] as SheetTile

		return cells.map(cell => {
			const position = ((Number(cell.slice(1)) - 1) << 5) | (cell.charCodeAt(0) - 65)

			for (let i = 0; i < tile.count; i++) {
				if (tile.positions[i] !== position) continue

				const text = tile.texts[i] ?? -1

				return text < 0 ? "" : (tile.strings[text] ?? "")
			}

			return ""
		})
	}

	function ok(result: SheetCommandResult): void {
		if (!result.ok) throw new Error(`refused: ${result.refusal.reason}`)
	}

	it("calculates again at once when a sheet is renamed, added or deleted, one that named no sheet too", async () => {
		// "INDIRECT function": a reference to a sheet the workbook lacks is #REF!; INDIRECT is volatile.
		const bytes = buildXlsx({
			sheets: [
				{
					name: "S",
					body: '<sheetData><row r="1"><c r="A1"><f>INDIRECT("T!A1")</f><v>5</v></c><c r="B1" t="e"><f>INDIRECT("\'U\'!A1")</f><v>#REF!</v></c></row></sheetData>'
				},
				{ name: "T", body: '<sheetData><row r="1"><c r="A1"><v>5</v></c></row></sheetData>' }
			],
			application: { name: "Microsoft Excel", version: "16.0300" },
			calcPr: 'calcId="191029"'
		})
		const result = await inThreadSheetHost().open(bytes, { kind: "xlsx", slicing: IMMEDIATE }, () => undefined)

		if (result.kind !== "opened") throw new Error(`expected an open, got ${result.kind}`)

		const { session, info } = result
		const sheet = info.sheets[0]?.id ?? 0
		const other = info.sheets[1]?.id ?? 0

		await session.tiles(sheet, Uint32Array.of(0, 0))
		await session.intent()
		ok(await session.command({ type: "cells.set", sheet, cells: { rows: Uint32Array.of(1), cols: Uint16Array.of(0), inputs: ["1"] } }))
		expect(await texts(session, sheet, ["A1", "B1"])).toEqual(["5", "#REF!"])

		ok(await session.command({ type: "sheet.rename", sheet: other, name: "U" }))
		expect(await texts(session, sheet, ["A1", "B1"])).toEqual(["#REF!", "5"])

		const saved = await session.save()

		if (!saved.ok) throw new Error(saved.reason)
		expect(strFromU8(unzipSync(saved.bytes)["xl/worksheets/sheet1.xml"] as Uint8Array)).toContain(
			'<c r="A1" t="e"><f>INDIRECT("T!A1")</f><v>#REF!</v></c>'
		)

		ok(await session.command({ type: "sheet.add", name: "T" }))
		expect(await texts(session, sheet, ["A1", "B1"])).toEqual(["0", "5"])

		ok(await session.undo())
		expect(await texts(session, sheet, ["A1", "B1"])).toEqual(["#REF!", "5"])
		ok(await session.undo())
		expect(await texts(session, sheet, ["A1", "B1"])).toEqual(["5", "#REF!"])

		ok(await session.command({ type: "sheet.delete", sheet: other }))
		expect(await texts(session, sheet, ["A1", "B1"])).toEqual(["#REF!", "#REF!"])
		session.close()
	})

	it("is taken back whole by undo after a sheet is added or renamed: undone, the save is the source", async () => {
		// B1 reads A1, whose function the app does not calculate (#NAME?): what INDIRECT gives is not stored as a value.
		const bytes = buildXlsx({
			sheets: [
				{
					name: "S",
					body:
						'<sheetData><row r="1"><c r="A1" t="e"><f>_xll.Custom()</f><v>#NAME?</v></c><c r="B1" t="e"><f>INDIRECT("A1")</f><v>#NAME?</v></c>' +
						'<c r="C1"><f>INDIRECT("T!A1")</f><v>5</v></c></row></sheetData>'
				},
				{ name: "T", body: '<sheetData><row r="1"><c r="A1"><v>5</v></c></row></sheetData>' }
			],
			application: { name: "Microsoft Excel", version: "16.0300" },
			calcPr: 'calcId="191029"'
		})
		const result = await inThreadSheetHost().open(bytes, { kind: "xlsx", slicing: IMMEDIATE }, () => undefined)

		if (result.kind !== "opened") throw new Error(`expected an open, got ${result.kind}`)

		const { session, info } = result
		const sheet = info.sheets[0]?.id ?? 0
		const source = await session.save()

		await session.intent()
		ok(await session.command({ type: "cells.set", sheet, cells: { rows: Uint32Array.of(4), cols: Uint16Array.of(0), inputs: ["1"] } }))
		ok(await session.command({ type: "sheet.add", name: "New" }))
		ok(await session.command({ type: "sheet.rename", sheet: info.sheets[1]?.id ?? 0, name: "U" }))
		expect(await texts(session, sheet, ["C1"])).toEqual(["#REF!"])

		for (let i = 0; i < 3; i++) ok(await session.undo())

		const undone = await session.save()

		if (!source.ok || !undone.ok) throw new Error("save refused")
		expect(undone.bytes).toEqual(source.bytes)
		session.close()
	})
})
