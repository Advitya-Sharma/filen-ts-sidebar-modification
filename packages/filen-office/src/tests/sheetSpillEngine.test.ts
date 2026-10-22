import { describe, expect, it } from "vitest"
import { FormulaEngine, type EngineCell, type EngineSheet, type RecalculatedCell } from "../core/sheet/formula/index"
import { engineCellFormula, engineSpillFormula, storedFormula } from "../core/sheet/refs/formulaRefs"

// Dynamic arrays in the formula engine (support.microsoft.com "Dynamic array formulas and spilled array behavior"): a
// formula's array spills over the cells beside and below it, reported for the workbook to write, and a cell in the way
// makes it #SPILL!. "=..." is an ordinary formula, "#=..." a dynamic-array one.

type Typed = string | number | boolean | null

function content(cell: Typed): EngineCell["content"] {
	if (typeof cell !== "string") return cell
	if (cell.startsWith("#=")) return engineSpillFormula(storedFormula(cell.slice(2))) ?? ""
	if (cell.startsWith("=")) return engineCellFormula(storedFormula(cell.slice(1))) ?? ""

	return cell
}

// `spills`: the dynamic arrays and what the file has each spill over (its values already in `rows`).
function build(rows: Typed[][], spills: { cell: string; rows: number; cols: number }[] = []): FormulaEngine {
	const sheet: EngineSheet = {
		name: "S",
		rows: rows.map(row => row.map(content)),
		formulas: spills.map(({ cell, rows: height, cols }) => {
			const { row, col } = at(cell)

			return { row, col, content: content(rows[row]?.[col] ?? null), spill: { rows: height, cols } }
		})
	}

	return new FormulaEngine([sheet], [])
}

function at(a1: string): { row: number; col: number } {
	return { row: Number(a1.slice(1)) - 1, col: a1.charCodeAt(0) - 65 }
}

function value(engine: FormulaEngine, a1: string): unknown {
	const { row, col } = at(a1)
	const result = engine.result(0, row, col)

	return result.error === true ? `!${String(result.value)}` : result.value
}

function set(engine: FormulaEngine, a1: string, typed: Typed, spill?: { rows: number; cols: number }): RecalculatedCell[] {
	return engine.set(0, [{ ...at(a1), content: content(typed), ...(spill !== undefined && { spill }) }])
}

function shown(changes: readonly RecalculatedCell[]): string[] {
	return changes
		.map(cell => {
			const name = `${String.fromCharCode(65 + cell.col)}${String(cell.row + 1)}`
			const extra =
				cell.spill !== undefined
					? ` spill ${String(cell.spill.rows)}x${String(cell.spill.cols)}`
					: cell.spilled === true
						? " spilled"
						: ""

			return `${name}=${String(cell.value)}${extra}`
		})
		.sort()
}

describe("dynamic arrays", () => {
	it("spill a new array formula over the cells below, and report them", () => {
		const engine = build([[1], [2], [3]])

		try {
			expect(shown(set(engine, "B1", "#=A1:A3*10", { rows: 1, cols: 1 }))).toEqual([
				"B1=10 spill 3x1",
				"B2=20 spilled",
				"B3=30 spilled"
			])
			expect([value(engine, "B2"), value(engine, "B3")]).toEqual([20, 30])
		} finally {
			engine.destroy()
		}
	})

	it("grow and shrink with their array, clearing what they no longer reach", () => {
		const engine = build(
			[
				[3, "#=SEQUENCE(A1)", null],
				[null, 2, null],
				[null, 3, null]
			],
			[{ cell: "B1", rows: 3, cols: 1 }]
		)

		try {
			expect(shown(set(engine, "A1", 2))).toEqual(["A1=2", "B1=1 spill 2x1", "B3=null spilled"])
			expect(shown(set(engine, "A1", 4))).toEqual(["A1=4", "B1=1 spill 4x1", "B3=3 spilled", "B4=4 spilled"])
		} finally {
			engine.destroy()
		}
	})

	it("are #SPILL! while a cell is in the way, and spill again once it is cleared", () => {
		const engine = build(
			[
				[1, "#=A1:A3"],
				[2, 2],
				[3, 3]
			],
			[{ cell: "B1", rows: 3, cols: 1 }]
		)

		try {
			// A value typed over one of its values blocks it; what it spilled elsewhere is cleared.
			expect(shown(set(engine, "B2", "x"))).toEqual(["B1=#SPILL! spill 1x1", "B2=x", "B3=null spilled"])
			expect(shown(set(engine, "B2", null))).toEqual(["B1=1 spill 3x1", "B2=2 spilled", "B3=3 spilled"])
			// A value beside it, where it would grow.
			set(engine, "B4", 9)
			expect(shown(set(engine, "A4", 4))).toEqual(["A4=4"])
		} finally {
			engine.destroy()
		}
	})

	it("leave the file's spilled values as they are until they calculate, which their dependents read", () => {
		const engine = build(
			[
				[1, "#=A1:A2*2", "=SUM(B1:B2)"],
				[5, 99]
			],
			[{ cell: "B1", rows: 2, cols: 1 }]
		)

		try {
			expect([value(engine, "B1"), value(engine, "B2"), value(engine, "C1")]).toEqual([2, 99, 101])
			expect(shown(set(engine, "A2", 6))).toEqual(["A2=6", "B1=2 spill 2x1", "B2=12 spilled", "C1=14"])
		} finally {
			engine.destroy()
		}
	})

	it("clear what they spilled when the formula goes", () => {
		const engine = build(
			[
				[1, "#=A1:A2"],
				[2, 2]
			],
			[{ cell: "B1", rows: 2, cols: 1 }]
		)

		try {
			expect(shown(set(engine, "B1", 7))).toEqual(["B1=7", "B2=null spilled"])
		} finally {
			engine.destroy()
		}
	})

	it("read a spill range (A1#) as the cells it spills over now", () => {
		const engine = build(
			[
				[2, "#=SEQUENCE(A1)", "=SUM(_xlfn.ANCHORARRAY(B1))", "=ROWS(_xlfn.ANCHORARRAY(B1))"],
				[null, 2]
			],
			[{ cell: "B1", rows: 2, cols: 1 }]
		)

		try {
			expect([value(engine, "C1"), value(engine, "D1")]).toEqual([3, 2])
			set(engine, "A1", 4)
			expect([value(engine, "C1"), value(engine, "D1")]).toEqual([10, 4])
		} finally {
			engine.destroy()
		}
	})

	it("move with their cell when rows are inserted above or through them", () => {
		const engine = build(
			[
				[1, "#=A1:A2"],
				[2, 2]
			],
			[{ cell: "B1", rows: 2, cols: 1 }]
		)

		try {
			engine.insert(0, "rows", 0, 1)
			expect(shown(set(engine, "A2", 5))).toEqual(["A2=5", "B2=5 spill 2x1"])
		} finally {
			engine.destroy()
		}
	})
})
