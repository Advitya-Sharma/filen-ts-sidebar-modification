import { describe, expect, it } from "vitest"
import { CYCLE, FormulaEngine, UNPARSEABLE, type EngineSheet } from "../core/sheet/formula/index"
import { budget } from "./timing"

function sheet(name: string, rows: EngineSheet["rows"], formulas: EngineSheet["formulas"] = []): EngineSheet {
	return { name, rows, formulas }
}

function values(cells: ReturnType<FormulaEngine["set"]>): string[] {
	return cells
		.map(cell => `${String(cell.sheet)}:${String(cell.row)},${String(cell.col)}=${String(cell.value)}${cell.error ? "!" : ""}`)
		.sort()
}

describe("FormulaEngine", () => {
	it("calculates across sheets and names, and reports what an edit changes", () => {
		const engine = new FormulaEngine(
			[sheet("Data", [[1, 2, "=A1+B1"]]), sheet("Sum", [["=SUM(Data!A1:C1)*Rate"]])],
			[{ name: "Rate", expression: "2" }]
		)

		try {
			expect(engine.result(0, 0, 2)).toEqual({ value: 3 })
			expect(engine.result(1, 0, 0)).toEqual({ value: 12 })
			expect(values(engine.set(0, [{ row: 0, col: 0, content: 10 }]))).toEqual(["0:0,0=10", "0:0,2=12", "1:0,0=48"])
		} finally {
			engine.destroy()
		}
	})

	it("gives Excel's error texts, and keeps the stored result of a formula it cannot calculate", () => {
		const engine = new FormulaEngine(
			[
				sheet(
					"S",
					[["=1/0", "=NOSUCHFUNCTION(1)", "=B1+1", "=D1"]],
					[
						{ row: 0, col: 1, content: "=NOSUCHFUNCTION(1)", fallback: 41 },
						{ row: 0, col: 3, content: "=D1", fallback: 0 }
					]
				)
			],
			[]
		)

		try {
			expect(engine.result(0, 0, 0)).toEqual({ value: "#DIV/0!", error: true })
			expect(engine.result(0, 0, 1)).toEqual({ value: 41 })
			expect(engine.result(0, 0, 2)).toEqual({ value: 42 })
			expect(FormulaEngine.unsupported(CYCLE)).toBe(true)
			expect(FormulaEngine.unsupported("#NAME?")).toBe(true)
			expect(FormulaEngine.unsupported("#REF!")).toBe(false)
			expect(UNPARSEABLE.startsWith("'")).toBe(true)
		} finally {
			engine.destroy()
		}
	})

	it("moves references with inserted and deleted rows, and follows a renamed sheet", () => {
		const engine = new FormulaEngine([sheet("A", [[1], [2], ["=SUM(A1:A2)"]]), sheet("B", [["=A!A3*10"]])], [])

		try {
			engine.insert(0, "rows", 1, 2)
			expect(engine.result(0, 4, 0)).toEqual({ value: 3 })
			expect(values(engine.set(0, [{ row: 1, col: 0, content: 5 }]))).toEqual(["0:1,0=5", "0:4,0=8", "1:0,0=80"])
			engine.remove(0, "rows", 0, 1)
			expect(engine.result(0, 3, 0)).toEqual({ value: 7 })
			engine.renameSheet(0, "Renamed")
			expect(engine.result(1, 0, 0)).toEqual({ value: 70 })
		} finally {
			engine.destroy()
		}
	})

	it("counts serial dates from 1900 (with its leap day) or from 1904", () => {
		const e1900 = new FormulaEngine([sheet("S", [["=DATE(1900,3,1)", "=DATE(2024,2,29)"]])], [])
		const e1904 = new FormulaEngine([sheet("S", [["=DATE(1904,1,2)"]])], [], true)

		try {
			expect(e1900.result(0, 0, 0)).toEqual({ value: 61 })
			expect(e1900.result(0, 0, 1)).toEqual({ value: 45351 })
			expect(e1904.result(0, 0, 0)).toEqual({ value: 1 })
		} finally {
			e1900.destroy()
			e1904.destroy()
		}
	})

	it("holds a sparse sheet sparsely: one value far right and one far down build in bounded memory", () => {
		const rows: EngineSheet["rows"] = Array.from({ length: 1_048_576 }, () => [])
		const wide = new Array<number | string | null>(16_384).fill(null)

		wide[16_383] = "far right"
		rows[1] = wide
		rows[0] = ["=B1+1", 2]
		rows[1_048_575] = [7]

		const before = process.memoryUsage().heapUsed
		const engine = new FormulaEngine([{ name: "Sparse", rows, formulas: [] }], [])

		try {
			expect(engine.result(0, 0, 0)).toEqual({ value: 3 })
			// One slot per cell of a 16,384 by 1,048,576 rectangle would not fit in memory at all.
			expect(process.memoryUsage().heapUsed - before).toBeLessThan(256 * 1024 * 1024)
		} finally {
			engine.destroy()
		}
	})
})

describe("FormulaEngine at scale", () => {
	// Setting many cells is linear in them: HyperFormula's own batch grouping walks every operation batched so far at
	// each one (16,000 sets took a second, 40,000 several).
	it("sets 40,000 cells over 40,000 dependents in time linear in them", () => {
		const count = 40_000
		const rows: EngineSheet["rows"] = Array.from({ length: count }, (_, row) => [row, `=A${String(row + 1)}*2`])
		const engine = new FormulaEngine([sheet("S", rows)], [])

		try {
			const started = performance.now()
			const changed = engine.set(
				0,
				Array.from({ length: count }, (_, row) => ({ row, col: 0, content: row + 1 }))
			)
			const elapsed = performance.now() - started

			expect(changed).toHaveLength(count * 2)
			expect(engine.result(0, count - 1, 1)).toEqual({ value: count * 2 })
			expect(elapsed).toBeLessThan(budget(2_500))
		} finally {
			engine.destroy()
		}
	})
})
