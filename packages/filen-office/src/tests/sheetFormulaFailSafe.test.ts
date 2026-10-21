import { FunctionArgumentType, FunctionPlugin, HyperFormula, type ImplementedFunctions } from "hyperformula"
import { describe, expect, it } from "vitest"
import { FormulaEngine, type EngineSheet } from "../core/sheet/formula/index"
import { guardFunctionPlugins } from "../core/sheet/formula/functions/plugin"
import { cells, firstTile, openFixture, tileEntry } from "./sheetSessionFixtures"

// A function failing in JavaScript gives its own cell an error and leaves every other cell, and every later edit,
// calculating: the engine would let the exception escape and fail the workbook's edits from then on.

class Failing extends FunctionPlugin {
	static override implementedFunctions: ImplementedFunctions = {
		"TEST.RANGE": { method: "range", parameters: [{ argumentType: FunctionArgumentType.NUMBER }] },
		"TEST.TYPE": { method: "type", parameters: [{ argumentType: FunctionArgumentType.NUMBER }] }
	}

	range(): unknown {
		return (1).toFixed(500)
	}

	type(): unknown {
		throw new TypeError("not a function")
	}
}

function engine(rows: EngineSheet["rows"]): FormulaEngine {
	return new FormulaEngine([{ name: "S", rows, formulas: [] }], [])
}

describe("functions that fail in JavaScript", () => {
	it("give #NUM! for a number out of range and #VALUE! otherwise, the other cells calculating", () => {
		HyperFormula.registerFunctionPlugin(Failing, { enGB: { "TEST.RANGE": "TEST.RANGE", "TEST.TYPE": "TEST.TYPE" } })
		guardFunctionPlugins()

		const built = engine([[1, "=TEST.RANGE(A1)", "=TEST.TYPE(A1)", "=A1*2"]])

		try {
			expect(built.result(0, 0, 1)).toEqual({ value: "#NUM!", error: true })
			expect(built.result(0, 0, 2)).toEqual({ value: "#VALUE!", error: true })
			expect(built.result(0, 0, 3)).toEqual({ value: 2 })
			expect(built.set(0, [{ row: 0, col: 0, content: 5 }])).toContainEqual({ sheet: 0, row: 0, col: 3, value: 10 })
		} finally {
			built.destroy()
		}
	})

	it("takes the largest and smallest of 200,000 numbers in SUBTOTAL and AGGREGATE", () => {
		const rows: (number | string)[][] = []

		for (let row = 0; row < 200_000; row++) rows.push([row])

		rows[0]?.push("=SUBTOTAL(4,A1:A200000)", "=SUBTOTAL(105,A1:A200000)", "=AGGREGATE(4,4,A1:A200000)", "=AGGREGATE(5,4,A1:A200000)")

		const built = engine(rows)

		try {
			expect([1, 2, 3, 4].map(col => built.result(0, 0, col))).toEqual([
				{ value: 199_999 },
				{ value: 0 },
				{ value: 199_999 },
				{ value: 0 }
			])
		} finally {
			built.destroy()
		}
	})
})

// support.microsoft.com "FIXED function": decimals can be as large as 127, and numbers never have more than 15
// significant digits. LibreOffice rounds the same way (FIXED(2.675;2) is 2.68, FIXED(123456789012345678;0) is
// 123,456,789,012,346,000).
describe("FIXED and DOLLAR", () => {
	it("give up to 127 decimals, at 15 significant digits, and refuse 128", () => {
		const formulas = ["=FIXED(1,101)", "=DOLLAR(1,127)", "=FIXED(1/3,20)", "=FIXED(1,128)", "=DOLLAR(-1234.567,2)"]
		const built = engine([formulas])

		try {
			expect(formulas.map((_, col) => built.result(0, 0, col))).toEqual([
				{ value: `1.${"0".repeat(101)}` },
				{ value: `$1.${"0".repeat(127)}` },
				{ value: `0.333333333333333${"0".repeat(5)}` },
				{ value: "#VALUE!", error: true },
				{ value: "($1,234.57)" }
			])
		} finally {
			built.destroy()
		}
	})

	it("write numbers from 1e21 with every digit, rounding the decimal digits half up", () => {
		const formulas = [
			"=FIXED(10^21,0)",
			"=FIXED(2.675,2)",
			"=FIXED(123456789012345678,0)",
			"=FIXED(5,-1)",
			"=FIXED(-0.004,2)",
			"=FIXED(9.995,2,1)"
		]
		const built = engine([formulas])

		try {
			expect(formulas.map((_, col) => built.result(0, 0, col).value)).toEqual([
				"1,000,000,000,000,000,000,000",
				"2.68",
				"123,456,789,012,346,000",
				"10",
				"0.00",
				"10.00"
			])
		} finally {
			built.destroy()
		}
	})

	it("leave the workbook's later edits working after a formula with 101 decimals is typed", async () => {
		const opened = await openFixture({
			sheets: [{ name: "S", body: '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row></sheetData>' }]
		})
		const { session } = opened
		const sheet = opened.info.sheets[0]?.id ?? 0

		await session.intent()

		expect((await session.command({ type: "cells.set", sheet, cells: cells([[2, 0, "=FIXED(1,101)"]]) })).ok).toBe(true)
		expect((await session.command({ type: "cells.set", sheet, cells: cells([[0, 0, "5"]]) })).ok).toBe(true)

		const tile = await firstTile(session, sheet)

		expect(tileEntry(tile, 0, 1)?.text).toBe("10")
		expect(tileEntry(tile, 2, 0)?.text).toBe(`1.${"0".repeat(101)}`)
		session.close()
	})
})
