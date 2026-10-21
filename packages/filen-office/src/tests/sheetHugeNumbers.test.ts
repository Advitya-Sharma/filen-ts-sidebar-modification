import { describe, expect, it } from "vitest"
import { FormulaEngine } from "../core/sheet/formula/index"
import { cells, firstTile, savedBytes, tileEntry } from "./sheetSessionFixtures"
import { openKeys, sheetBody } from "./sheetKeysFixtures"

// Results at a double's ends: shown to 15 significant digits like any other, a number below Excel's smallest
// (2.2251E-308, "Excel specifications and limits") as 0, never NaN or Infinity in a cell.

describe("results at a double's ends", () => {
	it("keeps the largest and the smallest numbers, and takes what is below Excel's smallest as 0", () => {
		const formulas = ["=A2", "=1/A3", "=1/A2", "=-1/A3", "=A4", "=A2*2"]
		const engine = new FormulaEngine([{ name: "S", rows: [formulas, [Number.MAX_VALUE], [1e300], [4e-320]], formulas: [] }], [])

		try {
			expect(formulas.map((_, col) => engine.result(0, 0, col))).toEqual([
				{ value: Number.MAX_VALUE },
				{ value: 1e-300 },
				{ value: 0 },
				{ value: -1e-300 },
				{ value: 0 },
				{ value: "#NUM!", error: true }
			])
		} finally {
			engine.destroy()
		}
	})

	it("gives huge and tiny results at their 15 significant digits, not one off in the last place", () => {
		const formulas = ["=A2", "=A2*1", "=-A2", "=A3", "=A4/3"]
		const engine = new FormulaEngine([{ name: "S", rows: [formulas, [1e300], [1.23456789012345e-290], [1e300]], formulas: [] }], [])

		try {
			expect(formulas.map((_, col) => engine.result(0, 0, col).value)).toEqual([
				1e300, 1e300, -1e300, 1.23456789012345e-290, 3.33333333333333e299
			])
		} finally {
			engine.destroy()
		}
	})

	it("saves a workbook whose formula divides by 1e308", async () => {
		const { session } = await openKeys({ sheets: [{ name: "S", body: sheetBody([[1, "=1/A1"]]) }] })

		await session.intent()
		expect((await session.command({ type: "cells.set", sheet: 0, cells: cells([[0, 0, "1e308"]]) })).ok).toBe(true)
		expect(tileEntry(await firstTile(session, 0), 0, 1)?.text).toBe("0")
		await expect(savedBytes(session)).resolves.toBeInstanceOf(Uint8Array)
		session.close()
	})
})
