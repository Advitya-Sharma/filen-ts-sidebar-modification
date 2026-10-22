import { describe, expect, it } from "vitest"
import { FormulaEngine, type EngineSheet } from "../core/sheet/formula/index"
import { engineCellFormula, storedFormula } from "../core/sheet/refs/formulaRefs"

// FILTER and IFS as Excel calculates them, where the engine's own differ (each case cites its page on
// support.microsoft.com).

type Typed = string | number | boolean | null

function calculate(rows: Typed[][]): Record<string, unknown> {
	const sheets: EngineSheet[] = [
		{
			name: "S",
			rows: rows.map(row =>
				row.map(cell =>
					typeof cell === "string" && cell.startsWith("=") ? (engineCellFormula(storedFormula(cell.slice(1))) ?? "") : cell
				)
			),
			formulas: []
		}
	]
	const engine = new FormulaEngine(sheets, [])
	const out: Record<string, unknown> = {}

	try {
		rows.forEach((row, r) =>
			row.forEach((cell, c) => {
				if (typeof cell !== "string" || !cell.startsWith("=")) return

				const result = engine.result(0, r, c)

				out[`${String.fromCharCode(65 + c)}${String(r + 1)}`] = result.error === true ? `!${String(result.value)}` : result.value
			})
		)
	} finally {
		engine.destroy()
	}

	return out
}

describe("FILTER", () => {
	it("keeps the rows or columns include marks, if_empty or #CALC! when none", () => {
		// "FILTER function": =FILTER(A5:D20,C5:C20=H2,"") returns the records for Apple; "If the include argument
		// returns no results ... FILTER will return a #CALC! error" unless if_empty is given.
		expect(
			calculate([
				["a", 1, 10, "=INDEX(FILTER(A1:C3,B1:B3>1),2,3)", '=FILTER(A1:C3,B1:B3>5,"No results")', "=FILTER(A1:C3,B1:B3>5)"],
				["b", 2, 20, "=ROWS(FILTER(A1:C3,B1:B3>1))", "=INDEX(FILTER(A1:C3,{1,0,1}),3,2)", "=FILTER(A1:C3,{1,0})"],
				["c", 3, 30, '=FILTER(A1:C3,{1;0;"x"})', "=SUM(FILTER(C1:C3,B1:B3<>2))", "=FILTER(A1:C3,B1:B3/0)"]
			])
		).toEqual({ D1: 30, E1: "No results", F1: "!#CALC!", D2: 2, E2: 30, F2: "!#VALUE!", D3: "!#VALUE!", E3: 40, F3: "!#DIV/0!" })
	})
})

describe("IFS", () => {
	it("gives the value of the first test that is TRUE, never calculating the rest", () => {
		// "IFS function": "If no TRUE conditions are found, this function returns #N/A error."
		expect(
			calculate([[85, '=IFS(A1>89,"A",A1>79,"B",A1>69,"C")', '=IFS(A1>0,"first",NA(),"never")', "=IFS(A1>100,1)", '=IFS("x",1)']])
		).toEqual({ B1: "B", C1: "first", D1: "!#N/A", E1: "!#VALUE!" })
	})
})

describe("approximate lookups", () => {
	it("search the values of the key's class alone, as Excel compares", () => {
		// "HLOOKUP function": with range_lookup TRUE "the next largest value that is less than lookup_value is returned";
		// a number in a row of text is no text to compare with, so the search does not stop at it.
		expect(
			calculate([
				["As New", "CCGT", "Oxy", 0, 0, 10, "a"],
				[1, 0.982, 0.9, 1, 1, 20, "b"],
				[null, null, null, null, null, 30, "c"],
				['=HLOOKUP("ccgt",A1:E2,2)', '=MATCH("D",A1:E1)', "=VLOOKUP(25,F1:G3,2)", "=MATCH(25,F1:F3,-1)", "=VLOOKUP(5,F1:G3,2)"]
			])
		).toEqual({ A4: 0.982, B4: 2, C4: "b", D4: "!#N/A", E4: "!#N/A" })
	})
})

describe("criteria naming an error", () => {
	it("are no error of the count's own", () => {
		// "COUNTIF function": criteria may be any value; an error there is what the cells are compared with.
		expect(calculate([[1, 2, "=COUNTIF(A1:B1,#REF!)", "=SUMIF(A1:B1,#N/A,A1:B1)"]])).toEqual({ C1: 0, D1: 0 })
	})
})
