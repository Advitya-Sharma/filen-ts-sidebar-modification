import { describe, expect, it } from "vitest"
import { FormulaEvaluator, ruleTruth, type FormulaHost } from "../core/sheet/cf/formula/evaluate"
import { parseFormula } from "../core/sheet/cf/formula/syntax"
import { FormulaError, Unsupported, numberText, type FormulaRef, type Scalar } from "../core/sheet/cf/formula/values"

// Rule formulas (conditional formats, validations) read and calculated as Excel does: array semantics, relative
// references moved from the anchor, the functions those rules use, and what is not calculated refused.

function grid(
	cells: Record<string, Scalar>,
	names: Record<string, string> = {},
	sheets: Record<string, number> = { One: 0 },
	extent: { readonly rows: number; readonly cols: number } = { rows: 100, cols: 30 }
): FormulaHost {
	const key = (sheet: number, row: number, col: number): string => `${String(sheet)}:${String(row)}:${String(col)}`
	const map = new Map<string, Scalar>()

	for (const [address, value] of Object.entries(cells)) {
		const match = /^(?:(\w+)!)?([A-Z]+)(\d+)$/u.exec(address)

		if (match === null) throw new Error(address)

		const col = [...(match[2] ?? "")].reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0) - 1

		map.set(key(sheets[match[1] ?? "One"] ?? 0, Number(match[3]) - 1, col), value)
	}

	return {
		value: (sheet, row, col) => map.get(key(sheet, row, col)) ?? null,
		forEach(ref: FormulaRef, visit) {
			const found: [number, number, Scalar][] = []

			for (const [k, value] of map) {
				const [sheet, row, col] = k.split(":").map(Number) as [number, number, number]

				if (
					sheet === ref.sheet &&
					row >= ref.startRow &&
					row <= ref.endRow &&
					col >= ref.startCol &&
					col <= ref.endCol &&
					value !== null
				)
					found.push([row, col, value])
			}

			found.sort((a, b) => a[0] - b[0] || a[1] - b[1])

			for (const [row, col, value] of found) visit(row, col, value)
		},
		extent: () => extent,
		sheetByName: name => sheets[Object.keys(sheets).find(candidate => candidate.toUpperCase() === name.toUpperCase()) ?? ""] ?? null,
		definedName: name => names[name.toUpperCase()] ?? null,
		today: () => 45_000,
		now: () => 45_000.5,
		date1904: false
	}
}

function evaluate(host: FormulaHost, formula: string, at: [number, number] = [0, 0], anchor: [number, number] = [0, 0]): Scalar {
	const evaluator = new FormulaEvaluator(host)

	return evaluator.scalar(evaluator.evaluate(formula, { sheet: 0, anchorRow: anchor[0], anchorCol: anchor[1], row: at[0], col: at[1] }))
}

describe("rule formulas", () => {
	it("moves relative references from the anchor and keeps absolute ones", () => {
		const host = grid({ A1: 1, A2: 5, B2: 7, C1: 10 })

		expect(evaluate(host, "A1>3", [1, 0])).toBe(true)
		expect(evaluate(host, "A1>3", [0, 0])).toBe(false)
		expect(evaluate(host, "$C$1-A1", [1, 1])).toBe(3)
		expect(evaluate(host, "A$1+1", [1, 0])).toBe(2)
		// Relative references wrap round the grid, as Excel's do in conditional formats.
		expect(evaluate(host, "A1", [0, 0], [1, 1])).toBe(null)
	})

	it("compares as Excel does: blank as the other side's empty, text above numbers, text without case", () => {
		const host = grid({ A1: "abc", A2: 0 })

		expect(evaluate(host, '"ABC"=A1')).toBe(true)
		expect(evaluate(host, "A1>99999")).toBe(true)
		expect(evaluate(host, "B1=0")).toBe(true)
		expect(evaluate(host, 'B1=""')).toBe(true)
		expect(evaluate(host, "TRUE>1")).toBe(true)
		expect(evaluate(host, "-2^2")).toBe(4)
		expect(evaluate(host, '"5"+1')).toBe(6)
		expect(evaluate(host, '"x"+1')).toBe(FormulaError.VALUE)
		expect(evaluate(host, "1/0")).toBe(FormulaError.DIV0)
		expect(evaluate(host, "50%")).toBe(0.5)
	})

	it("calculates the functions rules use", () => {
		const host = grid({ A1: "Apples", A2: 3, A3: 4, A4: 5, B1: "  x  y ", C1: 45_000 })

		expect(evaluate(host, 'NOT(ISERROR(SEARCH("pp",A1)))')).toBe(true)
		expect(evaluate(host, 'ISERROR(SEARCH("zz",A1))')).toBe(true)
		expect(evaluate(host, 'SEARCH("p?l",A1)')).toBe(2)
		expect(evaluate(host, 'FIND("a",A1)')).toBe(FormulaError.VALUE)
		expect(evaluate(host, "LEN(TRIM(B1))")).toBe(3)
		expect(evaluate(host, 'LEFT(A1,LEN("app"))="app"')).toBe(true)
		expect(evaluate(host, "MEDIAN(A2:A4)")).toBe(4)
		expect(evaluate(host, "SUM(A1:A4)")).toBe(12)
		expect(evaluate(host, "AVERAGE(A2:A4)")).toBe(4)
		expect(evaluate(host, 'COUNTIF(A1:A4,">3")')).toBe(2)
		expect(evaluate(host, 'COUNTIF(A1:A4,"app*")')).toBe(1)
		expect(evaluate(host, 'COUNTIFS(A2:A4,">2",A2:A4,"<5")')).toBe(2)
		expect(evaluate(host, "COUNTBLANK(A1:A6)")).toBe(2)
		expect(evaluate(host, "VLOOKUP(4,A2:A4,1,FALSE)")).toBe(4)
		expect(evaluate(host, "MATCH(5,A2:A4,0)")).toBe(3)
		expect(evaluate(host, "INDEX(A2:A4,2)")).toBe(4)
		expect(evaluate(host, 'IF(A2>2,"big","small")')).toBe("big")
		expect(evaluate(host, "AND(A2>2,A3>3)")).toBe(true)
		expect(evaluate(host, "OR(A2>9,A3>9)")).toBe(false)
		expect(evaluate(host, "MOD(ROW(),2)=0", [3, 0])).toBe(true)
		expect(evaluate(host, "FLOOR(C1,1)=TODAY()")).toBe(true)
		expect(evaluate(host, "YEAR(C1)")).toBe(2023)
		expect(evaluate(host, "MONTH(EDATE(C1,-1))")).toBe(2)
		expect(evaluate(host, "WEEKDAY(DATE(2023,3,15))")).toBe(4)
		expect(evaluate(host, "ISBLANK(Z9)")).toBe(true)
		expect(evaluate(host, "ADDRESS(2,3,4)")).toBe("C2")
		expect(evaluate(host, 'SUM(INDIRECT("A2:A3"))')).toBe(7)
		expect(evaluate(host, "SUM(OFFSET(A2,1,0,2,1))")).toBe(9)
		expect(evaluate(host, "SUMPRODUCT((A2:A4>3)*1)")).toBe(2)
		expect(evaluate(host, 'TEXT(C1,"yyyy-mm-dd")')).toBe("2023-03-15")
		// support.microsoft.com "PROPER function" examples.
		expect(evaluate(host, 'PROPER("this is a TITLE")')).toBe("This Is A Title")
		expect(evaluate(host, 'PROPER("2-way street")')).toBe("2-Way Street")
		expect(evaluate(host, 'PROPER("76BudGet")')).toBe("76Budget")
	})

	it("evaluates as an array formula, giving the top-left value", () => {
		const host = grid({ A1: "x", B1: "y", C1: "x" })

		expect(evaluate(host, 'OR($A1:$C1="y")')).toBe(true)
		expect(evaluate(host, 'AND($A1:$C1="x")')).toBe(false)
		expect(evaluate(host, "A1:C1")).toBe("x")
	})

	it("reads defined names and other sheets", () => {
		const host = grid({ A1: 2, "Two!B3": 9 }, { LIMIT: "One!$A$1*2" }, { One: 0, Two: 1 })

		expect(evaluate(host, "Limit")).toBe(4)
		expect(evaluate(host, "Two!B3+'Two'!$B$3")).toBe(18)
		expect(evaluate(host, "Missing")).toBe(FormulaError.NAME)
	})

	it("refuses what it does not calculate rather than guessing", () => {
		const host = grid({})

		expect(() => evaluate(host, 'WEBSERVICE("x")')).toThrow(Unsupported)
		// A name that is no function at all is Excel's #NAME? (a misspelt ISERROR).
		expect(evaluate(host, "ISERRROR(A1)")).toBe(FormulaError.NAME)
		expect(evaluate(host, "ISERROR(ISERRROR(A1))")).toBe(true)
		expect(() => evaluate(host, "Table1[Col]")).toThrow(Unsupported)
		expect(() => evaluate(host, "A1 B1")).toThrow(Unsupported)
		expect(() => evaluate(host, "Nowhere!A1")).toThrow(Unsupported)
		expect(() => parseFormula("SUM(")).toThrow(Unsupported)
	})

	it("knows which formulas move with the cell and which are volatile", () => {
		expect(parseFormula("$A$1>5").relative).toBe(false)
		expect(parseFormula("A1>5").relative).toBe(true)
		expect(parseFormula("ROW()=2").relative).toBe(true)
		expect(parseFormula("TODAY()").volatile).toBe(true)
	})

	it("takes only TRUE and non-zero numbers as a rule holding", () => {
		expect(ruleTruth(true)).toBe(true)
		expect(ruleTruth(2)).toBe(true)
		expect(ruleTruth(0)).toBe(false)
		expect(ruleTruth("TRUE")).toBe(false)
		expect(ruleTruth(null)).toBe(false)
	})

	it("writes numbers as Excel's text functions see them", () => {
		expect(numberText(0.1 + 0.2)).toBe("0.3")
		expect(numberText(1e20)).toBe("1E+20")
		expect(numberText(-12.5)).toBe("-12.5")
		expect(numberText(123456789012345680)).toBe("1.23456789012346E+17")
	})

	it("keeps a reference's whole shape past the sheet's used rows, its cells there blank", () => {
		const host = grid({ A1: 1, A2: 2, A3: 3, A4: 4, A5: 5, B3: 1, B4: 1, B5: 1 }, {}, { One: 0 }, { rows: 5, cols: 2 })

		expect(evaluate(host, 'SUMPRODUCT(--($A$1:$A$10=""))')).toBe(5)
		// Pairs A1 with B3: 1 + 2 + 3.
		expect(evaluate(host, "SUMPRODUCT($A$1:$A$8,$B$3:$B$10)")).toBe(6)
		expect(evaluate(host, "SUMPRODUCT($A$1:$A$4,$B$2:$B$5)")).toBe(2 + 3 + 4)
		expect(evaluate(host, "ROWS(--$A$1:$A$10)")).toBe(10)
		expect(evaluate(host, 'COUNTIF($A$1:$A$10,"")')).toBe(5)
		expect(evaluate(host, "COUNT(--($A$1:$A$10=0))")).toBe(10)
		expect(evaluate(host, "OR($A$6:$A$10=0)")).toBe(true)
		// A whole column is read as what it holds and a blank rest, not cell by cell.
		expect(evaluate(host, 'SUMPRODUCT(--($A:$A=""))')).toBe(1_048_576 - 5)
		expect(evaluate(host, 'SUMPRODUCT(($A:$A="")*($B:$B=""))')).toBe(1_048_576 - 5)
		expect(evaluate(host, "INDEX(--$A$1:$A$10,9)")).toBe(0)
		expect(evaluate(host, "INDEX(--$A$1:$A$10,2)")).toBe(2)
	})

	it("compares and subtracts numbers to Excel's 15 significant digits", () => {
		const host = grid({ A1: 0.1, A2: 0.2, A3: 0.3, B1: 1, B2: 1.00000000000001 })

		expect(evaluate(host, "A1+A2=A3")).toBe(true)
		expect(evaluate(host, "A1+A2<>A3")).toBe(false)
		expect(evaluate(host, "A1+A2>A3")).toBe(false)
		expect(evaluate(host, "A1+A2<=A3")).toBe(true)
		expect(evaluate(host, "A1+A2-A3")).toBe(0)
		expect(evaluate(host, "A3-A1-A2")).toBe(0)
		// A difference in the 15th digit is a difference.
		expect(evaluate(host, "B1=B2")).toBe(false)
		expect(evaluate(host, "B2>B1")).toBe(true)
		expect(evaluate(host, 'COUNTIF(A1:A3,"="&(A1+A2))')).toBe(1)
	})
})
