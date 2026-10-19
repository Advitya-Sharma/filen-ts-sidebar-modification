import { describe, expect, it } from "vitest"
import {
	allFormulajsCases,
	conversionOffice,
	convertedWorkbooks,
	formulajsCases,
	libreOfficeCases,
	poiCases,
	type FunctionCase
} from "./functionCases"
import { suiteMissing } from "./functionSuites"
import { sheetOffices } from "./libreOffice"

// The third-party function suites as { function, inputs, expected } cases, generated here from the fetched
// files (pnpm exec tsx src/tests/oracles/sheet/functionSuites.ts --fetch). Each suite skips until fetched;
// LibreOffice's and POI's also need FILEN_OFFICE_SOFFICE (they are converted to xlsx), and
// FILEN_OFFICE_SHEET_SUITES_LIMIT converts only the first n LibreOffice files.
//
// Licences of the generated cases' sources: LibreOffice sc/qa/unit/data/functions, MPL-2.0
// (https://www.mozilla.org/MPL/2.0/); formulajs test/, MIT, Copyright (c) 2014 Sutoiku, Inc.
// (https://github.com/formulajs/formulajs/blob/master/LICENSE); Apache POI test-data/spreadsheet, Apache-2.0
// (https://www.apache.org/licenses/LICENSE-2.0). Nothing derived from them is committed.

const OFFICE = conversionOffice(sheetOffices())
const LIMIT = Number(process.env["FILEN_OFFICE_SHEET_SUITES_LIMIT"] ?? Number.POSITIVE_INFINITY)

function wellFormed(cases: readonly FunctionCase[]): void {
	for (const item of cases) {
		expect(item.function, item.id).toMatch(/^([A-Z][A-Z0-9._]*|\(operators\))$/)

		if (item.inputs.kind === "cell") expect(item.inputs.formula, item.id).not.toBe("")
	}
}

function summary(cases: readonly FunctionCase[]): string {
	return `${cases.length} cases over ${new Set(cases.map(item => item.function)).size} functions`
}

describe("formulajs test parsing", () => {
	it("reads literal assertions and skips the rest", () => {
		const source = [
			"import * as math from '../src/math-trig.js'",
			"import * as error from '../src/utils/error.js'",
			"it('x', () => {",
			"  expect(math.ABS(-1)).to.equal(1)",
			"  expect(statistical.CHISQ.DIST.RT(3, [[1, 2], [3, 4]], null)).to.approximately(0.5, 1e-9)",
			"  expect(math.ACOS(100)).to.equal(error.num)",
			"  expect(math.ACOS()).to.equal(math.ACOS(0))",
			"  expect(math.ABS(someVariable)).to.equal(1)",
			"  expect(text.SPLIT('a b', ' ')).to.deep.equal(['a', 'b'])",
			"})"
		].join("\n")
		const { cases, tally } = formulajsCases(source, "t.js")

		expect(tally).toEqual({ cases: 4, skipped: 2 })
		expect(cases[0]).toEqual({
			suite: "formulajs",
			id: "t.js:4",
			function: "ABS",
			inputs: { kind: "args", args: [{ type: "n", value: -1 }] },
			expected: { type: "n", value: 1 }
		})
		expect(cases[1]?.function).toBe("CHISQ.DIST.RT")
		expect(cases[1]?.tolerance).toBe(1e-9)
		expect(cases[1]?.inputs).toEqual({
			kind: "args",
			args: [
				{ type: "n", value: 3 },
				{
					type: "array",
					rows: [
						[
							{ type: "n", value: 1 },
							{ type: "n", value: 2 }
						],
						[
							{ type: "n", value: 3 },
							{ type: "n", value: 4 }
						]
					]
				},
				{ type: "blank" }
			]
		})
		expect(cases[2]?.expected).toEqual({ type: "e", value: "#NUM!" })
		expect(cases[3]?.expected).toEqual({
			type: "array",
			rows: [
				[
					{ type: "s", value: "a" },
					{ type: "s", value: "b" }
				]
			]
		})
	})
})

describe.skipIf(suiteMissing("formulajs") !== undefined)("the formulajs suite", () => {
	it("generates cases", () => {
		const { cases, tally } = allFormulajsCases()

		wellFormed(cases)
		expect(cases.length).toBeGreaterThan(2000)
		expect(cases.some(item => item.function === "ABS" && item.expected.type === "n" && item.expected.value === 1)).toBe(true)
		console.log(`function suites: formulajs ${summary(cases)}, ${tally.skipped} assertions not literal`)
	})
})

describe.skipIf(suiteMissing("poi") !== undefined || OFFICE === undefined)("the POI suite", () => {
	it("generates cases", { timeout: 600_000 }, async () => {
		const { workbooks, failures } = await convertedWorkbooks("poi", OFFICE as NonNullable<typeof OFFICE>)
		const cases: FunctionCase[] = []
		let skipped = 0

		for (const [source, xlsx] of workbooks) {
			const result = await poiCases(source, xlsx)

			expect(result.cases.length, source).toBeGreaterThan(0)
			cases.push(...result.cases)
			skipped += result.tally.skipped
		}

		wellFormed(cases)
		expect(failures).toEqual([])
		expect(cases.some(item => item.function === "MATCH")).toBe(true)
		console.log(
			`function suites: poi ${summary(cases)} from ${workbooks.length} workbooks (${OFFICE?.name}), ${skipped} rows without an expected value`
		)
	})
})

describe.skipIf(suiteMissing("libreoffice") !== undefined || OFFICE === undefined)("the LibreOffice suite", () => {
	it("generates cases", { timeout: 1_800_000 }, async () => {
		const { workbooks, failures } = await convertedWorkbooks("libreoffice", OFFICE as NonNullable<typeof OFFICE>, LIMIT)
		const cases: FunctionCase[] = []
		const empty: string[] = []
		let skipped = 0

		for (const [source, xlsx] of workbooks) {
			const result = await libreOfficeCases(source, xlsx)

			if (result.cases.length === 0) empty.push(source)

			cases.push(...result.cases)
			skipped += result.tally.skipped
		}

		wellFormed(cases)
		expect(failures).toEqual([])
		expect(cases.length).toBeGreaterThan(workbooks.length)

		const average = cases.find(item => item.id === "statistical/fods/average.fods Sheet2!A2")

		if (average !== undefined) expect(average.expected).toEqual({ type: "n", value: 8 })

		console.log(
			`function suites: libreoffice ${summary(cases)} from ${workbooks.length} workbooks (${OFFICE?.name}), ${skipped} rows without an expected value; no cases in ${empty.length}: ${empty.join(", ")}`
		)
	})
})
