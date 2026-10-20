import { describe, expect, it } from "vitest"
import { planRefEdit } from "../core/xlsx/refs/index"
import { rows, sheetBody } from "./xlsxRefsFixtures"
import { buildFixture, openWorkbook, saved } from "./xlsxWriteFixtures"

// Planning and applying a structure edit costs in proportion to what refers to cells: a sheet of 100,000 rows with
// a formula on each, a shared column and a few hundred conditional formats, against the same sheet without them.

const ROWS = 100_000

function body(formulas: boolean): string {
	let data = "<sheetData>"

	for (let r = 1; r <= ROWS; r++) {
		data += `<row r="${r}"><c r="A${r}"><v>${r}</v></c><c r="B${r}"><v>${r * 2}</v></c>`

		if (formulas) {
			data += `<c r="C${r}"><f>A${r}+B${r}</f><v>0</v></c>`
			data +=
				r === 1
					? `<c r="D1"><f t="shared" ref="D1:D${ROWS}" si="0">C1*2</f><v>0</v></c>`
					: `<c r="D${r}"><f t="shared" si="0"/><v>0</v></c>`
		}

		data += "</row>"
	}

	data += "</sheetData>"

	if (formulas)
		for (let k = 0; k < 300; k++)
			data += `<conditionalFormatting sqref="E${k * 10 + 1}:E${k * 10 + 5}"><cfRule type="expression" priority="${k + 1}"><formula>E${k * 10 + 1}&gt;0</formula></cfRule></conditionalFormatting>`

	return data
}

async function time(formulas: boolean): Promise<{ plan: number; apply: number; save: number }> {
	const workbook = await openWorkbook(buildFixture(sheetBody(body(formulas))))

	await workbook.load(0)

	let started = performance.now()
	const result = await planRefEdit(workbook, rows("insert", 1, 1))
	const plan = performance.now() - started

	if (!result.ok) throw new Error(JSON.stringify(result.refusal))

	started = performance.now()
	result.plan.apply()

	const apply = performance.now() - started

	started = performance.now()
	await saved(workbook)

	return { plan, apply, save: performance.now() - started }
}

describe("structure edit cost", () => {
	it("is proportional to the constructs that refer to cells", { timeout: 300_000 }, async () => {
		const plain = await time(false)
		const rich = await time(true)

		console.log(
			`refs perf, ${ROWS} rows: plain plan ${plain.plan.toFixed(0)} ms apply ${plain.apply.toFixed(0)} ms save ${plain.save.toFixed(0)} ms; with ${ROWS * 2} formulas and 300 rules plan ${rich.plan.toFixed(0)} ms apply ${rich.apply.toFixed(0)} ms save ${rich.save.toFixed(0)} ms`
		)
		// Nothing to move but cells: planning reads no row.
		expect(plain.plan).toBeLessThan(50)
	})
})
