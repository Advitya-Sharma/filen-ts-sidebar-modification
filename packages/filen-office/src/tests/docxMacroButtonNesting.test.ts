import { describe, expect, it } from "vitest"
import { buildDocx } from "./docxSynthetic"
import { layOut } from "./layoutFixtures"
import { cpuMs } from "./timing"

// A MACROBUTTON's display text reads the fields nested in its instruction once each: as many nested fields cost
// about what they cost under any other field.

const char = (type: string): string => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`
const instr = (text: string): string => `<w:r><w:instrText xml:space="preserve">${text}</w:instrText></w:r>`
const run = (text: string): string => `<w:r><w:t>${text}</w:t></w:r>`

function body(field: string, nested: number): string {
	return `<w:p>${char("begin")}${instr(field)}${(char("begin") + instr("x") + char("separate") + run("y")).repeat(nested)}${char("end").repeat(nested)}${char("end")}</w:p>`
}

async function layoutMs(xml: string): Promise<number> {
	let least = Infinity

	for (let round = 0; round < 2; round++) {
		const bytes = buildDocx({ body: xml })
		const start = cpuMs()

		await layOut(bytes)
		least = Math.min(least, cpuMs() - start)
	}

	return least
}

describe("MACROBUTTON with nested fields", () => {
	it("lays out 30,000 nested result fields in time linear in their number", { timeout: 120_000 }, async () => {
		const macro = await layoutMs(body(" MACROBUTTON a ", 30_000))
		const page = await layoutMs(body(" PAGE ", 30_000))

		// Scanning the nested fields per text item made this about ten times the PAGE case.
		expect(macro).toBeLessThan(page * 3)
	})
})
