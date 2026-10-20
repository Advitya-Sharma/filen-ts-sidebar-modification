import { describe, expect, it } from "vitest"
import { bytesSource } from "@filen/office/core/opc"
import { openXlsx } from "../core/xlsx/read/index"
import { buildWorkbook, DEFAULT_STYLES } from "./xlsxReadFixtures"

// workbook.xml and styles.xml are parsed into trees: their elements are counted first, so a part of tiny elements
// is refused before a node is built for each.

describe("xlsx reader: meta parts", () => {
	it("refuses a workbook part with more elements than the cap, nested in one section", async () => {
		const workbook = buildWorkbook({ sheets: [{}], workbookTail: `<extLst><ext uri="x">${"<a/>".repeat(200)}</ext></extLst>` })
		const capped = await openXlsx(bytesSource(workbook), { readLimits: { maxMetaElements: 100 } })
		const open = await openXlsx(bytesSource(workbook), { readLimits: { maxMetaElements: 1000 } })

		expect(capped.ok ? "opened" : capped.reason).toBe("tooLarge")
		expect(open.ok).toBe(true)
	})

	it("refuses a styles part with more elements than the cap", async () => {
		const workbook = buildWorkbook({
			sheets: [{}],
			styles: `${DEFAULT_STYLES}<extLst><ext uri="x">${"<a/>".repeat(200)}</ext></extLst>`
		})
		const capped = await openXlsx(bytesSource(workbook), { readLimits: { maxMetaElements: 150 } })

		expect(capped.ok ? "opened" : capped.reason).toBe("tooLarge")
	})
})
