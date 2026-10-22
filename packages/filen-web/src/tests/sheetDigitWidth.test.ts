// @vitest-environment jsdom

import { describe, expect, it } from "vitest"
import type { SheetEvent, SheetId, SheetLayout, SheetOpened } from "@filen/office/core/sheet"
import { SheetEventRelay } from "@/features/spreadsheet/next/lib/sheetEvents"
import { withWorkbookDigit } from "@/features/spreadsheet/next/lib/sheetDigitWidth"

// A workbook's column-width unit, which the session works out from its default font once asked: Calibri 11's 7 px
// leaves the layouts as they are; another unit has the layouts made before it (the active sheet's, and those held
// events carry) made again before the grid takes them. Nothing is measured on the page.

function layoutOf(sheet: SheetId, defaultSize: number): SheetLayout {
	return { sheet, cols: { defaultSize } } as unknown as SheetLayout
}

function opened(px: number, calls: string[]): SheetOpened {
	return {
		info: { kind: "xlsx", activeSheet: 1, activeLayout: layoutOf(1, 64), styles: { from: 1, styles: [] } },
		session: {
			query: (name: string, request: object) => {
				calls.push(`${name} ${JSON.stringify(request)}`)

				return Promise.resolve({ px, font: {} })
			},
			layout: (sheet: SheetId) => {
				calls.push(`layout ${String(sheet)}`)

				return Promise.resolve(layoutOf(sheet, px * 8 + 8))
			}
		}
	} as unknown as SheetOpened
}

describe("the workbook's digit width", () => {
	it("keeps the layouts of a 7 px workbook", async () => {
		const calls: string[] = []
		const result = opened(7, calls)

		expect(await withWorkbookDigit(result, new SheetEventRelay())).toBe(result)
		expect(calls).toEqual(["digitWidth {}"])
	})

	it("makes the layouts before another unit again", async () => {
		const calls: string[] = []
		const relay = new SheetEventRelay()
		const result = opened(6, calls)

		relay.emit({
			type: "sheetReady",
			sheet: { id: 2, name: "Two", state: "visible", kind: "worksheet", ready: true },
			layout: layoutOf(2, 64)
		})

		const measured = await withWorkbookDigit(result, relay)
		const seen: SheetEvent[] = []

		relay.attach(event => seen.push(event))

		expect(calls).toEqual(["digitWidth {}", "layout 2", "layout 1"])
		expect(measured.info.activeLayout.cols.defaultSize).toBe(56)
		expect(seen[0]?.type === "sheetReady" ? seen[0].layout.cols.defaultSize : 0).toBe(56)
	})

	it("asks nothing for a text file", async () => {
		const calls: string[] = []
		const result = opened(6, calls)
		const text = { ...result, info: { ...result.info, kind: "csv" as const } }

		expect(await withWorkbookDigit(text, new SheetEventRelay())).toBe(text)
		expect(calls).toEqual([])
	})
})
