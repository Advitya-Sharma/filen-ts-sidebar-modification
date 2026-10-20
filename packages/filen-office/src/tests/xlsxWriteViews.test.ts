import { existsSync, readFileSync, readdirSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { bytesSource } from "../core/opc/index"
import { openXlsx, type XlsxWorkbook } from "../core/xlsx/read/index"
import type { XlsxSheetModel, XlsxSheetView } from "../core/xlsx/model/types"
import { saveXlsx } from "../core/xlsx/write/save"
import { CORPUS_DIR } from "./corpusDir"
import { buildFixture, changedEntries, openWorkbook, parts, saved } from "./xlsxWriteFixtures"

// Sheet views through the writer: the first view's pane and active selection follow the model, and everything the
// model does not hold (a pane split but not frozen, the other panes' selections, further views) is kept as written.
// The save's check reads the views back the way the reader does.

function sheet(views: string): Uint8Array {
	return buildFixture({ sheets: [{ name: "Data", body: `<sheetViews>${views}</sheetViews><sheetData/>` }] })
}

async function edit(source: Uint8Array, change: (view: XlsxSheetView) => XlsxSheetView): Promise<{ text: string; workbook: XlsxWorkbook }> {
	const workbook = await openWorkbook(source)
	const model = await workbook.load(0)

	model.view = change(model.view)
	workbook.changes.sheetSection(0, "sheetViews")

	const text = parts(await saved(workbook))["xl/worksheets/sheet1.xml"] ?? ""

	return { text, workbook }
}

function views(text: string): string {
	return /<sheetViews>.*<\/sheetViews>/s.exec(text)?.[0] ?? ""
}

const SPLIT =
	'<sheetView tabSelected="1" workbookViewId="0"><pane xSplit="16230" ySplit="3840" topLeftCell="AQ1000" activePane="bottomLeft"/>' +
	'<selection activeCell="A6" sqref="A6:A21"/><selection pane="topRight" activeCell="AQ16" sqref="AQ16"/>' +
	'<selection pane="bottomLeft" activeCell="L1008" sqref="L1008"/><selection pane="bottomRight" activeCell="AI1443" sqref="AI1443"/></sheetView>'
const FROZEN =
	'<sheetView workbookViewId="0"><pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>' +
	'<selection pane="topRight" activeCell="B1" sqref="B1"/><selection pane="bottomLeft" activeCell="A3" sqref="A3"/>' +
	'<selection pane="bottomRight" activeCell="C5" sqref="C5:D6"/></sheetView>'

describe("sheet views", () => {
	it("keep a split pane and the other panes' selections when the active selection moves", async () => {
		const { text } = await edit(sheet(SPLIT), view => ({
			...view,
			activeRow: 1009,
			activeCol: 11,
			selection: [{ startRow: 1009, startCol: 11, endRow: 1009, endCol: 11 }]
		}))

		expect(views(text)).toBe(
			`<sheetViews>${SPLIT.replace('pane="bottomLeft" activeCell="L1008" sqref="L1008"', 'pane="bottomLeft" activeCell="L1010" sqref="L1010"')}</sheetViews>`
		)
	})

	it("move a frozen split in place, keeping every selection", async () => {
		const { text } = await edit(sheet(FROZEN), view => ({ ...view, frozen: { rows: 3, cols: 1, topLeftRow: 3, topLeftCol: 1 } }))

		expect(views(text)).toBe(`<sheetViews>${FROZEN.replace('ySplit="2" topLeftCell="B3"', 'ySplit="3" topLeftCell="B4"')}</sheetViews>`)
	})

	it("unfreeze: the pane goes with its panes' selections, the top-left one stays", async () => {
		const { text } = await edit(sheet(FROZEN), view => ({ ...view, frozen: null }))

		expect(views(text)).toBe(
			'<sheetViews><sheetView workbookViewId="0"><selection activeCell="C5" sqref="C5:D6"/></sheetView></sheetViews>'
		)
	})

	it("freeze a view without a pane", async () => {
		const { text } = await edit(sheet('<sheetView workbookViewId="0"><selection activeCell="B2" sqref="B2"/></sheetView>'), view => ({
			...view,
			frozen: { rows: 1, cols: 0, topLeftRow: 1, topLeftCol: 0 }
		}))

		expect(views(text)).toBe(
			'<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="B2" sqref="B2"/>' +
				'<selection activeCell="B2" sqref="B2"/></sheetView></sheetViews>'
		)
	})

	it("keep further views as written", async () => {
		const second =
			'<sheetView workbookViewId="1"><pane ySplit="4" topLeftCell="A5" state="frozen"/><selection activeCell="A9" sqref="A9"/></sheetView>'
		const { text } = await edit(sheet(`<sheetView workbookViewId="0"/>${second}`), view => ({ ...view, zoom: 150 }))

		expect(views(text)).toBe(`<sheetViews><sheetView workbookViewId="0" zoomScale="150"/>${second}</sheetViews>`)
	})
})

const KINDS = ["xlsx", "xlsm", "xltx", "xltm"]
const FILES = KINDS.flatMap(kind =>
	existsSync(join(CORPUS_DIR, kind)) ? readdirSync(join(CORPUS_DIR, kind)).map(name => join(kind, name)) : []
)

function elements(text: string): string[] {
	return [...views(text).matchAll(/<(?:\w+:)?(sheetView|pane|selection)\b[^>]*>/g)].map(match => match[1] ?? "")
}

// Every element of `a`, in order, in `b` (a selection the active pane lacked may be added).
function subsequence(a: readonly string[], b: readonly string[]): boolean {
	let k = 0

	for (const item of b) if (item === a[k]) k++

	return k === a.length
}

describe.skipIf(FILES.length === 0)("sheet views over the corpus", () => {
	it(
		"save as written when unchanged, and keep every view, pane and selection when the active cell moves",
		{ timeout: 600_000 },
		async () => {
			const failures: string[] = []
			let checked = 0

			for (const file of FILES) {
				const source = new Uint8Array(readFileSync(join(CORPUS_DIR, file)))

				for (const moved of [false, true]) {
					const opened = await openXlsx(bytesSource(source))

					if (!opened.ok) continue

					const workbook = opened.workbook
					const entry = workbook.sheets.find(candidate => candidate.kind === "worksheet" && candidate.part !== "")
					let model: XlsxSheetModel

					try {
						model = await workbook.load(entry?.id ?? -1)
					} catch {
						continue
					}

					if (!model.sections.some(section => section.name === "sheetViews")) continue

					const view = model.view
					const row = Math.min(view.activeRow + 1, 1_048_575)

					if (moved)
						model.view = {
							...view,
							activeRow: row,
							selection: [{ startRow: row, startCol: view.activeCol, endRow: row, endCol: view.activeCol }]
						}

					workbook.changes.sheetSection(model.id, "sheetViews")

					const result = await saveXlsx({ opc: workbook.opc, workbook })

					if (!result.ok || result.bytes === undefined) {
						failures.push(`${file} ${moved ? "moved" : "unchanged"}: save refused (${result.ok ? "" : result.message})`)

						continue
					}

					checked++

					const before = parts(source)[model.part.replace(/^\//, "")] ?? ""
					const after = parts(result.bytes)[model.part.replace(/^\//, "")] ?? ""

					if (!moved) {
						const changed = await changedEntries(source, result.bytes)

						if (changed.changed.length > 0) failures.push(`${file} unchanged: ${changed.changed.join(", ")} differ`)
					} else if (!subsequence(elements(before), elements(after))) {
						failures.push(`${file} moved: ${elements(before).join(",")} became ${elements(after).join(",")}`)
					}
				}
			}

			expect(failures).toEqual([])
			expect(checked).toBeGreaterThan(100)
		}
	)
})
