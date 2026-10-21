import { describe, expect, it } from "vitest"
import { bytesSource } from "../core/opc/index"
import { RenderStyleTable } from "../core/sheet/session"
import { XlsxDocument } from "../core/xlsx/edit/document"
import type { SheetStore } from "../core/xlsx/model/store"
import { openXlsx } from "../core/xlsx/read/index"
import { buildXlsx } from "./oracles/sheet/synthetic"

// A shared formula whose master is not its ref's top-left (Excel writes such groups: the master is the cell holding
// the text, ECMA-376 Part 1 §18.3.1.40): every child's references move from the master's own cell, the children
// before it included.
const SHEET =
	'<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c><c r="D1"><v>4</v></c></row>' +
	'<row r="2"><c r="A2"><f t="shared" si="0"/><v>10</v></c><c r="B2"><f t="shared" si="0"/><v>20</v></c>' +
	'<c r="C2"><f t="shared" ref="A2:D3" si="0">C1*10</f><v>30</v></c><c r="D2"><f t="shared" si="0"/><v>40</v></c></row>' +
	'<row r="3"><c r="A3"><f t="shared" si="0"/><v>100</v></c></row></sheetData>'

describe("a shared formula whose master is not its range's first cell", () => {
	it("moves each child's references from the master's own cell", async () => {
		const opened = await openXlsx(bytesSource(buildXlsx({ sheets: [{ name: "One", body: SHEET }] })))

		if (!opened.ok) throw new Error(opened.reason)

		const document = new XlsxDocument(opened.workbook, new RenderStyleTable(), { kind: "xlsx" })

		try {
			const sheet = document.worksheets()[0]?.id ?? 0
			const model = await document.ready(sheet)
			const store = model.store as SheetStore
			const text = (row: number, col: number): string => document.formulaText(model, store.formula(store.locate(row, col)), row, col)

			expect([text(1, 0), text(1, 1), text(1, 2), text(1, 3), text(2, 0)]).toEqual(["A1*10", "B1*10", "C1*10", "D1*10", "A2*10"])

			const engine = document.buildEngine()

			expect([0, 1, 2, 3].map(col => engine.result(0, 1, col).value)).toEqual([10, 20, 30, 40])
		} finally {
			document.close()
		}
	})
})
