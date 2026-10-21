import { describe, expect, it } from "vitest"
import type { RefEdit } from "../core/xlsx/refs/index"
import { R, TYPES, edited, refused, relationship, sheetBody } from "./xlsxRefsFixtures"
import type { Fixture } from "./xlsxWriteFixtures"

// The catalogue's rules for cells inserted or deleted in a band (Excel's Insert and Delete of cells): what lies
// across the band moves as for whole rows or columns, what reaches outside it stays, and what cannot be split
// refuses with the construct.

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

// Cells of columns `from`..`to` shifted down (an insert) or up (a delete) at row `at`.
function down(type: "insert" | "delete", at: number, from: number, to: number, count = 1): RefEdit {
	return { kind: "axis", sheet: 0, shift: { type, axis: "rows", at, count, band: { start: from, end: to } } }
}

// Cells of rows `from`..`to` shifted right or left at column `at`.
function right(type: "insert" | "delete", at: number, from: number, to: number, count = 1): RefEdit {
	return { kind: "axis", sheet: 0, shift: { type, axis: "cols", at, count, band: { start: from, end: to } } }
}

describe("references, cells shifted in a band", () => {
	it("moves a reference across the band and leaves one reaching outside it", async () => {
		const fixture = sheetBody(
			'<sheetData><row r="9"><c r="A9"><f>SUM(B1:B5)+SUM(A1:C5)+$B$3+SUM(B:B)+SUM(2:3)</f><v>0</v></c></row></sheetData>'
		)
		const { sheet } = await edited(fixture, down("insert", 1, 1, 1), ["formulas"])

		expect(sheet).toContain("<f>SUM(B1:B6)+SUM(A1:C5)+$B$4+SUM(B:B)+SUM(2:3)</f>")

		const deleted = await edited(fixture, down("delete", 2, 1, 1), ["formulas"])

		expect(deleted.sheet).toContain("<f>SUM(B1:B4)+SUM(A1:C5)+#REF!+SUM(B:B)+SUM(2:3)</f>")
	})

	it("moves a defined name across the band", async () => {
		const fixture: Fixture = {
			sheets: [{ name: "Data", body: "<sheetData/>" }],
			workbookTail:
				'<definedNames><definedName name="In">Data!$C$2:$D$3</definedName><definedName name="Out">Data!$C$1:$D$3</definedName></definedNames>'
		}
		const { files } = await edited(fixture, right("insert", 1, 1, 2), ["definedNames"])

		expect(files["xl/workbook.xml"]).toContain(">Data!$D$2:$E$3<")
		expect(files["xl/workbook.xml"]).toContain(">Data!$C$1:$D$3<")
	})
})

describe("ranges, cells shifted in a band", () => {
	it("moves a hyperlink with its cell and refuses one only partly in the band", async () => {
		const fixture = sheetBody(
			'<sheetData/><hyperlinks><hyperlink ref="B3" location="Data!A1" display="a"/><hyperlink ref="D3:F3" location="Data!B2" display="b"/></hyperlinks>'
		)
		const { sheet } = await edited(fixture, down("insert", 1, 1, 1), ["hyperlinks"])

		expect(sheet).toContain('<hyperlink ref="B4" location="Data!A1" display="a"/>')
		// Its location is a reference: into the band, it moves too.
		expect(sheet).toContain('<hyperlink ref="D3:F3" location="Data!B3" display="b"/>')
		expect(await refused(fixture, down("insert", 1, 4, 4))).toEqual({ reason: "structureLocked", detail: "hyperlinks" })
	})

	it("moves an autofilter across the band and refuses to shift part of one", async () => {
		const fixture = sheetBody(
			'<sheetData/><autoFilter ref="A1:C4"><filterColumn colId="2"><filters><filter val="x"/></filters></filterColumn></autoFilter>'
		)
		const { sheet } = await edited(fixture, down("insert", 2, 0, 2), ["autoFilter"])

		expect(sheet).toContain('<autoFilter ref="A1:C5">')
		expect(await refused(fixture, down("insert", 2, 1, 1))).toEqual({
			reason: "structureLocked",
			detail: "autoFilter: part of the filtered range"
		})

		// Rows across the band alone: columns inserted inside it move its filter column.
		const across = await edited(fixture, right("insert", 1, 0, 3), ["autoFilter"])

		expect(across.sheet).toContain('<autoFilter ref="A1:D4"><filterColumn colId="3">')
	})

	it("leaves page breaks, column widths and the view where they are", async () => {
		const fixture = sheetBody(
			'<sheetViews><sheetView workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="B5" sqref="B5"/></sheetView></sheetViews>' +
				'<cols><col min="2" max="3" width="20" customWidth="1"/></cols><sheetData/><rowBreaks count="1" manualBreakCount="1"><brk id="5" max="16383" man="1"/></rowBreaks>'
		)
		const { sheet } = await edited(fixture, down("insert", 1, 1, 1))

		expect(sheet).toContain('<pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/>')
		expect(sheet).toContain('activeCell="B5" sqref="B5"')
		expect(sheet).toContain('<brk id="5" max="16383" man="1"/>')
		expect(sheet).toContain('<col min="2" max="3" width="20" customWidth="1"/>')
	})

	it("moves a sparkline with its cell", async () => {
		const X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
		const XM = "http://schemas.microsoft.com/office/excel/2006/main"
		const fixture = sheetBody(
			`<sheetData/><extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="${X14}"><x14:sparklineGroups xmlns:xm="${XM}"><x14:sparklineGroup><x14:sparklines>` +
				"<x14:sparkline><xm:f>Data!B1:B4</xm:f><xm:sqref>B6</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>"
		)
		const { sheet } = await edited(fixture, down("insert", 0, 1, 1), ["sparklines"])

		expect(sheet).toContain("<xm:f>Data!B2:B5</xm:f><xm:sqref>B7</xm:sqref>")
	})
})

describe("what refuses, cells shifted in a band", () => {
	it("refuses to move part of an array formula", async () => {
		const fixture = sheetBody('<sheetData><row r="2"><c r="B2"><f t="array" ref="B2:C3">A1:B2*2</f><v>0</v></c></row></sheetData>')

		expect(await refused(fixture, down("insert", 0, 1, 1))).toEqual({ reason: "arrayFormula", detail: "array formula" })
		expect(await refused(fixture, right("delete", 2, 2, 2))).toEqual({ reason: "arrayFormula", detail: "array formula" })
	})

	it("locks a pivot table and its source", async () => {
		const fixture: Fixture = {
			sheets: [{ name: "Data", body: "<sheetData/>", rels: relationship("rId1", "pivotTable", "../pivotTables/pivotTable1.xml") }],
			parts: {
				"xl/pivotCache/pivotCacheDefinition1.xml": {
					text: `${DECLARATION}<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}" r:id="rId1"><cacheSource type="worksheet"><worksheetSource ref="H1:I5" sheet="Data"/></cacheSource><cacheFields count="0"/></pivotCacheDefinition>`,
					type: TYPES.pivotCache
				},
				"xl/pivotTables/pivotTable1.xml": {
					text: `${DECLARATION}<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P1" cacheId="1" dataCaption="Values"><location ref="B3:C9" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/></pivotTableDefinition>`,
					type: TYPES.pivotTable
				}
			}
		}

		expect(await refused(fixture, down("insert", 0, 2, 2))).toEqual({ reason: "pivot", detail: "pivotTables" })
		expect(await refused(fixture, down("delete", 0, 8, 8))).toEqual({ reason: "pivot", detail: "pivotSources" })

		// Beside both, or below them: allowed.
		await edited(fixture, down("insert", 0, 5, 6))
		await edited(fixture, down("insert", 12, 2, 2))
	})
})

describe("range lists' formulas", () => {
	it("are written for the top-left of the rectangle around the ranges, not the first range listed", async () => {
		const fixture = sheetBody(
			'<sheetData/><conditionalFormatting sqref="A5:A6 A1:A2"><cfRule type="expression" priority="1"><formula>$B1=1</formula></cfRule></conditionalFormatting>' +
				'<dataValidations count="1"><dataValidation type="custom" sqref="C5:C6 C1:C2"><formula1>$D1&gt;0</formula1></dataValidation></dataValidations>'
		)
		// The rows of A1:A2 deleted: A5:A6 is the list, its first cell was A5, read from A1 as B5.
		const { sheet } = await edited(fixture, { kind: "axis", sheet: 0, shift: { type: "delete", axis: "rows", at: 0, count: 2 } }, [
			"conditionalFormatting",
			"dataValidations"
		])

		expect(sheet).toContain('<conditionalFormatting sqref="A3:A4"><cfRule type="expression" priority="1"><formula>$B3=1</formula>')
		expect(sheet).toContain('sqref="C3:C4"><formula1>$D3&gt;0</formula1>')

		// Cells shifted up in column A alone: only the conditional format's list moves.
		const band = await edited(fixture, down("delete", 0, 0, 0, 2), ["conditionalFormatting"])

		expect(band.sheet).toContain('<conditionalFormatting sqref="A3:A4"><cfRule type="expression" priority="1"><formula>$B5=1</formula>')
	})
})
