import { describe, expect, it } from "vitest"
import type { SheetCommandResult, SheetRange, SheetSessionApi } from "../core/sheet/protocol"
import { openTextSession } from "../core/sheet/text"
import { parts, type Fixture } from "./xlsxWriteFixtures"
import { firstTile, openFixture, savedBytes, tileEntry } from "./sheetSessionFixtures"
import { TYPES, relationship } from "./xlsxRefsFixtures"

// Cells inserted and deleted (Excel's Insert and Delete of cells): the cells after them in the band move, and what
// refers to cells follows them by the band rules, or the edit is refused with the construct that cannot.

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

function range(a1: string): SheetRange {
	const [from = "", to = from] = a1.split(":")
	const cell = (text: string) => {
		const match = /^([A-Z]+)(\d+)$/.exec(text)

		if (match === null) throw new Error(text)

		let col = 0

		for (const letter of match[1] ?? "") col = col * 26 + letter.charCodeAt(0) - 64

		return { row: Number(match[2]) - 1, col: col - 1 }
	}
	const start = cell(from)
	const end = cell(to)

	return { startRow: start.row, startCol: start.col, endRow: end.row, endCol: end.col }
}

async function shifted(
	fixture: Fixture,
	a1: string,
	edit: "insert" | "delete",
	axis: "rows" | "cols"
): Promise<{ result: SheetCommandResult; session: SheetSessionApi; sheet: number; bytes: Uint8Array }> {
	const opened = await openFixture(fixture)
	const sheet = opened.info.sheets[0]!.id
	const result = await opened.session.command({ type: "cells.shift", sheet, range: range(a1), edit, axis })

	return { result, session: opened.session, sheet, bytes: opened.bytes }
}

function sheetXml(bytes: Uint8Array, n = 1): string {
	return parts(bytes)[`xl/worksheets/sheet${String(n)}.xml`] ?? ""
}

// Every cell of a sheet part: "A1=value" or "A1=formula" for formula cells.
function cellsOf(xml: string): string[] {
	return [...xml.matchAll(/<c r="([A-Z]+\d+)"[^>]*?(?:\/>|>([\s\S]*?)<\/c>)/g)].map(match => {
		const body = match[2] ?? ""
		const formula = /<f[^>]*>([^<]*)<\/f>/.exec(body)?.[1]
		const value = /<v>([^<]*)<\/v>/.exec(body)?.[1] ?? /<t[^>]*>([^<]*)<\/t>/.exec(body)?.[1] ?? ""

		return `${match[1] ?? ""}=${formula === undefined ? value : `=${formula}`}`
	})
}

function sqrefs(xml: string, element: string): string[] {
	return [...xml.matchAll(new RegExp(`<${element}\\b[^>]*\\ssqref="([^"]*)"`, "g"))].map(match => match[1] ?? "")
}

const GRID: Fixture = {
	sheets: [
		{
			name: "One",
			body:
				'<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c></row>' +
				'<row r="2"><c r="A2"><v>5</v></c><c r="B2"><v>6</v></c><c r="C2"><v>7</v></c></row>' +
				'<row r="3"><c r="B3"><v>8</v></c></row>' +
				'<row r="10"><c r="A10"><f>SUM(B1:B5)</f><v>16</v></c><c r="C10"><f>SUM(A1:C5)</f><v>32</v></c>' +
				'<c r="D10"><f>B2+C2</f><v>13</v></c><c r="E10"><f>SUM(B:B)</f><v>16</v></c><c r="F10"><f>$B$3</f><v>8</v></c></row></sheetData>'
		},
		{ name: "Two", body: '<sheetData><row r="1"><c r="A1"><f>One!B2*2</f><v>12</v></c></row></sheetData>' }
	]
}

describe("cells shifted down and up", () => {
	it("moves the band's cells down and the references across the band alone", async () => {
		const { result, session } = await shifted(GRID, "B2", "insert", "rows")

		expect(result.ok).toBe(true)

		const saved = await savedBytes(session)
		const cells = cellsOf(sheetXml(saved))

		expect(cells).toEqual(
			expect.arrayContaining(["B1=2", "B3=6", "B4=8", "A2=5", "C2=7", "A10==SUM(B1:B6)", "C10==SUM(A1:C5)", "D10==B3+C2"])
		)
		expect(cells).toEqual(expect.arrayContaining(["E10==SUM(B:B)", "F10==$B$4"]))
		expect(cells.some(cell => cell.startsWith("B2="))).toBe(false)
		expect(cellsOf(sheetXml(saved, 2))).toEqual(["A1==One!B3*2"])
	})

	it("deletes the band's cells, moving the cells below up and their references to #REF!", async () => {
		const { result, session } = await shifted(GRID, "B2", "delete", "rows")

		expect(result.ok).toBe(true)

		const cells = cellsOf(sheetXml(await savedBytes(session)))

		expect(cells).toEqual(expect.arrayContaining(["B2=8", "A2=5", "A10==SUM(B1:B4)", "D10==#REF!+C2", "F10==$B$2"]))
		expect(cellsOf(sheetXml(await savedBytes(session), 2))).toEqual(["A1==One!#REF!*2"])
	})

	it("recalculates what reads the moved cells", async () => {
		const { session, info } = await openFixture(GRID)
		const sheet = info.sheets[0]!.id

		// An entry builds the formula engine.
		await session.command({ type: "cells.set", sheet, cells: { rows: Uint32Array.of(0), cols: Uint16Array.of(5), inputs: ["1"] } })
		await session.command({ type: "cells.shift", sheet, range: range("B1"), edit: "delete", axis: "rows" })

		const tile = await firstTile(session, sheet)

		// SUM(B1:B5) became SUM(B1:B4) over 6 and 8.
		expect(tileEntry(tile, 9, 0)?.text).toBe("14")
		expect(tileEntry(tile, 0, 1)?.text).toBe("6")
	})

	it("refuses to push the band's cells off the sheet", async () => {
		const fixture: Fixture = {
			sheets: [{ name: "One", body: '<sheetData><row r="1048576"><c r="B1048576"><v>1</v></c></row></sheetData>' }]
		}
		const { result } = await shifted(fixture, "B1", "insert", "rows")

		expect(result.ok ? null : result.refusal).toEqual({ reason: "tooLarge", detail: "cells" })
		expect((await shifted(fixture, "C1", "insert", "rows")).result.ok).toBe(true)
	})
})

describe("cells shifted right and left", () => {
	it("moves the band's cells right and grows a reference across the band", async () => {
		const fixture: Fixture = {
			sheets: [
				{
					name: "One",
					body:
						'<sheetData><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c><c r="C2"><v>3</v></c></row>' +
						'<row r="3"><c r="B3"><v>4</v></c></row>' +
						'<row r="5"><c r="A5"><f>SUM(A2:C3)</f><v>10</v></c><c r="B5"><f>SUM(A2:C2)</f><v>6</v></c></row></sheetData>'
				}
			],
			workbookTail: '<definedNames><definedName name="Row2">One!$B$2:$C$2</definedName></definedNames>'
		}
		const { result, session } = await shifted(fixture, "B2", "insert", "cols")

		expect(result.ok).toBe(true)

		const saved = await savedBytes(session)

		expect(cellsOf(sheetXml(saved))).toEqual(["A2=1", "C2=2", "D2=3", "B3=4", "A5==SUM(A2:C3)", "B5==SUM(A2:D2)"])
		expect(parts(saved)["xl/workbook.xml"]).toContain("One!$C$2:$D$2")
	})

	it("deletes cells to the left", async () => {
		const fixture: Fixture = {
			sheets: [
				{
					name: "One",
					body:
						'<sheetData><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c><c r="C2"><v>3</v></c></row>' +
						'<row r="5"><c r="A5"><f>B2+C2</f><v>5</v></c></row></sheetData>'
				}
			]
		}
		const { session } = await shifted(fixture, "B2", "delete", "cols")

		expect(cellsOf(sheetXml(await savedBytes(session)))).toEqual(["A2=1", "B2=3", "A5==#REF!+B2"])
	})
})

describe("constructs over the shifted cells", () => {
	it("moves conditional formats by LibreOffice's range-list rule and validations with their cells", async () => {
		const fixture: Fixture = {
			sheets: [
				{
					name: "One",
					body:
						'<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
						'<conditionalFormatting sqref="B1:B4"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>2</formula></cfRule></conditionalFormatting>' +
						'<conditionalFormatting sqref="A1:D4"><cfRule type="expression" priority="2"><formula>A1&gt;$B$2</formula></cfRule></conditionalFormatting>' +
						'<dataValidations count="1"><dataValidation type="whole" operator="between" allowBlank="1" sqref="A3:C3"><formula1>0</formula1><formula2>9</formula2></dataValidation></dataValidations>'
				}
			]
		}
		const inserted = await shifted(fixture, "B2", "insert", "rows")
		const insertedXml = sheetXml(await savedBytes(inserted.session))

		// A range across the band grows; one reaching outside it stays; a validation's cells in the band move.
		expect(sqrefs(insertedXml, "conditionalFormatting")).toEqual(["B1:B5", "A1:D4"])
		expect(insertedXml).toContain("A1&gt;$B$3")
		expect(sqrefs(insertedXml, "dataValidation")).toEqual(["A3 C3 B4"])

		const deleted = await shifted(fixture, "B2", "delete", "cols")
		const deletedXml = sheetXml(await savedBytes(deleted.session))

		// The deleted cell leaves each range; a piece lying across the band then moves left, any other stays.
		expect(sqrefs(deletedXml, "conditionalFormatting")).toEqual(["B1 B3:B4", "A1:D1 A3:D4 A2 B2:C2"])
		expect(deletedXml).toContain("A1&gt;#REF!")
	})

	it("moves a merge across the band and refuses one only partly in it", async () => {
		const fixture: Fixture = {
			sheets: [{ name: "One", body: '<sheetData/><mergeCells count="1"><mergeCell ref="B5:C6"/></mergeCells>' }]
		}
		const across = await shifted(fixture, "B2:C2", "insert", "rows")

		expect(sheetXml(await savedBytes(across.session))).toContain('<mergeCell ref="B6:C7"/>')

		const partly = await shifted(fixture, "B2", "insert", "rows")

		expect(partly.result.ok ? null : partly.result.refusal).toEqual({ reason: "merged", detail: "mergeCells" })

		// A merge before the shifted cells stays.
		expect((await shifted(fixture, "B8", "insert", "rows")).result.ok).toBe(true)
	})

	it("moves notes with their cells and drops the note of a deleted cell", async () => {
		const comments =
			`${DECLARATION}<comments xmlns="${MAIN}"><authors><author>a</author></authors><commentList>` +
			'<comment ref="B2" authorId="0"><text><t>x</t></text></comment><comment ref="C2" authorId="0"><text><t>y</t></text></comment></commentList></comments>'
		const vml =
			`${DECLARATION}<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">` +
			'<v:shape id="_x0000_s1025"><x:ClientData ObjectType="Note"><x:Anchor>2, 15, 0, 10, 4, 15, 4, 4</x:Anchor><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>' +
			'<v:shape id="_x0000_s1026"><x:ClientData ObjectType="Note"><x:Anchor>3, 15, 0, 10, 5, 15, 4, 4</x:Anchor><x:Row>1</x:Row><x:Column>2</x:Column></x:ClientData></v:shape></xml>'
		const fixture: Fixture = {
			sheets: [
				{
					name: "One",
					body: '<sheetData><row r="2"><c r="B2"><v>1</v></c><c r="C2"><v>2</v></c></row></sheetData><legacyDrawing r:id="rId2"/>',
					rels:
						relationship("rId1", "comments", "../comments1.xml") +
						relationship("rId2", "vmlDrawing", "../drawings/vmlDrawing1.vml")
				}
			],
			parts: {
				"xl/comments1.xml": { text: comments, type: TYPES.comments },
				"xl/drawings/vmlDrawing1.vml": { text: vml, type: TYPES.vml }
			}
		}
		const down = await shifted(fixture, "B1", "insert", "rows")
		const downParts = parts(await savedBytes(down.session))

		expect(downParts["xl/comments1.xml"]).toContain('ref="B3"')
		expect(downParts["xl/comments1.xml"]).toContain('ref="C2"')
		expect(downParts["xl/drawings/vmlDrawing1.vml"]).toContain("<x:Anchor>2, 15, 1, 10, 4, 15, 5, 4</x:Anchor><x:Row>2</x:Row>")
		expect(downParts["xl/drawings/vmlDrawing1.vml"]).toContain("<x:Anchor>3, 15, 0, 10, 5, 15, 4, 4</x:Anchor><x:Row>1</x:Row>")

		const removed = await shifted(fixture, "B2", "delete", "cols")
		const removedParts = parts(await savedBytes(removed.session))

		expect(removedParts["xl/comments1.xml"]).not.toContain('ref="B2" authorId="0"><text><t>x')
		expect(removedParts["xl/comments1.xml"]).toContain('ref="B2" authorId="0"><text><t>y')
	})

	it("moves a table across the band and refuses to shift part of one", async () => {
		const tablePart =
			`${DECLARATION}<table xmlns="${MAIN}" id="1" name="T" displayName="T" ref="B2:C4"><autoFilter ref="B2:C4"/>` +
			'<tableColumns count="2"><tableColumn id="1" name="K"/><tableColumn id="2" name="V"/></tableColumns></table>'
		const fixture: Fixture = {
			sheets: [
				{
					name: "One",
					body:
						'<sheetData><row r="2"><c r="B2" t="inlineStr"><is><t>K</t></is></c><c r="C2" t="inlineStr"><is><t>V</t></is></c></row>' +
						'<row r="3"><c r="B3"><v>1</v></c><c r="C3"><v>2</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts>',
					rels: relationship("rId1", "table", "../tables/table1.xml")
				}
			],
			parts: { "xl/tables/table1.xml": { text: tablePart, type: TYPES.table } }
		}
		const above = await shifted(fixture, "A1:D1", "insert", "rows")

		expect(parts(await savedBytes(above.session))["xl/tables/table1.xml"]).toContain('ref="B3:C5"')

		const inside = await shifted(fixture, "B3:C3", "insert", "rows")

		expect(parts(await savedBytes(inside.session))["xl/tables/table1.xml"]).toContain('ref="B2:C5"')

		const partly = await shifted(fixture, "B3", "insert", "rows")

		expect(partly.result.ok ? null : partly.result.refusal).toEqual({ reason: "structureLocked", detail: "tables: part of a table" })
	})

	it("moves a drawing with its top-left cell, as a whole when it lies across the band", async () => {
		const anchor = (fromCol: number, toCol: number) =>
			`<xdr:twoCellAnchor><xdr:from><xdr:col>${String(fromCol)}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
			`<xdr:to><xdr:col>${String(toCol)}</xdr:col><xdr:colOff>10</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>10</xdr:rowOff></xdr:to>` +
			'<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="s"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>'
		const drawing = (body: string) =>
			`${DECLARATION}<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${body}</xdr:wsDr>`
		const fixture = (body: string): Fixture => ({
			sheets: [
				{
					name: "One",
					body: '<sheetData/><drawing r:id="rId1"/>',
					rels: relationship("rId1", "drawing", "../drawings/drawing1.xml")
				}
			],
			parts: { "xl/drawings/drawing1.xml": { text: drawing(body), type: TYPES.drawing } }
		})
		const rowsOf = (xml: string) => [...xml.matchAll(/<xdr:row>(\d+)<\/xdr:row>/g)].map(match => Number(match[1]))

		// Across columns B:C alone, an insert of two rows above moves both edges.
		const whole = await shifted(fixture(anchor(1, 2)), "B1:C2", "insert", "rows")

		expect(rowsOf(parts(await savedBytes(whole.session))["xl/drawings/drawing1.xml"] ?? "")).toEqual([6, 10])

		// Its top-left cell in the band, reaching outside it: moves, keeping its size; inside it, a delete moves it up.
		const top = await shifted(fixture(anchor(1, 4)), "B1", "insert", "rows")

		expect(rowsOf(parts(await savedBytes(top.session))["xl/drawings/drawing1.xml"] ?? "")).toEqual([5, 9])

		// Its top-left cell outside the band: stays.
		const outside = await shifted(fixture(anchor(0, 4)), "B1", "insert", "rows")

		expect(rowsOf(parts(await savedBytes(outside.session))["xl/drawings/drawing1.xml"] ?? "")).toEqual([4, 8])
	})

	it("moves an array formula across the band, and reads a shared formula's moved members on their own", async () => {
		const fixture: Fixture = {
			sheets: [
				{
					name: "One",
					body:
						'<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c>' +
						'<c r="D1"><f t="array" ref="D1:E2">A1:B2*2</f><v>2</v></c></row>' +
						'<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>' +
						'<row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>6</v></c></row></sheetData>'
				}
			]
		}
		const array = await shifted(fixture, "D1:E1", "insert", "rows")
		const arrayXml = sheetXml(await savedBytes(array.session))

		expect(arrayXml).toMatch(/<c r="D2"><f t="array" ref="D2:E3">A1:B2\*2<\/f>/)

		// B2 and B3 move to B3 and B4 and still read A2 and A3, which no longer follow from B1's A1*2.
		const shared = await shifted(fixture, "B2", "insert", "rows")
		const cells = cellsOf(sheetXml(await savedBytes(shared.session)))

		expect(cells).toEqual(expect.arrayContaining(["B1==A1*2", "B3==A2*2", "B4==A3*2"]))
		expect((await shifted(fixture, "D2", "insert", "rows")).result).toMatchObject({ ok: false, refusal: { reason: "arrayFormula" } })
	})

	it("refuses on a protected sheet", async () => {
		const fixture: Fixture = {
			sheets: [{ name: "One", body: '<sheetData/><sheetProtection sheet="1" insertRows="0" deleteRows="0"/>' }]
		}
		const { result } = await shifted(fixture, "B2", "insert", "rows")

		expect(result.ok ? null : result.refusal.reason).toBe("protectedSheet")
	})
})

describe("cells shifted, undone and redone", () => {
	it("saves the source's bytes after an undo, and the edit's after a redo", async () => {
		const { session, bytes } = await shifted(GRID, "B2:C2", "delete", "rows")
		const edited = await savedBytes(session)

		await session.undo()

		const undone = await savedBytes(session)

		expect(parts(undone)["xl/worksheets/sheet1.xml"]).toBe(parts(bytes)["xl/worksheets/sheet1.xml"])
		expect(parts(undone)["xl/worksheets/sheet2.xml"]).toBe(parts(bytes)["xl/worksheets/sheet2.xml"])

		await session.redo()

		expect(parts(await savedBytes(session))["xl/worksheets/sheet1.xml"]).toBe(parts(edited)["xl/worksheets/sheet1.xml"])
	})

	it("shows again, after an undo, the result of a deleted formula the file keeps none for", async () => {
		const fixture: Fixture = {
			sheets: [
				{
					name: "One",
					body: '<sheetData><row r="1"><c r="A1"><v>4</v></c></row><row r="3"><c r="B3"><f>A1+1</f></c></row></sheetData>'
				}
			]
		}
		const { session, info } = await openFixture(fixture)
		const sheet = info.sheets[0]!.id

		await session.command({ type: "cells.set", sheet, cells: { rows: Uint32Array.of(0), cols: Uint16Array.of(5), inputs: ["1"] } })

		expect(tileEntry(await firstTile(session, sheet), 2, 1)?.text).toBe("5")

		await session.command({ type: "cells.shift", sheet, range: range("B3"), edit: "delete", axis: "rows" })

		expect(tileEntry(await firstTile(session, sheet), 2, 1)).toBeNull()

		await session.undo()

		expect(tileEntry(await firstTile(session, sheet), 2, 1)?.text).toBe("5")
	})

	it("gives a deleted row's cells back where they were after a later column insert and cell edit are undone", async () => {
		const fixture: Fixture = {
			sheets: [
				{
					name: "One",
					body: '<sheetData><row r="4"><c r="F4"><v>1</v></c></row><row r="5"><c r="F5"><v>2</v></c></row></sheetData>'
				}
			]
		}
		const { session, info, bytes } = await openFixture(fixture)
		const sheet = info.sheets[0]!.id

		await session.command({ type: "axis.delete", sheet, axis: "rows", at: 1, count: 1 })
		await session.command({ type: "axis.insert", sheet, axis: "cols", at: 1, count: 3 })
		await session.command({ type: "cells.set", sheet, cells: { rows: Uint32Array.of(3), cols: Uint16Array.of(4), inputs: ["x"] } })
		await session.undo()
		await session.undo()
		await session.undo()

		expect(cellsOf(sheetXml(await savedBytes(session)))).toEqual(cellsOf(sheetXml(bytes)))
	})

	it("inserts whole rows for a range across every column", async () => {
		const opened = await openFixture(GRID)
		const sheet = opened.info.sheets[0]!.id
		const result = await opened.session.command({
			type: "cells.shift",
			sheet,
			range: { startRow: 1, startCol: 0, endRow: 1, endCol: 16_383 },
			edit: "insert",
			axis: "rows"
		})

		expect(result.ok).toBe(true)
		expect(cellsOf(sheetXml(await savedBytes(opened.session)))).toEqual(expect.arrayContaining(["A3=5", "C3=7", "A11==SUM(B1:B6)"]))
	})
})

describe("cells shifted in a text file", () => {
	async function csv(text: string, a1: string, edit: "insert" | "delete", axis: "rows" | "cols"): Promise<SheetSessionApi> {
		const opened = await openTextSession(new TextEncoder().encode(text), { kind: "csv" }, () => undefined)

		if (opened.kind !== "opened") throw new Error(opened.kind)

		const result = await opened.session.command({ type: "cells.shift", sheet: 0, range: range(a1), edit, axis })

		expect(result.ok).toBe(true)

		return opened.session
	}

	const text = async (session: SheetSessionApi): Promise<string> => new TextDecoder().decode(await savedBytes(session))

	it("moves a column's texts down and up", async () => {
		expect(await text(await csv("a,b,c\r\n1,2,3\r\n", "B1", "insert", "rows"))).toBe("a,,c\r\n1,b,3\r\n,2\r\n")
		expect(await text(await csv("a,b,c\r\n1,2,3\r\n", "B1", "delete", "rows"))).toBe("a,2,c\r\n1,,3\r\n")
	})

	it("moves a row's texts right and left, and undoes as one", async () => {
		const session = await csv("a,b,c\r\n1,2,3\r\n", "A2:B2", "insert", "cols")

		expect(await text(session)).toBe("a,b,c\r\n,,1,2,3\r\n")

		await session.undo()

		expect(await text(session)).toBe("a,b,c\r\n1,2,3\r\n")
		// The cell left empty at the row's end stays a field, as clearing it does.
		expect(await text(await csv("a,b,c\r\n1,2,3\r\n", "A1", "delete", "cols"))).toBe("b,c,\r\n1,2,3\r\n")
	})
})
