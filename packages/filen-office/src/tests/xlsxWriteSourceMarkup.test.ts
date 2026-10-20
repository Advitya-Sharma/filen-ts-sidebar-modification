import { unzipSync, zipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { bytesSource } from "../core/opc/index"
import type { SheetCommandResult } from "../core/sheet/protocol"
import { openXlsx } from "../core/xlsx/read/index"
import { SML, cells, firstTile, openFixture, savedBytes } from "./sheetSessionFixtures"
import { CellKind } from "../core/xlsx/model/types"
import { saveXlsx } from "../core/xlsx/write/save"
import { buildFixture, changedEntries, equalBytes, openWorkbook, parts } from "./xlsxWriteFixtures"

// What a cell holds beyond the model (value metadata, phonetic flags, rich runs, a formula's own attributes) stays
// on that cell through edits and saves, and an edit undone saves as its source.

function ok(result: SheetCommandResult): void {
	if (!result.ok) throw new Error(`refused: ${result.refusal.reason} ${result.refusal.detail ?? ""}`)
}

const SST = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${SML}" count="1" uniqueCount="1"><si><t>hello</t></si></sst>`

async function sheetAfter(
	body: string,
	edit: (session: Awaited<ReturnType<typeof openFixture>>["session"], sheet: number) => Promise<void>,
	sharedStrings?: string
): Promise<string> {
	const { session, info } = await openFixture({
		sheets: [{ name: "One", body }],
		...(sharedStrings === undefined ? {} : { sharedStrings })
	})
	const one = info.sheets[0]!.id

	await firstTile(session, one)
	await session.intent()
	await edit(session, one)

	const sheet = parts(await savedBytes(session))["xl/worksheets/sheet1.xml"] ?? ""

	session.close()

	return sheet
}

describe("a cell's own markup", () => {
	it("stays on its cell when a value is typed into an empty cell before equal cells", async () => {
		const sheet = await sheetAfter(
			'<sheetData><row r="1"><c r="B1" vm="1"><v>1</v></c><c r="C1" vm="2"><v>1</v></c></row></sheetData>',
			async (session, one) => ok(await session.command({ type: "cells.set", sheet: one, cells: cells([[0, 0, "1"]]) }))
		)

		expect(sheet).toContain('<row r="1"><c r="A1"><v>1</v></c><c r="B1" vm="1"><v>1</v></c><c r="C1" vm="2"><v>1</v></c></row>')
	})

	it("moves with its cell when an inserted column moves equal cells", async () => {
		const sheet = await sheetAfter(
			'<sheetData><row r="1"><c r="A1" t="e" vm="1"><v>#VALUE!</v></c><c r="B1" t="e" vm="2"><v>#VALUE!</v></c></row></sheetData>',
			async (session, one) => ok(await session.command({ type: "axis.insert", sheet: one, axis: "cols", at: 0, count: 1 }))
		)

		expect(sheet).toContain('<c r="B1" t="e" vm="1">')
		expect(sheet).toContain('<c r="C1" t="e" vm="2">')
	})

	it("moves with its cell when a deleted column moves equal cells", async () => {
		const sheet = await sheetAfter(
			'<sheetData><row r="1"><c r="A1"><v>0</v></c><c r="B1" t="s" ph="1"><v>0</v></c><c r="C1" t="s"><v>0</v></c></row></sheetData>',
			async (session, one) => ok(await session.command({ type: "axis.delete", sheet: one, axis: "cols", at: 0, count: 1 })),
			SST
		)

		expect(sheet).toContain('<c r="A1" t="s" ph="1">')
		expect(sheet).toContain('<c r="B1" t="s">')
	})

	it("keeps an inline rich string's runs when only the cell's format changes", async () => {
		const sheet = await sheetAfter(
			'<sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b/><sz val="11"/></rPr><t>Bold</t></r><r><t xml:space="preserve"> plain</t></r></is></c></row></sheetData>',
			async (session, one) =>
				ok(
					await session.command({
						type: "format.apply",
						sheet: one,
						ranges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }],
						patch: { italic: true }
					})
				)
		)

		expect(sheet).toContain('<is><r><rPr><b/><sz val="11"/></rPr><t>Bold</t></r><r><t xml:space="preserve"> plain</t></r></is>')
	})

	it("keeps a value's text as written when only the cell's format changes", async () => {
		const sheet = await sheetAfter(
			'<sheetData><row r="1"><c r="A1"><v>0.25080000000000002</v></c><c r="B1" t="str"><f>"x"&amp;""</f><v/></c></row></sheetData>',
			async (session, one) =>
				ok(
					await session.command({
						type: "format.apply",
						sheet: one,
						ranges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }],
						patch: { bold: true }
					})
				)
		)

		expect(sheet).toContain('<v>0.25080000000000002</v></c><c r="B1" s="2" t="str"><f>"x"&amp;""</f><v/></c>')
	})

	it("keeps a formula's xml:space when only the cell's format changes", async () => {
		const sheet = await sheetAfter(
			'<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f xml:space="preserve"> A1*2</f><v>2</v></c></row></sheetData>',
			async (session, one) =>
				ok(
					await session.command({
						type: "format.apply",
						sheet: one,
						ranges: [{ startRow: 0, startCol: 1, endRow: 0, endCol: 1 }],
						patch: { bold: true }
					})
				)
		)

		expect(sheet).toContain('<f xml:space="preserve"> A1*2</f>')
	})
})

describe("an edited row", () => {
	it("keeps the calcChain when only a formula cell's format changes", async () => {
		const { session, info, bytes } = await openFixture({
			sheets: [
				{ name: "One", body: '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row></sheetData>' }
			],
			calcChain: true
		})
		const one = info.sheets[0]!.id

		await firstTile(session, one)
		await session.intent()
		ok(
			await session.command({
				type: "format.apply",
				sheet: one,
				ranges: [{ startRow: 0, startCol: 1, endRow: 0, endCol: 1 }],
				patch: { bold: true }
			})
		)

		expect((await changedEntries(bytes, await savedBytes(session))).removed).toEqual([])
		session.close()
	})

	it("saves an edit and its undo in a row whose cells have no r as its source", async () => {
		const { session, info, bytes } = await openFixture({
			sheets: [{ name: "One", body: '<sheetData><row r="1"><c t="inlineStr"><is><t>x</t></is></c><c><v>1</v></c></row></sheetData>' }]
		})
		const one = info.sheets[0]!.id

		await firstTile(session, one)
		await session.intent()
		ok(await session.command({ type: "cells.set", sheet: one, cells: cells([[0, 1, "2"]]) }))
		ok(await session.undo())

		const result = await session.save()

		expect(result.ok ? "saved" : `${result.reason}: ${result.message}`).toBe("saved")
		if (result.ok) expect(equalBytes(result.bytes!, bytes)).toBe(true)

		ok(await session.command({ type: "cells.set", sheet: one, cells: cells([[5, 0, "later"]]) }))

		const later = await session.save()

		expect(later.ok ? "saved" : `${later.reason}: ${later.message}`).toBe("saved")
		session.close()
	})
})

describe("an edit and its undo", () => {
	it.each([
		["a new string", "cells.set", '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>'],
		[
			"a new string in a workbook without shared strings",
			"cells.set",
			'<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'
		],
		["a formula", "formula", '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row></sheetData>'],
		["a new format", "format", '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>']
	])("saves %s and its undo as the source", async (_, edit, body) => {
		const { session, info, bytes } = await openFixture({
			sheets: [{ name: "One", body }],
			calcChain: true,
			...(body.includes('t="s"') ? { sharedStrings: SST } : {})
		})
		const one = info.sheets[0]!.id

		await firstTile(session, one)
		await session.intent()

		if (edit === "format") {
			ok(
				await session.command({
					type: "format.apply",
					sheet: one,
					ranges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }],
					patch: { italic: true }
				})
			)
		} else {
			ok(
				await session.command({
					type: "cells.set",
					sheet: one,
					cells: cells([[0, edit === "formula" ? 1 : 0, edit === "formula" ? "=A1*3" : "brand new"]])
				})
			)
		}

		ok(await session.undo())

		const output = await savedBytes(session)

		expect(await changedEntries(bytes, output)).toEqual({ changed: [], added: [], removed: [] })
		session.close()
	})

	it("saves a format and its undo in a workbook without a styles part as the source", async () => {
		const { session, info, bytes } = await openFixture({
			sheets: [{ name: "One", body: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' }],
			styles: null
		})
		const one = info.sheets[0]!.id

		await firstTile(session, one)
		await session.intent()
		ok(
			await session.command({
				type: "format.apply",
				sheet: one,
				ranges: [{ startRow: 0, startCol: 0, endRow: 3, endCol: 3 }],
				patch: { bold: true }
			})
		)
		ok(await session.undo())

		expect(await changedEntries(bytes, await savedBytes(session))).toEqual({ changed: [], added: [], removed: [] })
		session.close()
	})

	it("writes the appended records a later edit still uses", async () => {
		const { session, info } = await openFixture({
			sheets: [{ name: "One", body: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' }],
			sharedStrings: SST
		})
		const one = info.sheets[0]!.id

		await firstTile(session, one)
		await session.intent()
		ok(await session.command({ type: "cells.set", sheet: one, cells: cells([[1, 0, "first"]]) }))
		ok(await session.command({ type: "cells.set", sheet: one, cells: cells([[2, 0, "second"]]) }))
		ok(await session.undo())
		ok(
			await session.command({
				type: "format.apply",
				sheet: one,
				ranges: [{ startRow: 1, startCol: 0, endRow: 1, endCol: 0 }],
				patch: { bold: true }
			})
		)
		ok(
			await session.command({
				type: "format.apply",
				sheet: one,
				ranges: [{ startRow: 2, startCol: 0, endRow: 2, endCol: 0 }],
				patch: { italic: true }
			})
		)
		ok(await session.undo())

		const written = parts(await savedBytes(session))

		expect(written["xl/sharedStrings.xml"]).toContain("<si><t>first</t></si></sst>")
		expect(written["xl/sharedStrings.xml"]).not.toContain("second")
		expect(written["xl/styles.xml"]).toContain('<cellXfs count="3">')
		session.close()
	})
})

describe("a column insert or delete", () => {
	it("moves each row's cells as written, renaming them and widening spans", async () => {
		const sheet = await sheetAfter(
			'<sheetData><row r="1" spans="1:2"><c r="A1"><v>1.50</v></c><c r="B1"><f>A1*2</f><v>3.0</v></c></row><row r="2"><c r="A2"><v>2</v></c></row></sheetData>',
			async (session, one) => ok(await session.command({ type: "axis.insert", sheet: one, axis: "cols", at: 0, count: 1 }))
		)

		expect(sheet).toContain(
			'<row r="1" spans="1:3"><c r="B1"><v>1.50</v></c><c r="C1"><f>B1*2</f><v>3.0</v></c></row><row r="2"><c r="B2"><v>2</v></c></row>'
		)
	})

	it("cuts out the deleted cells and keeps the others as written", async () => {
		const sheet = await sheetAfter(
			'<sheetData><row r="1"><c r="A1"><v>1.50</v></c> <c r="B1"><v>2</v></c> <c r="C1"><v>3.0</v></c></row></sheetData>',
			async (session, one) => ok(await session.command({ type: "axis.delete", sheet: one, axis: "cols", at: 1, count: 1 }))
		)

		expect(sheet).toContain('<row r="1"><c r="A1"><v>1.50</v></c>  <c r="B1"><v>3.0</v></c></row>')
	})
})

describe("a shared formula", () => {
	const body =
		'<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c></row>' +
		'<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>' +
		'<row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>6</v></c></row></sheetData>'

	it("is never saved with a child outside every master's range", async () => {
		const workbook = await openWorkbook(buildFixture({ sheets: [{ name: "One", body }] }))
		const model = await workbook.load(0)

		// The master written over without its children being unshared.
		model.store.setCell(0, 1, CellKind.NUMBER, 5, 0, -1)
		workbook.changes.sheetRows(0)

		const result = await saveXlsx({ opc: workbook.opc, workbook })

		expect(result.ok ? "saved" : result.reason).toBe("reference")
	})

	it("saves an unrelated edit in a sheet whose source already has a child outside its master's range", async () => {
		const orphaned =
			'<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>2</v></c></row>' +
			'<row r="2"><c r="B2"><f t="shared" si="0"/><v>0</v></c></row><row r="3"><c r="B3"><f t="shared" si="0"/><v>0</v></c></row></sheetData>'
		const { session, info } = await openFixture({ sheets: [{ name: "One", body: orphaned }] })
		const one = info.sheets[0]!.id

		await firstTile(session, one)
		await session.intent()
		ok(await session.command({ type: "cells.set", sheet: one, cells: cells([[9, 5, "x"]]) }))

		const result = await session.save()

		expect(result.ok ? "saved" : `${result.reason}: ${result.message}`).toBe("saved")
		session.close()
	})

	it("saves with its master over each child", async () => {
		const workbook = await openWorkbook(buildFixture({ sheets: [{ name: "One", body }] }))
		const model = await workbook.load(0)

		model.store.setCell(2, 0, CellKind.NUMBER, 7, 0, -1)
		workbook.changes.sheetRows(0)

		const result = await saveXlsx({ opc: workbook.opc, workbook })

		expect(result.ok ? "saved" : `${result.reason}: ${result.message}`).toBe("saved")
	})
})

describe("a UTF-16 package", () => {
	it("takes a new relationship in its UTF-16 relationship part", async () => {
		const built = unzipSync(
			buildFixture({
				sheets: [{ name: "One", body: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' }],
				calcChain: true
			})
		)
		const utf16 = (part: string): void => {
			const text = new TextDecoder().decode(built[part]).replace('encoding="UTF-8"', 'encoding="UTF-16"')
			const out = new Uint8Array(2 + text.length * 2)

			out[0] = 0xff
			out[1] = 0xfe

			for (let i = 0; i < text.length; i++) {
				out[2 + i * 2] = text.charCodeAt(i) & 0xff
				out[3 + i * 2] = text.charCodeAt(i) >> 8
			}

			built[part] = out
		}

		utf16("xl/_rels/workbook.xml.rels")

		const { session, info } = await openFixture(zipSync(built))
		const one = info.sheets[0]!.id

		await firstTile(session, one)
		await session.intent()
		ok(await session.command({ type: "cells.set", sheet: one, cells: cells([[0, 1, "=A1*2"]]) }))
		ok(await session.command({ type: "sheet.add", name: "Two" }))

		const result = await session.save()

		expect(result.ok ? "saved" : `${result.reason}: ${result.message}`).toBe("saved")
		if (result.ok) expect((await openXlsx(bytesSource(result.bytes!))).ok).toBe(true)
		session.close()
	})
})
