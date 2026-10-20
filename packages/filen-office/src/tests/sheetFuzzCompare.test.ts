import { describe, expect, it } from "vitest"
import { readSaved } from "./fuzz/sheet/compare"
import { cells, firstTile, openFixture, savedBytes } from "./sheetSessionFixtures"
import { buildFixture, parts } from "./xlsxWriteFixtures"

// The fuzz's own reading of a saved workbook, against which the model is compared.

const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

describe("the fuzz's reading of a saved workbook", () => {
	it("reads a font wrapped in mc:AlternateContent as one font, its Fallback", async () => {
		const bytes = buildFixture({
			sheets: [{ name: "One", body: '<sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData>' }],
			styles:
				`<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><mc:AlternateContent xmlns:mc="${MC}">` +
				'<mc:Choice xmlns:hs="urn:other" Requires="hs"><font hs:extension="1"><sz val="9"/><name val="Arial"/></font></mc:Choice>' +
				'<mc:Fallback><font><b/><sz val="9"/><name val="Arial"/></font></mc:Fallback></mc:AlternateContent></fonts>' +
				'<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
				'<borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
				'<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/></cellXfs>'
		})
		const record = await readSaved(bytes)

		expect(record.problems).toEqual([])
		expect(record.fonts.map(font => [font.name, font.size, font.bold])).toEqual([
			["Calibri", 11, false],
			["Arial", 9, true]
		])
	})

	// A source whose numFmts count is wrong (3 declared, 2 held): a save that appends none keeps the part as written.
	const WRONG_COUNT =
		'<numFmts count="3"><numFmt numFmtId="164" formatCode="0.##"/><numFmt numFmtId="165" formatCode="0.000"/></numFmts>' +
		'<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
		'<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
		'<borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
		'<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>'

	it("takes a container count the source got wrong and the save kept for the source's, not a save's fault", async () => {
		const { session, info, bytes } = await openFixture({
			sheets: [{ name: "One", body: '<sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData>' }],
			styles: WRONG_COUNT
		})
		const one = info.sheets[0]!.id

		await firstTile(session, one)
		await session.intent()
		await session.command({ type: "cells.set", sheet: one, cells: cells([[0, 1, "2"]]) })

		const saved = await savedBytes(session)

		session.close()

		// The save left styles.xml as the source wrote it.
		expect(parts(saved)["xl/styles.xml"]).toBe(parts(bytes)["xl/styles.xml"])
		expect((await readSaved(saved, bytes)).problems).toEqual([])
		// Without the source, the count reads as wrong.
		expect((await readSaved(saved)).problems).toEqual(["styles numFmts count 3, 2 held"])
	})

	it("still finds a container count the save got wrong", async () => {
		const { session, info, bytes } = await openFixture({
			sheets: [{ name: "One", body: '<sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData>' }],
			styles: WRONG_COUNT
		})
		const one = info.sheets[0]!.id

		await firstTile(session, one)
		await session.intent()
		await session.command({
			type: "format.apply",
			sheet: one,
			ranges: [{ startRow: 0, startCol: 1, endRow: 0, endCol: 1 }],
			patch: { numFmt: "0.0000" }
		})

		const saved = await savedBytes(session)
		const styles = parts(saved)["xl/styles.xml"] ?? ""

		session.close()

		// A format appended: the count is written as what the part holds.
		expect(styles).toContain('<numFmts count="3">')
		expect((await readSaved(saved, bytes)).problems).toEqual([])
	})

	it("finds nothing after an edit is undone on a source whose counts are wrong", async () => {
		const { session, info, bytes } = await openFixture({
			sheets: [{ name: "One", body: '<sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData>' }],
			styles: WRONG_COUNT
		})
		const one = info.sheets[0]!.id

		await firstTile(session, one)
		await session.intent()
		await session.command({
			type: "format.apply",
			sheet: one,
			ranges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }],
			patch: { bold: true, numFmt: "0.0000" }
		})
		await session.undo()

		const saved = await savedBytes(session)

		session.close()
		expect(parts(saved)["xl/styles.xml"]).toBe(parts(bytes)["xl/styles.xml"])
		expect((await readSaved(saved, bytes)).problems).toEqual([])
	})
})
