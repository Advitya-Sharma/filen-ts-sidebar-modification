import { describe, expect, it } from "vitest"
import { CellFlag } from "../core/sheet/protocol"
import { buildFixture, firstTile, savedBytes, tileEntry } from "./sheetSessionFixtures"
import { NOTES, changeOf, entries, openAnnotated, range, refusalOf, settle } from "./sheetNotesFixtures"
import { TYPES, relationship } from "./xlsxRefsFixtures"

// Notes: their corner flag, their text, and their edits through the comments part and the legacy drawing, saved
// through the gate.

function flags(tile: Awaited<ReturnType<typeof firstTile>>, row: number, col: number): number {
	return tileEntry(tile, row, col)?.flags ?? 0
}

describe("sheet notes", () => {
	it("marks the cells with notes once the parts beside the sheet are read, and says what they hold", async () => {
		const { session, sheets } = await openAnnotated(NOTES)

		await settle()

		const tile = await firstTile(session, sheets[0] ?? 0)

		expect(flags(tile, 1, 1) & CellFlag.NOTE).toBe(CellFlag.NOTE)
		// D5 has no value: its entry is there for the note.
		expect(flags(tile, 4, 3) & CellFlag.NOTE).toBe(CellFlag.NOTE)
		expect(flags(tile, 0, 0) & CellFlag.NOTE).toBe(0)
		expect(await session.query("note.get", { sheet: sheets[0] ?? 0, row: 1, col: 1 })).toEqual({
			text: "Ann:\nfirst note",
			author: "Ann",
			shown: false
		})
		expect((await session.query("note.get", { sheet: sheets[0] ?? 0, row: 4, col: 3 }))?.shown).toBe(true)
		expect([...((await session.layout(sheets[0] ?? 0)).extensions?.shownNotes ?? [])]).toEqual([4, 3])
		session.close()
	})

	it("patches the first tile when the notes land after it", async () => {
		const { session, sheets, events } = await openAnnotated(NOTES)
		const tile = await firstTile(session, sheets[0] ?? 0)

		await settle()

		const patched = events.flatMap(event => (event.type === "change" ? event.change.patches : [])).find(patch => patch.tileRow === 0)

		expect(flags(patched ?? tile, 1, 1) & CellFlag.NOTE).toBe(CellFlag.NOTE)
		session.close()
	})

	it("adds a note to a sheet with none: comments part, legacy drawing, relationships and legacyDrawing, saved and read back", async () => {
		const { session, sheets } = await openAnnotated(NOTES)
		const sheet = sheets[1] ?? 0

		changeOf(await session.command({ type: "note.set", sheet, row: 2, col: 1, text: "Jan:\nnew & <one>" }))
		expect(flags(await firstTile(session, sheet), 2, 1) & CellFlag.NOTE).toBe(CellFlag.NOTE)

		const saved = await savedBytes(session)
		const files = entries(saved)
		const rels = files["xl/worksheets/_rels/sheet2.xml.rels"] ?? ""

		expect(rels).toMatch(/Type="[^"]*\/comments" Target="\.\.\/comments2\.xml"/)
		expect(rels).toMatch(/Type="[^"]*\/vmlDrawing" Target="\.\.\/drawings\/vmlDrawing2\.vml"/)
		expect(files["xl/worksheets/sheet2.xml"]).toMatch(/<legacyDrawing xmlns:r="[^"]*" r:id="rId\d+"\/><\/worksheet>$/)
		expect(files["xl/comments2.xml"]).toContain('<comment ref="B3" authorId="0" shapeId="0" xr:uid="{')
		expect(files["xl/comments2.xml"]).toContain("<authors><author>Jan</author></authors>")
		expect(files["xl/comments2.xml"]).toContain("<t>Jan:\nnew &amp; &lt;one&gt;</t>")
		// The first legacy drawing has block 1: the new one takes 2.
		expect(files["xl/drawings/vmlDrawing2.vml"]).toContain('<o:idmap v:ext="edit" data="2"/>')
		expect(files["xl/drawings/vmlDrawing2.vml"]).toContain('id="_x0000_s2049"')
		expect(files["xl/drawings/vmlDrawing2.vml"]).toContain("<x:Row>2</x:Row><x:Column>1</x:Column>")
		expect(files["[Content_Types].xml"]).toContain(
			'PartName="/xl/comments2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"'
		)
		// Untouched notes keep their bytes.
		expect(files["xl/comments1.xml"]).toBe(entries(NOTES_BYTES)["xl/comments1.xml"])

		const reopened = await openAnnotated(saved)

		expect((await reopened.session.query("note.get", { sheet: reopened.sheets[1] ?? 0, row: 2, col: 1 }))?.text).toBe(
			"Jan:\nnew & <one>"
		)
		reopened.session.close()

		changeOf(await session.undo())
		expect(await savedBytes(session)).toEqual(await savedBytes((await openAnnotated(NOTES)).session))
		session.close()
	})

	it("edits a note's text in place, keeping its author's run and the other notes as written", async () => {
		const { session, sheets } = await openAnnotated(NOTES)
		const sheet = sheets[0] ?? 0

		changeOf(await session.command({ type: "note.set", sheet, row: 1, col: 1, text: "Ann:\nchanged" }))

		const comments = entries(await savedBytes(session))["xl/comments1.xml"] ?? ""

		expect(comments).toContain(
			'<comment ref="B2" authorId="0" shapeId="0"><text><r><rPr><b/><sz val="9"/></rPr><t>Ann:</t></r><r><rPr><sz val="9"/></rPr><t xml:space="preserve">\nchanged</t></r></text></comment>'
		)
		expect(comments).toContain('<comment ref="D5" authorId="0" shapeId="0"><text><r><rPr><b/><sz val="9"/></rPr><t>Ann:</t></r>')
		// No new shape for a note that has one.
		expect(entries(await savedBytes(session))["xl/drawings/vmlDrawing1.vml"]).toBe(entries(NOTES_BYTES)["xl/drawings/vmlDrawing1.vml"])
		session.close()
	})

	it("adds a note beside others: a shape in the sheet's drawing with the next id, the note in cell order", async () => {
		const { session, sheets } = await openAnnotated(NOTES)
		const sheet = sheets[0] ?? 0

		changeOf(await session.command({ type: "note.set", sheet, row: 2, col: 0, text: "between" }))

		const files = entries(await savedBytes(session))

		expect(files["xl/comments1.xml"]).toMatch(/ref="B2".*ref="A3".*ref="D5"/s)
		expect(files["xl/comments1.xml"]).toContain("<authors><author>Ann</author><author>Jan</author></authors>")
		expect(files["xl/drawings/vmlDrawing1.vml"]).toContain('id="_x0000_s1027"')
		expect(files["xl/worksheets/sheet1.xml"]).toBe(entries(NOTES_BYTES)["xl/worksheets/sheet1.xml"])
		session.close()
	})

	it("deletes notes with their shapes, and undo saves the source again", async () => {
		const { session, sheets } = await openAnnotated(NOTES)
		const sheet = sheets[0] ?? 0
		const change = changeOf(await session.command({ type: "note.delete", sheet, ranges: [range(0, 0, 10, 10)] }))

		expect(change.layouts.some(layout => layout.sheet === sheet)).toBe(true)

		const files = entries(await savedBytes(session))

		expect(files["xl/comments1.xml"]).not.toContain("<comment ")
		expect(files["xl/drawings/vmlDrawing1.vml"]).not.toContain('ObjectType="Note"')
		expect(flags(await firstTile(session, sheet), 1, 1) & CellFlag.NOTE).toBe(0)

		changeOf(await session.undo())
		expect(entries(await savedBytes(session))["xl/comments1.xml"]).toBe(entries(NOTES_BYTES)["xl/comments1.xml"])
		expect(flags(await firstTile(session, sheet), 1, 1) & CellFlag.NOTE).toBe(CellFlag.NOTE)
		session.close()
	})

	it("shows a note all the time and hides it again", async () => {
		const { session, sheets } = await openAnnotated(NOTES)
		const sheet = sheets[0] ?? 0
		const change = changeOf(await session.command({ type: "note.show", sheet, row: 1, col: 1, shown: true }))

		expect([...(change.layouts.find(layout => layout.sheet === sheet)?.layout.extensions?.shownNotes ?? [])]).toEqual([1, 1, 4, 3])

		const vml = entries(await savedBytes(session))["xl/drawings/vmlDrawing1.vml"] ?? ""

		expect(vml).toContain("z-index:1;visibility:visible")
		expect(vml).toContain("<x:Row>1</x:Row><x:Column>1</x:Column><x:Visible/></x:ClientData>")

		changeOf(await session.command({ type: "note.show", sheet, row: 4, col: 3, shown: false }))
		expect(entries(await savedBytes(session))["xl/drawings/vmlDrawing1.vml"]).not.toContain(
			"<x:Row>4</x:Row><x:Column>3</x:Column><x:Visible/>"
		)
		session.close()
	})

	it("moves a note it added when rows are inserted above it", async () => {
		const { session, sheets } = await openAnnotated(NOTES)
		const sheet = sheets[1] ?? 0

		changeOf(await session.command({ type: "note.set", sheet, row: 2, col: 1, text: "moves" }))
		changeOf(await session.command({ type: "axis.insert", sheet, axis: "rows", at: 0, count: 2 }))

		const files = entries(await savedBytes(session))

		expect(files["xl/comments2.xml"]).toContain('ref="B5"')
		expect(files["xl/drawings/vmlDrawing2.vml"]).toContain("<x:Row>4</x:Row>")
		expect((await session.query("note.get", { sheet, row: 4, col: 1 }))?.text).toBe("moves")
		session.close()
	})

	it("writes notes, comments and links on a sheet added in the session", async () => {
		const { session } = await openAnnotated(NOTES)
		const added = changeOf(await session.command({ type: "sheet.add", name: "Fresh" }))
		const sheet = added.sheets?.find(summary => summary.name === "Fresh")?.id ?? -1

		changeOf(await session.command({ type: "note.set", sheet, row: 0, col: 0, text: "on a new sheet" }))
		changeOf(await session.command({ type: "thread.add", sheet, row: 1, col: 1, text: "a thread" }))
		changeOf(await session.command({ type: "link.set", sheet, ranges: [range(2, 2)], target: { url: "https://filen.io/" } }))

		const reopened = await openAnnotated(await savedBytes(session))
		const fresh = reopened.sheets[2] ?? -1

		expect((await reopened.session.query("note.get", { sheet: fresh, row: 0, col: 0 }))?.text).toBe("on a new sheet")
		expect((await reopened.session.query("thread.get", { sheet: fresh, row: 1, col: 1 }))?.comments[0]?.text).toBe("a thread")
		expect((await reopened.session.query("link.get", { sheet: fresh, row: 2, col: 2 }))?.url).toBe("https://filen.io/")
		reopened.session.close()
		session.close()
	})

	it("edits the notes of a duplicated sheet in its own parts, leaving the source sheet's", async () => {
		const { session, sheets } = await openAnnotated(NOTES)
		const duplicated = changeOf(await session.command({ type: "sheet.duplicate", sheet: sheets[0] ?? 0 }))
		const copy = duplicated.sheets?.find(summary => !sheets.includes(summary.id))?.id ?? -1

		await settle()
		expect((await session.query("note.get", { sheet: copy, row: 1, col: 1 }))?.text).toBe("Ann:\nfirst note")
		changeOf(await session.command({ type: "note.set", sheet: copy, row: 1, col: 1, text: "Ann:\nin the copy" }))
		changeOf(await session.command({ type: "note.set", sheet: copy, row: 6, col: 0, text: "new in the copy" }))

		const files = entries(await savedBytes(session))

		expect(files["xl/comments1.xml"]).toBe(entries(NOTES_BYTES)["xl/comments1.xml"])
		expect(Object.entries(files).some(([name, text]) => name.startsWith("xl/comments") && text.includes("in the copy"))).toBe(true)
		expect((await session.query("note.get", { sheet: sheets[0] ?? 0, row: 1, col: 1 }))?.text).toBe("Ann:\nfirst note")
		session.close()
	})

	it("asks for an author before a note is written", async () => {
		const { session, sheets } = await openAnnotated(NOTES, {})

		expect(refusalOf(await session.command({ type: "note.set", sheet: sheets[1] ?? 0, row: 0, col: 0, text: "x" }))).toBe(
			"authorNeeded"
		)
		session.close()
	})

	it("refuses note edits on a sheet whose comments part does not read, shows none, reports it and saves the file as it was", async () => {
		const damaged = buildFixture({
			...NOTES,
			parts: { ...NOTES.parts, "xl/comments1.xml": { text: "<comments><authors>", type: TYPES.comments } }
		})
		const { session, sheets, events } = await openAnnotated(damaged)
		const sheet = sheets[0] ?? 0

		await settle()

		expect(events).toContainEqual(expect.objectContaining({ type: "featureUnavailable", sheet, feature: "notes" }))
		expect(flags(await firstTile(session, sheet), 1, 1) & CellFlag.NOTE).toBe(0)
		expect(await session.query("note.get", { sheet, row: 1, col: 1 })).toBeNull()
		expect(refusalOf(await session.command({ type: "note.set", sheet, row: 0, col: 0, text: "x" }))).toBe("notesUnavailable")
		expect(refusalOf(await session.command({ type: "note.delete", sheet, ranges: [range(1, 1)] }))).toBe("notesUnavailable")
		expect((await session.query("commandState", { sheet, ranges: [range(0, 0)], types: ["note.set"] })).reasons).toEqual([
			"notesUnavailable"
		])
		// Another sheet's notes are not affected.
		changeOf(await session.command({ type: "note.set", sheet: sheets[1] ?? 0, row: 0, col: 0, text: "fine" }))
		changeOf(await session.undo())
		expect(await savedBytes(session)).toEqual(damaged)
		session.close()
	})

	it("keeps a form control's legacy drawing and adds note shapes beside it", async () => {
		const fixture = {
			sheets: [
				{
					name: "One",
					body: '<sheetData/><legacyDrawing r:id="rId1"/>',
					rels: relationship("rId1", "vmlDrawing", "../drawings/vmlDrawing1.vml")
				}
			],
			parts: {
				"xl/drawings/vmlDrawing1.vml": {
					text:
						'<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
						'<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="3"/></o:shapelayout>' +
						'<v:shape id="_x0000_s3073" type="#_x0000_t201"><x:ClientData ObjectType="Button"><x:Anchor>1, 0, 1, 0, 2, 0, 2, 0</x:Anchor></x:ClientData></v:shape></xml>',
					type: TYPES.vml
				}
			}
		}
		const { session, sheets } = await openAnnotated(fixture)

		changeOf(await session.command({ type: "note.set", sheet: sheets[0] ?? 0, row: 0, col: 0, text: "beside" }))

		const vml = entries(await savedBytes(session))["xl/drawings/vmlDrawing1.vml"] ?? ""

		expect(vml).toContain('ObjectType="Button"')
		expect(vml).toContain('<v:shapetype id="_x0000_t202"')
		expect(vml).toContain('id="_x0000_s3074"')
		session.close()
	})
})

const NOTES_BYTES = buildFixture(NOTES)
