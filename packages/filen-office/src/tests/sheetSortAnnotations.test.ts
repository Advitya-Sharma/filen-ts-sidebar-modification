import { describe, expect, it } from "vitest"
import { buildFixture, type Fixture } from "./xlsxWriteFixtures"
import { TYPES, relationship } from "./xlsxRefsFixtures"
import { VML_HEAD, commentsPart, entries, noteShape, openAnnotated, withWorkbookParts } from "./sheetNotesFixtures"

// A sort takes each moved line's notes (comment, shape and box), threaded comments, hyperlinks and merged cells with
// it, as Excel does, and an undo gives the parts back byte for byte. Merged cells of different sizes, a merge over
// more than one moved line and a hyperlink over more than one moved line refuse.

const TC = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
const SML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
const ROOT = "{FB8EA27E-C1B6-4481-B034-193455CC7425}"
const REPLY = "{789BBBE8-38C6-4228-9338-E9C75EB78FCF}"
const ANN = "{4A73A60B-EAFA-41CF-BAF3-B061059D7A5C}"
const HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

// Keys d, c, b, a in A2:A5 under a header: an ascending sort moves row 5 to 2, 4 to 3, 3 to 4 and 2 to 5.
function rowsXml(): string {
	const keys = ["d", "c", "b", "a"]

	return `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Key</t></is></c></row>${keys
		.map(
			(key, index) =>
				`<row r="${String(index + 2)}"><c r="A${String(index + 2)}" t="inlineStr"><is><t>${key}</t></is></c><c r="B${String(index + 2)}"><v>${String(index)}</v></c></row>`
		)
		.join("")}</sheetData>`
}

function workbook(tail: string): Uint8Array {
	const fixture: Fixture = {
		sheets: [
			{
				name: "One",
				body: `${rowsXml()}${tail}<legacyDrawing r:id="rId1"/>`,
				rels:
					relationship("rId1", "vmlDrawing", "../drawings/vmlDrawing1.vml") +
					relationship("rId2", "comments", "../comments1.xml") +
					'<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>' +
					`<Relationship Id="rId4" Type="${HYPERLINK}" Target="https://filen.io/" TargetMode="External"/>`
			}
		],
		parts: {
			"xl/threadedComments/threadedComment1.xml": {
				text:
					`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<ThreadedComments xmlns="${TC}" xmlns:x="${SML}">` +
					`<threadedComment ref="A3" dT="2026-03-24T09:25:49.80" personId="${ANN}" id="${ROOT}"><text>on c</text></threadedComment>` +
					`<threadedComment ref="A3" dT="2026-04-07T14:14:05.49" personId="${ANN}" id="${REPLY}" parentId="${ROOT}"><text>a reply</text></threadedComment>` +
					"</ThreadedComments>",
				type: TYPES.threaded
			},
			"xl/comments1.xml": {
				text: commentsPart(
					["Ann", `tc=${ROOT}`],
					[
						{ ref: "A2", author: 0, text: "note on d" },
						{ ref: "A3", author: 1, text: "placeholder" }
					]
				),
				type: TYPES.comments
			},
			"xl/drawings/vmlDrawing1.vml": {
				text: `${VML_HEAD}${noteShape(1025, 1, 0, true)}${noteShape(1026, 2, 0)}</xml>`,
				type: TYPES.vml
			}
		}
	}

	return withWorkbookParts(
		buildFixture(fixture),
		'<Relationship Id="rIdPersons" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/>',
		{
			"xl/persons/person.xml": {
				text: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<personList xmlns="${TC}" xmlns:x="${SML}"><person displayName="Ann" id="${ANN}" userId="Ann" providerId="None"/></personList>`,
				type: "application/vnd.ms-excel.person+xml"
			}
		}
	)
}

const LINKS = '<hyperlinks><hyperlink ref="A4" r:id="rId4"/><hyperlink ref="B5:C5" location="One!A1" display="top"/></hyperlinks>'
const MERGES = '<mergeCells count="2"><mergeCell ref="B2:C2"/><mergeCell ref="B3:C3"/></mergeCells>'

const SORT = {
	type: "sort.apply",
	range: { startRow: 0, startCol: 0, endRow: 4, endCol: 2 },
	orientation: "rows",
	header: true,
	caseSensitive: false,
	keys: [{ index: 0, descending: false }]
} as const

const PARTS = ["xl/worksheets/sheet1.xml", "xl/comments1.xml", "xl/drawings/vmlDrawing1.vml", "xl/threadedComments/threadedComment1.xml"]

describe("a sort's notes, threaded comments, hyperlinks and merged cells", () => {
	it("go with their lines, and an undo gives every part back byte for byte", async () => {
		const bytes = workbook(`${MERGES}${LINKS}`)
		const { session, sheets } = await openAnnotated(bytes)
		const sheet = sheets[0]!
		const sorted = await session.command({ ...SORT, sheet })

		expect(sorted.ok).toBe(true)

		const saved = await session.save()

		if (!saved.ok) throw new Error(saved.reason)

		const files = entries(saved.bytes)
		const sheetXml = files["xl/worksheets/sheet1.xml"] ?? ""

		expect(sheetXml).toContain('<mergeCell ref="B5:C5"/>')
		expect(sheetXml).toContain('<mergeCell ref="B4:C4"/>')
		expect(sheetXml).toContain('<hyperlink ref="A3" r:id="rId4"/>')
		expect(sheetXml).toContain('<hyperlink ref="B2:C2" location="One!A1" display="top"/>')
		expect(files["xl/comments1.xml"]).toContain('<comment ref="A5" authorId="0"')
		expect(files["xl/comments1.xml"]).toContain('<comment ref="A4" authorId="1"')
		expect(files["xl/threadedComments/threadedComment1.xml"]?.match(/ref="A4"/g)?.length).toBe(2)

		const vml = files["xl/drawings/vmlDrawing1.vml"] ?? ""

		// The shown note's box goes down three rows with its cell; the thread's placeholder one.
		expect(vml).toContain(
			"<x:Anchor>1, 15, 3, 2, 3, 15, 7, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>4</x:Row><x:Column>0</x:Column><x:Visible/>"
		)
		expect(vml).toContain(
			"<x:Anchor>1, 15, 1, 2, 3, 15, 5, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>3</x:Row><x:Column>0</x:Column>"
		)

		expect((await session.query("note.get", { sheet, row: 4, col: 0 }))?.text).toContain("note on d")
		expect(await session.query("note.get", { sheet, row: 1, col: 0 })).toBeNull()
		expect((await session.query("thread.get", { sheet, row: 3, col: 0 }))?.comments.length).toBe(2)
		expect((await session.query("link.get", { sheet, row: 2, col: 0 }))?.url).toBe("https://filen.io/")
		expect((await session.query("link.get", { sheet, row: 1, col: 2 }))?.location).toBe("One!A1")

		expect((await session.undo()).ok).toBe(true)

		const undone = await session.save()

		if (!undone.ok) throw new Error(undone.reason)

		for (const part of PARTS) expect(entries(undone.bytes)[part]).toBe(entries(bytes)[part])

		expect((await session.redo()).ok).toBe(true)

		const redone = await session.save()

		if (!redone.ok) throw new Error(redone.reason)

		for (const part of PARTS) expect(entries(redone.bytes)[part]).toBe(files[part])
	})

	it("refuse merged cells of different sizes, a merge over two moved lines and a hyperlink over two moved lines", async () => {
		const refusal = async (tail: string): Promise<string | null> => {
			const { session, sheets } = await openAnnotated(workbook(tail))
			const result = await session.command({ ...SORT, sheet: sheets[0]! })

			return result.ok ? null : `${result.refusal.reason}${result.refusal.detail === undefined ? "" : `/${result.refusal.detail}`}`
		}

		expect(await refusal('<mergeCells count="2"><mergeCell ref="B2:C2"/><mergeCell ref="B4:B4"/></mergeCells>')).toBe(
			"merged/mergeSize"
		)
		expect(await refusal('<mergeCells count="1"><mergeCell ref="B2:B3"/></mergeCells>')).toBe("merged")
		// Reaching past the sorted columns: the merge cannot go whole.
		expect(await refusal('<mergeCells count="1"><mergeCell ref="C2:D2"/></mergeCells>')).toBe("merged")
		expect(await refusal('<hyperlinks><hyperlink ref="A2:A3" location="One!A1"/></hyperlinks>')).toBe("structureLocked/hyperlinks")
		// A header's merge and a link on a line that stays are left alone.
		expect(await refusal('<mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells>')).toBeNull()
	})
})
