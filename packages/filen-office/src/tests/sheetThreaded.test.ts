import { describe, expect, it } from "vitest"
import { CellFlag } from "../core/sheet/protocol"
import { commentTime } from "../core/xlsx/threaded/edit"
import { placeholderText } from "../core/xlsx/threaded/threads"
import { buildFixture, firstTile, savedBytes, tileEntry } from "./sheetSessionFixtures"
import { NOTES, changeOf, entries, openAnnotated, range, refusalOf, settle, withWorkbookParts } from "./sheetNotesFixtures"
import { TYPES, relationship } from "./xlsxRefsFixtures"

// Threaded comments: their corner, the thread as asked for, new threads and replies under the session's author (a
// person added for it), the author's own comments edited, comments deleted; the placeholder note Excel writes kept to
// the thread.

const TC = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
const SML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
const ROOT = "{FB8EA27E-C1B6-4481-B034-193455CC7425}"
const REPLY = "{789BBBE8-38C6-4228-9338-E9C75EB78FCF}"
const ANN = "{4A73A60B-EAFA-41CF-BAF3-B061059D7A5C}"
const JAN = "{11111111-2222-4333-8444-555555555555}"

// As Excel writes them (lo__threadedComment.xlsx): a thread on A2 with a reply, its placeholder note and shape.
const THREADED = withWorkbookParts(
	buildFixture({
		sheets: [
			{
				name: "One",
				body: '<sheetData><row r="2"><c r="A2"><v>1</v></c></row></sheetData><legacyDrawing r:id="rId1"/>',
				rels:
					relationship("rId1", "vmlDrawing", "../drawings/vmlDrawing1.vml") +
					relationship("rId2", "comments", "../comments1.xml") +
					'<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>'
			}
		],
		parts: {
			"xl/threadedComments/threadedComment1.xml": {
				text:
					`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<ThreadedComments xmlns="${TC}" xmlns:x="${SML}">` +
					`<threadedComment ref="A2" dT="2026-03-24T09:25:49.80" personId="${ANN}" id="${ROOT}"><text>a comment on A2</text></threadedComment>` +
					`<threadedComment ref="A2" dT="2026-04-07T14:14:05.49" personId="${JAN}" id="${REPLY}" parentId="${ROOT}"><text>A reply</text></threadedComment>` +
					"</ThreadedComments>",
				type: TYPES.threaded
			},
			"xl/comments1.xml": {
				text:
					`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="${SML}"><authors><author>tc=${ROOT}</author></authors><commentList>` +
					`<comment ref="A2" authorId="0" shapeId="0"><text><t>${placeholderText(["a comment on A2", "A reply"])}</t></text></comment></commentList></comments>`,
				type: TYPES.comments
			},
			"xl/drawings/vmlDrawing1.vml": {
				text:
					'<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
					'<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>' +
					'<v:shape id="_x0000_s1025" type="#_x0000_t202"><x:ClientData ObjectType="Note"><x:Row>1</x:Row><x:Column>0</x:Column></x:ClientData></v:shape></xml>',
				type: TYPES.vml
			}
		}
	}),
	'<Relationship Id="rIdPersons" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/>',
	{
		"xl/persons/person.xml": {
			text:
				`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<personList xmlns="${TC}" xmlns:x="${SML}">` +
				`<person displayName="Ann" id="${ANN}" userId="5008b16217855fa4" providerId="Windows Live"/>` +
				`<person displayName="Jan" id="${JAN}" userId="Jan" providerId="None"/></personList>`,
			type: "application/vnd.ms-excel.person+xml"
		}
	}
)

describe("threaded comments", () => {
	it("marks a thread's cell with a comment corner, not its placeholder's note corner, and reads the thread", async () => {
		const { session, sheets } = await openAnnotated(THREADED)
		const sheet = sheets[0] ?? 0

		await settle()

		const flags = tileEntry(await firstTile(session, sheet), 1, 0)?.flags ?? 0

		expect(flags & CellFlag.THREADED_COMMENT).toBe(CellFlag.THREADED_COMMENT)
		expect(flags & CellFlag.NOTE).toBe(0)
		expect(await session.query("note.get", { sheet, row: 1, col: 0 })).toBeNull()
		expect(await session.query("thread.get", { sheet, row: 1, col: 0 })).toEqual({
			comments: [
				{ id: ROOT, author: "Ann", date: "2026-03-24T09:25:49.80", text: "a comment on A2", own: false, mentions: false },
				{ id: REPLY, author: "Jan", date: "2026-04-07T14:14:05.49", text: "A reply", own: true, mentions: false }
			],
			done: false
		})
		session.close()
	})

	it("replies under the author's person, keeping the placeholder note to the thread", async () => {
		const { session, sheets } = await openAnnotated(THREADED, {
			author: "Jan",
			clock: () => new Date(2026, 9, 20, 8, 5, 3, 470).getTime()
		})
		const sheet = sheets[0] ?? 0

		changeOf(await session.command({ type: "thread.add", sheet, row: 1, col: 0, text: "and another" }))

		const files = entries(await savedBytes(session))
		const threaded = files["xl/threadedComments/threadedComment1.xml"] ?? ""

		expect(threaded).toMatch(
			new RegExp(
				`id="\\${REPLY}" parentId="\\${ROOT}"><text>A reply</text></threadedComment><threadedComment ref="A2" dT="2026-10-20T08:05:03.47" personId="\\${JAN}" id="\\{[0-9A-F-]{36}\\}" parentId="\\${ROOT}"><text>and another</text>`
			)
		)
		// Jan is a person already: none added.
		expect(files["xl/persons/person.xml"]).toBe(entries(THREADED)["xl/persons/person.xml"])
		expect(files["xl/comments1.xml"]).toContain("Reply:\n    A reply\nReply:\n    and another</t>")
		session.close()
	})

	it("starts a thread on a sheet without one: its part, a person for the author, the placeholder note and shape", async () => {
		const { session, sheets } = await openAnnotated(NOTES, { author: "Zoë" })
		const sheet = sheets[1] ?? 0

		changeOf(await session.command({ type: "thread.add", sheet, row: 0, col: 2, text: "first" }))

		const files = entries(await savedBytes(session))

		expect(files["xl/worksheets/_rels/sheet2.xml.rels"]).toContain(
			'Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"'
		)
		expect(files["xl/_rels/workbook.xml.rels"]).toContain(
			'Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"'
		)
		expect(files["xl/persons/person.xml"]).toMatch(/<person displayName="Zoë" id="\{[0-9A-F-]{36}\}" userId="Zoë" providerId="None"\/>/)
		expect(files["xl/threadedComments/threadedComment1.xml"]).toMatch(
			/<threadedComment ref="C1" dT="[^"]+" personId="\{[0-9A-F-]{36}\}" id="(\{[0-9A-F-]{36}\})"><text>first<\/text>/
		)

		const id = /id="(\{[0-9A-F-]{36}\})"><text>first/.exec(files["xl/threadedComments/threadedComment1.xml"] ?? "")?.[1] ?? ""

		expect(files["xl/comments2.xml"]).toContain(`<author>tc=${id}</author>`)
		expect(files["xl/comments2.xml"]).toContain(`<comment ref="C1" authorId="0" shapeId="0" xr:uid="${id}"><text><t>[Threaded comment]`)
		expect(files["[Content_Types].xml"]).toContain('ContentType="application/vnd.ms-excel.threadedcomments+xml"')
		expect(files["[Content_Types].xml"]).toContain('ContentType="application/vnd.ms-excel.person+xml"')

		const thread = await session.query("thread.get", { sheet, row: 0, col: 2 })

		expect(thread?.comments.map(comment => [comment.author, comment.text, comment.own])).toEqual([["Zoë", "first", true]])

		changeOf(await session.undo())
		expect(await savedBytes(session)).toEqual(await savedBytes((await openAnnotated(NOTES)).session))
		session.close()
	})

	it("edits only the author's own comments", async () => {
		const { session, sheets } = await openAnnotated(THREADED)
		const sheet = sheets[0] ?? 0

		expect(refusalOf(await session.command({ type: "thread.edit", sheet, id: ROOT, text: "not mine" }))).toBe("notOwnComment")
		changeOf(await session.command({ type: "thread.edit", sheet, id: REPLY, text: "A better reply" }))

		const files = entries(await savedBytes(session))

		expect(files["xl/threadedComments/threadedComment1.xml"]).toContain(`parentId="${ROOT}"><text>A better reply</text>`)
		expect(files["xl/comments1.xml"]).toContain("Reply:\n    A better reply</t>")
		session.close()
	})

	it("deletes a reply, then the thread with its placeholder note and shape", async () => {
		const { session, sheets } = await openAnnotated(THREADED)
		const sheet = sheets[0] ?? 0

		changeOf(await session.command({ type: "thread.delete", sheet, id: REPLY }))
		expect((await session.query("thread.get", { sheet, row: 1, col: 0 }))?.comments.length).toBe(1)
		changeOf(await session.command({ type: "thread.delete", sheet, id: ROOT }))

		const files = entries(await savedBytes(session))

		expect(files["xl/threadedComments/threadedComment1.xml"]).not.toContain("<threadedComment ")
		expect(files["xl/comments1.xml"]).not.toContain("<comment ")
		expect(files["xl/drawings/vmlDrawing1.vml"]).not.toContain("ObjectType")
		expect((tileEntry(await firstTile(session, sheet), 1, 0)?.flags ?? 0) & CellFlag.THREADED_COMMENT).toBe(0)

		changeOf(await session.undo())
		changeOf(await session.undo())
		expect(await savedBytes(session)).toEqual(THREADED)
		session.close()
	})

	it("keeps a note and a thread apart on one cell", async () => {
		const { session, sheets } = await openAnnotated(NOTES)

		expect(await session.command({ type: "thread.add", sheet: sheets[0] ?? 0, row: 1, col: 1, text: "x" })).toMatchObject({
			ok: false,
			refusal: { reason: "unsupported", detail: "note" }
		})
		session.close()
	})

	it("writes the time as Excel does: local, to hundredths", () => {
		expect(commentTime(new Date(2026, 0, 2, 3, 4, 5, 69).getTime())).toBe("2026-01-02T03:04:05.06")
	})

	it("refuses thread edits when the threaded comments part does not read", async () => {
		const damaged = withWorkbookParts(
			buildFixture({
				sheets: [
					{
						name: "One",
						body: "<sheetData/>",
						rels: '<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>'
					}
				],
				parts: { "xl/threadedComments/threadedComment1.xml": { text: "<ThreadedComments", type: TYPES.threaded } }
			}),
			"",
			{}
		)
		const { session, sheets, events } = await openAnnotated(damaged)
		const sheet = sheets[0] ?? 0

		await settle()
		expect(events).toContainEqual(expect.objectContaining({ type: "featureUnavailable", feature: "threads" }))
		expect(refusalOf(await session.command({ type: "thread.add", sheet, row: 0, col: 0, text: "x" }))).toBe("notesUnavailable")
		expect(await savedBytes(session)).toEqual(damaged)
		session.close()
	})

	it("leaves a thread's placeholder to the thread when notes are deleted", async () => {
		const { session, sheets } = await openAnnotated(THREADED)

		expect((await session.command({ type: "note.delete", sheet: sheets[0] ?? 0, ranges: [range(1, 0)] })).ok).toBe(true)
		expect(entries(await savedBytes(session))["xl/comments1.xml"]).toBe(entries(THREADED)["xl/comments1.xml"])
		session.close()
	})
})
