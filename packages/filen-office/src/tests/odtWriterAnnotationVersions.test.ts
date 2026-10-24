import { describe, expect, it } from "vitest"
import { TextSelection } from "prosemirror-state"
import { ZipArchive, bytesSource } from "@filen/office/core/opc"
import { decodeUtf8, encodeUtf8 } from "@filen/office/core/xml"
import type { OdtWritten } from "../core/odt/write/writer"
import { buildZip } from "./zipBuilder"
import { odtBytes } from "./odtFixtures"
import { docxRtfServer } from "../dom/docx/rtfWorker"
import { editOdt, odtEditsOf, saveOdt } from "./odtWriteFixtures"

// A comment added to an ODF text document is written in the forms its office:version has: an author's initials as
// meta:creator-initials from ODF 1.3 (LibreOffice's loext:sender-initials before it), a range (office:name and
// office:annotation-end) from ODF 1.2. An ODF 1.0 or 1.1 annotation marks a point: a comment over a range there is
// kept at the start of its text, and the save says so.

// The document with office:version `version` in every part (undefined: none, which reads as 1.1).
async function versioned(version: string | undefined, first = "Plain para words"): Promise<Uint8Array> {
	const archive = await ZipArchive.open(bytesSource(odtBytes({ body: `<text:p>${first}</text:p><text:p>Last</text:p>` })))
	const entries = []

	for (const entry of archive.entries) {
		const data = await archive.read(entry)
		const text = entry.name.endsWith(".xml") ? decodeUtf8(data) : undefined

		entries.push({
			name: entry.name,
			data:
				text === undefined
					? data
					: encodeUtf8(
							text.replace(/ (office|manifest):version="1\.3"/gu, version === undefined ? "" : ` $1:version="${version}"`)
						),
			method: entry.method === 0 ? (0 as const) : (8 as const)
		})
	}

	return buildZip(entries)
}

async function commented(bytes: Uint8Array, length: number): Promise<string> {
	return (await commentedSave(bytes, length)).xml
}

async function commentedSave(bytes: Uint8Array, length: number): Promise<{ xml: string; written: OdtWritten }> {
	const editing = await editOdt(bytes)
	// Inside the first paragraph.
	const at = 1

	editing.opened.context.identity.author = "Ann Lee"
	editing.opened.context.identity.initials = "AL"
	editing.session.dispatch(editing.session.state.tr.setSelection(TextSelection.create(editing.session.state.doc, at, at + length)))
	await editing.editor.execute({ type: "comment.add", text: "Note" } as never)

	const saved = await saveOdt(editing)
	const written = saved.bytes
	const archive = await ZipArchive.open(bytesSource(written))
	const entry = archive.get("content.xml")

	if (entry === undefined) throw new Error("no content.xml")

	const xml = decodeUtf8(await archive.read(entry))

	expect((await editOdt(written)).locked).toBe(0)

	return { xml, written: saved }
}

describe("a comment added to an ODF text document", () => {
	it("records its author's initials as meta:creator-initials in ODF 1.3", async () => {
		const xml = await commented(await versioned("1.3"), 5)

		expect(xml).toMatch(
			/<office:annotation office:name="[^"]+"><dc:creator>Ann Lee<\/dc:creator>.*<meta:creator-initials>AL<\/meta:creator-initials>/u
		)
		expect(xml).toContain("<office:annotation-end ")
	})

	it("records them as loext:sender-initials before ODF 1.3, a range still named in ODF 1.2", async () => {
		const xml = await commented(await versioned("1.2"), 5)

		expect(xml).not.toContain("meta:creator-initials")
		expect(xml).toContain(
			'<loext:sender-initials office:process-content="false" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0">AL</loext:sender-initials>'
		)
		expect(xml).toMatch(/<office:annotation office:name="[^"]+">/u)
		expect(xml).toContain("<office:annotation-end ")
	})

	it.each([["1.1"], ["1.0"], [undefined]])("marks a point, unnamed, in ODF %s", async version => {
		// In an empty paragraph: no word around the caret for the comment to cover.
		const xml = await commented(await versioned(version, ""), 0)

		expect(xml).toMatch(/<office:annotation><dc:creator>Ann Lee<\/dc:creator>/u)
		expect(xml).not.toContain("office:annotation-end")
		expect(xml).not.toContain("office:name")
		expect(xml).not.toContain("meta:creator-initials")
		expect(xml).toMatch(/<loext:sender-initials office:process-content="false" [^>]*>AL<\/loext:sender-initials>/u)
	})

	it.each([["1.1"], ["1.0"]])("over a range is kept at the start of its text in ODF %s, and the save says so", async version => {
		const { xml, written } = await commentedSave(await versioned(version), 5)

		expect(xml).toMatch(
			/<text:p><office:annotation><dc:creator>Ann Lee<\/dc:creator>.*?<\/office:annotation>Plain para words<\/text:p>/u
		)
		expect(xml).not.toContain("office:annotation-end")
		expect(xml).not.toContain("office:name")
		expect(written.notices).toEqual([{ kind: "commentPoint", count: 1 }])
	})

	it("says nothing of comments a save keeps as they are", async () => {
		expect((await commentedSave(await versioned("1.2"), 5)).written.notices ?? []).toEqual([])
		expect((await commentedSave(await versioned("1.1", ""), 0)).written.notices ?? []).toEqual([])
	})

	it("reaches the editor's save from where the document is held", async () => {
		const bytes = await versioned("1.1")
		const editing = await editOdt(bytes)

		editing.opened.context.identity.author = "Ann Lee"
		editing.session.dispatch(editing.session.state.tr.setSelection(TextSelection.create(editing.session.state.doc, 1, 6)))
		await editing.editor.execute({ type: "comment.add", text: "Note" } as never)

		const edits = await odtEditsOf(editing)
		const server = docxRtfServer()

		await server.open(1, bytes, "odt")

		const written = await server.write(1, { parts: [...edits.parts], removed: [...(edits.removed ?? [])] })

		server.close(1)
		expect(written.ok && written.notices).toEqual([{ kind: "commentPoint", count: 1 }])
	})
})
