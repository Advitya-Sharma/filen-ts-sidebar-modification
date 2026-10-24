import { strFromU8, unzipSync } from "fflate"
import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { readRtf } from "../core/rtf/index"
import { RtfWriteRefusal } from "../core/rtf/write/out"
import { editRtf, latin1, paragraphs, saveRtf, type RtfEditing } from "./rtfWriteFixtures"

// Comments as RTF annotations (RTF 1.9.1, "Comments (Annotations)"): {\*\atrfstart N} and {\*\atrfend N} around the
// range, then {\*\atnid}{\*\atnauthor}\chatn{\*\annotation{\*\atnref N}{\*\atndate}...} where the reference is.

const HEADER = "{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0\\froman Times New Roman;}}"

const PLAIN = latin1(`${HEADER}\\pard\\plain\\f0\\fs24 First paragraph here\\par\n\\pard\\plain\\f0\\fs24 Second paragraph here\\par\n}`)

// A comment as Word writes one, on the second paragraph.
const COMMENTED = latin1(
	HEADER +
		"\\pard\\plain\\f0\\fs24 First paragraph here\\par\n" +
		"\\pard\\plain\\f0\\fs24 Second {\\*\\atrfstart 3}commented{\\*\\atrfend 3} words" +
		"{\\*\\atnid AL}{\\*\\atnauthor Ann Lee}\\chatn{\\*\\annotation{\\*\\atnref 3}{\\*\\atndate 1891047488}" +
		"\\pard\\plain {\\chatn}Old note\\par}\\par\n}"
)

async function parts(bytes: Uint8Array): Promise<{ body: string; comments: string }> {
	const read = await readRtf(bytes)

	if (!read.ok) throw new Error(read.reason)

	const zip = unzipSync(read.docx)

	return {
		body: strFromU8(zip["word/document.xml"] ?? new Uint8Array()),
		comments: strFromU8(zip["word/comments.xml"] ?? new Uint8Array())
	}
}

// Each comment: author, initials, date and text; and the text each range covers, in reading order.
function comments(read: { body: string; comments: string }): { comments: string[]; ranges: string[] } {
	const list = [...read.comments.matchAll(/<w:comment ([^>]*)>(.*?)<\/w:comment>/gsu)].map(match => {
		const attribute = (name: string): string => new RegExp(`${name}="([^"]*)"`, "u").exec(match[1] ?? "")?.[1] ?? ""
		const text = [...(match[2] ?? "").matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/gu)].map(t => t[1]).join("")

		return `${attribute("w:author")}|${attribute("w:initials")}|${attribute("w:date")}|${text}`
	})
	const ranges = [...read.body.matchAll(/<w:commentRangeStart w:id="(\d+)"\/>(.*?)<w:commentRangeEnd w:id="\1"\/>/gsu)].map(match =>
		[...(match[2] ?? "").matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/gu)].map(t => t[1]).join("")
	)

	return { comments: list, ranges }
}

async function comment(editing: RtfEditing, paragraph: number, from: number, to: number, text: string): Promise<void> {
	const session = editing.session
	const at = paragraphs(session.state.doc)[paragraph]

	if (at === undefined) throw new Error("no paragraph")

	session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, at.pos + 1 + from, at.pos + 1 + to)))
	expect(await editing.editor.execute({ type: "comment.add", text })).toBe(true)
}

describe("RTF comments", () => {
	it("writes a new comment as an annotation that reads back with its range, author, date and text", async () => {
		const editing = await editRtf(PLAIN)

		await comment(editing, 0, 6, 15, "A new note")

		const written = (await saveRtf(editing)).bytes
		const out = strFromU8(written, true)

		expect(out).toContain("{\\*\\atrfstart 1}")
		expect(out).toMatch(/\{\\\*\\atnauthor Filen\}\\chatn\{\\\*\\annotation\{\\\*\\atnref 1\}\{\\\*\\atndate \d+\}/u)
		expect(out).toContain("Second paragraph here\\par\n}")

		const read = comments(await parts(written))

		expect(read.ranges).toEqual(["paragraph"])
		expect(read.comments).toHaveLength(1)
		expect(read.comments[0]).toMatch(/^Filen\|\|\d{4}-\d\d-\d\dT\d\d:\d\d:00Z\|A new note$/u)
		expect((await editRtf(written)).locked).toBe(0)
	})

	it("keeps a source comment and numbers a new one before it past the source's range text", async () => {
		const editing = await editRtf(COMMENTED)

		expect(editing.locked).toBe(0)

		await comment(editing, 0, 0, 5, "Before")

		const written = (await saveRtf(editing)).bytes
		const out = strFromU8(written, true)
		const read = comments(await parts(written))

		expect(out).toContain("{\\*\\atrfstart 4}")
		expect(out).toContain("{\\*\\atrfstart 3}commented{\\*\\atrfend 3}")
		expect(read.ranges).toEqual(["First", "commented"])
		expect(read.comments.map(entry => entry.replace(/\|[^|]*\|(?=[^|]*$)/u, "||"))).toEqual(["Filen|||Before", "Ann Lee|AL||Old note"])
	})

	it("writes an edited paragraph holding a source comment, the comment kept", async () => {
		const editing = await editRtf(COMMENTED)
		const session = editing.session
		const second = paragraphs(session.state.doc)[1]

		if (second === undefined) throw new Error("no paragraph")

		session.dispatch(session.state.tr.insertText("New ", second.pos + 1))

		const written = await saveRtf(editing)
		const read = comments(await parts(written.bytes))

		expect(written.written).toBe(1)
		expect(read.ranges).toEqual(["commented"])
		expect(read.comments).toEqual(comments(await parts(COMMENTED)).comments)
		expect((await parts(written.bytes)).body).toContain("New Second ")
	})

	it("refuses a reply and a resolved comment, which an annotation cannot hold", async () => {
		for (const command of ["reply", "resolve"] as const) {
			const editing = await editRtf(COMMENTED)
			const id = /<w:comment w:id="(\d+)"/u.exec((await parts(COMMENTED)).comments)?.[1] ?? ""

			expect(
				await editing.editor.execute(
					command === "reply" ? { type: "comment.reply", id, text: "An answer" } : { type: "comment.resolve", id }
				)
			).toBe(true)

			const refusal = await saveRtf(editing).catch((error: unknown) => error)

			expect(refusal).toBeInstanceOf(RtfWriteRefusal)
			expect((refusal as RtfWriteRefusal).reason).toBe("comment")
			expect((refusal as RtfWriteRefusal).message).toContain(command === "reply" ? "a reply" : "a resolved comment")
		}
	})
})
