import { strFromU8, unzipSync } from "fflate"
import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { readRtf } from "../core/rtf/index"
import { editRtf, latin1, saveRtf, type RtfEditing } from "./rtfWriteFixtures"

// Rewritten RTF paragraphs keep what the source said: formatting removed reads back as removed, properties the
// translation holds are written back, and a paragraph under words it cannot hold stays as it was.

const H = "{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0\\froman Times New Roman;}}"

function find(editing: RtfEditing, text: string): number {
	let at = -1

	editing.session.state.doc.descendants((node, pos) => {
		if (at < 0 && node.isTextblock && node.textContent.includes(text)) at = pos

		return at < 0
	})

	if (at < 0) throw new Error(`no paragraph with ${text}`)

	return at
}

function select(editing: RtfEditing, text: string, from: number, length: number): void {
	const s = editing.session
	const at = find(editing, text)

	s.dispatch(s.state.tr.setSelection(TextSelection.create(s.state.doc, at + 1 + from, at + 1 + from + length)))
}

function type(editing: RtfEditing, text: string, inserted: string): void {
	const s = editing.session

	s.dispatch(s.state.tr.insertText(inserted, find(editing, text) + 1))
}

// The body as rewritten (after the header tables).
async function saved(editing: RtfEditing): Promise<string> {
	return strFromU8((await saveRtf(editing)).bytes, true)
}

async function bodyXml(bytes: Uint8Array): Promise<string> {
	const read = await readRtf(bytes)

	if (!read.ok) throw new Error(read.reason)

	return strFromU8(unzipSync(read.docx)["word/document.xml"] ?? new Uint8Array())
}

describe("RTF writer: formatting removed", () => {
	for (const [name, source, text, command] of [
		["bold from \\pard\\b", `${H}\\pard\\b Title\\b0\\par Body\\par}`, "Title", { type: "format.bold" }],
		["italic in a group", `${H}\\pard One {\\i para}\\par Two\\par}`, "para", { type: "format.italic" }],
		["underline", `${H}\\pard\\ul Under\\par Two\\par}`, "Under", { type: "format.underline" }],
		[
			"colour",
			`${H}{\\colortbl;\\red255\\green0\\blue0;}\\pard\\cf1 Red\\par Two\\par}`,
			"Red",
			{ type: "color.text", color: { kind: "auto" } }
		]
	] as const) {
		it(`saves ${name} turned off on text with no other properties`, async () => {
			const editing = await editRtf(latin1(source))

			const paragraph = editing.session.state.doc.nodeAt(find(editing, text))

			select(editing, text, paragraph?.textContent.indexOf(text) ?? 0, text.length)
			expect(editing.session.execute(command as never)).toBe(true)

			// The run holding the text reads back with no run properties.
			const xml = await bodyXml((await saveRtf(editing)).bytes)

			expect(xml).toMatch(new RegExp(`<w:r><w:t(?: xml:space="preserve")?>[^<]*${text}</w:t></w:r>`, "u"))
		})
	}
})

describe("RTF writer: run and paragraph properties", () => {
	it("keeps a run's languages when its paragraph is rewritten, with no emptied group left behind", async () => {
		const editing = await editRtf(latin1(`${H}\\pard {\\lang1031\\langfe1041\\alang1025 Ein Absatz}\\par Two\\par}`))

		type(editing, "Ein", "X")

		const out = await saved(editing)

		expect(out).toMatch(/\{\\plain[^{}]*\\lang1031\\langnp1031\\langfe1041\\langfenp1041\\alang1025 XEin Absatz\}/u)
		expect(out).not.toContain("{\\lang1031\\langfe1041\\alang1025 }")
	})

	it("starts a paragraph at a formatting group around its text, never at one holding a table", async () => {
		const editing = await editRtf(
			latin1(
				"{\\rtf1\\ansi\\deff0 {\\rtf1\\ansi\\deff0 {\\fonttbl{\\f0 Arial;}}\\f0 {\\line\n{\\b Level1}: \\par}\n{\\line Body\\par}}}"
			)
		)

		expect(editing.locked).toBe(0)
		type(editing, "Level1", "X")
		expect(await saved(editing)).toContain("{\\fonttbl{\\f0 Arial;}}\\f0 {\\pard")
	})

	it("reads a language as its tag and an unknown identifier as four hex digits", async () => {
		const xml = await bodyXml(latin1(`${H}\\pard {\\lang1031 a}{\\lang1024 b}\\par}`))

		expect(xml).toContain('<w:lang w:val="de-DE"/>')
		expect(xml).toContain('<w:lang w:val="0400"/>')
	})

	it("keeps an underline colour, emphasis marks, web-hidden text and text effects", async () => {
		const editing = await editRtf(
			latin1(
				`${H}{\\colortbl;\\red255\\green0\\blue0;}\\pard {\\ul\\ulc1 Under} {\\accdot Dots} {\\webhidden Web} {\\animtext3 Shine}\\par Two\\par}`
			)
		)

		type(editing, "Under", "X")

		const out = await saved(editing)

		expect(out).toMatch(/\\ul\\ulc1 XUnder/u)
		expect(out).toMatch(/\\accdot Dots/u)
		expect(out).toMatch(/\\webhidden Web/u)
		expect(out).toMatch(/\\animtext3 Shine/u)
	})

	it("keeps East Asian line rules, font alignment and kashida justification", async () => {
		const editing = await editRtf(
			latin1(`${H}\\pard\\nooverflow\\aspalpha\\aspnum\\adjustright\\faauto\\qk20 One para\\par\\pard Two\\par}`)
		)

		type(editing, "One", "X")

		const out = await saved(editing)

		for (const word of ["\\nooverflow", "\\aspalpha", "\\aspnum", "\\adjustright", "\\faauto", "\\qk20"]) expect(out).toContain(word)
	})

	it("locks a paragraph shown under words the translation cannot hold", async () => {
		for (const words of ["\\nocwrap", "\\nowwrap", "\\lisb200", "\\collapsed", "\\fafixed"]) {
			const editing = await editRtf(latin1(`${H}\\pard${words} One para\\par\\pard Two\\par}`))

			expect(editing.locked, words).toBe(1)
		}

		for (const run of ["{\\fittext1000 Fit}", "{\\ulc1 Colour}", "{\\animtext7 Wipe}"]) {
			const editing = await editRtf(latin1(`${H}{\\colortbl;\\red255\\green0\\blue0;}\\pard ${run} para\\par\\pard Two\\par}`))

			expect(editing.locked, run).toBe(1)
		}
	})

	it("writes automatic colour as \\cf0 instead of appending an empty colour-table entry", async () => {
		const editing = await editRtf(latin1(`${H}{\\colortbl;\\red255\\green255\\blue255;}\\pard\\cf0 One para\\par Two\\par}`))

		type(editing, "One", "X")

		const out = await saved(editing)

		expect(out).toContain("{\\colortbl;\\red255\\green255\\blue255;}")
		expect(out).toContain("\\cf0 XOne")
	})
})

describe("RTF writer: locked blocks deleted", () => {
	it("deletes a locked block holding an unknown destination, keeping the state it leaves for what follows", async () => {
		for (const source of [`${H}\\pard One {\\*\\zz x}\\qc para\\par Two\\par}`, `${H}\\pard One\\par\\pard Two {\\*\\zz x}}`]) {
			const editing = await editRtf(latin1(source))
			const s = editing.session
			let at = -1
			let size = 0

			s.state.doc.forEach((node, offset) => {
				if (at < 0 && node.type.name === "locked_block") {
					at = offset
					size = node.nodeSize
				}
			})
			expect(at).toBeGreaterThanOrEqual(0)
			expect(s.dispatch(s.state.tr.delete(at, at + size))).toBe(true)

			const out = await saved(editing)

			expect(out).not.toContain("\\zz")
		}
	})
})
