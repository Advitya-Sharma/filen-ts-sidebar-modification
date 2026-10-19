import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { saveDocx } from "@filen/office/core/docx/write"
import { parseFieldCommand } from "../../core/docx/edit/commands/fields"
import { parseFormCommand } from "../../core/docx/edit/commands/forms"
import { parseInsertCommand } from "../../core/docx/edit/commands/insert"
import { parseNoteCommand } from "../../core/docx/edit/commands/notes"
import { buildDocx, p } from "../docxSynthetic"
import { pagedEditor, settleAll } from "../pagedFixtures"

// Text a command writes into the file is refused at the command when XML cannot carry it (C0 controls,
// U+FFFE, U+FFFF, lone surrogates): accepting it only turns every later save into a refusal.

const ILLEGAL = [String.fromCharCode(0xfffe), String.fromCharCode(0xffff), String.fromCharCode(0xd800), `a${String.fromCharCode(0xdc00)}b`]
const EMOJI = String.fromCodePoint(0x1f600)

describe("commands refuse text XML cannot carry", () => {
	it.each(ILLEGAL)("refuses %j in bridge messages for link targets, field codes, control text and note marks", illegal => {
		expect(parseInsertCommand("link.set", { target: `https://e.example/${illegal}` })).toBeUndefined()
		expect(parseFieldCommand("insert.field", { instruction: `DATE \\@ "d${illegal}"` })).toBeUndefined()
		expect(parseFormCommand("control.text", { text: `x${illegal}` })).toBeUndefined()
		expect(parseFormCommand("control.items", { items: [{ text: `x${illegal}`, value: "1" }] })).toBeUndefined()
		expect(parseNoteCommand("note.insertFootnote", { mark: `*${illegal}` })).toBeUndefined()
	})

	it("still accepts text with paired surrogates", () => {
		expect(parseInsertCommand("link.set", { target: `https://e.example/${EMOJI}` })).toBeDefined()
		expect(parseFieldCommand("insert.field", { instruction: `DATE \\@ "d${EMOJI}"` })).toBeDefined()
	})

	it("refuses such text on the host's typed path, story and document commands alike, and the file still saves", async () => {
		const opened = await pagedEditor(buildDocx({ body: p("Alpha beta gamma") }))
		const { editor } = opened

		opened.context.identity.author = "A"
		await settleAll(editor)
		editor.session.dispatch(editor.session.state.tr.setSelection(TextSelection.create(editor.session.state.doc, 1, 6)))

		for (const illegal of [...ILLEGAL, String.fromCharCode(1)]) {
			expect(await editor.execute({ type: "link.set", target: `https://e.example/${illegal}` })).toBe(false)
			expect(await editor.execute({ type: "insert.field", instruction: `DATE \\@ "d${illegal}"` })).toBe(false)
			expect(await editor.execute({ type: "note.insertFootnote", mark: `*${illegal}` })).toBe(false)
			expect(await editor.execute({ type: "insert.date", format: `d${illegal}` })).toBe(false)
		}

		// Typed text is cleaned instead, as before.
		expect(await editor.execute({ type: "comment.add", text: `c${String.fromCharCode(1)}` })).toBe(true)

		expect(await editor.execute({ type: "link.set", target: `https://e.example/${EMOJI}` })).toBe(true)
		await expect(saveDocx({ pkg: opened.pkg, context: opened.context, sessions: editor.sessions })).resolves.toBeDefined()
	})
})
