import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { bytesSource } from "@filen/office/core/opc"
import type { FontExtents, TextMeasurer } from "@filen/office/core/text"
import { DocxEditorSession, docxSchema, type BlockProof } from "@filen/office/core/docx/edit"
import { saveDocx, writerProof } from "@filen/office/core/docx/write"
import type { DocxEngineServices } from "@filen/office/dom/docx"
import { inThreadEditingHost } from "@filen/office/dom/docx/editing"
import { nodeText, openBytes } from "./docxEditFixtures"
import { buildDocx, p } from "./docxSynthetic"

const measurer: TextMeasurer = {
	kerningControl: true,
	measure(font, _kerning, texts, count, widths) {
		const size = Number(/([\d.]+)px/.exec(font)?.[1] ?? "16")

		for (let i = 0; i < count; i++) widths[i] = (texts[i] ?? "").length * size * 0.5
	},
	fontExtents(font): FontExtents {
		const size = Number(/([\d.]+)px/.exec(font)?.[1] ?? "16")

		return { ascent: size * 0.8, descent: size * 0.2 }
	}
}

function services(pauses: { count: number }): DocxEngineServices {
	return {
		fontHost: { addBytes: () => Promise.resolve(true), addLocal: () => Promise.resolve(false), remove: () => undefined },
		replayFonts: false,
		fontAssets: { read: () => Promise.resolve(new Uint8Array(4)) },
		measurer,
		yield: () => {
			pauses.count++

			return Promise.resolve()
		},
		// Every slice pauses: the proof runs in its smallest steps.
		sliceMs: 0
	}
}

// Locks every paragraph whose text says so.
const LOCK_MARKED: BlockProof = {
	check: node => (node.type === docxSchema.nodes.paragraph && node.textContent.includes("LOCK") ? "proof" : undefined)
}

const BODY = Array.from({ length: 90 }, (_, k) => p(k % 30 === 7 ? `Paragraph ${k} LOCK` : `Paragraph ${k}`)).join("")

describe("inThreadEditingHost", () => {
	it("proves the body in slices, locks only the failing blocks and opens clean", async () => {
		const pauses = { count: 0 }
		const editing = inThreadEditingHost(() => services(pauses), { proof: () => LOCK_MARKED })
		const session = editing.host(bytesSource(buildDocx({ body: BODY })), () => undefined)

		expect(editing.saveInput()).toBeUndefined()

		const editor = await editing.editor
		let locked = 0

		editor.state.doc.forEach(node => {
			if (node.type === docxSchema.nodes.locked_block) locked++
		})

		expect(editing.proofLocked()).toBe(3)
		expect(locked).toBe(3)
		expect(editor.session.snapshot().dirty).toBe(false)
		expect(pauses.count).toBeGreaterThan(3)
		session.close()
		expect(editing.saveInput()).toBeUndefined()
	})

	it("saves every story's edits through the writer with the real proof", async () => {
		const editing = inThreadEditingHost(() => services({ count: 0 }), { proof: writerProof })
		const session = editing.host(bytesSource(buildDocx({ body: p("Hello world") + p("Second") })), () => undefined)
		const editor = await editing.editor
		const input = editing.saveInput()

		if (input === undefined) throw new Error("no save input")

		expect(editing.proofLocked()).toBe(0)

		const state = editor.state

		editor.dispatch(state.tr.setSelection(TextSelection.create(state.doc, 6)).insertText(","))
		expect(editor.session.snapshot().dirty).toBe(true)
		expect(input.sessions).toContain(editor.session)

		const written = input.sessions.map(story => story.state.doc)
		const result = await saveDocx(input)

		if (result.bytes === undefined) throw new Error("no bytes")

		input.sessions.forEach((story, index) => {
			const doc = written[index]

			if (doc !== undefined) story.markSaved(doc)
		})
		expect(editor.session.snapshot().dirty).toBe(false)

		const reopened = await openBytes(result.bytes)
		const body = reopened.model.body

		expect(body).toBeDefined()
		expect(result.changedParts).toContain("word/document.xml")

		const check = new DocxEditorSession(reopened.context)

		expect(nodeText(check.state.doc.child(0))).toBe("Hello, world")
		session.close()
	})
})
