import { describe, expect, it } from "vitest"
import { docxSchema, type BlockProof } from "@filen/office/core/docx/edit"
import { bytesSource } from "@filen/office/core/opc"
import type { FontExtents, TextMeasurer } from "@filen/office/core/text"
import type { DocxEngineServices } from "@filen/office/dom/docx"
import { inThreadEditingHost } from "@filen/office/dom/docx/editing"
import { buildDocx, p } from "./docxSynthetic"

// A Gate 1 partly made elsewhere (an RTF document's writer proves in its worker): the editor's open checks the
// body with the first part meanwhile and starts editing only once the rest has arrived and locked its blocks.

const measurer: TextMeasurer = {
	kerningControl: true,
	measure(_font, _kerning, texts, count, widths) {
		for (let i = 0; i < count; i++) widths[i] = (texts[i] ?? "").length * 8
	},
	fontExtents(): FontExtents {
		return { ascent: 13, descent: 3 }
	}
}

const services = (): DocxEngineServices => ({
	fontHost: { addBytes: () => Promise.resolve(true), addLocal: () => Promise.resolve(false), remove: () => undefined },
	replayFonts: false,
	fontAssets: { read: () => Promise.resolve(new Uint8Array(4)) },
	measurer,
	yield: () => Promise.resolve(),
	sliceMs: 0
})

const containing = (text: string): BlockProof => ({
	check: node => (node.type === docxSchema.nodes.paragraph && node.textContent.includes(text) ? "proof" : undefined)
})

describe("Gate 1 that arrives later", () => {
	it("opens the editor once the rest of the proof arrived, locked by both parts", async () => {
		let arrive: () => void = () => undefined
		let firstChecks = 0
		const first: BlockProof = {
			check(node, story) {
				firstChecks++

				return containing("FIRST").check(node, story)
			}
		}
		const later = new Promise<{ extra: BlockProof; full: BlockProof }>(resolve => {
			arrive = () => resolve({ extra: containing("LATER"), full: containing("LATER") })
		})
		const editing = inThreadEditingHost(services, { proof: () => ({ first, later }) })
		const handle = editing.host(
			bytesSource(buildDocx({ body: p("One") + p("FIRST locks") + p("Two") + p("LATER locks") })),
			() => undefined
		)
		let opened = false

		void editing.editor.then(() => {
			opened = true
		})

		for (let k = 0; k < 200 && firstChecks < 4; k++) await new Promise(resolve => setTimeout(resolve, 5))

		expect(firstChecks).toBe(4)
		await new Promise(resolve => setTimeout(resolve, 20))
		expect(opened).toBe(false)

		arrive()

		const editor = await editing.editor
		const locked: string[] = []

		editor.state.doc.descendants(node => {
			if (node.type === docxSchema.nodes.locked_block) locked.push(String(node.attrs["text"]))

			return node.type !== docxSchema.nodes.locked_block
		})

		expect(locked).toEqual(["FIRST locks", "LATER locks"])
		expect(editing.proofLocked()).toBe(2)
		expect(editor.session.snapshot().dirty).toBe(false)
		handle.close()
	})
})
