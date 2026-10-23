import { describe, expect, it } from "vitest"
import { bytesSource } from "@filen/office/core/opc"
import type { FontExtents, TextMeasurer } from "@filen/office/core/text"
import { prepareDocxSave, writerProof } from "@filen/office/core/docx/write"
import type { DocxEngineServices } from "@filen/office/dom/docx"
import { inThreadEditingHost } from "@filen/office/dom/docx/editing"
import { buildDocx, p } from "./docxSynthetic"

// A write of the editor's document (Download as from the editor) given a signal: an abort before or while it runs ends
// it at its next slice or stage and hands nothing out; the editor and later saves are unaffected.

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

function services(onPause: () => void): DocxEngineServices {
	return {
		fontHost: { addBytes: () => Promise.resolve(true), addLocal: () => Promise.resolve(false), remove: () => undefined },
		replayFonts: false,
		fontAssets: { read: () => Promise.resolve(new Uint8Array(4)) },
		measurer,
		yield: () => {
			onPause()

			return Promise.resolve()
		},
		sliceMs: 0
	}
}

const BODY = Array.from({ length: 60 }, (_, k) => p(`Paragraph ${String(k)}`)).join("")

describe("a write of the editor's document with a signal", () => {
	it("rejects as aborted, before it starts or at its next slice, and a later save still writes", async () => {
		const state = { pauses: 0, abortAt: Infinity, controller: new AbortController() }
		const editing = inThreadEditingHost(
			() =>
				services(() => {
					if (++state.pauses >= state.abortAt) state.controller.abort()
				}),
			{ proof: writerProof, writer: { prepare: prepareDocxSave } }
		)
		const session = editing.host(bytesSource(buildDocx({ body: BODY })), () => undefined)
		const editor = await editing.editor

		editor.session.dispatch(editor.session.state.tr.insertText("Edited ", 1))

		const before = new AbortController()

		before.abort()
		await expect(editing.save(undefined, { signal: before.signal })).rejects.toMatchObject({ name: "AbortError" })

		state.abortAt = state.pauses + 2
		await expect(editing.save(undefined, { signal: state.controller.signal })).rejects.toMatchObject({ name: "AbortError" })

		const written = await editing.save()

		expect(written.bytes.length).toBeGreaterThan(0)
		expect(editor.session.snapshot().dirty).toBe(true)
		session.close()
	})
})
