import type { Node } from "prosemirror-model"
import { describe, expect, it, vi } from "vitest"
import { StoryPart } from "../core/docx/model/part"
import { buildDocx, p } from "./docxSynthetic"
import { layOut } from "./layoutFixtures"
import { editedLayout, pageDigest, savedDigest, type EditedLayout } from "./pagedFixtures"

// The decoded text cache of the content builder (layout/content.ts): used while editing only, keyed by the
// part a text item belongs to (parts are immutable: an edited paragraph is laid out from a new overlay part),
// bounded by characters and emptied when full.

const SMALL = buildDocx({ body: Array.from({ length: 40 }, (_, k) => p(`Paragraph ${String(k)} ${"lorem ipsum ".repeat(20)}`)).join("") })
// More text than the cache holds (2^18 characters): it is emptied while the document is laid out.
const LARGE = buildDocx({ body: Array.from({ length: 300 }, (_, k) => p(`Paragraph ${String(k)} ${"lorem ipsum ".repeat(90)}`)).join("") })

function paragraphPos(doc: Node, index: number): number {
	let found = -1

	doc.forEach((_node, offset, k) => {
		if (k === index) found = offset
	})

	return found
}

// Text decodes during an edit: of the story's own part (laid out as it was) and of other parts (the edit's
// overlay).
async function decodesDuring(edited: EditedLayout, edit: () => Promise<boolean>): Promise<{ source: number; other: number }> {
	const text = vi.spyOn(StoryPart.prototype, "text")
	const source = edited.story.part

	await edit()

	const contexts = text.mock.contexts as StoryPart[]

	text.mockRestore()

	return { source: contexts.filter(part => part === source).length, other: contexts.filter(part => part !== source).length }
}

const insert =
	(edited: EditedLayout, paragraph: number, text: string): (() => Promise<boolean>) =>
	() =>
		edited.edit(edited.session.state.tr.insertText(text, paragraphPos(edited.session.state.doc, paragraph) + 3))

describe("decoded text cache", () => {
	it("decodes the paragraphs an edit lays out again once, and an edited paragraph's new text each time", async () => {
		const edited = await editedLayout(SMALL)

		// The paragraphs around the edit are laid out again from the source part: no decoding.
		const first = await decodesDuring(edited, insert(edited, 5, "zzq"))

		expect(first.source).toBe(0)
		expect(first.other).toBeGreaterThan(0)

		// The same paragraph again: a new overlay part, its new text (never what the cache held for the last one).
		const second = await decodesDuring(edited, insert(edited, 5, "xyw"))

		expect(second.other).toBeGreaterThan(0)

		const shown = pageDigest(edited.layout)

		expect(shown.join("\n")).toContain("xywzzq")
		expect(shown).toEqual(await savedDigest(await edited.save()))
	})

	it("is emptied when full, and lays out as before afterwards", async () => {
		const edited = await editedLayout(LARGE)

		// Laying the whole document out filled the cache past its bound: the first paragraphs were dropped and
		// are decoded again when laid out again.
		const decoded = await decodesDuring(edited, insert(edited, 1, "zzq"))

		expect(decoded.source).toBeGreaterThan(0)
		expect(pageDigest(edited.layout)).toEqual(await savedDigest(await edited.save()))
	})

	it("is not used by the viewer's layout", async () => {
		// The same parts laid out a second time: the viewer decodes again, an editing layout does not.
		const again = async (recordMaps: boolean): Promise<{ first: number; second: number; same: boolean }> => {
			const text = vi.spyOn(StoryPart.prototype, "text")
			const layout = await layOut(SMALL, { recordMaps })
			const first = text.mock.calls.length
			const before = pageDigest(layout)

			text.mockClear()
			layout.beginEdit(0, undefined)

			while (!layout.final) layout.advance(Infinity, Infinity)

			const second = text.mock.calls.length

			text.mockRestore()

			return { first, second, same: pageDigest(layout).join("\n") === before.join("\n") }
		}
		const viewer = await again(false)
		const editing = await again(true)

		expect(viewer.first).toBeGreaterThanOrEqual(40)
		expect(viewer.second).toBeGreaterThanOrEqual(40)
		expect(editing.second).toBe(0)
		expect(viewer.same && editing.same).toBe(true)
	})
})
