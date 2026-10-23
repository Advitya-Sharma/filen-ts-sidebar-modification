import { describe, expect, it } from "vitest"
import type { Node } from "prosemirror-model"
import { DocxEditorSession, proveBlocks, type BlockProof } from "@filen/office/core/docx/edit"
import { openSpec } from "./docxEditFixtures"
import { p } from "./docxSynthetic"

// Gate 1 proved in runs of top-level blocks (a host proving a long body in slices): a run given its first block
// walks only its own blocks, with the verdicts of a walk from the start.

const BLOCKS = 400

// Locks every paragraph whose text ends in 7.
const sevens: BlockProof = {
	check: node => (node.textContent.endsWith("7") ? "unmodelled" : undefined)
}

// The document, counting the top-level blocks a walk reads.
function counted(doc: Node): { doc: Node; reads: () => number } {
	let reads = 0
	const wrapped = new Proxy(doc, {
		get(target, key, receiver) {
			if (key === "child")
				return (index: number) => {
					reads++

					return target.child(index)
				}

			return Reflect.get(target, key, receiver) as unknown
		}
	})

	return { doc: wrapped, reads: () => reads }
}

describe("Gate 1 in runs", () => {
	it("walks only a run's own blocks when given where it starts", async () => {
		const opened = await openSpec({ body: Array.from({ length: BLOCKS }, (_, k) => p(`Block ${String(k)}`)).join("") })
		const session = new DocxEditorSession(opened.context, { proof: null })
		const doc = session.state.doc
		const index = BLOCKS - 32
		let from = 0

		for (let k = 0; k < index; k++) from += doc.child(k).nodeSize

		const to = doc.content.size
		const whole = counted(doc)
		const expected = proveBlocks(session.story, whole.doc, sevens, from, to)
		const run = counted(doc)
		const requests = proveBlocks(session.story, run.doc, sevens, from, to, { index, pos: from })

		expect(requests).toEqual(expected)
		expect(requests.length).toBe(3)
		expect(whole.reads()).toBe(BLOCKS)
		expect(run.reads()).toBe(32)
	})

	it("locks a run's failed blocks from where it starts", async () => {
		const opened = await openSpec({ body: Array.from({ length: 64 }, (_, k) => p(`Block ${String(k)}`)).join("") })
		const session = new DocxEditorSession(opened.context, { proof: null })
		const doc = session.state.doc
		let from = 0

		for (let k = 0; k < 32; k++) from += doc.child(k).nodeSize

		expect(session.prove(sevens, from, doc.content.size, { index: 32, pos: from })).toBe(3)
		expect(session.state.doc.child(37).type.name).toBe("locked_block")
		expect(session.state.doc.child(7).type.name).toBe("paragraph")
	})
})
