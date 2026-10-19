// @vitest-environment happy-dom
import { EditorState, type Transaction } from "prosemirror-state"
import { Fragment, Slice, type Node } from "prosemirror-model"
import { describe, expect, it } from "vitest"
import { docxSchema } from "../core/docx/edit/index"
import { LiveStarts, changedRange } from "../react/docx/edit/window"

// The body's block starts kept up to date from each transaction's steps (Mobile view's window and heights):
// they must equal a fresh count after any edit, found in one pass over the steps.

const nodes = docxSchema.nodes
const para = (text: string): Node => nodes.paragraph.create(null, text === "" ? null : docxSchema.text(text))
const docOf = (count: number, text: (k: number) => string): Node =>
	nodes.doc.create(
		null,
		Array.from({ length: count }, (_, k) => para(text(k)))
	)

function fresh(doc: Node): number[] {
	const out: number[] = []
	let pos = 0

	for (let k = 0; k < doc.childCount; k++) {
		out.push(pos)
		pos += doc.child(k).nodeSize
	}

	return out
}

function random(seed: number): () => number {
	let state = seed >>> 0

	return () => {
		state = (state * 1664525 + 1013904223) >>> 0

		return state / 2 ** 32
	}
}

// One to four steps of typing, deleting, inserting and replacing whole blocks, splitting and bold.
function randomTransaction(state: EditorState, next: () => number): Transaction {
	const tr = state.tr

	for (let step = 1 + Math.floor(next() * 4); step > 0; step--) {
		const doc = tr.doc
		const size = doc.content.size
		const kind = Math.floor(next() * 6)
		const a = Math.floor(next() * (size + 1))
		const b = Math.min(size, a + Math.floor(next() * 30))
		const inside = Math.max(1, Math.min(a, size - 1))

		try {
			if (kind === 0) tr.insertText("x".repeat(1 + Math.floor(next() * 3)), inside)
			else if (kind === 1) tr.delete(a, b)
			else if (kind === 2) {
				const $a = doc.resolve(a)

				tr.insert($a.depth === 0 ? a : $a.before(1), Fragment.from([para(`n${String(step)}`), para("")]))
			} else if (kind === 3) {
				if (doc.resolve(inside).parent.isTextblock) tr.split(inside)
			} else if (kind === 4) {
				const $a = doc.resolve(a)
				const $b = doc.resolve(b)

				tr.replace($a.depth === 0 ? a : $a.before(1), $b.depth === 0 ? b : $b.after(1), new Slice(Fragment.from(para("r")), 0, 0))
			} else tr.addMark(a, b, docxSchema.marks.bold.create())
		} catch {
			// A step the document does not take: the transaction goes on without it.
		}
	}

	return tr
}

describe("LiveStarts", () => {
	it("equals a fresh count after random multi-step transactions", () => {
		for (let seed = 1; seed <= 300; seed++) {
			const next = random(seed)
			let state = EditorState.create({ schema: docxSchema, doc: docOf(20, k => `para ${String(k)} text`) })
			const live = new LiveStarts(state.doc)

			for (let round = 0; round < 40; round++) {
				const tr = randomTransaction(state, next)

				if (tr.doc.childCount === 0) continue

				state = state.apply(tr)
				live.applyAll([tr])
				expect(Array.from(live.starts), `seed ${String(seed)} round ${String(round)}`).toEqual(fresh(state.doc))
				expect(live.document).toBe(state.doc)
			}
		}
	})

	it("takes a later step's insertion before an earlier step's range into account", () => {
		const doc = docOf(6, k => `block ${String(k)}`)
		const starts = fresh(doc)
		const at = (k: number): number => starts[k] ?? 0
		// A block replaced near the end, then text typed into the first block: the first step's range moves.
		const tr = EditorState.create({ schema: docxSchema, doc })
			.tr.replaceWith(at(4), at(5), Fragment.from([para("four"), para("and more")]))
			.insertText("typed ", at(0) + 1)
		const live = new LiveStarts(doc)

		expect(changedRange(tr)).toEqual({ from: at(0) + 1, to: at(5) })
		live.applyAll([tr])
		expect(Array.from(live.starts)).toEqual(fresh(tr.doc))
	})

	it("finds what thousands of steps changed in one pass (a Replace All, a bold over everything)", () => {
		const count = 8000
		const doc = docOf(count, k => `the cat number ${String(k)} sat on the mat`)
		const starts = fresh(doc)
		const replace = EditorState.create({ schema: docxSchema, doc }).tr

		for (let k = count - 1; k >= 0; k--) replace.insertText("THE", (starts[k] ?? 0) + 1, (starts[k] ?? 0) + 4)

		const bold = EditorState.create({ schema: docxSchema, doc }).tr

		for (let k = 0; k < count; k++) bold.addMark((starts[k] ?? 0) + 1, (starts[k] ?? 0) + 4, docxSchema.marks.bold.create())

		const live = new LiveStarts(doc)
		const started = performance.now()

		live.applyAll([replace])

		const replaced = performance.now() - started
		const boldLive = new LiveStarts(doc)
		const boldStarted = performance.now()

		boldLive.applyAll([bold])

		const bolded = performance.now() - boldStarted

		expect(Array.from(live.starts)).toEqual(fresh(replace.doc))
		expect(Array.from(boldLive.starts)).toEqual(starts)
		// A mapping inverted per step took seconds here; one pass takes a few milliseconds.
		expect(replaced).toBeLessThan(250)
		expect(bolded).toBeLessThan(250)
	})
})
