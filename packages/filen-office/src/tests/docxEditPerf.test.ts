import { existsSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { AllSelection, TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { DocxDocBuilder, DocxEditorSession, docxSchema } from "@filen/office/core/docx/edit"
import { key, memory, openEditor, percentile as share, placeCaret, spreadPosition, typeBursts } from "../bench/editHarness"
import { CORPUS_DIR } from "./corpusDir"
import { openBytes } from "./docxEditFixtures"

const FIXTURE = join(import.meta.dirname, "fixtures", "08-large-300p.docx")
const TARIFF = join(CORPUS_DIR, "docx", "govuk__Tariff_of_the_United_Kingdom_v1.34.docx")

function percentile(sorted: readonly number[], fraction: number): number {
	return sorted[Math.min(sorted.length - 1, Math.floor(sorted.length * fraction))] ?? 0
}

function heapMb(): number {
	globalThis.gc?.()

	return process.memoryUsage().heapUsed / 1048576
}

async function measureBuild(
	path: string
): Promise<{ session: DocxEditorSession; buildMs: number; sessionMs: number; heap: number; blocks: number }> {
	const opened = await openBytes(new Uint8Array(readFileSync(path)))
	const body = opened.model.body

	if (body === undefined) throw new Error("no body")

	const heapBefore = heapMb()
	const start = performance.now()
	const story = new DocxDocBuilder(opened.context).build(body)
	const buildMs = performance.now() - start
	const heap = heapMb() - heapBefore
	const sessionStart = performance.now()
	const session = new DocxEditorSession(opened.context)
	const sessionMs = performance.now() - sessionStart

	return { session, buildMs, sessionMs, heap, blocks: story.doc.childCount }
}

// Keystrokes at positions spread over the document: insert a character as a view would (a text
// selection, insertText through dispatch), the state's invariant plugins, history and snapshot included.
function keystrokes(session: DocxEditorSession, count: number, key: "type" | "backspace" | "enter" = "type"): number[] {
	const times: number[] = []

	for (let k = 0; k < count; k++) {
		const state = session.state
		// A position inside some text, spread over the document.
		let target = Math.floor((((k * 7919) % 10007) / 10007) * state.doc.content.size)
		let $pos = state.doc.resolve(target)

		while (!$pos.parent.inlineContent && target < state.doc.content.size) $pos = state.doc.resolve(++target)

		const start = performance.now()
		const tr = state.tr.setSelection(TextSelection.create(state.doc, target))

		if (key === "type") tr.insertText("a")
		else if (key === "enter") tr.split(target)
		else if ($pos.parentOffset > 0) tr.delete(target - 1, target)
		else tr.insertText("a")

		const applied = session.dispatch(tr)

		session.snapshot()
		times.push(performance.now() - start)

		if (!applied) throw new Error(`keystroke at ${target} refused`)
	}

	return times.sort((a, b) => a - b)
}

describe("docx editing performance", () => {
	it("builds the 300-page fixture and types into it under a frame per keystroke", { timeout: 120_000 }, async () => {
		const { session, buildMs, sessionMs, heap, blocks } = await measureBuild(FIXTURE)

		keystrokes(session, 50)

		const times = keystrokes(session, 300)
		const p50 = percentile(times, 0.5)
		const p95 = percentile(times, 0.95)
		const backspace = keystrokes(session, 200, "backspace")
		const enter = keystrokes(session, 200, "enter")

		session.dispatch(session.state.tr.setSelection(new AllSelection(session.state.doc)))

		const boldStart = performance.now()

		session.execute({ type: "format.bold" })
		session.snapshot()

		const boldMs = performance.now() - boldStart

		console.log(
			`300p: ${blocks} blocks, build ${buildMs.toFixed(1)} ms, session (build + proof) ${sessionMs.toFixed(1)} ms, heap ${heap.toFixed(1)} MB; ` +
				`keystroke p50 ${p50.toFixed(2)} ms p95 ${p95.toFixed(2)} ms max ${(times[times.length - 1] ?? 0).toFixed(2)} ms; ` +
				`backspace p95 ${percentile(backspace, 0.95).toFixed(2)} ms, enter p95 ${percentile(enter, 0.95).toFixed(2)} ms; select-all bold ${boldMs.toFixed(1)} ms`
		)
		expect(session.state.doc.type).toBe(docxSchema.nodes.doc)
		// The target is 16 ms; the bound leaves room for loaded CI machines.
		expect(p95).toBeLessThan(48)
		expect(percentile(enter, 0.95)).toBeLessThan(48)
		expect(percentile(backspace, 0.95)).toBeLessThan(48)
	})

	// As the web edits: the in-thread editing host (sliced open with the writer's proof), the pages in view laid
	// out in the key's task, the toolbar's snapshot read after each key.
	it(
		"opens the 300-page fixture for editing and types into it in page view under a frame per keystroke",
		{ timeout: 120_000 },
		async () => {
			const before = memory()
			const opened = await openEditor(new Uint8Array(readFileSync(FIXTURE)))
			const heap = memory().heapMb - before.heapMb
			const editor = opened.editor

			placeCaret(editor, spreadPosition(editor.session.state, 5000))

			const first = key(editor, "type")?.ms ?? 0
			const typed = await typeBursts(editor, 10, 20, ["type"], 0, 0)
			const enter = await typeBursts(editor, 5, 10, ["enter"], 50, 0)
			const p95 = share(typed.times, 0.95)

			console.log(
				`300p page view: open ${opened.editorMs.toFixed(0)} ms, heap ${heap.toFixed(1)} MB, longest open task ${opened.tasks.maxMs.toFixed(0)} ms; ` +
					`first key ${first.toFixed(2)} ms; keystroke p50 ${share(typed.times, 0.5).toFixed(2)} ms p95 ${p95.toFixed(2)} ms; enter p95 ${share(enter.times, 0.95).toFixed(2)} ms`
			)
			opened.close()
			expect(typed.refused).toBe(0)
			// The target is 16 ms (src/bench/editBench.ts measures the gate); the bound leaves room for loaded CI machines.
			expect(p95).toBeLessThan(48)
		}
	)

	it.skipIf(!existsSync(TARIFF))("builds the Tariff document", { timeout: 300_000 }, async () => {
		const { session, buildMs, sessionMs, heap, blocks } = await measureBuild(TARIFF)
		const times = keystrokes(session, 100)
		const all = (): void => {
			session.dispatch(session.state.tr.setSelection(new AllSelection(session.state.doc)))
		}
		let start = performance.now()

		all()
		session.snapshot()

		const cold = performance.now() - start

		session.dispatch(session.state.tr.insertText("z", session.state.doc.content.size - 3))
		start = performance.now()
		session.snapshot()

		const warm = performance.now() - start
		// Shift+arrow over half the document, a snapshot per step.
		const half = Math.floor(session.state.doc.content.size / 2)
		const arrows: number[] = []

		for (let k = 0; k < 20; k++) {
			const state = session.state
			const step = performance.now()

			session.dispatch(state.tr.setSelection(TextSelection.between(state.doc.resolve(1), state.doc.resolve(half + k))))
			session.snapshot()
			arrows.push(performance.now() - step)
		}

		arrows.sort((a, b) => a - b)

		console.log(
			`Tariff: ${blocks} blocks, build ${buildMs.toFixed(0)} ms, session ${sessionMs.toFixed(0)} ms, heap ${heap.toFixed(1)} MB; ` +
				`keystroke p50 ${percentile(times, 0.5).toFixed(2)} ms p95 ${percentile(times, 0.95).toFixed(2)} ms; ` +
				`select-all snapshot cold ${cold.toFixed(0)} ms, after an edit ${warm.toFixed(1)} ms; shift+arrow over half p50 ${percentile(arrows, 0.5).toFixed(2)} ms`
		)
		expect(blocks).toBeGreaterThan(0)
		expect(warm).toBeLessThan(150)
		expect(percentile(arrows, 0.5)).toBeLessThan(15)
	})
})
