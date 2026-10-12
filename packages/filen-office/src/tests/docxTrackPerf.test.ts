import { existsSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { AllSelection, TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { DocxEditorSession } from "@filen/office/core/docx/edit"
import { storyReviewIndex } from "../core/docx/review/accept"
import { CORPUS_DIR } from "./corpusDir"
import { openBytes } from "./docxEditFixtures"

// Recording tracked changes adds no document scan to a keystroke: typing, Backspace and Enter with tracking on
// cost what they cost with it off, on the 300-page fixture and on Tariff; select-all edits stay bounded.

const FIXTURE = join(import.meta.dirname, "fixtures", "08-large-300p.docx")
const TARIFF = join(CORPUS_DIR, "docx", "govuk__Tariff_of_the_United_Kingdom_v1.34.docx")

function percentile(sorted: readonly number[], fraction: number): number {
	return sorted[Math.min(sorted.length - 1, Math.floor(sorted.length * fraction))] ?? 0
}

async function sessionOf(path: string, tracking: boolean): Promise<DocxEditorSession> {
	const opened = await openBytes(new Uint8Array(readFileSync(path)))

	opened.context.identity.author = "Perf"
	opened.context.review.tracking = tracking

	const session = new DocxEditorSession(opened.context)

	// The review snapshot keeps the session's index (made once, as the first snapshot does).
	session.snapshot()

	return session
}

function keystrokes(session: DocxEditorSession, count: number, key: "type" | "backspace" | "enter"): number[] {
	const times: number[] = []

	for (let k = 0; k < count; k++) {
		const state = session.state
		let target = Math.floor((((k * 7919) % 10007) / 10007) * state.doc.content.size)
		let $pos = state.doc.resolve(target)

		while (!$pos.parent.inlineContent && target < state.doc.content.size) $pos = state.doc.resolve(++target)

		const start = performance.now()
		const tr = state.tr.setSelection(TextSelection.create(state.doc, target))

		if (key === "type") tr.insertText("a")
		else if (key === "enter") tr.split(target)
		else if ($pos.parentOffset > 0) tr.delete(target - 1, target)
		else tr.insertText("a")

		session.dispatch(tr)
		session.snapshot()
		times.push(performance.now() - start)
	}

	return times.sort((a, b) => a - b)
}

async function measure(path: string, label: string): Promise<{ off: number; on: number }> {
	const result = { off: 0, on: 0 }

	for (const tracking of [false, true]) {
		const session = await sessionOf(path, tracking)

		keystrokes(session, 30, "type")

		const type = percentile(keystrokes(session, 200, "type"), 0.95)
		const backspace = percentile(keystrokes(session, 100, "backspace"), 0.95)
		const enter = percentile(keystrokes(session, 100, "enter"), 0.95)

		console.log(
			`${label} tracking ${tracking ? "on " : "off"}: keystroke p95 ${type.toFixed(2)} ms, backspace p95 ${backspace.toFixed(2)} ms, enter p95 ${enter.toFixed(2)} ms` +
				(tracking ? `, changes ${storyReviewIndex(session).revisionCount}` : "")
		)
		result[tracking ? "on" : "off"] = Math.max(type, backspace, enter)
	}

	return result
}

describe("docx tracked changes performance", () => {
	it("types into the 300-page fixture with tracking on as fast as with it off", { timeout: 120_000 }, async () => {
		const { off, on } = await measure(FIXTURE, "300p")

		// The target is 16 ms; the bounds leave room for loaded machines.
		expect(on).toBeLessThan(48)
		expect(on).toBeLessThan(Math.max(off * 3, off + 8))
	})

	it.skipIf(!existsSync(TARIFF))("types into Tariff with tracking on, and deletes all of it", { timeout: 600_000 }, async () => {
		const { off, on } = await measure(TARIFF, "Tariff")

		expect(on).toBeLessThan(Math.max(off * 3, off + 8))

		const session = await sessionOf(TARIFF, true)
		const doc = session.state.doc
		const size = doc.content.size

		// Ctrl+A, Delete: everything stays, marked (section breaks included).
		session.dispatch(session.state.tr.setSelection(new AllSelection(doc)))

		let start = performance.now()
		const applied = session.dispatch(session.state.tr.deleteSelection())

		session.snapshot()

		const deleteMs = performance.now() - start
		const changes = storyReviewIndex(session).revisionCount

		start = performance.now()
		session.execute({ type: "history.undo" })

		const undoMs = performance.now() - start

		console.log(
			`Tariff tracked select-all delete of ${size} positions: ${String(applied)} in ${deleteMs.toFixed(0)} ms, ${changes} changes; undo ${undoMs.toFixed(0)} ms`
		)
		expect(applied).toBe(true)
		expect(session.state.doc.eq(doc)).toBe(true)
		expect(deleteMs).toBeLessThan(20_000)
	})
})
