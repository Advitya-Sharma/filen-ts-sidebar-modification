import { getHeapSpaceStatistics } from "node:v8"
import { EditorState, TextSelection } from "prosemirror-state"
import type { Node } from "prosemirror-model"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import { replaceAll } from "@filen/office/core/docx/edit"
import {
	addComments,
	commentIds,
	cpuMs,
	key,
	memory,
	nextTurn,
	openEditor,
	percentile,
	placeCaret,
	save,
	selectAll,
	spreadPosition,
	type KeyKind,
	type OpenedEditor
} from "./editHarness"

// A long editing session on one document (the editor's soak test): keys spread over the document (typing,
// Backspace, Enter, paste), tracking toggled every 500 keys, undo/redo bursts, select-all formatting undone,
// Replace All, comments added, replied to and resolved, table rows and columns added and removed, the markup
// view switched, and a save every 250 operations (as idle autosave would). After each save: memory after a
// full collection, the save's time and Gate 2's, the keys' p95 since the last save.
//
// Memory counts the session's data: the growth of V8's compiled code (code and trusted spaces: which code paths
// ran, not what the document holds) is taken out, and the harness keeps nothing of its own across saves (the
// last save's bytes and text are read after it). After the last save the undo history is cleared and memory
// measured again: what the history held is reported beside the rest.

export interface SoakPlan {
	readonly keys: number
	readonly undoBursts: number
	readonly formatAll: number
	readonly replaceAll: number
	readonly comments: number
	readonly tableOps: number
	readonly viewSwitches: number
	readonly trackEvery: number
	readonly saveEvery: number
	// Keys at one place before the caret moves elsewhere.
	readonly perSpot: number
}

export const SOAK_PLAN: SoakPlan = {
	keys: 5000,
	undoBursts: 200,
	formatAll: 20,
	replaceAll: 10,
	comments: 50,
	tableOps: 30,
	viewSwitches: 10,
	trackEvery: 500,
	saveEvery: 250,
	perSpot: 25
}

export interface SoakSave {
	readonly operations: number
	// Heap after a full collection over the heap before the open, less the growth of compiled code.
	readonly heapMb: number
	// The same with compiled code (the heap as process.memoryUsage reports it).
	readonly rawHeapMb: number
	// Growth of V8's code and trusted spaces since before the open.
	readonly codeMb: number
	// About what the documents' content takes (each session's document, and what the writer keeps of the story as
	// built that the document no longer shares): the part of the heap that grows with what was typed.
	readonly contentMb: number
	readonly externalMb: number
	readonly arrayBuffersMb: number
	readonly saveMs: number
	// The save's CPU time, and what was edited since the open (top-level blocks that are not the opened ones, MB
	// as contentMb counts): what a save costs is to depend on that alone.
	readonly saveCpuMs: number
	readonly editedMb: number
	readonly gate2Ms: number
	readonly longestTaskMs: number
	readonly longestTaskCpuMs: number
	// p95 of the keys since the previous save.
	readonly keysP95: number
}

export interface SoakResult {
	readonly openHeapMb: number
	// contentMb as opened.
	readonly openContentMb: number
	// After the last save, heap (as heapMb) with every session's undo history cleared, what the history held
	// (the last save's heapMb less that), and what it may hold: about one copy of the documents' content (a
	// format over a whole document keeps the content as it was).
	readonly clearedHeapMb: number
	readonly historyMb: number
	readonly historyLimitMb: number
	readonly saves: readonly SoakSave[]
	// Every key's time, in order.
	readonly keys: readonly number[]
	readonly refused: number
	readonly failedSaves: readonly string[]
	readonly passes: number
	readonly bytes: Uint8Array | undefined
	// The body's text as the session holds it at the last save.
	readonly text: string
}

type Special = "undo" | "format" | "replace" | "comment" | "table" | "view"

// The specials spread evenly among the keys.
function schedule(plan: SoakPlan): (Special | undefined)[] {
	const specials: Special[] = []
	const add = (kind: Special, count: number): void => {
		for (let k = 0; k < count; k++) specials.push(kind)
	}

	add("undo", plan.undoBursts)
	add("format", plan.formatAll)
	add("replace", plan.replaceAll)
	add("comment", plan.comments)
	add("table", plan.tableOps)
	add("view", plan.viewSwitches)

	// A fixed interleaving: specials of each kind spread out, not bunched.
	const order = specials.map((kind, k) => ({ kind, at: (k * 7919) % specials.length })).sort((a, b) => a.at - b.at)
	const out: (Special | undefined)[] = []
	const step = plan.keys / Math.max(1, order.length)
	let next = 0

	for (let k = 0; k < plan.keys; k++) {
		out.push(undefined)

		while (next < order.length && next * step <= k) out.push(order[next++]?.kind)
	}

	while (next < order.length) out.push(order[next++]?.kind)

	return out
}

const KINDS: readonly KeyKind[] = [
	"type",
	"type",
	"type",
	"type",
	"type",
	"type",
	"type",
	"backspace",
	"backspace",
	"enter",
	"type",
	"type",
	"type",
	"type",
	"type",
	"type",
	"backspace",
	"enter",
	"type",
	"paste"
]

// A word Replace All changes in a few hundred places: the one nearest 300 occurrences.
function replaceWord(text: string): string {
	const counts = new Map<string, number>()

	for (const word of text.split(/[^A-Za-z]+/)) if (word.length >= 5) counts.set(word, (counts.get(word) ?? 0) + 1)

	let best = "the"
	let distance = Number.POSITIVE_INFINITY

	for (const [word, count] of counts) {
		if (Math.abs(count - 300) < distance) {
			best = word
			distance = Math.abs(count - 300)
		}
	}

	return best
}

// The caret into a table cell after place `k` (false when the body has no table).
function intoTable(editor: PagedEditor, k: number): boolean {
	const doc = editor.session.state.doc
	const from = spreadPosition(editor.session.state, k)
	let found = -1
	let first = -1

	doc.forEach((child, offset) => {
		if (child.type.name !== "table") return
		if (first < 0) first = offset
		if (found < 0 && offset >= from) found = offset
	})

	const table = found >= 0 ? found : first

	if (table < 0) return false

	editor.dispatch(editor.session.state.tr.setSelection(TextSelection.near(doc.resolve(table + 1))))

	return true
}

export interface SoakOptions {
	// Run the plan again (other places) until this much time went by, and this many passes; one at least.
	readonly minutes?: number
	readonly passes?: number
	readonly plan?: SoakPlan
	readonly onSave?: (save: SoakSave) => void
}

export async function soak(bytes: Uint8Array, options: SoakOptions = {}): Promise<SoakResult> {
	const plan = options.plan ?? SOAK_PLAN
	const before = memory()
	const codeBefore = codeMb()
	// Zipped in a worker as the web zips: the main thread does the save's prepare alone.
	const opened: OpenedEditor = await openEditor(bytes, { zipWorker: true })
	const openHeapMb = memory().heapMb - before.heapMb - (codeMb() - codeBefore)
	const editor = opened.editor
	const openContentMb = documentsMb(opened.editing.saveInput()?.sessions ?? [])
	const word = replaceWord(editor.session.state.doc.textContent)
	const saves: SoakSave[] = []
	const keys: number[] = []
	const failedSaves: string[] = []
	const deadline = performance.now() + (options.minutes ?? 0) * 60_000
	let refused = 0
	let operations = 0
	let sinceSave = 0
	let keyCount = 0
	let replaced = false
	let markup = false
	let tableInserted = false
	let columnInserted = false
	let lastComment: string | undefined
	let commentStep = 0
	let bytesOut: Uint8Array | undefined
	let text = ""
	let clearedHeapMb = 0
	let historyLimitMb = 0
	let passes = 0

	// The last save's bytes and text are read after its memory is: the harness holds nothing across saves.
	const doSave = async (last = false): Promise<void> => {
		await editor.whenSettled()

		try {
			const cpu = cpuMs()
			const saved = await save(opened.editing)
			const saveCpuMs = cpuMs() - cpu
			const sessions = opened.editing.saveInput()?.sessions ?? []
			const content = documentsMb(sessions)
			const after = memory()
			const code = codeMb() - codeBefore
			const entry: SoakSave = {
				operations,
				heapMb: after.heapMb - before.heapMb - code,
				rawHeapMb: after.heapMb - before.heapMb,
				codeMb: code,
				contentMb: content,
				externalMb: after.externalMb - before.externalMb,
				arrayBuffersMb: after.arrayBuffersMb - before.arrayBuffersMb,
				saveMs: saved.ms,
				saveCpuMs,
				editedMb: editedMb(sessions),
				gate2Ms: saved.written.timings.verifyPartsMs + saved.written.timings.verifyZipMs,
				longestTaskMs: saved.tasks.maxMs,
				longestTaskCpuMs: saved.tasks.maxCpuMs,
				keysP95: percentile(keys.slice(keys.length - sinceSave), 0.95)
			}

			saves.push(entry)
			options.onSave?.(entry)

			if (last) {
				bytesOut = saved.written.bytes
				text = editor.body.session.state.doc.textContent
			}
		} catch (error) {
			failedSaves.push(error instanceof Error ? `${error.name}: ${error.message}` : String(error))
		}

		sinceSave = 0
	}

	const special = async (kind: Special): Promise<void> => {
		switch (kind) {
			case "undo": {
				const depth = 1 + (operations % 5)

				for (let k = 0; k < depth; k++) editor.execute({ type: "history.undo" })
				for (let k = 0; k < depth; k++) editor.execute({ type: "history.redo" })

				break
			}
			case "format":
				selectAll(editor)
				editor.execute({ type: "format.bold" })
				editor.snapshot()
				editor.execute({ type: "history.undo" })
				placeCaret(editor, spreadPosition(editor.session.state, operations))

				break
			case "replace": {
				const [from, to] = replaced ? [word.toUpperCase(), word] : [word, word.toUpperCase()]
				const found = replaceAll(editor.session.state, from, to, { caseSensitive: true })

				if (found !== null && "tr" in found) editor.dispatch(found.tr)

				replaced = !replaced

				break
			}
			case "comment": {
				const step = commentStep++ % 3

				if (step === 0 || lastComment === undefined) {
					const known = new Set(commentIds(editor))

					addComments(editor, 1)
					lastComment = commentIds(editor).find(id => !known.has(id))
				} else if (step === 1) {
					editor.execute({ type: "comment.reply", id: lastComment, text: "Soak reply" })
				} else {
					editor.execute({ type: "comment.resolve", id: lastComment })
				}

				break
			}
			case "table":
				if (!intoTable(editor, operations)) {
					placeCaret(editor, spreadPosition(editor.session.state, operations))
					editor.execute({ type: "table.insert", rows: 3, cols: 3 })

					break
				}

				if (operations % 2 === 0) {
					editor.execute({ type: tableInserted ? "table.row.delete" : "table.row.insertBelow" })
					tableInserted = !tableInserted
				} else {
					editor.execute({ type: columnInserted ? "table.column.delete" : "table.column.insertRight" })
					columnInserted = !columnInserted
				}

				placeCaret(editor, spreadPosition(editor.session.state, operations))

				break
			case "view":
				markup = !markup
				opened.editing.setView(markup ? "markup" : "final")
				await editor.whenSettled()

				break
		}

		editor.snapshot()
	}

	try {
		do {
			for (const kind of schedule(plan)) {
				if (kind === undefined) {
					if (keyCount % plan.perSpot === 0)
						placeCaret(editor, spreadPosition(editor.session.state, keyCount / plan.perSpot + passes * 997))
					if (keyCount % plan.trackEvery === 0 && keyCount > 0) editor.execute({ type: "review.track" })

					const done = key(editor, KINDS[keyCount % KINDS.length] ?? "type")

					keyCount++

					if (done !== undefined) {
						keys.push(done.ms)

						if (!done.applied) refused++
					}

					await nextTurn()
				} else {
					await special(kind)
				}

				operations++
				sinceSave++

				if (operations % plan.saveEvery === 0) await doSave()
			}

			passes++
		} while (performance.now() < deadline || passes < (options.passes ?? 1))

		await doSave(true)

		const sessions = opened.editing.saveInput()?.sessions ?? []

		historyLimitMb = sessions.reduce((sum, session) => sum + contentMb(session.state.doc), 0)
		clearHistories(sessions)
		await editor.whenSettled()
		clearedHeapMb = memory().heapMb - before.heapMb - (codeMb() - codeBefore)
	} finally {
		opened.close()
	}

	const historyMb = (saves[saves.length - 1]?.heapMb ?? clearedHeapMb) - clearedHeapMb

	return {
		openHeapMb,
		openContentMb,
		clearedHeapMb,
		historyMb,
		historyLimitMb,
		saves,
		keys,
		refused,
		failedSaves,
		passes,
		bytes: bytesOut,
		text
	}
}

// V8's compiled code (code and trusted spaces: bytecode, machine code), MB in use.
function codeMb(): number {
	let bytes = 0

	for (const space of getHeapSpaceStatistics())
		if (space.space_name === "code_space" || space.space_name === "trusted_space") bytes += space.space_used_size

	return bytes / 1048576
}

// About what the sessions' documents take, with the nodes of each story as built that its document no longer shares
// (the writer keeps those: a block an edit changed is written from its original's bytes where it can).
function documentsMb(sessions: readonly { readonly state: EditorState; readonly story: { readonly doc: Node } }[]): number {
	let bytes = 0

	for (const session of sessions) {
		const live = new Set<Node>()

		session.state.doc.descendants(node => {
			live.add(node)
			bytes += nodeBytes(node)
		})
		session.story.doc.descendants(node => {
			if (live.has(node)) return false

			bytes += nodeBytes(node)

			return true
		})
	}

	return bytes / 1048576
}

// What was edited since the open: the top-level blocks of each session's document that are not the ones it
// opened with, as documentsMb counts them.
function editedMb(sessions: readonly { readonly state: EditorState; readonly story: { readonly doc: Node } }[]): number {
	let bytes = 0

	for (const session of sessions) {
		const opened = new Set<Node>()

		session.story.doc.forEach(child => opened.add(child))
		session.state.doc.forEach(child => {
			if (opened.has(child)) return

			bytes += nodeBytes(child)
			child.descendants(node => {
				bytes += nodeBytes(node)
			})
		})
	}

	return bytes / 1048576
}

const nodeBytes = (node: Node): number => NODE_BYTES + (node.isText ? (node.text?.length ?? 0) * 2 : 0)

// About what one copy of a document's content takes: a node with its fragment and child array, its text.
function contentMb(doc: Node): number {
	let bytes = 0

	doc.descendants(node => {
		bytes += nodeBytes(node)
	})

	return bytes / 1048576
}

const NODE_BYTES = 200

// Every session's undo history dropped (a measurement after the session's last save: nothing is undone after).
function clearHistories(sessions: readonly { readonly state: EditorState }[]): void {
	for (const session of sessions) {
		const state = session.state

		// The session keeps its state in `current`; a state made from the same document and plugins has an empty
		// history.
		Reflect.set(session, "current", EditorState.create({ doc: state.doc, selection: state.selection, plugins: state.plugins }))
	}
}

export interface SoakLimits {
	// Final heap over the open's at most this (MB).
	readonly growthMb: number
}

// Least-squares slope of the last 15 of `values`, per step.
function tailSlope(values: readonly number[]): number {
	const tail = values.slice(-15)
	const n = tail.length
	const meanX = (n - 1) / 2
	const meanY = tail.reduce((sum, value) => sum + value, 0) / n
	let num = 0
	let den = 0

	tail.forEach((value, k) => {
		num += (k - meanX) * (value - meanY)
		den += (k - meanX) ** 2
	})

	return den === 0 ? 0 : num / den
}

// The soak's thresholds: what failed, empty when none did.
export function soakFailures(result: SoakResult, limits: SoakLimits): string[] {
	const failures: string[] = []
	const saves = result.saves

	if (result.failedSaves.length > 0) failures.push(`saves failed: ${result.failedSaves.join("; ")}`)
	if (saves.length < 3) return [...failures, `only ${String(saves.length)} saves`]

	const n = Math.min(15, saves.length)
	const content = tailSlope(saves.map(save => save.contentMb))
	const last = saves[saves.length - 1]
	const lastContent = last?.contentMb ?? result.openContentMb

	// Heap less the content: flat from save to save (under 1.25 MB per save over the last 15).
	const slope = tailSlope(saves.map(save => save.heapMb - save.contentMb))

	if (slope >= 1.25) failures.push(`heap less content grows ${slope.toFixed(2)} MB per save over the last ${String(n)} saves`)

	// Final heap within the open's, the limit, and 1.5 x what the content grew by.
	const bound = result.openHeapMb + limits.growthMb + Math.max(0, lastContent - result.openContentMb) * 1.5

	if ((last?.heapMb ?? 0) > bound) failures.push(`final heap ${(last?.heapMb ?? 0).toFixed(1)} MB over its bound ${bound.toFixed(1)} MB`)

	// Memory outside the heap: at most 6 x the content's growth + 0.5 MB per save (a buffer kept from each save shows).
	const buffers = tailSlope(saves.map(save => save.arrayBuffersMb))

	if (buffers > Math.max(0, content) * 6 + 0.5)
		failures.push(
			`array buffers grow ${buffers.toFixed(2)} MB per save over the last ${String(n)} saves (content ${content.toFixed(2)})`
		)

	if (result.historyMb > result.historyLimitMb)
		failures.push(
			`the undo history holds ${result.historyMb.toFixed(1)} MB, over the ${result.historyLimitMb.toFixed(1)} MB of one copy of the content`
		)

	const decile = Math.max(1, Math.floor(result.keys.length / 10))
	const firstP95 = percentile(result.keys.slice(0, decile), 0.95)
	const lastP95 = percentile(result.keys.slice(-decile), 0.95)

	if (lastP95 > firstP95 * 1.2) failures.push(`key p95 drifted from ${firstP95.toFixed(2)} to ${lastP95.toFixed(2)} ms`)

	// A save costs what was edited since the open: the last's CPU per edited MB at most 1.25 x the second's.
	const perMb = (save: SoakSave | undefined): number => (save === undefined ? 0 : save.saveCpuMs / Math.max(save.editedMb, 0.01))
	const second = perMb(saves[1])

	if (perMb(last) > second * 1.25)
		failures.push(`last save ${perMb(last).toFixed(0)} ms CPU per edited MB over 1.25 x the second's ${second.toFixed(0)}`)

	// No save task over 50 ms of CPU.
	const longest = Math.max(...saves.map(save => save.longestTaskCpuMs))

	if (longest > 50) failures.push(`a save task took ${longest.toFixed(0)} ms CPU`)

	return failures
}
