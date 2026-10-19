import { setFlagsFromString } from "node:v8"
import { runInNewContext } from "node:vm"
import { Worker } from "node:worker_threads"
import { Fragment, Slice } from "prosemirror-model"
import { AllSelection, TextSelection, type EditorState, type Transaction } from "prosemirror-state"
import { docxComments, type PagedEditor } from "@filen/office/core/docx/paged"
import { bytesSource } from "@filen/office/core/opc"
import type { FontExtents, TextMeasurer } from "@filen/office/core/text"
import { prepareDocxSave, writerProof } from "@filen/office/core/docx/write"
import type { DocxEngineServices, DocxOpenOptions } from "@filen/office/dom/docx"
import { inThreadEditingHost, type DocxEditingHost, type DocxWritten } from "@filen/office/dom/docx/editing"
import { docxZipWorker } from "@filen/office/dom/docx/save"

// The editor as the web opens it (in-thread editing host, Gate 1 proof in slices, the writer with its Gate 2,
// 6 ms slices), on node: text measured as half the font size per character, and each slice ends its event
// loop turn (setImmediate), so a turn is one task and TaskMonitor sees each.

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

export const nextTurn = (): Promise<void> =>
	new Promise(resolve => {
		setImmediate(resolve)
	})

function services(): DocxEngineServices {
	return {
		fontHost: { addBytes: () => Promise.resolve(true), addLocal: () => Promise.resolve(false), remove: () => undefined },
		replayFonts: false,
		fontAssets: { read: () => Promise.resolve(new Uint8Array(4)) },
		measurer,
		yield: nextTurn,
		sliceMs: 6,
		now: () => performance.now()
	}
}

export function percentile(values: readonly number[], share: number): number {
	const sorted = [...values].sort((a, b) => a - b)

	return sorted[Math.min(sorted.length - 1, Math.floor(sorted.length * share))] ?? 0
}

// node's gc() without --expose-gc on the command line (vitest's workers), as the flag exposes it.
const collect: () => void =
	(globalThis as { gc?: () => void }).gc ??
	(() => {
		setFlagsFromString("--expose-gc")

		return runInNewContext("gc") as () => void
	})()

const gc = (): void => {
	collect()
	collect()
}

export interface Memory {
	readonly heapMb: number
	readonly externalMb: number
	readonly arrayBuffersMb: number
}

// Memory after a full collection.
export function memory(): Memory {
	gc()

	const usage = process.memoryUsage()

	return { heapMb: usage.heapUsed / 1048576, externalMb: usage.external / 1048576, arrayBuffersMb: usage.arrayBuffers / 1048576 }
}

// This thread's CPU time, ms: what a task costs whatever else the machine runs (wall time also counts the
// time other processes held the core).
export function cpuMs(): number {
	const usage = process.threadCpuUsage()

	return (usage.user + usage.system) / 1000
}

export interface Tasks {
	// Longest turn of the event loop, ms (a turn is at least the 1 ms tick).
	readonly maxMs: number
	// The same in this thread's CPU time, and the turns over 50 ms of it.
	readonly maxCpuMs: number
	readonly cpuOver50: number
	readonly over50: number
	readonly over80: number
	readonly over100: number
	readonly turns: number
}

// Event loop turns while running: a 1 ms timer ticks once a turn, the gap between ticks is the turn's work.
export class TaskMonitor {
	private timer: ReturnType<typeof setTimeout> | undefined
	private last = 0
	private lastCpu = 0
	private max = 0
	private maxCpu = 0
	private cpuOver50 = 0
	private over50 = 0
	private over80 = 0
	private over100 = 0
	private turns = 0

	start(): this {
		this.last = performance.now()
		this.lastCpu = cpuMs()

		const tick = (): void => {
			const now = performance.now()
			const gap = now - this.last
			const cpu = cpuMs()
			const cpuGap = cpu - this.lastCpu

			this.last = now
			this.lastCpu = cpu
			this.turns++
			this.max = Math.max(this.max, gap)
			this.maxCpu = Math.max(this.maxCpu, cpuGap)

			if (cpuGap > 50) this.cpuOver50++

			if (gap > 50) this.over50++
			if (gap > 80) this.over80++
			if (gap > 100) this.over100++

			this.timer = setTimeout(tick, 1)
		}

		this.timer = setTimeout(tick, 1)

		return this
	}

	stop(): Tasks {
		clearTimeout(this.timer)
		this.timer = undefined

		// The turn running now counts too.
		const gap = performance.now() - this.last
		const cpuGap = cpuMs() - this.lastCpu

		return {
			maxMs: Math.max(this.max, gap),
			maxCpuMs: Math.max(this.maxCpu, cpuGap),
			cpuOver50: this.cpuOver50 + (cpuGap > 50 ? 1 : 0),
			over50: this.over50 + (gap > 50 ? 1 : 0),
			over80: this.over80 + (gap > 80 ? 1 : 0),
			over100: this.over100 + (gap > 100 ? 1 : 0),
			turns: this.turns
		}
	}
}

export interface OpenedEditor {
	readonly editing: DocxEditingHost
	readonly editor: PagedEditor
	// Open to an editable editor, and to the layout's last page.
	readonly editorMs: number
	// The first snapshot (what the toolbar reads once the editor opens).
	readonly firstSnapshotMs: number
	readonly settledMs: number
	readonly tasks: Tasks
	close(): void
}

export interface OpenOptions {
	readonly view?: DocxOpenOptions["view"]
	readonly author?: string
	// Zip saves in a worker thread, as the web does: what a save costs the main thread is its prepare alone.
	readonly zipWorker?: boolean
}

export async function openEditor(bytes: Uint8Array, options: OpenOptions = {}): Promise<OpenedEditor> {
	const editing = inThreadEditingHost(services, {
		proof: writerProof,
		writer:
			options.zipWorker === true
				? {
						prepare: prepareDocxSave,
						zipper: docxZipWorker(() => {
							const worker = new Worker(new URL("./zipWorker.ts", import.meta.url), { execArgv: ["--import", "tsx"] })

							return {
								post: (message, transfer) => {
									worker.postMessage(message, [...transfer])
								},
								listen: (onMessage, onError) => {
									worker.on("message", onMessage)
									worker.on("error", onError)
								},
								terminate: () => {
									void worker.terminate()
								}
							}
						})
					}
				: { prepare: prepareDocxSave },
		author: options.author ?? "Perf",
		clock: () => Date.UTC(2026, 0, 1)
	})
	const monitor = new TaskMonitor().start()
	const start = performance.now()
	const session = editing.host(bytesSource(bytes), () => undefined, {
		retain: true,
		...(options.view === undefined ? {} : { view: options.view })
	})
	const editor = await editing.editor
	const editorMs = performance.now() - start
	// The toolbar reads the editor's state as it shows it.
	const read = performance.now()

	editor.snapshot()

	const firstSnapshotMs = performance.now() - read

	await editor.whenSettled()

	const settledMs = performance.now() - start

	return { editing, editor, editorMs, firstSnapshotMs, settledMs, tasks: monitor.stop(), close: () => session.close() }
}

export type KeyKind = "type" | "backspace" | "enter" | "paste"

// A textblock position spread over the document by `k`, as a click there would put the caret.
export function spreadPosition(state: EditorState, k: number): number {
	const size = state.doc.content.size
	let target = Math.floor((((k * 7919) % 10007) / 10007) * size)

	while (target < size) {
		const $pos = state.doc.resolve(target)

		if ($pos.parent.inlineContent && $pos.parent.content.size > 0) return target

		target++
	}

	return TextSelection.atStart(state.doc).from
}

// Puts the caret at `pos` (a click), and scrolls its page into view as the web view does.
export function placeCaret(editor: PagedEditor, pos: number): void {
	const state = editor.session.state

	editor.dispatch(state.tr.setSelection(TextSelection.near(state.doc.resolve(pos))))

	const caret = editor.caret()

	if (caret !== undefined && !editor.suspended) editor.setVisible(caret.page, caret.page + 1)
}

let clipboard: Slice | undefined

function keyTransaction(state: EditorState, kind: KeyKind): Transaction | undefined {
	const { $from, from, empty } = state.selection

	switch (kind) {
		case "type":
			return state.tr.insertText("a")
		case "backspace":
			if (!empty) return state.tr.deleteSelection()

			return $from.parentOffset > 0 ? state.tr.delete(from - 1, from) : undefined
		case "enter":
			return $from.parent.inlineContent ? state.tr.split(from) : undefined
		case "paste":
			// A few plain words, as a copy from elsewhere brings.
			clipboard ??= new Slice(Fragment.from(state.schema.text("pasted text ")), 0, 0)

			return state.tr.replaceSelection(clipboard)
	}
}

// One key as the editor's input view applies it, and what the toolbar reads after it: the time in the key's
// task (the layout of the pages in view included; the rest of the layout runs in later turns). Undefined
// when the key does nothing there.
export function key(editor: PagedEditor, kind: KeyKind): { ms: number; cpuMs: number; applied: boolean } | undefined {
	const start = performance.now()
	const cpu = cpuMs()
	const tr = keyTransaction(editor.session.state, kind)

	if (tr === undefined) return undefined

	const applied = editor.dispatch(tr)

	editor.snapshot()

	return { ms: performance.now() - start, cpuMs: cpuMs() - cpu, applied }
}

export interface Burst {
	readonly times: number[]
	// The same keys in this thread's CPU time.
	readonly cpu: number[]
	refused: number
}

// Between two keys `intervalMs` apart: the layout carries on in its slices until it settles or the next key
// comes (one turn when 0, a key held down).
export async function between(editor: PagedEditor, intervalMs: number): Promise<void> {
	if (intervalMs <= 0 || editor.suspended) return nextTurn()

	let timer: ReturnType<typeof setTimeout> | undefined

	await Promise.race([
		editor.whenSettled(),
		new Promise<void>(resolve => {
			timer = setTimeout(resolve, intervalMs)
		})
	])
	clearTimeout(timer)
}

// `spots` places spread over the document, `perSpot` keys at each, `intervalMs` apart (100: a fast typist);
// the layout settles before the next place. `offset` picks other places.
export async function typeBursts(
	editor: PagedEditor,
	spots: number,
	perSpot: number,
	kinds: readonly KeyKind[] = ["type"],
	offset = 0,
	intervalMs = 100
): Promise<Burst> {
	const burst: Burst = { times: [], cpu: [], refused: 0 }

	for (let spot = 0; spot < spots; spot++) {
		placeCaret(editor, spreadPosition(editor.session.state, spot + offset))

		for (let k = 0; k < perSpot; k++) {
			const done = key(editor, kinds[k % kinds.length] ?? "type")

			if (done !== undefined) {
				burst.times.push(done.ms)
				burst.cpu.push(done.cpuMs)

				if (!done.applied) burst.refused++
			}

			await between(editor, intervalMs)
		}

		// Suspended (Mobile view shows the document) nothing is laid out until resume.
		if (!editor.suspended) await editor.whenSettled()
	}

	return burst
}

// Comments on words spread over the body (the word at the caret), as the review pane adds them.
export function addComments(editor: PagedEditor, count: number): number {
	let added = 0

	for (let k = 0; k < count; k++) {
		placeCaret(editor, spreadPosition(editor.session.state, k * 31 + 17))

		const state = editor.session.state
		const { $from } = state.selection
		const from = $from.start() + Math.min($from.parentOffset, Math.max(0, $from.parent.content.size - 1))
		const to = Math.min($from.end(), from + 4)

		if (to > from) editor.dispatch(state.tr.setSelection(TextSelection.create(state.doc, from, to)))
		if (editor.execute({ type: "comment.add", text: `Perf comment ${String(k)}` }) === true) added++
	}

	return added
}

export function commentIds(editor: PagedEditor): string[] {
	return docxComments(editor)
		.threads()
		.map(thread => thread.id)
}

export function selectAll(editor: PagedEditor): void {
	const state = editor.session.state

	editor.dispatch(state.tr.setSelection(new AllSelection(state.doc)))
}

export interface SaveResult {
	readonly ms: number
	readonly written: DocxWritten
	readonly tasks: Tasks
}

// A save as the web's writeDocx makes it (sessions marked saved with what was written).
export async function save(editing: DocxEditingHost): Promise<SaveResult> {
	const input = editing.saveInput()

	if (input === undefined) throw new Error("not open for editing")

	const sessions = [...input.sessions]
	const docs = sessions.map(session => session.state.doc)
	const monitor = new TaskMonitor().start()
	const start = performance.now()
	const written = await editing.save(sessions)
	const ms = performance.now() - start
	const tasks = monitor.stop()

	sessions.forEach((session, index) => {
		const doc = docs[index]

		if (doc !== undefined) session.markSaved(doc)
	})

	return { ms, written, tasks }
}
