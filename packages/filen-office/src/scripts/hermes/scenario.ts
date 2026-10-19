import { TextSelection } from "prosemirror-state"
import type { Node } from "prosemirror-model"
import { legacyText, type LegacyField } from "../../core/docx/forms/controls"
import { prepareDocxSave } from "../../core/docx/write/save"
import { writerProof } from "../../core/docx/write/proof"
import { bytesSource } from "../../core/opc/index"
import { FRAGMENT_INTS, FragmentInt, type FontExtents, type TextMeasurer } from "../../core/text/index"
import { inThreadEditingHost } from "../../dom/docx/editing"
import type { DocxEngineServices } from "../../dom/docx/engine"
import { inThreadDocxHost, macrotaskYield } from "../../dom/docx/host"
import { FIND_INTS, type DocxEvent, type DocxPageData } from "../../dom/docx/protocol"

// The core's open → layout → edit → save path, the way the mobile DOM component runs it (in-thread host, no
// Worker, MessageChannel or OffscreenCanvas), written so the same code runs under Node and Hermes and their
// results can be compared byte for byte. Only ECMAScript builtins and setTimeout are used.

export interface ScenarioFixture {
	readonly name: string
	readonly bytes: Uint8Array
}

export interface FixtureResult {
	readonly name: string
	readonly sourceBytes: number
	// Pages and a digest of their layout (structure: counts and texts; geometry: the float bits).
	readonly pages: number
	readonly structure: string
	readonly geometry: string
	// The save of the document as opened, and whether it equals the source.
	readonly unedited: Uint8Array
	readonly uneditedIdentical: boolean
	// What each edit step returned (false or "refused" where the session refused it).
	readonly steps: readonly string[]
	readonly edited: Uint8Array
	readonly changedParts: readonly string[]
	readonly editedPages: number
	readonly editedStructure: string
	readonly editedGeometry: string
	// The edited file opened again: its pages and the matches of the inserted text.
	readonly reopenedPages: number
	readonly reopenedMatches: number
	readonly ms: { readonly open: number; readonly save: number; readonly edit: number; readonly reopen: number }
}

export interface ProbeResult {
	readonly name: string
	readonly value: string
}

export interface ScenarioResult {
	readonly fixtures: readonly FixtureResult[]
	readonly probes: readonly ProbeResult[]
	readonly errors: readonly string[]
}

// Every character half an em wide: layout depends on nothing the engine does not compute itself.
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

// A clock that advances per reading, so time slices (and with them the order pages and styles are shipped
// in) fall the same way in every engine however fast it runs.
function virtualClock(): () => number {
	let now = 0

	return () => (now += 0.25)
}

function services(now: () => number): DocxEngineServices {
	return {
		fontHost: { addBytes: () => Promise.resolve(true), addLocal: () => Promise.resolve(false), remove: () => undefined },
		replayFonts: false,
		fontAssets: { read: () => Promise.resolve(new Uint8Array(4)) },
		measurer,
		// The DOM component's turn: a macrotask (setTimeout where there is no MessageChannel) every 6 ms.
		yield: macrotaskYield(),
		sliceMs: 6,
		now
	}
}

const CLOCK = Date.UTC(2026, 0, 2, 3, 4, 5)
const INSERTED = "Hermes \u2713 \u00E9 \u65E5\u672C \u{1F44D}\u{1F3FD} "

// New paragraph ids, comment ids and the like draw on Math.random: seeded, so both engines write the same.
function seedRandom(seed: number): void {
	let state = seed >>> 0

	Math.random = () => {
		state = (state + 0x6d2b79f5) >>> 0

		let t = state

		t = Math.imul(t ^ (t >>> 15), t | 1)
		t ^= t + Math.imul(t ^ (t >>> 7), t | 61)

		return ((t ^ (t >>> 14)) >>> 0) / 4294967296
	}
}

function fnv(hash: number, bytes: Uint8Array): number {
	let h = hash

	for (const byte of bytes) h = Math.imul(h ^ byte, 16777619) >>> 0

	return h
}

function fnvText(hash: number, text: string): number {
	let h = hash

	for (let i = 0; i < text.length; i++) h = Math.imul(h ^ text.charCodeAt(i), 16777619) >>> 0

	return h
}

function viewBytes(array: ArrayBufferView): Uint8Array {
	return new Uint8Array(array.buffer, array.byteOffset, array.byteLength)
}

// Pages ship with style indices into tables the events append to, which may repeat a style: the digest takes
// each fragment's style by value, so it does not depend on how the pages were batched.
function digest(pages: readonly (DocxPageData | undefined)[], styles: readonly string[]): { structure: string; geometry: string } {
	let structure = 2166136261
	let geometry = 2166136261

	for (const page of pages) {
		if (page === undefined) {
			structure = fnvText(structure, "missing")

			continue
		}

		structure = fnvText(
			structure,
			`${page.lineCount}/${page.fragmentCount}/${page.objectCount}/${page.graphicCount}/${page.sourceCount}`
		)

		for (const text of page.sourceTexts) structure = fnvText(structure, text)

		const fragments = page.fragmentInts.slice()

		for (let at = FragmentInt.STYLE; at < fragments.length; at += FRAGMENT_INTS) {
			fragments[at] = fnvText(2166136261, styles[fragments[at] ?? -1] ?? "") | 0
		}

		for (const array of [page.lineInts, fragments, page.sourceInts]) structure = fnv(structure, viewBytes(array))
		for (const array of [page.lineFloats, page.fragmentFloats, page.objectFloats, page.graphicFloats]) {
			geometry = fnv(geometry, viewBytes(array))
		}
	}

	return { structure: structure.toString(16), geometry: geometry.toString(16) }
}

// Font aliases carry a namespace per font manager (random, and counted within a realm): read without it.
const FONT_ALIAS = /\\"fo[0-9a-z]+-/g

class PageLog {
	readonly pages: (DocxPageData | undefined)[] = []
	readonly styles: string[] = []
	count = 0
	error: string | undefined
	private waiters: (() => void)[] = []
	private final = false

	readonly listener = (event: DocxEvent): void => {
		if (event.type === "error") {
			this.error = `${event.reason}: ${event.message}`
			this.release()
		} else if (event.type === "pages") {
			for (const page of event.pages) this.pages[page.index] = page
			for (const style of event.styles) this.styles.push(JSON.stringify(style).replace(FONT_ALIAS, '\\"fo-'))

			this.count = event.pageCount

			if (event.final) {
				this.final = true
				this.release()
			}
		}
	}

	settled(): Promise<void> {
		if (this.final || this.error !== undefined) return Promise.resolve()

		return new Promise(resolve => this.waiters.push(resolve))
	}

	snapshot(): (DocxPageData | undefined)[] {
		return this.pages.slice(0, this.count)
	}

	private release(): void {
		for (const waiter of this.waiters.splice(0)) waiter()
	}
}

function firstTextPos(doc: Node): number | undefined {
	let found: number | undefined

	doc.descendants((node, pos) => {
		if (found !== undefined) return false
		if (node.isTextblock && node.textContent.length > 0) {
			found = pos + 1

			return false
		}

		return true
	})

	return found
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
	if (a.length !== b.length) return false

	for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false

	return true
}

async function runFixture(fixture: ScenarioFixture): Promise<FixtureResult> {
	seedRandom(0x0ff1ce)

	const started = Date.now()
	const log = new PageLog()
	const now = virtualClock()
	const editing = inThreadEditingHost(() => services(now), {
		proof: writerProof,
		paged: { now },
		writer: { prepare: prepareDocxSave },
		author: "Hermes Runner",
		clock: () => CLOCK
	})
	const session = editing.host(bytesSource(fixture.bytes), log.listener)

	try {
		const editor = await editing.editor

		await log.settled()

		if (log.error !== undefined) throw new Error(log.error)

		const opened = digest(log.snapshot(), log.styles)
		const openedPages = log.count
		const openMs = Date.now() - started
		let at = Date.now()
		const unedited = (await editing.save()).bytes
		const saveMs = Date.now() - at

		at = Date.now()

		const steps: string[] = []
		const step = (label: string, run: () => boolean | Promise<boolean>): Promise<void> =>
			Promise.resolve(run()).then(
				done => {
					steps.push(`${label}:${done}`)
				},
				(error: unknown) => {
					steps.push(`${label}:threw ${error instanceof Error ? error.message : String(error)}`)
				}
			)
		const start = firstTextPos(editor.state.doc)

		if (start !== undefined) {
			await step("insert", () => {
				const state = editor.state

				return editor.dispatch(state.tr.setSelection(TextSelection.create(state.doc, start)).insertText(INSERTED))
			})
			await step("select", () => {
				const state = editor.state

				return editor.dispatch(state.tr.setSelection(TextSelection.create(state.doc, start, start + 6)))
			})
			await step("bold", () => editor.execute({ type: "format.bold" }))
			await step("split", () => {
				const state = editor.state
				const pos = start + INSERTED.length

				return editor.dispatch(state.tr.setSelection(TextSelection.create(state.doc, pos)).split(pos))
			})
			await step("table", () => editor.execute({ type: "table.insert", rows: 2, cols: 2 }))
			await step("track", () => editor.execute({ type: "review.track", on: true }))
			await step("tracked", () => {
				const state = editor.state

				return editor.dispatch(state.tr.setSelection(TextSelection.create(state.doc, start)).insertText("Tracked "))
			})
			await step("undo", () => editor.execute({ type: "history.undo" }))
			await step("redo", () => editor.execute({ type: "history.redo" }))
		}

		await editor.whenSettled()
		await new Promise<void>(resolve => setTimeout(resolve, 0))

		const editMs = Date.now() - at
		const editedLayout = digest(log.snapshot(), log.styles)
		const written = await editing.save()

		at = Date.now()

		const reopened = new PageLog()
		const viewer = inThreadDocxHost(() => services(virtualClock()))(bytesSource(written.bytes), reopened.listener)

		await reopened.settled()

		const matches = reopened.error === undefined ? (await viewer.find("Hermes")).length / FIND_INTS : -1

		viewer.close()

		return {
			name: fixture.name,
			sourceBytes: fixture.bytes.length,
			pages: openedPages,
			structure: opened.structure,
			geometry: opened.geometry,
			unedited,
			uneditedIdentical: sameBytes(unedited, fixture.bytes),
			steps,
			edited: written.bytes,
			changedParts: [...written.changedParts].sort(),
			editedPages: log.count,
			editedStructure: editedLayout.structure,
			editedGeometry: editedLayout.geometry,
			reopenedPages: reopened.error === undefined ? reopened.count : -1,
			reopenedMatches: matches,
			ms: { open: openMs, save: saveMs, edit: editMs, reopen: Date.now() - at }
		}
	} finally {
		session.close()
	}
}

const TEXT_FIELD: LegacyField = {
	type: "FORMTEXT",
	name: "Text1",
	enabled: true,
	checked: false,
	entries: [],
	result: 0,
	textType: "regular",
	maxLength: 3,
	format: "UPPERCASE"
}

// Builtins the core reads that engines differ on: what they give here.
function probes(): ProbeResult[] {
	const out: ProbeResult[] = []
	const probe = (name: string, run: () => unknown): void => {
		try {
			out.push({ name, value: String(run()) })
		} catch (error) {
			out.push({ name, value: `threw ${error instanceof Error ? `${error.name}: ${error.message}` : String(error)}` })
		}
	}

	probe("forms.legacyText maxLength", () => legacyText(TEXT_FIELD, "e\u0301a\u{1F44D}\u{1F3FD}\u{1F1E9}\u{1F1EA}x"))
	probe("forms.legacyText title case", () => legacyText({ ...TEXT_FIELD, maxLength: 0, format: "Title case" }, "\u00E9lan vital \u00DFx"))

	return out
}

export async function runScenario(fixtures: readonly ScenarioFixture[]): Promise<ScenarioResult> {
	const results: FixtureResult[] = []
	const errors: string[] = []

	for (const fixture of fixtures) {
		try {
			results.push(await runFixture(fixture))
		} catch (error) {
			errors.push(`${fixture.name}: ${error instanceof Error ? `${error.name}: ${error.message}` : String(error)}`)
		}
	}

	return { fixtures: results, probes: probes(), errors }
}
