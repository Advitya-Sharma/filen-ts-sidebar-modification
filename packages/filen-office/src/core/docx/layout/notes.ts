import type { DocxModel } from "../model/document"
import { Ns } from "../model/namespaces"
import { partReader } from "../model/reader"
import type { NotePositionProps } from "../props/types"
import type { NoteNumbers, PageValues, StoryContext } from "./content"
import { numberText } from "./fields"
import type { Engine, FootnoteSink, PageFrame } from "./flow"
import { PageGraphic, Layer, PageOutput, Region } from "./output"
import { EPSILON, points } from "./units"

// Note numbering (w:footnotePr / w:endnotePr in settings, overridden per section) and the footnote area at
// the bottom of each page.

export interface NoteSettings {
	readonly format: string
	readonly start: number
	// "continuous", "eachSect", "eachPage".
	readonly restart: string
	// "pageBottom", "beneathText", "sectEnd", "docEnd".
	readonly position: string
}

function readNotePr(model: DocxModel, node: number, endnote: boolean): NoteSettings {
	const out = { format: endnote ? "lowerRoman" : "decimal", start: 1, restart: "continuous", position: endnote ? "docEnd" : "pageBottom" }
	const document = model.settings.document

	if (document === undefined || node < 0) return out

	const reader = partReader(document)

	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) {
		const value = reader.attr(child, Ns.W, "val")

		if (value === undefined) continue

		switch (document.localName(child)) {
			case "numFmt":
				out.format = value
				break
			case "numStart":
				out.start = Number.parseInt(value, 10) || 1
				break
			case "numRestart":
				out.restart = value
				break
			case "pos":
				out.position = value
				break
		}
	}

	return out
}

function override(base: NoteSettings, section: NotePositionProps | undefined): NoteSettings {
	if (section === undefined) return base

	return {
		format: section.numFormat ?? base.format,
		start: section.numStart ?? base.start,
		restart: section.numRestart ?? base.restart,
		position: section.position ?? base.position
	}
}

export interface NoteNumberingSnapshot {
	readonly footnotes: NoteSettings
	readonly endnotes: NoteSettings
	readonly footnoteNext: number
	readonly endnoteNext: number
}

export class NoteNumbering implements NoteNumbers {
	readonly documentFootnotes: NoteSettings
	readonly documentEndnotes: NoteSettings
	footnotes: NoteSettings
	endnotes: NoteSettings
	private footnoteNext: number
	private endnoteNext: number
	private readonly assigned = new Map<number, string>()
	// The mark text of each note story.
	readonly marks = new Map<number, string>()
	// Footnotes already on the page (numbering that restarts on each page counts from them).
	pageNotes: () => number = () => 0
	// References numbered in the paragraph being built.
	private paragraphRefs = 0

	constructor(model: DocxModel) {
		this.documentFootnotes = readNotePr(model, model.settings.footnotePr, false)
		this.documentEndnotes = readNotePr(model, model.settings.endnotePr, true)
		this.footnotes = this.documentFootnotes
		this.endnotes = this.documentEndnotes
		this.footnoteNext = this.footnotes.start
		this.endnoteNext = this.endnotes.start
	}

	startSection(footnotes: NotePositionProps | undefined, endnotes: NotePositionProps | undefined, first: boolean): void {
		this.footnotes = override(this.documentFootnotes, footnotes)
		this.endnotes = override(this.documentEndnotes, endnotes)

		if (first || this.footnotes.restart === "eachSect") this.footnoteNext = this.footnotes.start
		if (first || this.endnotes.restart === "eachSect") this.endnoteNext = this.endnotes.start
	}

	startPage(): void {
		if (this.footnotes.restart === "eachPage") this.footnoteNext = this.footnotes.start
	}

	// Numbering state to restart layout from.
	snapshot(): NoteNumberingSnapshot {
		return { footnotes: this.footnotes, endnotes: this.endnotes, footnoteNext: this.footnoteNext, endnoteNext: this.endnoteNext }
	}

	restore(snapshot: NoteNumberingSnapshot): void {
		this.footnotes = snapshot.footnotes
		this.endnotes = snapshot.endnotes
		this.footnoteNext = snapshot.footnoteNext
		this.endnoteNext = snapshot.endnoteNext
	}

	beginParagraph(): void {
		this.paragraphRefs = 0
	}

	// Numbers assigned to references of `part` items for which `keep` holds, to carry into a snapshot.
	assignedWhere(keep: (part: number, item: number) => boolean): [number, string][] {
		const out: [number, string][] = []

		for (const [key, text] of this.assigned) {
			const reference = Math.floor(key / 2)
			const part = Math.floor(reference / 0x4000000)

			if (keep(part, reference - part * 0x4000000)) out.push([key, text])
		}

		return out
	}

	// Forgets the numbers of references in parts for which `drop` holds (layout numbers them again), then
	// takes `kept` back.
	forgetAssigned(drop: (part: number) => boolean, kept: readonly (readonly [number, string])[]): void {
		for (const key of this.assigned.keys()) if (drop(Math.floor(key / 2 / 0x4000000))) this.assigned.delete(key)

		for (const [key, text] of kept) this.assigned.set(key, text)
	}

	reference(endnote: boolean, part: number, item: number, story: number): string {
		// Per-page numbering is decided where the paragraph is laid out: after the notes already on its page.
		if (!endnote && this.footnotes.restart === "eachPage") {
			const text = numberText(this.footnotes.start + this.pageNotes() + this.paragraphRefs++, this.footnotes.format)

			this.marks.set(story, text)

			return text
		}

		const key = (part * 0x4000000 + item) * 2 + (endnote ? 1 : 0)
		let text = this.assigned.get(key)

		if (text === undefined) {
			const value = endnote ? this.endnoteNext++ : this.footnoteNext++

			text = numberText(value, endnote ? this.endnotes.format : this.footnotes.format)
			this.assigned.set(key, text)
		}

		// Kept from a snapshot, the number may differ from the one a later run gave the note.
		this.marks.set(story, text)

		return text
	}
}

interface LaidNote {
	readonly output: PageOutput
	readonly height: number
}

// A footnote area's state (FootnoteArea.save).
export interface FootnoteState {
	readonly notes: readonly PlacedNote[]
	readonly carry: readonly PlacedNote[]
	readonly left: number
	readonly width: number
	readonly reserved: number
	readonly splitAt: number
}

// A note to place on the page: its laid-out content from `from` (px into it; > 0 when continued).
export interface PlacedNote {
	readonly story: number
	readonly from: number
	readonly height: number
}

// The footnote area of the page being laid out.
export class FootnoteArea implements FootnoteSink {
	reserved = 0
	// The index of the last note, started on this page but too long to end here, or -1.
	private splitAt = -1
	private readonly engine: Engine
	private readonly notes: PlacedNote[] = []
	private readonly cache = new Map<number, LaidNote>()
	private carry: PlacedNote[] = []
	private separatorStory = -1
	private continuationStory = -1
	private width = 0
	private left = 0
	private values: PageValues = { page: 1, pageFormat: undefined, numPages: 1, sectionPages: 1 }
	private frame: PageFrame | undefined = undefined
	private separatorHeight = 0
	private continuationHeight = 0

	constructor(engine: Engine) {
		this.engine = engine

		const model = engine.model

		for (const id of model.settings.footnoteSeparators) {
			const story = model.footnotes.get(id)

			if (story?.noteType === "separator") this.separatorStory = story.index
			else if (story?.noteType === "continuationSeparator") this.continuationStory = story.index
		}

		for (const story of model.footnotes.values()) {
			if (story.noteType === "separator" && this.separatorStory < 0) this.separatorStory = story.index
			if (story.noteType === "continuationSeparator" && this.continuationStory < 0) this.continuationStory = story.index
		}
	}

	get carried(): boolean {
		return this.carry.length > 0
	}

	// Notes continued onto the next page, to restart layout there.
	get carriedNotes(): readonly PlacedNote[] {
		return this.carry.slice()
	}

	set carriedNotes(notes: readonly PlacedNote[]) {
		this.carry = notes.slice()
	}

	// Notes that start on this page.
	get count(): number {
		let count = 0

		for (const note of this.notes) if (note.from === 0) count++

		return count
	}

	// Layout restarts: notes are laid out again as they now read and are numbered (one laid out and left for the
	// next page, or continued there, stays in the cache meanwhile).
	restart(): void {
		this.cache.clear()
	}

	// A new page: the area spans [left, left + width); notes continued from the last page come first.
	startPage(left: number, width: number, values: PageValues, frame: PageFrame): void {
		if (Math.abs(width - this.width) > EPSILON) this.cache.clear()

		this.width = width
		this.values = values
		this.frame = frame
		this.separatorHeight = this.separatorStory < 0 ? points(12) : this.laid(this.separatorStory).height
		this.continuationHeight = this.continuationStory < 0 ? points(12) : this.laid(this.continuationStory).height
		this.startColumn(left, width)
	}

	// A new column with notes of its own, spanning [left, left + width); notes continued from the column
	// before come first.
	startColumn(left: number, width: number): void {
		if (Math.abs(width - this.width) > EPSILON) {
			this.cache.clear()
			this.width = width
			this.separatorHeight = this.separatorStory < 0 ? points(12) : this.laid(this.separatorStory).height
			this.continuationHeight = this.continuationStory < 0 ? points(12) : this.laid(this.continuationStory).height
		}

		this.left = left
		this.notes.length = 0
		this.reserved = 0

		for (const note of this.carry) this.notes.push(note)

		this.carry = []
		this.splitAt = -1
		this.reserved = this.notes.length === 0 ? 0 : this.continuationHeight + this.notes.reduce((sum, note) => sum + note.height, 0)
	}

	get split(): boolean {
		return this.splitAt >= 0
	}

	// The whole state, to lay a part of the page out again from here.
	save(): FootnoteState {
		return {
			notes: this.notes.slice(),
			carry: this.carry.slice(),
			left: this.left,
			width: this.width,
			reserved: this.reserved,
			splitAt: this.splitAt
		}
	}

	load(state: FootnoteState): void {
		if (Math.abs(state.width - this.width) > EPSILON) this.startColumn(state.left, state.width)

		this.notes.splice(0, this.notes.length, ...state.notes)
		this.carry = state.carry.slice()
		this.left = state.left
		this.reserved = state.reserved
		this.splitAt = state.splitAt
	}

	mark(): number {
		return this.notes.length
	}

	rollback(mark: number): void {
		if (mark >= this.notes.length) return

		this.notes.length = mark

		if (this.splitAt >= mark) this.splitAt = -1

		this.reserved = this.notes.length === 0 ? 0 : this.separator() + this.notes.reduce((sum, note) => sum + note.height, 0)
	}

	private separator(): number {
		return (this.notes[0]?.from ?? 0) > 0 ? this.continuationHeight : this.separatorHeight
	}

	add(story: number, lineBottom: number, limit: number, force: boolean): boolean {
		if (this.notes.some(note => note.story === story && note.from === 0)) return true

		if (this.splitAt >= 0) return force

		const laid = this.laid(story)
		const before = this.notes.length === 0 ? this.separatorHeight : this.reserved
		const needed = before + laid.height

		if (!force && lineBottom + needed > limit + EPSILON) {
			// A note too long for the room left starts here when its first line fits; the rest continues on the
			// next page, and the page takes no more text.
			const text = laid.output.text

			if (text.lineCount === 0 || lineBottom + before + text.lineTop(0) + text.lineHeight(0) > limit + EPSILON) return false

			this.splitAt = this.notes.length
			this.notes.push({ story, from: 0, height: laid.height })
			this.reserved = limit - lineBottom

			return true
		}

		this.notes.push({ story, from: 0, height: laid.height })
		this.reserved = needed

		return true
	}

	private laid(story: number): LaidNote {
		let laid = this.cache.get(story)

		if (laid === undefined) {
			const engine = this.engine
			const root = engine.model.stories[story]
			const part = root === undefined ? undefined : engine.model.parts[root.part]
			const output = new PageOutput()
			let height = 0

			if (root !== undefined && part !== undefined && this.frame !== undefined) {
				const context: StoryContext = {
					part,
					story,
					counters: undefined,
					noteMark: engine.notes.marks.get(story) ?? "",
					breaksAsLines: true
				}

				const trailing = engine.tuning.noteTrailingSpace
				const separator = root.noteType === "separator" || root.noteType === "continuationSeparator"

				height = engine.layoutBox(
					context,
					Region.FOOTNOTE,
					output,
					this.values,
					this.frame,
					0,
					0,
					this.width,
					trailing === "all" || (trailing === "notes" && !separator)
				)
			}

			laid = { output, height }
			this.cache.set(story, laid)
		}

		return laid
	}

	// Draws the notes at the bottom of the text area (bottom edge `bottom`); notes that do not fit continue
	// on the next page (or column). Returns the top of the notes, `bottom` without any.
	close(page: PageOutput, bottom: number, contentBottom: number): number {
		if (this.notes.length === 0) return bottom

		let y = Math.max(bottom - this.reserved, Math.min(contentBottom, bottom))
		const top = y
		const continued = (this.notes[0]?.from ?? 0) > 0
		const separator = continued ? this.continuationStory : this.separatorStory

		if (separator >= 0) page.append(this.laid(separator).output, this.left, y)
		else
			page.addGraphic(
				PageGraphic.NOTE_SEPARATOR,
				0,
				Layer.TEXT,
				this.left,
				y + points(6),
				continued ? this.width : Math.min(this.width, 192),
				1
			)

		y += continued ? this.continuationHeight : this.separatorHeight

		for (const note of this.notes) {
			const laid = this.laid(note.story)
			const room = bottom - y
			const rest = laid.height - note.from

			if (rest <= room + EPSILON) {
				page.append(laid.output, this.left, y - note.from, note.from, Infinity)
				y += rest
				this.cache.delete(note.story)

				continue
			}

			// Split at a line boundary: lines starting above the area's bottom stay.
			const cut = this.lineCut(laid.output, note.from, note.from + room)

			if (cut > note.from) page.append(laid.output, this.left, y - note.from, note.from, cut)

			this.carry.push({ story: note.story, from: cut, height: laid.height - cut })
			y = bottom
		}

		return top
	}

	// The top of the first line that does not end by `limit` (in the note's own coordinates).
	private lineCut(output: PageOutput, from: number, limit: number): number {
		const text = output.text
		let cut = limit

		for (let line = 0; line < text.lineCount; line++) {
			const top = text.lineTop(line)

			if (top < from - EPSILON) continue

			if (top + text.lineHeight(line) > limit + EPSILON) {
				cut = top

				break
			}
		}

		return Math.max(from, cut)
	}
}
