import type { Mark, MarkType, Node } from "prosemirror-model"
import { EditorState, type Transaction } from "prosemirror-state"
import { closeHistory, history, redo, redoDepth, undo, undoDepth } from "prosemirror-history"
import { AddMarkStep, RemoveMarkStep } from "prosemirror-transform"
import type { Story } from "../model/index"
import type { DocxGateReason } from "../package"
import { DocxDocBuilder, type StoryDoc } from "./build"
import {
	PARAGRAPH_STYLE_NAMES,
	isWriteCommand,
	type DocxAlignment,
	type DocxEditorCommand,
	type DocxEditorCommandType,
	type DocxParagraphStyle,
	DOCX_EDITOR_COMMANDS
} from "./commands"
import type { DocxEditContext } from "./context"
import { FormatResolver, RUN_TOGGLES, TOGGLE_ON, alignmentOf, type ListKind, type NewList, type RunToggle } from "./formats"
import { docxInvariants, markInternal, type ListNormalizer } from "./invariants"
import { conservativeProof, lockBlocks, proveBlocks, type BlockProof } from "./proof"
import { FORMATTING_MARKS, NO_SOURCE, docxSchema, inlineLevel, type IndentAttrs, type LockReason } from "./schema"

export type FormatState = "on" | "off" | "mixed"

// "textbox": text box stories are view-only until editing them also updates the VML fallback Word keeps
// of each text box (mc:Fallback), which would otherwise show the old text.
export type DocxReadOnlyReason = DocxGateReason | "host" | "textbox"

// Everything a toolbar needs, as plain data (it crosses a worker or WebView bridge as is).
export interface DocxEditorSnapshot {
	// Increases with every change of state.
	readonly version: number
	readonly readOnly: boolean
	readonly readOnlyReasons: readonly DocxReadOnlyReason[]
	// The document differs from the last saved (or opened) one.
	readonly dirty: boolean
	readonly canUndo: boolean
	readonly canRedo: boolean
	readonly selection: {
		readonly empty: boolean
		readonly from: number
		readonly to: number
		// The selection is (or is inside) a locked block, and why it is locked.
		readonly locked: LockReason | null
		readonly paragraphs: number
	}
	readonly formats: Readonly<Record<RunToggle, FormatState>>
	readonly paragraph: {
		// "custom" for any other style, "mixed" across paragraphs, null without paragraphs.
		readonly style: DocxParagraphStyle | "custom" | "mixed" | null
		// The style's display name when it is custom.
		readonly styleName: string | null
		readonly align: DocxAlignment | "other" | "mixed" | null
		readonly list: ListKind | "none" | "mixed" | null
		// List level (0-based) when every paragraph is at the same one.
		readonly level: number | null
	}
	readonly enabled: Readonly<Record<DocxEditorCommandType, boolean>>
}

export interface DocxEditorOptions {
	// The story to edit (default: the body).
	readonly story?: Story
	// The host's read-only mode (public links, versions, trash...). Gate 0 reasons make a session read-only too.
	readonly readOnly?: boolean
	// Gate 1 for the blocks of the story; null leaves it to the caller (see prove()).
	readonly proof?: BlockProof | null
	readonly normalizeList?: ListNormalizer
	// Undo steps kept (default 100, as ProseMirror's history).
	readonly historyDepth?: number
	// The story's editing document built ahead (DocxEditorSession.open builds it in slices).
	readonly built?: StoryDoc
}

const nodes = docxSchema.nodes
const marks = docxSchema.marks
const FORMATTING_TYPES: readonly MarkType[] = FORMATTING_MARKS.map(name => marks[name])
// What clearing direct formatting removes: everything but the language (Word's Ctrl+Space).
const CLEARED_TYPES: readonly MarkType[] = FORMATTING_TYPES.filter(type => type !== marks.lang && type !== marks.plain)
// Word's default indent step and tab stop: half an inch.
const INDENT_STEP = 720
const MAX_LIST_LEVEL = 8

interface RunPiece {
	readonly from: number
	readonly to: number
	readonly node: Node
	readonly paragraph: Node
	readonly barrier: boolean
}

// What a selection holds, additive over ranges so an extended selection updates by its difference.
// Characters (and run-level atoms) count by their size in the selection; paragraphs by their facts.
interface Tally {
	chars: number
	readonly on: number[]
	paragraphs: number
	outdentable: number
	readonly styles: Map<string, number>
	readonly aligns: Map<string, number>
	readonly lists: Map<string, number>
	readonly levels: Map<number, number>
}

interface ParagraphFacts {
	// A DocxParagraphStyle, or "custom:" and the style's name.
	readonly style: string
	readonly align: string
	readonly list: string
	readonly level: number
	readonly outdentable: boolean
}

function emptyTally(): Tally {
	return {
		chars: 0,
		on: [0, 0, 0, 0],
		paragraphs: 0,
		outdentable: 0,
		styles: new Map(),
		aligns: new Map(),
		lists: new Map(),
		levels: new Map()
	}
}

function bump<K>(map: Map<K, number>, key: K, by: number): void {
	const count = (map.get(key) ?? 0) + by

	if (count === 0) map.delete(key)
	else map.set(key, count)
}

function only<K>(map: ReadonlyMap<K, number>): K | undefined {
	return map.size === 1 ? map.keys().next().value : undefined
}

interface ParagraphAt {
	readonly node: Node
	readonly pos: number
}

// The headless editing session of one story: a ProseMirror state with the document invariants, undo
// history and the read-only gate, driven by serialisable commands. A view (when there is one) dispatches
// its transactions through dispatch(); a toolbar renders snapshot() and sends execute().
export class DocxEditorSession {
	readonly context: DocxEditContext
	readonly story: StoryDoc
	private readonly builder: DocxDocBuilder
	private readonly formats: FormatResolver
	private readonly hostReadOnly: boolean
	private readonly listeners = new Set<(snapshot: DocxEditorSnapshot) => void>()
	private readonly appliedListeners = new Set<(transactions: readonly Transaction[], state: EditorState) => void>()
	private current: EditorState
	private saved: Node
	private version = 0
	private cached: DocxEditorSnapshot | null = null
	private disposed = false
	private notifying = false
	private facts = new WeakMap<object, ParagraphFacts>()
	private blockTallies = new WeakMap<Node, Tally>()
	private selectionTally: { doc: Node; from: number; to: number; tally: Tally } | null = null
	private dirtyOf: { doc: Node; saved: Node; dirty: boolean } | null = null
	private savedFormsOf: { saved: Node; forms: Map<number, Node> } | null = null

	constructor(context: DocxEditContext, options: DocxEditorOptions = {}) {
		const story = options.story ?? context.model.body

		if (story === undefined) throw new RangeError("The document has no body")

		this.context = context
		this.hostReadOnly = options.readOnly ?? false
		this.builder = new DocxDocBuilder(context)
		this.story = options.built !== undefined && options.built.story === story ? options.built : this.builder.build(story)
		this.formats = new FormatResolver(context, this.story.part)
		this.current = EditorState.create({
			doc: this.story.doc,
			plugins: [
				history({ depth: options.historyDepth ?? 100 }),
				docxInvariants({
					readOnly: () => this.readOnly,
					...(options.normalizeList !== undefined ? { normalizeList: options.normalizeList } : {})
				})
			]
		})

		const proof = options.proof === undefined ? conservativeProof : options.proof

		if (proof !== null) this.prove(proof)

		this.saved = this.current.doc
	}

	// Opens a session building the story's document in slices (`pause` gives the host a turn whenever
	// `sliceMs` passed): no long task on a UI thread however long the document.
	static async open(
		context: DocxEditContext,
		options: DocxEditorOptions,
		slicing: { readonly pause: () => Promise<void>; readonly sliceMs: number; readonly now: () => number }
	): Promise<DocxEditorSession> {
		const story = options.story ?? context.model.body

		if (story === undefined) throw new RangeError("The document has no body")

		const built = await new DocxDocBuilder(context).buildSliced(story, slicing.pause, slicing.sliceMs, slicing.now)

		return new DocxEditorSession(context, { ...options, built })
	}

	get state(): EditorState {
		return this.current
	}

	get readOnly(): boolean {
		return this.hostReadOnly || this.context.gate.length > 0 || this.story.story.kind === "textbox"
	}

	// Ids of paragraph styles the document lacks that paragraphs use, with their built-in names: the writer
	// adds their definitions. Styles a command added that nothing uses any more are left out.
	get missingStyles(): ReadonlyMap<string, string> {
		if (this.context.newStyles.size === 0) return new Map()

		const used = this.usedParagraphValues("styleId")

		return new Map([...this.context.newStyles].filter(([id]) => used.has(id)))
	}

	// Lists commands started (negative numIds) that paragraphs use: the writer adds a w:num for each.
	get newLists(): ReadonlyMap<number, NewList> {
		if (this.formats.newLists.size === 0) return new Map()

		const used = this.usedParagraphValues("numId")

		return new Map([...this.formats.newLists].filter(([id]) => used.has(id)))
	}

	private usedParagraphValues(attr: string): Set<unknown> {
		const out = new Set<unknown>()

		this.current.doc.descendants(node => {
			if (node.type === nodes.paragraph) out.add(node.attrs[attr])

			return !node.isTextblock && node.type !== nodes.locked_block
		})

		return out
	}

	// Runs Gate 1 over [from, to) of the document and locks what fails; returns the number locked. Without
	// a proof at construction a host can call this in slices (mobile has no worker).
	prove(proof: BlockProof, from?: number, to?: number): number {
		const requests = proveBlocks(this.story, this.current.doc, proof, from, to)
		const tr = lockBlocks(this.current, this.story, this.builder, requests)

		if (tr !== null) this.apply(tr)

		return requests.length
	}

	// The source blocks of the last-saved document (paragraphs, tables, content controls) by source id, where
	// the id is that block's alone in it: what an edited block can go back to.
	private savedForms(): Map<number, Node> {
		const cached = this.savedFormsOf

		if (cached?.saved === this.saved) return cached.forms

		const forms = new Map<number, Node>()
		const repeated = new Set<number>()

		this.saved.descendants(node => {
			if (node.isInline) return false

			const src = node.attrs["src"]

			if (typeof src === "number" && src !== NO_SOURCE && !node.isInline) {
				if (forms.has(src)) repeated.add(src)
				else forms.set(src, node)
			}

			return !node.isTextblock && node.type !== nodes.locked_block
		})

		for (const src of repeated) forms.delete(src)

		this.savedFormsOf = { saved: this.saved, forms }

		return forms
	}

	// Blocks changed since the last save (or since the document opened) that can go back to how they were
	// saved: source paragraphs, and source tables or content controls changed other than in their source
	// paragraphs (rows, new paragraphs), each with its own form in the last-saved document. Positions in
	// state.doc, in document order. What a save that the writer refused falls back on.
	editedBlocks(): number[] {
		const forms = this.savedForms()
		const seen = new Map<number, number>()

		this.current.doc.descendants(node => {
			if (node.isInline) return false

			const src = node.attrs["src"]

			if (typeof src === "number" && src !== NO_SOURCE) seen.set(src, (seen.get(src) ?? 0) + 1)

			return !node.isTextblock && node.type !== nodes.locked_block
		})

		const revertible = (node: Node): boolean => {
			const src = node.attrs["src"]

			return typeof src === "number" && src !== NO_SOURCE && seen.get(src) === 1 && forms.has(src)
		}
		const out: number[] = []
		const visit = (parent: Node, offset: number): boolean => {
			let found = false

			parent.forEach((node, at) => {
				const pos = offset + at

				if (node.isInline || node.type === nodes.locked_block || forms.get(node.attrs["src"] as number) === node) return

				if (node.isTextblock) {
					if (revertible(node)) {
						out.push(pos)
						found = true
					}

					return
				}

				if (visit(node, pos + 1)) {
					found = true

					return
				}

				// Changed, but not in a source paragraph of its own: the container is what goes back.
				if ((node.type === nodes.table || node.type === nodes.sdt) && revertible(node)) {
					out.push(pos)
					found = true
				}
			})

			return found
		}

		visit(this.current.doc, 0)

		return out
	}

	// The step taking the blocks at `positions` of `state` back to their last-saved form: a block saved as the
	// file has it is locked (kept as its bytes, decision 5); one saved with edits goes back to that edited
	// form, which the writer accepted then.
	private revertStep(state: EditorState, positions: readonly number[]): Transaction | null {
		const forms = this.savedForms()
		const tr = state.tr

		for (const pos of [...positions].sort((a, b) => b - a)) {
			const node = tr.doc.nodeAt(pos)
			const form = node === null ? undefined : forms.get(node.attrs["src"] as number)

			if (node === null || form === undefined) continue

			const replacement =
				this.story.pristine.has(form) && (form.type === nodes.paragraph || form.type === nodes.table)
					? this.builder.lockedBlock(this.story, form, "proof")
					: form

			tr.replaceWith(pos, pos + node.nodeSize, replacement)
		}

		return tr.docChanged ? markInternal(tr) : null
	}

	// A read-only stand-in for this session whose document has the blocks at `positions` (from editedBlocks)
	// back at their last-saved form: what a save would write (pass it to the writer). The session is unchanged.
	withReverted(positions: readonly number[]): DocxEditorSession {
		const tr = this.revertStep(this.current, positions)
		const trial: DocxEditorSession = Object.create(this) as DocxEditorSession

		Object.defineProperty(trial, "current", { value: tr === null ? this.current : this.current.apply(tr) })

		return trial
	}

	// Takes these blocks (nodes of the document, found where they are now) back to their last-saved form, as
	// withReverted() wrote them: one undoable step, so the edits come back with an undo (and a later save takes
	// them back again). Blocks changed since are left. Returns the number taken back.
	revert(blocks: readonly Node[]): number {
		const wanted = new Set(blocks)
		const positions: number[] = []

		this.current.doc.descendants((node, pos) => {
			if (wanted.has(node)) positions.push(pos)

			return !node.isTextblock && wanted.size > positions.length
		})

		const tr = this.revertStep(this.current, positions)

		// An undo step of its own (never merged with the typing before it).
		if (tr === null || !this.apply(closeHistory(tr))) return 0

		return positions.length
	}

	// Applies a transaction from a view or a command. False when the invariants or the read-only gate
	// refused it.
	dispatch(tr: Transaction): boolean {
		return this.apply(tr)
	}

	execute(command: DocxEditorCommand): boolean {
		if (this.disposed || (this.readOnly && isWriteCommand(command.type))) return false

		const tr = this.transactionFor(command)

		if (tr === null) return false
		if (tr === true) return true

		return this.apply(tr)
	}

	canExecute(type: DocxEditorCommandType): boolean {
		return this.snapshot().enabled[type]
	}

	// Call after the document was written: it is no longer dirty. `doc` is the document as it was written
	// (state.doc when the save read it), so edits made while the file was stored stay unsaved.
	markSaved(doc: Node = this.current.doc): void {
		if (doc === this.saved) return

		this.saved = doc
		this.changed()
	}

	subscribe(listener: (snapshot: DocxEditorSnapshot) => void): () => void {
		this.listeners.add(listener)

		return () => {
			this.listeners.delete(listener)
		}
	}

	// Hears every change of state synchronously, with the transactions that made it (the one applied first,
	// then those the invariants appended): a paged view lays the document out again from them.
	onApplied(listener: (transactions: readonly Transaction[], state: EditorState) => void): () => void {
		this.appliedListeners.add(listener)

		return () => {
			this.appliedListeners.delete(listener)
		}
	}

	dispose(): void {
		this.disposed = true
		this.listeners.clear()
		this.appliedListeners.clear()
	}

	snapshot(): DocxEditorSnapshot {
		this.cached ??= this.computeSnapshot()

		return this.cached
	}

	// ---- internals --------------------------------------------------------------------------------------

	private apply(tr: Transaction): boolean {
		if (this.disposed) return false

		const result = this.current.applyTransaction(tr)

		if (result.transactions.length === 0) return false

		this.current = result.state

		for (const listener of this.appliedListeners) listener(result.transactions, result.state)

		this.changed()

		return true
	}

	// Listeners hear once per burst of changes (a microtask later), with the snapshot then current.
	private changed(): void {
		this.version++
		this.cached = null

		if (this.listeners.size === 0 || this.notifying) return

		this.notifying = true
		void Promise.resolve().then(() => {
			this.notifying = false

			if (this.disposed || this.listeners.size === 0) return

			const snapshot = this.snapshot()

			for (const listener of this.listeners) listener(snapshot)
		})
	}

	// The transaction a command makes; true when it ran without one (undo/redo), null when it cannot run.
	private transactionFor(command: DocxEditorCommand): Transaction | true | null {
		const state = this.current

		switch (command.type) {
			case "format.bold":
				return this.toggle("bold")
			case "format.italic":
				return this.toggle("italic")
			case "format.underline":
				return this.toggle("underline")
			case "format.strike":
				return this.toggle("strike")
			case "format.clear":
				return this.clearFormatting()
			case "style.set":
				return this.setStyle(command.style)
			case "align.start":
				return this.align("start")
			case "align.center":
				return this.align("center")
			case "align.end":
				return this.align("end")
			case "align.justify":
				return this.align("justify")
			case "list.bullet":
				return this.toggleList("bullet")
			case "list.numbered":
				return this.toggleList("numbered")
			case "list.indent":
				return this.indent(1)
			case "list.outdent":
				return this.indent(-1)
			case "history.undo":
				return undo(state, tr => this.apply(tr)) ? true : null
			case "history.redo":
				return redo(state, tr => this.apply(tr)) ? true : null
		}
	}

	private paragraphs(state: EditorState = this.current): ParagraphAt[] {
		const out: ParagraphAt[] = []
		const seen = new Set<number>()

		for (const range of state.selection.ranges) {
			state.doc.nodesBetween(range.$from.pos, range.$to.pos, (node, pos) => {
				if (node.type === nodes.paragraph) {
					if (!seen.has(pos)) {
						seen.add(pos)
						out.push({ node, pos })
					}

					return false
				}

				return !node.isInline && node.type !== nodes.locked_block
			})
		}

		return out
	}

	// Run-level content in the selection as [from, to) pieces with their paragraph. `barrier` is set on a
	// piece that follows content a formatting step must not touch (a marker, field or locked block).
	private runPieces(state: EditorState = this.current): RunPiece[] {
		const out: RunPiece[] = []
		let barrier = false

		for (const range of state.selection.ranges) {
			const from = range.$from.pos
			const to = range.$to.pos

			barrier = true
			state.doc.nodesBetween(from, to, (node, pos, parent) => {
				if (!node.isInline) {
					if (node.type !== nodes.locked_block) return true

					barrier = true

					return false
				}

				if (parent !== null && parent.type === nodes.paragraph && inlineLevel(node) === "run") {
					const start = Math.max(from, pos)
					const end = Math.min(to, pos + node.nodeSize)

					if (end > start) {
						out.push({ from: start, to: end, node, paragraph: parent, barrier })
						barrier = false
					}
				} else {
					barrier = true
				}

				return false
			})
		}

		return out
	}

	private toggleMark(toggle: RunToggle, on: boolean): Mark {
		switch (toggle) {
			case "bold":
			case "italic":
				return marks[toggle].create({ value: on, cs: on })
			case "underline":
				return on ? marks.underline.create({ style: "single" }) : marks.underline.create({ style: "none" })
			case "strike":
				return marks.strike.create({ value: on })
		}
	}

	private toggle(toggle: RunToggle): Transaction | null {
		const state = this.current
		const selection = state.selection
		const type = marks[toggle]

		if (selection.empty) {
			const paragraph = selection.$from.parent

			if (paragraph.type !== nodes.paragraph) return null

			const format = this.formats.paragraph(paragraph)
			const current = state.storedMarks ?? selection.$from.marks()
			const value = this.formats.toggle(toggle, format, current)
			const desired = !value.on
			const tr = state.tr.removeStoredMark(type)

			return desired === value.style ? tr : tr.addStoredMark(this.toggleMark(toggle, desired))
		}

		const pieces = this.runPieces(state)

		if (pieces.length === 0) return null

		const resolved = pieces.map(piece => this.formats.toggle(toggle, this.formats.paragraph(piece.paragraph), piece.node.marks))
		const desired = !resolved.every(value => value.on)
		const mark = this.toggleMark(toggle, desired)
		const tr = state.tr
		let group: { from: number; to: number; add: boolean; removed: Mark | undefined } | null = null
		const flush = (): void => {
			if (group === null) return

			if (group.removed !== undefined) tr.step(new RemoveMarkStep(group.from, group.to, group.removed))
			if (group.add) tr.step(new AddMarkStep(group.from, group.to, mark))

			group = null
		}

		// Direct formatting only where the styles do not already give the result. One step covers a run of
		// pieces across paragraphs (a step per piece would copy the document's block list each time) that
		// change alike: undo applies each step's inverse to its whole range, so a range holds only pieces
		// that had the same mark of this type (or none) and get the same change. A marker or field in
		// between ends the run.
		pieces.forEach((piece, index) => {
			const direct = desired !== (resolved[index]?.style ?? false)
			const existing = type.isInSet(piece.node.marks)
			const add = direct && existing?.eq(mark) !== true
			const removed = existing !== undefined && (add || !direct) ? existing : undefined

			if (!add && removed === undefined) {
				flush()

				return
			}

			if (group !== null && !piece.barrier && group.add === add && sameMark(group.removed, removed)) {
				group.to = piece.to
			} else {
				flush()
				group = { from: piece.from, to: piece.to, add, removed }
			}
		})
		flush()

		return tr.docChanged ? tr : null
	}

	// Word's Ctrl+Space: direct character formatting goes, the language stays, and the `plain` mark tells a
	// writer to ignore the run's source formatting except w:rtl and w:cs.
	private clearFormatting(): Transaction | null {
		const state = this.current
		const selection = state.selection
		const plain = marks.plain.create()

		if (selection.empty) {
			if (selection.$from.parent.type !== nodes.paragraph) return null

			const current = state.storedMarks ?? selection.$from.marks()

			return state.tr.setStoredMarks(plain.addToSet(current.filter(mark => !CLEARED_TYPES.includes(mark.type))))
		}

		const pieces = this.runPieces(state)

		if (pieces.length === 0) return null

		// Each step covers a run of pieces across paragraphs that all had the mark it removes (or all lack
		// the plain mark it adds): undo applies each step's inverse to its whole range.
		const found: Mark[] = []

		for (const piece of pieces) {
			for (const mark of piece.node.marks) {
				if (CLEARED_TYPES.includes(mark.type) && !found.some(known => known.eq(mark))) found.push(mark)
			}
		}

		const tr = state.tr
		const steps = (has: (piece: RunPiece) => boolean, step: (from: number, to: number) => void): void => {
			let start = -1
			let end = -1

			for (const piece of pieces) {
				if (start >= 0 && !piece.barrier && has(piece)) {
					end = piece.to

					continue
				}

				if (start >= 0) step(start, end)

				start = has(piece) ? piece.from : -1
				end = piece.to
			}

			if (start >= 0) step(start, end)
		}

		for (const mark of found) {
			steps(
				piece => mark.isInSet(piece.node.marks),
				(from, to) => tr.step(new RemoveMarkStep(from, to, mark))
			)
		}

		steps(
			piece => !plain.isInSet(piece.node.marks),
			(from, to) => tr.step(new AddMarkStep(from, to, plain))
		)

		return tr.docChanged ? tr : null
	}

	private setAttrs(tr: Transaction, at: ParagraphAt, attrs: Readonly<Record<string, unknown>>): void {
		for (const [name, value] of Object.entries(attrs)) {
			if (at.node.attrs[name] !== value) tr.setNodeAttribute(at.pos, name, value)
		}
	}

	// The id of a built-in paragraph style, adding it to the styles the writer must define when missing.
	private styleIdFor(style: DocxParagraphStyle): string | null {
		const styles = this.context.styles
		const name = PARAGRAPH_STYLE_NAMES[style]

		if (style === "normal") {
			const normal = styles.defaultOf("paragraph")

			return normal === undefined || normal.canonicalName === "normal" ? null : (styles.byName("normal", "paragraph")?.id ?? null)
		}

		const existing = styles.byName(name, "paragraph")

		if (existing !== undefined) return existing.id

		for (const [id, missing] of this.context.newStyles) if (missing === name) return id

		// Word's own id for the built-in style, made unique against the document's.
		const base = name.replace(/(^|\s)(\S)/g, (_match, _space: string, letter: string) => letter.toUpperCase())
		let id = base

		for (let n = 1; styles.byId(id) !== undefined || this.context.newStyles.has(id); n++) id = `${base}${n}`

		this.context.newStyles.set(id, name)
		this.facts = new WeakMap()
		this.blockTallies = new WeakMap()

		return id
	}

	private setStyle(style: DocxParagraphStyle): Transaction | null {
		const paragraphs = this.paragraphs()

		if (paragraphs.length === 0) return null

		const styleId = this.styleIdFor(style)
		const tr = this.current.tr

		for (const at of paragraphs) this.setAttrs(tr, at, { styleId })

		return tr
	}

	private align(alignment: DocxAlignment): Transaction | null {
		const paragraphs = this.paragraphs()

		if (paragraphs.length === 0) return null

		const value = alignment === "start" ? "left" : alignment === "end" ? "right" : alignment === "justify" ? "both" : "center"
		const tr = this.current.tr

		for (const at of paragraphs) {
			const fromStyle = alignmentOf(this.formats.styleParagraph(at.node.attrs["styleId"] as string | null).props.alignment)

			// No direct property where the style already aligns the paragraph so.
			this.setAttrs(tr, at, { align: fromStyle === alignment ? null : value })
		}

		return tr
	}

	private toggleList(kind: ListKind): Transaction | null {
		const state = this.current
		const paragraphs = this.paragraphs(state)

		if (paragraphs.length === 0) return null

		const lists = paragraphs.map(at => this.formats.listOf(at.node))
		const tr = state.tr

		if (lists.every(list => list?.kind === kind)) {
			for (const at of paragraphs) {
				// Numbering the style gives is switched off explicitly (numId 0); direct numbering is removed.
				const styled = this.formats.styleParagraph(at.node.attrs["styleId"] as string | null).props.numbering

				this.setAttrs(tr, at, styled !== undefined && styled.numId > 0 ? { numId: 0, ilvl: null } : { numId: null, ilvl: null })
			}

			return tr
		}

		// Continue the list next to the selection; else bullets take the document's first bullet list, and a
		// numbered list starts afresh on the definition of the document's first numbered list.
		const existing = this.formats.firstList(kind)
		const numId =
			this.adjacentList(state, paragraphs, kind) ?? (kind === "bullet" ? existing : undefined) ?? this.newList(kind, existing ?? null)

		paragraphs.forEach((at, index) => this.setAttrs(tr, at, { numId, ilvl: lists[index]?.ilvl ?? 0 }))

		return tr
	}

	// The list of the paragraph right before the selection (or right after it) when it is of `kind`: a
	// new item continues it.
	private adjacentList(state: EditorState, paragraphs: readonly ParagraphAt[], kind: ListKind): number | undefined {
		const first = paragraphs[0]
		const last = paragraphs[paragraphs.length - 1]

		if (first === undefined || last === undefined) return undefined

		const before = state.doc.resolve(first.pos).nodeBefore
		const after = state.doc.resolve(last.pos + last.node.nodeSize).nodeAfter

		for (const node of [before, after]) {
			if (node?.type !== nodes.paragraph) continue

			const list = this.formats.listOf(node)

			if (list?.kind === kind) return list.numId
		}

		return undefined
	}

	private newList(kind: ListKind, basedOn: number | null): number {
		// One numbering for every session of the document: lists of different stories never share an id.
		const id = -(this.formats.newLists.size + 1)

		this.formats.newLists.set(id, { kind, basedOn })
		this.facts = new WeakMap()
		this.blockTallies = new WeakMap()

		return id
	}

	private indent(direction: 1 | -1): Transaction | null {
		const paragraphs = this.paragraphs()

		if (paragraphs.length === 0) return null

		const tr = this.current.tr

		for (const at of paragraphs) {
			const list = this.formats.listOf(at.node)

			if (list !== undefined) {
				const ilvl = Math.max(0, Math.min(MAX_LIST_LEVEL, list.ilvl + direction))

				this.setAttrs(tr, at, { numId: list.numId, ilvl })

				continue
			}

			const indent = (at.node.attrs["indent"] as IndentAttrs | null) ?? {}
			const current = this.formats.paragraph(at.node).props.indent.start
			// Steps land on multiples of the step, as Word's indent buttons do.
			const next =
				direction > 0
					? (Math.floor(current / INDENT_STEP) + 1) * INDENT_STEP
					: Math.max(0, (Math.ceil(current / INDENT_STEP) - 1) * INDENT_STEP)

			if (next === current) continue

			const { startChars: _startChars, ...rest } = indent

			this.setAttrs(tr, at, { indent: { ...rest, start: next } })
		}

		return tr.docChanged ? tr : null
	}

	private paragraphFacts(node: Node): ParagraphFacts {
		let facts = this.facts.get(node.attrs)

		if (facts === undefined) {
			const format = this.formats.paragraph(node)
			const own = this.paragraphStyle(node, format.style?.canonicalName, format.style?.id)
			const list = this.formats.listOf(node)

			facts = {
				style: own.style === "custom" ? `custom:${own.name ?? ""}` : own.style,
				align: alignmentOf(format.props.alignment),
				list: list?.kind ?? "none",
				level: list?.ilvl ?? -1,
				outdentable: list !== undefined ? list.ilvl > 0 : format.props.indent.start > 0
			}
			this.facts.set(node.attrs, facts)
		}

		return facts
	}

	private addParagraph(tally: Tally, node: Node, sign: number): void {
		const facts = this.paragraphFacts(node)

		tally.paragraphs += sign
		tally.outdentable += facts.outdentable ? sign : 0
		bump(tally.styles, facts.style, sign)
		bump(tally.aligns, facts.align, sign)
		bump(tally.lists, facts.list, sign)

		if (facts.level >= 0) bump(tally.levels, facts.level, sign)
	}

	// The paragraph a position is inside of (after its start token), which a range starting there touches.
	private addContaining(tally: Tally, doc: Node, pos: number, sign: number): void {
		const $pos = doc.resolve(pos)

		for (let depth = $pos.depth; depth > 0; depth--) {
			if ($pos.node(depth).type === nodes.paragraph) {
				this.addParagraph(tally, $pos.node(depth), sign)

				return
			}
		}
	}

	private addTally(tally: Tally, add: Tally, sign: number): void {
		tally.chars += sign * add.chars
		tally.paragraphs += sign * add.paragraphs
		tally.outdentable += sign * add.outdentable

		for (let k = 0; k < tally.on.length; k++) tally.on[k] = (tally.on[k] ?? 0) + sign * (add.on[k] ?? 0)
		for (const [key, count] of add.styles) bump(tally.styles, key, sign * count)
		for (const [key, count] of add.aligns) bump(tally.aligns, key, sign * count)
		for (const [key, count] of add.lists) bump(tally.lists, key, sign * count)
		for (const [key, count] of add.levels) bump(tally.levels, key, sign * count)
	}

	private addInline(tally: Tally, paragraph: Node, node: Node, size: number, sign: number): void {
		const bits = this.formats.toggles(this.formats.paragraph(paragraph), node.marks)

		tally.chars += sign * size

		for (let k = 0; k < RUN_TOGGLES.length; k++) if ((bits & (TOGGLE_ON << k)) !== 0) tally.on[k] = (tally.on[k] ?? 0) + sign * size
	}

	// A block wholly in the selection. Tables and content controls keep their tally (by node identity: an
	// edit elsewhere leaves them as they were), so selecting everything again after an edit recounts only
	// what changed.
	private addBlock(tally: Tally, node: Node, sign: number): void {
		if (node.type === nodes.paragraph) {
			this.addParagraph(tally, node, sign)
			node.forEach(child => {
				if (inlineLevel(child) === "run") this.addInline(tally, node, child, child.nodeSize, sign)
			})

			return
		}

		if (node.type === nodes.locked_block || node.isInline) return

		if (node.type === nodes.table || node.type === nodes.sdt) {
			let own = this.blockTallies.get(node)

			if (own === undefined) {
				own = emptyTally()

				const target = own

				node.forEach(child => this.addBlock(target, child, 1))
				this.blockTallies.set(node, own)
			}

			this.addTally(tally, own, sign)

			return
		}

		node.forEach(child => this.addBlock(tally, child, sign))
	}

	// Run content clipped to [from, to) and the paragraphs starting there.
	private addRange(tally: Tally, doc: Node, from: number, to: number, sign: number): void {
		if (to <= from) return

		doc.nodesBetween(from, to, (node, pos, parent) => {
			if (!node.isInline) {
				if (pos >= from && pos + node.nodeSize <= to) {
					this.addBlock(tally, node, sign)

					return false
				}

				if (node.type === nodes.paragraph && pos >= from) this.addParagraph(tally, node, sign)

				return node.type !== nodes.locked_block
			}

			if (parent !== null && parent.type === nodes.paragraph && inlineLevel(node) === "run") {
				const size = Math.min(to, pos + node.nodeSize) - Math.max(from, pos)

				if (size > 0) this.addInline(tally, parent, node, size, sign)
			}

			return false
		})
	}

	// The tally of the selection: updated by the difference when only its ends moved on the same document
	// (extending a selection), else counted afresh.
	private tally(state: EditorState): Tally {
		const selection = state.selection
		const doc = state.doc

		if (selection.ranges.length !== 1) {
			const tally = emptyTally()

			for (const range of selection.ranges) {
				this.addContaining(tally, doc, range.$from.pos, 1)
				this.addRange(tally, doc, range.$from.pos, range.$to.pos, 1)
			}

			this.selectionTally = null

			return tally
		}

		const from = selection.from
		const to = selection.to
		const cached = this.selectionTally

		if (cached !== null && cached.doc === doc && Math.abs(from - cached.from) + Math.abs(to - cached.to) <= to - from) {
			const tally = cached.tally

			if (from !== cached.from) {
				this.addContaining(tally, doc, cached.from, -1)
				this.addContaining(tally, doc, from, 1)

				if (from < cached.from) this.addRange(tally, doc, from, cached.from, 1)
				else this.addRange(tally, doc, cached.from, from, -1)
			}

			if (to > cached.to) this.addRange(tally, doc, cached.to, to, 1)
			else if (to < cached.to) this.addRange(tally, doc, to, cached.to, -1)

			this.selectionTally = { doc, from, to, tally }

			return tally
		}

		const tally = emptyTally()

		this.addContaining(tally, doc, from, 1)
		this.addRange(tally, doc, from, to, 1)
		this.selectionTally = { doc, from, to, tally }

		return tally
	}

	private dirty(doc: Node): boolean {
		const cached = this.dirtyOf

		if (cached !== null && cached.doc === doc && cached.saved === this.saved) return cached.dirty

		const dirty = !this.saved.eq(doc)

		this.dirtyOf = { doc, saved: this.saved, dirty }

		return dirty
	}

	private computeSnapshot(): DocxEditorSnapshot {
		const state = this.current
		const selection = state.selection
		const readOnly = this.readOnly
		const tally = this.tally(state)
		const formats = { bold: "off", italic: "off", underline: "off", strike: "off" } as Record<RunToggle, FormatState>

		if (selection.empty) {
			const paragraph = selection.$from.parent

			if (paragraph.type === nodes.paragraph) {
				const bits = this.formats.toggles(this.formats.paragraph(paragraph), state.storedMarks ?? selection.$from.marks())

				RUN_TOGGLES.forEach((toggle, k) => {
					formats[toggle] = (bits & (TOGGLE_ON << k)) !== 0 ? "on" : "off"
				})
			}
		} else {
			RUN_TOGGLES.forEach((toggle, k) => {
				const on = tally.on[k] ?? 0

				formats[toggle] = tally.chars === 0 || on === 0 ? "off" : on === tally.chars ? "on" : "mixed"
			})
		}

		const styleKey = only(tally.styles)
		const style: DocxParagraphStyle | "custom" | "mixed" | null =
			tally.paragraphs === 0
				? null
				: styleKey === undefined
					? "mixed"
					: styleKey.startsWith("custom:")
						? "custom"
						: (styleKey as DocxParagraphStyle)
		const alignKey = only(tally.aligns)
		const listKey = only(tally.lists)
		const list = tally.paragraphs === 0 ? null : ((listKey ?? "mixed") as ListKind | "none" | "mixed")
		const lockedNode =
			selection.$from.nodeAfter?.type === nodes.locked_block && selection.to - selection.from <= 1 ? selection.$from.nodeAfter : null
		const writable = !readOnly && tally.paragraphs > 0
		const undoable = !readOnly && undoDepth(state) > 0
		const redoable = !readOnly && redoDepth(state) > 0
		const enabled = Object.fromEntries(DOCX_EDITOR_COMMANDS.map(command => [command.type, writable])) as Record<
			DocxEditorCommandType,
			boolean
		>

		enabled["history.undo"] = undoable
		enabled["history.redo"] = redoable
		enabled["list.outdent"] = writable && tally.outdentable > 0

		return {
			version: this.version,
			readOnly,
			readOnlyReasons: [
				...this.context.gate,
				...(this.hostReadOnly ? (["host"] as const) : []),
				...(this.story.story.kind === "textbox" ? (["textbox"] as const) : [])
			],
			dirty: this.dirty(state.doc),
			canUndo: undoable,
			canRedo: redoable,
			selection: {
				empty: selection.empty,
				from: selection.from,
				to: selection.to,
				locked: lockedNode === null ? null : (lockedNode.attrs["reason"] as LockReason),
				paragraphs: tally.paragraphs
			},
			formats,
			paragraph: {
				style,
				styleName: style === "custom" && styleKey !== undefined ? styleKey.slice("custom:".length) || null : null,
				align: tally.paragraphs === 0 ? null : ((alignKey ?? "mixed") as DocxAlignment | "other" | "mixed"),
				list,
				level: list !== null && list !== "none" && list !== "mixed" ? (only(tally.levels) ?? null) : null
			},
			enabled
		}
	}

	private paragraphStyle(
		node: Node,
		canonical: string | undefined,
		id: string | undefined
	): { style: DocxParagraphStyle | "custom"; name: string | null } {
		const styleId = node.attrs["styleId"] as string | null
		const missing = styleId === null ? undefined : this.context.newStyles.get(styleId)
		const name = missing ?? (id === this.context.styles.defaultOf("paragraph")?.id ? "normal" : canonical)

		for (const [style, builtIn] of Object.entries(PARAGRAPH_STYLE_NAMES))
			if (builtIn === name) return { style: style as DocxParagraphStyle, name: null }

		return { style: "custom", name: (id === undefined ? undefined : this.context.styles.byId(id)?.name) ?? null }
	}
}

function sameMark(a: Mark | undefined, b: Mark | undefined): boolean {
	return a === undefined ? b === undefined : b !== undefined && a.eq(b)
}
