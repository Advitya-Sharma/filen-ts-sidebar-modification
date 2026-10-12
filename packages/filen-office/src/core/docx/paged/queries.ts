import type { MarkType, Node } from "prosemirror-model"
import { TextSelection } from "prosemirror-state"
import { NO_SOURCE, docxSchema, type DocxEditorSession } from "../edit/index"
import { Item, RevisionView, type Story, type StoryPart } from "../model/index"
import { Ns } from "../model/namespaces"
import type { PagedEditor } from "./editor"

// Questions an app asks about the whole document now and then (a word count, the headings to navigate by):
// computed on demand, never per keystroke. A count is kept per document state (and per block, so counting
// again after an edit reads only the blocks it changed); the outline per document state.

const nodes = docxSchema.nodes
const marks = docxSchema.marks

export interface DocxTextCount {
	readonly words: number
	// Characters (code points) with and without spaces.
	readonly characters: number
	readonly charactersNoSpaces: number
	// Paragraphs holding something other than spaces.
	readonly paragraphs: number
}

export interface DocxWordCount extends DocxTextCount {
	// Pages laid out so far (all of them once the editor settled).
	readonly pages: number
}

export interface DocxWordCountOptions {
	// Word's "Include textboxes, footnotes and endnotes".
	readonly notesAndTextBoxes?: boolean
}

// Gives the host a turn between slices of a long count (no long task on a UI thread).
export interface DocxQuerySlicing {
	readonly pause: () => Promise<void>
	readonly sliceMs: number
	readonly now: () => number
}

// Word separates words by spaces; East Asian characters count one word each.
function isSpace(code: number): boolean {
	return code === 0x20 || code === 0x09 || code === 0x0a || code === 0x0d || (code >= 0x2000 && code <= 0x200a) || code === 0x3000
}

function isEastAsian(code: number): boolean {
	return (
		(code >= 0x3040 && code <= 0x30ff) ||
		(code >= 0x3400 && code <= 0x4dbf) ||
		(code >= 0x4e00 && code <= 0x9fff) ||
		(code >= 0xf900 && code <= 0xfaff) ||
		(code >= 0xff66 && code <= 0xff9f) ||
		(code >= 0x20000 && code <= 0x3ffff)
	)
}

class Counter {
	words = 0
	characters = 0
	charactersNoSpaces = 0
	paragraphs = 0
	private inWord = false
	private filled = false

	text(text: string): void {
		for (let k = 0; k < text.length; k++) {
			let code = text.charCodeAt(k)

			if (code >= 0xd800 && code <= 0xdbff && k + 1 < text.length) {
				const low = text.charCodeAt(k + 1)

				if (low >= 0xdc00 && low <= 0xdfff) {
					code = 0x10000 + ((code - 0xd800) << 10) + (low - 0xdc00)
					k++
				}
			}

			// An optional hyphen is invisible unless it breaks a line.
			if (code === 0xad) continue

			this.characters++

			if (isSpace(code)) {
				this.inWord = false

				continue
			}

			this.charactersNoSpaces++
			this.filled = true

			if (isEastAsian(code)) {
				this.words++
				this.inWord = false
			} else if (!this.inWord) {
				this.words++
				this.inWord = true
			}
		}
	}

	// A break between words that is no character (a line break, the end of a paragraph).
	gap(): void {
		this.inWord = false
	}

	endParagraph(): void {
		if (this.filled) this.paragraphs++

		this.filled = false
		this.inWord = false
	}

	add(count: DocxTextCount): void {
		this.words += count.words
		this.characters += count.characters
		this.charactersNoSpaces += count.charactersNoSpaces
		this.paragraphs += count.paragraphs
	}

	result(): DocxTextCount {
		return Object.freeze({
			words: this.words,
			characters: this.characters,
			charactersNoSpaces: this.charactersNoSpaces,
			paragraphs: this.paragraphs
		})
	}
}

function hiddenTypes(view: RevisionView): readonly MarkType[] {
	return view === RevisionView.Original ? [marks.insertion, marks.move_to] : [marks.deletion, marks.move_from]
}

function countInline(counter: Counter, node: Node): void {
	switch (node.type) {
		case nodes.text:
			counter.text(node.text ?? "")

			break
		case nodes.tab:
			counter.text("\t")

			break
		case nodes.break:
			counter.gap()

			break
		case nodes.hyphen:
			if (node.attrs["soft"] !== true) counter.text("‑")

			break
		case nodes.symbol: {
			const char = node.attrs["char"] as number

			// A w:sym without a valid w:char shows nothing (as layout draws it).
			if (char > 0 && char <= 0x10ffff) counter.text(String.fromCodePoint(char))

			break
		}
		case nodes.field:
			counter.text(node.attrs["result"] as string)

			break
		case nodes.run_object:
		case nodes.opaque_inline:
			counter.text(node.attrs["text"] as string)

			break
	}
}

// The paragraphs inside source item `item` (a block, or a story's container) as the file has them.
function countItem(counter: Counter, part: StoryPart, item: number, view: RevisionView): void {
	const end = part.end(item)
	const paragraph = (at: number): void => {
		for (const line of part.paragraphText(at, view).split("\n")) {
			counter.text(line)
			counter.gap()
		}

		counter.endParagraph()
	}

	if (part.kinds[item] === Item.Paragraph) {
		paragraph(item)

		return
	}

	for (let at = item + 1; at < end; at++) {
		if (part.kinds[at] !== Item.Paragraph) continue

		paragraph(at)
		at = part.end(at) - 1
	}
}

function countBlock(counter: Counter, part: StoryPart, node: Node, view: RevisionView, hidden: readonly MarkType[]): void {
	if (node.type === nodes.paragraph) {
		node.forEach(child => {
			if (!hidden.some(type => type.isInSet(child.marks) !== undefined)) countInline(counter, child)
		})
		counter.endParagraph()

		return
	}

	// Kept as its source bytes: counted from the file.
	if (node.type === nodes.locked_block) {
		const src = node.attrs["src"] as number

		if (src >= 0) countItem(counter, part, src, view)

		return
	}

	node.forEach(child => countBlock(counter, part, child, view, hidden))
}

// Counts per top-level block and view, kept while the block is (blocks an edit did not touch keep theirs).
const BLOCK_COUNTS: readonly WeakMap<Node, DocxTextCount>[] = [new WeakMap(), new WeakMap(), new WeakMap()]
const DOC_COUNTS: readonly WeakMap<Node, DocxTextCount>[] = [new WeakMap(), new WeakMap(), new WeakMap()]

function blockCount(part: StoryPart, node: Node, view: RevisionView): DocxTextCount {
	const cache = BLOCK_COUNTS[view] as WeakMap<Node, DocxTextCount>
	let count = cache.get(node)

	if (count === undefined) {
		const counter = new Counter()

		countBlock(counter, part, node, view, hiddenTypes(view))
		count = counter.result()
		cache.set(node, count)
	}

	return count
}

// The words, characters and paragraphs of a session's document as `view` shows it.
export async function countSession(
	session: DocxEditorSession,
	view: RevisionView = RevisionView.Final,
	slicing?: DocxQuerySlicing
): Promise<DocxTextCount> {
	const doc = session.state.doc
	const part = session.story.part
	const done = (DOC_COUNTS[view] as WeakMap<Node, DocxTextCount>).get(doc)

	if (done !== undefined) return done

	const counter = new Counter()
	let sliceStart = slicing?.now() ?? 0

	for (let k = 0; k < doc.childCount; k++) {
		counter.add(blockCount(part, doc.child(k), view))

		if (slicing !== undefined && slicing.now() - sliceStart >= slicing.sliceMs) {
			await slicing.pause()
			sliceStart = slicing.now()
		}
	}

	const count = counter.result()

	;(DOC_COUNTS[view] as WeakMap<Node, DocxTextCount>).set(doc, count)

	return count
}

function countModelStory(story: Story, editor: PagedEditor, view: RevisionView): DocxTextCount {
	const counter = new Counter()

	if (story.item >= 0) countItem(counter, editor.edits.model.partOf(story), story.item, view)

	return counter.result()
}

const NOTE_KINDS = new Set(["footnote", "endnote", "textbox"])

// Word's word count of the body (and, with the option, of text boxes, footnotes and endnotes), in the
// revision view the editor shows. Stories open in the editor count as edited; the others as the file has them.
export async function docxWordCount(
	editor: PagedEditor,
	options: DocxWordCountOptions = {},
	slicing?: DocxQuerySlicing
): Promise<DocxWordCount> {
	const view = editor.edits.view
	const counter = new Counter()

	counter.add(await countSession(editor.body.session, view, slicing))

	if (options.notesAndTextBoxes === true) {
		const open = new Map<Story, DocxEditorSession>()

		for (const session of editor.sessions) open.set(session.story.story, session)

		for (const story of editor.edits.model.stories) {
			if (story === undefined || !NOTE_KINDS.has(story.kind) || (story.noteType !== undefined && story.noteType !== "normal"))
				continue

			const session = open.get(story)

			counter.add(session === undefined ? countModelStory(story, editor, view) : await countSession(session, view, slicing))
		}
	}

	return { ...counter.result(), pages: editor.pageCount }
}

// ---- outline --------------------------------------------------------------------------------------------

export interface DocxOutlineEntry {
	// Outline level, 0 (Heading 1) to 8.
	readonly level: number
	readonly text: string
	// The heading paragraph's position in the body's document (`doc`).
	readonly pos: number
}

export interface DocxOutline {
	// The body's document the positions are in.
	readonly doc: Node
	readonly entries: readonly DocxOutlineEntry[]
}

// Characters of a heading an outline shows.
const MAX_HEADING = 255
const OUTLINES: readonly WeakMap<Node, DocxOutline>[] = [new WeakMap(), new WeakMap(), new WeakMap()]

// The body's headings (paragraphs with an outline level, from their style or direct formatting), in document
// order: Word's navigation pane.
export function docxOutline(editor: PagedEditor): DocxOutline {
	const session = editor.body.session
	const doc = session.state.doc
	const view = editor.edits.view
	const cache = OUTLINES[view] as WeakMap<Node, DocxOutline>
	let outline = cache.get(doc)

	if (outline !== undefined) return outline

	const formats = session.formats
	const part = session.story.part
	const hidden = hiddenTypes(view)
	const entries: DocxOutlineEntry[] = []
	const styleLevels = new Map<string | null, number | undefined>()
	// A paragraph's outline level: its style's, unless its own properties or a list may set one (then it is
	// resolved in full). Most paragraphs are resolved by their style alone.
	const levelOf = (node: Node): number | undefined => {
		const src = node.attrs["src"] as number
		const row = src === NO_SOURCE ? -1 : (part.aux[src] ?? -1)
		const pPr = row < 0 ? -1 : (part.paragraphs.pPr[row] ?? -1)

		if (node.attrs["numId"] !== null || (pPr >= 0 && part.reader.child(pPr, Ns.W, "outlineLvl") >= 0))
			return formats.paragraph(node).props.outlineLevel

		const styleId = node.attrs["styleId"] as string | null

		if (!styleLevels.has(styleId)) {
			const style = formats.styleParagraph(styleId)

			styleLevels.set(styleId, style.props.numbering === undefined ? style.props.outlineLevel : -1)
		}

		const level = styleLevels.get(styleId)

		return level === -1 ? formats.paragraph(node).props.outlineLevel : level
	}

	doc.descendants((node, pos) => {
		if (node.type !== nodes.paragraph) return node.type !== nodes.locked_block

		const level = levelOf(node)

		if (level !== undefined && level < 9) {
			let text = ""

			node.forEach(child => {
				if (child.isText && !hidden.some(type => type.isInSet(child.marks) !== undefined)) text += child.text ?? ""
			})

			text = text.trim()

			if (text !== "") entries.push(Object.freeze({ level, text: text.slice(0, MAX_HEADING), pos }))
		}

		return false
	})

	outline = Object.freeze({ doc, entries: Object.freeze(entries) })
	cache.set(doc, outline)

	return outline
}

// Puts the caret at the start of an outline entry's heading (in the body, which it enters). After edits the
// heading is found again by its level and text, the nearest to where it was; false when it is gone.
export function goToOutlineEntry(editor: PagedEditor, outline: DocxOutline, entry: DocxOutlineEntry): boolean {
	const session = editor.body.session
	const current = docxOutline(editor)
	let target: DocxOutlineEntry | undefined = current === outline ? entry : undefined

	if (target === undefined) {
		const index = outline.entries.indexOf(entry)
		let best = Number.POSITIVE_INFINITY

		current.entries.forEach((candidate, at) => {
			const distance = Math.abs(at - index)

			if (candidate.level === entry.level && candidate.text === entry.text && distance < best) {
				best = distance
				target = candidate
			}
		})
	}

	if (target === undefined) return false

	editor.exitStory()

	return editor.dispatch(session.state.tr.setSelection(TextSelection.create(current.doc, target.pos + 1)).scrollIntoView())
}
