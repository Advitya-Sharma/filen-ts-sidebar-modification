import { DOMParser, DOMSerializer, Fragment, Slice, type DOMOutputSpec, type Mark, type Node, type ResolvedPos } from "prosemirror-model"
import { EditorState, Plugin, type Transaction } from "prosemirror-state"
import { Mapping } from "prosemirror-transform"
import { EditorView, type NodeView, type NodeViewConstructor } from "prosemirror-view"
import { docxSchema, type DocxEditorSession, type MarkerList } from "../../../core/docx/edit/index"
import {
	besideNode,
	commandForKey,
	joinBackward,
	joinForward,
	lineBreak,
	selectAll,
	splitParagraph,
	type PagedEditor
} from "../../../core/docx/paged/index"
import { COMMENT_MOVE_META, freshSliceKeeping } from "../../../core/docx/edit/invariants/index"
import { docxComments } from "../../../core/docx/review/comments"

// The input sink of paged editing: a real ProseMirror view that holds focus and receives all input (keys,
// beforeinput, IME composition, clipboard, drag and drop) while the pages show the document. It is invisible
// and holds only the top-level blocks around the selection: each run of other blocks is one gap paragraph of
// the same size (never rendered), so positions are the document's, a transaction from the view applies to the
// session's state step for step, and a keystroke costs the same however long the document. It is moved so its
// own caret sits where the caret is painted, which is where IME candidate windows and magnifiers anchor.

export interface EditSinkOptions {
	// The accessible name of the editing area.
	readonly label: string
	readonly readOnly?: boolean
	// Cmd (Apple platforms) or Ctrl as the shortcut modifier.
	readonly apple?: boolean
	// Top-level blocks held on each side of the selection's ends (default 20).
	readonly window?: number
	// The painted caret on screen (client px) and the width of its column, or undefined when it is not drawn.
	readonly caretClientRect: () =>
		{ readonly left: number; readonly top: number; readonly height: number; readonly width?: number } | undefined
	// Page Up / Page Down.
	readonly pageMove?: (dir: 1 | -1, extend: boolean) => boolean
	// The app's find (Mod+F): the browser's would search the hidden text.
	readonly onFind?: () => void
}

const DEFAULT_WINDOW = 20
// The window is moved once the selection comes this close to its edge (so moving through the document moves
// it every few blocks, not on every block).
const WINDOW_MARGIN = 5
const IDENTITY = new Mapping()

// Gap paragraphs stand for runs of blocks outside the window: their text only gives them their size (a
// slice of one string, so a gap costs no copy of it).
const GAPS = new WeakSet<Node>()
let gapText = "-"

function gapParagraph(size: number): Node {
	const length = Math.max(0, size - 2)

	while (gapText.length < length) gapText += gapText

	const gap = nodes.paragraph.create(null, length === 0 ? null : docxSchema.text(gapText.slice(0, length)))

	GAPS.add(gap)

	return gap
}

function element<K extends keyof HTMLElementTagNameMap>(tag: K): HTMLElementTagNameMap[K] {
	return document.createElement(tag)
}

// Inline atoms as text a screen reader or a keyboard's text context can use.
function atom(text: (node: Node) => string, tag: "span" | "br" = "span"): NodeViewConstructor {
	return node => {
		const dom = element(tag)

		if (tag === "span") dom.textContent = text(node)

		dom.contentEditable = "false"

		return { dom, ignoreMutation: () => true }
	}
}

function symbolText(node: Node): string {
	const char = node.attrs["char"]

	return typeof char === "number" && char > 0 && char <= 0x10ffff ? String.fromCodePoint(char) : ""
}

function attrText(name: string): (node: Node) => string {
	return node => {
		const value = node.attrs[name]

		return typeof value === "string" ? value : ""
	}
}

const NODE_VIEWS: Record<string, NodeViewConstructor> = {
	paragraph: node => {
		if (GAPS.has(node)) {
			return {
				dom: element("p"),
				update: (next: Node) => GAPS.has(next),
				ignoreMutation: () => true
			} satisfies NodeView
		}

		const dom = element("p")

		// Right-to-left paragraphs: the browser's own caret motion and IME follow the text's direction.
		dom.dir = "auto"

		return { dom, contentDOM: dom, update: (next: Node) => next.type === node.type && !GAPS.has(next) } satisfies NodeView
	},
	table: () => {
		const dom = element("table")
		const contentDOM = element("tbody")

		dom.appendChild(contentDOM)

		return { dom, contentDOM }
	},
	table_row: () => {
		const dom = element("tr")

		return { dom, contentDOM: dom }
	},
	table_cell: () => {
		const dom = element("td")

		return { dom, contentDOM: dom }
	},
	sdt: () => {
		const dom = element("div")

		return { dom, contentDOM: dom }
	},
	locked_block: node => {
		const dom = element("div")

		dom.contentEditable = "false"
		dom.textContent = typeof node.attrs["text"] === "string" ? node.attrs["text"] : ""

		return { dom, ignoreMutation: () => true }
	},
	tab: atom(() => "\t"),
	break: atom(() => "", "br"),
	hyphen: atom(node => (node.attrs["soft"] === true ? "­" : "‑")),
	symbol: atom(symbolText),
	note_ref: atom(() => ""),
	comment_ref: atom(() => ""),
	drawing: atom(() => ""),
	run_object: atom(attrText("text")),
	field: atom(attrText("result")),
	marker: atom(() => ""),
	math: atom(() => ""),
	proof_err: atom(() => ""),
	rendered_break: atom(() => ""),
	opaque_inline: atom(attrText("text"))
}

const MARK_VIEWS = Object.fromEntries(Object.keys(docxSchema.marks).map(name => [name, () => ({ dom: element("span") })]))

// Copying writes this HTML (and the text): structure and basic formatting, never document-controlled URLs.
function clipboardSerializer(): DOMSerializer {
	const text =
		(value: (node: Node) => string) =>
		(node: Node): DOMOutputSpec => ["span", value(node)]
	const empty = (): DOMOutputSpec => ["span"]
	const nodeSpecs: Record<string, (node: Node) => DOMOutputSpec> = {
		paragraph: () => ["p", 0],
		table: () => ["table", ["tbody", 0]],
		table_row: () => ["tr", 0],
		table_cell: () => ["td", 0],
		sdt: () => ["div", 0],
		locked_block: node => ["p", attrText("text")(node)],
		tab: text(() => "\t"),
		break: () => ["br"],
		hyphen: text(node => (node.attrs["soft"] === true ? "" : "\u2011")),
		symbol: text(symbolText),
		note_ref: empty,
		comment_ref: empty,
		drawing: empty,
		run_object: text(attrText("text")),
		field: text(attrText("result")),
		marker: empty,
		math: empty,
		proof_err: empty,
		rendered_break: empty,
		opaque_inline: text(attrText("text"))
	}
	const toggled = (tag: string) => (mark: Mark) => (mark.attrs["value"] === false ? ["span", 0] : [tag, 0]) as DOMOutputSpec
	const markSpecs: Record<string, (mark: Mark, inline: boolean) => DOMOutputSpec> = Object.fromEntries(
		Object.keys(docxSchema.marks).map(name => [name, () => ["span", 0] as DOMOutputSpec])
	)

	markSpecs["bold"] = toggled("strong")
	markSpecs["italic"] = toggled("em")
	markSpecs["underline"] = () => ["u", 0]
	markSpecs["strike"] = toggled("s")

	return new DOMSerializer(nodeSpecs, markSpecs)
}

// Pasted HTML (after safeClipboardHtml) becomes paragraphs with basic formatting, breaks and tabs; the session's
// invariants make it new content.
function clipboardParser(): DOMParser {
	return new DOMParser(docxSchema, [
		{ tag: "p", node: "paragraph" },
		// Spaces kept; safeClipboardHtml made each line a block of its own, so no newline reaches a text node.
		{ tag: "pre", node: "paragraph", preserveWhitespace: true },
		{ tag: "br", node: "break" },
		{ tag: `span[${PASTE_TAB}]`, node: "tab" },
		{ tag: "b", mark: "bold", attrs: { value: true, cs: null } },
		{ tag: "i", mark: "italic", attrs: { value: true, cs: null } },
		{ tag: "u", mark: "underline" },
		{ tag: "s", mark: "strike", attrs: { value: true } }
	])
}

// Pasted HTML is rebuilt from its text and basic formatting before ProseMirror reads it: a flat list of <p>
// (or <pre>) blocks holding text, <br>, tabs and b/i/u/s only, without attributes; text is escaped. Content
// that is not text is dropped (scripts, styles, embedded and replaced content). A table becomes a paragraph
// per row, its cells separated by tabs and a cell's paragraphs by line breaks, as Word pastes a table as text
// (and as a copy's plain text has it): the editor cannot insert a table. Nesting is resolved here, never by
// the HTML parser or ProseMirror's clipboard reader, which would treat a leading cell as a table fragment and
// drop the rest. The source is parsed by DOMParser, which never runs scripts or loads anything, so what
// ProseMirror parses next references nothing; ProseMirror's own slice data (data-pm-slice) goes too.
const PASTE_TAB = "data-fo-tab"
const PASTE_DROPPED = new Set([
	"script",
	"style",
	"template",
	"noscript",
	"head",
	"title",
	"meta",
	"link",
	"base",
	"iframe",
	"frame",
	"frameset",
	"object",
	"embed",
	"applet",
	"img",
	"picture",
	"source",
	"track",
	"video",
	"audio",
	"canvas",
	"svg",
	"math",
	"map",
	"area",
	"form",
	"input",
	"button",
	"select",
	"textarea",
	"option",
	"dialog",
	"portal",
	"colgroup",
	"col"
])
// Elements that end the paragraph before them and start a new one.
const PASTE_BLOCKS = new Set([
	"p",
	"div",
	"h1",
	"h2",
	"h3",
	"h4",
	"h5",
	"h6",
	"li",
	"dt",
	"dd",
	"blockquote",
	"pre",
	"address",
	"article",
	"aside",
	"section",
	"header",
	"footer",
	"nav",
	"main",
	"figure",
	"figcaption",
	"center",
	"ul",
	"ol",
	"dl",
	"details",
	"summary",
	"hr",
	"caption",
	"table",
	"thead",
	"tbody",
	"tfoot",
	"tr",
	"td",
	"th"
])
// Blocks that are a paragraph of their own even when empty.
const PASTE_PARAGRAPHS = new Set(["p", "h1", "h2", "h3", "h4", "h5", "h6", "li", "dt", "dd", "pre"])
const ESCAPED: Readonly<Record<string, string>> = { "&": "&amp;", "<": "&lt;", ">": "&gt;" }
// Word's tab characters in its HTML (`<span style='mso-tab-count:2'>`).
const MSO_TABS = /mso-tab-count\s*:\s*(\d+)/i
const MARK_TAGS = ["b", "i", "u", "s"] as const

interface ParsedHtml {
	readonly body: ParsedNode | null
}

interface ParsedNode {
	readonly nodeType: number
	readonly nodeValue: string | null
	readonly localName?: string
	readonly childNodes: ArrayLike<ParsedNode>
	readonly style?: { readonly fontWeight?: string; readonly fontStyle?: string; readonly textDecoration?: string }
	getAttribute?(name: string): string | null
}

// Bold, italic, underline, strike: on, or off when unset.
type PasteMarks = readonly [boolean, boolean, boolean, boolean]

const NO_MARKS: PasteMarks = [false, false, false, false]

function pasteMarks(node: ParsedNode, tag: string, marks: PasteMarks): PasteMarks {
	let [bold, italic, underline, strike] = marks

	if (tag === "b" || tag === "strong") bold = true
	else if (tag === "i" || tag === "em") italic = true
	else if (tag === "u" || tag === "ins") underline = true
	else if (tag === "s" || tag === "del" || tag === "strike") strike = true

	const style = node.style

	if (style !== undefined) {
		const weight = style.fontWeight ?? ""

		// Google Docs wraps what it copies in <b style="font-weight:normal">.
		if (/^(bold|bolder|[6-9]00)$/.test(weight)) bold = true
		else if (/^(normal|lighter|[1-5]00)$/.test(weight)) bold = false

		const fontStyle = style.fontStyle ?? ""

		if (fontStyle === "italic" || fontStyle === "oblique") italic = true
		else if (fontStyle === "normal") italic = false

		const decoration = style.textDecoration ?? ""

		if (decoration.includes("underline")) underline = true
		if (decoration.includes("line-through")) strike = true
	}

	return bold === marks[0] && italic === marks[1] && underline === marks[2] && strike === marks[3]
		? marks
		: [bold, italic, underline, strike]
}

function marked(html: string, marks: PasteMarks): string {
	let open = ""
	let close = ""

	for (let k = 0; k < MARK_TAGS.length; k++) {
		if (!marks[k]) continue

		open += `<${MARK_TAGS[k] ?? ""}>`
		close = `</${MARK_TAGS[k] ?? ""}>` + close
	}

	return open + html + close
}

interface PasteBlock {
	readonly inline: string
	readonly pre: boolean
}

// Collects the paragraphs of a piece of HTML: inline content accumulates until a block boundary.
class PasteBlocks {
	readonly blocks: PasteBlock[] = []
	private inline = ""
	// The inline content holds more than collapsible white space.
	private content = false
	private pre = 0

	add(parent: ParsedNode, marks: PasteMarks): void {
		const children = parent.childNodes

		for (let k = 0; k < children.length; k++) {
			const child = children[k]

			if (child === undefined) continue

			if (child.nodeType === 3) {
				this.text(child.nodeValue ?? "", marks)

				continue
			}

			if (child.nodeType !== 1) continue

			const tag = child.localName ?? ""

			if (PASTE_DROPPED.has(tag)) continue

			const inner = pasteMarks(child, tag, marks)

			if (tag === "br") {
				this.inline += marked("<br>", inner)
				this.content = true
			} else if (tag === "table") {
				this.flush(false)
				this.table(child, inner)
			} else if (PASTE_BLOCKS.has(tag)) {
				this.block(child, tag, inner)
			} else {
				const tabs = MSO_TABS.exec(child.getAttribute?.("style") ?? "")

				if (tabs === null) {
					this.add(child, inner)
				} else {
					this.inline += marked(`<span ${PASTE_TAB}></span>`.repeat(Math.min(Number(tabs[1]), 16)), inner)
					this.content = true
				}
			}
		}
	}

	// What is left once the source is read.
	finish(): PasteBlock[] {
		this.flush(false)

		return this.blocks
	}

	private text(value: string, marks: PasteMarks): void {
		if (value.length === 0) return

		if (this.pre === 0 && !/[^ \t\r\n\f]/.test(value)) {
			// White space only: a separator, never a paragraph of its own.
			if (this.inline.length > 0) this.inline += " "

			return
		}

		if (this.pre > 0) {
			// Preformatted text keeps its lines (a paragraph each, as a plain-text paste makes them) and tabs.
			const lines = value.split(/\r\n?|\n/)

			lines.forEach((line, k) => {
				if (k > 0) this.flush(true)

				if (line.length === 0) return

				this.inline += marked(
					line
						.split("\t")
						.map(part => part.replace(/[&<>]/g, c => ESCAPED[c] ?? c))
						.join(`<span ${PASTE_TAB}></span>`),
					marks
				)
				this.content = true
			})

			return
		}

		this.inline += marked(
			value.replace(/[&<>]/g, c => ESCAPED[c] ?? c),
			marks
		)
		this.content = true
	}

	private block(node: ParsedNode, tag: string, marks: PasteMarks): void {
		this.flush(false)

		const before = this.blocks.length

		if (tag === "pre") this.pre++

		this.add(node, marks)

		const paragraph = PASTE_PARAGRAPHS.has(tag)

		this.flush(paragraph && this.blocks.length === before)

		if (tag === "pre") this.pre--
	}

	// A paragraph per row, cells separated by tabs.
	private table(node: ParsedNode, marks: PasteMarks): void {
		const rows: { row: ParsedNode; marks: PasteMarks }[] = []
		const findRows = (parent: ParsedNode, inherited: PasteMarks): void => {
			const children = parent.childNodes

			for (let k = 0; k < children.length; k++) {
				const child = children[k]
				const tag = child?.localName ?? ""

				if (child === undefined || child.nodeType !== 1 || PASTE_DROPPED.has(tag)) continue

				const inner = pasteMarks(child, tag, inherited)

				if (tag === "tr") rows.push({ row: child, marks: inner })
				else if (tag === "thead" || tag === "tbody" || tag === "tfoot") findRows(child, inner)
				else if (tag === "caption") this.block(child, tag, inner)
			}
		}

		findRows(node, marks)

		for (const { row, marks: rowMarks } of rows) {
			const cells: string[] = []
			const children = row.childNodes

			for (let k = 0; k < children.length; k++) {
				const cell = children[k]
				const tag = cell?.localName ?? ""

				if (cell === undefined || (tag !== "td" && tag !== "th")) continue

				const inner = new PasteBlocks()

				inner.pre = this.pre
				inner.add(cell, pasteMarks(cell, tag, rowMarks))
				cells.push(
					inner
						.finish()
						.map(block => block.inline)
						.join("<br>")
				)
			}

			this.blocks.push({ inline: cells.join(`<span ${PASTE_TAB}></span>`), pre: this.pre > 0 })
		}
	}

	// Ends the paragraph being collected: kept when it has content (a lone line break or no-break space is an
	// empty paragraph, as editors write one) or when `empty` asks for it.
	private flush(empty: boolean): void {
		if (this.content || empty) {
			// A line break that ends a block adds no line (as HTML shows it).
			const inline = this.inline.replace(/(?:<[bius]>)*<br>(?:<\/[bius]>)*\s*$/, "")

			this.blocks.push({ inline: /^\s*$/.test(inline) ? "" : inline, pre: this.pre > 0 })
		}

		this.inline = ""
		this.content = false
	}
}

export function safeClipboardHtml(html: string): string {
	const Parser: unknown = Reflect.get(globalThis, "DOMParser")

	if (typeof Parser !== "function") return ""

	const parsed = new (Parser as new () => { parseFromString(source: string, type: string): ParsedHtml })().parseFromString(
		html,
		"text/html"
	)

	if (parsed.body === null) return ""

	const blocks = new PasteBlocks()

	blocks.add(parsed.body, NO_MARKS)

	return blocks
		.finish()
		.map(block => (block.pre ? `<pre>${block.inline}</pre>` : `<p>${block.inline}</p>`))
		.join("")
}

const nodes = docxSchema.nodes
// Leaves as plain text (what a copy writes and assistive technology reads).
const PLAIN_LEAVES: Record<string, (node: Node) => string> = {
	tab: () => "\t",
	break: () => "\n",
	hyphen: node => (node.attrs["soft"] === true ? "" : "\u2011"),
	symbol: symbolText,
	run_object: attrText("text"),
	field: attrText("result"),
	opaque_inline: attrText("text")
}
const plainLeaf = (node: Node): string => PLAIN_LEAVES[node.type.name]?.(node) ?? ""

// [from, to) of `doc` as plain text, as Word copies it: paragraphs on lines of their own (empty ones
// kept), a table's cells separated by tabs and its rows by line breaks, list numbers as text.
export function plainText(doc: Node, from: number, to: number, marker: (pos: number) => string): string {
	const blocks = (parent: Node, offset: number, out: string[]): string[] => {
		parent.forEach((child, at) => {
			const pos = offset + at
			const end = pos + child.nodeSize

			if (end <= from || pos >= to) return

			if (child.isTextblock) {
				const a = Math.max(from, pos + 1) - pos - 1
				const b = Math.max(a, Math.min(to, end - 1) - pos - 1)
				const prefix = from <= pos + 1 && to > pos + 1 ? marker(pos + 1) : ""

				out.push(prefix + child.textBetween(a, b, undefined, plainLeaf))
			} else if (child.type === nodes.table) {
				child.forEach((row, rowAt) => {
					const rowPos = pos + 1 + rowAt

					if (rowPos + row.nodeSize <= from || rowPos >= to) return

					const cells: string[] = []

					row.forEach((cell, cellAt) => {
						const cellPos = rowPos + 1 + cellAt

						if (cellPos + cell.nodeSize > from && cellPos < to) cells.push(blocks(cell, cellPos + 1, []).join("\n"))
					})
					out.push(cells.join("\t"))
				})
			} else if (child.type === nodes.locked_block) {
				out.push(attrText("text")(child))
			} else {
				blocks(child, pos + 1, out)
			}
		})

		return out
	}

	return blocks(doc, 0, []).join("\n")
}

// Without the control characters XML cannot hold (tabs and vertical tabs become nodes).
function withoutControls(text: string): string {
	let out = ""
	let last = 0

	for (let k = 0; k < text.length; k++) {
		const c = text.charCodeAt(k)

		if (c < 0x20 && c !== 0x09 && c !== 0x0b) {
			out += text.slice(last, k)
			last = k + 1
		}
	}

	return last === 0 ? text : out + text.slice(last)
}

// Pasted plain text: one paragraph per line (empty lines kept), tabs and vertical tabs (Word's line breaks)
// as nodes, formatted as the text at the insertion point.
export function plainSlice(text: string, $context: ResolvedPos): Slice {
	const marks = $context.marks()
	const paragraphs = text.split(/\r\n?|\n/).map(line => {
		const inline: Node[] = []
		const parts = withoutControls(line).split(/([\t\v])/)

		for (const part of parts) {
			if (part === "\t") inline.push(nodes.tab.create(null, null, marks))
			else if (part === "\v") inline.push(nodes.break.create({ type: "textWrapping" }, null, marks))
			else if (part.length > 0) inline.push(docxSchema.text(part, marks))
		}

		return nodes.paragraph.create(null, inline)
	})

	return new Slice(Fragment.from(paragraphs), 1, 1)
}

const COMMENT_MARKERS: ReadonlySet<unknown> = new Set(["commentRangeStart", "commentRangeEnd"])

// Ids of the comments anchored in `fragment` (ranges, references, marker lists of blocks).
function commentIds(fragment: Fragment): string[] {
	const out = new Set<string>()
	const list = (value: unknown): void => {
		if (!Array.isArray(value)) return

		for (const marker of value as MarkerList)
			if (COMMENT_MARKERS.has(marker.kind) || marker.kind === "commentReference") if (marker.id !== null) out.add(marker.id)
	}

	fragment.descendants(node => {
		const id: unknown = node.attrs["id"]

		if (
			node.type === docxSchema.nodes.comment_ref ||
			(node.type === docxSchema.nodes.marker && COMMENT_MARKERS.has(node.attrs["kind"]))
		) {
			if (typeof id === "string") out.add(id)
		}

		list(node.attrs["before"])
		list(node.attrs["end"])

		return true
	})

	return [...out]
}

export class EditSink {
	readonly element: HTMLDivElement
	readonly view: EditorView
	private readonly editor: PagedEditor
	private readonly options: EditSinkOptions
	private readonly windowSize: number
	// The session's state the view's was made from, the top-level blocks the view holds ([from, to) index
	// ranges, ascending) and its gap paragraphs (kept while their size does not change).
	private real: EditorState
	private windows: number[] = []
	private gaps: Node[] = []
	private offsetX = 0
	private offsetY = 0
	private compositionStart = -1
	// The text being composed (the last compositionupdate's).
	private compositionText = ""
	private dispatching = false
	private readOnlyOption: boolean
	private readOnlyShown: boolean
	// The input view's font size and width as last set (px, client space).
	private lineHeight = 0
	private width = 0
	private readonly stopListening: () => void
	// What the last cut took out with comment anchors in it: pasted back as text of the same clipboard, it moves
	// the comments (as in Word) rather than leaving them out.
	private cut:
		{ readonly session: DocxEditorSession; readonly text: string; readonly slice: Slice; readonly ids: readonly string[] } | undefined

	constructor(editor: PagedEditor, parent: HTMLElement, options: EditSinkOptions) {
		this.editor = editor
		this.options = options
		this.windowSize = Math.max(1, options.window ?? DEFAULT_WINDOW)
		this.element = element("div")
		this.element.className = "fo-docx-sink"
		parent.appendChild(this.element)

		this.readOnlyOption = options.readOnly === true
		this.readOnlyShown = this.readOnly
		this.real = editor.state
		this.view = new EditorView(this.element, {
			state: this.viewState(editor.state),
			editable: () => !this.readOnly,
			dispatchTransaction: tr => this.dispatch(tr),
			// The session's plugins act on its own state; what stateless ones do in a view (paste, drop) acts here.
			// A stateful plugin's props read its state from the view's, which has none (history's undo is below).
			plugins: editor.state.plugins.flatMap(plugin =>
				plugin.spec.props === undefined || plugin.spec.state !== undefined ? [] : [new Plugin({ props: plugin.spec.props })]
			),
			nodeViews: NODE_VIEWS,
			markViews: MARK_VIEWS,
			clipboardSerializer: clipboardSerializer(),
			clipboardParser: clipboardParser(),
			transformPastedHTML: safeClipboardHtml,
			clipboardTextSerializer: () => {
				const state = this.editor.state
				const { from, to } = state.selection

				return plainText(state.doc, from, to, pos => this.editor.geometry.markerText(pos))
			},
			clipboardTextParser: (text, $context) => plainSlice(text, $context),
			handlePaste: (_view, event) => this.pasteCut(event.clipboardData?.getData("text/plain")),
			// The pages scroll to the painted caret; the sink never scrolls anything.
			handleScrollToSelection: () => true,
			handleKeyDown: (view, event) => this.keyDown(view, event),
			handleDOMEvents: {
				// Read-only, the view is not editable and ProseMirror's own key handling is off: selection keys only.
				keydown: (view, event) => (view.editable ? false : this.readOnlyKeyDown(view, event)),
				// Undo and redo from outside the keyboard (Edit menu, shake, a soft keyboard's key): the editor's,
				// in its order across stories.
				beforeinput: (_view, event) => this.historyInput(event),
				// The document's selection, not the view's (whose gaps hold no content).
				copy: (_view, event) => this.copyEvent(event, false),
				cut: (_view, event) => this.copyEvent(event, true),
				compositionstart: view => {
					this.compositionStart = view.state.selection.from
					this.compositionText = ""

					return false
				},
				compositionupdate: (_view, event) => {
					this.compositionText = event.data

					return false
				},
				compositionend: () => {
					this.compositionStart = -1
					this.compositionText = ""
					this.editor.setComposition(undefined)

					return false
				}
			},
			attributes: () => ({
				role: "textbox",
				"aria-multiline": "true",
				"aria-label": options.label,
				spellcheck: "false",
				translate: "no",
				// Not editable but still focused (selection, copy).
				...(this.readOnly ? { "aria-readonly": "true", inputmode: "none", tabindex: "0" } : {})
			})
		})

		// Commands from a toolbar change the session without the view, and editing can move to another story
		// (a header): the view follows.
		this.stopListening = editor.subscribe(event => {
			if (event.type === "pages" || this.dispatching) return

			if (this.real !== editor.state) {
				if (event.type === "story") this.windows = []

				this.sync()
			}

			this.refreshReadOnly()
		})
	}

	// Read-only: the host's option or the story's session (a gate, a text box).
	get readOnly(): boolean {
		return this.readOnlyOption || this.editor.session.readOnly
	}

	setReadOnly(readOnly: boolean): void {
		this.readOnlyOption = readOnly
		this.refreshReadOnly()
	}

	// The view's editability and attributes follow read-only.
	private refreshReadOnly(): void {
		const readOnly = this.readOnly

		if (readOnly === this.readOnlyShown) return

		this.readOnlyShown = readOnly
		this.view.setProps({})
	}

	destroy(): void {
		this.stopListening()
		this.view.destroy()
		this.element.remove()
		// The gaps' text grows to a document's size: a closed document does not keep it.
		gapText = "-"
	}

	focus(): void {
		const view = this.view

		if (!view.editable) {
			// ProseMirror only syncs a read-only view's selection once the document's selection is inside it.
			const selection = view.dom.ownerDocument.getSelection()

			view.dom.focus({ preventScroll: true })

			if (selection !== null && !view.dom.contains(selection.anchorNode)) selection.collapse(view.dom, 0)
		}

		view.focus()
	}

	get focused(): boolean {
		return this.view.hasFocus()
	}

	// The selection as a copy writes it (the same HTML and text as a copy event), or undefined when it is
	// empty. For an app's own Cut and Copy commands (menus, touch), which write the system clipboard itself.
	clipboardContent(): { readonly html: string; readonly text: string } | undefined {
		const state = this.editor.state

		if (state.selection.empty) return undefined

		const { dom, text } = this.view.serializeForClipboard(state.selection.content())

		return { html: dom.innerHTML, text }
	}

	// Deletes the selection (the second half of an app's Cut); false when read-only or nothing is selected.
	deleteSelection(): boolean {
		if (this.readOnly || this.editor.state.selection.empty) return false

		this.rememberCut()

		return this.run(this.editor.state.tr.deleteSelection())
	}

	// Pastes what an app read from the system clipboard, through the same parsers as a paste event: HTML
	// keeps paragraphs and basic formatting only (parsed in an inert document; nothing it references loads),
	// text is taken as typed. False when read-only or there is nothing to paste.
	paste(content: { readonly html?: string; readonly text?: string }): boolean {
		if (this.readOnly) return false
		if (this.pasteCut(content.text)) return true
		if (content.html !== undefined && content.html.length > 0) return this.view.pasteHTML(content.html)
		if (content.text !== undefined && content.text.length > 0) return this.view.pasteText(content.text)

		return false
	}

	selectAll(): boolean {
		return this.run(selectAll(this.editor.state))
	}

	get composing(): boolean {
		return this.view.composing
	}

	// Sizes the sink's text like the painted caret's line (IME windows take the caret's height) and its
	// column, and moves it so its own caret lies on the painted one.
	align(): void {
		const target = this.options.caretClientRect()

		if (target === undefined) return

		const style = this.element.style
		const lineHeight = Math.round(target.height * 10) / 10
		const width = target.width === undefined ? 0 : Math.max(1, Math.round(target.width))

		if (lineHeight > 0 && lineHeight !== this.lineHeight) {
			this.lineHeight = lineHeight
			style.fontSize = `${String(Math.round((lineHeight / 1.2) * 10) / 10)}px`
			style.lineHeight = `${String(lineHeight)}px`
			style.height = `${String(lineHeight)}px`
		}

		if (width > 0 && width !== this.width) {
			this.width = width
			style.width = `${String(width)}px`
		}

		let coords: { left: number; top: number }

		try {
			coords = this.view.coordsAtPos(this.view.state.selection.head)
		} catch {
			return
		}

		const dx = target.left - coords.left
		const dy = target.top - coords.top

		if (Math.abs(dx) < 0.5 && Math.abs(dy) < 0.5) return

		this.offsetX += dx
		this.offsetY += dy
		this.element.style.transform = `translate(${Math.round(this.offsetX * 10) / 10}px,${Math.round(this.offsetY * 10) / 10}px)`
	}

	// A transaction of the view: applied to the session's state (positions are the same), and the view brought
	// to the result (whatever the browser changed that the session did not take is put back).
	private dispatch(viewTr: Transaction): void {
		const tr = viewTr.docChanged && this.readOnly ? undefined : this.sessionTransaction(viewTr)

		if (tr === undefined) {
			this.sync(true)

			return
		}

		this.apply(tr, true)
	}

	// `tr` (on the view's state) as a transaction on the session's: its steps, selection, stored marks, time
	// and metadata; undefined when a step does not apply there.
	private sessionTransaction(tr: Transaction): Transaction | undefined {
		const real = this.editor.state

		// The view holds the whole document: the transaction is the session's as it is.
		if (this.view.state.doc === real.doc) return tr

		const out = real.tr

		for (const step of tr.steps) if (out.maybeStep(step).failed !== null) return undefined

		if (tr.selectionSet) out.setSelection(tr.selection.map(out.doc, IDENTITY))
		if (tr.storedMarksSet) out.setStoredMarks(tr.storedMarks)
		if (tr.scrolledIntoView) out.scrollIntoView()

		const meta: unknown = Reflect.get(tr, "meta")

		if (typeof meta === "object" && meta !== null) for (const key in meta) out.setMeta(key, Reflect.get(meta, key))

		return out.setTime(tr.time)
	}

	// Applies a transaction on the session's state; `fromView`: it came from the view, which is brought up to
	// date even when the session refused it.
	private apply(tr: Transaction, fromView = false): void {
		if (tr.docChanged && this.readOnly) {
			this.sync(fromView)

			return
		}

		if (this.compositionStart >= 0 && tr.docChanged) this.compositionStart = tr.mapping.map(this.compositionStart, -1)

		this.dispatching = true

		try {
			this.editor.dispatch(tr)
		} finally {
			this.dispatching = false
		}

		this.sync(fromView)

		if (this.view.composing && this.compositionStart >= 0) {
			const state = this.editor.state
			const from = compositionFrom(state.doc, state.selection.head, this.compositionStart, this.compositionText)

			this.compositionStart = from
			this.editor.setComposition({
				from,
				to: this.compositionText.length > 0 ? from + this.compositionText.length : Math.max(from, state.selection.head)
			})
		}
	}

	// A command's transaction on the session's state.
	private run(tr: Transaction | null | false): boolean {
		if (tr === null) return false
		if (tr !== false) this.apply(tr)

		return true
	}

	// beforeinput historyUndo / historyRedo: the editor undoes or redoes in its order across stories (the
	// browser's own never runs).
	private historyInput(event: InputEvent): boolean {
		const type = event.inputType === "historyUndo" ? "history.undo" : event.inputType === "historyRedo" ? "history.redo" : undefined

		if (type === undefined) return false

		event.preventDefault()

		if (!this.readOnly) void this.editor.execute({ type })

		return true
	}

	// A copy or cut writes the document's selection, as ProseMirror's own would write the view's.
	private copyEvent(event: ClipboardEvent, cut: boolean): boolean {
		const data = event.clipboardData
		const content = data === null ? undefined : this.clipboardContent()

		if (data === null || content === undefined) return false

		event.preventDefault()
		data.clearData()
		data.setData("text/html", content.html)
		data.setData("text/plain", content.text)
		this.cut = undefined

		if (cut && !this.readOnly) {
			this.rememberCut(content.text)
			this.apply(this.editor.state.tr.deleteSelection().scrollIntoView().setMeta("uiEvent", "cut"))
		}

		return true
	}

	// Keeps the selection as it is about to be cut when it holds comment anchors (see pasteCut).
	private rememberCut(text?: string): void {
		const slice = this.editor.state.selection.content()
		const ids = commentIds(slice.content)

		this.cut =
			ids.length === 0
				? undefined
				: { session: this.editor.session, text: text ?? this.view.serializeForClipboard(slice).text, slice, ids }
	}

	// Pastes the last cut as it was, comment anchors included, when `text` is what it wrote and its comments are
	// anchored nowhere now (the cut took them out, untracked, and no paste put them back): the comments move, also
	// into another story (as in Word). Once only; a copy, or anything else, pastes as usual.
	private pasteCut(text: string | undefined): boolean {
		const cut = this.cut

		if (cut === undefined || text === undefined || this.readOnly) return false

		this.cut = undefined

		const comments = docxComments(this.editor)

		if (text.replace(/\r\n?/g, "\n") !== cut.text.replace(/\r\n?/g, "\n") || cut.ids.some(id => comments.comment(id) !== undefined))
			return false

		// Into another part, the cut's source identities are not that part's: new content, the anchors kept.
		const slice =
			this.editor.session.story.part === cut.session.story.part
				? cut.slice
				: freshSliceKeeping(
						cut.slice,
						new Set(
							cut.ids.flatMap(id => ["commentRangeStart", "commentRangeEnd", "commentReference"].map(kind => `${kind}:${id}`))
						)
					)

		return this.run(
			this.editor.state.tr.replaceSelection(slice).scrollIntoView().setMeta("uiEvent", "paste").setMeta(COMMENT_MOVE_META, cut.ids)
		)
	}

	// Brings the view to the session's state; `force`: also when that did not change (the view's DOM may
	// differ from its state after a refused change).
	private sync(force = false): void {
		const real = this.editor.state

		if (real === this.real && !force) return

		const state = real === this.real ? this.view.state : this.viewState(real)

		this.real = real
		this.view.updateState(state)
	}

	// The view's state for the session's `real`: its document's top-level blocks around the selection's ends,
	// each run of other blocks one gap paragraph of the run's size, and the selection there.
	private viewState(real: EditorState): EditorState {
		const doc = real.doc
		const count = doc.childCount
		const story = this.editor.story
		const starts = story.document === doc ? story.starts : undefined
		const indexOf = (pos: number): number => (starts === undefined ? doc.resolve(pos).index(0) : story.entryAt(pos))
		const { from, to } = real.selection
		const a = Math.max(0, Math.min(count - 1, indexOf(from)))
		const b = Math.max(a, Math.min(count - 1, indexOf(to)))
		const n = this.windowSize
		const margin = Math.min(WINDOW_MARGIN, n >> 1)
		const [lo = 0, hi = 0] = this.windows
		// The window stays while the selection is well inside it (or it reaches the document's edge).
		const kept =
			this.windows.length === 2 &&
			hi <= count &&
			a >= lo &&
			b < hi &&
			(lo === 0 || a - lo >= margin) &&
			(hi === count || hi - 1 - b >= margin)

		if (!kept) {
			// A selection spanning many blocks: the blocks around each end.
			this.windows =
				b - a > 2 * n
					? [Math.max(0, a - n), Math.min(count, a + n + 1), Math.max(0, b - n), Math.min(count, b + n + 1)]
					: [Math.max(0, a - n), Math.min(count, b + n + 1)]
		}

		const windows = this.windows

		if (windows.length === 2 && windows[0] === 0 && windows[1] === count) {
			this.gaps = []

			return EditorState.create({ doc, selection: real.selection, storedMarks: real.storedMarks })
		}

		const children: Node[] = []
		const gaps: Node[] = []
		const positionOfChild = (index: number): number =>
			index >= count ? doc.content.size : starts === undefined ? positionOf(doc, index) : (starts[index] ?? 0)
		const gap = (start: number, end: number): void => {
			const size = positionOfChild(end) - positionOfChild(start)
			const node = this.gaps.find(other => other.nodeSize === size && !gaps.includes(other)) ?? gapParagraph(size)

			gaps.push(node)
			children.push(node)
		}
		let at = 0

		for (let w = 0; w + 1 < windows.length; w += 2) {
			const start = Math.max(at, windows[w] ?? 0)
			const end = Math.min(count, windows[w + 1] ?? 0)

			if (start > at) gap(at, start)

			for (let k = start; k < end; k++) children.push(doc.child(k))

			at = Math.max(at, end)
		}

		if (at < count) gap(at, count)

		this.gaps = gaps

		const viewDoc = doc.copy(Fragment.from(children))

		return EditorState.create({ doc: viewDoc, selection: real.selection.map(viewDoc, IDENTITY), storedMarks: real.storedMarks })
	}

	// Keys of a read-only view: what moves the selection; character and word steps through the document's
	// selection (the browser moves no caret in content that is not editable).
	private readOnlyKeyDown(view: EditorView, event: KeyboardEvent): boolean {
		if (this.keyDown(view, event, true)) {
			event.preventDefault()

			return true
		}

		// Some browsers go back in history on Backspace outside editable content.
		if (event.key === "Backspace") {
			event.preventDefault()

			return true
		}

		if (event.key !== "ArrowLeft" && event.key !== "ArrowRight") return false

		const apple = this.options.apple === true
		const selection = view.dom.ownerDocument.getSelection()

		if (selection === null || (apple ? event.ctrlKey : event.metaKey)) return false

		selection.modify(
			event.shiftKey ? "extend" : "move",
			event.key === "ArrowLeft" ? "left" : "right",
			(apple ? event.altKey : event.ctrlKey) ? "word" : "character"
		)
		event.preventDefault()

		return true
	}

	private keyDown(view: EditorView, event: KeyboardEvent, readOnly = false): boolean {
		if (view.composing || event.isComposing) return false

		const editor = this.editor
		const apple = this.options.apple === true
		const mod = apple ? event.metaKey : event.ctrlKey
		const extend = event.shiftKey
		const state = editor.state

		switch (event.key) {
			case "ArrowUp":
			case "ArrowDown": {
				if (event.altKey) return false

				const dir = event.key === "ArrowDown" ? 1 : -1

				if (apple && event.metaKey) editor.moveDocumentBoundary(dir, extend)
				else if (!mod) editor.moveVertical(dir, extend)
				else return false

				return true
			}
			case "ArrowLeft":
			case "ArrowRight": {
				const dir = event.key === "ArrowRight" ? 1 : -1

				// Cmd+Arrow, and Ctrl+Arrow as macOS has it: the line's start or end.
				if (apple && (event.metaKey || (event.ctrlKey && !event.altKey))) return editor.moveLineBoundary(dir, extend)
				// Word steps: the browser's.
				if (mod || event.altKey || event.metaKey || event.ctrlKey) return false

				if (!extend) {
					const beside = besideNode(state, dir)

					if (beside !== null) return this.run(beside)
				}

				if (state.selection.empty || extend) {
					const head = state.selection.head
					// Right-to-left text: steps on screen, as the arrows show.
					const hit = editor.geometry.visualMove(head, editor.affinity, dir)

					if (hit !== undefined) return editor.moveTo(hit, extend, false)

					// Where a line wraps the input view wraps too, and a browser may only move its caret to the other
					// line's side of the same position: one character, as Word steps.
					const pos = editor.geometry.wrapsAt(head) ? stepCharacter(state.doc, head, dir) : undefined

					if (pos !== undefined) return editor.moveTo({ pos, affinity: 1, page: -1, line: -1, atom: false }, extend, false)
				}

				return false
			}
			case "Home":
			case "End": {
				const dir = event.key === "End" ? 1 : -1

				return mod ? editor.moveDocumentBoundary(dir, extend) : editor.moveLineBoundary(dir, extend)
			}
			case "PageUp":
			case "PageDown":
				return this.options.pageMove?.(event.key === "PageDown" ? 1 : -1, extend) ?? false
			case "Enter":
				if (mod || event.altKey || readOnly) return false

				return this.run(extend ? lineBreak(state) : splitParagraph(state))
			// At a paragraph's edge any deletion (a character, a word, to the line's start) joins paragraphs.
			case "Backspace":
				return readOnly ? false : this.run(joinBackward(state))
			case "Delete":
				// Shift+Delete cuts.
				return readOnly || extend ? false : this.run(joinForward(state))
			case "Escape":
				// The format painter goes first; then out of a header, footer or note, back to the body.
				if (editor.painter !== undefined) {
					editor.setPainter(undefined)

					return true
				}

				if (editor.activeStory === editor.body.session.story.story.index) return false

				editor.exitStory()

				return true
		}

		if (mod && !event.altKey && !extend) {
			const key = event.key.toLowerCase()

			if (key === "a") return this.run(selectAll(state))

			if (key === "f" || key === "g" || key === "h") {
				this.options.onFind?.()

				return this.options.onFind !== undefined
			}
		}

		const command = readOnly ? undefined : commandForKey(event, apple)

		if (command !== undefined) {
			// Undo and redo follow the order across stories; story commands go to the story being edited.
			void editor.execute(command)

			return true
		}

		return false
	}
}

// Where the text being composed starts: the occurrence of `text` around the caret nearest to where the
// composition started (`start`, mapped through the edits; a browser may rewrite more of the paragraph than
// the composition), or `start` itself.
function compositionFrom(doc: Node, head: number, start: number, text: string): number {
	const $head = doc.resolve(head)
	const parent = $head.parent

	if (!parent.isTextblock || text.length === 0) return Math.min(start, doc.content.size)

	const content = $head.start()
	const offset = $head.parentOffset
	const lo = Math.max(0, offset - text.length)
	const around = parent.textBetween(lo, Math.min(parent.content.size, offset + text.length), undefined, "\ufffc")
	let best = -1

	for (let at = around.indexOf(text); at >= 0; at = around.indexOf(text, at + 1)) {
		const pos = content + lo + at

		// The caret is inside or at either end of the composed text.
		if (pos <= head && pos + text.length >= head && (best < 0 || Math.abs(pos - start) < Math.abs(best - start))) best = pos
	}

	return best >= 0 ? best : Math.min(Math.max(start, content), content + parent.content.size)
}

// The position one character (a surrogate pair, an inline atom) after (dir 1) or before (-1) `pos` in its
// textblock, or undefined at the textblock's edge.
function stepCharacter(doc: Node, pos: number, dir: 1 | -1): number | undefined {
	const $pos = doc.resolve(pos)
	const parent = $pos.parent
	const offset = $pos.parentOffset

	if (!parent.isTextblock || (dir > 0 ? offset >= parent.content.size : offset <= 0)) return undefined

	const pair =
		dir > 0
			? parent.textBetween(offset, Math.min(parent.content.size, offset + 2), undefined, "\ufffc")
			: parent.textBetween(Math.max(0, offset - 2), offset, undefined, "\ufffc")
	const surrogates = pair.length === 2 && /^[\ud800-\udbff][\udc00-\udfff]$/.test(pair)

	return pos + dir * (surrogates ? 2 : 1)
}

function positionOf(doc: Node, index: number): number {
	let pos = 0

	for (let k = 0; k < index; k++) pos += doc.child(k).nodeSize

	return pos
}
