import { DOMParser, DOMSerializer, Fragment, Slice, type DOMOutputSpec, type Mark, type Node, type ResolvedPos } from "prosemirror-model"
import type { EditorState, Transaction } from "prosemirror-state"
import { Decoration, DecorationSet, EditorView, type NodeView, type NodeViewConstructor } from "prosemirror-view"
import { docxSchema } from "../../../core/docx/edit/index"
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

// The input sink of paged editing: a real ProseMirror view over the session's state that holds focus and
// receives all input (keys, beforeinput, IME composition, clipboard, drag and drop) while the pages show the
// document. It is invisible and renders only the blocks around the selection (every other top-level block is
// an empty placeholder), so it costs little however long the document; it is moved so its own caret sits
// where the caret is painted, which is where IME candidate windows and magnifiers anchor.

export interface EditSinkOptions {
	// The accessible name of the editing area.
	readonly label: string
	readonly readOnly?: boolean
	// Cmd (Apple platforms) or Ctrl as the shortcut modifier.
	readonly apple?: boolean
	// Top-level blocks rendered on each side of the selection's ends (default 20).
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

interface WindowSpec {
	readonly window: true
}

function inWindow(decorations: readonly Decoration[]): boolean {
	return decorations.some(decoration => (decoration.spec as Partial<WindowSpec> | undefined)?.window === true)
}

function element<K extends keyof HTMLElementTagNameMap>(tag: K): HTMLElementTagNameMap[K] {
	return document.createElement(tag)
}

// A block node's view: its real DOM inside the window (or nested in a block that is), an empty placeholder
// otherwise. Crossing the window's edge recreates the view.
function windowed(make: () => { dom: HTMLElement; contentDOM?: HTMLElement }, placeholder: () => HTMLElement): NodeViewConstructor {
	return (node, view, getPos, decorations) => {
		const pos = getPos()
		const top = pos === undefined || view.state.doc.resolve(pos).depth === 0

		if (!top || inWindow(decorations)) {
			const made = make()

			return {
				dom: made.dom,
				...(made.contentDOM === undefined ? {} : { contentDOM: made.contentDOM }),
				update: (next: Node, nextDecorations: readonly Decoration[]) =>
					next.type === node.type && (!top || inWindow(nextDecorations))
			} satisfies NodeView
		}

		return {
			dom: placeholder(),
			update: (next: Node, nextDecorations: readonly Decoration[]) => next.type === node.type && !inWindow(nextDecorations),
			ignoreMutation: () => true
		} satisfies NodeView
	}
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
	paragraph: windowed(
		() => {
			const dom = element("p")

			// Right-to-left paragraphs: the browser's own caret motion and IME follow the text's direction.
			dom.dir = "auto"

			return { dom, contentDOM: dom }
		},
		() => element("p")
	),
	table: windowed(
		() => {
			const dom = element("table")
			const contentDOM = element("tbody")

			dom.appendChild(contentDOM)

			return { dom, contentDOM }
		},
		() => element("div")
	),
	table_row: () => {
		const dom = element("tr")

		return { dom, contentDOM: dom }
	},
	table_cell: () => {
		const dom = element("td")

		return { dom, contentDOM: dom }
	},
	sdt: windowed(
		() => {
			const dom = element("div")

			return { dom, contentDOM: dom }
		},
		() => element("div")
	),
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

// Pasted HTML becomes paragraphs with basic formatting; the session's invariants make it new content.
function clipboardParser(): DOMParser {
	const block = (tag: string) => ({ tag, node: "paragraph" })

	return new DOMParser(docxSchema, [
		...["p", "h1", "h2", "h3", "h4", "h5", "h6", "li", "blockquote", "pre", "td", "th"].map(block),
		{ tag: "br", node: "break" },
		{ tag: "strong", mark: "bold", attrs: { value: true, cs: null } },
		{ tag: "b", mark: "bold", attrs: { value: true, cs: null } },
		{ tag: "em", mark: "italic", attrs: { value: true, cs: null } },
		{ tag: "i", mark: "italic", attrs: { value: true, cs: null } },
		{ tag: "u", mark: "underline" },
		{ tag: "s", mark: "strike", attrs: { value: true } },
		{ tag: "del", mark: "strike", attrs: { value: true } }
	])
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

export class EditSink {
	readonly element: HTMLDivElement
	readonly view: EditorView
	private readonly editor: PagedEditor
	private readonly options: EditSinkOptions
	private readonly windowSize: number
	private windowCache: { doc: Node; key: string; set: DecorationSet } | undefined = undefined
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

	constructor(editor: PagedEditor, parent: HTMLElement, options: EditSinkOptions) {
		this.editor = editor
		this.options = options
		this.windowSize = Math.max(1, options.window ?? DEFAULT_WINDOW)
		this.element = element("div")
		this.element.className = "fo-docx-sink"
		parent.appendChild(this.element)

		this.readOnlyOption = options.readOnly === true
		this.readOnlyShown = this.readOnly

		this.view = new EditorView(this.element, {
			state: editor.state,
			editable: () => !this.readOnly,
			dispatchTransaction: tr => this.dispatch(tr),
			decorations: state => this.windowDecorations(state),
			nodeViews: NODE_VIEWS,
			markViews: MARK_VIEWS,
			clipboardSerializer: clipboardSerializer(),
			clipboardParser: clipboardParser(),
			clipboardTextSerializer: (_slice, view) => {
				const { from, to } = view.state.selection

				return plainText(view.state.doc, from, to, pos => this.editor.geometry.markerText(pos))
			},
			clipboardTextParser: (text, $context) => plainSlice(text, $context),
			// The pages scroll to the painted caret; the sink never scrolls anything.
			handleScrollToSelection: () => true,
			handleKeyDown: (view, event) => this.keyDown(view, event),
			handleDOMEvents: {
				// Read-only, the view is not editable and ProseMirror's own key handling is off: selection keys only.
				keydown: (view, event) => (view.editable ? false : this.readOnlyKeyDown(view, event)),
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

			const state = editor.state

			if (this.view.state !== state) {
				if (event.type === "story") this.windowCache = undefined

				this.view.updateState(state)
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

	private dispatch(tr: Transaction): void {
		if (tr.docChanged && this.readOnly) {
			// Whatever the browser changed in the view is put back.
			this.view.updateState(this.editor.state)

			return
		}

		if (this.compositionStart >= 0 && tr.docChanged) this.compositionStart = tr.mapping.map(this.compositionStart, -1)

		this.dispatching = true

		try {
			this.editor.dispatch(tr)
		} finally {
			this.dispatching = false
		}

		this.view.updateState(this.editor.state)

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

	private run(tr: Transaction | null | false): boolean {
		if (tr === null) return false
		if (tr !== false) this.dispatch(tr)

		return true
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
		const state = view.state

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
				// Out of a header, footer or note, back to the body.
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
			editor.session.execute(command)

			return true
		}

		return false
	}

	private windowDecorations(state: EditorState): DecorationSet {
		const doc = state.doc
		const { from, to } = state.selection
		const count = doc.childCount
		const n = this.windowSize
		const a = doc.resolve(from).index(0)
		const b = Math.min(count - 1, doc.resolve(to).index(0))
		// A selection spanning many blocks renders around both ends only.
		const ranges: [number, number][] =
			b - a > 2 * n
				? [
						[Math.max(0, a - n), Math.min(count, a + n + 1)],
						[Math.max(0, b - n), Math.min(count, b + n + 1)]
					]
				: [[Math.max(0, a - n), Math.min(count, b + n + 1)]]
		const key = ranges.map(range => range.join("-")).join(",")
		const cached = this.windowCache

		if (cached !== undefined && cached.doc === doc && cached.key === key) return cached.set

		const story = this.editor.story
		const starts = story.document === doc ? story.starts : undefined
		const decorations: Decoration[] = []

		for (const [lo, hi] of ranges) {
			let pos = starts === undefined ? positionOf(doc, lo) : (starts[lo] ?? 0)

			for (let k = lo; k < hi; k++) {
				const child = doc.child(k)

				decorations.push(Decoration.node(pos, pos + child.nodeSize, {}, { window: true } satisfies WindowSpec))
				pos += child.nodeSize
			}
		}

		const set = DecorationSet.create(doc, decorations)

		this.windowCache = { doc, key, set }

		return set
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
