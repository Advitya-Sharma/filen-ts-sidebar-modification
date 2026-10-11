// @vitest-environment happy-dom
import { act } from "react"
import { createRoot, type Root } from "react-dom/client"
import { TextSelection } from "prosemirror-state"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { SOURCE_INTS, type DocumentLayout, type PageDisplayList } from "@filen/office/core/docx"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import type { DocxPageData, DocxSessionApi, DocxTextStyle } from "@filen/office/dom/docx"
import { DocxMediaCache, paintPage, patchPage, type DocxDocument, type DocxTables } from "@filen/office/react/docx"
import { DocxEditor, EditSink, type DocxEditorHandle, type DocxEditorMark } from "@filen/office/react/docx/editor"
import { plainSlice, plainText } from "../react/docx/edit/sink"
import { buildDocx } from "./docxSynthetic"
import { pagedEditor, settle } from "./pagedDomFixtures"

Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true)

const SESSION: DocxSessionApi = {
	objects: () => Promise.resolve([]),
	media: () => Promise.resolve(null),
	locate: () => Promise.resolve(null),
	find: () => Promise.resolve(new Int32Array(0)),
	annotations: () => Promise.resolve(null),
	close: () => undefined
}

// The renderer's view of a layout page (no objects or links in these documents).
function pageData(page: PageDisplayList): DocxPageData {
	const text = page.text

	return {
		index: page.index,
		number: page.number,
		numberText: page.numberText,
		section: page.section,
		width: page.width,
		height: page.height,
		lineCount: text.lineCount,
		fragmentCount: text.fragmentCount,
		decorationCount: text.decorationCount,
		lineFloats: text.lineFloats,
		lineInts: text.lineInts,
		fragmentFloats: text.fragmentFloats,
		fragmentInts: text.fragmentInts,
		decorationFloats: text.decorationFloats,
		decorationInts: text.decorationInts,
		graphicCount: page.graphicCount,
		graphicFloats: page.graphicFloats,
		graphicInts: page.graphicInts,
		objectCount: 0,
		objectFloats: new Float32Array(0),
		objectInts: new Int32Array(0),
		objectKinds: new Uint8Array(0),
		sourceCount: page.sourceCount,
		sourceInts: page.sourceInts.slice(0, page.sourceCount * SOURCE_INTS),
		sourceTexts: page.sourceTexts,
		linkCount: 0,
		linkInts: new Int32Array(0)
	}
}

function tablesOf(layout: DocumentLayout): DocxTables {
	const styles: DocxTextStyle[] = []

	for (let id = 0; id < layout.styles.count; id++) {
		const style = layout.styles.get(id)

		styles.push({
			font: `400 ${Math.round(style.size)}px sans-serif`,
			smallCapsFont: "400 10px sans-serif",
			color: "",
			kerning: true,
			letterSpacing: 0,
			scaleX: 1,
			shift: 0,
			ascent: style.size * 0.8,
			descent: style.size * 0.2,
			revision: 0
		})
	}

	return { styles, decorations: [], links: [] }
}

function documentOf(editor: PagedEditor): DocxDocument {
	const pages: DocxPageData[] = []

	for (let index = 0; index < editor.pageCount; index++) {
		const page = editor.page(index)

		if (page !== undefined) pages.push(pageData(page))
	}

	return {
		status: "ready",
		info: undefined,
		pages,
		pageCount: pages.length,
		tables: tablesOf(editor.layout),
		session: SESSION,
		media: new DocxMediaCache(SESSION),
		error: undefined
	}
}

const p = (text: string): string => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
const LONG = buildDocx({
	body: Array.from({ length: 60 }, (_, k) => p(`Paragraph ${k} ${"lorem ipsum dolor sit amet ".repeat(12)}`)).join("")
})

function paragraphPos(editor: PagedEditor, index: number): number {
	let found = -1

	editor.state.doc.forEach((_node, offset, k) => {
		if (k === index) found = offset
	})

	return found
}

const frame = (): Promise<void> => new Promise(resolve => requestAnimationFrame(() => setTimeout(resolve, 0)))

describe("line patching", () => {
	it("repaints only changed lines and equals a fresh paint", async () => {
		const { editor } = await pagedEditor(LONG)
		const before = editor.page(1)

		if (before === undefined) throw new Error("no page")

		const tables = tablesOf(editor.layout)
		const painted = paintPage(pageData(before), { document, tables })
		// A paragraph in the middle of page 2.
		const target = editor.geometry.hitTest(1, 300, 500)

		if (target === undefined) throw new Error("no hit")

		editor.dispatch(editor.state.tr.insertText("typed words ", target.pos))
		await settle(editor)

		const after = editor.page(1)

		if (after === undefined || after === before) throw new Error("page not laid out again")

		const inserted = vi.spyOn(painted.parts.textElements, "insertBefore")
		const patched = patchPage(painted, pageData(after), { document, tables: tablesOf(editor.layout) })
		const fresh = paintPage(pageData(after), { document, tables: tablesOf(editor.layout) })

		expect(patched.painted.root.innerHTML).toBe(fresh.root.innerHTML)

		// Lines above the edit kept their elements.
		const kept = patched.painted.lines.filter(line => painted.lines.some(old => old.text === line.text))

		expect(kept.length).toBeGreaterThan(5)
		// Only new lines were put in the DOM: kept ones stay where they are (a moved element is restyled and
		// laid out again).
		expect(inserted.mock.calls.length).toBe(patched.painted.lines.length - kept.length)
		expect(patched.painted.lines[0]?.text).toBe(painted.lines[0]?.text)
	})
})

describe("DocxEditor", () => {
	let container: HTMLDivElement
	let root: Root

	beforeEach(() => {
		container = document.createElement("div")
		container.style.height = "800px"
		document.body.appendChild(container)
		root = createRoot(container)
	})

	afterEach(() => {
		act(() => root.unmount())
		container.remove()
	})

	it("renders pages, an input view windowed around the selection, caret and selection", async () => {
		const { editor } = await pagedEditor(LONG)
		const handle = { current: null as DocxEditorHandle | null }

		act(() =>
			root.render(
				<DocxEditor
					ref={value => {
						handle.current = value
					}}
					editor={editor}
					document={documentOf(editor)}
					zoom={1}
					label="Document"
					window={3}
				/>
			)
		)
		await act(frame)

		const textbox = container.querySelector<HTMLElement>("[role=textbox]")

		expect(textbox?.getAttribute("aria-label")).toBe("Document")
		expect(container.querySelector(".fo-docx-pages")?.getAttribute("aria-hidden")).toBe("true")
		// Only the blocks near the selection have content in the input view.
		const filled = Array.from(textbox?.children ?? []).filter(child => child.childNodes.length > 0)

		expect(textbox?.children.length).toBe(editor.state.doc.childCount)
		expect(filled.length).toBeLessThanOrEqual(7)
		expect(container.querySelectorAll(".fo-docx-page").length).toBeGreaterThan(0)

		// A caret in paragraph 2, drawn on its page.
		await act(async () => {
			editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, paragraphPos(editor, 2) + 5)))
			handle.current?.focus()
			await frame()
		})
		expect(container.querySelectorAll(".fo-docx-caret").length).toBe(1)

		// The window followed the selection to paragraph 40.
		await act(async () => {
			const at = paragraphPos(editor, 40) + 3

			editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, at, at + 30)))
			await frame()
		})

		const blocks = Array.from(textbox?.children ?? [])

		expect(blocks[40]?.childNodes.length).toBeGreaterThan(0)
		expect(blocks[2]?.childNodes.length).toBe(0)

		// A range: selection boxes on the mounted page that shows it, no caret.
		const ends = editor.ends()

		if (ends.start !== undefined) {
			act(() => handle.current?.scrollToPage(ends.start?.page ?? 0))
			await act(frame)
		}

		expect(container.querySelectorAll(".fo-docx-caret").length).toBe(0)
	})

	it("types through the input view and patches the page", async () => {
		const { editor } = await pagedEditor(LONG)
		let document = documentOf(editor)
		const render = (): void =>
			root.render(
				<DocxEditor
					editor={editor}
					document={document}
					zoom={1}
					label="Document"
				/>
			)

		act(render)
		await act(frame)

		const before = container.querySelector('[data-fo-page="0"] .fo-docx-text')?.textContent ?? ""

		await act(async () => {
			editor.dispatch(editor.state.tr.insertText("HELLO ", paragraphPos(editor, 0) + 1))
			await settle(editor)
			document = documentOf(editor)
			render()
			await frame()
		})

		const after = container.querySelector('[data-fo-page="0"] .fo-docx-text')?.textContent ?? ""

		expect(after).not.toBe(before)
		expect(after).toContain("HELLO")
	})
})

const GAP = 16

// A point of page `index` on screen (happy-dom lays nothing out: client space is the content's own, the
// pages stacked from the gap at the left edge).
function clientPoint(editor: PagedEditor, index: number, x: number, y: number): { clientX: number; clientY: number } {
	let top = GAP

	for (let k = 0; k < index; k++) top += (editor.page(k)?.height ?? 0) + GAP

	return { clientX: GAP + x, clientY: top + y }
}

function pointer(type: string, point: { clientX: number; clientY: number }, init: PointerEventInit = {}): PointerEvent {
	return new PointerEvent(type, { bubbles: true, cancelable: true, button: 0, pointerId: 1, pointerType: "mouse", ...point, ...init })
}

function key(target: Element, init: KeyboardEventInit): KeyboardEvent {
	const event = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init })

	target.dispatchEvent(event)

	return event
}

// The first textblock whose text starts with `text`: its content's start.
function textStart(editor: PagedEditor, text: string): number {
	let found = -1

	editor.state.doc.descendants((node, pos) => {
		if (found < 0 && node.isTextblock && node.textContent.startsWith(text)) found = pos + 1

		return found < 0
	})

	return found
}

const caretAt = (editor: PagedEditor, pos: number): void => {
	editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos)))
}

describe("DocxEditor input", () => {
	let container: HTMLDivElement
	let root: Root

	beforeEach(() => {
		container = document.createElement("div")
		document.body.appendChild(container)
		root = createRoot(container)
	})

	afterEach(() => {
		act(() => root.unmount())
		container.remove()
	})

	const mount = async (editor: PagedEditor, props: { readOnly?: boolean; marks?: readonly DocxEditorMark[] } = {}) => {
		const handle = { current: null as DocxEditorHandle | null }

		act(() =>
			root.render(
				<DocxEditor
					ref={value => {
						handle.current = value
					}}
					editor={editor}
					document={documentOf(editor)}
					zoom={1}
					gap={GAP}
					label="Document"
					{...props}
				/>
			)
		)
		await act(frame)

		const content = container.querySelector<HTMLElement>(".fo-docx-pages-content")
		const textbox = container.querySelector<HTMLElement>("[role=textbox]")

		if (content === null || textbox === null) throw new Error("not rendered")

		return { handle, content, textbox }
	}

	it("counts clicks itself: a double click selects the word, a triple click the paragraph", async () => {
		const { editor } = await pagedEditor(LONG)
		const { content } = await mount(editor)
		const box = editor.geometry.locate(paragraphPos(editor, 1) + 14)

		if (box === undefined) throw new Error("no caret")

		const point = clientPoint(editor, box.page, box.x + 1, box.top + box.height / 2)
		const press = (init: PointerEventInit = {}): void =>
			act(() => {
				content.dispatchEvent(pointer("pointerdown", point, init))
				content.dispatchEvent(pointer("pointerup", point, init))
			})
		const paragraph = editor.state.doc.child(1)

		press()
		expect(editor.selection().empty).toBe(true)

		press()

		const word = editor.selection()

		expect(word.empty).toBe(false)
		expect(word.to - word.from).toBeLessThan(10)

		press()
		expect(editor.selection().to - editor.selection().from).toBe(paragraph.content.size)

		// Shift+click extends instead, from the paragraph's start.
		press({ shiftKey: true })
		expect(editor.selection().anchor).toBe(paragraphPos(editor, 1) + 1)
		expect(editor.selection().head).toBeLessThan(paragraphPos(editor, 1) + 1 + paragraph.content.size)
	})

	it("places the caret on the nearest line for presses between and beside pages", async () => {
		const { editor } = await pagedEditor(LONG)
		const { content } = await mount(editor)
		const page = editor.page(0)

		if (page === undefined) throw new Error("no page")

		// Just below page 0: its last line; just above page 1: its first.
		act(() => content.dispatchEvent(pointer("pointerdown", clientPoint(editor, 0, 300, page.height + 3))))
		act(() => content.dispatchEvent(pointer("pointerup", clientPoint(editor, 0, 300, page.height + 3))))
		expect(editor.caret()?.page).toBe(0)

		act(() => content.dispatchEvent(pointer("pointerdown", clientPoint(editor, 1, 300, -3), { clientX: GAP + 300 + 1000 })))
		act(() => content.dispatchEvent(pointer("pointerup", clientPoint(editor, 1, 300, -3))))

		const caret = editor.caret()

		expect(caret?.page).toBe(1)
		// Far right of the page: the line's end.
		expect(caret?.x ?? 0).toBeGreaterThan(page.width / 2)
	})

	it("leaves touch scrolling to the browser and keeps focus after a tap", async () => {
		const { editor } = await pagedEditor(LONG)
		const { content, textbox } = await mount(editor)
		const before = editor.selection().head
		const point = clientPoint(editor, 0, 300, 300)
		const down = pointer("pointerdown", point, { pointerType: "touch" })

		act(() => content.dispatchEvent(down))
		expect(down.defaultPrevented).toBe(false)

		// The browser took the gesture over: nothing selected or focused.
		act(() => content.dispatchEvent(pointer("pointercancel", point, { pointerType: "touch" })))
		expect(editor.selection().head).toBe(before)
		expect(document.activeElement).not.toBe(textbox)

		// The compatibility mousedown after a tap moves no focus.
		const mouse = new MouseEvent("mousedown", { bubbles: true, cancelable: true, button: 0, ...point })

		content.dispatchEvent(mouse)
		expect(mouse.defaultPrevented).toBe(true)
	})

	it("keeps a touch handle's other end where it was when the drag began", async () => {
		const { editor } = await pagedEditor(LONG)
		const { handle } = await mount(editor)
		const start = paragraphPos(editor, 1) + 20

		act(() => {
			editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, start, start + 5)))
		})

		const end = editor.geometry.locate(start + 5)

		if (end === undefined) throw new Error("no end")

		// The start handle dragged past the end, then further.
		for (const dx of [60, 120]) {
			const point = clientPoint(editor, end.page, end.x + dx, end.top + end.height / 2)

			act(() => handle.current?.dragSelectionEnd("start", point.clientX, point.clientY))
			expect(Math.min(editor.selection().anchor, editor.selection().head)).toBe(start + 5)
		}

		act(() => handle.current?.endSelectionDrag())
	})

	it("is read-only through the view: no edits, selection and focus still", async () => {
		const { editor } = await pagedEditor(LONG)
		const { textbox, handle } = await mount(editor, { readOnly: true })
		const doc = editor.state.doc

		expect(textbox.getAttribute("contenteditable")).toBe("false")
		expect(textbox.getAttribute("aria-readonly")).toBe("true")
		expect(textbox.getAttribute("inputmode")).toBe("none")
		expect(textbox.getAttribute("tabindex")).toBe("0")
		expect(container.querySelector(".fo-docx-editor")?.hasAttribute("data-fo-read-only")).toBe(true)

		act(() => handle.current?.focus())
		expect(document.activeElement).toBe(textbox)

		const pos = paragraphPos(editor, 2) + 3

		act(() => caretAt(editor, pos))
		act(() => {
			key(textbox, { key: "End" })
			key(textbox, { key: "Enter" })
			key(textbox, { key: "Backspace" })
			key(textbox, { key: "b", ctrlKey: true })
		})
		expect(editor.state.doc).toBe(doc)
		expect(editor.selection().head).toBeGreaterThan(pos)

		// Editable again.
		await mount(editor, { readOnly: false })
		expect(textbox.getAttribute("contenteditable")).toBe("true")
		expect(textbox.hasAttribute("aria-readonly")).toBe(false)
	})

	it("draws the marks of each page in one pass, keeping the boxes that did not move", async () => {
		const { editor } = await pagedEditor(LONG)
		const marks: DocxEditorMark[] = []

		editor.state.doc.forEach((_node, offset) => marks.push({ from: offset + 1, to: offset + 10, kind: "spelling" }))

		const rangeRects = vi.spyOn(editor, "rangeRects")
		const rangesRects = vi.spyOn(editor.geometry, "rangesRects")
		const { handle } = await mount(editor, { marks })
		const mounted = Array.from(container.querySelectorAll<HTMLElement>("[data-fo-page]")).map(page => Number(page.dataset["foPage"]))
		const drawn = (): Element[] => Array.from(container.querySelectorAll('.fo-docx-mark[data-fo-mark="spelling"]'))
		// A box per mark on a line: a wave starting at each box.
		const boxes = (): number => drawn().reduce((sum, path) => sum + (path.getAttribute("d")?.match(/M/g)?.length ?? 0), 0)
		let expected = 0

		for (const index of mounted) for (const mark of marks) expected += editor.rangeRects(mark.from, mark.to, index).length

		expect(mounted.length).toBeGreaterThan(0)
		expect(expected).toBeGreaterThan(0)
		expect(boxes()).toBe(expected)
		// One call per page drawn, never one per mark.
		expect(rangeRects.mock.calls.length).toBe(mounted.length * marks.length)
		expect(rangesRects.mock.calls.length).toBeLessThanOrEqual(mounted.length * 2)

		const before = drawn()
		const calls = rangesRects.mock.calls.length

		// The caret moving redraws no mark.
		await act(async () => {
			editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, paragraphPos(editor, 3) + 2)))
			handle.current?.focus()
			await frame()
		})
		expect(rangesRects.mock.calls.length).toBe(calls)
		expect(drawn()).toEqual(before)

		const caret = container.querySelector(".fo-docx-caret")

		expect(caret).not.toBeNull()

		// Typing: paths of lines it did not touch are the same elements, and so is the caret.
		await act(async () => {
			editor.dispatch(editor.state.tr.insertText("x"))
			await frame()
		})

		const after = drawn()

		expect(after.filter(element => before.includes(element)).length).toBeGreaterThan(before.length / 2)
		expect(container.querySelector(".fo-docx-caret")).toBe(caret)
		expect(container.querySelectorAll(".fo-docx-caret").length).toBe(1)
	})

	it("draws a path of marks per line and touches only the path of the line typed in", async () => {
		const { editor } = await pagedEditor(LONG)
		let marks: DocxEditorMark[] = []

		editor.state.doc.forEach((_node, offset) => marks.push({ from: offset + 1, to: offset + 10, kind: "spelling" }))
		await mount(editor, { marks })

		const rows = (): Map<Element, string> =>
			new Map(Array.from(container.querySelectorAll(".fo-docx-mark"), path => [path, path.getAttribute("d") ?? ""]))
		const before = rows()

		// One path per line: every mark here sits on its paragraph's first line.
		expect(before.size).toBeGreaterThan(5)
		expect([...before.values()].every(d => d.match(/M/g)?.length === 1)).toBe(true)
		expect(container.querySelectorAll("svg.fo-docx-marks").length).toBe(container.querySelectorAll("[data-fo-page]").length)

		const tr = editor.state.tr.insertText("x", paragraphPos(editor, 3) + 3)

		marks = marks.map(mark => ({ ...mark, from: tr.mapping.map(mark.from), to: tr.mapping.map(mark.to) }))
		editor.dispatch(tr)
		await mount(editor, { marks })

		const after = rows()
		const changed = [...after].filter(([row, content]) => before.get(row) !== content)

		expect(after.size).toBe(before.size)
		expect(changed.length).toBe(1)
	})

	it("draws find results in a drawing of their own, under no other kind", async () => {
		const { editor } = await pagedEditor(LONG)
		const start = paragraphPos(editor, 0) + 1

		await mount(editor, {
			marks: [
				{ from: start, to: start + 9, kind: "find-active" },
				{ from: start + 10, to: start + 15, kind: "repeated" },
				{ from: start + 20, to: start + 25, kind: "find" }
			]
		})

		const blended = container.querySelector("svg.fo-docx-marks[data-fo-blend]")
		const plain = container.querySelector("svg.fo-docx-marks:not([data-fo-blend])")

		expect(blended?.querySelectorAll('[data-fo-mark="find"], [data-fo-mark="find-active"]').length).toBe(2)
		expect(plain?.querySelectorAll('[data-fo-mark="repeated"]').length).toBe(1)
		expect(blended?.querySelector('[data-fo-mark="find-active"]')?.getAttribute("d")).toMatch(/^M[\d.]+ 0h[\d.]+v[\d.]+h-[\d.]+z$/)
		// The plain drawing comes first: highlights blend over it.
		expect(plain?.compareDocumentPosition(blended as Node)).toBe(Node.DOCUMENT_POSITION_FOLLOWING)
	})
})

describe("EditSink", () => {
	let host: HTMLDivElement

	beforeEach(() => {
		host = document.createElement("div")
		document.body.appendChild(host)
	})

	afterEach(() => host.remove())

	const sinkOf = (editor: PagedEditor, options: { readOnly?: boolean; apple?: boolean } = {}): EditSink =>
		new EditSink(editor, host, { label: "Document", caretClientRect: () => undefined, ...options })

	it("refuses edits read-only and allows them once writable", async () => {
		const { editor } = await pagedEditor(LONG)
		const sink = sinkOf(editor, { readOnly: true })
		const doc = editor.state.doc
		const pos = paragraphPos(editor, 1) + 2

		sink.view.dispatch(sink.view.state.tr.insertText("typed", pos))
		expect(editor.state.doc).toBe(doc)

		sink.view.dispatch(sink.view.state.tr.setSelection(TextSelection.create(sink.view.state.doc, pos)))
		expect(editor.selection().head).toBe(pos)

		sink.setReadOnly(false)
		expect(sink.view.editable).toBe(true)
		expect(sink.view.dom.getAttribute("contenteditable")).toBe("true")

		sink.view.dispatch(sink.view.state.tr.insertText("typed", pos))
		expect(editor.state.doc.textBetween(pos, pos + 5)).toBe("typed")
		sink.destroy()
	})

	it("underlines the composed text where a browser rewrote the whole paragraph", async () => {
		const { editor } = await pagedEditor(LONG)
		const sink = sinkOf(editor)
		const start = paragraphPos(editor, 1) + 1
		const caret = start + 10
		const paragraph = editor.state.doc.child(1)
		const composing = (text: string): void => {
			const before = paragraph.textBetween(0, 10)
			const after = paragraph.textBetween(10, paragraph.content.size)
			const update = Object.assign(new Event("compositionupdate", { bubbles: true }), { data: text })

			sink.view.dom.dispatchEvent(update)

			// The browser's caret: after the composed text.
			const tr = sink.view.state.tr.insertText(before + text + after, start, start + editor.state.doc.child(1).content.size)

			sink.view.dispatch(tr.setSelection(TextSelection.create(tr.doc, caret + text.length)))
		}

		caretAt(editor, caret)
		sink.view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }))
		expect(sink.view.composing).toBe(true)

		composing("あ")
		composing("あい")
		expect(editor.composition).toEqual({ from: caret, to: caret + 2 })

		sink.view.dom.dispatchEvent(Object.assign(new Event("compositionend", { bubbles: true }), { data: "あい" }))
		expect(editor.composition).toBeUndefined()
		sink.destroy()
	})

	it("sizes its text like the painted caret's line and column", async () => {
		const { editor } = await pagedEditor(LONG)
		const sink = new EditSink(editor, host, {
			label: "Document",
			caretClientRect: () => ({ left: 100, top: 200, height: 24, width: 480 })
		})

		sink.align()
		expect(sink.element.style.lineHeight).toBe("24px")
		expect(sink.element.style.fontSize).toBe("20px")
		expect(sink.element.style.width).toBe("480px")
		sink.destroy()
	})

	it("joins paragraphs on word and line deletions at their edges, Ctrl+Arrow to line ends on macOS", async () => {
		const { editor } = await pagedEditor(LONG)
		const sink = sinkOf(editor, { apple: true })
		const count = editor.state.doc.childCount

		caretAt(editor, paragraphPos(editor, 2) + 1)
		key(sink.view.dom, { key: "Backspace", altKey: true })
		expect(editor.state.doc.childCount).toBe(count - 1)

		const end = paragraphPos(editor, 3) + editor.state.doc.child(3).nodeSize - 1

		caretAt(editor, end)
		key(sink.view.dom, { key: "Delete", altKey: true })
		expect(editor.state.doc.childCount).toBe(count - 2)

		caretAt(editor, paragraphPos(editor, 5) + 3)
		key(sink.view.dom, { key: "ArrowRight", ctrlKey: true })

		const line = editor.geometry.lineBoundary(paragraphPos(editor, 5) + 3, 1, 1)

		expect(editor.selection().head).toBe(line?.pos)
		sink.destroy()
	})

	it("steps on screen through right-to-left lines", async () => {
		const { editor } = await pagedEditor(
			buildDocx({ body: '<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:t xml:space="preserve">שלום עולם יפה</w:t></w:r></w:p>' })
		)
		const sink = sinkOf(editor)
		const pos = textStart(editor, "שלום") + 4

		caretAt(editor, pos)

		const x = editor.caret()?.x ?? 0

		key(sink.view.dom, { key: "ArrowLeft" })
		expect(editor.selection().head).toBe(pos + 1)
		expect(editor.caret()?.x ?? 0).toBeLessThan(x)
		expect(sink.view.dom.querySelector("p")?.dir).toBe("auto")
		sink.destroy()
	})
})

describe("plain text clipboard", () => {
	it("copies paragraphs, table cells and rows, empty paragraphs and list numbers like Word", async () => {
		const { editor } = await pagedEditor(
			buildDocx({
				body:
					p("Before") +
					p("") +
					'<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
					`<w:tr><w:tc>${p("A1")}</w:tc><w:tc>${p("B1")}</w:tc></w:tr><w:tr><w:tc>${p("A2")}</w:tc><w:tc>${p("B2")}</w:tc></w:tr></w:tbl>` +
					p("Item one") +
					p("After")
			})
		)
		const doc = editor.state.doc
		const item = textStart(editor, "Item")
		const marker = (pos: number): string => (pos === item ? "1.\t" : "")

		expect(plainText(doc, 0, doc.content.size, marker)).toBe("Before\n\nA1\tB1\nA2\tB2\n1.\tItem one\nAfter")
		// From inside a paragraph: no marker for a paragraph started before the range.
		expect(plainText(doc, item + 2, item + 6, marker)).toBe("em o")
	})

	it("pastes plain text as one paragraph per line, tabs as tabs", async () => {
		const { editor } = await pagedEditor(LONG)
		const slice = plainSlice("a\tb\r\n\nc\u0001", editor.state.doc.resolve(paragraphPos(editor, 1) + 3))
		const paragraphs: string[] = []

		slice.content.forEach(node => paragraphs.push(node.textBetween(0, node.content.size, undefined, leaf => leaf.type.name)))
		expect(paragraphs).toEqual(["atabb", "", "c"])
		expect(slice.openStart).toBe(1)
		expect(slice.openEnd).toBe(1)
	})
})
