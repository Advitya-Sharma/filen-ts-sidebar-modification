import type { Node } from "prosemirror-model"
import type { PageDisplayList } from "../../../core/docx/layout/index"
import type { CaretBox, PagedEditor, PageRect } from "../../../core/docx/paged/index"
import type { DocxPageData, DocxSessionApi } from "../../../dom/docx/protocol"
import type { DocxMediaCache } from "../media"
import { paintObjects } from "../paint/objects"
import { paintPage, patchPage, type PaintedPage } from "../paint/page"
import type { DocxTables } from "../store"

// Draws the pages of a paged editor: each mounted page is painted from its display list and, when laid out
// again, patched line by line; above it a layer with the caret, the selection, the text being composed and
// marks from outside (spelling). Structure and data-* attributes only; the look is --fo-* variables.

// A range to mark on the pages (a spelling mistake, a search result) and what kind of mark it is.
export interface DocxEditorMark {
	readonly from: number
	readonly to: number
	readonly kind: string
}

interface MountedPage {
	readonly host: HTMLElement
	readonly layer: HTMLElement
	// Inside the layer, in drawing order: the dimmed body, the marks, selection and composition, the caret.
	readonly under: HTMLElement
	readonly marks: HTMLElement
	readonly over: HTMLElement
	readonly caret: HTMLElement
	// What the marks were last drawn from (redrawn only when one of these changed), the drawings that hold
	// them (made when first needed) and the rows drawn, by what they show: a redraw keeps the rows that did not
	// change (typing changes the line typed in) and moves those that only moved.
	drawn: { page: PageDisplayList | undefined; doc: Node | undefined; marks: readonly DocxEditorMark[]; story: number }
	// The marks drawn on the page, relative to its first position (from, to per mark) and their kinds: an edit
	// elsewhere moves them all alike and changes nothing drawn.
	shown: { offsets: number[]; kinds: string[] }
	drawings: { plain: SVGSVGElement | undefined; blend: SVGSVGElement | undefined }
	rows: Map<string, MarkRow[]>
	painted: PaintedPage | undefined
	release: () => void
	signal: { aborted: boolean }
}

// The marks of one kind on one line of a page: one path in the page's drawing, in line coordinates, moved
// as a whole when the line moves. A page with a thousand marks is a path per line, and a keystroke changes the
// path of the line typed in.
interface MarkRow {
	readonly element: SVGPathElement
	// Its kind, line height and boxes: equal content draws the same.
	content: string
	top: number
}

interface MarkLine {
	readonly kind: string
	readonly top: number
	readonly height: number
	// Left and width of each box.
	readonly boxes: number[]
}

const SVG = "http://www.w3.org/2000/svg"
// Mark kinds drawn as highlights over the text: they multiply with it, in a drawing of their own.
const BLENDED = new Set(["find", "find-active"])
// A misspelling's wave: half-waves this wide, peaks this far from the line's bottom.
const WAVE = 3
const WAVE_TOP = 2.25
const WAVE_BOTTOM = 0.25

// The path of a line's marks of `kind` (line coordinates: 0 is the line's top): a wave under a misspelling, a
// double line under a repeated word, a box over a find result, a dotted line under anything else.
function markPath(kind: string, height: number, boxes: readonly number[]): string {
	let d = ""

	for (let b = 0; b + 1 < boxes.length; b += 2) {
		const x = boxes[b] ?? 0
		const width = boxes[b + 1] ?? 0

		if (kind === "spelling") {
			const waves = Math.max(1, Math.round(width / WAVE))
			const middle = height - (WAVE_TOP + WAVE_BOTTOM) / 2
			// A quadratic's extreme lies halfway to its control point.
			const lift = WAVE_TOP - WAVE_BOTTOM

			d += `M${round(x)} ${round(middle)}q${WAVE / 2} ${-lift} ${WAVE} 0${`t${WAVE} 0`.repeat(waves - 1)}`
		} else if (kind === "repeated") {
			d += `M${round(x)} ${round(height - 0.5)}h${round(width)}M${round(x)} ${round(height - 2.5)}h${round(width)}`
		} else if (BLENDED.has(kind)) {
			d += `M${round(x)} 0h${round(width)}v${round(height)}h${round(-width)}z`
		} else {
			d += `M${round(x)} ${round(height - 1)}h${round(width)}`
		}
	}

	return d
}

export interface SurfaceSources {
	readonly session: DocxSessionApi
	readonly media: DocxMediaCache
	readonly tables: DocxTables
}

export class EditSurface {
	private readonly editor: PagedEditor
	private readonly mounted = new Map<number, MountedPage>()
	private sources: SurfaceSources | undefined = undefined
	private pages: readonly (DocxPageData | undefined)[] = []
	// Sorted by `from`; the longest's length bounds where a line's marks start.
	private marks: readonly DocxEditorMark[] = []
	private longestMark = 0
	private focused = false
	private caretPage: { index: number; box: CaretBox } | undefined = undefined
	// One caret element, moved rather than made again: its blink runs on the compositor, and a new one each
	// keystroke would change the page's layers (a costly commit on a page with many marks).
	private caretElement: HTMLElement | undefined = undefined
	private caretStyle = ""
	private readonly boxesScratch: number[] = []
	private readonly ownersScratch: number[] = []
	private blinkPhase = false

	constructor(editor: PagedEditor) {
		this.editor = editor
	}

	mount(index: number, host: HTMLElement, layer: HTMLElement): void {
		this.unmount(index)

		const document = layer.ownerDocument
		const under = document.createElement("div")
		const marks = document.createElement("div")
		const over = document.createElement("div")
		const caret = document.createElement("div")

		layer.replaceChildren(under, marks, over, caret)
		this.mounted.set(index, {
			host,
			layer,
			under,
			marks,
			over,
			caret,
			drawn: { page: undefined, doc: undefined, marks: [], story: -1 },
			shown: { offsets: [], kinds: [] },
			drawings: { plain: undefined, blend: undefined },
			rows: new Map(),
			painted: undefined,
			release: () => undefined,
			signal: { aborted: false }
		})
		this.paint(index)
		this.drawPage(index)
	}

	unmount(index: number): void {
		const page = this.mounted.get(index)

		if (page === undefined) return

		page.signal.aborted = true
		page.release()
		page.host.replaceChildren()
		page.layer.replaceChildren()
		this.mounted.delete(index)
	}

	dispose(): void {
		for (const index of [...this.mounted.keys()]) this.unmount(index)
	}

	// The renderer's page data and tables; mounted pages whose data changed are patched. Whether any was.
	update(pages: readonly (DocxPageData | undefined)[], sources: SurfaceSources): boolean {
		let repainted = false

		this.pages = pages
		this.sources = sources

		for (const index of this.mounted.keys()) if (this.paint(index)) repainted = true

		return repainted
	}

	// Whether the marks changed (draw() shows them).
	setMarks(marks: readonly DocxEditorMark[]): boolean {
		if (this.marks === marks) return false

		let sorted = true
		let longest = 0

		for (let k = 0; k < marks.length; k++) {
			const mark = marks[k]

			if (mark === undefined) continue

			longest = Math.max(longest, mark.to - mark.from)

			if (k > 0 && (marks[k - 1]?.from ?? 0) > mark.from) sorted = false
		}

		this.marks = sorted ? marks : [...marks].sort((a, b) => a.from - b.from)
		this.longestMark = longest

		return true
	}

	setFocused(focused: boolean): boolean {
		if (this.focused === focused) return false

		this.focused = focused

		return true
	}

	private paint(index: number): boolean {
		const mounted = this.mounted.get(index)
		const page = this.pages[index]
		const sources = this.sources

		if (mounted === undefined || page === undefined || sources === undefined || mounted.painted?.page === page) return false

		const context = { document: mounted.host.ownerDocument, tables: sources.tables }
		let objectsChanged = true

		if (mounted.painted === undefined || mounted.painted.page.index !== page.index) {
			const painted = paintPage(page, context)

			mounted.host.replaceChildren(painted.root)
			mounted.painted = painted
		} else {
			const patched = patchPage(mounted.painted, page, context)

			mounted.painted = patched.painted
			objectsChanged = patched.objectsChanged
		}

		if (objectsChanged) {
			mounted.signal.aborted = true
			mounted.release()

			const signal = { aborted: false }

			mounted.signal = signal
			mounted.release = paintObjects(mounted.painted.objects, {
				document: mounted.host.ownerDocument,
				session: sources.session,
				media: sources.media,
				signal
			})
		}

		return true
	}

	// ---- overlays -----------------------------------------------------------------------------------------

	// Redraws caret, selection, composition and marks on every mounted page.
	draw(): void {
		this.caretPage = undefined

		const caret = this.editor.selection().empty ? this.editor.caret() : undefined

		if (caret !== undefined) this.caretPage = { index: caret.page, box: caret }

		for (const index of this.mounted.keys()) this.drawPage(index)
	}

	private drawPage(index: number): void {
		const mounted = this.mounted.get(index)

		if (mounted === undefined) return

		const document = mounted.layer.ownerDocument
		const editor = this.editor
		const selection = editor.selection()

		// Editing a header, footer or note: the body is shown dimmed.
		const band = editor.activeStory === editor.body.session.story.story.index ? undefined : editor.bodyBand(index)
		const shown = band === undefined ? undefined : editor.page(index)

		if (band !== undefined && shown !== undefined) {
			mounted.under.replaceChildren(
				rectElement(document, "fo-docx-dim", { page: index, x: 0, y: band.top, width: shown.width, height: band.bottom - band.top })
			)
		} else if (mounted.under.firstChild !== null) {
			mounted.under.replaceChildren()
		}

		this.drawMarks(index, mounted)

		const fragment = document.createDocumentFragment()

		if (!selection.empty) {
			for (const rect of editor.selectionRects(index)) fragment.appendChild(rectElement(document, "fo-docx-selection", rect))
		}

		const composition = editor.composition

		if (composition !== undefined && composition.to > composition.from) {
			for (const rect of editor.rangeRects(composition.from, composition.to, index))
				fragment.appendChild(rectElement(document, "fo-docx-composition", rect))
		}

		mounted.over.replaceChildren(fragment)

		const caret = this.caretPage

		if (caret !== undefined && caret.index === index) this.placeCaret(mounted, caret.box, composition !== undefined)
		else if (this.caretElement !== undefined && this.caretElement.parentNode === mounted.caret) this.caretElement.remove()
	}

	private placeCaret(mounted: MountedPage, box: CaretBox, composing: boolean): void {
		let element = this.caretElement

		if (element === undefined) {
			element = mounted.layer.ownerDocument.createElement("div")
			element.className = "fo-docx-caret"
			this.caretElement = element
		}

		if (element.parentNode !== mounted.caret) mounted.caret.appendChild(element)

		const style = `left:${round(box.x)}px;top:${round(box.top)}px;height:${round(box.height)}px;`

		if (style !== this.caretStyle) {
			element.style.cssText = style
			this.caretStyle = style
			// A caret that moved shows at once: its blink starts over, by switching between two equal animations
			// (restarting one through getAnimations() would make the browser style the page there and then).
			this.blinkPhase = !this.blinkPhase
		}

		const blink = this.focused ? (this.blinkPhase ? "b" : "a") : undefined

		if (element.dataset["foBlink"] !== blink) {
			if (blink === undefined) delete element.dataset["foBlink"]
			else element.dataset["foBlink"] = blink
		}

		setFlag(element, "foComposing", composing)
	}

	// The marks within the positions the page shows, all boxes found in one pass over its lines, drawn as a
	// path per line and kind. Only when the page, the document or the marks changed; paths that show what they
	// showed are kept (moved when their line moved), so a keystroke touches the path of the line typed in.
	private drawMarks(index: number, mounted: MountedPage): void {
		const geometry = this.editor.geometry
		const page = geometry.pages.page(index)
		const doc = geometry.story.document
		const story = this.editor.activeStory
		const drawn = mounted.drawn

		if (drawn.page === page && drawn.doc === doc && drawn.marks === this.marks && drawn.story === story) return

		const samePage = drawn.page === page && drawn.story === story

		drawn.page = page
		drawn.doc = doc
		drawn.marks = this.marks
		drawn.story = story

		// The same layout of the page with the same marks on it, only moved (an edit on another page): as drawn.
		if (samePage && this.sameMarksShown(index, mounted.shown)) return

		const marks = this.marks
		// Reused from page to page and keystroke to keystroke (thousands of numbers each).
		const boxes = this.boxesScratch
		const owners = this.ownersScratch

		boxes.length = 0
		owners.length = 0
		geometry.rangesBoxes(marks, this.longestMark, index, boxes, owners)
		this.recordMarksShown(index, mounted.shown)

		// Boxes come line by line: the lines of one top, one per kind.
		const lines: MarkLine[] = []
		let firstOfTop = 0

		for (let k = 0; k < owners.length; k++) {
			const kind = marks[owners[k] ?? -1]?.kind
			const top = boxes[4 * k] ?? 0
			const height = boxes[4 * k + 1] ?? 0

			if (kind === undefined) continue

			const last = lines[lines.length - 1]

			if (last === undefined || last.top !== top || last.height !== height) firstOfTop = lines.length

			let line: MarkLine | undefined

			for (let l = firstOfTop; l < lines.length && line === undefined; l++) if (lines[l]?.kind === kind) line = lines[l]

			if (line === undefined) {
				line = { kind, top, height, boxes: [] }
				lines.push(line)
			}

			line.boxes.push(boxes[4 * k + 2] ?? 0, boxes[4 * k + 3] ?? 0)
		}

		const previous = mounted.rows
		const next = new Map<string, MarkRow[]>()
		const added: { line: MarkLine; content: string }[] = []
		const keep = (content: string, row: MarkRow): void => {
			const list = next.get(content)

			if (list === undefined) next.set(content, [row])
			else list.push(row)
		}

		for (const line of lines) {
			let content = `${line.kind}|${round(line.height)}`

			for (let b = 0; b + 1 < line.boxes.length; b += 2) content += `|${round(line.boxes[b] ?? 0)}:${round(line.boxes[b + 1] ?? 0)}`

			const kept = takeNearest(previous.get(content), line.top)

			if (kept === undefined) {
				added.push({ line, content })

				continue
			}

			if (kept.top !== line.top) {
				kept.top = line.top
				kept.element.setAttribute("transform", `translate(0 ${round(line.top)})`)
			}

			keep(content, kept)
		}

		// Paths no longer drawn are drawn again for the new rows (in the drawing of their kind), the rest removed.
		const spare: MarkRow[] = []

		for (const list of previous.values()) spare.push(...list)

		for (const { line, content } of added) {
			const blend = BLENDED.has(line.kind)
			const drawing = this.drawing(index, mounted, blend)
			let row = spare.pop()

			if (row === undefined) {
				const element = drawing.ownerDocument.createElementNS(SVG, "path")

				element.setAttribute("class", "fo-docx-mark")
				row = { element, content, top: Number.NaN }
			}

			if (row.element.parentNode !== drawing) drawing.appendChild(row.element)
			if (row.element.getAttribute("data-fo-mark") !== line.kind) row.element.setAttribute("data-fo-mark", line.kind)

			row.element.setAttribute("d", markPath(line.kind, line.height, line.boxes))

			if (row.top !== line.top) row.element.setAttribute("transform", `translate(0 ${round(line.top)})`)

			row.content = content
			row.top = line.top
			keep(content, row)
		}

		for (const row of spare) row.element.remove()

		mounted.rows = next
	}

	// The marks that may show on page `index` (those within its positions), each from and to relative to its
	// first position.
	private eachMarkShown(index: number, visit: (from: number, to: number, kind: string) => boolean): void {
		const range = this.editor.geometry.pageRange(index)

		if (range === undefined) return

		const marks = this.marks
		let low = 0
		let high = marks.length

		while (low < high) {
			const middle = (low + high) >> 1

			if ((marks[middle]?.from ?? 0) < range.from - this.longestMark) low = middle + 1
			else high = middle
		}

		for (let k = low; k < marks.length; k++) {
			const mark = marks[k]

			if (mark === undefined || mark.from > range.to) break
			if (mark.to < range.from) continue
			if (!visit(mark.from - range.from, mark.to - range.from, mark.kind)) return
		}
	}

	private sameMarksShown(index: number, shown: { readonly offsets: readonly number[]; readonly kinds: readonly string[] }): boolean {
		let k = 0
		let same = true

		this.eachMarkShown(index, (from, to, kind) => {
			same = shown.offsets[2 * k] === from && shown.offsets[2 * k + 1] === to && shown.kinds[k] === kind
			k++

			return same
		})

		return same && k === shown.kinds.length
	}

	private recordMarksShown(index: number, shown: { offsets: number[]; kinds: string[] }): void {
		shown.offsets.length = 0
		shown.kinds.length = 0
		this.eachMarkShown(index, (from, to, kind) => {
			shown.offsets.push(from, to)
			shown.kinds.push(kind)

			return true
		})
	}

	// The page's drawing for plain or blended marks, made when first needed and sized to the page.
	private drawing(index: number, mounted: MountedPage, blend: boolean): SVGSVGElement {
		const drawings = mounted.drawings
		const shown = this.editor.page(index)
		const width = shown?.width ?? 0
		const height = shown?.height ?? 0
		let drawing = blend ? drawings.blend : drawings.plain

		if (drawing === undefined) {
			drawing = mounted.layer.ownerDocument.createElementNS(SVG, "svg")
			drawing.setAttribute("class", "fo-docx-marks")
			drawing.setAttribute("aria-hidden", "true")

			if (blend) {
				drawing.setAttribute("data-fo-blend", "")
				drawings.blend = drawing
			} else {
				drawings.plain = drawing
			}

			// Plain marks under the highlights.
			if (blend) mounted.marks.appendChild(drawing)
			else mounted.marks.prepend(drawing)
		}

		if (drawing.getAttribute("width") !== String(round(width))) drawing.setAttribute("width", String(round(width)))
		if (drawing.getAttribute("height") !== String(round(height))) drawing.setAttribute("height", String(round(height)))

		return drawing
	}

	// The caret's box on its page as last drawn, or undefined when it is not drawn (a range is selected or
	// its page is not mounted).
	caret(): { readonly index: number; readonly box: CaretBox } | undefined {
		return this.caretPage === undefined || !this.mounted.has(this.caretPage.index) ? undefined : this.caretPage
	}

	isMounted(index: number): boolean {
		return this.mounted.has(index)
	}
}

// Of rows that show the same (repeated text has the same marks), the one nearest `top`: one that did not move
// stays untouched.
function takeNearest(rows: MarkRow[] | undefined, top: number): MarkRow | undefined {
	if (rows === undefined || rows.length === 0) return undefined

	let best = 0

	for (let k = 1; k < rows.length && rows[best]?.top !== top; k++) {
		if (Math.abs((rows[k]?.top ?? 0) - top) < Math.abs((rows[best]?.top ?? 0) - top)) best = k
	}

	return rows.splice(best, 1)[0]
}

function round(value: number): number {
	return Math.round(value * 100) / 100
}

function setFlag(element: HTMLElement, name: string, on: boolean): void {
	if (on === (element.dataset[name] !== undefined)) return

	if (on) element.dataset[name] = ""
	else delete element.dataset[name]
}

function boxStyle(rect: PageRect): string {
	return `left:${round(rect.x)}px;top:${round(rect.y)}px;width:${round(rect.width)}px;height:${round(rect.height)}px;`
}

function rectElement(document: Document, className: string, rect: PageRect): HTMLElement {
	const element = document.createElement("div")

	element.className = className
	element.style.cssText = boxStyle(rect)

	return element
}
