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
	painted: PaintedPage | undefined
	release: () => void
	signal: { aborted: boolean }
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
	// Sorted by `from`; the longest's length bounds where a page's marks start.
	private marks: readonly DocxEditorMark[] = []
	private longestMark = 0
	private focused = false
	private caretPage: { index: number; box: CaretBox } | undefined = undefined

	constructor(editor: PagedEditor) {
		this.editor = editor
	}

	mount(index: number, host: HTMLElement, layer: HTMLElement): void {
		this.unmount(index)
		this.mounted.set(index, { host, layer, painted: undefined, release: () => undefined, signal: { aborted: false } })
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
		const fragment = document.createDocumentFragment()
		const editor = this.editor
		const selection = editor.selection()

		// Editing a header, footer or note: the body is shown dimmed.
		if (editor.activeStory !== editor.body.session.story.story.index) {
			const band = editor.bodyBand(index)
			const page = editor.page(index)

			if (band !== undefined && page !== undefined) {
				fragment.appendChild(
					rectElement(document, "fo-docx-dim", {
						page: index,
						x: 0,
						y: band.top,
						width: page.width,
						height: band.bottom - band.top
					})
				)
			}
		}

		// Only the marks within the positions the page shows.
		const marks = this.marks
		const range = marks.length === 0 ? undefined : editor.geometry.pageRange(index)

		if (range !== undefined) {
			for (let k = firstFrom(marks, range.from - this.longestMark); k < marks.length; k++) {
				const mark = marks[k]

				if (mark === undefined || mark.from > range.to) break
				if (mark.to < range.from) continue

				for (const rect of editor.rangeRects(mark.from, mark.to, index)) {
					const box = rectElement(document, "fo-docx-mark", rect)

					box.dataset["foMark"] = mark.kind
					fragment.appendChild(box)
				}
			}
		}

		if (!selection.empty) {
			for (const rect of editor.selectionRects(index)) fragment.appendChild(rectElement(document, "fo-docx-selection", rect))
		}

		const composition = editor.composition

		if (composition !== undefined && composition.to > composition.from) {
			for (const rect of editor.rangeRects(composition.from, composition.to, index))
				fragment.appendChild(rectElement(document, "fo-docx-composition", rect))
		}

		const caret = this.caretPage

		if (caret !== undefined && caret.index === index) {
			const element = document.createElement("div")

			element.className = "fo-docx-caret"
			element.style.cssText = `left:${round(caret.box.x)}px;top:${round(caret.box.top)}px;height:${round(caret.box.height)}px;`

			if (this.focused) element.dataset["foBlink"] = ""
			if (composition !== undefined) element.dataset["foComposing"] = ""

			fragment.appendChild(element)
		}

		mounted.layer.replaceChildren(fragment)
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

function round(value: number): number {
	return Math.round(value * 100) / 100
}

function rectElement(document: Document, className: string, rect: PageRect): HTMLElement {
	const element = document.createElement("div")

	element.className = className
	element.style.cssText = `left:${round(rect.x)}px;top:${round(rect.y)}px;width:${round(rect.width)}px;height:${round(rect.height)}px;`

	return element
}

// The first mark whose `from` is at least `pos` (marks sorted by `from`).
function firstFrom(marks: readonly DocxEditorMark[], pos: number): number {
	let low = 0
	let high = marks.length

	while (low < high) {
		const middle = (low + high) >> 1

		if ((marks[middle]?.from ?? 0) < pos) low = middle + 1
		else high = middle
	}

	return low
}
