import { TILE_COLS, TILE_ROWS, type SheetId } from "../../core/sheet/protocol"
import type { SheetGridModel } from "./model"
import { BorderPainter } from "./paint/borders"
import type { GridCanvas } from "./paint/canvas"
import { paintCells, TEXT_PADDING } from "./paint/cells"
import { paintChrome } from "./paint/chrome"
import { canvasSwatches, PatternCache } from "./paint/patterns"
import { TextMeasurer, textFont } from "./paint/text"
import { domExtent, toDom, toVirtual } from "./scroll"
import type { GridPoint } from "./selection"
import { readTheme } from "./readTheme"
import { DEFAULT_THEME, type SheetGridTheme } from "./theme"
import { bodyTotal, gridView, hitAxis, revealOffset, screenOf, edgeAt, type GridView } from "./view"

// Draws a grid on two canvases (cells; headers and selection over them) at most once a frame, sized to the
// scroller's viewport. The scroller holds only an empty extent element, so the browser scrolls nothing it has to
// paint; each scroll event maps the scroll position to the sheet's (scroll.ts) and asks for a frame. The frame
// fetches the tiles it lacks (they paint in a later frame), positions the editor and the accessible active cell,
// and grows the extent when the user scrolled to its end.
//
// A right-to-left sheet (SheetLayout.rightToLeft) is laid out as any other and shown mirrored: the painters
// mirror what they draw, hit-testing mirrors the pointer, and the scroller's horizontal offset runs from its end,
// so the scrollbar starts at the right and dragging it left brings in the columns past the left edge.
//
// Canvas rather than rows of memoised cell elements: both cost about the same for a step of a few rows, but a step
// that replaces the window (a page, a jump, a fast fling) makes the DOM build and lay out every cell again, several
// times the canvas's draw in Chromium (and past a frame on a slow CPU); and a canvas draws every border style,
// diagonals and turned text, which CSS borders cannot.

export interface GridElements {
	readonly root: HTMLElement
	readonly scroller: HTMLElement
	readonly extent: HTMLElement
	readonly cells: HTMLCanvasElement
	readonly chrome: HTMLCanvasElement
	// The accessible active cell, kept over the real one.
	readonly active: HTMLElement
}

export type GridHit =
	| { readonly kind: "cell"; readonly row: number; readonly col: number }
	| { readonly kind: "colHeader"; readonly col: number; readonly edge: number | null }
	| { readonly kind: "rowHeader"; readonly row: number; readonly edge: number | null }
	| { readonly kind: "corner" }

interface Scroll {
	top: number
	left: number
	// The browser's offsets the virtual ones were last set from, to tell our own scrolling from the user's.
	domTop: number
	domLeft: number
}

// Tile rows and columns fetched past the edges of what is on screen.
const MARGIN_ROWS = 16
const MARGIN_COLS = 8
// Rows (columns) left before the extent's end that make it grow.
const GROW_SLACK = 8

export class GridRenderer {
	private elements: GridElements | null = null
	private cells: GridCanvas | null = null
	private chrome: GridCanvas | null = null
	private measurer: TextMeasurer | null = null
	private borders: BorderPainter | null = null
	private readonly patterns = new PatternCache(canvasSwatches())
	private theme: SheetGridTheme = DEFAULT_THEME
	// The variables changed (or may have): read again when next used. An app moves its theme class in an effect of
	// its own, which runs after the grid's, so they are read at the frame, not at once.
	private themeStale = false
	private frame = 0
	private width = 0
	private height = 0
	private pixelRatio = 1
	private readonly scroll: Scroll = { top: 0, left: 0, domTop: 0, domLeft: 0 }
	private sheet: SheetId | null = null
	// Each sheet's scroll position when the user left it.
	private readonly scrolls = new Map<SheetId, { top: number; left: number }>()
	private view: GridView | null = null
	private extentWidth = -1
	private extentHeight = -1
	private readonly detach: (() => void)[] = []
	// The editor, placed over its cell on every frame while it is open.
	private editor: HTMLElement | null = null
	// A resize's guide: the edge after `index`.
	private guide: { axis: "rows" | "cols"; index: number } | null = null
	// Frame costs (ms), for the perf harness.
	readonly costs: number[] = []
	measure = false

	private readonly model: SheetGridModel

	constructor(model: SheetGridModel) {
		this.model = model
	}

	attach(elements: GridElements): void {
		this.dispose()
		this.elements = elements
		this.cells = elements.cells.getContext("2d")
		this.chrome = elements.chrome.getContext("2d")

		if (this.cells !== null) {
			this.measurer = new TextMeasurer(this.cells)
			this.borders = new BorderPainter(this.cells)
		}

		this.theme = readTheme(elements.root)
		this.themeStale = false

		const onScroll = (): void => {
			this.readScroll()
			this.schedule()
		}
		const resize = new ResizeObserver(() => {
			this.readSize()
			this.schedule()
		})

		elements.scroller.addEventListener("scroll", onScroll, { passive: true })
		resize.observe(elements.scroller)
		this.detach.push(
			() => elements.scroller.removeEventListener("scroll", onScroll),
			() => resize.disconnect(),
			this.model.onPaint(() => this.schedule())
		)
		// A theme that follows the system's scheme changes the variables without the app saying so.
		if (typeof globalThis.matchMedia === "function") {
			const scheme = globalThis.matchMedia("(prefers-color-scheme: dark)")
			const onScheme = (): void => this.refreshTheme()

			scheme.addEventListener("change", onScheme)
			this.detach.push(() => scheme.removeEventListener("change", onScheme))
		}

		this.readSize()
		this.restoreScroll()
		this.paint()
	}

	dispose(): void {
		for (const detach of this.detach.splice(0)) detach()

		if (this.frame !== 0) cancelAnimationFrame(this.frame)

		this.frame = 0
		this.elements = null
	}

	refreshTheme(): void {
		if (this.elements === null) return

		this.themeStale = true
		this.schedule()
	}

	private currentTheme(): SheetGridTheme {
		if (this.themeStale && this.elements !== null) {
			this.themeStale = false
			this.theme = readTheme(this.elements.root)
			this.measurer?.reset()
		}

		return this.theme
	}

	// Placed at once, so it never shows a frame away from its cell.
	setEditor(editor: HTMLElement | null): void {
		this.editor = editor

		const view = this.view

		if (view !== null && editor !== null) this.place(view, this.model.getSnapshot().selection.active)
	}

	setGuide(guide: { axis: "rows" | "cols"; index: number } | null): void {
		this.guide = guide
		this.schedule()
	}

	schedule(): void {
		if (this.frame === 0 && this.elements !== null) this.frame = requestAnimationFrame(() => this.paint())
	}

	private readSize(): void {
		const scroller = this.elements?.scroller

		if (scroller === undefined) return

		this.width = scroller.clientWidth
		this.height = scroller.clientHeight
	}

	// The current sheet is shown right to left.
	private mirrored(): boolean {
		return this.model.layout(this.model.getSnapshot().sheet)?.rightToLeft === true
	}

	// The scroller's horizontal offset as laid out: from its end in a mirrored sheet.
	private laidOutLeft(scroller: HTMLElement): number {
		return this.mirrored() ? Math.max(0, scroller.scrollWidth - scroller.clientWidth - scroller.scrollLeft) : scroller.scrollLeft
	}

	// Keeps the virtual offsets when the browser's are what we set them to (rounding would move them otherwise).
	private readScroll(): void {
		const scroller = this.elements?.scroller

		if (scroller === undefined || this.view === null) return

		const top = scroller.scrollTop
		const left = this.laidOutLeft(scroller)

		if (Math.abs(top - this.scroll.domTop) >= 1) {
			this.scroll.top = toVirtual(top, bodyTotal(this.view.rows), this.view.rows.bodyView)
			this.scroll.domTop = top
		}

		if (Math.abs(left - this.scroll.domLeft) >= 1) {
			this.scroll.left = toVirtual(left, bodyTotal(this.view.cols), this.view.cols.bodyView)
			this.scroll.domLeft = left
		}
	}

	// A sheet opens where the user left it in this session, else where its file did: the scrolled pane's first cell
	// at the top left.
	private restoreScroll(): void {
		const sheet = this.model.getSnapshot().sheet
		const layout = this.model.layout(sheet)
		const axes = this.model.axes(sheet)

		if (layout === null || axes === null) return

		if (this.sheet !== null) this.scrolls.set(this.sheet, { top: this.scroll.top, left: this.scroll.left })

		this.sheet = sheet

		const kept = this.scrolls.get(sheet)
		const pane = layout.pane
		// The scrolled part starts past the frozen block.
		const bodyRow = pane === null ? 0 : pane.rows > 0 ? pane.firstRow + pane.rows : 0
		const bodyCol = pane === null ? 0 : pane.cols > 0 ? pane.firstCol + pane.cols : 0

		this.scroll.top = kept?.top ?? (pane === null ? 0 : Math.max(0, axes.rows.offset(pane.topRow) - axes.rows.offset(bodyRow)))
		this.scroll.left = kept?.left ?? (pane === null ? 0 : Math.max(0, axes.cols.offset(pane.leftCol) - axes.cols.offset(bodyCol)))
		this.view = this.computeView()
		this.writeScroll()
	}

	private computeView(): GridView | null {
		const sheet = this.model.getSnapshot().sheet
		const layout = this.model.layout(sheet)
		const axes = this.model.axes(sheet)

		if (layout === null || axes === null) return null

		return gridView(axes, layout, this.scroll, this.width, this.height)
	}

	// The extent element sized for the view, then the browser's offsets set from the virtual ones.
	private writeScroll(): void {
		const elements = this.elements
		const view = this.view

		if (elements === null || view === null) return

		this.sizeExtent(view)

		const top = toDom(view.rows.scroll, bodyTotal(view.rows), view.rows.bodyView)
		const left = toDom(view.cols.scroll, bodyTotal(view.cols), view.cols.bodyView)

		this.scroll.top = view.rows.scroll
		this.scroll.left = view.cols.scroll
		elements.scroller.scrollTop = top
		elements.scroller.scrollLeft = this.mirrored()
			? Math.max(0, elements.scroller.scrollWidth - elements.scroller.clientWidth - left)
			: left
		this.scroll.domTop = elements.scroller.scrollTop
		this.scroll.domLeft = this.laidOutLeft(elements.scroller)
	}

	// Whether the extent's size changed.
	private sizeExtent(view: GridView): boolean {
		const extent = this.elements?.extent

		if (extent === undefined) return false

		const width = Math.ceil(view.cols.header + view.cols.frozenSize + domExtent(bodyTotal(view.cols)))
		const height = Math.ceil(view.rows.header + view.rows.frozenSize + domExtent(bodyTotal(view.rows)))
		const changed = width !== this.extentWidth || height !== this.extentHeight

		if (width !== this.extentWidth) {
			extent.style.width = `${String(width)}px`
			this.extentWidth = width
		}

		if (height !== this.extentHeight) {
			extent.style.height = `${String(height)}px`
			this.extentHeight = height
		}

		return changed
	}

	currentView(): GridView | null {
		return this.view
	}

	paint(): void {
		this.frame = 0

		const elements = this.elements

		if (elements === null) return

		const started = this.measure ? performance.now() : 0
		const snapshot = this.model.getSnapshot()

		if (snapshot.sheet !== this.sheet) this.restoreScroll()

		const layout = this.model.layout(snapshot.sheet)
		const axes = this.model.axes(snapshot.sheet)
		const theme = this.currentTheme()
		const pixelRatio = Math.max(1, Math.min(4, globalThis.devicePixelRatio || 1))

		this.resizeCanvases(pixelRatio)

		if (layout === null || axes === null) {
			this.blank()

			return
		}

		const view = gridView(axes, layout, this.scroll, this.width, this.height)

		this.view = view

		// A grown or shrunk extent changes what a scaled browser offset means: the browser's follows the sheet's.
		if (this.sizeExtent(view)) this.writeScroll()

		this.requestTiles(snapshot.sheet, view)

		if (this.cells !== null && this.measurer !== null && this.borders !== null) {
			paintCells(
				{ canvas: this.cells, measurer: this.measurer, borders: this.borders, patterns: this.patterns },
				{ model: this.model, sheet: snapshot.sheet, layout, view, theme, scale: axes.scale, pixelRatio }
			)
		}

		if (this.chrome !== null) {
			paintChrome(this.chrome, {
				view,
				theme,
				pixelRatio,
				rightToLeft: layout.rightToLeft,
				selection: snapshot.selection,
				merges: this.model.merges(layout),
				guide:
					this.guide === null
						? null
						: {
								axis: this.guide.axis,
								at:
									screenOf(this.guide.axis === "rows" ? view.rows : view.cols, this.guide.index) +
									(this.guide.axis === "rows" ? view.rows : view.cols).axis.size(this.guide.index)
							}
			})
		}

		this.place(view, snapshot.selection.active)
		this.growIfAtEnd(snapshot.sheet, view)

		if (this.measure) this.costs.push(performance.now() - started)
	}

	private blank(): void {
		const theme = this.currentTheme()

		for (const canvas of [this.cells, this.chrome]) {
			if (canvas === null) continue

			canvas.setTransform(this.pixelRatio, 0, 0, this.pixelRatio, 0, 0)
			canvas.fillStyle = canvas === this.cells ? theme.background : "rgba(0,0,0,0)"
			canvas.fillRect(0, 0, this.width, this.height)
		}
	}

	private resizeCanvases(pixelRatio: number): void {
		const elements = this.elements

		if (elements === null) return

		const width = Math.max(1, Math.round(this.width * pixelRatio))
		const height = Math.max(1, Math.round(this.height * pixelRatio))

		if (elements.cells.width === width && elements.cells.height === height && this.pixelRatio === pixelRatio) return

		this.pixelRatio = pixelRatio

		for (const canvas of [elements.cells, elements.chrome]) {
			canvas.width = width
			canvas.height = height
			canvas.style.width = `${String(this.width)}px`
			canvas.style.height = `${String(this.height)}px`
		}

		// A resized canvas loses its state.
		this.measurer?.reset()
		this.borders?.reset()
	}

	// The tiles under the view (frozen panes included) plus a margin, as tile row and column pairs.
	private requestTiles(sheet: SheetId, view: GridView): void {
		const rowSpans: [number, number][] = []
		const colSpans: [number, number][] = []

		if (view.rows.frozen > 0) rowSpans.push([view.rows.frozenFirst, view.rows.frozenEnd - 1])
		if (view.rows.last >= view.rows.first)
			rowSpans.push([Math.max(view.rows.frozenEnd, view.rows.first - MARGIN_ROWS), view.rows.last + MARGIN_ROWS])
		if (view.cols.frozen > 0) colSpans.push([view.cols.frozenFirst, view.cols.frozenEnd - 1])
		if (view.cols.last >= view.cols.first)
			colSpans.push([Math.max(view.cols.frozenEnd, view.cols.first - MARGIN_COLS), view.cols.last + MARGIN_COLS])

		const tileRows = new Set<number>()
		const tileCols = new Set<number>()

		for (const [first, last] of rowSpans)
			for (let tile = Math.floor(first / TILE_ROWS); tile <= Math.floor(Math.min(view.rows.axis.count - 1, last) / TILE_ROWS); tile++)
				tileRows.add(tile)
		for (const [first, last] of colSpans)
			for (let tile = Math.floor(first / TILE_COLS); tile <= Math.floor(Math.min(view.cols.axis.count - 1, last) / TILE_COLS); tile++)
				tileCols.add(tile)

		const coordinates: number[] = []

		for (const tileRow of tileRows) for (const tileCol of tileCols) coordinates.push(tileRow, tileCol)

		this.model.need(sheet, coordinates)
	}

	// The accessible active cell and the editor over the active cell.
	private place(view: GridView, active: GridPoint): void {
		const elements = this.elements

		if (elements === null) return

		const box = this.cellBox(view, active)
		const a11y = elements.active
		const mirrored = this.mirrored()
		// Where a box laid out from x starts on screen.
		const screenX = (x: number, width: number): number => (mirrored ? view.width - x - width : x)

		if (box === null) {
			a11y.style.transform = "translate(-9999px, 0)"
		} else {
			a11y.style.transform = `translate(${String(elements.scroller.scrollLeft + screenX(box.x, box.width))}px, ${String(this.scroll.domTop + box.y)}px)`
			a11y.style.width = `${String(box.width)}px`
			a11y.style.height = `${String(box.height)}px`
		}

		const editor = this.editor
		const editing = this.model.getSnapshot().editing

		if (editor === null || editing === null) return

		const editBox = this.cellBox(view, { row: editing.row, col: editing.col })

		if (editBox === null) {
			editor.style.transform = "translate(-9999px, 0)"

			return
		}

		const layout = this.model.layout(editing.sheet)
		const found = this.model.entry(editing.sheet, editing.row, editing.col)
		const style = this.model.style(found === null ? 0 : (found.tile.styles[found.entry] ?? 0))
		const scale = layout === null ? 1 : this.model.scale(layout)

		const editX = screenX(editBox.x, editBox.width)

		// Mirrored, the editor keeps its cell's right edge and grows to the left.
		editor.style.transform = mirrored
			? `translate(calc(${String(editX + editBox.width)}px - 100%), ${String(editBox.y)}px)`
			: `translate(${String(editX)}px, ${String(editBox.y)}px)`
		editor.style.minWidth = `${String(editBox.width)}px`
		editor.style.minHeight = `${String(editBox.height)}px`
		editor.style.maxWidth = `${String(Math.max(editBox.width, mirrored ? editX + editBox.width : view.width - editX))}px`
		editor.style.font = textFont(style, this.model.style(0), this.currentTheme().font, scale).css
		editor.style.paddingLeft = `${String(Math.max(1, Math.round(TEXT_PADDING * scale)))}px`
	}

	// A cell's (or its merge's) box on screen, or null when it is not on screen.
	cellBox(view: GridView, point: GridPoint): { x: number; y: number; width: number; height: number } | null {
		const layout = this.model.layout(this.model.getSnapshot().sheet)
		const merge =
			layout === null
				? null
				: this.model
						.merges(layout)
						.find(
							range =>
								point.row >= range.startRow &&
								point.row <= range.endRow &&
								point.col >= range.startCol &&
								point.col <= range.endCol
						)
		const startRow = merge?.startRow ?? point.row
		const startCol = merge?.startCol ?? point.col
		const endRow = merge?.endRow ?? point.row
		const endCol = merge?.endCol ?? point.col
		const x = screenOf(view.cols, startCol)
		const y = screenOf(view.rows, startRow)
		const width = screenOf(view.cols, endCol) + view.cols.axis.size(endCol) - x
		const height = screenOf(view.rows, endRow) + view.rows.axis.size(endRow) - y

		if (
			width <= 0 ||
			height <= 0 ||
			x + width <= view.cols.header ||
			y + height <= view.rows.header ||
			x >= view.width ||
			y >= view.height
		)
			return null

		return { x, y, width, height }
	}

	private growIfAtEnd(sheet: SheetId, view: GridView): void {
		const rowsLeft = view.rows.axis.count - 1 - view.rows.last
		const colsLeft = view.cols.axis.count - 1 - view.cols.last

		if (view.rows.last >= view.rows.first && rowsLeft < GROW_SLACK) this.model.grow(sheet, view.rows.last + GROW_SLACK, 0)
		if (view.cols.last >= view.cols.first && colsLeft < GROW_SLACK) this.model.grow(sheet, 0, view.cols.last + GROW_SLACK)
	}

	// --- Hit-testing ---

	hit(clientX: number, clientY: number, slop: number): GridHit | null {
		const elements = this.elements
		const view = this.view

		if (elements === null || view === null) return null

		const rect = elements.scroller.getBoundingClientRect()
		const screenX = clientX - rect.left
		const y = clientY - rect.top

		// The scrollbars.
		if (screenX >= this.width || y >= this.height) return null

		const x = this.layoutX(screenX)

		const row = hitAxis(view.rows, y)
		const col = hitAxis(view.cols, x)

		if (row === null || col === null) return null
		if (row.part === "header" && col.part === "header") return { kind: "corner" }
		if (row.part === "header" && col.part !== "header") return { kind: "colHeader", col: col.index, edge: edgeAt(view.cols, x, slop) }
		if (col.part === "header" && row.part !== "header") return { kind: "rowHeader", row: row.index, edge: edgeAt(view.rows, y, slop) }
		if (row.part === "header" || col.part === "header") return null

		return { kind: "cell", row: row.index, col: col.index }
	}

	// A screen x in the scroller as the grid is laid out (mirrored in a right-to-left sheet).
	layoutX(x: number): number {
		return this.mirrored() ? this.width - x : x
	}

	// Where a pointer is, in the scroller's own coordinates (on screen: layoutX gives the laid-out x).
	local(clientX: number, clientY: number): { x: number; y: number } | null {
		const scroller = this.elements?.scroller

		if (scroller === undefined) return null

		const rect = scroller.getBoundingClientRect()

		return { x: clientX - rect.left, y: clientY - rect.top }
	}

	// --- Scrolling ---

	// Scrolls the least that brings a cell fully into view (frozen rows and columns are always in view).
	reveal(point: GridPoint): void {
		const sheet = this.model.getSnapshot().sheet

		this.model.grow(sheet, point.row, point.col)
		this.view = this.computeView()

		const view = this.view

		if (view === null) return

		const top = revealOffset(view.rows, point.row)
		const left = revealOffset(view.cols, point.col)

		if (top === view.rows.scroll && left === view.cols.scroll) return

		this.scroll.top = top
		this.scroll.left = left
		this.view = this.computeView()
		this.writeScroll()
		this.schedule()
	}

	// Scrolls by pixels (dragging a selection past the edge).
	scrollBy(dx: number, dy: number): void {
		const view = this.view

		if (view === null || (dx === 0 && dy === 0)) return

		this.scroll.top = Math.max(0, view.rows.scroll + dy)
		this.scroll.left = Math.max(0, view.cols.scroll + dx)
		this.view = this.computeView()
		this.writeScroll()
		this.schedule()
	}

	// Rows a page holds (PageUp, PageDown).
	pageRows(): number {
		const view = this.view

		return view === null ? 20 : Math.max(1, view.rows.last - view.rows.first)
	}

	// Columns a screen holds (Alt+Page Up, Alt+Page Down).
	pageCols(): number {
		const view = this.view

		return view === null ? 8 : Math.max(1, view.cols.last - view.cols.first)
	}

	// A screen position along an axis's edge, for a resize's guide: where `index` ends.
	edgePosition(axis: "rows" | "cols", index: number): number | null {
		const view = this.view

		if (view === null) return null

		const own = axis === "rows" ? view.rows : view.cols

		return screenOf(own, index) + own.axis.size(index)
	}

	// Measures text as the grid draws it (autofit).
	measurerFor(): TextMeasurer | null {
		return this.measurer
	}

	themeFont(): string {
		return this.currentTheme().font
	}
}
