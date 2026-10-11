import { Item, RevisionView } from "../model/kinds"
import { Ns } from "../model/namespaces"
import type { StoryPart } from "../model/part"
import type { ResolvedBorder, ResolvedTableRowProps, TableCellProps, WidthProps } from "../props/types"
import type { CellFormat, TableFormat } from "../styles/resolve"
import { Area, BlockFlow, FlowStatus, nextToken, type Engine } from "./flow"
import { PageGraphic, Layer, OutputMark } from "./output"
import { borderWidth } from "./styles"
import { EPSILON, twips } from "./units"

// Table layout: the grid from w:tblGrid (scaled to a percentage width), rows of cells laid out as block
// flows, vertical merges, row heights, borders with the heavier border winning a shared edge, header rows
// repeated on every page and rows split across pages unless w:cantSplit.

interface Slot {
	readonly item: number
	readonly format: CellFormat
	readonly column: number
	readonly span: number
	readonly x: number
	readonly width: number
	readonly continued: boolean
	// Rows a vertical merge starting here spans (1 otherwise; -1 until counted).
	rows: number
}

interface RowInfo {
	readonly item: number
	readonly table: TableFormat
	readonly props: ResolvedTableRowProps
	readonly slots: readonly Slot[]
}

function widthPx(width: WidthProps | undefined, reference: number): number {
	if (width === undefined) return 0
	if (width.type === "dxa") return twips(width.value)
	if (width.type === "pct") return (width.value / 5000) * reference

	return 0
}

// Border numbers and precedence for resolving two borders on one edge, as Word does it (MS-OI29500 notes on
// ISO/IEC 29500-1 §17.4.66): the heavier border (width in eighths of a point times its number; dotted and
// dashed weigh 1) wins, then the style earlier in the precedence list, then the darker colour.
const BORDER_NUMBERS: Readonly<Record<string, number>> = {
	single: 1,
	thick: 2,
	double: 3,
	dotDash: 8,
	dotDotDash: 9,
	triple: 10,
	thinThickSmallGap: 11,
	thickThinSmallGap: 12,
	thinThickThinSmallGap: 13,
	thinThickMediumGap: 14,
	thickThinMediumGap: 15,
	thinThickThinMediumGap: 16,
	thinThickLargeGap: 17,
	thickThinLargeGap: 18,
	thinThickThinLargeGap: 19,
	wave: 20,
	doubleWave: 21,
	dashSmallGap: 22,
	dashDotStroked: 23,
	threeDEmboss: 24,
	threeDEngrave: 25,
	outset: 26,
	inset: 27
}

const BORDER_PRECEDENCE = [
	"single",
	"thick",
	"double",
	"dotted",
	"dashed",
	"dotDash",
	"dotDotDash",
	"triple",
	"thinThickSmallGap",
	"thickThinSmallGap",
	"thinThickThinSmallGap",
	"thinThickMediumGap",
	"thickThinMediumGap",
	"thinThickThinMediumGap",
	"thinThickLargeGap",
	"thickThinLargeGap",
	"thinThickThinLargeGap",
	"wave",
	"doubleWave",
	"dashSmallGap",
	"dashDotStroked",
	"threeDEmboss",
	"threeDEngrave",
	"outset",
	"inset"
]

function borderWeight(border: ResolvedBorder): number {
	if (border.style === "dotted" || border.style === "dashed") return 1

	return border.size * (BORDER_NUMBERS[border.style] ?? 1)
}

function precedence(style: string): number {
	const at = BORDER_PRECEDENCE.indexOf(style)

	return at < 0 ? BORDER_PRECEDENCE.length : at
}

// Brightness tests in turn: R + B + 2G, then B + 2G, then G (automatic colour is black).
function brightness(color: string | undefined, test: number): number {
	const value = color === undefined ? 0 : Number.parseInt(color, 16)
	const r = (value >> 16) & 255
	const g = (value >> 8) & 255
	const b = value & 255

	return test === 0 ? r + b + 2 * g : test === 1 ? b + 2 * g : g
}

// The border drawn where `a` (first in reading order) and `b` meet.
function heavier(a: ResolvedBorder | undefined, b: ResolvedBorder | undefined): ResolvedBorder | undefined {
	if (a === undefined) return b
	if (b === undefined || a === b) return a

	const weight = borderWeight(b) - borderWeight(a)

	if (weight !== 0) return weight > 0 ? b : a

	const style = precedence(b.style) - precedence(a.style)

	if (style !== 0) return style < 0 ? b : a

	for (let test = 0; test < 3; test++) {
		const difference = brightness(b.color, test) - brightness(a.color, test)

		if (difference !== 0) return difference < 0 ? b : a
	}

	return a
}

export class TableLayout {
	readonly flow: BlockFlow
	readonly engine: Engine
	readonly part: StoryPart
	readonly item: number
	y = 0
	// Body rows placed in the current area (repeated header rows not counted).
	rowsPlaced = 0
	private readonly rows: number[] = []
	private readonly direct: TableFormat
	private readonly grid: number[] = []
	private readonly gridX: number[] = []
	private tableX = 0
	private tableWidth = 0
	private headerRows = 0
	private row = 0
	private continuing = false
	private token = -1
	private started = false
	private breakChecked = false
	private readonly infos = new Map<number, RowInfo>()
	private readonly pool: BlockFlow[] = []
	private readonly headerPool: BlockFlow[] = []
	private readonly cellArea = new Area()
	private readonly mark = new OutputMark()
	private readonly mergeBottoms: number[] = []
	private readonly cells = new Map<number, TableCellProps>()

	constructor(flow: BlockFlow, item: number, area: Area) {
		const engine = flow.engine
		const part = flow.part

		if (part === undefined) throw new Error("A table needs its part")

		this.flow = flow
		this.engine = engine
		this.part = part
		this.item = item

		const tableRow = part.aux[item] ?? -1
		const tblPr = part.tables.tblPr[tableRow] ?? -1

		this.direct = engine.resolver.table(engine.store.table(part.document, tblPr))
		this.collectRows()
		this.computeGrid(area)
	}

	private collectRows(): void {
		const part = this.part
		const view = this.engine.view

		for (let item = this.item + 1; item < part.end(this.item);) {
			const kind = part.kinds[item]

			if (kind === Item.Row) {
				if (view === RevisionView.Markup || !part.hidden(item, view)) this.rows.push(item)

				item = part.end(item)
			} else if (
				kind === Item.Sdt ||
				kind === Item.CustomXml ||
				kind === Item.AlternateContent ||
				kind === Item.Insertion ||
				kind === Item.Deletion
			) {
				item++
			} else {
				item = part.end(item)
			}
		}
	}

	private cellItems(row: number): number[] {
		const part = this.part
		const view = this.engine.view
		const out: number[] = []

		for (let item = row + 1; item < part.end(row);) {
			const kind = part.kinds[item]

			if (kind === Item.Cell) {
				if (view === RevisionView.Markup || !part.hidden(item, view)) out.push(item)

				item = part.end(item)
			} else if (
				kind === Item.Sdt ||
				kind === Item.CustomXml ||
				kind === Item.AlternateContent ||
				kind === Item.Insertion ||
				kind === Item.Deletion
			) {
				item++
			} else {
				item = part.end(item)
			}
		}

		return out
	}

	private computeGrid(area: Area): void {
		const part = this.part
		const document = part.document
		const reader = part.reader
		const grid = this.grid
		const tableRow = part.aux[this.item] ?? -1
		const tblGrid = part.tables.tblGrid[tableRow] ?? -1

		if (tblGrid >= 0) {
			for (let child = document.firstChildElement(tblGrid); child >= 0; child = document.nextSiblingElement(child)) {
				if (document.localName(child) === "gridCol") grid.push(Math.max(0, twips(reader.int(child, Ns.W, "w", 0))))
			}
		}

		const props = this.direct.props
		let columns = 0

		for (const row of this.rows.slice(0, 32)) {
			const rowProps = this.engine.store.row(document, part.rows.trPr[part.aux[row] ?? -1] ?? -1)
			let count = (rowProps.gridBefore ?? 0) + (rowProps.gridAfter ?? 0)

			for (const cell of this.cellItems(row)) count += this.cellProps(cell).gridSpan ?? 1

			columns = Math.max(columns, count)
		}

		while (grid.length < columns) grid.push(0)

		const known = grid.reduce((sum, w) => sum + w, 0)
		const wanted = widthPx(props.width, area.width)
		const zero = grid.filter(w => w <= 0).length

		if (zero > 0) {
			const share = Math.max(0, (wanted > 0 ? wanted : area.width) - known) / zero

			for (let c = 0; c < grid.length; c++) if ((grid[c] ?? 0) <= 0) grid[c] = Math.max(share, 8)
		} else if (
			wanted > 0 &&
			known > 0 &&
			(props.width.type === "pct" || (props.layout === "fixed" && Math.abs(wanted - known) > 1) || wanted > known + 1)
		) {
			const scale = wanted / known

			for (let c = 0; c < grid.length; c++) grid[c] = (grid[c] ?? 0) * scale
		}

		let x = 0

		for (const w of grid) {
			this.gridX.push(x)
			x += w
		}

		this.tableWidth = x

		const firstMargin = twips(props.cellMargins.start.type === "dxa" ? props.cellMargins.start.value : 0)
		let left: number

		if (props.physicalAlignment === "center") left = area.left + (area.width - x) / 2
		else if (props.physicalAlignment === "right") left = area.left + area.width - x
		else {
			left = area.left + widthPx(props.indent, area.width)

			// Before Word 2013 the first cell's text, not its border, lines up with the indent.
			if (this.engine.compat.compatibilityMode < 15) left -= firstMargin
		}

		const position = props.position

		if (position !== undefined && position.horzAnchor !== undefined) {
			const frame = area.frame
			const offset = twips(position.x ?? 0)

			if (position.horzAnchor === "page") left = offset
			else if (position.horzAnchor === "margin") left = frame.marginLeft + offset
			else left = area.left + offset

			if (position.xSpec === "center")
				left =
					(position.horzAnchor === "page" ? 0 : area.left) + ((position.horzAnchor === "page" ? frame.width : area.width) - x) / 2
			else if (position.xSpec === "right") left = (position.horzAnchor === "page" ? frame.width : area.left + area.width) - x
		}

		this.tableX = left
	}

	// A cell's direct properties, parsed once (the grid reads the first rows before they are laid out).
	private cellProps(cell: number): TableCellProps {
		let props = this.cells.get(cell)

		if (props === undefined) {
			const part = this.part

			props = this.engine.store.cell(part.document, part.cells.tcPr[part.aux[cell] ?? -1] ?? -1)
			this.cells.set(cell, props)
		}

		return props
	}

	private info(index: number): RowInfo | undefined {
		let info = this.infos.get(index)

		if (info !== undefined) return info

		const item = this.rows[index]

		if (item === undefined) return undefined

		const engine = this.engine
		const part = this.part
		const document = part.document
		const rowRow = part.aux[item] ?? -1
		const tblPrEx = part.rows.tblPrEx[rowRow] ?? -1
		const table =
			tblPrEx < 0
				? this.direct
				: engine.resolver.table(
						engine.store.table(document, part.tables.tblPr[part.aux[this.item] ?? -1] ?? -1),
						engine.store.table(document, tblPrEx)
					)
		const props = engine.resolver.row(table, engine.store.row(document, part.rows.trPr[rowRow] ?? -1), {
			rowIndex: index,
			rowCount: this.rows.length
		})
		const slots: Slot[] = []
		let column = props.gridBefore
		const bidi = this.direct.props.bidiVisual

		for (const cell of this.cellItems(item)) {
			const direct = this.cellProps(cell)

			this.cells.delete(cell)

			const span = Math.max(1, direct.gridSpan ?? 1)
			const format = engine.resolver.cell(table, direct, {
				rowIndex: index,
				rowCount: this.rows.length,
				gridColumn: column,
				gridSpan: span,
				gridColumnCount: this.grid.length
			})
			const end = Math.min(this.grid.length, column + span)
			const start = this.gridX[Math.min(column, this.gridX.length - 1)] ?? 0
			const width = (this.gridX[end - 1] ?? start) + (this.grid[end - 1] ?? 0) - start

			slots.push({
				item: cell,
				format,
				column,
				span,
				x: bidi ? this.tableX + this.tableWidth - start - width : this.tableX + start,
				width: Math.max(1, width),
				continued: format.props.vMerge === "continue",
				rows: format.props.vMerge === "restart" ? -1 : 1
			})
			column += span
		}

		info = { item, table, props, slots }
		this.infos.set(index, info)

		return info
	}

	// Rows a vertical merge starting at `slot` of row `index` spans, counted once (iteratively: rows are read
	// ahead without counting their own merges).
	private spanOf(index: number, slot: Slot): number {
		if (slot.rows >= 0) return slot.rows

		let rows = 1

		for (let next = index + 1; next < this.rows.length && rows < 4096; next++) {
			const below = this.info(next)?.slots.find(candidate => candidate.column === slot.column)

			if (below === undefined || !below.continued) break

			rows++
		}

		slot.rows = rows

		return rows
	}

	// The row layout continues at.
	get currentRow(): number {
		return this.row
	}

	// Lays out rows from `top` until the table ends or the area is full.
	run(area: Area, top: number): FlowStatus {
		const page = area.page

		if (this.token !== area.token) {
			const fresh = !this.started

			this.token = area.token
			this.rowsPlaced = 0
			this.y = top
			this.started = true

			if (!fresh && this.headerRows > 0 && this.row >= this.headerRows && area.breaks) {
				for (let r = 0; r < this.headerRows; r++) this.layoutRow(area, r, this.headerPool, true, false)
			}
		}

		if (this.row === 0 && !this.continuing && !this.breakChecked) {
			this.breakChecked = true

			while (this.headerRows < this.rows.length && this.info(this.headerRows)?.props.header === true) this.headerRows++

			// A page break before the table's first paragraph breaks before the table.
			if (area.breaks && !area.empty && this.firstParagraphBreaks()) return FlowStatus.PAGE_BREAK

			const position = this.direct.props.position

			if (position !== undefined && (position.vertAnchor === undefined || position.vertAnchor === "text") && (position.y ?? 0) > 0)
				this.y += twips(position.y ?? 0)
		}

		while (this.row < this.rows.length) {
			const info = this.info(this.row)

			if (info === undefined || info.props.hidden) {
				this.row++

				continue
			}

			page.mark(this.mark)

			const status = this.layoutRow(area, this.row, this.pool, false, this.continuing)

			if (status === FlowStatus.SUSPEND) return FlowStatus.SUSPEND

			if (status === FlowStatus.DONE) {
				this.continuing = false
				this.infos.delete(this.row - 64)
				this.row++
				this.rowsPlaced++
				continue
			}

			return FlowStatus.OVERFLOW
		}

		return FlowStatus.DONE
	}

	private cellMargins(slot: Slot): { top: number; bottom: number; start: number; end: number } {
		const margins = slot.format.props.margins
		const px = (width: WidthProps): number => (width.type === "dxa" ? twips(width.value) : 0)

		return { top: px(margins.top), bottom: px(margins.bottom), start: px(margins.start), end: px(margins.end) }
	}

	// Lays out one row (or the rest of a split row). DONE when it ended in the area; OVERFLOW when it moved
	// (nothing kept) or split (continuing); SUSPEND when fonts must load first (nothing kept).
	private layoutRow(area: Area, index: number, pool: BlockFlow[], header: boolean, continuing: boolean): FlowStatus {
		const engine = this.engine
		const info = this.info(index)
		const page = area.page

		if (info === undefined) return FlowStatus.DONE

		const top = this.y
		const props = info.props
		const exact = props.height?.rule === "exact"
		const rowMark = page.mark(new OutputMark())
		const atTop = this.rowsPlaced === 0 && (area.empty || header)

		const footnoteMark = area.footnotes?.mark() ?? 0
		const exclusionMark = area.exclusions?.count ?? 0
		const missingBefore = engine.styles.hasMissing

		if (!continuing) {
			info.slots.forEach((slot, k) => {
				let flow = pool[k]

				if (flow === undefined) {
					flow = new BlockFlow(engine, { ...this.flow.story, breaksAsLines: true }, this.flow.region, false)
					pool[k] = flow
				}

				if (!slot.continued) flow.setContainer(this.part, slot.item, slot.format)
				else flow.setRange(this.part, 0, 0)
			})
		}

		const cellArea = this.cellArea.copy(area)
		const token = nextToken()
		let overflow = false
		let anyPlaced = false
		let height = 0
		const cellMarks: OutputMark[] = []

		cellArea.breaks = false
		cellArea.inTable = true
		cellArea.token = token

		info.slots.forEach((slot, k) => {
			const flow = pool[k]

			cellMarks.push(page.mark(new OutputMark()))

			if (flow === undefined || slot.continued) return

			const margins = this.cellMargins(slot)

			cellArea.left = slot.x + margins.start
			cellArea.width = Math.max(1, slot.width - margins.start - margins.end)
			cellArea.top = top + (continuing ? 0 : margins.top)
			const merged = this.spanOf(index, slot) > 1

			cellArea.bottom = merged || exact ? Infinity : area.bottom - margins.bottom
			cellArea.empty = atTop

			const before = page.text.lineCount
			const status = flow.run(cellArea)

			if (page.text.lineCount > before) anyPlaced = true
			if (status === FlowStatus.OVERFLOW) overflow = true

			const bottom = flow.contentBottom() + margins.bottom

			if (merged) this.mergeBottoms[slot.column] = bottom
			else height = Math.max(height, bottom - top)
		})

		for (const slot of info.slots) {
			if (!slot.continued) continue

			const merged = this.mergeBottoms[slot.column] ?? Number.NaN

			if (!Number.isNaN(merged) && this.lastOfMerge(index, slot.column)) {
				height = Math.max(height, merged - top)
				this.mergeBottoms[slot.column] = Number.NaN
			}
		}

		if (props.height !== undefined) {
			const h = twips(props.height.value)

			if (props.height.rule === "exact") height = h
			else height = Math.max(height, h)
		}

		// A fresh row that needed fonts not loaded yet is laid out again once they are.
		if (!continuing && !missingBefore && this.flow.suspendable && engine.styles.hasMissing) {
			page.rollback(rowMark)
			area.footnotes?.rollback(footnoteMark)

			if (area.exclusions !== null) area.exclusions.count = exclusionMark

			return FlowStatus.SUSPEND
		}

		const fits = !overflow && top + height <= area.bottom + EPSILON

		if (!fits && !header) {
			const splittable = !props.cantSplit && !exact

			if (!atTop && (!splittable || !anyPlaced || !overflow)) {
				page.rollback(rowMark)
				area.footnotes?.rollback(footnoteMark)

				if (area.exclusions !== null) area.exclusions.count = exclusionMark

				return FlowStatus.OVERFLOW
			}

			if (overflow) {
				// Split: this part reaches the area's bottom, the rest continues on the next page.
				height = Math.max(0, area.bottom - top)
				this.drawRow(area, info, index, top, height, true)
				this.y = top + height
				this.continuing = true
				this.rowsPlaced++

				return FlowStatus.OVERFLOW
			}
		}

		this.alignCells(area, info, pool, cellMarks, top, height)

		// A row of exact height cuts what its cells hold beyond it.
		if (exact) {
			const end = page.mark(new OutputMark())

			info.slots.forEach((slot, k) => {
				const from = cellMarks[k]
				const to = cellMarks[k + 1] ?? end

				if (from !== undefined && !slot.continued)
					page.addClip(slot.x, top, slot.width, height, from.lines, to.lines, from.objects, to.objects)
			})
		}

		this.drawRow(area, info, index, top, height, false)
		this.y = top + height

		return FlowStatus.DONE
	}

	private firstParagraphBreaks(): boolean {
		const slot = this.info(0)?.slots[0]
		const part = this.part

		if (slot === undefined) return false

		for (let item = slot.item + 1; item < part.end(slot.item); item++) {
			if (part.kinds[item] === Item.Paragraph) return this.engine.content.formatOf(part, item, slot.format).props.pageBreakBefore
		}

		return false
	}

	private lastOfMerge(index: number, column: number): boolean {
		const next = this.info(index + 1)
		const slot = next?.slots.find(candidate => candidate.column === column)

		return slot === undefined || !slot.continued
	}

	// Cell vertical alignment once the row's height is known.
	private alignCells(area: Area, info: RowInfo, pool: BlockFlow[], marks: OutputMark[], top: number, height: number): void {
		const page = area.page

		info.slots.forEach((slot, k) => {
			const align = slot.format.props.vAlign
			const flow = pool[k]
			const from = marks[k]

			if ((align !== "center" && align !== "bottom") || flow === undefined || from === undefined || slot.continued || slot.rows !== 1)
				return

			const margins = this.cellMargins(slot)
			const free = top + height - margins.bottom - flow.contentBottom()

			if (free <= EPSILON) return

			const to = marks[k + 1] ?? null

			page.move(from, to, 0, align === "center" ? free / 2 : free)
		})
	}

	private drawRow(area: Area, info: RowInfo, index: number, top: number, height: number, split: boolean): void {
		const page = area.page
		const decorations = this.engine.decorations
		const tableBorders = info.table.props.borders
		const first = index === 0 || this.rowsPlaced === 0
		const last = index === this.rows.length - 1 || split
		const next = last ? undefined : this.info(index + 1)
		const tableShading = decorations.shading(info.table.props.shading)
		const line = (border: ResolvedBorder | undefined, x: number, y: number, horizontal: boolean, length: number): void => {
			if (border === undefined || length <= 0) return

			const width = borderWidth(border)

			page.addGraphic(
				PageGraphic.BORDER,
				decorations.border(border),
				Layer.TEXT,
				horizontal ? x : x - width / 2,
				horizontal ? y - width / 2 : y,
				horizontal ? length : width,
				horizontal ? width : length
			)
		}

		info.slots.forEach((slot, k) => {
			const borders = slot.format.props.borders
			const shading = decorations.shading(slot.format.props.shading)

			if (shading !== 0 || tableShading !== 0)
				page.addGraphic(PageGraphic.SHADING, shading !== 0 ? shading : tableShading, Layer.BEHIND, slot.x, top, slot.width, height)

			if (first && !slot.continued)
				line(
					index === 0 ? (borders.top ?? tableBorders.top) : (borders.top ?? tableBorders.insideH ?? tableBorders.top),
					slot.x,
					top,
					true,
					slot.width
				)

			const below = next?.slots.find(candidate => candidate.column === slot.column)

			if (below === undefined || !below.continued) {
				const own = last ? (borders.bottom ?? tableBorders.bottom) : (borders.bottom ?? tableBorders.insideH)

				line(below === undefined ? own : heavier(own, below.format.props.borders.top), slot.x, top + height, true, slot.width)
			}

			const previous = info.slots[k - 1]
			const start =
				k === 0
					? (borders.start ?? tableBorders.start)
					: heavier(borders.start ?? tableBorders.insideV, previous?.format.props.borders.end ?? tableBorders.insideV)

			line(start, slot.x, top, false, height)

			if (k === info.slots.length - 1) line(borders.end ?? tableBorders.end, slot.x + slot.width, top, false, height)
		})
	}
}
