import { HyperFormula } from "hyperformula"
import { MAX_COLUMNS, MAX_ROWS } from "../model/model"

// Whole-column (A:A) and whole-row (1:1) references, as far as the engine reads them. HyperFormula ends one at the
// sheet's height (or width), so a single cell written in the grid's last row makes every whole-column reference of
// that sheet read a million cells at each recalculation, whichever column the cell is in. Here they end one row past
// the last cell of the columns (rows) that the sheet's whole-column (whole-row) references cover, one end for all of
// them and for the sheets a formula pairs with it, so paired ranges keep one size (SUMIFS(A:A,B:B,…),
// SUMIFS(One!B:B,Two!A:A,…)). Past that end every covered column is empty, so what a reference sums, counts or finds
// is unchanged; what reads Excel's full 1,048,576 rows (ROWS, blank counts, the last match) adds the empty rest
// itself (functions/plugin.ts wholeTail).

type Graph = HyperFormula["dependencyGraph"]

interface Address {
	readonly row: number
	readonly col: number
}

interface WholeRange {
	readonly sheet: number
	readonly start: Address
	readonly end: Address
}

// A sheet's covered columns (or rows) and where their last cell is, along one axis.
interface Axis {
	// Covered indices, ascending, as [first, last] pairs.
	spans: [number, number][]
	// One past the last cell of the covered indices; -1 until read.
	end: number
}

interface SheetState {
	columns: Axis
	rows: Axis
}

function covers(spans: readonly [number, number][], first: number, last: number): boolean {
	for (const [from, to] of spans) if (from <= first && last <= to) return true

	return false
}

function cover(spans: [number, number][], first: number, last: number): [number, number][] {
	const merged: [number, number][] = []
	let next: [number, number] = [first, last]

	for (const span of [...spans].sort((a, b) => a[0] - b[0])) {
		if (span[1] + 1 < next[0] || next[1] + 1 < span[0]) merged.push(span)
		else next = [Math.min(span[0], next[0]), Math.max(span[1], next[1])]
	}

	merged.push(next)

	return merged.sort((a, b) => a[0] - b[0])
}

// The engine's cell storage of one sheet (AddressMapping's strategies), read for where the last cell is: sparse
// keeps a map of rows per column, dense an array of columns per row.
interface Storage {
	readonly mapping: unknown
	getHeight(): number
}

function lastRowIn(storage: Storage, spans: readonly [number, number][]): number {
	if (spans.length === 0) return 0

	const mapping = storage.mapping
	let last = -1

	if (mapping instanceof Map) {
		for (const [first, end] of spans) {
			for (let col = first; col <= end; col++) {
				const rows: unknown = mapping.get(col)

				if (!(rows instanceof Map)) continue

				for (const row of rows.keys() as Iterable<number>) if (row > last) last = row
			}
		}

		return last + 1
	}

	if (!Array.isArray(mapping)) return storage.getHeight()

	for (let row = Math.min(mapping.length, storage.getHeight()) - 1; row >= 0; row--) {
		const line: unknown = mapping[row]

		if (!Array.isArray(line)) continue

		for (const [first, end] of spans) {
			for (let col = first; col <= Math.min(end, line.length - 1); col++) if (line[col] !== undefined) return row + 1
		}
	}

	return 0
}

function lastColumnIn(storage: Storage, spans: readonly [number, number][]): number {
	if (spans.length === 0) return 0

	const mapping = storage.mapping
	let last = -1

	if (mapping instanceof Map) {
		for (const [col, rows] of mapping as Map<number, unknown>) {
			if (col <= last || !(rows instanceof Map)) continue

			for (const [first, end] of spans) {
				for (let row = first; row <= end && col > last; row++) if (rows.has(row)) last = col
			}
		}

		return last + 1
	}

	if (!Array.isArray(mapping)) return Number.POSITIVE_INFINITY

	for (const [first, end] of spans) {
		for (let row = first; row <= Math.min(end, mapping.length - 1); row++) {
			const line: unknown = mapping[row]

			if (!Array.isArray(line)) continue

			for (let col = line.length - 1; col > last; col--) {
				if (line[col] !== undefined) {
					last = col
					break
				}
			}
		}
	}

	return last + 1
}

// One engine's whole-range ends, by engine sheet id.
export class WholeRanges {
	private graph: Graph | null = null
	private readonly sheets = new Map<number, SheetState>()
	// Each sheet's group's end (one past its last covered cell), by engine sheet id, as read last; cleared on a change.
	private readonly columnEnds = new Map<number, number>()
	private readonly rowEnds = new Map<number, number>()
	// Sheets whose whole columns (rows) one formula pairs (SUMIFS(One!B:B,Two!A:A,…)), by upper-case name: a union of
	// them, by each one's parent. Their ranges end in one place, so the pairs keep one size.
	private readonly links = new Map<string, string>()

	private changed(): void {
		this.columnEnds.clear()
		this.rowEnds.clear()
	}

	private use(graph: Graph): void {
		if (graph === this.graph) return

		this.graph = graph
		this.reset()
	}

	private state(graph: Graph, sheet: number): SheetState {
		this.use(graph)

		let state = this.sheets.get(sheet)

		if (state === undefined) {
			state = { columns: { spans: [], end: -1 }, rows: { spans: [], end: -1 } }
			this.sheets.set(sheet, state)
			this.seed(graph, sheet, state)
		}

		return state
	}

	// Every whole-column and whole-row reference the engine holds, as covered: ranges of one sheet then end at the
	// same place however they come to be read.
	private seed(graph: Graph, sheet: number, state: SheetState): void {
		for (const vertex of graph.graph.getInfiniteRanges()) {
			const range = (vertex as { range?: WholeRange }).range

			if (range?.sheet !== sheet) continue

			if (range.end.row === Number.POSITIVE_INFINITY) state.columns.spans = cover(state.columns.spans, range.start.col, range.end.col)
			else if (range.end.col === Number.POSITIVE_INFINITY) state.rows.spans = cover(state.rows.spans, range.start.row, range.end.row)
		}
	}

	// The sheet's cells, or null for a sheet the engine holds none for (one a formula names that the workbook lacks).
	private storage(graph: Graph, sheet: number): Storage | null {
		try {
			return graph.addressMapping.getStrategyForSheetOrThrow(sheet) as unknown as Storage
		} catch {
			return null
		}
	}

	private root(name: string): string {
		let at = name

		for (let parent = this.links.get(at); parent !== undefined && parent !== at; parent = this.links.get(at)) at = parent

		return at
	}

	// Sheets (by name, any case) that one formula reads whole columns or rows of together.
	link(names: readonly string[]): void {
		const roots = [...new Set(names.map(name => this.root(name.toUpperCase())))]
		const [first] = roots

		if (first === undefined || roots.length < 2) return

		for (const root of roots) this.links.set(root, first)

		this.changed()
	}

	rename(from: string, to: string): void {
		const before = from.toUpperCase()
		const after = to.toUpperCase()

		if (before === after) return

		for (const [name, parent] of [...this.links]) {
			this.links.delete(name)
			this.links.set(name === before ? after : name, parent === before ? after : parent)
		}
	}

	// The engine sheet ids whose whole ranges end together with this sheet's.
	private group(graph: Graph, sheet: number): number[] {
		const name = graph.sheetMapping.getSheetName(sheet)?.toUpperCase()

		if (name === undefined || this.links.size === 0) return [sheet]

		const root = this.root(name)
		const ids = new Set([sheet])

		for (const member of this.links.keys()) {
			if (this.root(member) !== root) continue

			const id = graph.sheetMapping.getSheetId(member)

			if (id !== undefined) ids.add(id)
		}

		return [...ids]
	}

	private localEnd(graph: Graph, sheet: number, columns: boolean): number {
		const state = this.state(graph, sheet)
		const axis = columns ? state.columns : state.rows

		if (axis.end < 0) {
			const storage = this.storage(graph, sheet)

			axis.end =
				storage === null || axis.spans.length === 0
					? 0
					: columns
						? lastRowIn(storage, axis.spans)
						: lastColumnIn(storage, axis.spans)
		}

		return axis.end
	}

	// One past the last covered cell of the sheet's group, along whole columns (`columns`) or whole rows.
	private groupEnd(graph: Graph, sheet: number, columns: boolean): number {
		const cache = columns ? this.columnEnds : this.rowEnds
		const known = cache.get(sheet)

		if (known !== undefined) return known

		let end = 0

		for (const member of this.group(graph, sheet)) end = Math.max(end, this.localEnd(graph, member, columns))

		cache.set(sheet, end)

		return end
	}

	private covering(graph: Graph, range: WholeRange, columns: boolean): void {
		const state = this.state(graph, range.sheet)
		const axis = columns ? state.columns : state.rows
		const first = columns ? range.start.col : range.start.row
		const last = columns ? range.end.col : range.end.row

		if (covers(axis.spans, first, last)) return

		axis.spans = cover(axis.spans, first, last)
		axis.end = -1
		this.changed()
	}

	// Where a whole-column range ends: one row past the last cell of the columns its group covers, so a search for a
	// blank finds one (MATCH(TRUE,ISBLANK(A:A),0)); every row from there down is empty.
	height(graph: Graph, range: WholeRange): number {
		this.covering(graph, range, true)

		return Math.min(MAX_ROWS, this.groupEnd(graph, range.sheet, true) + 1)
	}

	width(graph: Graph, range: WholeRange): number {
		this.covering(graph, range, false)

		return Math.min(MAX_COLUMNS, this.groupEnd(graph, range.sheet, false) + 1)
	}

	// Where the sheet's group's whole-column references end (`rows`) and its whole-row references (`cols`) as covered
	// now, without the blank row.
	ends(graph: Graph, sheet: number): { readonly rows: number; readonly cols: number } {
		return { rows: this.groupEnd(graph, sheet, true), cols: this.groupEnd(graph, sheet, false) }
	}

	// The sheets (upper-case names) whose whole ranges end with this sheet's.
	groupNames(graph: Graph, sheet: number): Set<string> {
		const names = new Set<string>()

		for (const id of this.group(graph, sheet)) {
			const name = graph.sheetMapping.getSheetName(id)

			if (name !== undefined) names.add(name.toUpperCase())
		}

		return names
	}

	// What the group's whole-column references read per row they reach (`rows`: each one's columns times the formulas
	// reading it), and its whole-row references per column (`cols`).
	load(graph: Graph, sheet: number): { readonly rows: number; readonly cols: number } {
		const members = new Set(this.group(graph, sheet))
		let rows = 0
		let cols = 0

		for (const vertex of graph.graph.getInfiniteRanges()) {
			const range = (vertex as { range?: WholeRange }).range

			if (range === undefined || !members.has(range.sheet)) continue

			const readers = graph.graph.adjacentNodesCount(vertex)

			if (range.end.row === Number.POSITIVE_INFINITY) rows += readers * (range.end.col - range.start.col + 1)
			else if (range.end.col === Number.POSITIVE_INFINITY) cols += readers * (range.end.row - range.start.row + 1)
		}

		return { rows, cols }
	}

	// A non-empty cell written (a held result's too): the ends move past it when its column (row) is covered.
	written(sheet: number, row: number, col: number): void {
		const state = this.sheets.get(sheet)

		if (state === undefined) return

		if (state.columns.end >= 0 && row >= state.columns.end && covers(state.columns.spans, col, col)) {
			state.columns.end = row + 1
			this.columnEnds.clear()
		}

		if (state.rows.end >= 0 && col >= state.rows.end && covers(state.rows.spans, row, row)) {
			state.rows.end = col + 1
			this.rowEnds.clear()
		}
	}

	// Cells moved or arrived other than one by one (rows or columns inserted or removed, a spilled array): every end and
	// what is covered read again. Linked sheets stay linked.
	reset(): void {
		this.sheets.clear()
		this.changed()
	}

	// What a cell written at the sheet's far edge makes whole-column (whole-row) references read: whether its column
	// (row) is one they cover.
	reads(sheet: number, row: number, col: number): { readonly rows: boolean; readonly cols: boolean } {
		const state = this.sheets.get(sheet)

		if (state === undefined) return { rows: true, cols: true }

		return { rows: covers(state.columns.spans, col, col), cols: covers(state.rows.spans, row, row) }
	}
}

// The ends in force for the engine calculating now (FormulaEngine sets them around every call into it).
let active: WholeRanges | null = null

export function withWholeRanges<T>(ranges: WholeRanges | null, work: () => T): T {
	const previous = active

	active = ranges

	try {
		return work()
	} finally {
		active = previous
	}
}

let installed = false

type EffectiveExtent = (this: WholeRange, graph: Graph) => number

// Ends the engine's whole-column and whole-row ranges as above. Its range classes are not exported: they are found
// from the ranges of a throwaway engine, and their extent methods replaced once (HyperFormula is pinned exactly).
export function installWholeRanges(licenseKey: string): void {
	if (installed) return

	installed = true

	const probe = HyperFormula.buildFromArray([[], [null, null, "=SUM(1:1)"], [null, "=SUM(A:A)"]], { licenseKey })

	try {
		for (const vertex of probe.dependencyGraph.graph.getInfiniteRanges()) {
			const range = (vertex as { range?: WholeRange }).range

			if (range === undefined) continue

			const prototype = Object.getPrototypeOf(range) as { effectiveHeight?: EffectiveExtent; effectiveWidth?: EffectiveExtent }

			if (range.end.row === Number.POSITIVE_INFINITY) {
				prototype.effectiveHeight = function (graph) {
					return active === null ? graph.getSheetHeight(this.sheet) : active.height(graph, this)
				}
			} else if (range.end.col === Number.POSITIVE_INFINITY) {
				prototype.effectiveWidth = function (graph) {
					return active === null ? graph.getSheetWidth(this.sheet) : active.width(graph, this)
				}
			}
		}
	} finally {
		probe.destroy()
	}
}
