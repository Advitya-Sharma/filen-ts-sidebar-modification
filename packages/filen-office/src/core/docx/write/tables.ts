import type { Node } from "prosemirror-model"
import { decodeUtf8 } from "../../xml/index"
import { NO_SOURCE, docxSchema, endsWithTable } from "../edit/schema"
import {
	cellOverrides,
	previousGrid,
	propertySource,
	rowGrid,
	rowOverrides,
	sourceCellProps,
	sourceChanges,
	sourceRowProps,
	sourceTableProps,
	tableMap,
	tableOverrides,
	type CellOverrides,
	type RowOverrides,
	type SourceResolution,
	type TableMap,
	type TableOverrides,
	type TrackStamp
} from "../edit/tables/map"
import { dataAttr, dataChildren, type ElementData } from "../edit/copies"
import { Item, Revision, RevisionFlag } from "../model/index"
import { Ns } from "../model/namespaces"
import type { TableCellProps, TableProps, TableRowProps } from "../props/index"
import { closeTag, copyNode, freshContainer, openTag, Out, spliceChildren } from "./markup"
import { refuseBlockMarks, src } from "./source"
import type { StoryWriting } from "./story"
import { DocxWriteError, integer, name as checkName } from "./values"

// Tables, rows and cells. What is as the source has it is copied as its bytes; what table commands changed
// is spliced into a copy of the source's property elements (w:tblPr, w:tblGrid, w:trPr, w:tcPr) in schema
// order, and new rows and cells start from the properties of the source row or cell they were modelled on.
// A cell's width in twips (w:tcW) is its `colwidth` once an edit set one; a grid a row outgrew grows with it
// (and the table's width in twips with its grid). A property change recorded while tracking is written as
// Word writes one (w:tblPrChange, w:tblGridChange, w:trPrChange, w:tcPrChange: the whole property element as
// it was before, ECMA-376 Part 1 §17.13.5.33-37); a source's own change accepted is left out, rejected puts
// the properties it holds back in place of the current ones.

const marks = docxSchema.marks

// Children of container properties a copy leaves out: tracked changes, whose ids are the original's.
const TRACKED_TABLE = ["tblPrChange", "tblGridChange"]
const TRACKED_ROW = ["ins", "del", "trPrChange", "tblPrExChange"]
const TRACKED_CELL = ["cellIns", "cellDel", "cellMerge", "tcPrChange"]
// CT_TblPrExBase: child order.
const EXCEPTION_PROPERTIES = [
	"tblW",
	"jc",
	"tblCellSpacing",
	"tblInd",
	"tblBorders",
	"shd",
	"tblLayout",
	"tblCellMar",
	"tblLook",
	"tblPrExChange"
]

// ECMA-376 Part 1 (Transitional) CT_TblPr, CT_TrPr and CT_TcPr: child order.
const TABLE_PROPERTIES = [
	"tblStyle",
	"tblpPr",
	"tblOverlap",
	"bidiVisual",
	"tblStyleRowBandSize",
	"tblStyleColBandSize",
	"tblW",
	"jc",
	"tblCellSpacing",
	"tblInd",
	"tblBorders",
	"shd",
	"tblLayout",
	"tblCellMar",
	"tblLook",
	"tblCaption",
	"tblDescription",
	"tblPrChange"
]
const ROW_PROPERTIES = [
	"cnfStyle",
	"divId",
	"gridBefore",
	"gridAfter",
	"wBefore",
	"wAfter",
	"cantSplit",
	"trHeight",
	"tblHeader",
	"tblCellSpacing",
	"jc",
	"hidden",
	"ins",
	"del",
	"trPrChange"
]
const CELL_PROPERTIES = [
	"cnfStyle",
	"tcW",
	"gridSpan",
	"hMerge",
	"vMerge",
	"tcBorders",
	"shd",
	"noWrap",
	"tcMar",
	"textDirection",
	"tcFitText",
	"vAlign",
	"hideMark",
	"headers",
	"cellIns",
	"cellDel",
	"cellMerge",
	"tcPrChange"
]

const JC = { start: "left", center: "center", end: "right" } as const
const FILL = /^(?:auto|[0-9A-F]{6})$/
// Grid column width of a new table without a grid (twips).
const NEW_COLUMN = 2000

function markup(write: (out: Out) => void): string {
	const out = new Out()

	write(out)

	return decodeUtf8(out.finish())
}

// A row's or cell's property element left with no child is left out (both optional: CT_Row, CT_Tc).
function unlessEmpty(element: string): string {
	return /^<([^\s/>]+)>\s*<\/\1>$/.test(element) ? "" : element
}

function sameNumbers(a: readonly number[], b: readonly number[]): boolean {
	return a === b || (a.length === b.length && a.every((value, k) => value === b[k]))
}

// The w:gridCol widths of a source w:tblGrid.
function sourceGrid(w: StoryWriting, tblGrid: number): readonly number[] {
	if (tblGrid < 0) return []

	const reader = w.part.reader
	const document = w.part.document
	const out: number[] = []

	for (let child = document.firstChildElement(tblGrid); child >= 0; child = document.nextSiblingElement(child)) {
		if (reader.is(child, Ns.W, "gridCol")) out.push(reader.int(child, Ns.W, "w", 0))
	}

	return out
}

// A w:tblGrid of `grid`, with the w:tblGridChange of a grid change recorded while tracking.
function gridMarkup(w: StoryWriting, grid: readonly number[], change?: { readonly grid: readonly number[] }): string {
	const names = w.names
	const columns = (widths: readonly number[]): string => {
		for (const width of widths) integer("w:gridCol w", width, 0)

		return widths.map(width => names.empty("gridCol", [["w", String(width)]])).join("")
	}
	const recorded =
		change === undefined
			? ""
			: `<${names.el("tblGridChange")} ${names.attr("id")}="${String(w.env.nextRevisionId())}"><${names.el("tblGrid")}>${columns(change.grid)}</${names.el("tblGrid")}></${names.el("tblGridChange")}>`

	return `<${names.el("tblGrid")}>${columns(grid)}${recorded}</${names.el("tblGrid")}>`
}

// ---- tracked property changes ---------------------------------------------------------------------------

// A w:tblPrChange / w:trPrChange / w:tcPrChange the editor recorded: its stamp and the property element as it
// was before (`previous`, markup).
function changeElement(w: StoryWriting, local: string, stamp: TrackStamp, previous: string): string {
	if (stamp.author === "") throw new DocxWriteError("missingSource", `A tracked ${local} without an author`)

	checkName(`w:${local} author`, stamp.author)

	const open = w.names.empty(local, [
		["id", String(w.env.nextRevisionId())],
		["author", stamp.author],
		["date", stamp.date]
	])

	return `${open.slice(0, -2)}>${previous}</${w.names.el(local)}>`
}

// The W children of source element `node` by local name.
function childrenOf(w: StoryWriting, node: number): Map<string, number[]> {
	const out = new Map<string, number[]>()
	const reader = w.part.reader
	const document = w.part.document

	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) {
		if (reader.group(child) !== Ns.W) continue

		const local = document.localName(child)
		const list = out.get(local)

		if (list === undefined) out.set(local, [child])
		else list.push(child)
	}

	return out
}

// What a source property element (`container`: w:tblPr, w:trPr, w:tblPrEx, w:tcPr) becomes once the source's
// own tracked change in it (`local`, holding the element before it as `inner`) is resolved: accepted, the
// change leaves; rejected, the children it holds replace the others (the tracked ones, `kept`, stay). `also`:
// further records the resolution takes out (w:cellMerge).
function resolvedChanges(
	w: StoryWriting,
	container: number,
	local: string,
	inner: string,
	resolution: SourceResolution | undefined,
	kept: readonly string[],
	also: readonly string[] = []
): Map<string, string> {
	const changes = new Map<string, string>()

	if (resolution === undefined || container < 0) return changes

	for (const name of [local, ...also]) changes.set(name, "")

	if (resolution === "accepted") return changes

	const change = childrenOf(w, container).get(local)?.[0]
	const previous = change === undefined ? undefined : childrenOf(w, change).get(inner)?.[0]

	if (previous === undefined) return changes

	const now = childrenOf(w, container)
	const before = childrenOf(w, previous)
	const document = w.part.document

	for (const name of now.keys()) if (!kept.includes(name) && !before.has(name)) changes.set(name, "")

	for (const [name, nodes] of before) {
		if (kept.includes(name)) continue

		changes.set(
			name,
			markup(out => nodes.forEach(node => copyNode(out, document, node)))
		)
	}

	return changes
}

// CT_TblBorders, CT_TcBorders, CT_TblCellMar, CT_TcMar: child order.
const SIDES = ["top", "start", "left", "bottom", "end", "right", "insideH", "insideV", "tl2br", "tr2bl"]
const NESTED: ReadonlySet<string> = new Set(["tblBorders", "tcBorders", "tblCellMar", "tcMar"])

// A copy of source element `node` with its WordprocessingML children in schema order (`order`), or undefined when
// they are in it already.
function inOrder(w: StoryWriting, node: number, order: readonly string[]): string | undefined {
	const document = w.part.document
	const reader = w.part.reader
	const children: { node: number; rank: number }[] = []
	let previous = -1
	let disordered = false

	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) {
		const index = reader.group(child) === Ns.W ? order.indexOf(document.localName(child)) : -1
		const rank = index < 0 ? order.length : index

		disordered ||= rank < previous
		previous = rank
		children.push({ node: child, rank })
	}

	if (!disordered) return undefined

	children.sort((a, b) => a.rank - b.rank)

	return markup(out => {
		openTag(out, document, node)

		for (const child of children) copyNode(out, document, child.node)

		closeTag(out, document, node)
	})
}

// The property element as it was before a change recorded while tracking: the base element (`container`, or
// a new one) with `changes`, without what carries ids (`tracked`). Its copies of the base's children are in
// schema order (Word writes the record from its own model; a source out of order is not copied into it so).
function previousElement(
	w: StoryWriting,
	container: number,
	local: string,
	changes: Map<string, string>,
	tracked: readonly string[],
	order: readonly string[]
): string {
	strip(changes, tracked)

	if (container >= 0) {
		for (const [name, nodes] of childrenOf(w, container)) {
			if (!NESTED.has(name) || changes.has(name)) continue

			const node = nodes[0]
			const ordered = node === undefined ? undefined : inOrder(w, node, SIDES)

			if (ordered !== undefined) changes.set(name, ordered)
		}

		return markup(sub => spliceChildren(sub, w.part.reader, container, changes, order))
	}

	const fresh = freshContainer(w.names, local, changes, order)

	return fresh === "" ? w.names.empty(local) : fresh
}

// A w:ins / w:del / w:cellIns / w:cellDel recording a change the editor tracked.
function trackElement(w: StoryWriting, local: string, stamp: TrackStamp | undefined): string {
	const author = stamp?.author ?? w.env.context.identity.author

	if (author === undefined || author === "") throw new DocxWriteError("missingSource", `A tracked ${local} without an author`)

	checkName(`w:${local} author`, author)

	return w.names.empty(local, [
		["id", String(w.env.nextRevisionId())],
		["author", author],
		["date", stamp?.date]
	])
}

// The tracked insertion and deletion children (`locals`) where the revision bits differ from what the
// source has (`had`). A copy (`owned` false) whose record stayed with its original is no tracked change, as a
// copy of the source's leaves its tracked children out: records are written from stamps, never invented.
function trackChanges(
	w: StoryWriting,
	changes: Map<string, string>,
	bits: number,
	had: number,
	locals: readonly [string, string],
	stamps: { readonly ins?: TrackStamp; readonly del?: TrackStamp },
	owned: boolean
): void {
	const flags = [RevisionFlag.Inserted, RevisionFlag.Deleted] as const

	flags.forEach((flag, k) => {
		const local = locals[k] as string
		const want = (bits & flag) !== 0
		const stamp = k === 0 ? stamps.ins : stamps.del

		if (want === ((had & flag) !== 0)) return
		if (want && !owned && stamp === undefined) return

		changes.set(local, want ? trackElement(w, local, stamp) : "")
	})
}

// What a command recorded must be what the writer knows: anything else refuses the save rather than being
// left out.
const TABLE_KEYS = new Set(["from", "copy", "align", "width", "borders", "look", "change", "gridChange", "source"])
const ROW_KEYS = new Set(["from", "copy", "header", "gridBefore", "gridAfter", "ins", "del", "change", "source"])
const CELL_KEYS = new Set(["from", "copy", "shading", "vAlign", "ins", "del", "change", "source"])
// What a change record keeps of the properties before it.
const TABLE_PROPS = new Set(["align", "width", "borders", "look"])
const ROW_PROPS = new Set(["header", "gridBefore", "gridAfter"])
const CELL_PROPS = new Set(["shading", "vAlign"])
const V_ALIGN = new Set(["top", "center", "bottom"])

function checkKeys(value: object, known: ReadonlySet<string>, what: string): void {
	for (const key of Object.keys(value))
		if (!known.has(key)) throw new DocxWriteError("unsupported", `A ${what} change the writer does not know: ${key}`)
}

function checkStamp(stamp: TrackStamp | undefined): void {
	if (stamp === undefined) return
	if (typeof stamp.author !== "string" || (stamp.date !== undefined && typeof stamp.date !== "string"))
		throw new DocxWriteError("invalidValue", "A tracked change stamp is not valid")
}

// A change record and a source resolution as the writer knows them.
function checkRecord(
	value: { readonly change?: TrackStamp & { readonly props?: object }; readonly source?: string },
	known: ReadonlySet<string>,
	what: string
): void {
	if (value.source !== undefined && value.source !== "accepted" && value.source !== "rejected")
		throw new DocxWriteError("invalidValue", `A ${what} change resolution: ${JSON.stringify(value.source)}`)

	if (value.change === undefined) return

	checkStamp(value.change)

	if (typeof value.change.props !== "object" || value.change.props === null)
		throw new DocxWriteError("invalidValue", `A tracked ${what} change without the properties before it`)

	checkKeys(value.change.props, known, `${what} property`)
}

// The children of the properties of another part's table, row or cell a copy keeps (its overrides' `copy`), as
// markup of this part by local name.
function copiedProperties(w: StoryWriting, data: ElementData | undefined): Map<string, string> {
	return data === undefined ? new Map() : dataChildren(w.part, w.names.w, data)
}

// Children a copy leaves out, unless a change sets them.
function strip(changes: Map<string, string>, locals: readonly string[]): void {
	for (const local of locals) if (!changes.has(local)) changes.set(local, "")
}

// ---- tables ---------------------------------------------------------------------------------------------

// The w:tblPr children table commands set: style `styleId` (`baseStyle` the base element's), `overrides`, and
// the width in twips a changed grid gives a table whose width is in twips (`width`: the base's).
function tableChanges(
	w: StoryWriting,
	styleId: string | null,
	baseStyle: string | null,
	overrides: TableOverrides,
	width: TableProps["width"],
	grid: readonly number[],
	gridChanged: boolean,
	fresh: boolean
): Map<string, string> {
	const names = w.names

	checkKeys(overrides, TABLE_KEYS, "table")

	if (overrides.align !== undefined && !(overrides.align in JC))
		throw new DocxWriteError("invalidValue", `w:jc: ${JSON.stringify(overrides.align)}`)
	if (
		overrides.width !== undefined &&
		overrides.width.type !== "dxa" &&
		overrides.width.type !== "pct" &&
		overrides.width.type !== "auto"
	)
		throw new DocxWriteError("invalidValue", `w:tblW type: ${JSON.stringify(overrides.width.type)}`)
	const changes = new Map<string, string>()

	if (styleId !== baseStyle) {
		if (styleId !== null) checkName("w:tblStyle", styleId)

		changes.set("tblStyle", styleId === null ? "" : names.val("tblStyle", styleId))
	}

	if (overrides.width !== undefined) {
		integer("w:tblW w", overrides.width.value, 0)
		changes.set(
			"tblW",
			names.empty("tblW", [
				["w", String(overrides.width.value)],
				["type", overrides.width.type]
			])
		)
	} else if (fresh) {
		changes.set(
			"tblW",
			names.empty("tblW", [
				["w", "0"],
				["type", "auto"]
			])
		)
	} else if (gridChanged && width?.type === "dxa") {
		changes.set(
			"tblW",
			names.empty("tblW", [
				["w", String(grid.reduce((sum, value) => sum + value, 0))],
				["type", "dxa"]
			])
		)
	}

	if (overrides.align !== undefined) changes.set("jc", names.val("jc", JC[overrides.align]))

	if (overrides.borders === true) {
		const side = (local: string): string =>
			names.empty(local, [
				["val", "single"],
				["sz", "4"],
				["space", "0"],
				["color", "auto"]
			])

		changes.set(
			"tblBorders",
			`<${names.el("tblBorders")}>${["top", "left", "bottom", "right", "insideH", "insideV"].map(side).join("")}</${names.el("tblBorders")}>`
		)
	}

	if (overrides.look === true) {
		changes.set(
			"tblLook",
			names.empty("tblLook", [
				["val", "04A0"],
				["firstRow", "1"],
				["lastRow", "0"],
				["firstColumn", "1"],
				["lastColumn", "0"],
				["noHBand", "0"],
				["noVBand", "1"]
			])
		)
	}

	return changes
}

// What the w:tblPr of a table becomes: the source's resolution (`base`), the commands' changes, and a change
// recorded while tracking (`baseGrid`: the grid the base element goes with).
function tableProperties(
	w: StoryWriting,
	node: Node,
	tblPr: number,
	base: Map<string, string>,
	baseStyle: string | null,
	props: TableProps,
	baseGrid: readonly number[],
	grid: readonly number[],
	fresh: boolean
): Map<string, string> {
	const overrides = tableOverrides(node)
	const gridChanged = !sameNumbers(grid, baseGrid)
	const styleId = node.attrs["styleId"] as string | null

	checkRecord(overrides, TABLE_PROPS, "table")
	checkStamp(overrides.gridChange)

	if (overrides.gridChange !== undefined && !Array.isArray(overrides.gridChange.grid))
		throw new DocxWriteError("invalidValue", "A tracked grid change without the grid before it")

	const changes = new Map([...base, ...tableChanges(w, styleId, baseStyle, overrides, props.width, grid, gridChanged, fresh)])
	const change = overrides.change

	if (change !== undefined) {
		const before = overrides.gridChange?.grid ?? grid
		const previous = new Map([
			...base,
			...tableChanges(w, change.styleId, baseStyle, change.props, props.width, before, !sameNumbers(before, baseGrid), fresh)
		])

		if (tblPr < 0 && !fresh && !previous.has("tblW")) previous.set("tblW", autoWidth(w))

		changes.set(
			"tblPrChange",
			changeElement(w, "tblPrChange", change, previousElement(w, tblPr, "tblPr", previous, TRACKED_TABLE, TABLE_PROPERTIES))
		)
	}

	return changes
}

// The style of copied w:tblPr `data`.
function dataStyle(data: ElementData): string | undefined {
	const style = data.children.find(child => child.name === "tblStyle" && child.uri === "w")

	return style === undefined ? undefined : dataAttr(style, "val")
}

// The w:tblW Word gives a table without one.
function autoWidth(w: StoryWriting): string {
	return w.names.empty("tblW", [
		["w", "0"],
		["type", "auto"]
	])
}

// Rows of `map` wider than `columns` grid columns.
function wideRows(map: TableMap, columns: number): number {
	let count = 0

	for (let r = 0; r < map.height; r++) if (map.row(r).end + map.row(r).after > columns) count++

	return count
}

// The grid as written: as wide as the table's rows need when one outgrew it. A table whose grid and rows
// wider than it are as the story was opened keeps its grid.
function fittedGrid(map: TableMap, original: TableMap | undefined): readonly number[] {
	const grid = map.grid
	const columns = map.width

	if (columns <= grid.length) return grid
	if (original !== undefined && sameNumbers(grid, original.grid) && wideRows(map, grid.length) <= wideRows(original, grid.length))
		return grid
	if (grid.length === 0) return Array.from({ length: columns }, () => NEW_COLUMN)

	const average = Math.round(grid.reduce((sum, value) => sum + value, 0) / grid.length)

	return [...grid, ...Array.from({ length: columns - grid.length }, () => average)]
}

export function* writeTable(w: StoryWriting, out: Out, node: Node): Generator<void, void> {
	const part = w.part
	const document = part.document
	const names = w.names
	const item = src(node)
	const store = w.env.context.store
	const overrides = tableOverrides(node)

	if (item === NO_SOURCE || part.kinds[item] !== Item.Table) {
		const grid = fittedGrid(tableMap(node, part, store), undefined)
		const model = propertySource(part, node, Item.Table)

		if (model === NO_SOURCE) {
			// A copy of another part's table starts from its properties.
			const copied = overrides.copy
			const changes = tableProperties(
				w,
				node,
				-1,
				copiedProperties(w, copied),
				copied === undefined ? null : (dataStyle(copied) ?? null),
				store.table(document, -1),
				[],
				grid,
				copied === undefined
			)

			if (copied !== undefined && !changes.has("tblW")) changes.set("tblW", autoWidth(w))

			out.text(
				`<${names.el("tbl")}>${freshContainer(names, "tblPr", changes, TABLE_PROPERTIES)}${gridMarkup(w, grid, overrides.gridChange)}`
			)
		} else {
			// A new table modelled on a source one (a copy): that table's properties without its tracked change.
			const modelRow = part.aux[model] ?? -1
			const modelPr = part.tables.tblPr[modelRow] ?? -1
			const props = sourceTableProps(part, store, model, false)
			const modelGrid = sourceGrid(w, part.tables.tblGrid[modelRow] ?? -1)
			const changes = tableProperties(w, node, modelPr, new Map(), props.styleId ?? null, props, modelGrid, grid, false)

			out.text(`<${names.el("tbl")}>`)

			if (changes.size === 0 && modelPr >= 0) {
				w.copyProperties(out, [modelPr], TRACKED_TABLE)
			} else {
				strip(changes, ["tblPrChange"])
				out.text(
					modelPr >= 0
						? markup(sub => spliceChildren(sub, part.reader, modelPr, changes, TABLE_PROPERTIES))
						: freshContainer(names, "tblPr", new Map([["tblW", autoWidth(w)], ...changes]), TABLE_PROPERTIES)
				)
			}

			out.text(gridMarkup(w, grid, overrides.gridChange))
		}

		yield* w.writeChildren(out, node, NO_SOURCE, (row, rowOut) => writeRow(w, rowOut, row, gridBefore(node)))
		out.text(`</${names.el("tbl")}>`)

		return
	}

	const row = part.aux[item] ?? -1
	const tblPr = part.tables.tblPr[row] ?? -1
	const tblGrid = part.tables.tblGrid[row] ?? -1
	const original = w.original(node)
	const resolution = overrides.source
	const rejected = resolution === "rejected"
	const props = sourceTableProps(part, store, item, rejected)
	const sourceStyle = rejected || original === undefined ? (props.styleId ?? null) : (original.attrs["styleId"] as string | null)
	const written = fittedGrid(tableMap(node, part, store), original === undefined ? undefined : tableMap(original, part, store))
	const fileGrid = (original?.attrs["grid"] as readonly number[] | undefined) ?? sourceGrid(w, tblGrid)
	const gridRecord = sourceChanges(part, item).get(Revision.TableGrid)
	// The grid the properties go with: the one before the source's change once that was rejected.
	const baseGrid = (rejected ? previousGrid(part, gridRecord) : undefined) ?? fileGrid
	// Written anew: changed, with a change recorded, or without the source's change once that was resolved.
	const gridWritten =
		!sameNumbers(written, fileGrid) || overrides.gridChange !== undefined || (resolution !== undefined && gridRecord !== undefined)
	const base = resolvedChanges(w, tblPr, "tblPrChange", "tblPr", resolution, [])
	const changes = tableProperties(w, node, tblPr, base, sourceStyle, props, baseGrid, written, false)
	const tblPrMarkup = (): string =>
		tblPr >= 0
			? markup(sub => spliceChildren(sub, part.reader, tblPr, changes, TABLE_PROPERTIES))
			: freshContainer(names, "tblPr", new Map([["tblW", autoWidth(w)], ...changes]), TABLE_PROPERTIES)

	if (!w.owns(node)) {
		// A second table from one source (a split, a move of part of it): its own properties, no ids.
		out.text(`<${names.el("tbl")}>`)

		if (changes.size === 0) {
			w.copyProperties(out, [tblPr], TRACKED_TABLE)
		} else {
			strip(changes, ["tblPrChange"])
			out.text(tblPrMarkup())
		}

		if (!gridWritten && tblGrid >= 0) w.copyProperties(out, [tblGrid], TRACKED_TABLE)
		else out.text(gridMarkup(w, written, overrides.gridChange))

		yield* w.writeChildren(out, node, NO_SOURCE, (child, rowOut) => writeRow(w, rowOut, child, gridBefore(node)))
		out.text(`</${names.el("tbl")}>`)

		return
	}

	const element = part.node(item)

	openTag(out, document, element)

	let from = document.openEnd(element)

	if (tblPr < 0 && changes.size > 0) out.text(tblPrMarkup())

	for (const property of [tblPr, tblGrid].sort((a, b) => a - b)) {
		if (property < 0) continue

		if (property === tblPr && changes.size > 0) out.text(tblPrMarkup())
		else if (property === tblGrid && gridWritten) out.text(gridMarkup(w, written, overrides.gridChange))
		else w.copyWhole(out, property)

		from = document.end(property)
	}

	if (tblGrid < 0 && gridWritten) out.text(gridMarkup(w, written, overrides.gridChange))

	yield* w.writeChildren(out, node, item, (child, rowOut) => writeRow(w, rowOut, child, gridBefore(node)), {
		from,
		to: document.contentEnd(element),
		tail: []
	})
	closeTag(out, document, element)
}

// The grid a table's recorded changes were made against: the one before its recorded grid change.
function gridBefore(table: Node): readonly number[] {
	return tableOverrides(table).gridChange?.grid ?? (table.attrs["grid"] as readonly number[])
}

// ---- rows -----------------------------------------------------------------------------------------------

// The w:trPr children row commands set (`props`: the base element's).
function rowChanges(w: StoryWriting, overrides: RowOverrides, props: TableRowProps): Map<string, string> {
	const names = w.names
	const changes = new Map<string, string>()

	if (overrides.header !== undefined && typeof overrides.header !== "boolean") throw new DocxWriteError("invalidValue", "w:tblHeader")
	if (overrides.header !== undefined && overrides.header !== (props.header ?? false))
		changes.set("tblHeader", overrides.header ? names.empty("tblHeader") : "")

	for (const [local, width] of [
		["gridBefore", "wBefore"],
		["gridAfter", "wAfter"]
	] as const) {
		const value = overrides[local]

		if (value === undefined || value === (props[local] ?? 0)) continue

		integer(`w:${local}`, value, 0)
		changes.set(local, value > 0 ? names.val(local, String(value)) : "")
		// Its width was the old columns'.
		changes.set(width, "")
	}

	return changes
}

function* writeRow(w: StoryWriting, out: Out, node: Node, grid: readonly number[]): Generator<void, void> {
	refuseBlockMarks(node, marks.row_props)

	const part = w.part
	const document = part.document
	const names = w.names
	const overrides = rowOverrides(node)
	const own = src(node)
	const owned = own !== NO_SOURCE && part.kinds[own] === Item.Row && w.owns(node)
	const model = propertySource(part, node, Item.Row)
	const modelRow = model === NO_SOURCE ? -1 : (part.aux[model] ?? -1)
	const trPr = modelRow < 0 ? -1 : (part.rows.trPr[modelRow] ?? -1)
	const tblPrEx = modelRow < 0 ? -1 : (part.rows.tblPrEx[modelRow] ?? -1)
	const resolution = overrides.source
	const props =
		model === NO_SOURCE
			? w.env.context.store.row(document, -1)
			: sourceRowProps(part, w.env.context.store, model, resolution === "rejected")

	checkKeys(overrides, ROW_KEYS, "row")
	checkStamp(overrides.ins)
	checkStamp(overrides.del)
	checkRecord(overrides, ROW_PROPS, "row")

	const base =
		model === NO_SOURCE ? copiedProperties(w, overrides.copy) : resolvedChanges(w, trPr, "trPrChange", "trPr", resolution, TRACKED_ROW)
	const changes = new Map([...base, ...rowChanges(w, overrides, props)])

	trackChanges(
		w,
		changes,
		node.attrs["revision"] as number,
		owned ? (part.rows.revision[modelRow] ?? 0) : 0,
		["ins", "del"],
		overrides,
		owned
	)

	if (overrides.change !== undefined) {
		const previous = new Map([...base, ...rowChanges(w, overrides.change.props, props)])

		changes.set(
			"trPrChange",
			changeElement(w, "trPrChange", overrides.change, previousElement(w, trPr, "trPr", previous, TRACKED_ROW, ROW_PROPERTIES))
		)
	}

	const exceptions = resolvedChanges(w, tblPrEx, "tblPrExChange", "tblPrEx", resolution, [])
	const spans = reshapedSpans(w, node, grid)
	const cell = (child: Node, cellOut: Out): Generator<void, void> => writeCell(w, cellOut, child, spans?.get(child)?.shift())

	if (owned) {
		const element = part.node(own)

		openTag(out, document, element)

		// w:tblPrEx and w:trPr in the order the source has them (CT_Row puts tblPrEx first).
		let from = document.openEnd(element)

		for (const property of [trPr, tblPrEx].sort((a, b) => a - b)) {
			if (property < 0) continue

			if (property === trPr && changes.size > 0)
				out.text(unlessEmpty(markup(sub => spliceChildren(sub, part.reader, trPr, changes, ROW_PROPERTIES))))
			else if (property === tblPrEx && exceptions.size > 0)
				out.text(markup(sub => spliceChildren(sub, part.reader, tblPrEx, exceptions, EXCEPTION_PROPERTIES)))
			else w.copyWhole(out, property)

			from = document.end(property)
		}

		if (trPr < 0 && changes.size > 0) out.text(freshContainer(names, "trPr", changes, ROW_PROPERTIES))

		yield* w.writeChildren(out, node, own, cell, { from, to: document.contentEnd(element), tail: [] })
		closeTag(out, document, element)

		return
	}

	out.text(`<${names.el("tr")}>`)

	if (model !== NO_SOURCE && changes.size === 0 && exceptions.size === 0) {
		// A copy of a source row, or a new row modelled on one: its properties without tracked changes.
		w.copyProperties(out, [trPr, tblPrEx], TRACKED_ROW)
	} else if (model !== NO_SOURCE) {
		if (tblPrEx >= 0 && exceptions.size === 0) w.copyProperties(out, [tblPrEx], TRACKED_ROW)
		else if (tblPrEx >= 0) out.text(markup(sub => spliceChildren(sub, part.reader, tblPrEx, exceptions, EXCEPTION_PROPERTIES)))

		strip(changes, TRACKED_ROW)
		out.text(
			trPr >= 0
				? unlessEmpty(markup(sub => spliceChildren(sub, part.reader, trPr, changes, ROW_PROPERTIES)))
				: freshContainer(names, "trPr", changes, ROW_PROPERTIES)
		)
	} else {
		out.text(freshContainer(names, "trPr", changes, ROW_PROPERTIES))
	}

	yield* w.writeChildren(out, node, NO_SOURCE, cell)
	out.text(`</${names.el("tr")}>`)
}

// ---- cells ----------------------------------------------------------------------------------------------

// For the cells of `row` merged or split since their recorded change: the width of the grid columns each spans
// in `grid` (the grid before), what rejecting the change gives it (in a queue per node: a row may hold one
// node twice). Undefined without such a cell.
function reshapedSpans(w: StoryWriting, row: Node, grid: readonly number[]): Map<Node, number[]> | undefined {
	let out: Map<Node, number[]> | undefined
	let column = rowGrid(w.part, w.env.context.store, row).before

	row.forEach(cell => {
		const span = Math.max(1, cell.attrs["colspan"] as number)

		if (cellOverrides(cell).change?.reshaped === true) {
			let width = 0

			for (let c = column; c < column + span; c++) width += grid[c] ?? 0

			out ??= new Map()
			out.set(cell, [...(out.get(cell) ?? []), width])
		}

		column += span
	})

	return out
}

function* writeCell(w: StoryWriting, out: Out, node: Node, spanned: number | undefined): Generator<void, void> {
	w.cells++

	try {
		yield* writeCellContent(w, out, node, spanned)
	} finally {
		w.cells--
	}
}

// The w:tcPr children cell commands set: the cell's span and merge against the base element's (`baseSpan`,
// `baseVMerge`), its width `colwidth` and `overrides` (`props`: the base element's).
function cellChanges(
	w: StoryWriting,
	node: Node,
	overrides: CellOverrides,
	colwidth: number | null,
	props: TableCellProps,
	baseSpan: number,
	baseVMerge: string | null,
	modelled: boolean
): Map<string, string> {
	const names = w.names
	const span = node.attrs["colspan"] as number
	const vMerge = node.attrs["vMerge"] as string | null
	const changes = new Map<string, string>()

	if (overrides.vAlign !== undefined && !V_ALIGN.has(overrides.vAlign))
		throw new DocxWriteError("invalidValue", `w:vAlign: ${JSON.stringify(overrides.vAlign)}`)

	if (modelled) for (const local of ["gridSpan", "vMerge", "hMerge"]) changes.set(local, "")

	if (span !== baseSpan) changes.set("gridSpan", span > 1 ? names.val("gridSpan", String(span)) : "")
	if (vMerge !== baseVMerge)
		changes.set("vMerge", vMerge === null ? "" : vMerge === "continue" ? names.empty("vMerge") : names.val("vMerge", "restart"))

	// A width an edit gave the cell (twips).
	if (colwidth !== null && (props.width?.type !== "dxa" || props.width.value !== colwidth)) {
		integer("w:tcW w", colwidth, 0)
		changes.set(
			"tcW",
			names.empty("tcW", [
				["w", String(colwidth)],
				["type", "dxa"]
			])
		)
	}

	if (overrides.shading !== undefined) {
		const fill = overrides.shading ?? "auto"

		if (!FILL.test(fill)) throw new DocxWriteError("invalidValue", `w:shd fill: ${JSON.stringify(fill)} is not a valid value`)

		changes.set(
			"shd",
			names.empty("shd", [
				["val", "clear"],
				["color", "auto"],
				["fill", fill]
			])
		)
	}

	if (overrides.vAlign !== undefined) changes.set("vAlign", names.val("vAlign", overrides.vAlign))

	return changes
}

function* writeCellContent(w: StoryWriting, out: Out, node: Node, spanned: number | undefined): Generator<void, void> {
	refuseBlockMarks(node, marks.cell_props)

	const part = w.part
	const document = part.document
	const names = w.names
	// Word wants a cell to end with a paragraph, which a table as its last block would not.
	const endsWithParagraph = !endsWithTable(node)
	const overrides = cellOverrides(node)
	const own = src(node)
	const isSource = own !== NO_SOURCE && part.kinds[own] === Item.Cell
	const owned = isSource && w.owns(node)
	const model = propertySource(part, node, Item.Cell)
	const modelRow = model === NO_SOURCE ? -1 : (part.aux[model] ?? -1)
	const tcPr = modelRow < 0 ? -1 : (part.cells.tcPr[modelRow] ?? -1)
	const resolution = overrides.source
	const props =
		model === NO_SOURCE
			? w.env.context.store.cell(document, -1)
			: sourceCellProps(part, w.env.context.store, model, resolution === "rejected")
	const span = node.attrs["colspan"] as number
	// A new cell modelled on a source cell takes its properties, not its span and merge.
	const modelled = model !== NO_SOURCE && !isSource
	// As the editor reads them (a span below 1 is 1).
	const baseSpan = isSource ? Math.max(1, props.gridSpan ?? 1) : 1
	const baseVMerge = isSource ? (props.vMerge ?? null) : null

	checkKeys(overrides, CELL_KEYS, "cell")
	checkStamp(overrides.ins)
	checkStamp(overrides.del)
	checkRecord(overrides, CELL_PROPS, "cell")
	integer("w:gridSpan", span, 1)

	// A cell carries one tracked change at most (CT_TcPr: cellIns, cellDel or cellMerge).
	if (
		((node.attrs["revision"] as number) & (RevisionFlag.Inserted | RevisionFlag.Deleted)) ===
		(RevisionFlag.Inserted | RevisionFlag.Deleted)
	)
		throw new DocxWriteError("unsupported", "A table cell both inserted and deleted as tracked changes")

	const base =
		model === NO_SOURCE
			? copiedProperties(w, overrides.copy)
			: resolvedChanges(w, tcPr, "tcPrChange", "tcPr", resolution, ["cellIns", "cellDel"], ["cellMerge"])
	const changes = new Map([
		...base,
		...cellChanges(w, node, overrides, node.attrs["colwidth"] as number | null, props, baseSpan, baseVMerge, modelled)
	])

	trackChanges(
		w,
		changes,
		node.attrs["revision"] as number,
		owned ? (part.cells.revision[modelRow] ?? 0) : 0,
		["cellIns", "cellDel"],
		overrides,
		owned
	)

	const change = overrides.change

	if (change !== undefined) {
		// Merged or split since: the width of the columns it spans, as rejecting it in the editor gives it.
		const colwidth = change.reshaped === true && spanned !== undefined ? spanned : change.colwidth
		const previous = new Map([...base, ...cellChanges(w, node, change.props, colwidth, props, baseSpan, baseVMerge, modelled)])

		changes.set(
			"tcPrChange",
			changeElement(w, "tcPrChange", change, previousElement(w, tcPr, "tcPr", previous, TRACKED_CELL, CELL_PROPERTIES))
		)
	}

	const content = (region?: { readonly from: number; readonly to: number }): Generator<void, boolean> =>
		w.writeChildren(
			out,
			node,
			owned ? own : NO_SOURCE,
			(child, blockOut) => w.writeBlock(blockOut, child),
			region === undefined ? undefined : { ...region, tail: [] }
		)

	if (owned) {
		const element = part.node(own)

		openTag(out, document, element)

		if (tcPr >= 0 && changes.size === 0) w.copyWhole(out, tcPr)
		else if (tcPr >= 0) out.text(unlessEmpty(markup(sub => spliceChildren(sub, part.reader, tcPr, changes, CELL_PROPERTIES))))
		else out.text(freshContainer(names, "tcPr", changes, CELL_PROPERTIES))

		yield* content({ from: tcPr >= 0 ? document.end(tcPr) : document.openEnd(element), to: document.contentEnd(element) })

		if (!endsWithParagraph) out.text(`<${names.el("p")}/>`)

		closeTag(out, document, element)

		return
	}

	out.text(`<${names.el("tc")}>`)

	if (model !== NO_SOURCE && changes.size === 0) {
		w.copyProperties(out, [tcPr], TRACKED_CELL)
	} else if (model !== NO_SOURCE && tcPr >= 0) {
		strip(changes, TRACKED_CELL)
		out.text(unlessEmpty(markup(sub => spliceChildren(sub, part.reader, tcPr, changes, CELL_PROPERTIES))))
	} else {
		out.text(freshContainer(names, "tcPr", changes, CELL_PROPERTIES))
	}

	yield* content()

	if (!endsWithParagraph) out.text(`<${names.el("p")}/>`)

	out.text(`</${names.el("tc")}>`)
}
