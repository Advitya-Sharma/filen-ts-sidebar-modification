import type { Node } from "prosemirror-model"
import { decodeUtf8 } from "../../xml/index"
import { NO_SOURCE, docxSchema, endsWithTable } from "../edit/schema"
import { cellOverrides, propertySource, rowOverrides, tableMap, tableOverrides, type TableMap, type TrackStamp } from "../edit/tables/map"
import { Item, RevisionFlag } from "../model/index"
import { Ns } from "../model/namespaces"
import { closeTag, freshContainer, openTag, Out, spliceChildren } from "./markup"
import { refuseBlockMarks, src } from "./source"
import type { StoryWriting } from "./story"
import { DocxWriteError, integer, name as checkName } from "./values"

// Tables, rows and cells. What is as the source has it is copied as its bytes; what table commands changed
// is spliced into a copy of the source's property elements (w:tblPr, w:tblGrid, w:trPr, w:tcPr) in schema
// order, and new rows and cells start from the properties of the source row or cell they were modelled on.
// A cell's width in twips (w:tcW) is its `colwidth` once an edit set one; a grid a row outgrew grows with it
// (and the table's width in twips with its grid).

const marks = docxSchema.marks

// Children of container properties a copy leaves out: tracked changes, whose ids are the original's.
const TRACKED_TABLE = ["tblPrChange", "tblGridChange"]
const TRACKED_ROW = ["ins", "del", "trPrChange", "tblPrExChange"]
const TRACKED_CELL = ["cellIns", "cellDel", "cellMerge", "tcPrChange"]

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

function gridMarkup(w: StoryWriting, grid: readonly number[]): string {
	const names = w.names

	for (const width of grid) integer("w:gridCol w", width, 0)

	return `<${names.el("tblGrid")}>${grid.map(width => names.empty("gridCol", [["w", String(width)]])).join("")}</${names.el("tblGrid")}>`
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
const TABLE_KEYS = new Set(["align", "width", "borders", "look"])
const ROW_KEYS = new Set(["from", "header", "gridBefore", "gridAfter", "ins", "del"])
const CELL_KEYS = new Set(["from", "shading", "vAlign", "ins", "del"])
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

// Children a copy leaves out, unless a change sets them.
function strip(changes: Map<string, string>, locals: readonly string[]): void {
	for (const local of locals) if (!changes.has(local)) changes.set(local, "")
}

// ---- tables ---------------------------------------------------------------------------------------------

function tableChanges(
	w: StoryWriting,
	node: Node,
	tblPr: number,
	styleChanged: boolean,
	grid: readonly number[],
	gridChanged: boolean,
	fresh: boolean
): Map<string, string> {
	const names = w.names
	const overrides = tableOverrides(node)
	const width = w.env.context.store.table(w.part.document, tblPr).width

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
	const styleId = node.attrs["styleId"] as string | null

	if (styleChanged) {
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

	if (item === NO_SOURCE || part.kinds[item] !== Item.Table) {
		const grid = fittedGrid(tableMap(node, part, store), undefined)
		const changes = tableChanges(w, node, -1, node.attrs["styleId"] !== null, grid, true, true)

		out.text(`<${names.el("tbl")}>${freshContainer(names, "tblPr", changes, TABLE_PROPERTIES)}${gridMarkup(w, grid)}`)
		yield* w.writeChildren(out, node, NO_SOURCE, (row, rowOut) => writeRow(w, rowOut, row))
		out.text(`</${names.el("tbl")}>`)

		return
	}

	const row = part.aux[item] ?? -1
	const tblPr = part.tables.tblPr[row] ?? -1
	const tblGrid = part.tables.tblGrid[row] ?? -1
	const original = w.original(node)
	const sourceStyle = original === undefined ? (store.table(document, tblPr).styleId ?? null) : original.attrs["styleId"]
	const written = fittedGrid(tableMap(node, part, store), original === undefined ? undefined : tableMap(original, part, store))
	const gridChanged = !sameNumbers(written, (original?.attrs["grid"] as readonly number[] | undefined) ?? sourceGrid(w, tblGrid))
	const changes = tableChanges(w, node, tblPr, node.attrs["styleId"] !== sourceStyle, written, gridChanged, false)
	const tblPrMarkup = (): string =>
		tblPr >= 0
			? markup(sub => spliceChildren(sub, part.reader, tblPr, changes, TABLE_PROPERTIES))
			: freshContainer(
					names,
					"tblPr",
					new Map([
						[
							"tblW",
							names.empty("tblW", [
								["w", "0"],
								["type", "auto"]
							])
						],
						...changes
					]),
					TABLE_PROPERTIES
				)

	if (!w.owns(node)) {
		// A second table from one source (a split, a move of part of it): its own properties, no ids.
		out.text(`<${names.el("tbl")}>`)

		if (changes.size === 0) {
			w.copyProperties(out, [tblPr], TRACKED_TABLE)
		} else {
			strip(changes, ["tblPrChange"])
			out.text(tblPrMarkup())
		}

		if (!gridChanged && tblGrid >= 0) w.copyProperties(out, [tblGrid], TRACKED_TABLE)
		else out.text(gridMarkup(w, written))

		yield* w.writeChildren(out, node, NO_SOURCE, (child, rowOut) => writeRow(w, rowOut, child))
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
		else if (property === tblGrid && gridChanged) out.text(gridMarkup(w, written))
		else w.copyWhole(out, property)

		from = document.end(property)
	}

	if (tblGrid < 0 && gridChanged) out.text(gridMarkup(w, written))

	yield* w.writeChildren(out, node, item, (child, rowOut) => writeRow(w, rowOut, child), {
		from,
		to: document.contentEnd(element),
		tail: []
	})
	closeTag(out, document, element)
}

// ---- rows -----------------------------------------------------------------------------------------------

function* writeRow(w: StoryWriting, out: Out, node: Node): Generator<void, void> {
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
	const props = w.env.context.store.row(document, trPr)
	const changes = new Map<string, string>()

	checkKeys(overrides, ROW_KEYS, "row")
	checkStamp(overrides.ins)
	checkStamp(overrides.del)

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

	trackChanges(
		w,
		changes,
		node.attrs["revision"] as number,
		owned ? (part.rows.revision[modelRow] ?? 0) : 0,
		["ins", "del"],
		overrides,
		owned
	)

	const cell = (child: Node, cellOut: Out): Generator<void, void> => writeCell(w, cellOut, child)

	if (owned) {
		const element = part.node(own)

		openTag(out, document, element)

		// w:tblPrEx and w:trPr in the order the source has them (CT_Row puts tblPrEx first).
		let from = document.openEnd(element)

		for (const property of [trPr, tblPrEx].sort((a, b) => a - b)) {
			if (property < 0) continue

			if (property === trPr && changes.size > 0)
				out.text(markup(sub => spliceChildren(sub, part.reader, trPr, changes, ROW_PROPERTIES)))
			else w.copyWhole(out, property)

			from = document.end(property)
		}

		if (trPr < 0 && changes.size > 0) out.text(freshContainer(names, "trPr", changes, ROW_PROPERTIES))

		yield* w.writeChildren(out, node, own, cell, { from, to: document.contentEnd(element), tail: [] })
		closeTag(out, document, element)

		return
	}

	out.text(`<${names.el("tr")}>`)

	if (model !== NO_SOURCE && changes.size === 0) {
		// A copy of a source row, or a new row modelled on one: its properties without tracked changes.
		w.copyProperties(out, [trPr, tblPrEx], TRACKED_ROW)
	} else if (model !== NO_SOURCE) {
		if (tblPrEx >= 0) w.copyProperties(out, [tblPrEx], TRACKED_ROW)

		strip(changes, TRACKED_ROW)
		out.text(
			trPr >= 0
				? markup(sub => spliceChildren(sub, part.reader, trPr, changes, ROW_PROPERTIES))
				: freshContainer(names, "trPr", changes, ROW_PROPERTIES)
		)
	} else {
		out.text(freshContainer(names, "trPr", changes, ROW_PROPERTIES))
	}

	yield* w.writeChildren(out, node, NO_SOURCE, cell)
	out.text(`</${names.el("tr")}>`)
}

// ---- cells ----------------------------------------------------------------------------------------------

function* writeCell(w: StoryWriting, out: Out, node: Node): Generator<void, void> {
	w.cells++

	try {
		yield* writeCellContent(w, out, node)
	} finally {
		w.cells--
	}
}

function* writeCellContent(w: StoryWriting, out: Out, node: Node): Generator<void, void> {
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
	const props = w.env.context.store.cell(document, tcPr)
	const span = node.attrs["colspan"] as number
	const vMerge = node.attrs["vMerge"] as string | null
	// A new cell modelled on a source cell takes its properties, not its span and merge.
	const modelled = model !== NO_SOURCE && !isSource
	// As the editor reads them (a span below 1 is 1).
	const baseSpan = isSource ? Math.max(1, props.gridSpan ?? 1) : 1
	const baseVMerge = isSource ? (props.vMerge ?? null) : null
	const changes = new Map<string, string>()

	checkKeys(overrides, CELL_KEYS, "cell")
	checkStamp(overrides.ins)
	checkStamp(overrides.del)
	integer("w:gridSpan", span, 1)

	// A cell carries one tracked change at most (CT_TcPr: cellIns, cellDel or cellMerge).
	if (
		((node.attrs["revision"] as number) & (RevisionFlag.Inserted | RevisionFlag.Deleted)) ===
		(RevisionFlag.Inserted | RevisionFlag.Deleted)
	)
		throw new DocxWriteError("unsupported", "A table cell both inserted and deleted as tracked changes")

	if (overrides.vAlign !== undefined && !V_ALIGN.has(overrides.vAlign))
		throw new DocxWriteError("invalidValue", `w:vAlign: ${JSON.stringify(overrides.vAlign)}`)

	if (modelled) for (const local of ["gridSpan", "vMerge", "hMerge"]) changes.set(local, "")

	if (span !== baseSpan) changes.set("gridSpan", span > 1 ? names.val("gridSpan", String(span)) : "")
	if (vMerge !== baseVMerge)
		changes.set("vMerge", vMerge === null ? "" : vMerge === "continue" ? names.empty("vMerge") : names.val("vMerge", "restart"))

	// A width an edit gave the cell (twips).
	const colwidth = node.attrs["colwidth"] as number | null

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

	trackChanges(
		w,
		changes,
		node.attrs["revision"] as number,
		owned ? (part.cells.revision[modelRow] ?? 0) : 0,
		["cellIns", "cellDel"],
		overrides,
		owned
	)

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
		else if (tcPr >= 0) out.text(markup(sub => spliceChildren(sub, part.reader, tcPr, changes, CELL_PROPERTIES)))
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
		out.text(markup(sub => spliceChildren(sub, part.reader, tcPr, changes, CELL_PROPERTIES)))
	} else {
		out.text(freshContainer(names, "tcPr", changes, CELL_PROPERTIES))
	}

	yield* content()

	if (!endsWithParagraph) out.text(`<${names.el("p")}/>`)

	out.text(`</${names.el("tc")}>`)
}
