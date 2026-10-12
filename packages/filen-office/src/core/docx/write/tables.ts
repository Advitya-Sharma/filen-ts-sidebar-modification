import type { Node } from "prosemirror-model"
import { NO_SOURCE, endsWithTable } from "../edit/schema"
import { Item } from "../model/index"
import { closeTag, openTag, type Out } from "./markup"
import { refuseBlockMarks, src } from "./source"
import type { StoryWriting } from "./story"
import { DocxWriteError } from "./values"

// Children of container properties a copy leaves out: tracked changes, whose ids are the original's.
const TRACKED_TABLE = ["tblPrChange", "tblGridChange"]
const TRACKED_ROW = ["ins", "del", "trPrChange", "tblPrExChange"]
const TRACKED_CELL = ["cellIns", "cellDel", "cellMerge", "tcPrChange"]

export function* writeTable(w: StoryWriting, out: Out, node: Node): Generator<void, void> {
	const part = w.part
	const document = part.document
	const names = w.names

	const item = src(node)

	if (!w.owns(node) && item !== NO_SOURCE && part.kinds[item] === Item.Table) {
		// A second table from one source (a split, a move of part of it): its own properties, no ids.
		w.checkUnchanged(node, ["styleId", "grid", "props"], "Changed table properties")

		const row = part.aux[item] ?? -1

		out.text(`<${names.el("tbl")}>`)
		w.copyProperties(out, [part.tables.tblPr[row] ?? -1, part.tables.tblGrid[row] ?? -1], TRACKED_TABLE)
		yield* w.writeChildren(out, node, NO_SOURCE, (child, rowOut) => writeRow(w, rowOut, child))
		out.text(`</${names.el("tbl")}>`)

		return
	}

	if (!w.owns(node)) {
		if ((node.attrs["props"] ?? null) !== null) throw new DocxWriteError("unsupported", "Changed table properties")

		let columns = (node.attrs["grid"] as readonly number[]).length

		node.forEach(row => {
			let count = 0

			row.forEach(cell => {
				count += cell.attrs["colspan"] as number
			})
			columns = Math.max(columns, count)
		})

		const grid =
			(node.attrs["grid"] as readonly number[]).length > 0
				? (node.attrs["grid"] as readonly number[])
				: Array.from({ length: columns }, () => 2000)
		const styleId = node.attrs["styleId"] as string | null

		out.text(
			`<${names.el("tbl")}><${names.el("tblPr")}>${styleId === null ? "" : names.val("tblStyle", styleId)}${names.empty("tblW", [
				["w", "0"],
				["type", "auto"]
			])}</${names.el("tblPr")}><${names.el("tblGrid")}>${grid.map(width => names.empty("gridCol", [["w", String(width)]])).join("")}</${names.el("tblGrid")}>`
		)
		yield* w.writeChildren(out, node, NO_SOURCE, (row, rowOut) => writeRow(w, rowOut, row))
		out.text(`</${names.el("tbl")}>`)

		return
	}

	const row = part.aux[item] ?? -1

	w.checkUnchanged(node, ["styleId", "grid", "props"], "Changed table properties")

	const element = part.node(item)

	openTag(out, document, element)

	let from = document.openEnd(element)

	for (const property of [part.tables.tblPr[row] ?? -1, part.tables.tblGrid[row] ?? -1].sort((a, b) => a - b)) {
		if (property < 0) continue

		w.copyWhole(out, property)
		from = document.end(property)
	}

	yield* w.writeChildren(out, node, item, (child, rowOut) => writeRow(w, rowOut, child), {
		from,
		to: document.contentEnd(element),
		tail: []
	})
	closeTag(out, document, element)
}

function* writeRow(w: StoryWriting, out: Out, node: Node): Generator<void, void> {
	refuseBlockMarks(node)

	const part = w.part
	const document = part.document
	const own = src(node)

	if (!w.owns(node) && own !== NO_SOURCE && part.kinds[own] === Item.Row) {
		const row = part.aux[own] ?? -1

		out.text(`<${w.names.el("tr")}>`)
		w.copyProperties(out, [part.rows.trPr[row] ?? -1, part.rows.tblPrEx[row] ?? -1], TRACKED_ROW)
		yield* w.writeChildren(out, node, NO_SOURCE, (cell, cellOut) => writeCell(w, cellOut, cell))
		out.text(`</${w.names.el("tr")}>`)

		return
	}

	if (!w.owns(node)) {
		out.text(`<${w.names.el("tr")}>`)
		yield* w.writeChildren(out, node, NO_SOURCE, (cell, cellOut) => writeCell(w, cellOut, cell))
		out.text(`</${w.names.el("tr")}>`)

		return
	}

	const item = src(node)
	const row = part.aux[item] ?? -1
	const element = part.node(item)

	openTag(out, document, element)

	// w:tblPrEx and w:trPr in the order the source has them (CT_Row puts tblPrEx first).
	let from = document.openEnd(element)

	for (const property of [part.rows.trPr[row] ?? -1, part.rows.tblPrEx[row] ?? -1].sort((a, b) => a - b)) {
		if (property < 0) continue

		w.copyWhole(out, property)
		from = document.end(property)
	}

	yield* w.writeChildren(out, node, item, (cell, cellOut) => writeCell(w, cellOut, cell), {
		from,
		to: document.contentEnd(element),
		tail: []
	})
	closeTag(out, document, element)
}

function* writeCell(w: StoryWriting, out: Out, node: Node): Generator<void, void> {
	w.cells++

	try {
		yield* writeCellContent(w, out, node)
	} finally {
		w.cells--
	}
}

function* writeCellContent(w: StoryWriting, out: Out, node: Node): Generator<void, void> {
	refuseBlockMarks(node)

	const part = w.part
	const document = part.document
	const names = w.names
	// Word wants a cell to end with a paragraph, which a table as its last block would not.
	const endsWithParagraph = !endsWithTable(node)

	const own = src(node)

	if (!w.owns(node) && own !== NO_SOURCE && part.kinds[own] === Item.Cell) {
		w.checkUnchanged(node, ["colspan", "vMerge"], "Changed cell merging")
		out.text(`<${names.el("tc")}>`)
		w.copyProperties(out, [part.cells.tcPr[part.aux[own] ?? -1] ?? -1], TRACKED_CELL)
		yield* w.writeChildren(out, node, NO_SOURCE, (child, blockOut) => w.writeBlock(blockOut, child))

		if (!endsWithParagraph) out.text(`<${names.el("p")}/>`)

		out.text(`</${names.el("tc")}>`)

		return
	}

	if (!w.owns(node)) {
		const span = node.attrs["colspan"] as number
		const vMerge = node.attrs["vMerge"] as string | null
		const tcPr = `${span > 1 ? names.val("gridSpan", String(span)) : ""}${vMerge === null ? "" : vMerge === "continue" ? names.empty("vMerge") : names.val("vMerge", vMerge)}`

		out.text(`<${names.el("tc")}>${tcPr === "" ? "" : `<${names.el("tcPr")}>${tcPr}</${names.el("tcPr")}>`}`)
		yield* w.writeChildren(out, node, NO_SOURCE, (child, blockOut) => w.writeBlock(blockOut, child))

		if (!endsWithParagraph) out.text(`<${names.el("p")}/>`)

		out.text(`</${names.el("tc")}>`)

		return
	}

	const item = own
	const row = part.aux[item] ?? -1

	w.checkUnchanged(node, ["colspan", "vMerge"], "Changed cell merging")

	const element = part.node(item)
	const tcPr = part.cells.tcPr[row] ?? -1

	openTag(out, document, element)

	if (tcPr >= 0) w.copyWhole(out, tcPr)

	yield* w.writeChildren(out, node, item, (child, blockOut) => w.writeBlock(blockOut, child), {
		from: tcPr >= 0 ? document.end(tcPr) : document.openEnd(element),
		to: document.contentEnd(element),
		tail: []
	})

	if (!endsWithParagraph) out.text(`<${names.el("p")}/>`)

	closeTag(out, document, element)
}
