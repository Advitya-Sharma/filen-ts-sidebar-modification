import { axisLimit, type AxisShift } from "./axis"
import type { RefConstruct } from "./catalogue"
import { childByLocal, type Edited, type RefJob } from "./markup"

// Objects anchored to cells (ECMA-376 Part 1 §20.5.2.33 twoCellAnchor, §20.5.2.24 oneCellAnchor, §20.5.3.1
// editAs; CT_ObjectAnchor's moveWithCells / sizeWithCells for controls and OLE objects): "twoCell" moves and sizes
// with its cells, "oneCell" moves with its top-left cell and keeps its size, "absolute" stays. An edge at offset
// 0 of a row lies on the boundary above it: an object's top edge there belongs to that row, its bottom edge to the
// row above, so an insert at the boundary moves the object and does not grow it.

export type AnchorMode = "twoCell" | "oneCell" | "absolute"

export interface AnchorPoint {
	index: number
	offset: number
}

// The two edges after the edit (to null for a one-cell anchor), or the same objects when nothing moved.
export function shiftEdges(
	from: AnchorPoint,
	to: AnchorPoint | null,
	mode: AnchorMode,
	shift: AxisShift
): { from: AnchorPoint; to: AnchorPoint | null } {
	if (mode === "absolute") return { from, to }

	const end = shift.at + shift.count
	const last = axisLimit(shift.axis) - 1
	const clamp = (point: AnchorPoint): AnchorPoint => (point.index > last ? { index: last, offset: point.offset } : point)

	if (mode === "oneCell" || to === null) {
		let moved: number

		if (shift.type === "insert") moved = from.index >= shift.at ? shift.count : 0
		else moved = from.index >= end ? -shift.count : from.index >= shift.at ? shift.at - from.index : 0

		if (moved === 0) return { from, to }

		return {
			from: clamp({ index: from.index + moved, offset: from.offset }),
			to: to === null ? null : clamp({ index: to.index + moved, offset: to.offset })
		}
	}

	const top = (point: AnchorPoint): AnchorPoint => {
		if (shift.type === "insert")
			return point.index >= shift.at ? clamp({ index: point.index + shift.count, offset: point.offset }) : point
		if (point.index >= end) return { index: point.index - shift.count, offset: point.offset }

		return point.index >= shift.at ? { index: shift.at, offset: 0 } : point
	}
	// The bottom edge at offset 0 belongs to the row above.
	const bottom = (point: AnchorPoint): AnchorPoint => {
		const boundary = point.offset === 0 && point.index > 0

		if (shift.type === "insert") {
			const moves = boundary ? point.index > shift.at : point.index >= shift.at

			return moves ? clamp({ index: point.index + shift.count, offset: point.offset }) : point
		}

		if (boundary ? point.index > end : point.index >= end) return { index: point.index - shift.count, offset: point.offset }
		if (boundary ? point.index > shift.at : point.index >= shift.at) return { index: shift.at, offset: 0 }

		return point
	}

	return { from: top(from), to: bottom(to) }
}

function readPoint(edited: Edited, node: number, rows: boolean): { point: AnchorPoint; index: number; offset: number } | null {
	const { document } = edited
	const index = childByLocal(document, node, rows ? "row" : "col")
	const offset = childByLocal(document, node, rows ? "rowOff" : "colOff")

	if (index < 0) return null

	const value = Number(document.textContent(index).trim())
	const off = offset < 0 ? 0 : Number(document.textContent(offset).trim())

	if (!Number.isInteger(value) || !Number.isFinite(off)) return null

	return { point: { index: value, offset: off }, index, offset }
}

function writePoint(edited: Edited, read: { point: AnchorPoint; index: number; offset: number }, next: AnchorPoint): boolean {
	if (next === read.point || (next.index === read.point.index && next.offset === read.point.offset)) return false

	edited.editor.setText(read.index, String(next.index))

	if (read.offset >= 0 && next.offset !== read.point.offset) edited.editor.setText(read.offset, String(next.offset))

	return true
}

// An anchor element holding from (and to) markers: xdr:twoCellAnchor, xdr:oneCellAnchor, a control's or OLE
// object's anchor. Unreadable markers are left (an object without cells to follow).
export function shiftAnchor(
	edited: Edited,
	node: number,
	job: RefJob,
	construct: RefConstruct,
	mode: AnchorMode = anchorMode(edited, node)
): void {
	const shift = job.shift

	if (shift === null || mode === "absolute") return

	const { document } = edited
	const rows = shift.axis === "rows"
	const fromNode = childByLocal(document, node, "from")
	const toNode = childByLocal(document, node, "to")
	const from = fromNode < 0 ? null : readPoint(edited, fromNode, rows)
	const to = toNode < 0 ? null : readPoint(edited, toNode, rows)

	if (from === null) return

	const next = shiftEdges(from.point, to?.point ?? null, to === null ? "oneCell" : mode, shift)
	let changed = writePoint(edited, from, next.from)

	if (to !== null && next.to !== null) changed = writePoint(edited, to, next.to) || changed
	if (changed) job.touch(construct)
}

function anchorMode(edited: Edited, node: number): AnchorMode {
	const { document } = edited
	const local = document.localName(node)

	if (local === "absoluteAnchor") return "absolute"
	if (local === "oneCellAnchor") return "oneCell"

	if (local === "anchor") {
		const flag = (name: string) => /^(1|true)$/.test(document.attribute(node, name) ?? "")

		return flag("sizeWithCells") ? "twoCell" : flag("moveWithCells") ? "oneCell" : "absolute"
	}

	const editAs = document.attribute(node, "editAs")

	return editAs === "oneCell" ? "oneCell" : editAs === "absolute" ? "absolute" : "twoCell"
}

// A legacy drawing (VML) x:Anchor: "LeftColumn, LeftOffset, TopRow, TopOffset, RightColumn, RightOffset,
// BottomRow, BottomOffset". The shape moves and sizes with its cells.
export function shiftVmlAnchor(text: string, shift: AxisShift): string | null {
	const values = text.split(",").map(part => Number(part.trim()))

	if (values.length !== 8 || values.some(value => !Number.isInteger(value))) return null

	const rows = shift.axis === "rows"
	const base = rows ? 2 : 0
	const from = { index: values[base] ?? 0, offset: values[base + 1] ?? 0 }
	const to = { index: values[base + 4] ?? 0, offset: values[base + 5] ?? 0 }
	const next = shiftEdges(from, to, "twoCell", shift)

	if (next.from === from && next.to === to) return text

	values[base] = next.from.index
	values[base + 1] = next.from.offset
	values[base + 4] = next.to?.index ?? to.index
	values[base + 5] = next.to?.offset ?? to.offset

	return values.join(", ")
}
