import { SHEET_MAX_COLS, SHEET_MAX_ROWS, type SheetRange, type SheetRefusal } from "../../sheet/protocol"
import { decodeUtf8, encodeUtf8, XmlEditor, type XmlDocument } from "../../xml/index"
import { AnnotationWrites, COMMENTS_TYPE, THREADED_TYPE, VML_TYPE } from "../comments/parts"
import { annotationStore } from "../comments/store"
import type { XlsxSheetModel } from "../model/types"
import { childElements, descendants, editMarkup } from "../refs/markup"
import { RefRefused } from "../refs/catalogue"
import { formatRef, parseRef } from "../refs/sqref"
import { refState } from "../refs/state"
import type { XlsxDocument } from "./document"

// What a sort moves besides the cells, as Excel moves it ("Sort data in a range or table", support.microsoft.com):
// the notes (their comments, their shapes in the legacy drawing, a box shown all the time going with its cell),
// threaded comments and hyperlinks of the moved lines go with them. A hyperlink over more than one moved line, or
// over a moved line and cells the sort leaves, cannot go anywhere whole: the sort refuses.

// The lines a sort moves, as absolute rows (columns for a left-to-right sort): each source line's target. Cells move
// only between `span`'s cross-axis bounds.
export interface SortLineMap {
	readonly rows: boolean
	readonly span: SheetRange
	readonly targets: ReadonlyMap<number, number>
}

function at(map: SortLineMap, row: number, col: number): { row: number; col: number } | null {
	const { rows, span } = map
	const across = rows ? col : row

	if (across < (rows ? span.startCol : span.startRow) || across > (rows ? span.endCol : span.endRow)) return null

	const target = map.targets.get(rows ? row : col)

	if (target === undefined) return null

	return rows ? { row: target, col } : { row, col: target }
}

// A range after the sort: moved when it lies on one moved line within the span, as it is when it meets no moved
// line, null when it cannot move whole.
function movedRange(map: SortLineMap, range: SheetRange): SheetRange | "kept" | null {
	const { rows, span } = map
	const first = rows ? range.startRow : range.startCol
	const last = rows ? range.endRow : range.endCol
	const crossFirst = rows ? range.startCol : range.startRow
	const crossLast = rows ? range.endCol : range.endRow
	const spanFirst = rows ? span.startCol : span.startRow
	const spanLast = rows ? span.endCol : span.endRow

	if (crossLast < spanFirst || crossFirst > spanLast) return "kept"

	let moving = false

	// Whichever is shorter: the range's lines, or the moved ones.
	if (last - first < map.targets.size) for (let line = first; line <= last && !moving; line++) moving = map.targets.has(line)
	else for (const line of map.targets.keys()) if (line >= first && line <= last) moving = true

	if (!moving) return "kept"
	if (first !== last || crossFirst < spanFirst || crossLast > spanLast) return null

	const target = map.targets.get(first) ?? first

	return rows ? { ...range, startRow: target, endRow: target } : { ...range, startCol: target, endCol: target }
}

function moveRefAttribute(map: SortLineMap, editor: XmlEditor, document: XmlDocument, node: number, name: string): boolean {
	const range = parseRef(document.attribute(node, name) ?? "")

	if (range === null || range.startRow !== range.endRow || range.startCol !== range.endCol) return false

	const to = at(map, range.startRow, range.startCol)

	if (to === null) return false

	editor.setAttribute(node, name, formatRef({ startRow: to.row, startCol: to.col, endRow: to.row, endCol: to.col }))

	return true
}

// Comments (§18.7.3 comment ref) and threaded comments ([MS-XLSX] §2.6.204 ref), each a cell.
function moveCellRefs(map: SortLineMap, document: XmlDocument, element: string): Uint8Array | null {
	const editor = new XmlEditor(document)
	let moved = false

	for (const node of descendants(document, document.documentElement()))
		if (document.localName(node) === element && moveRefAttribute(map, editor, document, node, "ref")) moved = true

	return moved ? editor.serialize() : null
}

// Note shapes ([MS-OI29500] §2.1.1716 x:ClientData ObjectType="Note"): x:Row and x:Column name the cell; its box
// (x:Anchor: LeftColumn, LeftOffset, TopRow, TopOffset, RightColumn, RightOffset, BottomRow, BottomOffset) moves as
// far as the cell does, kept on the sheet.
function moveNoteShapes(map: SortLineMap, document: XmlDocument): Uint8Array | null {
	const editor = new XmlEditor(document)
	let moved = false

	for (const node of descendants(document, document.documentElement())) {
		if (document.localName(node) !== "ClientData" || document.attribute(node, "ObjectType") !== "Note") continue

		const children = childElements(document, node)
		const rowNode = children.find(child => document.localName(child) === "Row") ?? -1
		const colNode = children.find(child => document.localName(child) === "Column") ?? -1
		const anchorNode = children.find(child => document.localName(child) === "Anchor") ?? -1

		if (rowNode < 0 || colNode < 0) continue

		const row = Number(document.textContent(rowNode).trim())
		const col = Number(document.textContent(colNode).trim())

		if (!Number.isInteger(row) || !Number.isInteger(col)) continue

		const to = at(map, row, col)

		if (to === null) continue

		moved = true

		if (to.row !== row) editor.setText(rowNode, String(to.row))
		if (to.col !== col) editor.setText(colNode, String(to.col))

		if (anchorNode >= 0) {
			const values = document
				.textContent(anchorNode)
				.split(",")
				.map(part => Number(part.trim()))

			if (values.length === 8 && values.every(value => Number.isInteger(value))) {
				const by = map.rows ? to.row - row : to.col - col
				const base = map.rows ? 2 : 0
				const last = (map.rows ? SHEET_MAX_ROWS : SHEET_MAX_COLS) - 1

				values[base] = Math.max(0, Math.min(last, (values[base] ?? 0) + by))
				values[base + 4] = Math.max(0, Math.min(last, (values[base + 4] ?? 0) + by))
				editor.setText(anchorNode, values.join(", "))
			}
		}
	}

	return moved ? editor.serialize() : null
}

// The hyperlinks section's occurrences after the sort, null when no link moves.
function moveLinks(map: SortLineMap, occurrences: readonly string[]): string[] | SheetRefusal | null {
	let moved = false
	const out: string[] = []

	for (const markup of occurrences) {
		if (markup.trim() === "") {
			out.push(markup)

			continue
		}

		const { document, editor } = editMarkup(encodeUtf8(markup), "hyperlinks")

		for (const node of childElements(document, document.documentElement())) {
			if (document.localName(node) !== "hyperlink") continue

			const range = parseRef(document.attribute(node, "ref") ?? "")

			if (range === null) continue

			const next = movedRange(map, range)

			if (next === null) return { reason: "structureLocked", detail: "hyperlinks" }
			if (next === "kept") continue

			editor.setAttribute(node, "ref", formatRef(next))
			moved = true
		}

		out.push(editor.changed ? decodeUtf8(editor.serialize()) : markup)
	}

	return moved ? out : null
}

// The sheet's notes, threaded comments and hyperlinks moved with the lines a sort moves: the writes (applied), null
// when none moves, or why they cannot move.
export async function moveAnnotations(
	document: XlsxDocument,
	model: XlsxSheetModel,
	map: SortLineMap
): Promise<AnnotationWrites | { readonly refused: SheetRefusal } | null> {
	const workbook = document.workbook
	const store = annotationStore(workbook)
	const sheet = model.id

	await store.load(document, sheet)

	const parts = store.parts(sheet)
	const failure =
		(parts.comments === null ? null : store.failure(sheet, "notes")) ??
		(parts.vml === null ? null : store.failure(sheet, "drawing")) ??
		(parts.threaded === null ? null : store.failure(sheet, "threads"))

	if (failure !== null) return { refused: { reason: "notesUnavailable", detail: failure } }

	let links: string[] | null

	try {
		const moved = moveLinks(map, refState(workbook).sectionMarkup(model, "hyperlinks"))

		if (moved !== null && !Array.isArray(moved)) return { refused: moved }

		links = moved
	} catch (error) {
		if (error instanceof RefRefused) return { refused: error.refusal }

		throw error
	}

	const comments = parts.comments === null ? null : store.commentsOf(sheet)
	const vml = parts.vml === null ? null : store.vmlOf(sheet)
	const threaded = parts.threaded === null ? null : store.threadedOf(sheet)
	const movedComments = comments === null ? null : moveCellRefs(map, comments.document, "comment")
	const movedShapes = vml === null ? null : moveNoteShapes(map, vml.document)
	const movedThreads = threaded === null ? null : moveCellRefs(map, threaded.document, "threadedComment")

	if (links === null && movedComments === null && movedShapes === null && movedThreads === null) return null

	const writes = new AnnotationWrites(workbook)

	if (parts.comments !== null && movedComments !== null) writes.setPart(parts.comments, COMMENTS_TYPE, movedComments)
	if (parts.vml !== null && movedShapes !== null) writes.setPart(parts.vml, VML_TYPE, movedShapes)
	if (parts.threaded !== null && movedThreads !== null) writes.setPart(parts.threaded, THREADED_TYPE, movedThreads)
	if (links !== null) writes.setSection(model, "hyperlinks", links)

	return writes
}
