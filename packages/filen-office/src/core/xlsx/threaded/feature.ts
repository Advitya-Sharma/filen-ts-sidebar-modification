import { CellFlag, SHEET_MAX_COLS, TILE_COLS, TILE_ROWS, type SheetId, type SheetRange } from "../../sheet/protocol"
import { defineSheetFeature, type SheetReadyResult, type SheetTileBuilder } from "../../sheet/registry"
import { annotationStore, type AnnotationStore } from "../comments/store"
import type { Slicing } from "../../slicing"
import type { XlsxDocument } from "../edit/document"
import { cellKeyOf } from "../edit/history"
import { addThreadComment, deleteThreadComment, editThreadComment, threadsRefusal } from "./edit"
import type { Thread, ThreadedPart } from "./threads"

// Threaded comments: a corner on the first cell of each thread (CellFlag.THREADED_COMMENT), the thread when asked
// (thread.get), new threads and replies under the session's author, the author's own comments edited, comments and
// whole threads deleted. A sheet whose threaded comments or persons do not read shows none and refuses their edits.

declare module "../../sheet/protocol" {
	interface SheetCommandMap {
		// A comment on a cell: a new thread, or a reply to the cell's thread.
		"thread.add": { readonly sheet: SheetId; readonly row: number; readonly col: number; readonly text: string }
		// A comment of the session's author edited.
		"thread.edit": { readonly sheet: SheetId; readonly id: string; readonly text: string }
		// A reply deleted, or a thread with its first comment.
		"thread.delete": { readonly sheet: SheetId; readonly id: string }
	}
	interface SheetQueryMap {
		"thread.get": { request: SheetCell; result: SheetThread | null }
	}
	interface SheetRefusalReasons {
		// Only its author edits a comment.
		notOwnComment: true
	}
}

export interface SheetThreadComment {
	readonly id: string
	readonly author: string
	// As the file holds it (xsd:dateTime, Excel's local time, no zone).
	readonly date: string
	readonly text: string
	// The session's author wrote it: it may be edited.
	readonly own: boolean
	// It mentions people (their names are in the text): it is not edited here.
	readonly mentions: boolean
}

export interface SheetThread {
	readonly comments: readonly SheetThreadComment[]
	// Resolved.
	readonly done: boolean
}

const byTile = new WeakMap<ThreadedPart, Map<number, Thread[]>>()

function tileOf(row: number, col: number): number {
	return Math.floor(row / TILE_ROWS) * (SHEET_MAX_COLS / TILE_COLS) + Math.floor(col / TILE_COLS)
}

function threadsOf(store: AnnotationStore, sheet: SheetId): ThreadedPart | null {
	return store.failure(sheet, "threads") === null ? store.threadedOf(sheet) : null
}

function flagThreads(document: XlsxDocument, tile: SheetTileBuilder): void {
	const part = threadsOf(annotationStore(document.workbook), tile.sheet)

	if (part === null) return

	let tiles = byTile.get(part)

	if (tiles === undefined) {
		tiles = new Map()

		for (const thread of part.threads.values()) {
			const key = tileOf(thread.row, thread.col)
			const list = tiles.get(key)

			if (list === undefined) tiles.set(key, [thread])
			else list.push(thread)
		}

		byTile.set(part, tiles)
	}

	for (const thread of tiles.get(tileOf(tile.firstRow, tile.firstCol)) ?? [])
		tile.addFlags(tile.entry(thread.row, thread.col), CellFlag.THREADED_COMMENT)
}

async function ready(document: XlsxDocument, sheet: SheetId, slicing: Slicing | undefined): Promise<SheetReadyResult> {
	const store = annotationStore(document.workbook)

	await store.load(document, sheet, slicing)

	const failed = store.failure(sheet, "threads")

	if (failed !== null) throw new Error(failed)

	const ranges: SheetRange[] = []

	for (const thread of threadsOf(store, sheet)?.threads.values() ?? [])
		ranges.push({ startRow: thread.row, startCol: thread.col, endRow: thread.row, endCol: thread.col })

	return { ranges }
}

async function threadAt(document: XlsxDocument, sheet: SheetId, row: number, col: number): Promise<SheetThread | null> {
	if (document.entry(sheet)?.kind !== "worksheet") return null

	const store = annotationStore(document.workbook)

	try {
		await store.load(document, sheet)
	} catch {
		return null
	}

	const thread = threadsOf(store, sheet)?.threads.get(cellKeyOf(row, col))

	if (thread === undefined) return null

	const persons = store.personsOf()?.persons
	const author = document.settings.author

	return {
		comments: [thread.root, ...thread.replies].map(comment => {
			const name = persons?.get(comment.personId)?.displayName ?? ""

			return {
				id: comment.id,
				author: name,
				date: comment.date,
				text: comment.text,
				own: author !== undefined && author !== "" && name === author,
				mentions: comment.mentions
			}
		}),
		done: thread.root.done
	}
}

export const threadsFeature = defineSheetFeature<XlsxDocument>({
	id: "threads",
	commands: {
		"thread.add": {
			info: { write: true, protection: "objects" },
			check: (context, target) => threadsRefusal(context.doc, target.sheet),
			run: (context, command) => addThreadComment(context.doc, command.sheet, command.row, command.col, command.text, context.changes)
		},
		"thread.edit": {
			info: { write: true, protection: "objects" },
			check: (context, target) => threadsRefusal(context.doc, target.sheet),
			run: (context, command) => editThreadComment(context.doc, command.sheet, command.id, command.text, context.changes)
		},
		"thread.delete": {
			info: { write: true, protection: "objects" },
			check: (context, target) => threadsRefusal(context.doc, target.sheet),
			run: (context, command) => deleteThreadComment(context.doc, command.sheet, command.id, context.changes)
		}
	},
	queries: {
		"thread.get": { run: (context, request) => threadAt(context.doc, request.sheet, request.row, request.col) }
	},
	ready: (context, sheet) => ready(context.doc, sheet, context.slicing),
	tiles: (context, tile) => flagThreads(context.doc, tile)
})
