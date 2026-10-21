import { CellFlag, SHEET_MAX_COLS, TILE_COLS, TILE_ROWS, type SheetId, type SheetRange } from "../../sheet/protocol"
import { defineSheetFeature, type SheetReadyResult, type SheetTileBuilder } from "../../sheet/registry"
import type { Slicing } from "../../slicing"
import type { XlsxDocument } from "../edit/document"
import { cellKeyOf } from "../edit/history"
import { deleteNotes, notesRefusal, setNote, showNote } from "./edit"
import type { CommentsPart, NoteEntry } from "./notes"
import { annotationStore, type AnnotationStore } from "./store"

// Notes (the legacy comments Excel now calls notes): a red corner on their cells (CellFlag.NOTE), the text when
// asked (note.get), added, edited, deleted and shown all the time. A thread's placeholder note is the thread's, not a
// note: the threads feature shows it. A sheet whose comments part or legacy drawing does not read shows no notes and
// refuses their edits ("notesUnavailable").

declare module "../../sheet/protocol" {
	interface SheetCommandMap {
		// A cell's note: its text replaced, or a new note written under the session's author (Excel starts a new note's
		// text with the author's name and a colon on a line of its own: the app puts that in the text).
		"note.set": { readonly sheet: SheetId; readonly row: number; readonly col: number; readonly text: string }
		// The notes of the cells in `ranges` deleted.
		"note.delete": { readonly sheet: SheetId; readonly ranges: readonly SheetRange[] }
		// A note shown all the time (true), or only when its cell is pointed at.
		"note.show": { readonly sheet: SheetId; readonly row: number; readonly col: number; readonly shown: boolean }
	}
	interface SheetQueryMap {
		"note.get": { request: SheetCell; result: SheetNote | null }
	}
	interface SheetLayoutExtensionMap {
		// The cells whose notes are shown all the time, as row, col pairs.
		readonly shownNotes?: Uint32Array
	}
	interface SheetRefusalReasons {
		// The sheet's notes or threaded comments could not be read (a damaged part beside the sheet; detail: which).
		notesUnavailable: true
		// The edit records its author, and the session has none (SheetSessionSettings.author).
		authorNeeded: true
	}
}

export interface SheetNote {
	readonly text: string
	readonly author: string
	readonly shown: boolean
}

// Notes by tile, per parsed part: a tile visits only its own.
const byTile = new WeakMap<CommentsPart, Map<number, NoteEntry[]>>()

function tileOf(row: number, col: number): number {
	return Math.floor(row / TILE_ROWS) * (SHEET_MAX_COLS / TILE_COLS) + Math.floor(col / TILE_COLS)
}

function tileNotes(part: CommentsPart): Map<number, NoteEntry[]> {
	let tiles = byTile.get(part)

	if (tiles === undefined) {
		tiles = new Map()

		for (const note of part.notes.values()) {
			const key = tileOf(note.row, note.col)
			let list = tiles.get(key)

			if (list === undefined) {
				list = []
				tiles.set(key, list)
			}

			list.push(note)
		}

		byTile.set(part, tiles)
	}

	return tiles
}

// The notes a sheet shows: its comments part's, but threads' placeholders. Null while unread or unavailable.
export function visibleNotes(store: AnnotationStore, sheet: SheetId): { part: CommentsPart; shows: (note: NoteEntry) => boolean } | null {
	if (store.failure(sheet, "notes") !== null || store.failure(sheet, "drawing") !== null) return null

	const part = store.commentsOf(sheet)

	if (part === null) return null

	const threads = store.failure(sheet, "threads") === null ? store.threadedOf(sheet) : null

	return { part, shows: note => note.thread === null || threads?.comments.has(note.thread) !== true }
}

function flagNotes(document: XlsxDocument, tile: SheetTileBuilder): void {
	const notes = visibleNotes(annotationStore(document.workbook), tile.sheet)

	if (notes === null) return

	const list = tileNotes(notes.part).get(tileOf(tile.firstRow, tile.firstCol))

	for (const note of list ?? []) if (notes.shows(note)) tile.addFlags(tile.entry(note.row, note.col), CellFlag.NOTE)
}

function shownNotes(store: AnnotationStore, sheet: SheetId): Uint32Array | undefined {
	const notes = visibleNotes(store, sheet)
	const vml = notes === null ? null : store.vmlOf(sheet)

	if (notes === null || vml === null) return undefined

	const cells: number[] = []

	for (const [key, shape] of vml.notes) {
		const note = notes.part.notes.get(key)

		if (shape.shown && note !== undefined && notes.shows(note)) cells.push(note.row, note.col)
	}

	return cells.length === 0 ? undefined : Uint32Array.from(cells)
}

async function ready(document: XlsxDocument, sheet: SheetId, slicing: Slicing | undefined): Promise<SheetReadyResult> {
	const store = annotationStore(document.workbook)

	await store.load(document, sheet, slicing)

	const failed = store.failure(sheet, "notes") ?? store.failure(sheet, "drawing")

	if (failed !== null) throw new Error(failed)

	const notes = visibleNotes(store, sheet)
	const ranges: SheetRange[] = []

	for (const note of notes?.part.notes.values() ?? [])
		if (notes?.shows(note) === true) ranges.push({ startRow: note.row, startCol: note.col, endRow: note.row, endCol: note.col })

	return { ranges, layout: shownNotes(store, sheet) !== undefined }
}

export const notesFeature = defineSheetFeature<XlsxDocument>({
	id: "notes",
	commands: {
		"note.set": {
			info: { write: true, protection: "objects" },
			check: (context, target) => notesRefusal(context.doc, target.sheet),
			run: (context, command) => setNote(context.doc, command.sheet, command.row, command.col, command.text, context.changes)
		},
		"note.delete": {
			info: { write: true, protection: "objects" },
			check: (context, target) => notesRefusal(context.doc, target.sheet),
			run: (context, command) => deleteNotes(context.doc, command.sheet, command.ranges, context.changes)
		},
		"note.show": {
			info: { write: true, protection: "objects" },
			check: (context, target) => notesRefusal(context.doc, target.sheet),
			run: (context, command) => showNote(context.doc, command.sheet, command.row, command.col, command.shown, context.changes)
		}
	},
	queries: {
		"note.get": {
			run: async (context, request) => {
				const store = annotationStore(context.doc.workbook)

				if (context.doc.entry(request.sheet)?.kind !== "worksheet") return null

				try {
					await store.load(context.doc, request.sheet)
					// Whether it is shown is in the sheet's legacy drawing, which its legacyDrawing names.
					await context.doc.ready(request.sheet)
				} catch {
					return null
				}

				const notes = visibleNotes(store, request.sheet)
				const note = notes?.part.notes.get(cellKeyOf(request.row, request.col))

				if (notes === null || note === undefined || !notes.shows(note)) return null

				return {
					text: note.text,
					author: note.author,
					shown: store.vmlOf(request.sheet)?.notes.get(cellKeyOf(request.row, request.col))?.shown === true
				}
			}
		}
	},
	ready: (context, sheet) => ready(context.doc, sheet, context.slicing),
	tiles: (context, tile) => flagNotes(context.doc, tile),
	layout: {
		shownNotes: (context, sheet) => shownNotes(annotationStore(context.doc.workbook), sheet)
	}
})
