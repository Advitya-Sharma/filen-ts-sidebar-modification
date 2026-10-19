import * as Comlink from "comlink"
import type { SpreadsheetFileKind, SpreadsheetOpened } from "@/features/spreadsheet/workers/spreadsheet.worker"
import { MAX_SHEET_CELLS, type DocState, type EditOp, type EditResult } from "@/features/spreadsheet/lib/edits"
import { parseCsvFile } from "@/features/spreadsheet/lib/csvView"
import { CsvDocument } from "@/features/spreadsheet/lib/csvDocument"

// Owns every open CSV and TSV file, as spreadsheet.worker.ts owns workbooks, without the workbook reader and
// the formula engine a text file never needs: opening one loads only this. Answers the same calls for an
// open document.

const documents = new Map<number, CsvDocument>()
let nextId = 1

function document(id: number): CsvDocument {
	const found = documents.get(id)

	if (found === undefined) {
		throw new Error(`spreadsheet: no open document ${String(id)}`)
	}

	return found
}

const api = {
	// A .tsv is read tab-separated; anything else sent here (CSV, or a file sniffed as text) by its separator.
	open: (bytes: Uint8Array, kind: SpreadsheetFileKind): SpreadsheetOpened => {
		const { rows, format } = parseCsvFile(bytes, kind === "tsv")

		if (rows.reduce((cells, row) => cells + row.length, 0) > MAX_SHEET_CELLS) {
			throw new Error("spreadsheet: too many cells")
		}

		const id = nextId++
		const opened = new CsvDocument(rows, format)

		documents.set(id, opened)

		return { id, doc: opened.doc() }
	},
	// Opening text has nothing to stop.
	cancelOpen: (): void => undefined,
	apply: (id: number, op: EditOp): EditResult => document(id).apply(op),
	undo: (id: number): EditResult => document(id).undo(),
	redo: (id: number): EditResult => document(id).redo(),
	serialize: (id: number): { bytes: Uint8Array; version: number } => {
		const serialized = document(id).serialize()

		return Comlink.transfer(serialized, [serialized.bytes.buffer as ArrayBuffer])
	},
	// Text is never proven: it saves as its format allows, and has nothing only a save needs to let go.
	writability: (id: number): boolean => document(id).writable,
	viewOnly: (): void => undefined,
	// A text file's one sheet is never deferred.
	sheet: (id: number): never => {
		throw new Error(`spreadsheet: document ${String(id)} has no deferred sheet`)
	},
	markSaved: (id: number, version: number): DocState => document(id).markSaved(version),
	close: (id: number): void => {
		documents.delete(id)
	}
}

export type CsvWorkerApi = typeof api

Comlink.expose(api)
