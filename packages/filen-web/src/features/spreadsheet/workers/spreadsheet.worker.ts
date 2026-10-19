import * as Comlink from "comlink"
import { openXlsx, readXls } from "hucre/xlsx"
import type { SpreadsheetDoc, SpreadsheetKind } from "@/features/spreadsheet/lib/model"
import { MAX_SHEET_CELLS, type DocState, type EditOp, type EditResult } from "@/features/spreadsheet/lib/edits"
import { workbookDoc, WorkbookViews } from "@/features/spreadsheet/lib/xlsxView"
import { parseCsvFile } from "@/features/spreadsheet/lib/csvView"
import { XlsxDocument } from "@/features/spreadsheet/lib/xlsxDocument"
import { CsvDocument } from "@/features/spreadsheet/lib/csvDocument"
import { checkZipLimits } from "@/features/spreadsheet/lib/zipLimits"
import { xlsToXlsx } from "@/features/spreadsheet/lib/xlsConvert"
import {
	openEncryptedWorkbook,
	sealWorkbook,
	unlockWorkbook,
	type WorkbookLocked,
	type WorkbookProtection,
	type WorkbookRefusal,
	type WorkbookUnlocked
} from "@/features/spreadsheet/lib/xlsxEncryption"

// Owns every open spreadsheet: the parsed workbook stays here, with its edits and undo history, and the
// page gets views of it (model.ts) and patches after each edit (edits.ts). A file is untrusted input, so
// parsing runs off the main thread, inside limits: a zip whose entries inflate past DECOMPRESSED_LIMIT in
// all (or number past ZIP_ENTRY_LIMIT) or a grid past CELL_LIMIT fails to open rather than taking the tab's
// memory.

export type SpreadsheetFileKind = SpreadsheetKind | "tsv"

const CELL_LIMIT = MAX_SHEET_CELLS
const DECOMPRESSED_LIMIT = 512 * 1024 * 1024
const ZIP_ENTRY_LIMIT = 20_000

// An opened document, and what the page is told of it. `encrypted`: a password-protected workbook, opened
// decrypted (its saves are sealed), and whether its package's integrity was verified.
export interface SpreadsheetOpened {
	id: number
	doc: SpreadsheetDoc
	encrypted?: "verified" | "unchecked"
}

// A protected workbook waiting for its password (unlock), or one that is not opened (no prompt).
export type SpreadsheetOpen = SpreadsheetOpened | { locked: number } | { refused: WorkbookRefusal }

// An unlock's answer: opened, the password was wrong (ask again), refused, or stopped (cancelled: ask again;
// discarded: gone).
export type SpreadsheetUnlock = SpreadsheetOpened | { wrongPassword: true } | { refused: WorkbookRefusal } | { aborted: true }

// An .xls opens to be looked at only: nothing is kept, as nothing can be written back.
const documents = new Map<number, XlsxDocument | CsvDocument>()
// A protected workbook's key, its saves' seal, and its plaintext until the proof read it (the page holds only
// the file as stored, which is sealed).
const protections = new Map<number, { protection: WorkbookProtection; plaintext: Uint8Array | null }>()
// Opens under way, by the page's token: `aborted` stops the default password's try (its spin can be long).
const openings = new Map<number, { aborted: boolean }>()
// Workbooks waiting for a password, and the try under way (its `aborted` stops it).
const locks = new Map<number, { locked: WorkbookLocked; attempt: { aborted: boolean } | null }>()
let nextId = 1

async function openXlsxDocument(bytes: Uint8Array): Promise<XlsxDocument> {
	// The whole archive's declared sizes first: hucre caps each entry, not their sum.
	checkZipLimits(bytes, { maxEntries: ZIP_ENTRY_LIMIT, maxBytes: DECOMPRESSED_LIMIT })

	const document = new XlsxDocument(
		await openXlsx(bytes, { readStyles: true, maxTotalCells: CELL_LIMIT, maxDecompressedBytes: DECOMPRESSED_LIMIT })
	)

	// View-only until proven: the proof runs once the page shows it may edit (writability(id, bytes): its
	// first edit waits for it), and is skipped once it says it won't (viewOnly(id)). Until then only what
	// saving copies is kept; the proof reads the rest again from the bytes it is given.
	document.dropOriginals()

	return document
}

// A decrypted workbook opened as any other; its key is released if it fails to open.
async function openDecrypted({ plaintext, protection, integrity }: WorkbookUnlocked): Promise<SpreadsheetOpened> {
	let document: XlsxDocument

	try {
		document = await openXlsxDocument(plaintext)
	} catch (error) {
		protection.release()

		throw error
	}

	const id = nextId++
	const doc = document.doc()

	documents.set(id, document)
	// Kept only for a proof that may still run: a workbook known at open to stay view-only is never proven.
	protections.set(id, { protection, plaintext: doc.unprovable === undefined ? plaintext : null })

	return { id, doc, encrypted: integrity }
}

// `token` names the open for cancelOpen.
async function open(bytes: Uint8Array, kind: SpreadsheetFileKind, token?: number): Promise<SpreadsheetOpen> {
	// A compound file named or sniffed as a workbook may be a password-protected .xlsx (its name may still say
	// .xls: it then opens read-only, as any .xls).
	if (kind === "xlsx" || kind === "xls") {
		const signal = { aborted: false }

		if (token !== undefined) {
			openings.set(token, signal)
		}

		let encrypted: Awaited<ReturnType<typeof openEncryptedWorkbook>>

		try {
			encrypted = await openEncryptedWorkbook(bytes, signal)
		} finally {
			if (token !== undefined) {
				openings.delete(token)
			}
		}

		switch (encrypted.kind) {
			case "plain":
				break
			case "refused":
				return { refused: encrypted.reason }
			case "aborted":
				// The page went: nothing was opened.
				throw new Error("spreadsheet: opening was cancelled")
			case "needsPassword": {
				const lock = nextId++

				locks.set(lock, { locked: encrypted.locked, attempt: null })

				return { locked: lock }
			}
			case "unlocked":
				return await openDecrypted(encrypted)
		}
	}

	const id = nextId++

	switch (kind) {
		case "xlsx": {
			const document = await openXlsxDocument(bytes)

			documents.set(id, document)

			return { id, doc: document.doc() }
		}
		case "xls": {
			const workbook = await readXls(bytes, { maxTotalCells: CELL_LIMIT })

			return { id, doc: { ...workbookDoc(workbook, new WorkbookViews(workbook.themeColors), false), kind: "xls" } }
		}
		case "csv":
		case "tsv": {
			const { rows, format } = parseCsvFile(bytes, kind === "tsv")

			if (rows.reduce((cells, row) => cells + row.length, 0) > CELL_LIMIT) {
				throw new Error("spreadsheet: too many cells")
			}

			const document = new CsvDocument(rows, format)

			documents.set(id, document)

			return { id, doc: document.doc() }
		}
	}
}

function document(id: number): XlsxDocument | CsvDocument {
	const found = documents.get(id)

	if (found === undefined) {
		throw new Error(`spreadsheet: no open document ${String(id)}`)
	}

	return found
}

// Tries a password on a locked workbook; one try at a time. A wrong one, or a try cancelled, leaves it locked
// for the next.
async function unlock(lockId: number, password: string): Promise<SpreadsheetUnlock> {
	const lock = locks.get(lockId)

	if (lock?.attempt !== null) {
		throw new Error(`spreadsheet: no locked workbook ${String(lockId)} to unlock`)
	}

	const attempt = { aborted: false }
	const discarded = () => locks.get(lockId) !== lock

	lock.attempt = attempt

	try {
		const result = await unlockWorkbook(lock.locked, password, attempt)

		if (!result.ok) {
			if (result.reason === "wrongPassword" || result.reason === "aborted") {
				return result.reason === "wrongPassword" && !discarded() ? { wrongPassword: true } : { aborted: true }
			}

			locks.delete(lockId)

			return { refused: result.reason }
		}

		if (discarded()) {
			result.protection.release()

			return { aborted: true }
		}

		locks.delete(lockId)

		return await openDecrypted(result)
	} finally {
		lock.attempt = null
	}
}

function stopAttempt(lockId: number): void {
	const attempt = locks.get(lockId)?.attempt

	if (attempt !== null && attempt !== undefined) {
		attempt.aborted = true
	}
}

const api = {
	open,
	// The page went before the open answered: a default password's try under way stops.
	cancelOpen: (token: number): void => {
		const signal = openings.get(token)

		if (signal !== undefined) {
			signal.aborted = true
		}
	},
	unlock,
	// Stops the try under way; the workbook stays locked.
	cancelUnlock: (lockId: number): void => {
		stopAttempt(lockId)
	},
	// The page will not unlock this workbook: a try under way stops, and what it read goes.
	discardLocked: (lockId: number): void => {
		stopAttempt(lockId)
		locks.delete(lockId)
	},
	apply: async (id: number, op: EditOp): Promise<EditResult> => {
		const found = document(id)

		return found instanceof XlsxDocument ? await found.edit(op) : found.apply(op)
	},
	undo: (id: number): EditResult => document(id).undo(),
	redo: (id: number): EditResult => document(id).redo(),
	// The file's bytes as they now stand, handed over rather than copied, and the edit state they hold.
	// A protected workbook's are sealed, never handed out as its zip.
	// Its protection is read before anything is awaited: a close meanwhile releases the key, which fails the
	// seal, and never turns the save into its zip.
	serialize: async (id: number): Promise<{ bytes: Uint8Array; version: number }> => {
		const protection = protections.get(id)?.protection
		const serialized = await document(id).serialize()
		const bytes = protection === undefined ? serialized.bytes : await sealWorkbook(protection, serialized.bytes)

		return Comlink.transfer({ bytes, version: serialized.version }, [bytes.buffer as ArrayBuffer])
	},
	// Whether the open document can be saved: for a workbook, once proven from the file's bytes (`source`,
	// a copy the proof lets go; a protected workbook's plaintext is kept here instead, as `source` is sealed)
	// that saving loses nothing (its doc opens with writable false until then); for text, as its format
	// allows.
	writability: async (id: number, source: Uint8Array | null): Promise<boolean> => {
		const found = document(id)

		if (!(found instanceof XlsxDocument)) {
			return found.writable
		}

		const kept = protections.get(id)
		const bytes = kept === undefined ? source : kept.plaintext

		if (kept !== undefined) {
			kept.plaintext = null
		}

		return await found.verifyWritable(bytes ?? undefined)
	},
	// The page will not edit this document: a workbook skips its proof and drops what only saving needs.
	viewOnly: (id: number): void => {
		const found = document(id)
		const kept = protections.get(id)

		if (kept !== undefined) {
			kept.plaintext = null
		}

		if (found instanceof XlsxDocument) {
			found.viewOnly()
		}
	},
	// An .xls converted to an .xlsx (values and merges only: lib/xlsConvert.ts), handed over.
	xlsToXlsx: async (bytes: Uint8Array): Promise<Uint8Array> => {
		const converted = await xlsToXlsx(bytes, CELL_LIMIT)

		return Comlink.transfer(converted, [converted.buffer as ArrayBuffer])
	},
	// The bytes serialize returned at `version` are now the file's.
	markSaved: (id: number, version: number): DocState => document(id).markSaved(version),
	// A protected workbook's key goes with it: a save still under way fails rather than sealing.
	close: (id: number): void => {
		const found = documents.get(id)

		if (found instanceof XlsxDocument) {
			found.close()
		}

		documents.delete(id)
		protections.get(id)?.protection.release()
		protections.delete(id)
	}
}

export type SpreadsheetWorkerApi = typeof api

Comlink.expose(api)
