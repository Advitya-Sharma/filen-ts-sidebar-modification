import * as Comlink from "comlink"
import SpreadsheetWorker from "@/features/spreadsheet/workers/spreadsheet.worker.ts?worker"
import type { SpreadsheetKind } from "@/features/spreadsheet/lib/model"
import type {
	SpreadsheetFileKind,
	SpreadsheetOpened,
	SpreadsheetUnlock,
	SpreadsheetWorkerApi
} from "@/features/spreadsheet/workers/spreadsheet.worker"
import type { WorkbookRefusal } from "@/features/spreadsheet/lib/xlsxEncryption"
import { idleResource } from "@/lib/idleResource"
import { recoverIfNewerBuild } from "@/lib/appUpdate"

type SpreadsheetRemote = Comlink.Remote<SpreadsheetWorkerApi>

// The worker keeps the hucre and HyperFormula code and whatever heap the largest workbook grew, so it is
// torn down once no document is open and no call is in flight; the next spreadsheet spins up a fresh one.
const SPREADSHEET_WORKER_IDLE_MS = 30_000

// One worker for every open spreadsheet, created the first time one opens.
const spreadsheetWorker = idleResource(
	() => {
		const worker = new SpreadsheetWorker()
		// A deploy that changed this worker removes its old file (appUpdate.ts).
		worker.addEventListener("error", () => {
			recoverIfNewerBuild()
		})

		return { worker, remote: Comlink.wrap<SpreadsheetWorkerApi>(worker) }
	},
	({ worker, remote }) => {
		remote[Comlink.releaseProxy]()
		worker.terminate()
	},
	SPREADSHEET_WORKER_IDLE_MS
)

// The open documents, under ids never reused in this tab: a fresh worker numbers its own from 1 again, and
// a late call from a closed viewer must not reach a newer document. `release` ends the hold that keeps the
// worker while the document is open.
const documents = new Map<number, { workerId: number; release: () => void }>()
let nextId = 1
let nextOpen = 1

// A password-protected workbook waiting for its password. `unlock` tries one (one at a time): opened, or
// `wrongPassword` (and `aborted`, after `cancel` stopped it) to ask again; `discard` gives up on it (a try
// under way stops). The worker is kept until it opens, is refused or is discarded.
export interface SpreadsheetLock {
	unlock: (password: string) => Promise<SpreadsheetUnlock>
	cancel: () => void
	discard: () => void
}

export type SpreadsheetOpen = SpreadsheetOpened | { locked: SpreadsheetLock } | { refused: WorkbookRefusal }

// Opens a document and keeps the worker until closeSpreadsheet (or, for a locked workbook, until it is
// discarded or refused): the hold starts before the bytes are sent, so a long parse is never torn down under.
// `signal` stops a protected workbook's default-password try once the caller has gone (the open rejects); what
// opens regardless is still the caller's to close.
export function openSpreadsheet(bytes: Uint8Array, kind: SpreadsheetFileKind, signal?: AbortSignal): Promise<SpreadsheetOpen> {
	const opened = Promise.withResolvers<SpreadsheetOpen>()
	const token = nextOpen++

	spreadsheetWorker
		.use(async ({ remote }) => {
			const cancel = () => {
				void remote.cancelOpen(token)
			}

			signal?.addEventListener("abort", cancel, { once: true })

			let result: Awaited<ReturnType<typeof remote.open>>

			try {
				result = await remote.open(bytes, kind, token)
			} finally {
				signal?.removeEventListener("abort", cancel)
			}

			if ("refused" in result) {
				opened.resolve(result)

				return
			}

			const closed = Promise.withResolvers<undefined>()
			const release = () => {
				closed.resolve(undefined)
			}

			opened.resolve("locked" in result ? { locked: lockOf(result.locked, release) } : registered(result, release))
			await closed.promise
		})
		.catch(opened.reject)

	return opened.promise
}

// Registers a document the worker opened under a tab-wide id; `release` ends the hold once it closes.
function registered(result: SpreadsheetOpened, release: () => void): SpreadsheetOpened {
	const id = nextId++

	documents.set(id, { workerId: result.id, release })

	return { ...result, id }
}

function lockOf(lockId: number, release: () => void): SpreadsheetLock {
	// Opened, refused or discarded: nothing more is asked of the lock.
	let settled = false

	const discard = (): void => {
		if (settled) {
			return
		}

		settled = true
		// Sent before the hold ends, so the worker outlives it.
		void spreadsheetWorker.use(({ remote }) => remote.discardLocked(lockId))
		release()
	}

	return {
		unlock: password => {
			if (settled) {
				return Promise.reject(new Error("spreadsheet: this workbook is no longer locked"))
			}

			// The hold is still on (not settled), so this is the worker holding the lock.
			return spreadsheetWorker
				.use<SpreadsheetUnlock>(({ remote }) => remote.unlock(lockId, password))
				.then(
					answer => {
						if (settled) {
							// Discarded meanwhile: what the worker opened before it heard is closed there.
							if ("id" in answer) {
								void spreadsheetWorker.use(({ remote }) => remote.close(answer.id))
							}

							return { aborted: true }
						}

						if ("id" in answer) {
							settled = true

							return registered(answer, release)
						}

						if ("refused" in answer) {
							settled = true
							release()
						}

						return answer
					},
					(error: unknown) => {
						// The try failed in the worker: the lock is given up there and the hold ends, whoever holds it.
						discard()

						throw error
					}
				)
		},
		cancel: () => {
			if (!settled) {
				void spreadsheetWorker.use(({ remote }) => remote.cancelUnlock(lockId))
			}
		},
		discard
	}
}

export function closeSpreadsheet(id: number): void {
	const document = documents.get(id)

	if (document === undefined) {
		return
	}

	documents.delete(id)
	// Sent before the hold ends, so the worker outlives it.
	void spreadsheetWorker.use(({ remote }) => remote.close(document.workerId))
	document.release()
}

// A call on an open document. Once the document is closed it rejects at once, as the worker itself would,
// without spinning a worker up.
export function withOpenSpreadsheet<R>(id: number, run: (remote: SpreadsheetRemote, workerId: number) => Promise<R>): Promise<R> {
	const document = documents.get(id)

	if (document === undefined) {
		return Promise.reject(new Error(`spreadsheet: no open document ${String(id)}`))
	}

	return spreadsheetWorker.use(({ remote }) => run(remote, document.workerId))
}

// A call that needs no open document.
export function withSpreadsheetWorker<R>(run: (remote: SpreadsheetRemote) => Promise<R>): Promise<R> {
	return spreadsheetWorker.use(({ remote }) => run(remote))
}

const ZIP = [0x50, 0x4b, 0x03, 0x04]
const OLE2 = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]

function startsWith(bytes: Uint8Array, magic: readonly number[]): boolean {
	return bytes.length >= magic.length && magic.every((byte, index) => bytes[index] === byte)
}

// A file whose name does not say what it is, read from its first bytes: a zip is taken for an .xlsx, an
// OLE2 compound file for an .xls (or a password-protected workbook, which the worker tells apart), anything
// else for text.
export function sniffSpreadsheetKind(bytes: Uint8Array): SpreadsheetKind {
	if (startsWith(bytes, ZIP)) return "xlsx"
	if (startsWith(bytes, OLE2)) return "xls"

	return "csv"
}
