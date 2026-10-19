import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import type { SpreadsheetOpen } from "@/features/spreadsheet/lib/spreadsheetClient"

// The spreadsheet worker's lifetime: kept while a document is open or a call is in flight, torn down once
// idle, and never reached by a call on a document that has closed. The worker and its Comlink proxy are
// plain fakes; spreadsheet.worker.ts itself is covered by the engine tests.

const { WorkerCtor, wrap, terminate, release, recoverIfNewerBuild } = vi.hoisted(() => ({
	recoverIfNewerBuild: vi.fn(),
	WorkerCtor: vi.fn(),
	wrap: vi.fn(),
	terminate: vi.fn(),
	release: vi.fn()
}))

vi.mock("@/features/spreadsheet/workers/spreadsheet.worker.ts?worker", () => ({ default: WorkerCtor }))
vi.mock("@/lib/appUpdate", () => ({ recoverIfNewerBuild }))
vi.mock("comlink", async importOriginal => ({ ...(await importOriginal<typeof import("comlink")>()), wrap }))

const IDLE_MS = 30_000
const DOC = { kind: "csv", sheets: [], activeSheet: 0, styles: [], writable: true }

interface FakeRemote {
	open: ReturnType<typeof vi.fn>
	close: ReturnType<typeof vi.fn>
	markSaved: ReturnType<typeof vi.fn>
	unlock: ReturnType<typeof vi.fn>
	cancelOpen: ReturnType<typeof vi.fn>
	cancelUnlock: ReturnType<typeof vi.fn>
	discardLocked: ReturnType<typeof vi.fn>
}

let remotes: FakeRemote[] = []
// What every open waits on before it answers: settled unless a test holds it.
let openGate: Promise<unknown> = Promise.resolve()
// What the next opens answer instead of a document: a locked or refused workbook.
let openAnswers: unknown[] = []

async function freshClient() {
	vi.resetModules()

	const { releaseProxy } = await import("comlink")

	remotes = []
	openGate = Promise.resolve()
	openAnswers = []
	WorkerCtor.mockReset()
	wrap.mockReset()
	terminate.mockReset()
	release.mockReset()
	recoverIfNewerBuild.mockReset()
	WorkerCtor.mockImplementation(function FakeWorker() {
		return Object.assign(new EventTarget(), { terminate })
	})
	wrap.mockImplementation(() => {
		// Like the real worker, each one numbers its documents from 1.
		let nextWorkerId = 1
		const remote: FakeRemote = {
			open: vi.fn(async () => {
				await openGate

				return openAnswers.shift() ?? { id: nextWorkerId++, doc: DOC }
			}),
			close: vi.fn(() => Promise.resolve()),
			markSaved: vi.fn((workerId: number) => Promise.resolve({ dirty: false, canUndo: false, canRedo: false, workerId })),
			unlock: vi.fn(() => Promise.resolve({ wrongPassword: true })),
			cancelOpen: vi.fn(() => Promise.resolve()),
			cancelUnlock: vi.fn(() => Promise.resolve()),
			discardLocked: vi.fn(() => Promise.resolve())
		}

		remotes.push(remote)

		return { ...remote, [releaseProxy]: release }
	})

	return await import("@/features/spreadsheet/lib/spreadsheetClient")
}

// An open that opened a document (not a locked or refused workbook).
async function opened(pending: Promise<SpreadsheetOpen>): Promise<Extract<SpreadsheetOpen, { id: number }>> {
	const result = await pending

	if (!("id" in result)) {
		throw new Error("expected an opened document")
	}

	return result
}

beforeEach(() => {
	vi.useFakeTimers()
})

afterEach(() => {
	vi.useRealTimers()
})

describe("spreadsheetClient", () => {
	it("keeps the worker while a document is open, and tears it down once idle after it closes", async () => {
		const client = await freshClient()
		const { id } = await opened(client.openSpreadsheet(new Uint8Array([1]), "csv"))

		await vi.advanceTimersByTimeAsync(IDLE_MS * 2)
		expect(terminate).not.toHaveBeenCalled()

		client.closeSpreadsheet(id)
		await vi.advanceTimersByTimeAsync(0)
		expect(remotes[0]?.close).toHaveBeenCalledExactlyOnceWith(1)
		expect(terminate).not.toHaveBeenCalled()

		await vi.advanceTimersByTimeAsync(IDLE_MS)
		expect(release).toHaveBeenCalledOnce()
		expect(terminate).toHaveBeenCalledOnce()
	})

	it("checks for a newer build when the worker fails to start", async () => {
		const client = await freshClient()

		await opened(client.openSpreadsheet(new Uint8Array([1]), "csv"))

		const worker = WorkerCtor.mock.results[0]?.value as EventTarget | undefined

		expect(recoverIfNewerBuild).not.toHaveBeenCalled()
		worker?.dispatchEvent(new Event("error"))
		expect(recoverIfNewerBuild).toHaveBeenCalledOnce()
	})

	it("never tears the worker down under an open still parsing past the idle window", async () => {
		const client = await freshClient()
		let answer: () => void = () => undefined

		openGate = new Promise<void>(resolve => {
			answer = resolve
		})

		const pending = client.openSpreadsheet(new Uint8Array([1]), "csv")

		await vi.advanceTimersByTimeAsync(IDLE_MS * 2)
		expect(terminate).not.toHaveBeenCalled()

		answer()

		const { id } = await opened(pending)

		client.closeSpreadsheet(id)
		await vi.advanceTimersByTimeAsync(IDLE_MS)
		expect(terminate).toHaveBeenCalledOnce()
	})

	it("rejects a call on a closed document without spinning up a worker", async () => {
		const client = await freshClient()
		const { id } = await opened(client.openSpreadsheet(new Uint8Array([1]), "csv"))

		client.closeSpreadsheet(id)
		await vi.advanceTimersByTimeAsync(IDLE_MS)
		expect(terminate).toHaveBeenCalledOnce()

		const run = vi.fn(() => Promise.resolve())

		await expect(client.withOpenSpreadsheet(id, run)).rejects.toThrow("no open document")
		expect(run).not.toHaveBeenCalled()
		expect(WorkerCtor).toHaveBeenCalledOnce()
	})

	it("never lets a closed document's id reach a newer worker's document of the same worker id", async () => {
		const client = await freshClient()
		const first = await opened(client.openSpreadsheet(new Uint8Array([1]), "csv"))

		client.closeSpreadsheet(first.id)
		await vi.advanceTimersByTimeAsync(IDLE_MS)

		const second = await opened(client.openSpreadsheet(new Uint8Array([2]), "csv"))

		// Both workers numbered their document 1; the ids handed out differ.
		expect(WorkerCtor).toHaveBeenCalledTimes(2)
		expect(second.id).not.toBe(first.id)

		await expect(client.withOpenSpreadsheet(first.id, (remote, workerId) => remote.markSaved(workerId, 1))).rejects.toThrow(
			"no open document"
		)
		expect(remotes[1]?.markSaved).not.toHaveBeenCalled()

		await client.withOpenSpreadsheet(second.id, (remote, workerId) => remote.markSaved(workerId, 1))
		expect(remotes[1]?.markSaved).toHaveBeenCalledExactlyOnceWith(1, 1)
	})

	it("keeps the worker for a call still in flight when its document closes", async () => {
		const client = await freshClient()
		const { id } = await opened(client.openSpreadsheet(new Uint8Array([1]), "csv"))
		let answer: (value: unknown) => void = () => undefined

		remotes[0]?.markSaved.mockReturnValueOnce(
			new Promise(resolve => {
				answer = resolve
			})
		)

		const saving = client.withOpenSpreadsheet(id, (remote, workerId) => remote.markSaved(workerId, 1))

		client.closeSpreadsheet(id)
		await vi.advanceTimersByTimeAsync(IDLE_MS * 2)
		expect(terminate).not.toHaveBeenCalled()

		answer({ dirty: false, canUndo: false, canRedo: false })
		await saving
		await vi.advanceTimersByTimeAsync(IDLE_MS)
		expect(terminate).toHaveBeenCalledOnce()
	})

	it("runs a one-off call in its own use, disposed once idle", async () => {
		const client = await freshClient()

		await client.withSpreadsheetWorker(() => Promise.resolve())
		expect(terminate).not.toHaveBeenCalled()
		await vi.advanceTimersByTimeAsync(IDLE_MS)
		expect(terminate).toHaveBeenCalledOnce()
	})
})

describe("spreadsheetClient: password-protected workbooks", () => {
	async function lockedClient() {
		const client = await freshClient()

		openAnswers = [{ locked: 7 }]

		const result = await client.openSpreadsheet(new Uint8Array([1]), "xlsx")

		if (!("locked" in result)) {
			throw new Error("expected a locked workbook")
		}

		return { client, lock: result.locked, remote: remotes[0] }
	}

	it("keeps the worker while a workbook waits for its password, through wrong passwords and cancelled tries", async () => {
		const { lock, remote } = await lockedClient()

		await expect(lock.unlock("wrong")).resolves.toEqual({ wrongPassword: true })
		remote?.unlock.mockResolvedValueOnce({ aborted: true })
		lock.cancel()
		await expect(lock.unlock("slow")).resolves.toEqual({ aborted: true })
		await vi.advanceTimersByTimeAsync(IDLE_MS * 2)

		expect(remote?.cancelUnlock).toHaveBeenCalledWith(7)
		expect(remote?.unlock).toHaveBeenNthCalledWith(1, 7, "wrong")
		expect(terminate).not.toHaveBeenCalled()
	})

	it("hands an unlocked workbook a tab-wide id, held until it closes", async () => {
		const { client, lock, remote } = await lockedClient()

		remote?.unlock.mockResolvedValueOnce({ id: 3, doc: DOC, encrypted: true })

		const answer = await lock.unlock("right")

		if (!("id" in answer)) {
			throw new Error("expected the workbook opened")
		}

		expect(answer.encrypted).toBe(true)
		await expect(lock.unlock("again")).rejects.toThrow("no longer locked")
		await vi.advanceTimersByTimeAsync(IDLE_MS * 2)
		expect(terminate).not.toHaveBeenCalled()

		client.closeSpreadsheet(answer.id)
		await vi.advanceTimersByTimeAsync(0)
		expect(remote?.close).toHaveBeenCalledExactlyOnceWith(3)
		await vi.advanceTimersByTimeAsync(IDLE_MS)
		expect(terminate).toHaveBeenCalledOnce()
	})

	it("lets the worker go once a locked workbook is discarded or refused", async () => {
		const discarded = await lockedClient()
		const terminated = terminate.mock.calls.length

		discarded.lock.discard()
		discarded.lock.discard()
		await vi.advanceTimersByTimeAsync(0)
		expect(discarded.remote?.discardLocked).toHaveBeenCalledExactlyOnceWith(7)
		await vi.advanceTimersByTimeAsync(IDLE_MS)
		expect(terminate).toHaveBeenCalledTimes(terminated + 1)

		const refused = await lockedClient()

		refused.remote?.unlock.mockResolvedValueOnce({ refused: "damaged" })
		await expect(refused.lock.unlock("right")).resolves.toEqual({ refused: "damaged" })
		await vi.advanceTimersByTimeAsync(IDLE_MS)
		expect(terminate).toHaveBeenCalledOnce()
		expect(refused.remote?.discardLocked).not.toHaveBeenCalled()
	})

	it("closes, in the worker, a workbook a try opened after the page discarded it", async () => {
		const { lock, remote } = await lockedClient()
		let answer: (value: unknown) => void = () => undefined

		remote?.unlock.mockReturnValueOnce(
			new Promise(resolve => {
				answer = resolve
			})
		)

		const trying = lock.unlock("right")

		lock.discard()
		answer({ id: 4, doc: DOC, encrypted: true })

		await expect(trying).resolves.toEqual({ aborted: true })
		await vi.advanceTimersByTimeAsync(0)
		expect(remote?.close).toHaveBeenCalledExactlyOnceWith(4)
	})

	it("cancels an open in the worker once its caller has gone", async () => {
		const client = await freshClient()
		let answer: () => void = () => undefined

		openGate = new Promise<void>(resolve => {
			answer = resolve
		})

		const leaving = new AbortController()
		const pending = client.openSpreadsheet(new Uint8Array([1]), "xlsx", leaving.signal)

		await vi.advanceTimersByTimeAsync(0)

		const token = remotes[0]?.open.mock.calls[0]?.[2] as number

		leaving.abort()
		expect(remotes[0]?.cancelOpen).toHaveBeenCalledExactlyOnceWith(token)
		answer()
		await pending

		// Once answered, its caller's signal is no longer heard.
		const later = new AbortController()

		await client.openSpreadsheet(new Uint8Array([2]), "xlsx", later.signal)
		later.abort()
		expect(remotes[0]?.cancelOpen).toHaveBeenCalledOnce()
	})

	it("passes a refused workbook on without holding the worker", async () => {
		const client = await freshClient()

		openAnswers = [{ refused: "rightsManaged" }]

		await expect(client.openSpreadsheet(new Uint8Array([1]), "xlsx")).resolves.toEqual({ refused: "rightsManaged" })
		await vi.advanceTimersByTimeAsync(IDLE_MS)
		expect(terminate).toHaveBeenCalledOnce()
	})

	it("gives the lock up and ends the hold when a try fails in the worker", async () => {
		const { lock, remote } = await lockedClient()

		remote?.unlock.mockRejectedValueOnce(new Error("spreadsheet: too many cells"))
		await expect(lock.unlock("right")).rejects.toThrow("too many cells")
		expect(remote?.discardLocked).toHaveBeenCalledWith(7)
		await vi.advanceTimersByTimeAsync(IDLE_MS)
		expect(terminate).toHaveBeenCalledOnce()
		await expect(lock.unlock("right")).rejects.toThrow("no longer locked")
	})
})
