import { MessageChannel, type MessagePort } from "node:worker_threads"
import { describe, expect, it } from "vitest"
import type { SheetHost, SheetOpenResult, SheetSessionApi } from "../core/sheet/protocol"
import { checkSheetPassword, sheetJobs, sheetWorkerHost, type SheetWorkerLike } from "../dom/sheet/client"
import { sheetPasswordJob } from "../dom/sheet/jobs"
import { workbookOpener } from "../dom/sheet/host"
import { serveSheetWorker, type SheetWorkerJob } from "../dom/sheet/worker"
import { SHEET_PDF_JOB, sheetPdf, type SheetPdfRequest } from "../react/sheet/print/pdfClient"
import { sheetPdfJob } from "../react/sheet/print/pdfJob"
import { DEFAULT_PAGE_SETUP } from "../core/sheet/print/setup"
import { nodeFontAssets } from "./sheetPdfFixtures"
import { agileStreams, encryptedOfficeFile } from "./encryption/testing"
import { BASIC, buildFixture, firstTile } from "./sheetSessionFixtures"

// Jobs a sheet worker runs beside a session: progress reaches the page, an abort reaches the job, a protected file's
// password is checked against it as stored, and nothing a job reads is held for the page's grid.

function workerPair(jobs: Readonly<Record<string, SheetWorkerJob>>): { host: SheetHost; close: () => void } {
	const { port1, port2 } = new MessageChannel()
	const page: SheetWorkerLike = {
		postMessage: (message, transfer) => port1.postMessage(message, transfer ?? []),
		addEventListener: (type, listener) => {
			if (type === "message") port1.on("message", (data: unknown) => listener({ data }))
		}
	}

	serveSheetWorker(
		{
			post: (message, transfer) => port2.postMessage(message, [...transfer]),
			listen: onMessage => {
				port2.on("message", onMessage)
			}
		},
		workbookOpener(),
		jobs
	)

	return {
		host: sheetWorkerHost(page),
		close: () => {
			;(port1 as MessagePort).close()
			;(port2 as MessagePort).close()
		}
	}
}

function opened(result: SheetOpenResult): SheetSessionApi {
	if (result.kind !== "opened") throw new Error(`expected an open, got ${result.kind}`)

	return result.session
}

describe("sheet worker jobs", () => {
	it("runs a job beside the session, with its progress on the page and its answer back", async () => {
		const { host, close } = workerPair({
			count: async (args, context) => {
				const layout = await context.session.layout(0)

				context.progress({ step: 1 })
				context.progress({ step: 2 })

				return { args, rows: layout.rowCount }
			}
		})
		const session = opened(await host.open(buildFixture(BASIC), { kind: "xlsx" }, () => undefined))
		const jobs = sheetJobs(session)
		const progress: unknown[] = []

		expect(jobs).not.toBeNull()
		expect(await jobs?.job("count", { x: 1 }, { onProgress: value => progress.push(value) })).toEqual({ args: { x: 1 }, rows: 3 })
		expect(progress).toEqual([{ step: 1 }, { step: 2 }])
		await expect(jobs?.job("missing", null)).rejects.toThrow("Unknown job missing")
		host.dispose()
		close()
	})

	it("aborts a job when the page's signal does", async () => {
		const { host, close } = workerPair({
			wait: async (_args, context) => {
				while (!context.signal.aborted) await new Promise(resolve => setTimeout(resolve, 1))

				throw new Error("aborted in the worker")
			}
		})
		const session = opened(await host.open(buildFixture(BASIC), { kind: "xlsx" }, () => undefined))
		const controller = new AbortController()
		const running = sheetJobs(session)?.job("wait", null, { signal: controller.signal })

		controller.abort()
		await expect(running).rejects.toThrow("aborted in the worker")
		host.dispose()
		close()
	})

	it("checks a protected workbook's password against the file as stored, keeping nothing", async () => {
		const { host, close } = workerPair(sheetPasswordJob())
		const stored = await encryptedOfficeFile(agileStreams(buildFixture(BASIC), "secret"))
		const locked = await host.open(stored.slice(), { kind: "xlsx" }, () => undefined)

		if (locked.kind !== "needsPassword") throw new Error(locked.kind)

		const unlocked = await locked.locked.unlock("secret")

		if (!unlocked.ok) throw new Error(unlocked.reason)

		expect(await checkSheetPassword(unlocked.session, stored, "wrong")).toBe("wrong")
		expect(await checkSheetPassword(unlocked.session, stored, "secret")).toBe("accepted")

		const defaulted = await encryptedOfficeFile(agileStreams(buildFixture(BASIC), "VelvetSweatshop"))

		expect(await checkSheetPassword(unlocked.session, defaulted, "")).toBe("default")
		expect(await checkSheetPassword(unlocked.session, buildFixture(BASIC), "secret")).toBe("failed")
		host.dispose()
		close()
	}, 60_000)

	it("gives tiles for a print without holding them for the grid or counting their looks as sent", async () => {
		const { host, close } = workerPair({})
		const session = opened(await host.open(buildFixture(BASIC), { kind: "xlsx" }, () => undefined))
		const printed = await session.query("print.tiles", { sheet: 0, coordinates: Uint32Array.of(0, 0), known: 0 })

		expect(printed.tiles).toHaveLength(1)
		expect(printed.styles.from).toBe(0)
		expect(printed.styles.styles.length).toBeGreaterThan(0)

		// The grid's own first tile still brings every look it needs.
		const tile = await firstTile(session, 0)

		expect(tile.count).toBe(printed.tiles[0]?.count)
		expect(sheetJobs({ ...session, job: undefined } as unknown as SheetSessionApi)).toBeNull()
		host.dispose()
		close()
	})

	it("answers the print rows again after a page took them across (each answer owns its buffers)", async () => {
		const { host, close } = workerPair({
			rows: async (_args, context) => {
				const rows = await context.session.query("print.rows", { sheet: 0 })

				return rows?.indices.length ?? -1
			}
		})
		const session = opened(await host.open(buildFixture(BASIC), { kind: "xlsx" }, () => undefined))
		const first = await session.query("print.rows", { sheet: 0 })
		const second = await session.query("print.rows", { sheet: 0 })

		expect(second?.indices.length).toBe(first?.indices.length)
		expect(await sheetJobs(session)?.job("rows", null)).toBe(first?.indices.length)
		host.dispose()
		close()
	})
	it("answers a PDF export whose session is closed or whose host is disposed as closed, not as a failure", async () => {
		const pending: SheetWorkerJob = async (_args, context) => {
			while (!context.signal.aborted) await new Promise(resolve => setTimeout(resolve, 1))

			throw new Error("aborted in the worker")
		}

		for (const end of ["close", "dispose"] as const) {
			const { host, close } = workerPair({ [SHEET_PDF_JOB]: pending })
			const session = opened(await host.open(buildFixture(BASIC), { kind: "xlsx" }, () => undefined))
			const pdf = sheetPdf(session, {} as never)

			await new Promise(resolve => setTimeout(resolve, 5))

			if (end === "close") session.close()
			else host.dispose()

			expect(await pdf).toEqual({ ok: false, reason: "closed" })
			host.dispose()
			close()
		}
	})
	it("makes the same PDF twice when asked for a deterministic one, and a new file identifier otherwise", async () => {
		const { host, close } = workerPair(sheetPdfJob(nodeFontAssets))
		const result = await host.open(buildFixture(BASIC), { kind: "xlsx" }, () => undefined)

		if (result.kind !== "opened") throw new Error(`expected an open, got ${result.kind}`)

		const used = { startRow: 0, startCol: 0, endRow: 2, endCol: 1 }
		const request: SheetPdfRequest = {
			sheets: [{ sheet: result.info.activeSheet, name: "Sheet1", setup: DEFAULT_PAGE_SETUP, used, ranges: [used] }],
			values: { file: "book.xlsx", path: "", date: "", time: "" },
			title: "book",
			lang: "en",
			kind: "xlsx",
			summaries: result.info.sheets,
			date1904: false,
			showFormulas: false
		}
		const bytes = async (deterministic: boolean): Promise<Uint8Array> => {
			const pdf = await sheetPdf(result.session, request, { deterministic })

			if (!pdf.ok) throw new Error(`expected a PDF, got ${pdf.reason}`)

			return new Uint8Array(await pdf.blob.arrayBuffer())
		}

		const first = await bytes(true)

		expect(await bytes(true)).toEqual(first)
		expect(await bytes(false)).not.toEqual(first)
		host.dispose()
		close()
	})
})
