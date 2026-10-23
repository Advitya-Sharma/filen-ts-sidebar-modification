import { beforeEach, describe, expect, it, vi } from "vitest"

const mocks = vi.hoisted(() => ({ openSheet: vi.fn(), textToXlsx: vi.fn() }))

vi.mock("@/features/spreadsheet/lib/sheetHosts", () => ({
	openSheet: mocks.openSheet,
	textToXlsx: mocks.textToXlsx,
	isTextSheetKind: (kind: string) => kind === "csv" || kind === "tsv" || kind === "tab"
}))

import type { SheetHandle } from "@/features/spreadsheet/hooks/useSheetSession"
import { commitForExport, liveConversion } from "@/features/spreadsheet/save/sheetConvert"

// Download as ▸ .xlsx / .ods from the spreadsheet editor: the open session's cells as they are now, the cell being
// typed committed first, as Download as CSV and PDF read them. A text file's cells are made a workbook as Save as
// .xlsx makes it, and converted on from there; a workbook converts in its own session. A protected workbook has none
// (its copy is never written unprotected from here).

const CELLS = new TextEncoder().encode("a\tb\n")
const XLSX = new TextEncoder().encode("xlsx")

function handle(answer: unknown = { bytes: CELLS }, encrypted?: string) {
	const query = vi.fn((..._args: unknown[]) => Promise.resolve(answer))

	return { query, handle: { info: { encrypted }, session: { query } } as unknown as SheetHandle }
}

beforeEach(() => {
	mocks.openSheet.mockReset()
	mocks.textToXlsx.mockReset()
})

describe("liveConversion", () => {
	it("makes an edited CSV's live cells a workbook, the open entry committed first", async () => {
		const { handle: text, query } = handle()
		const commit = vi.fn(() => Promise.resolve(true))

		mocks.textToXlsx.mockResolvedValue({ ok: true, bytes: XLSX })

		const live = liveConversion(text, "csv", "data.csv", commit)
		const signal = new AbortController().signal

		expect(await live?.("xlsx", signal)).toEqual({ ok: true, bytes: XLSX, lossy: [] })
		expect(commit.mock.invocationCallOrder[0]).toBeLessThan(query.mock.invocationCallOrder[0] ?? 0)
		expect(query).toHaveBeenCalledWith("saveAs.cells", {}, { signal })
		expect(mocks.textToXlsx).toHaveBeenCalledWith(CELLS, "data")
	})

	it("converts an edited TSV's live cells on to .ods in a workbook session of its own, closed after", async () => {
		const { handle: text } = handle()
		const ods = { ok: true, bytes: new TextEncoder().encode("ods"), lossy: [] }
		const workbook = { query: vi.fn(() => Promise.resolve(ods)), close: vi.fn() }

		mocks.textToXlsx.mockResolvedValue({ ok: true, bytes: XLSX })
		mocks.openSheet.mockResolvedValue({ kind: "opened", session: workbook, info: {} })

		const signal = new AbortController().signal

		expect(await liveConversion(text, "tsv", "data.tsv", () => Promise.resolve(true))?.("ods", signal)).toEqual(ods)
		expect(mocks.openSheet.mock.calls[0]?.[0]).toBe(XLSX)
		expect(mocks.openSheet.mock.calls[0]?.[1]).toMatchObject({ kind: "xlsx" })
		expect(workbook.query).toHaveBeenCalledWith("saveAs.format", { format: "ods" }, { signal })
		expect(workbook.close).toHaveBeenCalledOnce()
	})

	it("converts a workbook in its own session, the open entry committed first", async () => {
		const ods = { ok: true, bytes: new TextEncoder().encode("ods"), lossy: [] }
		const { handle: book, query } = handle(ods)
		const commit = vi.fn(() => Promise.resolve(true))
		const signal = new AbortController().signal

		expect(await liveConversion(book, "xlsx", "book.xlsx", commit)?.("ods", signal)).toEqual(ods)
		expect(commit.mock.invocationCallOrder[0]).toBeLessThan(query.mock.invocationCallOrder[0] ?? 0)
		expect(query).toHaveBeenCalledWith("saveAs.format", { format: "ods" }, { signal })
	})

	it("refuses when the open entry cannot be committed, converting nothing", async () => {
		const { handle: book, query } = handle()

		expect(
			await liveConversion(book, "xlsx", "book.xlsx", () => Promise.resolve(false))?.("ods", new AbortController().signal)
		).toMatchObject({
			ok: false,
			reason: "uncommitted"
		})
		expect(query).not.toHaveBeenCalled()
	})

	it("answers aborted once cancelled while the cells were made a workbook, handing nothing out", async () => {
		const { handle: text } = handle()
		const controller = new AbortController()

		mocks.textToXlsx.mockImplementation(() => {
			controller.abort()

			return Promise.resolve({ ok: true, bytes: XLSX })
		})

		expect(await liveConversion(text, "csv", "data.csv", () => Promise.resolve(true))?.("xlsx", controller.signal)).toMatchObject({
			ok: false,
			reason: "aborted"
		})
		expect(mocks.openSheet).not.toHaveBeenCalled()
	})

	it("answers aborted once cancelled while the workbook converted", async () => {
		const controller = new AbortController()
		const query = vi.fn(() => {
			controller.abort()

			return Promise.resolve({ ok: true, bytes: XLSX, lossy: [] })
		})
		const book = { info: { encrypted: undefined }, session: { query } } as unknown as SheetHandle

		expect(await liveConversion(book, "xlsx", "book.xlsx", () => Promise.resolve(true))?.("ods", controller.signal)).toMatchObject({
			ok: false,
			reason: "aborted"
		})
	})

	it("has none for a password-protected workbook", () => {
		expect(liveConversion(handle(undefined, "verified").handle, "xlsx", "book.xlsx", () => Promise.resolve(true))).toBeUndefined()
	})
})

describe("commitForExport", () => {
	it("commits the open entry and goes on", async () => {
		const say = vi.fn()

		expect(await commitForExport({ commitEditing: () => Promise.resolve(true), getSnapshot: () => ({ editing: null }) }, say)).toBe(
			true
		)
		expect(say).not.toHaveBeenCalled()
	})

	it("says why an entry the sheet reopened as a malformed formula stops the export", async () => {
		const say = vi.fn()
		const model = { commitEditing: () => Promise.resolve(false), getSnapshot: () => ({ editing: { invalid: true } }) }

		expect(await commitForExport(model, say)).toBe(false)
		expect(say).toHaveBeenCalledOnce()
	})

	it("adds nothing where the sheet already told why (its refusal toast, the validation alert)", async () => {
		const say = vi.fn()
		const model = { commitEditing: () => Promise.resolve(false), getSnapshot: () => ({ editing: null }) }

		expect(await commitForExport(model, say)).toBe(false)
		expect(say).not.toHaveBeenCalled()
	})
})
