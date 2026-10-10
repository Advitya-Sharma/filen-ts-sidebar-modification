// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { createElement } from "react"
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react"
import type { DriveItem } from "@/features/drive/lib/item"
import type { EditResult } from "@/features/spreadsheet/lib/edits"
import { cellKey, type SpreadsheetDoc } from "@/features/spreadsheet/lib/model"
import { mockSheetView } from "@/tests/mockSheetView"

const BYTES = new Uint8Array([1, 2, 3])
const NOTHING = { dirty: false, canUndo: false, canRedo: false }

const { remote, toastError, layerUpdate } = vi.hoisted(() => ({
	remote: {
		apply: vi.fn(),
		undo: vi.fn(),
		redo: vi.fn(),
		writability: vi.fn(),
		viewOnly: vi.fn(),
		serialize: vi.fn(),
		markSaved: vi.fn()
	},
	toastError: vi.fn(),
	layerUpdate: vi.fn()
}))

vi.mock("sonner", () => ({ toast: { success: vi.fn(), error: toastError } }))
vi.mock("@/features/preview/hooks/usePreviewBytes", () => ({
	usePreviewBytes: () => ({ status: "success", bytes: BYTES, refetch: () => undefined })
}))
vi.mock("@/features/spreadsheet/hooks/useSizeLayer", () => ({
	useSizeLayer: () => ({ layer: new Map(), update: layerUpdate, follow: vi.fn() })
}))
vi.mock(import("@/features/spreadsheet/lib/spreadsheetClient"), async importOriginal => ({
	...(await importOriginal()),
	openSpreadsheet: () => Promise.resolve({ id: 1, doc: DOC }),
	closeSpreadsheet: vi.fn(),
	withOpenSpreadsheet: (id: number, run: (remote: never, workerId: number) => Promise<unknown>) => run(remote as never, id) as never
}))

import "@/lib/i18n"
import { SpreadsheetViewer } from "@/features/spreadsheet/components/spreadsheetViewer"

// Two sheets, A1 holding a value, the second sheet's name "Two".
const DOC: SpreadsheetDoc = {
	kind: "xlsx",
	sheets: [
		mockSheetView({ name: "One", rowCount: 3, colCount: 3, cells: new Map([[cellKey(0, 0), { text: "1", numeric: true }]]) }),
		mockSheetView({ name: "Two", rowCount: 3, colCount: 3 })
	],
	activeSheet: 0,
	styles: [],
	writable: false
}

function file(): DriveItem {
	return { type: "file", data: { uuid: "u1", decryptedMeta: { name: "book.xlsx" } } } as unknown as DriveItem
}

function viewer(editable = true) {
	return createElement(SpreadsheetViewer, { item: file(), documentKey: "book", alt: "book.xlsx", editable })
}

const refused = (reason: "readOnly" | "formula"): EditResult => ({ type: "refused", reason, state: NOTHING })

async function opened(editable = true) {
	const rendered = render(viewer(editable))

	await screen.findByRole("grid")

	return rendered
}

function bar(): HTMLInputElement {
	const input = screen.getAllByLabelText("Cell contents").find(element => element.tagName === "INPUT")

	if (!(input instanceof HTMLInputElement)) throw new Error("no formula bar")

	return input
}

// Each test sets what the worker answers: none carries over from another.
beforeEach(() => {
	for (const mock of [...Object.values(remote), toastError, layerUpdate]) mock.mockReset()

	remote.writability.mockReturnValue(new Promise(() => undefined))
	remote.apply.mockResolvedValue({ type: "none", state: NOTHING })
})

afterEach(() => {
	cleanup()
})

describe("SpreadsheetViewer editing an unproven workbook", () => {
	it("proves the workbook only once the user sets out to edit it, sending the file as opened", async () => {
		await opened()

		expect(remote.writability).not.toHaveBeenCalled()
		expect(screen.queryByText("Checking this file can be saved…")).toBeNull()

		fireEvent.focus(bar())

		await waitFor(() => {
			expect(remote.writability).toHaveBeenCalledWith(1, BYTES)
		})
		expect(screen.getByText("Checking this file can be saved…")).toBeTruthy()
	})

	it("keeps a malformed formula open, holding the selection and the sheet, until editing stops", async () => {
		remote.writability.mockResolvedValue(true)

		const { rerender } = await opened()

		fireEvent.focus(bar())
		fireEvent.change(bar(), { target: { value: "=SUM(A1;B1)" } })
		fireEvent.keyDown(bar(), { key: "Enter" })

		expect(screen.getByRole("alert").textContent).toContain("This formula can't be read")
		expect(bar().getAttribute("aria-invalid")).toBe("true")
		expect(remote.apply).not.toHaveBeenCalled()

		fireEvent.keyDown(screen.getByRole("grid"), { key: "ArrowDown" })
		fireEvent.click(screen.getByRole("tab", { name: "Two" }))

		expect(screen.getByLabelText("Selected cells").textContent).toBe("A1")
		expect(screen.getByRole("tab", { name: "One" }).getAttribute("aria-selected")).toBe("true")
		expect(screen.getByRole("alert")).toBeTruthy()

		// Mended, it commits.
		fireEvent.change(bar(), { target: { value: "=SUM(A1,B1)" } })

		expect(screen.queryByRole("alert")).toBeNull()

		fireEvent.change(bar(), { target: { value: "=SUM(A1" } })
		fireEvent.keyDown(bar(), { key: "Enter" })

		expect(screen.getByRole("alert")).toBeTruthy()

		// Editing turns off (the viewer goes read-only): the entry that can no longer be mended is cancelled.
		rerender(viewer(false))

		expect(screen.queryByRole("alert")).toBeNull()
		expect(screen.getAllByLabelText("Cell contents").some(element => element.tagName === "INPUT")).toBe(false)
		expect(remote.apply).not.toHaveBeenCalled()
	})

	it("shows one toast for edits refused alike, as when the proof fails under them", async () => {
		remote.writability.mockResolvedValue(false)
		remote.apply.mockResolvedValue(refused("readOnly"))

		await opened()

		const grid = screen.getByRole("grid")

		fireEvent.keyDown(grid, { key: "Delete" })
		fireEvent.keyDown(grid, { key: "Delete" })

		await waitFor(() => {
			expect(toastError).toHaveBeenCalledTimes(2)
		})

		for (const call of toastError.mock.calls) expect(call[1]).toEqual({ id: "readOnly" })
	})

	it("keeps a size beside the file when the workbook turns out read-only", async () => {
		remote.writability.mockResolvedValue(false)
		remote.apply.mockResolvedValue(refused("readOnly"))

		await opened()

		fireEvent.contextMenu(screen.getByRole("grid"))
		fireEvent.click(await screen.findByRole("menuitem", { name: "Reset column width" }))

		await waitFor(() => {
			expect(layerUpdate).toHaveBeenCalledWith(0, "cols", [[0, null]])
		})
		expect(remote.apply).toHaveBeenCalledWith(1, { type: "resize", sheet: 0, axis: "cols", sizes: [[0, null]] })
	})
})
