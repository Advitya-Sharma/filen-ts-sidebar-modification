// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { createElement, createRef } from "react"
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react"
import type { DriveItem } from "@/features/drive/lib/item"
import type { EditResult } from "@/features/spreadsheet/lib/edits"
import { cellKey, type SpreadsheetDoc } from "@/features/spreadsheet/lib/model"
import { mockSheetView } from "@/tests/mockSheetView"

const BYTES = new Uint8Array([1, 2, 3])
const NOTHING = { dirty: false, canUndo: false, canRedo: false }

const { remote, toastError, layerUpdate, openSpreadsheet } = vi.hoisted(() => ({
	openSpreadsheet: vi.fn(),
	remote: {
		apply: vi.fn(),
		undo: vi.fn(),
		redo: vi.fn(),
		writability: vi.fn(),
		viewOnly: vi.fn(),
		serialize: vi.fn(),
		markSaved: vi.fn(),
		sheet: vi.fn()
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
	openSpreadsheet,
	closeSpreadsheet: vi.fn(),
	withOpenSpreadsheet: (id: number, run: (remote: never, workerId: number) => Promise<unknown>) => run(remote as never, id) as never
}))

import "@/lib/i18n"
import { SpreadsheetViewer, type SpreadsheetSaveSource } from "@/features/spreadsheet/components/spreadsheetViewer"

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
	openSpreadsheet.mockReset()
	openSpreadsheet.mockImplementation(() => Promise.resolve({ id: 1, doc: DOC }))
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

describe("SpreadsheetViewer with a deferred sheet", () => {
	it("asks the worker for a sheet the first time it is shown, and shows it once built", async () => {
		const deferred = { ...DOC, sheets: [DOC.sheets[0], mockSheetView({ name: "Two", rowCount: 3, colCount: 3, deferred: true })] }

		openSpreadsheet.mockImplementation(() => Promise.resolve({ id: 1, doc: deferred }))
		remote.sheet.mockResolvedValue({
			type: "sheets",
			sheets: [null, mockSheetView({ name: "Two", rowCount: 3, colCount: 3, cells: new Map([[cellKey(1, 1), { text: "built" }]]) })],
			styles: [],
			state: NOTHING
		} satisfies EditResult)

		// jsdom has no scrolling; a sheet switch scrolls the grid to its top.
		Object.defineProperty(HTMLElement.prototype, "scrollTo", { configurable: true, value: () => undefined })
		await opened()

		expect(remote.sheet).not.toHaveBeenCalled()

		fireEvent.click(screen.getByRole("tab", { name: "Two" }))

		expect(await screen.findByText("built")).toBeTruthy()
		expect(remote.sheet).toHaveBeenCalledExactlyOnceWith(1, 1)

		fireEvent.click(screen.getByRole("tab", { name: "One" }))
		fireEvent.click(screen.getByRole("tab", { name: "Two" }))

		expect(screen.getByText("built")).toBeTruthy()
		expect(remote.sheet).toHaveBeenCalledOnce()
		Reflect.deleteProperty(HTMLElement.prototype, "scrollTo")
	})
})

describe("SpreadsheetViewer with a password-protected workbook", () => {
	const CFB = Uint8Array.of(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0)
	const ZIP = Uint8Array.of(0x50, 0x4b, 0x03, 0x04, 0)

	function lock() {
		return { unlock: vi.fn(), cancel: vi.fn(), discard: vi.fn() }
	}

	async function submit(password: string): Promise<void> {
		fireEvent.change(await screen.findByLabelText("Password"), { target: { value: password } })
		fireEvent.click(screen.getByRole("button", { name: "Open" }))
	}

	it("shows the workbook unreadable and gives its lock up when a try fails", async () => {
		const locked = lock()

		locked.unlock.mockRejectedValueOnce(new Error("spreadsheet: too many cells"))
		openSpreadsheet.mockResolvedValue({ locked })

		const { unmount } = render(viewer())

		await submit("right")
		await waitFor(() => {
			expect(screen.queryByRole("dialog")).toBeNull()
		})
		expect(locked.discard).toHaveBeenCalledOnce()
		unmount()
		expect(locked.discard).toHaveBeenCalledOnce()
	})

	it("asks for the password in the workbook's place, says when it is wrong, and opens with the right one", async () => {
		const locked = lock()

		locked.unlock.mockResolvedValueOnce({ wrongPassword: true }).mockResolvedValueOnce({ id: 1, doc: DOC, encrypted: "verified" })
		openSpreadsheet.mockResolvedValue({ locked })
		render(viewer())

		expect(await screen.findByText("This document is password-protected.")).toBeTruthy()
		expect(screen.getByRole("dialog", { name: "Password-protected document" })).toBeTruthy()

		await submit("wrong")
		expect(await screen.findByText("That password isn't right. Try again.")).toBeTruthy()

		await submit("right")
		await screen.findByRole("grid")

		expect(locked.unlock.mock.calls).toEqual([["wrong"], ["right"]])
		expect(screen.getByRole("button", { name: "Password-protected" })).toBeTruthy()
		expect(locked.discard).not.toHaveBeenCalled()
	})

	it("gives up on the lock when the viewer goes, and stops a try the user cancels", async () => {
		const locked = lock()
		let answer: (value: unknown) => void = () => undefined

		locked.unlock.mockReturnValueOnce(
			new Promise(resolve => {
				answer = resolve
			})
		)
		openSpreadsheet.mockResolvedValue({ locked })

		const { unmount } = render(viewer())

		await submit("slow")
		fireEvent.click(await screen.findByRole("button", { name: "Cancel" }))
		expect(locked.cancel).toHaveBeenCalledOnce()
		answer({ aborted: true })

		unmount()
		expect(locked.discard).toHaveBeenCalledOnce()
	})

	it.each([
		["rightsManaged", "This document is protected with rights management (IRM) and can't be opened in Filen."],
		["unsupported", "This document uses an encryption method Filen can't open."],
		["damaged", "This document is damaged or was modified after it was encrypted."]
	] as const)("refuses a %s workbook without asking for a password", async (reason, message) => {
		openSpreadsheet.mockResolvedValue({ refused: reason })
		render(viewer())

		expect(await screen.findByText(message)).toBeTruthy()
		expect(screen.queryByRole("dialog")).toBeNull()
	})

	it("warns, beside the badge's wording, when the workbook's integrity can't be verified", async () => {
		openSpreadsheet.mockResolvedValue({ id: 1, doc: DOC, encrypted: "unchecked" })
		render(viewer(false))
		await screen.findByRole("grid")

		const badge = screen.getByRole("button", { name: "Password-protected" })

		expect(badge.getAttribute("aria-describedby")).toBeTruthy()
		expect(screen.getByText("Its integrity can't be verified: it may have been changed after it was encrypted.")).toBeTruthy()
		fireEvent.click(badge)
		expect(await screen.findByText("Opened with its password, in this browser only.")).toBeTruthy()
	})

	it("proves it from the worker's plaintext and saves only sealed bytes", async () => {
		const saveRef = createRef<SpreadsheetSaveSource | null>() as { current: SpreadsheetSaveSource | null }

		openSpreadsheet.mockResolvedValue({ id: 1, doc: DOC, encrypted: "verified" })
		remote.writability.mockResolvedValue(true)
		render(createElement(SpreadsheetViewer, { item: file(), documentKey: "book", alt: "book.xlsx", editable: true, saveRef }))
		await screen.findByRole("grid")

		fireEvent.focus(bar())
		await waitFor(() => {
			expect(remote.writability).toHaveBeenCalledWith(1, null)
		})
		await waitFor(() => {
			expect(saveRef.current).not.toBeNull()
		})

		remote.serialize.mockResolvedValueOnce({ bytes: ZIP, version: 1 }).mockResolvedValueOnce({ bytes: CFB, version: 1 })

		await expect(saveRef.current?.()).rejects.toThrow("not sealed")
		await expect(saveRef.current?.().then(snapshot => snapshot.bytes)).resolves.toEqual(CFB)
	})
})
