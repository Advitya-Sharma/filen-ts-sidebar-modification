// @vitest-environment jsdom

import { describe, expect, it, vi } from "vitest"
import { useState } from "react"
import { act, renderHook, waitFor } from "@testing-library/react"
import type { DriveItem } from "@/features/drive/lib/item"
import { cellKey, type GridDoc, type SpreadsheetDoc } from "@filen/office/core/sheet"

const { open, close, undo, redo, apply, writability, viewOnly } = vi.hoisted(() => ({
	viewOnly: vi.fn(() => Promise.resolve()),
	writability: vi.fn(),
	open: vi.fn(),
	close: vi.fn(),
	undo: vi.fn(),
	redo: vi.fn(),
	apply: vi.fn()
}))

vi.mock(import("@/features/spreadsheet/lib/spreadsheetClient"), async importOriginal => ({
	...(await importOriginal()),
	openSpreadsheet: open,
	closeSpreadsheet: close,
	withOpenSpreadsheet: (id: number, run: (remote: never, workerId: number) => Promise<unknown>) =>
		run({ undo, redo, apply, writability, viewOnly } as never, id) as never,
	sniffSpreadsheetKind: () => "csv"
}))

const BYTES = new Uint8Array([1, 2, 3])

vi.mock("@/features/preview/hooks/usePreviewBytes", () => ({
	usePreviewBytes: () => ({ status: "success", bytes: BYTES, refetch: () => undefined })
}))

const { useSpreadsheetDoc } = await import("@/features/spreadsheet/hooks/useSpreadsheetDoc")
const { useSpreadsheetEdits, useSpreadsheetWritability } = await import("@/features/spreadsheet/hooks/useSpreadsheetEdits")
const { gridDoc } = await import("@filen/office/core/sheet")
const { csvDoc } = await import("@filen/office/core/sheet")

const DOC: SpreadsheetDoc = { kind: "csv", sheets: [], activeSheet: 0, styles: [], writable: true }

// The edits over a view held as useSpreadsheetDoc holds it.
function useEditsOver(initial: GridDoc, onShift?: Parameters<typeof useSpreadsheetEdits>[2]) {
	const [doc, setDoc] = useState(initial)

	return { doc, ...useSpreadsheetEdits(1, setDoc, onShift) }
}

function file(uuid: string, name: string): DriveItem {
	return { type: "file", data: { uuid, decryptedMeta: { name } } } as unknown as DriveItem
}

describe("useSpreadsheetDoc", () => {
	it("keeps the document across an own save, and stops saving once renamed to another format", async () => {
		open.mockResolvedValue({ id: 7, doc: DOC })

		const { result, rerender } = renderHook(({ item }) => useSpreadsheetDoc(item, "key"), {
			initialProps: { item: file("u1", "t.csv") }
		})

		await waitFor(() => {
			expect(result.current.status).toBe("ready")
		})
		expect(result.current).toMatchObject({ id: 7, unnamed: false, renamed: false })

		rerender({ item: file("u2", "t.csv") })
		expect(result.current).toMatchObject({ id: 7, renamed: false })

		rerender({ item: file("u2", "t.xlsx") })
		expect(result.current).toMatchObject({ id: 7, renamed: true })
		expect(open).toHaveBeenCalledTimes(1)
	})
})

describe("useSpreadsheetDoc held view", () => {
	it("holds one view, which edits replace rather than keep beside the view as opened", async () => {
		open.mockResolvedValue({ id: 9, doc: csvDoc([["a"]], true) })

		const { result } = renderHook(() => useSpreadsheetDoc(file("v1", "v.csv"), "view"))

		await waitFor(() => {
			expect(result.current.status).toBe("ready")
		})

		const opened = result.current.status === "ready" ? result.current.doc : null

		act(() => {
			if (result.current.status === "ready") {
				result.current.updateDoc(doc => ({ ...doc, activeSheet: 1 }))
			}
		})

		expect(result.current).toMatchObject({ status: "ready", id: 9, doc: { activeSheet: 1 } })
		expect(result.current.status === "ready" && result.current.doc).not.toBe(opened)
	})
})

describe("useSpreadsheetDoc renames", () => {
	it("treats .xlsm and .xlsx as different formats, but a changed letter case as the same", async () => {
		open.mockResolvedValue({ id: 8, doc: DOC })

		const { result, rerender } = renderHook(({ item }) => useSpreadsheetDoc(item, "macros"), {
			initialProps: { item: file("m1", "book.xlsm") }
		})

		await waitFor(() => {
			expect(result.current.status).toBe("ready")
		})

		rerender({ item: file("m1", "Book.XLSM") })
		expect(result.current).toMatchObject({ renamed: false })

		rerender({ item: file("m1", "book.xlsx") })
		expect(result.current).toMatchObject({ renamed: true })
	})
})

describe("useSpreadsheetEdits", () => {
	it("sends no undo or redo when there is nothing to undo or redo, so nothing shows as pending", async () => {
		const { result } = renderHook(() => useEditsOver(gridDoc(DOC)))
		let outcome: unknown

		await act(async () => {
			outcome = await result.current.undo()
			await result.current.redo()
		})

		expect(outcome).toMatchObject({ type: "none" })
		expect(undo).not.toHaveBeenCalled()
		expect(redo).not.toHaveBeenCalled()
		expect(result.current.pending).toBe(false)
	})

	it("counts an edit as pending until the worker answers", async () => {
		let answer: (value: unknown) => void = () => undefined

		apply.mockReturnValue(
			new Promise(resolve => {
				answer = resolve
			})
		)

		const { result } = renderHook(() => useEditsOver(gridDoc(DOC)))
		let done: Promise<unknown> = Promise.resolve()

		act(() => {
			done = result.current.apply({ type: "addSheet", name: "S" })
		})
		expect(result.current.pending).toBe(true)

		await act(async () => {
			answer({ type: "none", state: { dirty: true, canUndo: true, canRedo: false } })
			await done
		})
		expect(result.current.pending).toBe(false)
		expect(result.current.state.dirty).toBe(true)
	})

	it("tells onShift where a CSV's rows or columns moved, and moves the held cells with them", async () => {
		const shift = { type: "insert", axis: "rows", at: 0, count: 1, revert: false } as const
		const onShift = vi.fn()

		apply.mockResolvedValue({
			type: "shifted",
			sheet: 0,
			shift,
			cells: [],
			rowCount: 2,
			colCount: 1,
			state: { dirty: true, canUndo: true, canRedo: false }
		})

		const { result } = renderHook(() => useEditsOver(gridDoc(csvDoc([["a"]], true)), onShift))

		await act(async () => {
			await result.current.apply({ type: "insert", sheet: 0, axis: "rows", at: 0, count: 1 })
		})

		expect(onShift).toHaveBeenCalledExactlyOnceWith(0, shift)
		expect(result.current.doc.sheets[0]?.cells.get(cellKey(1, 0))).toEqual({ text: "a" })
		expect(result.current.doc.sheets[0]?.rowCount).toBe(2)
	})
})

describe("useSpreadsheetWritability", () => {
	const XLSX = { kind: "xlsx", writable: false } as const

	it("keeps a workbook read-only until the worker proves it saves intact", async () => {
		let answer: (writable: boolean) => void = () => undefined

		writability.mockReturnValue(
			new Promise<boolean>(resolve => {
				answer = resolve
			})
		)

		const { result } = renderHook(() => useSpreadsheetWritability(3, XLSX, BYTES, true, false))

		expect(result.current).toBe("checking")
		expect(writability).toHaveBeenCalledWith(3, BYTES)

		await act(async () => {
			answer(true)
			await Promise.resolve()
		})
		expect(result.current).toBe("writable")
	})

	it("stays read-only when the proof fails or the check throws", async () => {
		writability.mockResolvedValueOnce(false).mockRejectedValueOnce(new Error("closed"))

		const refused = renderHook(() => useSpreadsheetWritability(4, XLSX, BYTES, true, false))
		const failed = renderHook(() => useSpreadsheetWritability(5, XLSX, BYTES, true, false))

		await waitFor(() => {
			expect(refused.result.current).toBe("readOnly")
			expect(failed.result.current).toBe("readOnly")
		})
	})

	it("asks nothing for text, or when editing could not follow, and ignores an answer after unmount", async () => {
		writability.mockClear()

		viewOnly.mockClear()

		expect(renderHook(() => useSpreadsheetWritability(6, { kind: "csv", writable: true }, BYTES, true, false)).result.current).toBe(
			"writable"
		)
		// Not asked to edit yet: nothing is proven, and editing is still on offer.
		expect(renderHook(() => useSpreadsheetWritability(7, XLSX, BYTES, false, false)).result.current).toBe("unproven")
		expect(writability).not.toHaveBeenCalled()
		// Not editable for now is not "never": the proof stays possible.
		expect(viewOnly).not.toHaveBeenCalled()

		// A public link: nothing will edit it, so the worker skips the proof.
		renderHook(() => useSpreadsheetWritability(9, XLSX, BYTES, false, true))
		expect(viewOnly).toHaveBeenCalledWith(9)
		expect(writability).not.toHaveBeenCalled()

		let answer: (writable: boolean) => void = () => undefined

		writability.mockReturnValue(
			new Promise<boolean>(resolve => {
				answer = resolve
			})
		)

		const gone = renderHook(() => useSpreadsheetWritability(8, XLSX, BYTES, true, false))

		gone.unmount()
		await act(async () => {
			answer(true)
			await Promise.resolve()
		})
		expect(gone.result.current).toBe("checking")
	})

	it("asks nothing about a workbook known at open to stay view-only, and says why", () => {
		writability.mockClear()
		viewOnly.mockClear()

		expect(
			renderHook(() => useSpreadsheetWritability(11, { kind: "xlsx", writable: false, unprovable: "unsafe" }, BYTES, true, false))
				.result.current
		).toBe("readOnly")
		expect(
			renderHook(() => useSpreadsheetWritability(12, { kind: "xlsx", writable: false, unprovable: "tooLarge" }, BYTES, true, false))
				.result.current
		).toBe("tooLarge")
		expect(writability).not.toHaveBeenCalled()
		expect(viewOnly).not.toHaveBeenCalled()
	})

	it("asks once editing becomes possible later, as after a rename back", async () => {
		writability.mockReset()
		viewOnly.mockClear()
		writability.mockResolvedValue(true)

		const { result, rerender } = renderHook(({ wanted }) => useSpreadsheetWritability(10, XLSX, BYTES, wanted, false), {
			initialProps: { wanted: false }
		})

		expect(writability).not.toHaveBeenCalled()
		rerender({ wanted: true })
		await waitFor(() => {
			expect(result.current).toBe("writable")
		})
		expect(writability).toHaveBeenCalledWith(10, BYTES)
		expect(viewOnly).not.toHaveBeenCalled()
	})
})
