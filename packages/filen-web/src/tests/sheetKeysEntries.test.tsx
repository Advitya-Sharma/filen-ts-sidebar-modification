// @vitest-environment jsdom

import { beforeEach, describe, expect, it, vi } from "vitest"
import { act, fireEvent, render, screen, waitFor } from "@testing-library/react"
import { inThreadSheetHost } from "@filen/office/dom/sheet"
import type { DriveItem } from "@/features/drive/lib/item"
import { registerAllActions } from "@/lib/keymap/actions"
import { preview } from "@/locales/en/preview"

// Entries through the formula bar over a real text session on this thread, read back as the active cell comes back.

const { bytes } = vi.hoisted(() => {
	const held: { current: Uint8Array } = { current: new Uint8Array(0) }

	return { bytes: held }
})

vi.mock("@/providers/themeProvider", async importOriginal => ({
	...(await importOriginal<typeof import("@/providers/themeProvider")>()),
	useTheme: () => ({ theme: "light" })
}))

vi.mock("@/features/preview/hooks/usePreviewBytes", () => ({
	usePreviewBytes: () => ({ status: "success", bytes: bytes.current, refetch: () => undefined })
}))

vi.mock("@/features/spreadsheet/next/lib/sheetHosts", async importOriginal => {
	const host = inThreadSheetHost()

	return {
		...(await importOriginal<typeof import("@/features/spreadsheet/next/lib/sheetHosts")>()),
		openSheet: (source: Uint8Array, options: Parameters<typeof host.open>[1], onEvent: Parameters<typeof host.open>[2]) =>
			host.open(source, options, onEvent)
	}
})

const { SheetNextViewer } = await import("@/features/spreadsheet/next/sheetNextViewer")

registerAllActions()

beforeEach(() => {
	vi.stubGlobal(
		"ResizeObserver",
		class {
			observe = vi.fn()
			unobserve = vi.fn()
			disconnect = vi.fn()
		}
	)
	vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockReturnValue(null)
})

function file(name: string): DriveItem {
	return { type: "file", data: { uuid: `u-${name}`, decryptedMeta: { name }, size: 1n } } as unknown as DriveItem
}

function goTo(ref: string): void {
	const box = screen.getByRole("textbox", { name: preview.previewSheetNameBox })

	fireEvent.focus(box)
	fireEvent.change(box, { target: { value: ref } })
	fireEvent.keyDown(box, { key: "Enter" })
}

function contents(): string {
	return screen.getByRole<HTMLInputElement>("textbox", { name: preview.previewSpreadsheetCellContents }).value
}

describe("entries in a CSV", () => {
	it("show again in the formula bar when the active cell comes back to a row the file did not have", async () => {
		bytes.current = Uint8Array.from(new TextEncoder().encode("a,b\r\n1,2\r\n"))

		render(
			<SheetNextViewer
				item={file("rows.csv")}
				documentKey="rows"
				alt="rows.csv"
				editable
				saveRef={{ current: null }}
			/>
		)

		await screen.findByRole("textbox", { name: preview.previewSpreadsheetCellContents })
		goTo("A2")
		await waitFor(() => {
			expect(contents()).toBe("1")
		})

		goTo("A3")

		const bar = screen.getByRole("textbox", { name: preview.previewSpreadsheetCellContents })

		await act(async () => {
			bar.focus()
			await Promise.resolve()
		})
		fireEvent.change(bar, { target: { value: "x" } })
		fireEvent.keyDown(bar, { key: "Enter" })

		goTo("A5")
		await waitFor(() => {
			expect(contents()).toBe("")
		})
		goTo("A3")
		await waitFor(() => {
			expect(contents()).toBe("x")
		})
	})
})
