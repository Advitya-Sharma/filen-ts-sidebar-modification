// @vitest-environment jsdom

import { beforeEach, describe, expect, it, vi } from "vitest"
import { fireEvent, render, screen, waitFor, within } from "@testing-library/react"
import { QueryClient, QueryClientProvider } from "@tanstack/react-query"
import { inThreadSheetHost } from "@filen/office/dom/sheet"
import type { DriveItem } from "@/features/drive/lib/item"
import { registerAllActions } from "@/lib/keymap/actions"
import { common } from "@/locales/en/common"
import { preview } from "@/locales/en/preview"
import { workbook } from "@/tests/sheetMenusFixtures"

// The link dialog (Ctrl/Cmd+K, Edit link) gives the grid its focus back however it closes, as the other dialogs do.

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

vi.mock("@/features/spreadsheet/lib/sheetHosts", async importOriginal => {
	const host = inThreadSheetHost()

	return {
		...(await importOriginal<typeof import("@/features/spreadsheet/lib/sheetHosts")>()),
		openSheet: (source: Uint8Array, options: Parameters<typeof host.open>[1], onEvent: Parameters<typeof host.open>[2]) =>
			host.open(source, options, onEvent)
	}
})

const { SheetViewer } = await import("@/features/spreadsheet/sheetViewer")

// The find field is the document editor's, which binds Mod+F through the app's keymap.
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

function nameBox(): string {
	return screen.getByRole<HTMLInputElement>("textbox", { name: preview.previewSheetNameBox }).value
}

async function goTo(reference: string): Promise<void> {
	const box = screen.getByRole<HTMLInputElement>("textbox", { name: preview.previewSheetNameBox })

	fireEvent.focus(box)
	fireEvent.change(box, { target: { value: reference } })
	fireEvent.keyDown(box, { key: "Enter" })
	await waitFor(() => {
		expect(nameBox()).toBe(reference)
	})
}

async function openViewer(name: string, source: Uint8Array): Promise<HTMLElement> {
	bytes.current = source

	render(
		<QueryClientProvider client={new QueryClient()}>
			<SheetViewer
				item={file(name)}
				documentKey={name}
				alt={name}
				editable
				saveRef={{ current: null }}
			/>
		</QueryClientProvider>
	)

	await screen.findByRole("textbox", { name: preview.previewSpreadsheetCellContents })

	return await screen.findByRole("grid", {}, { timeout: 5000 })
}

async function openLinkDialog(): Promise<{ grid: HTMLElement; dialog: HTMLElement }> {
	const grid = await openViewer(
		"links.xlsx",
		await workbook([
			{
				name: "One",
				rows: [
					["a", "b"],
					["c", "d"]
				]
			}
		])
	)

	await goTo("B2")
	grid.focus()
	fireEvent.keyDown(grid, { key: "k", code: "KeyK", ctrlKey: true })

	return { grid, dialog: await screen.findByRole("dialog", {}, { timeout: 3000 }) }
}

async function closed(grid: HTMLElement): Promise<void> {
	await waitFor(() => {
		expect(screen.queryByRole("dialog")).toBeNull()
	})
	// Still there once the dialog's own closing work is done.
	await new Promise(resolve => setTimeout(resolve, 300))
	await waitFor(() => {
		expect(document.activeElement).toBe(grid)
	})
}

describe("the link dialog", () => {
	it("gives the grid its focus back when cancelled", { timeout: 30_000 }, async () => {
		const { grid, dialog } = await openLinkDialog()

		fireEvent.click(within(dialog).getByRole("button", { name: common.cancel }))
		await closed(grid)
	})

	it("gives the grid its focus back on Escape", { timeout: 30_000 }, async () => {
		const { grid, dialog } = await openLinkDialog()

		fireEvent.keyDown(within(dialog).getByRole("textbox", { name: preview.previewSheetLinkAddress }), { key: "Escape" })
		await closed(grid)
	})

	it("gives the grid its focus back when applied, and starts from the cell on the next opening", { timeout: 30_000 }, async () => {
		const { grid, dialog } = await openLinkDialog()

		fireEvent.change(within(dialog).getByRole("textbox", { name: preview.previewSheetLinkAddress }), { target: { value: "filen.io" } })
		fireEvent.click(within(dialog).getByRole("button", { name: preview.previewSheetLinkApply }))
		await closed(grid)

		await goTo("A1")
		grid.focus()
		fireEvent.keyDown(grid, { key: "k", code: "KeyK", ctrlKey: true })

		const next = await screen.findByRole("dialog", {}, { timeout: 3000 })

		expect(within(next).getByRole<HTMLInputElement>("textbox", { name: preview.previewSheetLinkAddress }).value).toBe("")
		expect(within(next).getByRole<HTMLInputElement>("textbox", { name: preview.previewSheetLinkText }).value).toBe("a")
	})
})
