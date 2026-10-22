// @vitest-environment jsdom

import { beforeEach, describe, expect, it, vi } from "vitest"
import { act, fireEvent, render, screen } from "@testing-library/react"
import { writeXlsx } from "hucre/xlsx"
import { inThreadSheetHost } from "@filen/office/dom/sheet"
import type { DriveItem } from "@/features/drive/lib/item"
import { menuGroups, rendersItems } from "@/features/spreadsheet/components/sheetExtension"
import { registerAllActions } from "@/lib/keymap/actions"

// The grid's menu puts a separator only between groups that show something: none first, last or twice in a row,
// also when a feature's group comes out empty (a read-only file, a cell with nothing to offer).

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

describe("the grid menu's separators", () => {
	it("leaves out a group whose parts all came out empty", () => {
		expect(rendersItems(null)).toBe(false)
		expect(
			rendersItems(
				<>
					{false}
					{null}
					<>{undefined}</>
				</>
			)
		).toBe(false)
		expect(rendersItems(<>{[null, <span key="a">a</span>]}</>)).toBe(true)
		expect(
			menuGroups([{ id: "empty", menuItems: <>{null}</> }, { id: "none" }, { id: "item", menuItems: <span>a</span> }]).map(
				group => group.id
			)
		).toEqual(["item"])
	})

	it("starts, ends and runs without a separator on a read-only file", async () => {
		bytes.current = await writeXlsx({ sheets: [{ name: "One", rows: [["a", "b"]] }] })
		render(
			<SheetViewer
				item={{ type: "file", data: { uuid: "u-ro", decryptedMeta: { name: "ro.xlsx" }, size: 1n } } as unknown as DriveItem}
				documentKey="ro"
				alt="ro.xlsx"
				editable={false}
				saveRef={{ current: null }}
			/>
		)

		const grid = await screen.findByRole("grid", {}, { timeout: 5000 })

		await act(async () => {
			fireEvent.contextMenu(grid, { clientX: 100, clientY: 60 })
			await new Promise(resolve => setTimeout(resolve, 100))
		})

		const menu = await screen.findByRole("menu", {}, { timeout: 3000 })
		const kinds = [...menu.children].map(child => (child.getAttribute("role") === "separator" ? "|" : "item")).join(" ")

		expect(kinds).toContain("item")
		expect(kinds).not.toMatch(/\| \||^\||\|$/)
	})
})
