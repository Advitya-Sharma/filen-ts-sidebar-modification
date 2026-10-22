// @vitest-environment happy-dom

import { afterEach, describe, expect, it, vi } from "vitest"
import { useState } from "react"
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react"
import { writeXlsx } from "hucre/xlsx"
import { inThreadSheetHost } from "@filen/office/dom/sheet"
import type { SheetOpened, SheetRenderStyle } from "@filen/office/core/sheet"
import { createSheetGrid, useSheetGrid, type SheetGridModel } from "@filen/office/react/sheet"
import { act } from "react"
import { borderShortcut, setFormatCellsOpen } from "@/features/spreadsheet/toolbar/formatCellsState"

// The spreadsheet toolbar over a real session on this thread: it shows the active cell's format, and what its
// controls do reaches the session (formats, borders, decimals, the Format Cells dialog's code).

const size = vi.hoisted(() => ({ width: 2000 }))

vi.mock("@/lib/useElementSize", () => ({ useElementSize: () => ({ width: size.width, height: 40 }) }))

// What the grid's last frame drew for a General number it rounded (a narrow cell), as the toolbar asks it.
const rounded = vi.hoisted(() => ({ text: undefined as string | undefined }))

vi.mock("@filen/office/react/sheet", async importOriginal => ({
	...(await importOriginal<typeof import("@filen/office/react/sheet")>()),
	roundedGeneral: () => rounded.text
}))

// What the Insert menu asks the features that insert for.
const requests = vi.hoisted(() => ({ image: vi.fn(), shortcut: vi.fn() }))

vi.mock("@/features/spreadsheet/objects/insertRequests", () => ({
	requestImage: requests.image,
	requestShortcut: requests.shortcut
}))

const { SheetToolbar } = await import("@/features/spreadsheet/toolbar/sheetToolbar")

afterEach(() => {
	cleanup()
	size.width = 2000
	rounded.text = undefined
})

async function opened(): Promise<SheetOpened> {
	const bytes = await writeXlsx({
		sheets: [
			{
				name: "Sheet1",
				rows: [
					[3.14159, "text"],
					[2, 3]
				]
			}
		]
	})
	const result = await inThreadSheetHost().open(bytes, { kind: "xlsx" }, () => undefined)

	if (result.kind !== "opened") {
		throw new Error(result.kind)
	}

	await result.session.intent()

	return result
}

function Harness({ open, onModel, disabled = false }: { open: SheetOpened; onModel: (model: SheetGridModel) => void; disabled?: boolean }) {
	const [model] = useState(() => {
		const created = createSheetGrid(open)

		// The grid's first tile, as a drawn grid would hold it: changes patch it and bring their styles.
		created.need(open.info.activeSheet, [0, 0])
		onModel(created)

		return created
	})
	const grid = useSheetGrid(model)

	return (
		<SheetToolbar
			grid={grid}
			disabled={disabled}
			formats
			onFormat={patch => {
				void model.command({ type: "format.apply", sheet: grid.sheet, ranges: grid.selection.ranges, patch })
			}}
		/>
	)
}

async function mounted(): Promise<{
	model: SheetGridModel
	open: SheetOpened
	style: () => Promise<SheetRenderStyle>
	numFmt: () => Promise<string>
}> {
	const open = await opened()
	let model: SheetGridModel | null = null

	render(
		<Harness
			open={open}
			onModel={created => {
				model = created
			}}
		/>
	)

	const grid = model as unknown as SheetGridModel
	const cell = () => open.session.query("cell", { sheet: open.info.activeSheet, row: 0, col: 0 })

	await waitFor(() => {
		expect(grid.entry(open.info.activeSheet, 0, 0)).not.toBeNull()
	})

	// The controls turn on once the active cell's format is known.
	await waitFor(() => {
		expect(screen.getByRole("button", { name: "Text color" }).hasAttribute("disabled")).toBe(false)
	})

	return {
		model: grid,
		open,
		style: async () => grid.style((await cell()).style),
		numFmt: async () => (await cell()).numFmt
	}
}

describe("spreadsheet toolbar", () => {
	it("lays out Excel's groups in one row: history, font, toggles, colours, borders, alignment, merge, number, freeze, sort and filter, insert, find", async () => {
		await mounted()

		const toolbar = screen.getByRole("toolbar", { name: "Formatting" })
		const names = Array.from(toolbar.querySelectorAll("button, [role=combobox]")).map(element => element.getAttribute("aria-label"))

		expect(names).toEqual([
			"Undo",
			"Redo",
			"Font",
			"Font size",
			"Bold",
			"Italic",
			"Underline",
			"Strikethrough",
			"Text color",
			"Fill color",
			"Borders",
			"Horizontal alignment",
			"Vertical alignment",
			"Wrap text",
			"Merge",
			"Number format",
			"Decrease decimal",
			"Increase decimal",
			"Percent style",
			"Freeze",
			"Sort and filter",
			"Insert",
			"Find and replace",
			"More formatting"
		])
	})

	it("toggles bold on the selection as the active cell shows it", async () => {
		const { style } = await mounted()

		fireEvent.click(screen.getByRole("button", { name: "Bold" }))
		await waitFor(async () => {
			expect((await style()).bold).toBe(true)
		})
		await waitFor(() => {
			expect(screen.getByRole("button", { name: "Bold" }).getAttribute("aria-pressed")).toBe("true")
		})
		fireEvent.click(screen.getByRole("button", { name: "Bold" }))
		await waitFor(async () => {
			expect((await style()).bold).toBeUndefined()
		})
	})

	it("adds a decimal place to a General number from the places it shows", async () => {
		const { numFmt } = await mounted()

		fireEvent.click(screen.getByRole("button", { name: "Increase decimal" }))
		await waitFor(async () => {
			expect(await numFmt()).toBe("0.000000")
		})
		fireEvent.click(screen.getByRole("button", { name: "Decrease decimal" }))
		await waitFor(async () => {
			expect(await numFmt()).toBe("0.00000")
		})
	})

	it("adds a decimal place to a General number a narrow cell rounds from the places the cell shows", async () => {
		const { numFmt } = await mounted()

		rounded.text = "3.14"
		fireEvent.click(screen.getByRole("button", { name: "Increase decimal" }))
		await waitFor(async () => {
			expect(await numFmt()).toBe("0.000")
		})
	})

	it("draws all borders from the Borders menu in the chosen line style", async () => {
		const { style } = await mounted()

		fireEvent.click(screen.getByRole("button", { name: "Borders" }))
		fireEvent.click(await screen.findByRole("menuitem", { name: "Line style" }))
		fireEvent.click(await screen.findByRole("menuitemradio", { name: "Dashed" }))
		fireEvent.click(screen.getByRole("button", { name: "Borders" }))
		fireEvent.click(await screen.findByRole("menuitem", { name: "All borders" }))
		await waitFor(async () => {
			const look = await style()

			expect([look.top?.style, look.right?.style, look.bottom?.style, look.left?.style]).toEqual([
				"dashed",
				"dashed",
				"dashed",
				"dashed"
			])
		})
	})

	it("applies a number format from the menu and a custom code from the Format Cells dialog", async () => {
		const { numFmt } = await mounted()

		fireEvent.click(screen.getByRole("button", { name: "Number format" }))
		fireEvent.click(await screen.findByRole("menuitemradio", { name: /^Percentage/u }))
		await waitFor(async () => {
			expect(await numFmt()).toBe("0.00%")
		})
		fireEvent.click(screen.getByRole("button", { name: "Number format" }))
		fireEvent.click(await screen.findByRole("menuitem", { name: "More number formats…" }))

		const code = await screen.findByRole("textbox", { name: "Format code" })

		fireEvent.change(code, { target: { value: '0.0" kg' } })
		expect(await screen.findByText("Close every quote and bracket, and follow \\, _ and * with a character.")).toBeTruthy()
		expect(screen.getByRole("button", { name: "Apply" }).hasAttribute("disabled")).toBe(true)
		// A pasted control character: a file cannot hold it, so the dialog refuses it as the session would.
		fireEvent.change(code, { target: { value: "0.0\u0001" } })
		expect(await screen.findByText("Remove the control characters from the format code.")).toBeTruthy()
		fireEvent.change(code, { target: { value: '0.0" kg"' } })
		expect(screen.getByLabelText("Sample").textContent).toBe("3.1 kg")
		fireEvent.click(screen.getByRole("button", { name: "Apply" }))
		await waitFor(async () => {
			expect(await numFmt()).toBe('0.0" kg"')
		})
	})

	it("moves the last groups into the ⋯ menu on a narrow row and keeps them working there", async () => {
		size.width = 520

		const { style } = await mounted()

		expect(screen.queryByRole("button", { name: "Percent style" })).toBeNull()
		expect(screen.queryByRole("button", { name: "Horizontal alignment" })).toBeNull()
		fireEvent.click(screen.getByRole("button", { name: "More formatting" }))
		fireEvent.click(await screen.findByRole("menuitem", { name: "Horizontal alignment" }))
		fireEvent.click(await screen.findByRole("menuitemradio", { name: "Center" }))
		await waitFor(async () => {
			expect((await style()).align).toBe("center")
		})
	})

	it("opens Format Cells from what opens it outside the toolbar (Ctrl+1)", async () => {
		const { model } = await mounted()

		expect(screen.queryByRole("dialog")).toBeNull()
		act(() => {
			setFormatCellsOpen(model, true)
		})
		expect(await screen.findByRole("dialog", { name: "Format cells" })).toBeTruthy()
		fireEvent.click(screen.getByRole("button", { name: "Cancel" }))
		await waitFor(() => {
			expect(screen.queryByRole("dialog", { name: "Format cells" })).toBeNull()
		})
	})

	it("does not keep Ctrl+1 waiting while the toolbar has no cell format to start from", async () => {
		const open = await opened()
		let model: SheetGridModel | null = null
		const onModel = (created: SheetGridModel): void => {
			model = created
		}
		const { rerender } = render(
			<Harness
				open={open}
				onModel={onModel}
				disabled
			/>
		)
		const grid = model as unknown as SheetGridModel

		await waitFor(() => {
			expect(grid.entry(open.info.activeSheet, 0, 0)).not.toBeNull()
		})
		act(() => {
			setFormatCellsOpen(grid, true)
		})
		rerender(
			<Harness
				open={open}
				onModel={onModel}
			/>
		)
		await waitFor(() => {
			expect(screen.getByRole("button", { name: "Text color" }).hasAttribute("disabled")).toBe(false)
		})
		expect(screen.queryByRole("dialog", { name: "Format cells" })).toBeNull()
	})

	it("names the Borders menu's groups for screen readers", async () => {
		await mounted()

		fireEvent.click(screen.getByRole("button", { name: "Borders" }))

		for (const name of ["Edges", "Outlines", "Top and bottom", "Inside and diagonal"]) {
			expect(await screen.findByRole("group", { name })).toBeTruthy()
		}
	})

	it("draws an outline and takes borders off with the shortcuts' command", async () => {
		const { model, open, style } = await mounted()
		const ranges = [{ startRow: 0, startCol: 0, endRow: 1, endCol: 1 }]

		await model.command(borderShortcut("borderOutline", open.info.activeSheet, ranges))
		await waitFor(async () => {
			const look = await style()

			expect([look.top?.style, look.left?.style, look.bottom, look.right]).toEqual(["thin", "thin", undefined, undefined])
		})
		await model.command(borderShortcut("borderNone", open.info.activeSheet, ranges))
		await waitFor(async () => {
			const look = await style()

			expect([look.top, look.left]).toEqual([undefined, undefined])
		})
	})

	it("clears the selection's formats from the ⋯ menu, keeping its contents", async () => {
		const { style, numFmt } = await mounted()

		fireEvent.click(screen.getByRole("button", { name: "Bold" }))
		await waitFor(async () => {
			expect((await style()).bold).toBe(true)
		})
		fireEvent.click(screen.getByRole("button", { name: "More formatting" }))
		fireEvent.click(await screen.findByRole("menuitem", { name: "Clear formats" }))
		await waitFor(async () => {
			expect((await style()).bold).toBeUndefined()
		})
		expect(await numFmt()).toBe("General")
	})

	it("asks for a picture from this device or the drive, a link and a note from the Insert menu", async () => {
		const { model } = await mounted()
		const pick = async (name: string): Promise<void> => {
			fireEvent.click(screen.getByRole("button", { name: "Insert" }))
			fireEvent.click(await screen.findByRole("menuitem", { name }))
		}

		await pick("Image from file…")
		await pick("Image from drive…")
		await pick("Link…")
		await pick("Note")
		expect(requests.image.mock.calls).toEqual([
			[model, "file"],
			[model, "drive"]
		])
		expect(requests.shortcut.mock.calls).toEqual([
			[model, "hyperlink"],
			[model, "note"]
		])
		// The chart types are a submenu of their own.
		fireEvent.click(screen.getByRole("button", { name: "Insert" }))
		expect(await screen.findByRole("menuitem", { name: "Chart" })).toBeDefined()
	})
})
