// @vitest-environment jsdom

import { describe, expect, it, vi } from "vitest"
import { act, fireEvent, render, screen, waitFor } from "@testing-library/react"
import { useEffect } from "react"
import { CellFlag } from "@filen/office/core/sheet"
import type { SheetCellPointer, SheetGridHandle, SheetGridModel, SheetGridState, SheetShortcut } from "@filen/office/react/sheet"
import type { SheetHandle } from "@/features/spreadsheet/hooks/useSheetSession"
import type { SheetBodyExtension } from "@/features/spreadsheet/components/sheetExtension"
import { openGrid, workbook } from "@/tests/sheetMenusFixtures"
import { GridHarness } from "@/tests/sheetMenusHarness"
import { previewSheetNotes } from "@/locales/en/previewSheet/notes"

// Notes, comments and links in the sheet view, over a session on the in-thread host: Shift+F2 writes a note under
// the documents settings' author (its text starting with the author's line), a pointer resting on the cell shows it,
// Ctrl/Cmd+K links the cell, and Ctrl/Cmd+click on the link goes through the untrusted-domain confirmation.

const { intercept } = vi.hoisted(() => ({ intercept: vi.fn((_href: string) => true) }))

vi.mock("@/features/office/lib/officeSettings", () => ({
	useOfficeSettings: () => ({ authorName: "Jan", autosave: false, loaded: true }),
	ensureAuthorName: () => Promise.resolve("Jan")
}))

vi.mock("@/lib/trustedLinks/useTrustedLinkGate", () => ({
	useTrustedLinkGate: () => ({ intercept, onClick: () => undefined, onAuxClick: () => undefined, dialog: null })
}))

const { useSheetAnnotations } = await import("@/features/spreadsheet/comments/useSheetAnnotations")

type NotesKey = keyof typeof previewSheetNotes

// A string as shown: its English text once the catalog holds it, else its key.
function label(key: NotesKey): RegExp {
	const escape = (text: string): string => text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")

	return new RegExp(`^(${escape(previewSheetNotes[key])}|${escape(key)})$`)
}

// The grid's handle as overlays use it: every cell on screen at a fixed box.
const HANDLE = {
	cellRect: () => ({ x: 40, y: 20, width: 64, height: 20 }),
	onFrame: () => () => undefined,
	focus: () => undefined
} as unknown as SheetGridHandle

interface Probe {
	pointer: ((event: SheetCellPointer) => unknown) | null
	shortcut: ((key: SheetShortcut) => boolean) | null
}

function Annotations({ grid, onExtension }: { grid: SheetGridState; onExtension: (extension: SheetBodyExtension) => void }) {
	const extension = useSheetAnnotations({
		grid,
		gridRef: { current: HANDLE },
		handle: { session: grid.model.session } as unknown as SheetHandle,
		canEdit: true,
		text: false
	})

	useEffect(() => {
		onExtension(extension)
	})

	return (
		<>
			<div style={{ position: "relative" }}>{extension.overlay}</div>
			{extension.dialogs}
		</>
	)
}

function select(model: SheetGridModel, row: number, col: number): void {
	act(() => {
		model.choose({
			active: { row, col },
			anchor: { row, col },
			focus: { row, col },
			ranges: [{ startRow: row, startCol: col, endRow: row, endCol: col }]
		})
	})
}

describe("sheet notes and links in the view", () => {
	it(
		"writes a note from Shift+F2 under the author, shows it on hover, links a cell and asks before following it",
		{ timeout: 30_000 },
		async () => {
			const model = await openGrid(
				await workbook([
					{
						name: "One",
						rows: [
							["a", "b"],
							["c", "d"]
						]
					}
				]),
				"xlsx"
			)
			const probe: Probe = { pointer: null, shortcut: null }

			render(
				<GridHarness model={model}>
					{grid => (
						<Annotations
							grid={grid}
							onExtension={extension => {
								probe.pointer = extension.grid?.pointer ?? null
								probe.shortcut = extension.shortcut ?? null
							}}
						/>
					)}
				</GridHarness>
			)

			select(model, 1, 1)
			act(() => {
				expect(probe.shortcut?.("note")).toBe(true)
			})

			const field = await screen.findByRole("textbox", { name: label("previewSheetNoteText") }, { timeout: 10_000 })

			expect((field as HTMLTextAreaElement).value).toBe("Jan:\n")
			fireEvent.change(field, { target: { value: "Jan:\nlooks right" } })
			fireEvent.click(screen.getByRole("button", { name: label("previewSheetNoteSave") }))

			await waitFor(async () => {
				expect(await model.session.query("note.get", { sheet: model.getSnapshot().sheet, row: 1, col: 1 })).toMatchObject({
					text: "Jan:\nlooks right",
					author: "Jan"
				})
			})

			// A pointer resting on the cell shows the note, read-only.
			await waitFor(() => {
				expect(model.entry(model.getSnapshot().sheet, 1, 1)).not.toBeNull()
			})
			act(() => {
				probe.pointer?.({ kind: "hover", row: 1, col: 1, clientX: 0, clientY: 0 })
			})

			const card = await screen.findByRole("tooltip", {}, { timeout: 2000 })

			expect(card.textContent).toContain("looks right")
			act(() => {
				probe.pointer?.({ kind: "leave" })
			})
			await waitFor(() => {
				expect(screen.queryByRole("tooltip")).toBeNull()
			})

			// Ctrl/Cmd+K links the cell; Ctrl/Cmd+click on it asks before the domain opens.
			select(model, 0, 0)
			act(() => {
				expect(probe.shortcut?.("hyperlink")).toBe(true)
			})
			fireEvent.change(await screen.findByRole("textbox", { name: label("previewSheetLinkAddress") }), {
				target: { value: "filen.io/docs" }
			})
			fireEvent.click(screen.getByRole("button", { name: label("previewSheetLinkApply") }))

			await waitFor(async () => {
				expect(await model.session.query("link.get", { sheet: model.getSnapshot().sheet, row: 0, col: 0 })).toMatchObject({
					url: "https://filen.io/docs"
				})
			})
			await waitFor(() => {
				expect(model.getSnapshot().active.flags & CellFlag.HYPERLINK).not.toBe(0)
			})
			act(() => {
				probe.pointer?.({ kind: "hover", row: 0, col: 0, clientX: 0, clientY: 0 })
			})
			await waitFor(() => {
				const taken = probe.pointer?.({
					kind: "down",
					row: 0,
					col: 0,
					clientX: 0,
					clientY: 0,
					pointerType: "mouse",
					ctrlKey: true,
					metaKey: true,
					shiftKey: false,
					altKey: false,
					editing: false
				})

				expect(taken).toBe(true)
				expect(intercept).toHaveBeenCalledWith("https://filen.io/docs")
			})
		}
	)
})
