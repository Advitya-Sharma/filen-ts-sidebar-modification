// @vitest-environment jsdom

import { webcrypto } from "node:crypto"
import { beforeEach, describe, expect, it, vi } from "vitest"
import { act, fireEvent, render, screen, waitFor } from "@testing-library/react"
import type { SheetCommand } from "@filen/office/core/sheet"
import type { SheetGridModel } from "@filen/office/react/sheet"
import { label, openGrid, workbook } from "@/tests/sheetMenusFixtures"
import { GridHarness } from "@/tests/sheetMenusHarness"

// The office engine viewer's sheet tabs: each tab's menu deletes, moves, hides and shows again, duplicates and
// protects sheets, and the workbook's structure protection turns those off.

const { toasts } = vi.hoisted(() => ({ toasts: [] as string[] }))

vi.mock("sonner", () => ({ toast: { error: (message: string) => toasts.push(message), success: () => undefined } }))

const { SheetTabs } = await import("@/features/spreadsheet/next/tabs/sheetTabs")

async function opened(): Promise<SheetGridModel> {
	return await openGrid(
		await workbook([
			{ name: "One", rows: [["a"]] },
			{ name: "Two", rows: [["b"]] },
			{ name: "Three", rows: [["c"]] }
		]),
		"xlsx"
	)
}

function names(model: SheetGridModel, state: "visible" | "hidden" = "visible"): string[] {
	return model
		.getSnapshot()
		.sheets.filter(sheet => sheet.state === state)
		.map(sheet => sheet.name)
}

function show(model: SheetGridModel) {
	const run = (command: SheetCommand): void => {
		void model.intent()
		void model.command(command)
	}

	return render(
		<GridHarness model={model}>
			{grid => (
				<SheetTabs
					grid={grid}
					canEdit
					run={run}
				/>
			)}
		</GridHarness>
	)
}

function tabMenu(name: string): void {
	fireEvent.contextMenu(screen.getByRole("tab", { name }))
}

async function choose(key: Parameters<typeof label>[0]): Promise<void> {
	fireEvent.click(await screen.findByText(label(key)))
}

describe("the sheet tabs' menu", () => {
	beforeEach(() => {
		toasts.length = 0
		vi.stubGlobal("crypto", webcrypto)
	})

	it("hides a sheet and shows it again from the unhide dialog", async () => {
		const model = await opened()

		show(model)
		tabMenu("Two")
		await choose("previewSheetHide")
		await waitFor(() => {
			expect(names(model)).toEqual(["One", "Three"])
		})
		expect(screen.queryByRole("tab", { name: "Two" })).toBeNull()

		tabMenu("One")
		await choose("previewSheetUnhide")
		fireEvent.click(await screen.findByRole("radio", { name: "Two" }))
		fireEvent.click(screen.getByRole("button", { name: label("previewSheetUnhideAction") }))
		await waitFor(() => {
			expect(names(model)).toEqual(["One", "Two", "Three"])
		})
		expect(model.getSnapshot().sheets.find(sheet => sheet.id === model.getSnapshot().sheet)?.name).toBe("Two")
	})

	it("deletes a sheet once confirmed", async () => {
		const model = await opened()

		show(model)
		tabMenu("Three")
		await choose("previewSheetDelete")
		fireEvent.click(await screen.findByRole("button", { name: label("previewSheetDelete") }))
		await waitFor(() => {
			expect(names(model)).toEqual(["One", "Two"])
		})
	})

	it("moves a sheet right, past the next tab shown", async () => {
		const model = await opened()

		show(model)
		tabMenu("One")
		await choose("previewSheetMoveRight")
		await waitFor(() => {
			expect(names(model)).toEqual(["Two", "One", "Three"])
		})
	})

	it("duplicates a sheet and shows the copy", async () => {
		const model = await opened()

		show(model)
		tabMenu("One")
		await choose("previewSheetDuplicate")
		await waitFor(() => {
			expect(names(model)).toHaveLength(4)
		})

		const snapshot = model.getSnapshot()
		const copy = snapshot.sheets.find(sheet => !["One", "Two", "Three"].includes(sheet.name))

		expect(copy).toBeDefined()
		expect(snapshot.sheet).toBe(copy?.id)
	})

	it("protects a sheet with a password and takes it off only with that password", async () => {
		const model = await opened()

		show(model)
		tabMenu("One")
		await choose("previewSheetProtect")
		fireEvent.change(await screen.findByLabelText(label("previewSheetProtectPassword")), { target: { value: "secret" } })
		fireEvent.change(screen.getByLabelText(label("previewSheetProtectConfirm")), { target: { value: "secret" } })
		fireEvent.click(screen.getByRole("button", { name: label("previewSheetProtectAction") }))
		await waitFor(
			() => {
				expect(model.getSnapshot().layout?.protection?.hasPassword).toBe(true)
			},
			{ timeout: 20_000 }
		)

		tabMenu("One")
		await choose("previewSheetUnprotect")

		const field = await screen.findByLabelText(label("previewSheetUnprotectPassword"))

		fireEvent.change(field, { target: { value: "wrong" } })
		fireEvent.click(screen.getByRole("button", { name: label("previewSheetUnprotectAction") }))
		expect(await screen.findByText(label("previewSheetWrongPassword"), undefined, { timeout: 20_000 })).toBeTruthy()
		expect(model.getSnapshot().layout?.protection).not.toBeNull()

		fireEvent.change(field, { target: { value: "secret" } })
		fireEvent.click(screen.getByRole("button", { name: label("previewSheetUnprotectAction") }))
		await waitFor(
			() => {
				expect(model.getSnapshot().layout?.protection).toBeNull()
			},
			{ timeout: 20_000 }
		)
	}, 60_000)

	it("protects without a password once a typed one is cleared", async () => {
		const model = await opened()

		show(model)
		tabMenu("One")
		await choose("previewSheetProtect")

		const password = await screen.findByLabelText(label("previewSheetProtectPassword"))

		fireEvent.change(password, { target: { value: "secret" } })
		fireEvent.change(screen.getByLabelText(label("previewSheetProtectConfirm")), { target: { value: "secret" } })
		fireEvent.change(password, { target: { value: "" } })
		expect(screen.getByLabelText(label("previewSheetProtectConfirm"))).toHaveProperty("value", "")
		fireEvent.click(screen.getByRole("button", { name: label("previewSheetProtectAction") }))
		await waitFor(() => {
			expect(model.getSnapshot().layout?.protection).toEqual(expect.objectContaining({ hasPassword: false }))
		})
	})

	it("stops checking a file's password when the unprotect prompt is cancelled", async () => {
		const model = await openGrid(await workbook([{ name: "One", rows: [["a"]] }]), "xlsx")

		// A file's own hash, spun as many times as Office allows.
		void model.intent()
		await model.command({
			type: "sheet.protect",
			sheet: model.getSnapshot().sheet,
			allows: [],
			hash: {
				algorithmName: "SHA-512",
				hashValue: `${"A".repeat(86)}==`,
				saltValue: "AAAAAAAAAAAAAAAAAAAAAA==",
				spinCount: 10_000_000
			}
		})

		const digest = vi.spyOn(webcrypto.subtle, "digest")

		show(model)
		tabMenu("One")
		await choose("previewSheetUnprotect")
		fireEvent.change(await screen.findByLabelText(label("previewSheetUnprotectPassword")), { target: { value: "x" } })
		fireEvent.click(screen.getByRole("button", { name: label("previewSheetUnprotectAction") }))
		await waitFor(() => {
			expect(digest.mock.calls.length).toBeGreaterThan(64)
		})
		fireEvent.click(screen.getByRole("button", { name: /^(Cancel|cancel)$/ }))

		const spun = digest.mock.calls.length

		await new Promise(resolve => setTimeout(resolve, 100))
		expect(digest.mock.calls.length - spun).toBeLessThan(200)
		expect(model.getSnapshot().layout?.protection).not.toBeNull()
		digest.mockRestore()
	})

	it("turns the sheet changes off while the workbook's structure is protected", async () => {
		const model = await opened()

		show(model)
		tabMenu("One")
		await choose("previewSheetProtectWorkbook")
		fireEvent.click(await screen.findByRole("button", { name: label("previewSheetProtectAction") }))
		await waitFor(async () => {
			const state = await model.session.query("commandState", { sheet: model.getSnapshot().sheet, ranges: [], types: ["sheet.add"] })

			expect(state.reasons[0]).toBe("protectedWorkbook")
		})

		await act(async () => {
			await new Promise(resolve => setTimeout(resolve, 50))
		})
		tabMenu("Two")

		const insert = (await screen.findByText(label("previewSheetInsertSheet"))).closest("[role=menuitem]")

		await waitFor(() => {
			expect(insert?.getAttribute("aria-disabled") ?? insert?.hasAttribute("data-disabled")).toBeTruthy()
		})
		expect(await screen.findByText(label("previewSheetUnprotectWorkbook"))).toBeTruthy()
	})
})
