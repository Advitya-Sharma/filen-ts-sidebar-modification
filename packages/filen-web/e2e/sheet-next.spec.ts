import { readFileSync } from "node:fs"
import { fileURLToPath } from "node:url"
import type { Locator, Page } from "@playwright/test"
import { writeXlsx } from "hucre/xlsx"
import { test, expect } from "./fixtures"
import { withScratchDirectory, LAZY_VIEWER_TIMEOUT_MS, LIVE_WRITE_TIMEOUT_MS, uploadFiles } from "./helpers/listing"
import { resolveModKey } from "./helpers/modkey"
import { trackCspViolations } from "./helpers/csp"
import { storedPassword, trackPasswordCanary } from "./helpers/passwordCanary"
import { WORKBOOK_PASSWORD, encryptedWorkbook } from "@/tests/encryptedWorkbookFixtures"
import { preview } from "@/locales/en/preview"

// The spreadsheet viewer on the office engine end to end, in builds with VITE_SHEET_NEXT=1: a real open in the
// sheet workers (the workbook reader and HyperFormula under the app's CSP, which forbids eval), the canvas grid,
// entries, formats, rows and sheets, and saving. Each leg reloads after its save, so what it reads back is the
// file the drive now holds. The grid draws its cells: a cell is read as the active cell, which the grid names
// for assistive technology, after the name box went to it.
test.skip(process.env["VITE_SHEET_NEXT"] !== "1", "the office engine's viewer is only in builds with VITE_SHEET_NEXT=1")
test.describe.configure({ mode: "default" })

const XLS_FIXTURE = fileURLToPath(new URL("../src/tests/fixtures/spreadsheet/budget.xls", import.meta.url))
const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
const CSV_BYTES = Buffer.from("Name;Qty\r\nApples;3\r\nPears;5\r\n", "utf8")

async function xlsxBytes(): Promise<Buffer> {
	return Buffer.from(
		await writeXlsx({
			sheets: [
				{
					name: "Budget",
					rows: [
						["Item", "Cost"],
						["Rent", 1200],
						["Food", 300],
						["Total", 1500]
					],
					cells: new Map([["3,1", { value: 1500, type: "formula", formula: "SUM(B2:B3)", formulaResult: 1500 }]])
				},
				{ name: "Notes", rows: [["second sheet"]] }
			]
		})
	)
}

// Makes `ref` ("B4") the active cell through the name box; the grid has focus after.
async function goTo(dialog: Locator, ref: string): Promise<void> {
	const box = dialog.getByRole("textbox", { name: preview.previewSheetNameBox })

	await box.click()
	await box.fill(ref)
	await box.press("Enter")
	await expect(box).toHaveValue(ref)
	await expect(dialog.getByRole("grid")).toBeFocused()
}

// The active cell's shown text.
function activeCell(dialog: Locator): Locator {
	return dialog.getByRole("grid").getByRole("gridcell")
}

async function expectCell(dialog: Locator, ref: string, text: string, timeout?: number): Promise<void> {
	await goTo(dialog, ref)
	await expect(activeCell(dialog)).toHaveText(text, timeout === undefined ? {} : { timeout })
}

async function typeInto(page: Page, dialog: Locator, ref: string, text: string): Promise<void> {
	await goTo(dialog, ref)
	await page.keyboard.type(text)
	await page.keyboard.press("Enter")
}

// Opens the grid's menu on the active cell's row header (row `row`, 1-based), the way a right-click there does.
async function rowMenu(dialog: Locator, row: number): Promise<void> {
	await goTo(dialog, `A${String(row)}`)
	// Column headers 20 px, rows 20 px at 100 %: the middle of row `row`'s header.
	await dialog.getByRole("grid").click({ button: "right", position: { x: 8, y: 20 + (row - 1) * 20 + 10 } })
}

test("csv and xlsx open on the office engine, edit, format, insert a row, add and rename a sheet, save and reopen", async ({ page }) => {
	test.slow()

	const cspViolations = trackCspViolations(page)
	const dialog = page.getByRole("dialog").first()
	const saveButton = dialog.getByRole("button", { name: "Save", exact: true })
	const bold = dialog.getByRole("button", { name: preview.previewSpreadsheetBold, exact: true })
	const mod = await resolveModKey(page)

	await withScratchDirectory(page, "sheet-next", async ({ listbox, runId }) => {
		const nameCsv = `e2e-sheet-next-${runId}.csv`
		const nameXlsx = `e2e-sheet-next-${runId}.xlsx`

		await uploadFiles(
			page,
			[
				{ name: nameCsv, mimeType: "text/csv", buffer: CSV_BYTES },
				{ name: nameXlsx, mimeType: XLSX_MIME, buffer: await xlsxBytes() }
			],
			listbox
		)

		// CSV: the delimiter is sniffed, an entry marks the file dirty, and the save keeps it.
		const csvRow = listbox.getByRole("option", { name: nameCsv })

		await csvRow.dblclick()
		await expectCell(dialog, "A2", "Apples", LAZY_VIEWER_TIMEOUT_MS)
		await expectCell(dialog, "B1", "Qty")
		await typeInto(page, dialog, "B4", "42")
		await expectCell(dialog, "B4", "42")
		await expect(saveButton).toBeEnabled()
		await saveButton.click()
		await expect(saveButton).toHaveCount(0, { timeout: LIVE_WRITE_TIMEOUT_MS })
		await page.keyboard.press("Escape")
		await expect(dialog).toHaveCount(0)

		// XLSX: a formula recalculates as its inputs change, undo takes it back.
		const xlsxRow = listbox.getByRole("option", { name: nameXlsx })

		await xlsxRow.dblclick()
		await expectCell(dialog, "B4", "1500", LAZY_VIEWER_TIMEOUT_MS)
		await expect(dialog.getByRole("tab", { name: "Notes" })).toBeVisible()
		await expect(bold).toBeEnabled()

		await typeInto(page, dialog, "B3", "500")
		await expectCell(dialog, "B4", "1700", LAZY_VIEWER_TIMEOUT_MS)
		await page.keyboard.press(`${mod}+Z`)
		await expect(activeCell(dialog)).toHaveText("1500")
		await page.keyboard.press(`${mod}+Shift+Z`)
		await expect(activeCell(dialog)).toHaveText("1700")

		await typeInto(page, dialog, "B5", "=B4*2")
		await expectCell(dialog, "B5", "3400")
		await expect(dialog.getByRole("textbox", { name: preview.previewSpreadsheetCellContents })).toHaveValue("=B4*2")

		// A format, from the toolbar.
		await goTo(dialog, "A1")
		await bold.click()
		await expect(bold).toHaveAttribute("aria-pressed", "true")

		// A row inserted above row 2 moves the rows under it, formulas included.
		await rowMenu(dialog, 2)
		await page.getByRole("menuitem", { name: "Insert row above" }).click()
		await expectCell(dialog, "A3", "Rent")
		await expectCell(dialog, "B5", "1700")
		await expectCell(dialog, "B6", "3400")
		await expect(dialog.getByRole("textbox", { name: preview.previewSpreadsheetCellContents })).toHaveValue("=B5*2")
		await expectCell(dialog, "A2", "")

		// A sheet added, then renamed.
		await dialog.getByRole("button", { name: preview.previewSpreadsheetAddSheet }).click()
		await expect(dialog.getByRole("tab", { name: "Sheet 3" })).toBeVisible()
		await dialog.getByRole("tab", { name: "Sheet 3" }).dblclick()

		const rename = page.getByRole("dialog", { name: preview.previewSpreadsheetRenameSheet })

		await rename.getByLabel(preview.previewSpreadsheetSheetName).fill("Summary")
		await rename.getByRole("button", { name: preview.previewSpreadsheetRenameSheet }).click()
		await expect(rename).toHaveCount(0)
		await expect(dialog.getByRole("tab", { name: "Summary" })).toBeVisible()

		// Saved from the grid, as a user would after renaming.
		await goTo(dialog, "A1")
		await page.keyboard.press(`${mod}+S`)
		await expect(saveButton).toHaveCount(0, { timeout: LIVE_WRITE_TIMEOUT_MS })
		await page.keyboard.press("Escape")
		await expect(dialog).toHaveCount(0)

		// What the drive now holds, read fresh.
		await page.reload()
		await expect(xlsxRow).toBeVisible({ timeout: LIVE_WRITE_TIMEOUT_MS })
		await xlsxRow.dblclick()
		await expectCell(dialog, "B5", "1700", LAZY_VIEWER_TIMEOUT_MS)
		await expectCell(dialog, "B6", "3400")
		await expectCell(dialog, "A3", "Rent")
		await goTo(dialog, "A1")
		await expect(bold).toHaveAttribute("aria-pressed", "true")
		await expect(dialog.getByRole("tab", { name: "Summary" })).toBeVisible()

		// A sheet other than the active one is read once first shown.
		await dialog.getByRole("tab", { name: "Notes" }).click()
		await expectCell(dialog, "A1", "second sheet")
		await page.keyboard.press("Escape")
		await expect(dialog).toHaveCount(0)

		await csvRow.dblclick()
		await expectCell(dialog, "B4", "42", LAZY_VIEWER_TIMEOUT_MS)
		await page.keyboard.press("Escape")
		await expect(dialog).toHaveCount(0)

		expect(cspViolations).toEqual([])
	})
})

test("an .xls opens view-only on the office engine and saves as an editable .xlsx beside it", async ({ page }) => {
	const dialog = page.getByRole("dialog").first()

	await withScratchDirectory(page, "sheet-next-xls", async ({ listbox, runId }) => {
		const nameXls = `e2e-sheet-next-${runId}.xls`
		const nameXlsx = `e2e-sheet-next-${runId}.xlsx`
		const xlsRowName = new RegExp(`${runId}\\.xls(?!x)`)

		await uploadFiles(page, [{ name: nameXls, mimeType: "application/vnd.ms-excel", buffer: readFileSync(XLS_FIXTURE) }])
		await expect(listbox.getByRole("option", { name: xlsRowName })).toBeVisible({ timeout: LIVE_WRITE_TIMEOUT_MS })

		await listbox.getByRole("option", { name: xlsRowName }).dblclick()
		await expectCell(dialog, "A2", "Rent", LAZY_VIEWER_TIMEOUT_MS)
		await expect(dialog.getByText(preview.previewSpreadsheetReadOnlyXls)).toBeVisible()

		await dialog.getByRole("button", { name: preview.previewSpreadsheetSaveAsXlsx, exact: true }).click()

		const confirm = page.getByRole("alertdialog", { name: preview.previewSpreadsheetSaveAsXlsxTitle })

		await confirm.getByRole("button", { name: preview.previewSpreadsheetSaveAsXlsx, exact: true }).click()
		await expect(confirm).toHaveCount(0, { timeout: LIVE_WRITE_TIMEOUT_MS })

		// The copy takes the preview's place, with the same values, and is editable.
		await expect(dialog.getByText(nameXlsx)).toBeVisible({ timeout: LIVE_WRITE_TIMEOUT_MS })
		await expectCell(dialog, "B2", "1200", LAZY_VIEWER_TIMEOUT_MS)
		await expect(dialog.getByRole("button", { name: preview.previewSpreadsheetBold, exact: true })).toBeEnabled()

		await page.keyboard.press("Escape")
		await expect(dialog).toHaveCount(0)
		await expect(listbox.getByRole("option", { name: nameXlsx })).toBeVisible()
	})
})

test("a password-protected xlsx is asked its password on the office engine, saved sealed and asked again; Excel's default password opens it silently", async ({
	page
}) => {
	test.slow()

	const cspViolations = trackCspViolations(page)
	const canary = trackPasswordCanary(page, WORKBOOK_PASSWORD)
	const dialog = page.getByRole("dialog").first()
	const saveButton = page.getByRole("button", { name: "Save", exact: true })
	const prompt = page.getByRole("dialog", { name: preview.previewDocxPasswordTitle })
	const badge = page.getByRole("button", { name: preview.previewDocxPasswordProtected })

	async function unlock(password: string): Promise<void> {
		const field = prompt.getByLabel(preview.previewDocxPasswordLabel, { exact: true })

		await expect(field).toBeFocused({ timeout: LAZY_VIEWER_TIMEOUT_MS })
		await field.fill(password)
		await prompt.getByRole("button", { name: preview.previewDocxPasswordSubmit, exact: true }).click()
	}

	await withScratchDirectory(page, "sheet-next-encrypted", async ({ listbox, runId }) => {
		const encrypted = `e2e-sheet-next-${runId}-encrypted.xlsx`
		const silent = `e2e-sheet-next-${runId}-default.xlsx`

		await uploadFiles(
			page,
			[
				{ name: encrypted, mimeType: XLSX_MIME, buffer: Buffer.from(await encryptedWorkbook("agile")) },
				{ name: silent, mimeType: XLSX_MIME, buffer: Buffer.from(await encryptedWorkbook("default")) }
			],
			listbox
		)

		await listbox.getByRole("option", { name: encrypted }).dblclick()
		await expect(prompt).toBeVisible({ timeout: LAZY_VIEWER_TIMEOUT_MS })
		await unlock("wrong")
		await expect(prompt.getByText(preview.previewDocxPasswordWrong)).toBeVisible({ timeout: LAZY_VIEWER_TIMEOUT_MS })
		await unlock(WORKBOOK_PASSWORD)
		await expect(prompt).toHaveCount(0, { timeout: LAZY_VIEWER_TIMEOUT_MS })
		await expectCell(dialog, "A1", "Item", LAZY_VIEWER_TIMEOUT_MS)
		await expect(badge).toBeVisible()

		await typeInto(page, dialog, "A2", "Sealed edit")
		await expect(saveButton).toBeEnabled()
		await saveButton.click()
		await expect(saveButton).toHaveCount(0, { timeout: LIVE_WRITE_TIMEOUT_MS })
		await expect(prompt).toHaveCount(0)
		await page.keyboard.press("Escape")
		await expect(dialog).toHaveCount(0)

		// What the drive now holds is still protected: asked again, it holds the edit.
		await page.reload()
		await expect(listbox.getByRole("option", { name: encrypted })).toBeVisible({ timeout: LIVE_WRITE_TIMEOUT_MS })
		await listbox.getByRole("option", { name: encrypted }).dblclick()
		await expect(prompt).toBeVisible({ timeout: LAZY_VIEWER_TIMEOUT_MS })
		await unlock(WORKBOOK_PASSWORD)
		await expectCell(dialog, "A2", "Sealed edit", LAZY_VIEWER_TIMEOUT_MS)
		await page.keyboard.press("Escape")
		await expect(dialog).toHaveCount(0)

		// Excel's default password: no prompt at all.
		await listbox.getByRole("option", { name: silent }).dblclick()
		await expectCell(dialog, "A1", "Item", LAZY_VIEWER_TIMEOUT_MS)
		await expect(badge).toBeVisible()
		await expect(prompt).toHaveCount(0)
		await page.keyboard.press("Escape")
		await expect(dialog).toHaveCount(0)

		expect(await storedPassword(page, WORKBOOK_PASSWORD)).toEqual([])
	})

	expect(canary.leaks).toEqual([])
	expect(cspViolations).toEqual([])
})
