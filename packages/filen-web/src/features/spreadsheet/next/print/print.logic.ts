import type { SheetRange, SheetSummary } from "@filen/office/core/sheet"
import { COMMON_PAPER_SIZES, paperPoints, type SheetPageSetup, type SheetPrintPatch } from "@filen/office/core/sheet/print"
import type { PrintSheetRequest } from "@filen/office/react/sheet/print"
import type { PreviewKey } from "@/lib/i18n"

// The print dialog's choices and what they print: the sheets and ranges, each sheet's setup with the dialog's choices
// over the active sheet's, and the patch that writes them into the file.

export type PrintScope = "sheet" | "workbook" | "selection"

export type PrintScaling = "none" | "fitSheet" | "fitColumns" | "fitRows" | "fitPages" | "custom"

export interface PrintForm {
	readonly scope: PrintScope
	readonly ignorePrintArea: boolean
	readonly paperSize: number
	readonly orientation: "portrait" | "landscape"
	readonly scaling: PrintScaling
	// The custom percent as typed.
	readonly percent: string
	readonly gridLines: boolean
	readonly headings: boolean
}

// What the browser lays out for one print at most (each page an SVG in the document).
export const MAX_BROWSER_PAGES = 500

export const PAPER_KEYS: ReadonlyMap<number, PreviewKey> = new Map<number, PreviewKey>([
	[1, "previewSheetPrintPaperLetter"],
	[5, "previewSheetPrintPaperLegal"],
	[3, "previewSheetPrintPaperTabloid"],
	[7, "previewSheetPrintPaperExecutive"],
	[9, "previewSheetPrintPaperA4"],
	[8, "previewSheetPrintPaperA3"],
	[11, "previewSheetPrintPaperA5"],
	[13, "previewSheetPrintPaperB5"]
])

export const COMMON_PAPERS: readonly number[] = COMMON_PAPER_SIZES

export function scalingOf(setup: SheetPageSetup): PrintScaling {
	if (!setup.fitToPage) return setup.scale === 100 ? "none" : "custom"
	if (setup.fitToWidth === 1 && setup.fitToHeight === 1) return "fitSheet"
	if (setup.fitToWidth === 1 && setup.fitToHeight === 0) return "fitColumns"
	if (setup.fitToWidth === 0 && setup.fitToHeight === 1) return "fitRows"

	return "fitPages"
}

export function formOf(setup: SheetPageSetup, scope: PrintScope = "sheet"): PrintForm {
	return {
		scope,
		ignorePrintArea: false,
		paperSize: setup.customPaper ? 0 : setup.paperSize,
		orientation: setup.orientation,
		scaling: scalingOf(setup),
		percent: String(setup.scale),
		gridLines: setup.gridLines,
		headings: setup.headings
	}
}

// A whole number from 10 to 400, or null.
export function percentOf(text: string): number | null {
	const trimmed = text.trim()

	if (!/^[0-9]{1,3}$/.test(trimmed)) return null

	const value = Number(trimmed)

	return value >= 10 && value <= 400 ? value : null
}

// The dialog's choices over a sheet's setup (the paper of a sheet with custom paper stays its own unless one was
// chosen; "fitPages" keeps the file's page counts).
export function applyForm(setup: SheetPageSetup, form: PrintForm): SheetPageSetup | null {
	const percent = form.scaling === "custom" ? percentOf(form.percent) : 100

	if (percent === null) return null

	const fit =
		form.scaling === "fitSheet"
			? { width: 1, height: 1 }
			: form.scaling === "fitColumns"
				? { width: 1, height: 0 }
				: form.scaling === "fitRows"
					? { width: 0, height: 1 }
					: form.scaling === "fitPages"
						? { width: setup.fitToWidth, height: setup.fitToHeight }
						: null

	return {
		...setup,
		...(form.paperSize !== 0 &&
			(form.paperSize !== setup.paperSize || setup.customPaper) && {
				paperSize: form.paperSize,
				paper: paperPoints(form.paperSize),
				customPaper: false
			}),
		orientation: form.orientation,
		fitToPage: fit !== null,
		fitToWidth: fit?.width ?? setup.fitToWidth,
		fitToHeight: fit?.height ?? setup.fitToHeight,
		scale: fit === null ? percent : setup.scale,
		gridLines: form.gridLines,
		headings: form.headings
	}
}

// What writes the choices into the file: only what differs from the sheet's setup (null: nothing).
export function patchOf(setup: SheetPageSetup, form: PrintForm): SheetPrintPatch | null {
	const next = applyForm(setup, form)

	if (next === null) return null

	const patch: {
		-readonly [K in keyof SheetPrintPatch]: SheetPrintPatch[K]
	} = {}

	if (next.orientation !== setup.orientation) patch.orientation = next.orientation
	if (form.paperSize !== 0 && (setup.customPaper || form.paperSize !== setup.paperSize)) patch.paperSize = form.paperSize
	if (next.fitToPage && (!setup.fitToPage || next.fitToWidth !== setup.fitToWidth || next.fitToHeight !== setup.fitToHeight))
		patch.fit = { width: next.fitToWidth, height: next.fitToHeight }
	if (!next.fitToPage && (setup.fitToPage || next.scale !== setup.scale)) patch.scale = next.scale
	if (next.gridLines !== setup.gridLines) patch.gridLines = next.gridLines
	if (next.headings !== setup.headings) patch.headings = next.headings

	return Object.keys(patch).length === 0 ? null : patch
}

export interface PrintSheetSettings {
	readonly sheet: SheetSummary
	readonly setup: SheetPageSetup
	readonly used: SheetRange | null
}

// The sheets and ranges a print takes: the active sheet (its print area unless ignored, else its used range), every
// visible worksheet, or the selection's ranges. The dialog's choices apply to the active sheet; the other sheets of a
// workbook print with their own settings, as in Excel.
export function printRequests(
	form: PrintForm,
	active: PrintSheetSettings,
	// Every worksheet, in the workbook's order (the active one included).
	sheets: readonly PrintSheetSettings[],
	selection: readonly SheetRange[]
): PrintSheetRequest[] | null {
	const own = applyForm(active.setup, form)

	if (own === null) return null

	const request = (settings: PrintSheetSettings, setup: SheetPageSetup): PrintSheetRequest => ({
		sheet: settings.sheet.id,
		name: settings.sheet.name,
		setup,
		used: settings.used,
		ranges: setup.printArea !== null && !form.ignorePrintArea ? setup.printArea : settings.used === null ? [] : [settings.used]
	})

	if (form.scope === "selection") {
		return [
			{ sheet: active.sheet.id, name: active.sheet.name, setup: { ...own, printArea: null }, used: active.used, ranges: selection }
		]
	}

	if (form.scope === "sheet") return [request(active, own)]

	return sheets
		.filter(settings => settings.sheet.state === "visible")
		.map(settings => (settings.sheet.id === active.sheet.id ? request(settings, own) : request(settings, settings.setup)))
}
