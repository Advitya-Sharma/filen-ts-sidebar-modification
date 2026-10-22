import type { DocxTextColor } from "@filen/office/core/docx/edit"
import type { PreviewKey } from "@/lib/i18n"

// What the office editors' controls share (the DOCX editor and the spreadsheet toolbar), in a module that reaches no
// editor code.

// Keeps focus where it was (the document): a pressed toolbar or chip button never takes it.
export function keepFocus(event: { preventDefault: () => void }): void {
	event.preventDefault()
}

// What a colour swatch is called: a colour's name, with how much lighter or darker a theme variation is.
export interface DocxSwatchLabel {
	readonly name: PreviewKey
	readonly tone?: { readonly lighter: boolean; readonly percent: number }
}

export interface DocxSwatch {
	// The colour as the snapshot's values name it ("RRGGBB", "theme:accent1/tint:153"): the chosen one is marked.
	readonly key: string
	// What it shows, "RRGGBB".
	readonly rgb: string
	readonly label: DocxSwatchLabel
	readonly color: DocxTextColor
}

// Word's standard colours (Excel offers the same).
export const STANDARD_SWATCHES: readonly DocxSwatch[] = (
	[
		["C00000", "previewOfficeColorDarkRed"],
		["FF0000", "previewOfficeColorRed"],
		["FFC000", "previewOfficeColorOrange"],
		["FFFF00", "previewOfficeColorYellow"],
		["92D050", "previewOfficeColorLightGreen"],
		["00B050", "previewOfficeColorGreen"],
		["00B0F0", "previewOfficeColorLightBlue"],
		["0070C0", "previewOfficeColorBlue"],
		["002060", "previewOfficeColorDarkBlue"],
		["7030A0", "previewOfficeColorPurple"]
	] as const
).map(([rgb, name]) => ({ key: rgb, rgb, label: { name }, color: { kind: "rgb", value: rgb } }))
