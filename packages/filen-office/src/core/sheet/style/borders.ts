import type { SheetBorderStyle, SheetColorInput, SheetId, SheetRange } from "../protocol"

// Borders as a command sets them: lines for the edges of each range (its outline, the lines inside it, the
// diagonals), and the presets of Excel's Borders menu (Home tab, "Apply or remove cell borders on a worksheet",
// support.microsoft.com) as edge sets.

// A line to set; an absent or null colour is automatic.
export interface SheetBorderLineInput {
	readonly style: SheetBorderStyle
	readonly color?: SheetColorInput | null
}

// What a border edit does to each edge of a range: a line sets it, null clears it, absent leaves it as it is.
// Inside lines go on both cells that share them, as Excel and LibreOffice write them.
export interface SheetBorderEdges {
	readonly top?: SheetBorderLineInput | null
	readonly bottom?: SheetBorderLineInput | null
	readonly left?: SheetBorderLineInput | null
	readonly right?: SheetBorderLineInput | null
	readonly insideHorizontal?: SheetBorderLineInput | null
	readonly insideVertical?: SheetBorderLineInput | null
	readonly diagonalUp?: SheetBorderLineInput | null
	readonly diagonalDown?: SheetBorderLineInput | null
}

// The neighbours' facing lines of a cleared outline edge go too (LibreOffice's ScDocument::ApplySelectionFrame
// with RemoveAdjacentCellBorder, as its Borders menu and dialog set it; Microsoft documents nothing else), so a
// cleared edge shows no line.
export interface SheetBorderPatch {
	readonly sheet: SheetId
	readonly ranges: readonly SheetRange[]
	readonly edges: SheetBorderEdges
}

declare module "../protocol" {
	interface SheetCommandMap {
		"format.borders": SheetBorderPatch
		// Excel's Clear Formats: every cell of the ranges (whole columns and rows their axes too) back to the
		// workbook's default format, contents kept.
		"format.clear": { readonly sheet: SheetId; readonly ranges: readonly SheetRange[] }
	}
}

export type SheetBorderPreset =
	| "bottom"
	| "top"
	| "left"
	| "right"
	| "none"
	| "all"
	| "outside"
	| "thickOutside"
	| "inside"
	| "insideHorizontal"
	| "insideVertical"
	| "bottomDouble"
	| "thickBottom"
	| "topBottom"
	| "topThickBottom"
	| "topDoubleBottom"
	| "diagonalUp"
	| "diagonalDown"

// The menu's order: Excel's, then the inside lines and the diagonals its dialog offers.
export const SHEET_BORDER_PRESETS: readonly SheetBorderPreset[] = [
	"bottom",
	"top",
	"left",
	"right",
	"none",
	"all",
	"outside",
	"thickOutside",
	"bottomDouble",
	"thickBottom",
	"topBottom",
	"topThickBottom",
	"topDoubleBottom",
	"inside",
	"insideHorizontal",
	"insideVertical",
	"diagonalUp",
	"diagonalDown"
]

// The edges a preset sets with `line` (the menu's line style and colour; thick and double presets keep their own
// style in its colour). "none" clears the outline and the inside lines, not the diagonals.
export function borderPresetEdges(preset: SheetBorderPreset, line: SheetBorderLineInput = { style: "thin" }): SheetBorderEdges {
	const styled = (style: SheetBorderStyle): SheetBorderLineInput => ({ ...line, style })

	switch (preset) {
		case "bottom":
			return { bottom: line }
		case "top":
			return { top: line }
		case "left":
			return { left: line }
		case "right":
			return { right: line }
		case "none":
			return { top: null, bottom: null, left: null, right: null, insideHorizontal: null, insideVertical: null }
		case "all":
			return { top: line, bottom: line, left: line, right: line, insideHorizontal: line, insideVertical: line }
		case "outside":
			return { top: line, bottom: line, left: line, right: line }
		case "thickOutside": {
			const thick = styled("thick")

			return { top: thick, bottom: thick, left: thick, right: thick }
		}
		case "inside":
			return { insideHorizontal: line, insideVertical: line }
		case "insideHorizontal":
			return { insideHorizontal: line }
		case "insideVertical":
			return { insideVertical: line }
		case "bottomDouble":
			return { bottom: styled("double") }
		case "thickBottom":
			return { bottom: styled("thick") }
		case "topBottom":
			return { top: line, bottom: line }
		case "topThickBottom":
			return { top: line, bottom: styled("thick") }
		case "topDoubleBottom":
			return { top: line, bottom: styled("double") }
		case "diagonalUp":
			return { diagonalUp: line }
		case "diagonalDown":
			return { diagonalDown: line }
	}
}
