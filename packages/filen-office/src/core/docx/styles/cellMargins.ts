import type { CellMarginsProps, ResolvedCellMargins, WidthProps } from "../props/types"

// A table's cell margins from its layers (style hierarchy first, then the table's own, a row's exceptions, a
// cell's), each side on its own: the last layer that sets a side wins. A side no layer sets takes the standard's
// default: 115 twips (0.08 in) leading and trailing, no top or bottom padding (ISO/IEC 29500-1 §17.4, the start,
// end, top and bottom children of w:tblCellMar; Transitional left and right); MS-OI29500 lists no other default
// for Word. Layout and the editing commands both read margins through here.

const NO_PADDING: WidthProps = { value: 0, type: "dxa" }
const SIDE_PADDING: WidthProps = { value: 115, type: "dxa" }

export function resolveCellMargins(layers: readonly (CellMarginsProps | undefined)[]): ResolvedCellMargins {
	let top: WidthProps | undefined
	let start: WidthProps | undefined
	let bottom: WidthProps | undefined
	let end: WidthProps | undefined

	for (const layer of layers) {
		if (layer === undefined) continue

		top = layer.top ?? top
		start = layer.start ?? start
		bottom = layer.bottom ?? bottom
		end = layer.end ?? end
	}

	return { top: top ?? NO_PADDING, start: start ?? SIDE_PADDING, bottom: bottom ?? NO_PADDING, end: end ?? SIDE_PADDING }
}

// A margin in twips: a width in another unit counts as none (Word takes pct and auto as 0, MS-OI29500 2.1.146).
export function marginTwips(width: WidthProps): number {
	return width.type === "dxa" ? width.value : 0
}
