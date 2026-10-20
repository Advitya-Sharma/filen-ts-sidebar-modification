import type {
	CellStyleView,
	FormatPatch,
	HorizontalAlign,
	SheetFormatPatch,
	SheetHorizontalAlign,
	SheetRenderStyle,
	SheetRgb
} from "@filen/office/core/sheet"

function hex(color: SheetRgb): string {
	return `#${color.toString(16).padStart(6, "0")}`
}

// "#rrggbb" as a colour, null as "no colour"; anything else changes nothing (undefined).
function colorInput(color: string | null | undefined): SheetRgb | null | undefined {
	if (color === null || color === undefined) {
		return color
	}

	return /^#[0-9a-f]{6}$/i.test(color) ? Number.parseInt(color.slice(1), 16) : undefined
}

// The alignments the toolbar offers; any other shows none of its buttons pressed.
function toolbarAlign(align: SheetHorizontalAlign | undefined): HorizontalAlign | undefined {
	return align === "left" || align === "center" || align === "right" ? align : undefined
}

// The active cell's look as the formatting toolbar reads it.
export function toolbarStyle(style: SheetRenderStyle, numFmt: string): CellStyleView {
	const align = toolbarAlign(style.align)

	return {
		...(style.bold === true && { bold: true }),
		...(style.italic === true && { italic: true }),
		...(style.underline !== undefined && { underline: true }),
		...(style.strike === true && { strike: true }),
		...(style.color !== undefined && { color: hex(style.color) }),
		...(style.fill !== undefined && { fill: hex(style.fill) }),
		...(align !== undefined && { align }),
		...(style.wrap === true && { wrap: true }),
		...(style.size !== undefined && { size: style.size }),
		numFmt
	}
}

// The toolbar's change as the session's format command takes it.
export function sheetFormatPatch(patch: FormatPatch): SheetFormatPatch {
	const color = colorInput(patch.color)
	const fill = colorInput(patch.fill)

	return {
		...(patch.bold !== undefined && { bold: patch.bold }),
		...(patch.italic !== undefined && { italic: patch.italic }),
		...(patch.underline !== undefined && { underline: patch.underline ? "single" : null }),
		...(patch.strike !== undefined && { strike: patch.strike }),
		...(color !== undefined && { color }),
		...(fill !== undefined && { fill }),
		...(patch.align !== undefined && { align: patch.align }),
		...(patch.numFmt !== undefined && { numFmt: patch.numFmt })
	}
}
