import { defineSheetFeature } from "../../sheet/registry"
import { DEFAULT_DIGIT_PX, MAX_DIGIT_PX, workbookDigitPx } from "../../sheet/sizes/sizes"
import type { XlsxDocument } from "./document"

// The workbook's column-width unit: the widest digit of its default font in whole pixels (ECMA-376 Part 1
// §18.3.1.13), worked out from the font's name and size alone (workbookDigitPx), so it is the same on every device.
// The app asks for it once the workbook opens, before it shows a layout; every layout made after it, and every width
// written, uses it. Widths in the file are unchanged.

declare module "../../sheet/protocol" {
	interface SheetQueryMap {
		// The default font (the Normal style's, as cell style 0 shows it) and the digit width in use: the font's own
		// unless `px` sets another (whole pixels, held to 1–MAX_DIGIT_PX).
		digitWidth: {
			request: { readonly px?: number }
			result: { readonly px: number; readonly font: SheetRenderStyle }
		}
	}
}

export { MAX_DIGIT_PX }

// A digit width as a whole number of pixels the unit can be: Calibri 11's 7 px when it is no number.
export function validDigit(px: number): number {
	return Number.isFinite(px) ? Math.max(1, Math.min(MAX_DIGIT_PX, Math.round(px))) : DEFAULT_DIGIT_PX
}

export const digitWidthFeature = defineSheetFeature<XlsxDocument>({
	id: "digitWidth",
	queries: {
		digitWidth: {
			run: (context, request) => {
				const { font, size } = context.doc.styles.renderStyle(0)

				context.doc.digitPx = validDigit(request.px ?? workbookDigitPx(font, size))

				return { px: context.doc.digitPx, font: { ...(font !== undefined && { font }), ...(size !== undefined && { size }) } }
			}
		}
	}
})
