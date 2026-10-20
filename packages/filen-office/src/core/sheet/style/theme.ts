import type { SheetRgb } from "../protocol"

// What a formatting toolbar offers from the workbook itself: its theme's colours (the colour pickers' theme
// part) and fonts (the font picker's theme fonts), and the font every unformatted cell shows.
export interface SheetFormatChoices {
	// Theme colours by SpreadsheetML theme index (§18.8.3: 0 Background 1, 1 Text 1, 2 Background 2, 3 Text 2,
	// 4-9 Accent 1-6, 10 Hyperlink, 11 Followed Hyperlink), with the Office default theme's where the file has none.
	readonly themeColors: readonly SheetRgb[]
	// The theme's heading and body fonts ("" when the file has no theme).
	readonly majorFont: string
	readonly minorFont: string
	// Families the workbook's fonts name, each once, in the order the styles part holds them.
	readonly fonts: readonly string[]
}

declare module "../protocol" {
	interface SheetQueryMap {
		"format.choices": { request: Record<string, never>; result: SheetFormatChoices }
	}
}

// The tints of the theme colour picker's variations (Excel's: lighter 80/60/40 %, darker 25/50 %; Text 1 and
// Background 1, black and white in the default theme, get steps of their own), by theme index.
export function themeTints(index: number): readonly number[] {
	if (index === 0) return [-0.05, -0.15, -0.25, -0.35, -0.5]
	if (index === 1) return [0.5, 0.35, 0.25, 0.15, 0.05]
	if (index === 2) return [-0.1, -0.25, -0.5, -0.75, -0.9]

	return [0.8, 0.6, 0.4, -0.25, -0.5]
}
