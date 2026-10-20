// The default indexed palette (ECMA-376 Part 1 §18.8.27), 0xRRGGBB. Index 64 is the system foreground (window
// text) and 65 the system background (window); they are drawn black and white.
export const DEFAULT_PALETTE: readonly number[] = [
	0x000000, 0xffffff, 0xff0000, 0x00ff00, 0x0000ff, 0xffff00, 0xff00ff, 0x00ffff, 0x000000, 0xffffff, 0xff0000, 0x00ff00, 0x0000ff,
	0xffff00, 0xff00ff, 0x00ffff, 0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xc0c0c0, 0x808080, 0x9999ff, 0x993366,
	0xffffcc, 0xccffff, 0x660066, 0xff8080, 0x0066cc, 0xccccff, 0x000080, 0xff00ff, 0xffff00, 0x00ffff, 0x800080, 0x800000, 0x008080,
	0x0000ff, 0x00ccff, 0xccffff, 0xccffcc, 0xffff99, 0x99ccff, 0xff99cc, 0xcc99ff, 0xffcc99, 0x3366ff, 0x33cccc, 0x99cc00, 0xffcc00,
	0xff9900, 0xff6600, 0x666699, 0x969696, 0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333, 0x000000,
	0xffffff
]

export const SYSTEM_FOREGROUND = 64
export const SYSTEM_BACKGROUND = 65

// A workbook's palette: its indexedColors (0xAARRGGBB) over the default, the system colours kept.
export function workbookPalette(indexedColors: readonly number[] | null): readonly number[] {
	if (indexedColors === null || indexedColors.length === 0) return DEFAULT_PALETTE

	const palette = DEFAULT_PALETTE.slice()

	for (let i = 0; i < indexedColors.length && i < SYSTEM_FOREGROUND; i++) palette[i] = (indexedColors[i] as number) & 0xffffff

	return palette
}
