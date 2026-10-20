import type { XlsxColor, XlsxTheme } from "../../xlsx/model/types"
import { SYSTEM_FOREGROUND } from "./palette"

// Colours of SpreadsheetML styles (ECMA-376 Part 1 §18.8.3 CT_Color) as 0xRRGGBB.

// The theme index of a colour (0-based, §18.8.3 theme) into the theme's clrScheme order (dk1, lt1, dk2, lt2,
// accent1-6, hlink, folHlink, §20.1.6.2): SpreadsheetML counts the first two pairs light first, so theme 0 is lt1
// (Background 1) and theme 1 dk1 (Text 1).
const THEME_ORDER = [1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11]

// The Office 2007-2013 default theme's colours, for a workbook without a theme part (clrScheme order).
export const DEFAULT_THEME_COLORS: readonly number[] = [
	0x000000, 0xffffff, 0x1f497d, 0xeeece1, 0x4f81bd, 0xc0504d, 0x9bbb59, 0x8064a2, 0x4bacc6, 0xf79646, 0x0000ff, 0x800080
]

// The tint of §18.8.3: the colour's luminance in HLS (0..HLSMAX, HLSMAX 255) darkened by a negative tint
// (L * (1 + tint)) or lightened by a positive one (L * (1 - tint) + HLSMAX - HLSMAX * (1 - tint)).
export function applyTint(rgb: number, tint: number): number {
	if (tint === 0 || !Number.isFinite(tint)) return rgb

	const r = ((rgb >> 16) & 0xff) / 255
	const g = ((rgb >> 8) & 0xff) / 255
	const b = (rgb & 0xff) / 255
	const max = Math.max(r, g, b)
	const min = Math.min(r, g, b)
	let hue = 0
	let saturation = 0
	let lum = (max + min) / 2

	if (max !== min) {
		const delta = max - min

		saturation = lum > 0.5 ? delta / (2 - max - min) : delta / (max + min)
		hue = max === r ? (g - b) / delta + (g < b ? 6 : 0) : max === g ? (b - r) / delta + 2 : (r - g) / delta + 4
		hue /= 6
	}

	const clamped = Math.max(-1, Math.min(1, tint))

	lum = clamped < 0 ? lum * (1 + clamped) : lum * (1 - clamped) + clamped

	if (saturation === 0) {
		const grey = Math.round(lum * 255)

		return (grey << 16) | (grey << 8) | grey
	}

	const q = lum < 0.5 ? lum * (1 + saturation) : lum + saturation - lum * saturation
	const p = 2 * lum - q
	const channel = (t: number): number => {
		const x = t < 0 ? t + 1 : t > 1 ? t - 1 : t
		const value = x < 1 / 6 ? p + (q - p) * 6 * x : x < 1 / 2 ? q : x < 2 / 3 ? p + (q - p) * (2 / 3 - x) * 6 : p

		return Math.max(0, Math.min(255, Math.round(value * 255)))
	}

	return (channel(hue + 1 / 3) << 16) | (channel(hue) << 8) | channel(hue - 1 / 3)
}

// Resolves a colour: theme over rgb over indexed over auto (MS-OI29500 §2.1.694: the order Office reads them),
// then its tint. `auto` is what an automatic colour is in its place (a font's: the system foreground). -1 when the
// colour is absent or names nothing this workbook has.
export function resolveColor(
	color: XlsxColor | undefined,
	theme: XlsxTheme | null,
	palette: readonly number[],
	auto: number = palette[SYSTEM_FOREGROUND] ?? 0
): number {
	if (color === undefined) return -1

	let rgb = -1

	if (color.theme !== undefined) {
		const index = THEME_ORDER[color.theme]
		const colors = theme?.colors ?? DEFAULT_THEME_COLORS

		rgb = index === undefined ? -1 : (colors[index] ?? -1)
	} else if (color.argb !== undefined) rgb = color.argb & 0xffffff
	else if (color.indexed !== undefined) rgb = palette[color.indexed] ?? -1
	else if (color.auto === true) rgb = auto

	if (rgb < 0) return -1

	return color.tint === undefined ? rgb : applyTint(rgb, color.tint)
}
