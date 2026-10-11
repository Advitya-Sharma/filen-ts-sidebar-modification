import { applyColorTransforms, type ColorTransform } from "../../docx/theme/index"
import type { DrawingTheme } from "../color"
import { rgba, type Rgba } from "../displayList"

// Default series colours when a series has no c:spPr fill, after the built-in chart styles (c:style 1-48,
// eight colour columns per row): column 1 greys, column 2 the six accents, columns 3-8 shades of one
// accent. Accents repeat in darker and lighter variants after six series, as Office does.

const ACCENTS = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]

// Variations per pass over the six accents.
const CYCLES: readonly (readonly ColorTransform[])[] = [
	[],
	[{ kind: "lumMod", value: 0.6 }],
	[
		{ kind: "lumMod", value: 0.8 },
		{ kind: "lumOff", value: 0.2 }
	],
	[{ kind: "lumMod", value: 0.8 }],
	[
		{ kind: "lumMod", value: 0.6 },
		{ kind: "lumOff", value: 0.4 }
	],
	[{ kind: "lumMod", value: 0.5 }],
	[
		{ kind: "lumMod", value: 0.7 },
		{ kind: "lumOff", value: 0.3 }
	],
	[{ kind: "lumMod", value: 0.7 }],
	[
		{ kind: "lumMod", value: 0.5 },
		{ kind: "lumOff", value: 0.5 }
	]
]

const FALLBACK_ACCENTS = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"]

function accent(theme: DrawingTheme, slot: number): string {
	return theme.scheme(ACCENTS[slot] ?? "accent1") ?? FALLBACK_ACCENTS[slot] ?? "4472C4"
}

// The colour of series (or point, for varied colours) `index` of `count` under chart style `style`.
export function seriesColor(theme: DrawingTheme, style: number, index: number, count: number): Rgba {
	const column = (Math.max(1, Math.min(48, style)) - 1) % 8
	const position = Math.max(0, index)

	if (column === 1) {
		const cycle = CYCLES[Math.floor(position / 6) % CYCLES.length] ?? []

		return rgba(applyColorTransforms(accent(theme, position % 6), cycle))
	}

	// Monochrome styles run from dark to light across the series.
	const base = column === 0 ? (theme.scheme("tx1") ?? "000000") : accent(theme, column - 2)
	const t = count <= 1 ? 0.5 : position / (count - 1)

	if (column === 0) {
		const gray = Math.round(0x40 + (0xd0 - 0x40) * t)
		const hex = gray.toString(16).padStart(2, "0").toUpperCase()

		return rgba(hex + hex + hex)
	}

	const transforms: ColorTransform[] = t < 0.5 ? [{ kind: "shade", value: 0.5 + t }] : [{ kind: "tint", value: 1 - (t - 0.5) * 1.2 }]

	return rgba(applyColorTransforms(base, transforms))
}

// tx1 shifted toward bg1, for gridlines and axis lines.
export function mutedText(theme: DrawingTheme, keep: number): Rgba {
	return rgba(
		applyColorTransforms(theme.scheme("tx1") ?? "000000", [
			{ kind: "lumMod", value: keep },
			{ kind: "lumOff", value: 1 - keep }
		])
	)
}

export function textColor(theme: DrawingTheme): Rgba {
	return rgba(theme.scheme("tx1") ?? "000000")
}

export function backgroundColor(theme: DrawingTheme): Rgba {
	return rgba(theme.scheme("bg1") ?? "FFFFFF")
}
