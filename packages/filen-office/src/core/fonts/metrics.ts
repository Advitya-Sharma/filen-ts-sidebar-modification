// A face's vertical metrics in font units, as stored: descenders are negative in hhea and OS/2 typo,
// positive in OS/2 win.
export interface VerticalMetrics {
	readonly unitsPerEm: number
	readonly hheaAscender: number
	readonly hheaDescender: number
	readonly hheaLineGap: number
	readonly typoAscender: number
	readonly typoDescender: number
	readonly typoLineGap: number
	readonly winAscent: number
	readonly winDescent: number
	// OS/2 fsSelection bit 7 (USE_TYPO_METRICS).
	readonly useTypoMetrics: boolean
}

// Positional constructor for generated tables.
export function verticalMetrics(
	unitsPerEm: number,
	hheaAscender: number,
	hheaDescender: number,
	hheaLineGap: number,
	typoAscender: number,
	typoDescender: number,
	typoLineGap: number,
	winAscent: number,
	winDescent: number,
	useTypoMetrics: boolean
): VerticalMetrics {
	return {
		unitsPerEm,
		hheaAscender,
		hheaDescender,
		hheaLineGap,
		typoAscender,
		typoDescender,
		typoLineGap,
		winAscent,
		winDescent,
		useTypoMetrics
	}
}

export type LineSpacingRule = "gdi" | "typoWhenFlagged"

// Default line spacing in font units. "gdi" is TEXTMETRIC tmHeight + tmExternalLeading as the OpenType
// spec maps them from the win and hhea values; "typoWhenFlagged" switches to the typo values when the font
// sets USE_TYPO_METRICS, as the spec recommends for new layout code. Which one Word follows is not
// documented, so the layout chooses.
export function lineSpacingUnits(metrics: VerticalMetrics, rule: LineSpacingRule): number {
	if (rule === "typoWhenFlagged" && metrics.useTypoMetrics) {
		return metrics.typoAscender - metrics.typoDescender + metrics.typoLineGap
	}

	const win = metrics.winAscent + metrics.winDescent
	const hhea = metrics.hheaAscender - metrics.hheaDescender + metrics.hheaLineGap

	return win + Math.max(0, hhea - win)
}
