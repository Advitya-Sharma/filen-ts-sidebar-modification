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

// How a face's default line is measured. "gdi" is TEXTMETRIC tmHeight + tmExternalLeading as the OpenType spec maps
// them from the win and hhea values, with the leading above the ascent; "typoWhenFlagged" switches to the typo values
// when the font sets USE_TYPO_METRICS, as the spec recommends for new layout code. Which one Word follows is not
// documented, so the layout chooses. "hheaGapBelow" is LibreOffice's: the hhea ascent and descent (typo when
// USE_TYPO_METRICS is set, win when hhea has none) with the line gap (external leading) under the descent, so a taller
// font's descent on the same line swallows it; "hhea" leaves the line gap out, as LibreOffice does when a document
// turns external leading off (measured, LibreOffice 7.5).
export type LineSpacingRule = "gdi" | "typoWhenFlagged" | "hheaGapBelow" | "hhea"

// Ascent and descent of a face's default line in font units, adding up to its line spacing.
export function lineAscentDescent(metrics: VerticalMetrics, rule: LineSpacingRule): { ascent: number; descent: number } {
	const typo = metrics.useTypoMetrics && (metrics.typoAscender !== 0 || metrics.typoDescender !== 0)

	if (rule === "hheaGapBelow" || rule === "hhea") {
		const [ascent, descent, gap] = typo
			? [metrics.typoAscender, -metrics.typoDescender, metrics.typoLineGap]
			: metrics.hheaAscender !== 0 || metrics.hheaDescender !== 0
				? [metrics.hheaAscender, -metrics.hheaDescender, metrics.hheaLineGap]
				: [metrics.winAscent, metrics.winDescent, 0]

		return { ascent, descent: descent + (rule === "hheaGapBelow" ? Math.max(0, gap) : 0) }
	}

	if (rule === "typoWhenFlagged" && metrics.useTypoMetrics) {
		return { ascent: metrics.typoAscender + metrics.typoLineGap, descent: -metrics.typoDescender }
	}

	const win = metrics.winAscent + metrics.winDescent
	const hhea = metrics.hheaAscender - metrics.hheaDescender + metrics.hheaLineGap

	return { ascent: metrics.winAscent + Math.max(0, hhea - win), descent: metrics.winDescent }
}

// A missing symbol font (Symbol, Wingdings and the like) as LibreOffice measures it: OpenSymbol, the font it ships for
// them (LibreOffice 7.5's opens___.ttf: head unitsPerEm, hhea, OS/2 typo and win values).
export const SYMBOL_FALLBACK_METRICS: VerticalMetrics = verticalMetrics(2048, 1420, -442, 307, 1420, -442, 307, 1897, 641, true)

// Default line spacing in font units.
export function lineSpacingUnits(metrics: VerticalMetrics, rule: LineSpacingRule): number {
	const { ascent, descent } = lineAscentDescent(metrics, rule)

	return ascent + descent
}
