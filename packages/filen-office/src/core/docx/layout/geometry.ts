import {
	Alignment,
	LineRule,
	TabAlignment,
	TabLeader,
	TabStops,
	type Exclusions,
	type ParagraphGeometry,
	type TabStop
} from "../../text/index"
import type { DocxSettings } from "../model/settings"
import type { ResolvedParagraphProps, ResolvedSectionProps } from "../props/types"
import { points, twips } from "./units"

// Paragraph and page geometry in px from resolved properties.

// A reusable ParagraphGeometry.
export class Geometry implements ParagraphGeometry {
	left = 0
	width = 0
	startIndent = 0
	endIndent = 0
	firstLineIndent = 0
	rtl = false
	alignment: Alignment = Alignment.START
	tabs: TabStops | null = null
	defaultTabStop = 48
	lineRule: LineRule = LineRule.AUTO
	lineSpacing = 1
	expandLineBreaks = true
	hangingIndentTab = true
	exclusions: Exclusions | null = null
	decimalSeparator = 0x2e
}

export interface LayoutCompat {
	readonly defaultTabStop: number
	readonly expandLineBreaks: boolean
	readonly hangingIndentTab: boolean
	readonly decimalSeparator: number
	readonly compatibilityMode: number
	// w:doNotUseHTMLParagraphAutoSpacing is not set.
	readonly htmlSpacing: boolean
	// w:suppressSpBfAfterPgBrk: no space before at the top of a page after a page or column break.
	readonly suppressSpaceAfterBreak: boolean
	// w:suppressTopSpacing: the first line of a page ignores exact and at-least line heights.
	readonly suppressTopSpacing: boolean
	readonly mirrorMargins: boolean
	readonly gutterAtTop: boolean
	readonly evenAndOddHeaders: boolean
}

export function layoutCompat(settings: DocxSettings): LayoutCompat {
	const legacy = new Set(settings.legacyCompat)

	return {
		defaultTabStop: twips(settings.defaultTabStop ?? 720) || 48,
		expandLineBreaks: !legacy.has("doNotExpandShiftReturn"),
		hangingIndentTab: !legacy.has("doNotUseIndentAsNumberingTabStop"),
		decimalSeparator: (settings.decimalSymbol ?? ".").charCodeAt(0) || 0x2e,
		compatibilityMode: settings.compatibilityMode ?? 11,
		htmlSpacing: !legacy.has("doNotUseHTMLParagraphAutoSpacing"),
		suppressSpaceAfterBreak: legacy.has("suppressSpBfAfterPgBrk"),
		suppressTopSpacing: legacy.has("suppressTopSpacing"),
		mirrorMargins: settings.mirrorMargins,
		gutterAtTop: settings.gutterAtTop,
		evenAndOddHeaders: settings.evenAndOddHeaders
	}
}

const TAB_ALIGNMENTS: Readonly<Record<string, TabAlignment>> = {
	start: TabAlignment.START,
	center: TabAlignment.CENTER,
	end: TabAlignment.END,
	decimal: TabAlignment.DECIMAL,
	bar: TabAlignment.BAR,
	num: TabAlignment.START
}

const TAB_LEADERS: Readonly<Record<string, TabLeader>> = {
	dot: TabLeader.DOT,
	hyphen: TabLeader.HYPHEN,
	underscore: TabLeader.UNDERSCORE,
	heavy: TabLeader.HEAVY,
	middleDot: TabLeader.MIDDLE_DOT
}

const tabCache = new WeakMap<ResolvedParagraphProps, { readonly end: number; readonly stops: TabStops | null }>()

// A paragraph's tab stops for a line ending at `end` (px from the column's start). A right, centre or
// decimal stop past the end aligns there: Word keeps the text after it on the line, running into the margin
// rather than wrapping (tables of contents set their page number stop a hair past the margin).
function tabStops(props: ResolvedParagraphProps, end: number): TabStops | null {
	const cached = tabCache.get(props)

	if (cached !== undefined && cached.end === end) return cached.stops

	const stops =
		props.tabs.length === 0
			? null
			: new TabStops(
					props.tabs.map((tab): TabStop => {
						const alignment = TAB_ALIGNMENTS[tab.alignment] ?? TabAlignment.START
						const position = twips(tab.position)

						return {
							position: alignment !== TabAlignment.START && alignment !== TabAlignment.BAR && position > end ? end : position,
							alignment,
							leader: TAB_LEADERS[tab.leader] ?? TabLeader.NONE
						}
					})
				)

	tabCache.set(props, { end, stops })

	return stops
}

function alignmentOf(props: ResolvedParagraphProps): Alignment {
	switch (props.alignment) {
		case "center":
			return Alignment.CENTER
		case "end":
		case "right":
			return Alignment.END
		case "both":
		case "mediumKashida":
		case "highKashida":
		case "lowKashida":
		case "thaiDistribute":
			return Alignment.BOTH
		case "distribute":
			return Alignment.DISTRIBUTE
		default:
			return Alignment.START
	}
}

// Fills `g` for a paragraph in a column [left, left + width). `gridPitch` > 0 snaps lines to the
// document grid (approximated as a minimum height).
export function paragraphGeometry(
	g: Geometry,
	props: ResolvedParagraphProps,
	left: number,
	width: number,
	compat: LayoutCompat,
	gridPitch: number
): Geometry {
	const indent = props.indent
	const spacing = props.spacing

	g.left = left
	g.width = width
	g.startIndent = twips(indent.start)
	g.endIndent = twips(indent.end)
	g.firstLineIndent = twips(indent.firstLine)
	g.rtl = props.bidi
	g.alignment = alignmentOf(props)
	g.tabs = tabStops(props, width - twips(indent.end))
	g.defaultTabStop = compat.defaultTabStop
	g.expandLineBreaks = compat.expandLineBreaks
	g.hangingIndentTab = compat.hangingIndentTab
	g.decimalSeparator = compat.decimalSeparator
	g.exclusions = null

	if (spacing.lineRule === "exact") {
		g.lineRule = LineRule.EXACT
		g.lineSpacing = twips(Math.abs(spacing.line))
	} else if (spacing.lineRule === "atLeast") {
		g.lineRule = LineRule.AT_LEAST
		g.lineSpacing = twips(spacing.line)
	} else {
		const multiple = spacing.line > 0 ? spacing.line / 240 : 1

		if (gridPitch > 0 && props.snapToGrid) {
			g.lineRule = LineRule.AT_LEAST
			g.lineSpacing = gridPitch * multiple
		} else {
			g.lineRule = LineRule.AUTO
			g.lineSpacing = multiple
		}
	}

	return g
}

// Auto spacing (w:beforeAutospacing / w:afterAutospacing), before and after in px: under
// w:doNotUseHTMLParagraphAutoSpacing a fixed 5 pt before and 10 pt after (ECMA-376 Part 4); otherwise
// HTML-style values the layout tuning gives.
export interface AutoSpacing {
	readonly before: number
	readonly after: number
}

export const FIXED_AUTO_SPACING: AutoSpacing = { before: points(5), after: points(10) }

// w:beforeLines / w:afterLines count hundredths of a 12 pt line.
const LINE_UNIT = points(12)

export function spaceBefore(props: ResolvedParagraphProps, auto: AutoSpacing): number {
	const spacing = props.spacing

	if (spacing.beforeAutospacing) return auto.before
	if (spacing.beforeLines !== undefined && spacing.beforeLines > 0) return (spacing.beforeLines / 100) * LINE_UNIT

	return twips(Math.max(0, spacing.before))
}

export function spaceAfter(props: ResolvedParagraphProps, auto: AutoSpacing): number {
	const spacing = props.spacing

	if (spacing.afterAutospacing) return auto.after
	if (spacing.afterLines !== undefined && spacing.afterLines > 0) return (spacing.afterLines / 100) * LINE_UNIT

	return twips(Math.max(0, spacing.after))
}

// A page of a section in px.
export interface PageGeometry {
	readonly width: number
	readonly height: number
	readonly left: number
	readonly right: number
	// Text area top and bottom edges before headers and footers push them.
	readonly top: number
	readonly bottom: number
	// Negative margins fix the edge however tall the header or footer is.
	readonly topFixed: boolean
	readonly bottomFixed: boolean
	readonly header: number
	readonly footer: number
	// Columns: offsets from `left` and widths.
	readonly columnX: readonly number[]
	readonly columnWidths: readonly number[]
	readonly separator: boolean
	readonly gridPitch: number
}

// `odd`: the page number is odd (mirrored margins swap the inside and outside on even pages).
export function pageGeometry(section: ResolvedSectionProps, compat: LayoutCompat, odd: boolean): PageGeometry {
	const width = twips(section.pageWidth)
	const height = twips(section.pageHeight)
	const margins = section.margins
	const gutter = twips(Math.max(0, margins.gutter))
	let left = twips(margins.left)
	let right = twips(margins.right)
	let top = twips(Math.abs(margins.top))
	const bottomMargin = twips(Math.abs(margins.bottom))

	if (compat.mirrorMargins && !odd) [left, right] = [right, left]

	if (compat.gutterAtTop) top += gutter
	else if (section.rtlGutter) right += gutter
	else if (compat.mirrorMargins && !odd) right += gutter
	else left += gutter

	const textWidth = Math.max(1, width - left - right)
	const columns = section.columns
	const count = Math.max(1, Math.min(columns.count, 45))
	const columnX: number[] = []
	const columnWidths: number[] = []

	if (count > 1 && !columns.equalWidth && columns.columns !== undefined && columns.columns.length > 0) {
		let x = 0

		for (const column of columns.columns) {
			const w = twips(column.width ?? 0)

			columnX.push(x)
			columnWidths.push(Math.max(1, w))
			x += w + twips(column.space ?? 0)
		}
	} else {
		const space = count > 1 ? twips(columns.space) : 0
		const w = Math.max(1, (textWidth - space * (count - 1)) / count)

		for (let c = 0; c < count; c++) {
			columnX.push(c * (w + space))
			columnWidths.push(w)
		}
	}

	if (section.bidi) {
		for (let c = 0; c < columnX.length; c++) columnX[c] = textWidth - (columnX[c] ?? 0) - (columnWidths[c] ?? 0)
	}

	const grid = section.docGrid

	return {
		width,
		height,
		left,
		right,
		top,
		bottom: height - bottomMargin,
		topFixed: margins.top < 0,
		bottomFixed: margins.bottom < 0,
		header: twips(margins.header),
		footer: twips(margins.footer),
		columnX,
		columnWidths,
		separator: columns.separator && count > 1,
		gridPitch:
			(grid.type === "lines" || grid.type === "linesAndChars") && grid.linePitch !== undefined && grid.linePitch > 0
				? twips(grid.linePitch)
				: 0
	}
}
