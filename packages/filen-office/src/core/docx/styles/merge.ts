// Portions derived from SuperDoc 1.46.3 (packages/layout-engine/style-engine/src/cascade.ts: font slot/theme
// pairing, firstLine/hanging exclusivity, tab clearing), copyright Harbour and the SuperDoc contributors,
// licensed under the GNU Affero General Public License v3.0. Modified for @filen/office.

import type {
	CellBordersProps,
	CellMarginsProps,
	FontsProps,
	IndentProps,
	LanguageProps,
	ParagraphBordersProps,
	ParagraphProps,
	RunProps,
	SpacingProps,
	TabStop,
	TableBordersProps,
	TableCellProps,
	TableProps,
	TableRowProps
} from "../props/types"

// Cascade steps: `source` overrides `target` field by field. Compound properties merge per attribute where
// Word does (spacing, indents, fonts, language, borders by side); the rest replace whole.

export type Mutable<T> = { -readonly [K in keyof T]: T[K] }

// §17.7.3: in styles these flip the inherited state; in direct formatting they set it. Word measures a
// style level against the document defaults (MS-OI29500 §2.1.230): a level whose value differs from the
// default flips the state, one that matches leaves it. With the default off that is plain XOR of "on".
export const TOGGLE_KEYS = [
	"bold",
	"boldCs",
	"italic",
	"italicCs",
	"caps",
	"smallCaps",
	"strike",
	"outline",
	"shadow",
	"emboss",
	"imprint",
	"vanish"
] as const

const PLAIN_RUN_KEYS = [
	"doubleStrike",
	"noProof",
	"snapToGrid",
	"webHidden",
	"specVanish",
	"rtl",
	"cs",
	"oMath",
	"color",
	"spacing",
	"scale",
	"kern",
	"position",
	"size",
	"sizeCs",
	"highlight",
	"underline",
	"effect",
	"border",
	"shading",
	"fitText",
	"vertAlign",
	"emphasis",
	"eastAsianLayout"
] as const

const FONT_PAIRS = [
	["ascii", "asciiTheme"],
	["hAnsi", "hAnsiTheme"],
	["eastAsia", "eastAsiaTheme"],
	["cs", "csTheme"]
] as const

// A slot set one way (name or theme reference) at a higher level drops the other form from below, or the
// lower concrete name would defeat the theme reference.
function mergeFonts(target: FontsProps | undefined, source: FontsProps): FontsProps {
	const out: Mutable<FontsProps> = { ...target }

	for (const [concrete, theme] of FONT_PAIRS) {
		const sourceTheme = source[theme]
		const sourceConcrete = source[concrete]

		if (sourceTheme !== undefined) {
			delete out[concrete]
			out[theme] = sourceTheme
		} else if (sourceConcrete !== undefined) {
			delete out[theme]
		}

		if (sourceConcrete !== undefined) out[concrete] = sourceConcrete
	}

	if (source.hint !== undefined) out.hint = source.hint

	return out
}

function mergeLanguage(target: LanguageProps | undefined, source: LanguageProps): LanguageProps {
	return { ...target, ...source }
}

// `defaults` (the document defaults) marks a style level: its toggles flip instead of setting.
export function mergeRun(target: Mutable<RunProps>, source: RunProps, defaults?: RunProps): void {
	if (source === target) return

	if (source.fonts !== undefined) target.fonts = mergeFonts(target.fonts, source.fonts)
	if (source.lang !== undefined) target.lang = mergeLanguage(target.lang, source.lang)

	for (const key of TOGGLE_KEYS) {
		const value = source[key]

		if (value === undefined) continue

		if (defaults !== undefined) {
			if (value !== (defaults[key] ?? false)) target[key] = !(target[key] ?? false)
		} else {
			target[key] = value
		}
	}

	for (const key of PLAIN_RUN_KEYS) {
		const value = source[key]

		if (value !== undefined) (target as Record<string, unknown>)[key] = value
	}
}

function mergeSpacing(target: SpacingProps | undefined, source: SpacingProps): SpacingProps {
	const out: Mutable<SpacingProps> = { ...target, ...source }

	// A plain twips value set above replaces a line-based value from below: *Lines would otherwise win.
	if (source.before !== undefined && source.beforeLines === undefined) delete out.beforeLines
	if (source.after !== undefined && source.afterLines === undefined) delete out.afterLines

	return out
}

export function mergeIndent(target: Mutable<IndentProps>, source: IndentProps): void {
	if (source.firstLine !== undefined) delete target.hanging
	if (source.hanging !== undefined) delete target.firstLine
	if (source.firstLineChars !== undefined) delete target.hangingChars
	if (source.hangingChars !== undefined) delete target.firstLineChars

	Object.assign(target, source)
}

// Tabs accumulate; a "clear" tab removes the inherited stop at its position.
function mergeTabs(target: readonly TabStop[] | undefined, source: readonly TabStop[]): TabStop[] {
	const cleared = new Set<number>()

	for (const tab of source) {
		if (tab.alignment === "clear") cleared.add(tab.position)
	}

	const out = (target ?? []).filter(tab => !cleared.has(tab.position) && !source.some(next => next.position === tab.position))

	for (const tab of source) {
		if (tab.alignment !== "clear") out.push(tab)
	}

	return out
}

function mergeSides<T extends object>(target: T | undefined, source: T): T {
	return { ...target, ...source }
}

const PLAIN_PARAGRAPH_KEYS = [
	"keepNext",
	"keepLines",
	"pageBreakBefore",
	"frame",
	"widowControl",
	"suppressLineNumbers",
	"shading",
	"suppressAutoHyphens",
	"kinsoku",
	"wordWrap",
	"overflowPunct",
	"topLinePunct",
	"autoSpaceDE",
	"autoSpaceDN",
	"bidi",
	"adjustRightInd",
	"snapToGrid",
	"contextualSpacing",
	"mirrorIndents",
	"suppressOverlap",
	"alignment",
	"textDirection",
	"textAlignment",
	"textboxTightWrap",
	"outlineLevel"
] as const

// Everything but the paragraph style, the mark's run properties and (optionally) the indent.
export function mergeParagraph(target: Mutable<ParagraphProps>, source: ParagraphProps, indent: boolean): void {
	if (source === target) return

	for (const key of PLAIN_PARAGRAPH_KEYS) {
		const value = source[key]

		if (value !== undefined) (target as Record<string, unknown>)[key] = value
	}

	if (source.spacing !== undefined) target.spacing = mergeSpacing(target.spacing, source.spacing)
	if (source.tabs !== undefined) target.tabs = mergeTabs(target.tabs, source.tabs)
	if (source.borders !== undefined) target.borders = mergeSides<ParagraphBordersProps>(target.borders, source.borders)
	if (source.numbering !== undefined) target.numbering = { ...target.numbering, ...source.numbering }

	if (indent && source.indent !== undefined) {
		const merged: Mutable<IndentProps> = { ...target.indent }

		mergeIndent(merged, source.indent)
		target.indent = merged
	}
}

const PLAIN_TABLE_KEYS = [
	"position",
	"overlap",
	"bidiVisual",
	"rowBandSize",
	"colBandSize",
	"width",
	"alignment",
	"cellSpacing",
	"indent",
	"shading",
	"layout",
	"look",
	"caption",
	"description"
] as const

export function mergeTable(target: Mutable<TableProps>, source: TableProps): void {
	for (const key of PLAIN_TABLE_KEYS) {
		const value = source[key]

		if (value !== undefined) (target as Record<string, unknown>)[key] = value
	}

	if (source.borders !== undefined) target.borders = mergeSides<TableBordersProps>(target.borders, source.borders)
	if (source.cellMargins !== undefined) target.cellMargins = mergeSides<CellMarginsProps>(target.cellMargins, source.cellMargins)
}

const PLAIN_ROW_KEYS = [
	"gridBefore",
	"gridAfter",
	"widthBefore",
	"widthAfter",
	"cantSplit",
	"height",
	"header",
	"cellSpacing",
	"alignment",
	"hidden"
] as const

export function mergeRow(target: Mutable<TableRowProps>, source: TableRowProps): void {
	for (const key of PLAIN_ROW_KEYS) {
		const value = source[key]

		if (value !== undefined) (target as Record<string, unknown>)[key] = value
	}
}

const PLAIN_CELL_KEYS = [
	"width",
	"gridSpan",
	"hMerge",
	"vMerge",
	"shading",
	"noWrap",
	"textDirection",
	"fitText",
	"vAlign",
	"hideMark"
] as const

export function mergeCell(target: Mutable<TableCellProps>, source: TableCellProps): void {
	for (const key of PLAIN_CELL_KEYS) {
		const value = source[key]

		if (value !== undefined) (target as Record<string, unknown>)[key] = value
	}

	if (source.borders !== undefined) target.borders = mergeCellBorders(target.borders, source.borders)
	if (source.margins !== undefined) target.margins = mergeSides<CellMarginsProps>(target.margins, source.margins)
}

// A cell side of "none" leaves the border beneath it (the table style's, then the table's) in place: only
// "nil" removes one (MS-OI29500 §17.4.66 a).
function mergeCellBorders(target: CellBordersProps | undefined, source: CellBordersProps): CellBordersProps {
	const out: Record<string, unknown> = { ...target }

	for (const [side, border] of Object.entries(source) as [string, CellBordersProps["top"]][])
		if (border !== undefined && border.style !== "none") out[side] = border

	return out as CellBordersProps
}
