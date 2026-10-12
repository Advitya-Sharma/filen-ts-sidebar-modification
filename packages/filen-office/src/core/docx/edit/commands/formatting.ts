import { MAX_TWIPS, choiceArg, field, hexArg, intArg, optional, stringArg, type DocxCommandInfo, type DocxCommandParser } from "./base"

// Character and paragraph formatting beyond the base toggles, and the format painter.

// ST_HighlightColor.
export const DOCX_HIGHLIGHT_COLORS = [
	"none",
	"yellow",
	"green",
	"cyan",
	"magenta",
	"blue",
	"red",
	"darkBlue",
	"darkCyan",
	"darkGreen",
	"darkMagenta",
	"darkRed",
	"darkYellow",
	"darkGray",
	"lightGray",
	"black"
] as const

export type DocxHighlightColor = (typeof DOCX_HIGHLIGHT_COLORS)[number]

// ST_ThemeColor.
export const DOCX_THEME_COLORS = [
	"dark1",
	"light1",
	"dark2",
	"light2",
	"accent1",
	"accent2",
	"accent3",
	"accent4",
	"accent5",
	"accent6",
	"hyperlink",
	"followedHyperlink",
	"background1",
	"text1",
	"background2",
	"text2"
] as const

export type DocxThemeColor = (typeof DOCX_THEME_COLORS)[number]

// A text colour: automatic, an RGB value ("RRGGBB"), or a theme colour with an optional tint or shade (0–255).
export type DocxTextColor =
	| { readonly kind: "auto" }
	| { readonly kind: "rgb"; readonly value: string }
	| { readonly kind: "theme"; readonly theme: DocxThemeColor; readonly tint?: number; readonly shade?: number }

export type DocxLineSpacingRule = "auto" | "exact" | "atLeast"

export type FormattingCommand =
	| { readonly type: "format.superscript" }
	| { readonly type: "format.subscript" }
	| { readonly type: "font.family"; readonly family: string }
	// Half-points.
	| { readonly type: "font.size"; readonly halfPoints: number }
	| { readonly type: "font.grow" }
	| { readonly type: "font.shrink" }
	| { readonly type: "color.text"; readonly color: DocxTextColor }
	| { readonly type: "color.highlight"; readonly name: DocxHighlightColor }
	// "auto": `value` in 240ths of a line; "exact" and "atLeast": twips.
	| { readonly type: "paragraph.lineSpacing"; readonly rule: DocxLineSpacingRule; readonly value: number }
	// Twips.
	| { readonly type: "paragraph.spacing"; readonly before?: number; readonly after?: number }
	| {
			readonly type: "paragraph.indent"
			readonly start?: number
			readonly end?: number
			readonly firstLine?: number
			readonly hanging?: number
	  }
	| { readonly type: "format.copy" }
	| { readonly type: "format.paste" }

const story = { write: true, scope: "story" } as const

export const FORMATTING_COMMANDS: readonly DocxCommandInfo<FormattingCommand["type"]>[] = [
	{ type: "format.superscript", labelKey: "docx.format.superscript", icon: "superscript", group: "format", ...story },
	{ type: "format.subscript", labelKey: "docx.format.subscript", icon: "subscript", group: "format", ...story },
	{ type: "font.family", labelKey: "docx.font.family", icon: "font-family", group: "font", input: "font", ...story },
	{ type: "font.size", labelKey: "docx.font.size", icon: "font-size", group: "font", input: "size", ...story },
	{ type: "font.grow", labelKey: "docx.font.grow", icon: "font-grow", group: "font", ...story },
	{ type: "font.shrink", labelKey: "docx.font.shrink", icon: "font-shrink", group: "font", ...story },
	{ type: "color.text", labelKey: "docx.color.text", icon: "text-color", group: "color", input: "color", ...story },
	{ type: "color.highlight", labelKey: "docx.color.highlight", icon: "highlight", group: "color", input: "highlight", ...story },
	{ type: "paragraph.lineSpacing", labelKey: "docx.paragraph.lineSpacing", icon: "line-spacing", group: "paragraph", ...story },
	{ type: "paragraph.spacing", labelKey: "docx.paragraph.spacing", icon: "paragraph-spacing", group: "paragraph", ...story },
	{ type: "paragraph.indent", labelKey: "docx.paragraph.indent", icon: "indent", group: "paragraph", ...story },
	// The painter's formatting is the editor's (it outlives story switches): document commands. Copying it
	// changes nothing; applying it does.
	{ type: "format.copy", labelKey: "docx.format.copy", icon: "format-painter", group: "painter", write: false, scope: "document" },
	{ type: "format.paste", labelKey: "docx.format.paste", icon: "format-painter", group: "painter", write: true, scope: "document" }
]

const HIGHLIGHTS = new Set<string>(DOCX_HIGHLIGHT_COLORS)
const THEMES = new Set<string>(DOCX_THEME_COLORS)
const RULES = new Set<string>(["auto", "exact", "atLeast"])
// Font sizes Word accepts: 1 to 1638 points.
const MAX_HALF_POINTS = 3276

function textColor(value: object): DocxTextColor | undefined {
	const color = field(value, "color")

	if (typeof color !== "object" || color === null) return undefined

	const kind = field(color, "kind")

	if (kind === "auto") return { kind }

	if (kind === "rgb") {
		const rgb = hexArg(color, "value")

		return rgb === undefined ? undefined : { kind, value: rgb }
	}

	if (kind !== "theme") return undefined

	const theme = choiceArg<DocxThemeColor>(color, "theme", THEMES)
	const tint = optional(color, "tint", (holder, key) => intArg(holder, key, 0, 255))
	const shade = optional(color, "shade", (holder, key) => intArg(holder, key, 0, 255))

	if (theme === undefined || !tint.ok || !shade.ok) return undefined

	return {
		kind,
		theme,
		...(tint.value === undefined ? {} : { tint: tint.value }),
		...(shade.value === undefined ? {} : { shade: shade.value })
	}
}

const twips = (min: number) => (holder: object, key: string) => intArg(holder, key, min, MAX_TWIPS)

export const parseFormattingCommand: DocxCommandParser<FormattingCommand> = (type, value) => {
	switch (type) {
		case "font.family": {
			// w:rFonts names: Word keeps at most 31 characters.
			const family = stringArg(value, "family", 31, 1)

			return family === undefined ? undefined : { type, family }
		}
		case "font.size": {
			const halfPoints = intArg(value, "halfPoints", 2, MAX_HALF_POINTS)

			return halfPoints === undefined ? undefined : { type, halfPoints }
		}
		case "color.text": {
			const color = textColor(value)

			return color === undefined ? undefined : { type, color }
		}
		case "color.highlight": {
			const name = choiceArg<DocxHighlightColor>(value, "name", HIGHLIGHTS)

			return name === undefined ? undefined : { type, name }
		}
		case "paragraph.lineSpacing": {
			const rule = choiceArg<DocxLineSpacingRule>(value, "rule", RULES)
			const amount =
				rule === undefined ? undefined : intArg(value, "value", rule === "auto" ? 1 : 0, rule === "auto" ? 31680 : MAX_TWIPS)

			return rule === undefined || amount === undefined ? undefined : { type, rule, value: amount }
		}
		case "paragraph.spacing": {
			const before = optional(value, "before", twips(0))
			const after = optional(value, "after", twips(0))

			if (!before.ok || !after.ok || (before.value === undefined && after.value === undefined)) return undefined

			return {
				type,
				...(before.value === undefined ? {} : { before: before.value }),
				...(after.value === undefined ? {} : { after: after.value })
			}
		}
		case "paragraph.indent": {
			const start = optional(value, "start", twips(-MAX_TWIPS))
			const end = optional(value, "end", twips(-MAX_TWIPS))
			const firstLine = optional(value, "firstLine", twips(0))
			const hanging = optional(value, "hanging", twips(0))

			if (!start.ok || !end.ok || !firstLine.ok || !hanging.ok) return undefined
			if (start.value === undefined && end.value === undefined && firstLine.value === undefined && hanging.value === undefined)
				return undefined
			// w:ind takes one of the two.
			if (firstLine.value !== undefined && hanging.value !== undefined) return undefined

			return {
				type,
				...(start.value === undefined ? {} : { start: start.value }),
				...(end.value === undefined ? {} : { end: end.value }),
				...(firstLine.value === undefined ? {} : { firstLine: firstLine.value }),
				...(hanging.value === undefined ? {} : { hanging: hanging.value })
			}
		}
		default:
			return { type } as FormattingCommand
	}
}
