import type { ColorAttrs, IndentAttrs, SpacingAttrs } from "../edit/schema"
import type { ParagraphBordersProps } from "../props/index"
import { THEME_COLOR } from "./order.generated"

export type WriteFailure =
	// An edit the writer cannot express yet (a changed table grid, a new relationship without an r prefix).
	| "unsupported"
	// Content that only exists as source bytes appeared without them.
	| "missingSource"
	// A value the schema does not allow where the writer would emit it.
	| "invalidValue"

export class DocxWriteError extends Error {
	readonly reason: WriteFailure

	constructor(reason: WriteFailure, message: string) {
		super(message)
		this.name = "DocxWriteError"
		this.reason = reason
	}
}

// Checks of the values the writer emits, against their ECMA-376 simple types: a value outside its type
// refuses the save rather than reaching the file.

const INT32 = 2 ** 31

function invalid(what: string, value: unknown): never {
	throw new DocxWriteError("invalidValue", `${what}: ${JSON.stringify(value)} is not a valid value`)
}

// ST_DecimalNumber and the integer forms of the twips and half-point measures.
export function integer(what: string, value: unknown, min = -INT32, max = INT32 - 1): void {
	if (typeof value !== "number" || !Number.isInteger(value) || value < min || value > max) invalid(what, value)
}

export function optionalInteger(what: string, value: unknown, min?: number, max?: number): void {
	if (value !== undefined && value !== null) integer(what, value, min, max)
}

export function oneOf(what: string, values: ReadonlySet<string>, value: unknown): void {
	if (typeof value !== "string" || !values.has(value)) invalid(what, value)
}

export function optionalOneOf(what: string, values: ReadonlySet<string>, value: unknown): void {
	if (value !== undefined && value !== null) oneOf(what, values, value)
}

// ST_String values the writer names things by (style ids): non-empty, without C0 controls, noncharacters
// or lone surrogates.
export function name(what: string, value: unknown): void {
	if (typeof value !== "string" || value === "") invalid(what, value)

	for (let k = 0; k < value.length; k++) {
		const code = value.charCodeAt(k)

		if (code >= 0xd800 && code <= 0xdbff && (value.charCodeAt(k + 1) & 0xfc00) === 0xdc00) {
			k++

			continue
		}

		if (code < 0x20 || (code >= 0xd800 && code <= 0xdfff) || code === 0xfffe || code === 0xffff) invalid(what, value)
	}
}

// ST_HexColor: "auto" or six hex digits.
const HEX_COLOR = /^(?:auto|[0-9A-Fa-f]{6})$/

export function color(what: string, value: ColorAttrs | null, themes: ReadonlySet<string>): void {
	if (value === null) return
	if (value.value !== null && !HEX_COLOR.test(value.value)) invalid(what, value.value)

	optionalOneOf(`${what} theme`, themes, value.themeColor)
	optionalInteger(`${what} tint`, value.themeTint, 0, 255)
	optionalInteger(`${what} shade`, value.themeShade, 0, 255)
}

// CT_Ind: start/end are signed twips, hanging/firstLine unsigned twips, the *Chars forms decimal numbers.
export function indent(value: IndentAttrs | null): void {
	if (value === null) return

	optionalInteger("w:ind start", value.start)
	optionalInteger("w:ind end", value.end)
	optionalInteger("w:ind hanging", value.hanging, 0)
	optionalInteger("w:ind firstLine", value.firstLine, 0)
	optionalInteger("w:ind startChars", value.startChars)
	optionalInteger("w:ind endChars", value.endChars)
	optionalInteger("w:ind hangingChars", value.hangingChars)
	optionalInteger("w:ind firstLineChars", value.firstLineChars)
}

const BORDER_SIDES = ["top", "start", "bottom", "end", "between", "bar"] as const

// CT_PBdr: each side a CT_Border (ST_Border style, eighths of a point, 0-31 points of space, a colour).
export function paragraphBorders(value: ParagraphBordersProps | null): void {
	if (value === null) return

	for (const side of BORDER_SIDES) {
		const border = value[side]

		if (border === undefined) continue

		name(`w:pBdr ${side}`, border.style)
		optionalInteger(`w:pBdr ${side} sz`, border.size, 0)
		optionalInteger(`w:pBdr ${side} space`, border.space, 0, 31)
		color(
			`w:pBdr ${side} color`,
			{
				value: border.val ?? null,
				themeColor: border.themeColor ?? null,
				themeTint: border.themeTint ?? null,
				themeShade: border.themeShade ?? null
			},
			THEME_COLOR
		)
	}
}

const LINE_RULES: ReadonlySet<string> = new Set(["auto", "exact", "atLeast"])

// CT_Spacing: before/after unsigned twips, line signed twips (240ths of a line for auto), *Lines decimal numbers.
export function spacing(value: SpacingAttrs | null): void {
	if (value === null) return

	optionalInteger("w:spacing before", value.before, 0)
	optionalInteger("w:spacing after", value.after, 0)
	optionalInteger("w:spacing beforeLines", value.beforeLines)
	optionalInteger("w:spacing afterLines", value.afterLines)
	optionalInteger("w:spacing line", value.line)
	optionalOneOf("w:spacing lineRule", LINE_RULES, value.lineRule)
}

// w14:paraId (ST_LongHexNumber as Word uses it): eight hex digits, above 0 and below 0x80000000.
const PARA_ID = /^[0-7][0-9A-Fa-f]{7}$/

export function paraId(value: unknown): void {
	if (typeof value !== "string" || !PARA_ID.test(value) || value === "00000000") invalid("w14:paraId", value)
}
