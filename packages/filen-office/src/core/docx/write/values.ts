import type { ColorAttrs, IndentAttrs } from "../edit/schema"

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
