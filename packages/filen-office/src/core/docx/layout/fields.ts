import { formatNumber } from "../numbering/formats"

// Fields whose result layout computes per page instead of showing the cached result Word stored.
export const FieldKind = { OTHER: 0, PAGE: 1, NUMPAGES: 2, SECTIONPAGES: 3 } as const

export type FieldKind = (typeof FieldKind)[keyof typeof FieldKind]

export interface FieldInfo {
	readonly kind: FieldKind
	// The ST_NumberFormat a \* switch asks for, or undefined to use the section's (PAGE) or decimal.
	readonly format: string | undefined
}

const OTHER: FieldInfo = { kind: FieldKind.OTHER, format: undefined }

// \* general formatting switches that format numbers (ECMA-376 §17.16.4.3), as ST_NumberFormat values.
const NUMBER_SWITCHES: Readonly<Record<string, string>> = {
	arabic: "decimal",
	roman: "lowerRoman",
	Roman: "upperRoman",
	ROMAN: "upperRoman",
	alphabetic: "lowerLetter",
	ALPHABETIC: "upperLetter",
	Alphabetic: "upperLetter",
	ordinal: "ordinal",
	Ordinal: "ordinal",
	cardtext: "cardinalText",
	CardText: "cardinalText",
	ordtext: "ordinalText",
	OrdText: "ordinalText",
	hex: "hexLower",
	HEX: "hexUpper"
}

const FIELD = /^\s*([A-Za-z]+)/
const SWITCH = /\\\*\s*([A-Za-z]+)/g

export function fieldInfo(instruction: string): FieldInfo {
	const type = FIELD.exec(instruction)?.[1]?.toUpperCase()
	const kind =
		type === "PAGE"
			? FieldKind.PAGE
			: type === "NUMPAGES"
				? FieldKind.NUMPAGES
				: type === "SECTIONPAGES"
					? FieldKind.SECTIONPAGES
					: FieldKind.OTHER

	if (kind === FieldKind.OTHER) return OTHER

	let format: string | undefined

	for (const match of instruction.matchAll(SWITCH)) {
		const name = match[1] ?? ""
		// Hex switches are case-sensitive (\* hex lower case digits, \* HEX upper case); the others are not.
		const mapped = NUMBER_SWITCHES[name] ?? (name.toLowerCase() === "hex" ? undefined : NUMBER_SWITCHES[name.toLowerCase()])

		if (mapped !== undefined) format = mapped
	}

	return { kind, format }
}

// A page number as a field shows it: the switch's format, else (PAGE) the section's w:pgNumType format.
export function numberText(value: number, format: string | undefined): string {
	if (format === "hexLower") return value.toString(16)
	if (format === "hexUpper") return value.toString(16).toUpperCase()

	return formatNumber(value, format ?? "decimal")
}
