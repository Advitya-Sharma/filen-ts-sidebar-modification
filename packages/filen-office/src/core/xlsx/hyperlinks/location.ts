import type { SheetId, SheetRange } from "../../sheet/protocol"
import type { XlsxWorkbookModel } from "../model/types"
import { parseRangeRef } from "../write/text"

// Where a link inside the workbook goes: a location (§18.3.1.47 location; HYPERLINK's "#" link_location) is a
// reference, optionally on a named sheet, or a defined name (local to the link's sheet first, then the workbook's).

export interface LinkPlace {
	readonly sheet: SheetId
	readonly range: SheetRange
}

// The last "!" outside a quoted sheet name, -1 for none.
function sheetSeparator(text: string): number {
	let quoted = false
	let found = -1

	for (let k = 0; k < text.length; k++) {
		const char = text[k]

		if (char === "'") quoted = !quoted
		else if (char === "!" && !quoted) found = k
	}

	return found
}

function sheetName(text: string): string {
	const trimmed = text.trim()

	return trimmed.startsWith("'") && trimmed.endsWith("'") && trimmed.length >= 2 ? trimmed.slice(1, -1).replaceAll("''", "'") : trimmed
}

function reference(text: string): SheetRange | null {
	return parseRangeRef(text.trim().replaceAll("$", "")) ?? null
}

export function resolvePlace(workbook: XlsxWorkbookModel, own: SheetId, location: string, depth = 0): LinkPlace | null {
	let text = location.trim()

	if (text.startsWith("#")) text = text.slice(1)
	if (text === "" || depth > 4) return null

	const bang = sheetSeparator(text)

	if (bang >= 0) {
		const name = sheetName(text.slice(0, bang)).toUpperCase()
		const sheet = workbook.sheets.find(entry => entry.name.toUpperCase() === name)
		const range = reference(text.slice(bang + 1))

		return sheet === undefined || sheet.kind !== "worksheet" || range === null ? null : { sheet: sheet.id, range }
	}

	const range = reference(text)

	if (range !== null) return { sheet: own, range }

	const upper = text.toUpperCase()
	const defined =
		workbook.definedNames.find(name => name.sheet === own && name.name.toUpperCase() === upper) ??
		workbook.definedNames.find(name => name.sheet === null && name.name.toUpperCase() === upper)

	return defined === undefined ? null : resolvePlace(workbook, defined.sheet ?? own, defined.formula, depth + 1)
}

// A HYPERLINK formula's arguments as written: up to its first two, split at top-level commas (strings, parentheses
// and array constants kept whole). Null when the formula is not one call of HYPERLINK.
export function hyperlinkArguments(formula: string): string[] | null {
	const match = /^\s*=?\s*(?:_xlfn\.)?HYPERLINK\s*\(/i.exec(formula)

	if (match === null) return null

	const args: string[] = []
	let depth = 0
	let quoted = false
	let start = match[0].length

	for (let k = start; k < formula.length; k++) {
		const char = formula[k]

		if (quoted) {
			if (char === '"') {
				if (formula[k + 1] === '"') k++
				else quoted = false
			}

			continue
		}

		if (char === '"') quoted = true
		else if (char === "(" || char === "{") depth++
		else if (char === ")" || char === "}") {
			if (depth === 0) {
				args.push(formula.slice(start, k).trim())

				return formula.slice(k + 1).trim() === "" ? args : null
			}

			depth--
		} else if (char === "," && depth === 0) {
			args.push(formula.slice(start, k).trim())
			start = k + 1
		}
	}

	return null
}

// A string literal's text, null when the argument is not one.
export function stringLiteral(argument: string): string | null {
	return /^"(?:[^"]|"")*"$/.test(argument) ? argument.slice(1, -1).replaceAll('""', '"') : null
}
