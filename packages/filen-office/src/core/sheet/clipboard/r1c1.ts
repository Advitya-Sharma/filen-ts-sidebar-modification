import { columnName } from "../refs/cellRef"
import { SHEET_MAX_COLS, SHEET_MAX_ROWS } from "../protocol"

// R1C1 references (as Google Sheets puts formulas on the clipboard, data-sheets-formula) as A1 ones for the cell at
// `row`, `col` (0-based). R[n] and C[n] are relative, Rn and Cn absolute ($), R or C alone the cell's own row or
// column; a reference naming only rows (columns) is whole rows (columns). Text in quotes ("…" literals, '…' sheet
// names) is left as written, as are words that only start like a reference (ROUND, COUNT). A reference past the sheet
// becomes #REF!.

const PART = /^(?:\[(-?\d{1,7})\]|(\d{1,7}))?/

interface Part {
	readonly kind: "row" | "col"
	readonly index: number
	readonly absolute: boolean
}

function isWordChar(code: number): boolean {
	return (
		(code >= 48 && code <= 57) || (code >= 65 && code <= 90) || (code >= 97 && code <= 122) || code === 95 || code === 46 || code >= 128
	)
}

// A reference at `at`: its parts and where it ends; null when the text there is not one.
function readReference(text: string, at: number, row: number, col: number): { parts: Part[]; end: number } | null {
	const parts: Part[] = []
	let i = at

	for (const [letter, kind] of [
		["R", "row"],
		["C", "col"]
	] as const) {
		if ((text[i] ?? "").toUpperCase() !== letter) continue

		const match = PART.exec(text.slice(i + 1, i + 12))
		const relative = match?.[1]
		const absolute = match?.[2]
		const own = kind === "row" ? row : col

		parts.push(
			absolute !== undefined
				? { kind, index: Number(absolute) - 1, absolute: true }
				: { kind, index: own + (relative === undefined ? 0 : Number(relative)), absolute: false }
		)
		i += 1 + (match?.[0].length ?? 0)
	}

	if (parts.length === 0 || (i < text.length && (isWordChar(text.charCodeAt(i)) || text[i] === "("))) return null

	return { parts, end: i }
}

function a1(parts: readonly Part[]): string | null {
	let out = ""

	for (const part of parts) {
		if (part.index < 0 || part.index >= (part.kind === "row" ? SHEET_MAX_ROWS : SHEET_MAX_COLS)) return null

		const name = part.kind === "row" ? String(part.index + 1) : columnName(part.index)
		const written = `${part.absolute ? "$" : ""}${name}`

		// A1 names the column first.
		out = part.kind === "col" ? written + out : out + written
	}

	return out
}

export function r1c1ToA1(formula: string, row: number, col: number): string {
	let out = ""
	let i = 0

	while (i < formula.length) {
		const char = formula[i] ?? ""

		if (char === '"' || char === "'") {
			let end = i + 1

			while (end < formula.length) {
				if (formula[end] === char) {
					if (formula[end + 1] === char) {
						end += 2

						continue
					}

					break
				}

				end++
			}

			out += formula.slice(i, end + 1)
			i = end + 1

			continue
		}

		const previous = i === 0 ? 32 : formula.charCodeAt(i - 1)
		const reference = isWordChar(previous) ? null : readReference(formula, i, row, col)

		if (reference === null) {
			out += char
			i++

			continue
		}

		const only = reference.parts.length === 1 ? (reference.parts[0]?.kind ?? null) : null
		let end = reference.end
		let last: Part[] | null = null

		// A whole-row or whole-column range: both ends name the same axis only.
		if (only !== null && formula[end] === ":") {
			const second = readReference(formula, end + 1, row, col)

			if (second !== null && second.parts.length === 1 && second.parts[0]?.kind === only) {
				last = second.parts
				end = second.end
			}
		}

		const first = a1(reference.parts)
		const second = last === null ? first : a1(last)

		if (first === null || second === null) out += "#REF!"
		else out += only === null ? first : `${first}:${second}`

		i = end
	}

	return out
}
