import { SHEET_MAX_COLS, SHEET_MAX_ROWS } from "../../sheet/protocol"
import type { AxisEdit } from "../../sheet/model/edits"
import { shiftSpan } from "../../sheet/refs/formulaRefs"

// Formula text a quick scan can move: cell references and ranges among operators, numbers and function names; no
// sheet names, strings, structured references, errors, arrays, whole rows or columns. Anything else (or anything
// next to a reference that could make it a name) goes through the formula parser (shiftFormula), whose text this
// matches on what it takes.

function isLetter(code: number): boolean {
	return (code >= 65 && code <= 90) || (code >= 97 && code <= 122)
}

function isDigit(code: number): boolean {
	return code >= 48 && code <= 57
}

// What may continue a name or a number, as the parser reads a word: letters, digits, _ . $ ? \ and every code
// unit from 128 on.
function isNameChar(code: number): boolean {
	return isLetter(code) || isDigit(code) || code === 95 || code === 46 || code === 36 || code === 63 || code === 92 || code >= 128
}

// Characters the scan leaves to the parser: ' " ! [ ] # { } @ \ and a colon outside a cell range.
function isParserOnly(code: number): boolean {
	return (
		code === 39 ||
		code === 34 ||
		code === 33 ||
		code === 91 ||
		code === 93 ||
		code === 35 ||
		code === 123 ||
		code === 125 ||
		code === 64 ||
		code === 92 ||
		code === 58
	)
}

// One cell reference at `at`: [$]letters[$]digits, or -1 when there is none.
interface QuickCell {
	end: number
	colAbs: boolean
	col: number
	rowAbs: boolean
	row: number
}

const scratch: QuickCell = { end: -1, colAbs: false, col: 0, rowAbs: false, row: 0 }
const second: QuickCell = { end: -1, colAbs: false, col: 0, rowAbs: false, row: 0 }

function quickCell(text: string, at: number, out: QuickCell): boolean {
	let i = at

	out.colAbs = text.charCodeAt(i) === 36

	if (out.colAbs) i++

	let col = 0
	const lettersStart = i

	while (i < text.length && isLetter(text.charCodeAt(i)) && i - lettersStart < 4) {
		col = col * 26 + ((text.charCodeAt(i) | 0x20) - 96)
		i++
	}

	if (i === lettersStart || i - lettersStart > 3) return false

	out.rowAbs = text.charCodeAt(i) === 36

	if (out.rowAbs) i++

	let row = 0
	const digitsStart = i

	while (i < text.length && isDigit(text.charCodeAt(i)) && i - digitsStart < 8) {
		row = row * 10 + text.charCodeAt(i) - 48
		i++
	}

	if (i === digitsStart || i - digitsStart > 7) return false

	out.col = col - 1
	out.row = row - 1
	out.end = i

	return true
}

function colLetters(col: number): string {
	let name = ""

	for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name

	return name
}

function cellText(colAbs: boolean, col: number, rowAbs: boolean, row: number): string {
	return `${colAbs ? "$" : ""}${colLetters(col)}${rowAbs ? "$" : ""}${String(row + 1)}`
}

// A formula on the target sheet after rows or columns of it moved, or null when the quick scan cannot tell.
export function quickShift(text: string, edit: AxisEdit): string | null {
	const rows = edit.axis === "rows"
	const limit = rows ? SHEET_MAX_ROWS : SHEET_MAX_COLS
	let out = ""
	let copied = 0
	let i = 0

	while (i < text.length) {
		const code = text.charCodeAt(i)

		if (isParserOnly(code)) return null

		if (!isLetter(code) && code !== 36) {
			// A number or a name not starting with a letter: skipped whole (1E5 is not E5).
			if (isNameChar(code)) {
				while (i < text.length && isNameChar(text.charCodeAt(i))) i++
			} else {
				i++
			}

			continue
		}

		if (!quickCell(text, i, scratch)) {
			// A word (SUM, TRUE): skipped whole.
			while (i < text.length && isNameChar(text.charCodeAt(i))) i++

			continue
		}

		let end = scratch.end
		const range = text.charCodeAt(end) === 58

		if (range && !quickCell(text, end + 1, second)) return null
		if (range) end = second.end

		const next = end < text.length ? text.charCodeAt(end) : -1

		// Followed by more of a name, or a call (LOG10( ): the parser tells.
		if (next !== -1 && (isNameChar(next) || next === 40)) return null

		const startCol = scratch.col
		const startRow = scratch.row
		const endCol = range ? second.col : startCol
		const endRow = range ? second.row : startRow

		if (
			Math.max(startCol, endCol) >= SHEET_MAX_COLS ||
			Math.max(startRow, endRow) >= SHEET_MAX_ROWS ||
			Math.min(startRow, endRow, startCol, endCol) < 0 ||
			(rows ? startRow > endRow : startCol > endCol)
		)
			return null

		const span = shiftSpan(rows ? startRow : startCol, rows ? endRow : endCol, edit, limit, range)

		if (span === null || span[0] !== (rows ? startRow : startCol) || span[1] !== (rows ? endRow : endCol)) {
			let moved = "#REF!"

			if (span !== null) {
				const first = rows
					? cellText(scratch.colAbs, startCol, scratch.rowAbs, span[0])
					: cellText(scratch.colAbs, span[0], scratch.rowAbs, startRow)

				moved = range
					? `${first}:${rows ? cellText(second.colAbs, endCol, second.rowAbs, span[1]) : cellText(second.colAbs, span[1], second.rowAbs, endRow)}`
					: first
			}

			out += text.slice(copied, i) + moved
			copied = end
		}

		i = end
	}

	return copied === 0 ? text : out + text.slice(copied)
}
