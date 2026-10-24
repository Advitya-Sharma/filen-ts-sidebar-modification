// Delimiter-separated text as RFC 4180 (https://www.rfc-editor.org/rfc/rfc4180) describes it, read
// leniently where a file breaks it:
// - a field that opens with a double quote runs to the quote that closes it, and holds the separator and line
//   ends; a doubled quote inside it stands for one;
// - what follows the closing quote, up to the next separator or line end, belongs to the same field as it
//   stands, and so does a quote anywhere in a field that does not open with one;
// - a quote never closed takes the rest of the file;
// - CRLF, LF and CR alone all end a record (RFC 4180 has CRLF only); a line end after the last record adds no
//   empty one, and records may differ in their number of fields.
// Every field stays the text it was: nothing is trimmed or read as a number.

const QUOTE = 0x22
const CR = 0x0d
const LF = 0x0a

// Where the unquoted rest of a field ends: at the separator, a line end, or the end of the text.
function fieldEnd(text: string, from: number, separator: number): number {
	let i = from

	while (i < text.length) {
		const code = text.charCodeAt(i)

		if (code === separator || code === CR || code === LF) {
			break
		}

		i++
	}

	return i
}

// How much a parse may build: past any of these it throws, before the rest is read (an empty record counts as one
// cell). `keepFirstRows`: a record past maxRows ends the parse instead, its first maxRows records kept (as Excel and
// LibreOffice open a file longer than a sheet).
export interface CsvLimits {
	readonly maxCells?: number
	readonly maxRows?: number
	readonly maxColumns?: number
	readonly keepFirstRows?: boolean
}

// The records of `text`, separated by `delimiter` (one character).
export function parseCsv(text: string, delimiter: string, limits: CsvLimits = {}): string[][] {
	return parseCsvRecords(text, delimiter, limits).rows
}

// The records of `text`; `complete` is false when keepFirstRows stopped the parse before the end of the text.
export function parseCsvRecords(text: string, delimiter: string, limits: CsvLimits = {}): { rows: string[][]; complete: boolean } {
	const separator = delimiter.charCodeAt(0)
	const maxCells = limits.maxCells ?? Infinity
	const maxRows = limits.maxRows ?? Infinity
	const maxColumns = limits.maxColumns ?? Infinity
	const rows: string[][] = []
	const length = text.length
	// The record being read; each is kept as an array of its own size (a copy of these), not one grown by pushes.
	const fields: string[] = []
	let width = 0
	let cellCount = 0
	let i = 0

	while (i < length) {
		if (width === 0 && rows.length >= maxRows) {
			if (limits.keepFirstRows === true) return { rows, complete: false }

			throw new Error("spreadsheet: too many rows")
		}

		let field: string

		if (text.charCodeAt(i) === QUOTE) {
			let from = i + 1

			field = ""

			for (;;) {
				const quote = text.indexOf('"', from)

				if (quote === -1) {
					field += text.slice(from)
					i = length

					break
				}

				field += text.slice(from, quote)

				if (text.charCodeAt(quote + 1) === QUOTE) {
					field += '"'
					from = quote + 2

					continue
				}

				i = quote + 1

				break
			}

			const end = fieldEnd(text, i, separator)

			if (end > i) {
				field += text.slice(i, end)
			}

			i = end
		} else {
			const end = fieldEnd(text, i, separator)

			field = text.slice(i, end)
			i = end
		}

		if (++cellCount > maxCells) throw new Error("spreadsheet: too many cells")

		fields[width++] = field

		if (width > maxColumns) throw new Error("spreadsheet: too many columns")

		if (i >= length) {
			rows.push(fields.slice(0, width))

			return { rows, complete: true }
		}

		const code = text.charCodeAt(i++)

		if (code === separator) {
			// A separator ends the text: the record closes with an empty field.
			if (i >= length) {
				fields[width++] = ""

				if (++cellCount > maxCells) throw new Error("spreadsheet: too many cells")
				if (width > maxColumns) throw new Error("spreadsheet: too many columns")

				rows.push(fields.slice(0, width))

				return { rows, complete: true }
			}

			continue
		}

		if (code === CR && text.charCodeAt(i) === LF) {
			i++
		}

		rows.push(fields.slice(0, width))
		width = 0
	}

	return { rows, complete: true }
}

const DELIMITERS = [",", ";", "\t", "|"]
const SAMPLE_LINES = 10

// The first non-empty lines of the text, a quoted line end kept inside its line.
function sampleLines(text: string): string[] {
	const lines: string[] = []
	let quoted = false
	let start = 0
	let i = 0

	for (; i < text.length && lines.length < SAMPLE_LINES; i++) {
		const code = text.charCodeAt(i)

		if (code === QUOTE) {
			if (quoted && text.charCodeAt(i + 1) === QUOTE) {
				i++
			} else {
				quoted = !quoted
			}

			continue
		}

		if (!quoted && (code === LF || code === CR)) {
			if (i > start) {
				lines.push(text.slice(start, i))
			}

			if (code === CR && text.charCodeAt(i + 1) === LF) {
				i++
			}

			start = i + 1
		}
	}

	if (lines.length < SAMPLE_LINES && i >= text.length && text.length > start) {
		lines.push(text.slice(start))
	}

	return lines
}

// The separators in a line outside quotes.
function countUnquoted(line: string, separator: number): number {
	let count = 0
	let quoted = false

	for (let i = 0; i < line.length; i++) {
		const code = line.charCodeAt(i)

		if (code === QUOTE) {
			if (quoted && line.charCodeAt(i + 1) === QUOTE) {
				i++
			} else {
				quoted = !quoted
			}
		} else if (!quoted && code === separator) {
			count++
		}
	}

	return count
}

// How each separator found in the first ten lines is spread: on how many lines it is found its most common number
// of times (the first line's), and that number.
function delimiterScores(text: string): { delimiter: string; lines: number; mode: number }[] {
	const lines = sampleLines(text)
	const scores: { delimiter: string; lines: number; mode: number }[] = []

	for (const delimiter of DELIMITERS) {
		const separator = delimiter.charCodeAt(0)
		const counts = lines.map(line => countUnquoted(line, separator)).filter(count => count > 0)
		const mode = counts[0]

		if (mode !== undefined) scores.push({ delimiter, lines: counts.filter(count => count === mode).length, mode })
	}

	return scores
}

// The separator a CSV file uses, which it does not say: of comma, semicolon, tab and bar, the one found the
// same number of times on the most of the first ten lines, a comma when none is found. Excel itself takes its
// locale's list separator: `preferred` (the session locale's) wins where another is found on as many lines; else
// the one found the most often per line does.
export function detectDelimiter(text: string, preferred?: string): string {
	let best: { delimiter: string; lines: number; mode: number } | undefined

	for (const score of delimiterScores(text)) {
		if (
			best === undefined ||
			score.lines > best.lines ||
			(score.lines === best.lines && best.delimiter !== preferred && (score.delimiter === preferred || score.mode > best.mode))
		)
			best = score
	}

	return best?.delimiter ?? ","
}

// Whether `text` detects as separated by `delimiter` whatever the locale prefers.
export function detectsStably(text: string, delimiter: string): boolean {
	const scores = delimiterScores(text)
	const own = scores.find(score => score.delimiter === delimiter)

	if (own === undefined) return scores.length === 0 && delimiter === ","

	return scores.every(score => score === own || score.lines < own.lines)
}

// Records as text: a field is quoted when it holds the separator, a quote or a line end, its quotes doubled; with
// `stable`, also when it holds any other separator detectDelimiter weighs, so the text written detects as written.
// Records are joined by `lineSeparator`, with none after the last.
export function writeCsv(rows: readonly (readonly string[])[], delimiter: string, lineSeparator: string, stable = false): string {
	const lines = new Array<string>(rows.length)

	for (let r = 0; r < rows.length; r++) {
		const row = rows[r] ?? []
		const fields = new Array<string>(row.length)

		for (let c = 0; c < row.length; c++) fields[c] = csvField(row[c] ?? "", delimiter, stable)

		lines[r] = fields.join(delimiter)
	}

	return lines.join(lineSeparator)
}

const NEEDS_QUOTES = /["\r\n]/
const NEEDS_QUOTES_STABLE = /["\r\n,;\t|]/

// One field as writeCsv writes it.
export function csvField(value: string, delimiter: string, stable: boolean): string {
	return (stable ? NEEDS_QUOTES_STABLE : NEEDS_QUOTES).test(value) || value.includes(delimiter)
		? `"${value.replaceAll('"', '""')}"`
		: value
}

// The bytes `text` takes as UTF-8, as TextEncoder writes it (a lone surrogate as U+FFFD).
export function utf8Length(text: string): number {
	let length = 0

	for (let i = 0; i < text.length; i++) {
		const code = text.charCodeAt(i)

		if (code < 0x80) {
			length += 1
		} else if (code < 0x800) {
			length += 2
		} else if (code >= 0xd800 && code <= 0xdbff && (text.charCodeAt(i + 1) & 0xfc00) === 0xdc00) {
			length += 4
			i++
		} else {
			length += 3
		}
	}

	return length
}
