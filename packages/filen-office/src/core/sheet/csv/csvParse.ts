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

// The records of `text`, separated by `delimiter` (one character).
export function parseCsv(text: string, delimiter: string): string[][] {
	const separator = delimiter.charCodeAt(0)
	const rows: string[][] = []
	const length = text.length
	let row: string[] = []
	let i = 0

	while (i < length) {
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

		row.push(field)

		if (i >= length) {
			rows.push(row)

			return rows
		}

		const code = text.charCodeAt(i++)

		if (code === separator) {
			// A separator ends the text: the record closes with an empty field.
			if (i >= length) {
				row.push("")
				rows.push(row)

				return rows
			}

			continue
		}

		if (code === CR && text.charCodeAt(i) === LF) {
			i++
		}

		rows.push(row)
		row = []
	}

	return rows
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

// The separator a CSV file uses, which it does not say: of comma, semicolon, tab and bar, the one found the
// same number of times on the most of the first ten lines (then the most often per line), a comma when none
// is found. Excel itself takes its locale's list separator.
export function detectDelimiter(text: string): string {
	const lines = sampleLines(text)
	let best = ","
	let bestScore = -1

	for (const delimiter of DELIMITERS) {
		const separator = delimiter.charCodeAt(0)
		const counts = lines.map(line => countUnquoted(line, separator)).filter(count => count > 0)
		const mode = counts[0]

		if (mode === undefined) {
			continue
		}

		const score = counts.filter(count => count === mode).length * 1000 + mode

		if (score > bestScore) {
			bestScore = score
			best = delimiter
		}
	}

	return best
}

// Records as text: a field is quoted only when it holds the separator, a quote or a line end, its quotes
// doubled. Records are joined by `lineSeparator`, with none after the last.
export function writeCsv(rows: readonly (readonly string[])[], delimiter: string, lineSeparator: string): string {
	const lines = new Array<string>(rows.length)

	for (let r = 0; r < rows.length; r++) {
		const row = rows[r] ?? []
		const fields = new Array<string>(row.length)

		for (let c = 0; c < row.length; c++) {
			const value = row[c] ?? ""

			fields[c] =
				value.includes(delimiter) || value.includes('"') || value.includes("\n") || value.includes("\r")
					? `"${value.replaceAll('"', '""')}"`
					: value
		}

		lines[r] = fields.join(delimiter)
	}

	return lines.join(lineSeparator)
}
