import { FIND_INTS, type DocxFindOptions } from "./protocol"

// Find over laid-out pages. A match belongs to the page whose line holds its first character, so a paragraph
// split over two pages (whose text both pages carry) reports each match once.

export interface FindablePage {
	readonly index: number
	readonly sourceCount: number
	readonly sourceTexts: readonly string[]
	readonly lineCount: number
	// lineInts as the text layout writes them (LINE_INTS per line; text start 2, end 3, source 5).
	readonly lineInts: Int32Array
}

const LINE_INTS = 7
const SPECIAL = /[.*+?^${}()|[\]\\]/g
const MARK = /\p{M}/u

// `text` in NFC, with each normalised code unit's cluster in the original (start, end): a match maps back to
// whole characters however the document composed them. Null maps when the text is already NFC.
interface Normalised {
	readonly text: string
	readonly starts: Int32Array | null
	readonly ends: Int32Array | null
}

function normalise(text: string): Normalised {
	const nfc = text.normalize("NFC")

	if (nfc === text) return { text, starts: null, ends: null }

	const starts: number[] = []
	const ends: number[] = []
	let out = ""
	let at = 0

	while (at < text.length) {
		// A cluster: one code point and the combining marks after it.
		let end = at + ((text.codePointAt(at) ?? 0) > 0xffff ? 2 : 1)

		while (end < text.length) {
			const next = text.codePointAt(end) ?? 0

			if (!MARK.test(String.fromCodePoint(next))) break

			end += next > 0xffff ? 2 : 1
		}

		const piece = text.slice(at, end).normalize("NFC")

		for (let k = 0; k < piece.length; k++) {
			starts.push(at)
			ends.push(end)
		}

		out += piece
		at = end
	}

	return { text: out, starts: Int32Array.from(starts), ends: Int32Array.from(ends) }
}

// The page's line that holds offset `at` of `source`, or -1.
export function lineOf(page: FindablePage, source: number, at: number): number {
	for (let line = 0; line < page.lineCount; line++) {
		const i = line * LINE_INTS

		if (page.lineInts[i + 5] === source && at >= (page.lineInts[i + 2] ?? 0) && at < (page.lineInts[i + 3] ?? 0)) return line
	}

	return -1
}

export function findMatches(pages: Iterable<FindablePage | undefined>, query: string, options: DocxFindOptions = {}): Int32Array {
	const out: number[] = []

	if (query.length === 0) return new Int32Array(0)

	const limit = options.limit ?? 10000
	const pattern = new RegExp(query.normalize("NFC").replace(SPECIAL, "\\$&"), options.caseSensitive === true ? "gu" : "giu")

	for (const page of pages) {
		if (page === undefined) continue

		for (let source = 0; source < page.sourceCount; source++) {
			const { text, starts, ends } = normalise(page.sourceTexts[source] ?? "")

			pattern.lastIndex = 0

			for (let match = pattern.exec(text); match !== null; match = pattern.exec(text)) {
				const found = match.index
				const foundEnd = found + match[0].length

				if (foundEnd === found) {
					pattern.lastIndex++

					continue
				}

				// Back to the original text, widened to whole characters.
				const start = starts === null ? found : (starts[found] ?? found)
				const end = ends === null ? foundEnd : (ends[foundEnd - 1] ?? foundEnd)

				if (lineOf(page, source, start) < 0) continue

				out.push(page.index, source, start, end)

				if (out.length / FIND_INTS >= limit) return Int32Array.from(out)
			}
		}
	}

	return Int32Array.from(out)
}
