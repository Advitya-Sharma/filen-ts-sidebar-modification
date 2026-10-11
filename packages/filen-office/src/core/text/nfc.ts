// Text in Unicode NFC with a map back to the original, for searches that must not care how a document composed
// its characters.

const MARK = /\p{M}/u

// `text` in NFC, with each normalised code unit's cluster in the original (start, end): a match maps back to
// whole characters however the document composed them. Null maps when the text is already NFC.
export interface NfcText {
	readonly text: string
	readonly starts: Int32Array | null
	readonly ends: Int32Array | null
}

export function nfcWithMap(text: string): NfcText {
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
