// The RTF a writer produces: control words, text already encoded (encode.ts) and source bytes copied as they
// were, kept as chunks (copied bytes are views of the source, never copies until the file is joined). A
// control word that could run on into what follows (a letter, digit, space or minus) gets its delimiting space.

export class RtfWriteRefusal extends Error {
	// What the writer cannot represent: a WordprocessingML element or attribute, or the edit's kind.
	readonly reason: string

	constructor(reason: string, message: string = reason) {
		super(message)
		this.name = "RtfWriteRefusal"
		this.reason = reason
	}
}

function isLetter(code: number): boolean {
	return (code >= 0x61 && code <= 0x7a) || (code >= 0x41 && code <= 0x5a)
}

function isDigit(code: number): boolean {
	return code >= 0x30 && code <= 0x39
}

// Whether `at(length - 1)` ends an open control word (\name or \nameN with no delimiter after it).
function endsInWord(length: number, at: (index: number) => number): boolean {
	let k = length - 1

	while (k >= 0 && isDigit(at(k))) k--

	if (k >= 0 && at(k) === 0x2d) k--

	const lettersEnd = k

	while (k >= 0 && isLetter(at(k))) k--

	if (k === lettersEnd || k < 0 || at(k) !== 0x5c) return false

	// The backslash is not itself escaped.
	let slashes = 0

	for (let j = k; j >= 0 && at(j) === 0x5c; j--) slashes++

	return slashes % 2 === 1
}

function needsDelimiter(code: number): boolean {
	return isLetter(code) || isDigit(code) || code === 0x20 || code === 0x2d
}

export class RtfOut {
	private readonly chunks: Uint8Array[] = []
	private pending = ""
	private open = false
	length = 0

	// A control word (`name` letters, `param` its parameter).
	word(name: string, param?: number): this {
		this.put(`\\${name}${param === undefined ? "" : String(Math.trunc(param))}`)
		this.open = true

		return this
	}

	// RTF that ends no control word open: braces, encoded text, \* and the like.
	rtf(text: string): this {
		if (text === "") return this

		this.put(text)
		this.open = false

		return this
	}

	// Source bytes as they were.
	bytes(bytes: Uint8Array): this {
		if (bytes.length === 0) return this

		this.flush(bytes[0] ?? 0)
		this.chunks.push(bytes)
		this.length += bytes.length
		this.open = endsInWord(bytes.length, index => bytes[index] ?? 0)

		return this
	}

	// Everything written so far, as chunks.
	finish(): Uint8Array[] {
		this.flush(-1)

		return this.chunks
	}

	// Whether the last thing written is a control word without its delimiter.
	get wordOpen(): boolean {
		return this.open
	}

	private put(text: string): void {
		const first = text.charCodeAt(0)

		if (this.open && needsDelimiter(first)) this.pending += " "

		this.pending += text
		this.open = false
	}

	private flush(next: number): void {
		if (this.open && next >= 0 && needsDelimiter(next)) {
			this.pending += " "
			this.open = false
		}

		if (this.pending === "") return

		const out = new Uint8Array(this.pending.length)

		for (let k = 0; k < this.pending.length; k++) out[k] = this.pending.charCodeAt(k) & 0xff

		this.chunks.push(out)
		this.length += out.length
		this.pending = ""
	}
}

export function joinChunks(chunks: readonly Uint8Array[]): Uint8Array {
	let length = 0

	for (const chunk of chunks) length += chunk.length

	const out = new Uint8Array(length)
	let at = 0

	for (const chunk of chunks) {
		out.set(chunk, at)
		at += chunk.length
	}

	return out
}
