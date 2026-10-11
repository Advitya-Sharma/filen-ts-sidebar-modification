import { NodeKind, XmlDocument, type XmlEncoding } from "./document"
import { XmlError } from "./errors"
import { decodeUtf8, transcodeUtf16 } from "./utf8"

const LT = 60
const GT = 62
const SLASH = 47
const BANG = 33
const QUESTION = 63
const DQUOTE = 34
const SQUOTE = 39

function isSpace(b: number | undefined): boolean {
	return b === 32 || b === 9 || b === 10 || b === 13
}

function startsWith(bytes: Uint8Array, at: number, ascii: string): boolean {
	for (let k = 0; k < ascii.length; k++) {
		if (bytes[at + k] !== ascii.charCodeAt(k)) return false
	}

	return true
}

// Index of `ascii` at or after `from`, or -1.
function indexOfSequence(bytes: Uint8Array, ascii: string, from: number): number {
	const first = ascii.charCodeAt(0)

	for (let i = bytes.indexOf(first, from); i >= 0; i = bytes.indexOf(first, i + 1)) {
		if (startsWith(bytes, i, ascii)) return i
	}

	return -1
}

function isWhitespaceOnly(bytes: Uint8Array, start: number, end: number): boolean {
	for (let i = start; i < end; i++) {
		if (!isSpace(bytes[i])) return false
	}

	return true
}

function hashName(bytes: Uint8Array, start: number, end: number): number {
	let hash = 0

	for (let i = start; i < end; i++) hash = (Math.imul(hash, 31) + (bytes[i] ?? 0)) | 0

	return hash
}

// Names interned by their bytes: hashed, then compared byte for byte, so a hash collision never merges two
// names. Most parts use a few dozen distinct names, so the table stays tiny.
class NameTable {
	readonly names: string[] = []
	readonly lengths: number[] = []
	private readonly starts: number[] = []
	private readonly chain: number[] = []
	private readonly buckets = new Map<number, number>()
	private readonly bytes: Uint8Array

	constructor(bytes: Uint8Array) {
		this.bytes = bytes
	}

	// `hash` is hashName over the same bytes, computed by the caller while it scanned them.
	intern(start: number, end: number, hash: number = hashName(this.bytes, start, end)): number {
		const head = this.buckets.get(hash)

		for (let id = head ?? -1; id >= 0; id = this.chain[id] ?? -1) {
			if (this.matches(id, start, end)) return id
		}

		const id = this.names.length

		this.names.push(decodeUtf8(this.bytes, start, end))
		this.lengths.push(end - start)
		this.starts.push(start)
		this.chain.push(head ?? -1)
		this.buckets.set(hash, id)

		return id
	}

	matches(id: number, start: number, end: number): boolean {
		const length = this.lengths[id] ?? -1

		if (length !== end - start) return false

		const bytes = this.bytes
		const offset = (this.starts[id] ?? 0) - start

		for (let k = start; k < end; k++) {
			if (bytes[k] !== bytes[k + offset]) return false
		}

		return true
	}
}

function detectEncoding(bytes: Uint8Array): XmlEncoding {
	const b0 = bytes[0]
	const b1 = bytes[1]

	if ((b0 === 0xff && b1 === 0xfe) || (b0 === 0x3c && b1 === 0)) return "utf-16le"
	if ((b0 === 0xfe && b1 === 0xff) || (b0 === 0 && b1 === 0x3c)) return "utf-16be"
	if (bytes.length >= 2 && (b0 === 0 || b1 === 0)) throw new XmlError("encoding", "Neither a UTF-8 nor a UTF-16 part", 0)

	return "utf-8"
}

// Start-tag attribute name spans, reused across tags.
let attributeNames = new Int32Array(32)

function sameName(bytes: Uint8Array, spans: Int32Array, a: number, b: number): boolean {
	const aStart = spans[a * 2] ?? 0
	const bStart = spans[b * 2] ?? 0
	const length = (spans[a * 2 + 1] ?? 0) - aStart

	if (length !== (spans[b * 2 + 1] ?? 0) - bStart) return false

	for (let k = 0; k < length; k++) if (bytes[aStart + k] !== bytes[bStart + k]) return false

	return true
}

// Whether two of a start tag's attribute names are the same bytes. Pairwise for the few attributes real tags
// carry; hashed above that, so a hostile tag with thousands of attributes stays linear.
function repeatsName(bytes: Uint8Array, spans: Int32Array, count: number): boolean {
	if (count > 8) {
		const seen = new Map<number, number[]>()

		for (let a = 0; a < count; a++) {
			const hash = hashName(bytes, spans[a * 2] ?? 0, spans[a * 2 + 1] ?? 0)
			const bucket = seen.get(hash)

			if (bucket === undefined) {
				seen.set(hash, [a])

				continue
			}

			for (const b of bucket) if (sameName(bytes, spans, a, b)) return true

			bucket.push(a)
		}

		return false
	}

	for (let a = 1; a < count; a++) {
		const aStart = spans[a * 2] ?? 0
		const length = (spans[a * 2 + 1] ?? 0) - aStart

		for (let b = 0; b < a; b++) {
			const bStart = spans[b * 2] ?? 0

			if (length !== (spans[b * 2 + 1] ?? 0) - bStart || bytes[aStart + length - 1] !== bytes[bStart + length - 1]) continue

			let same = true

			for (let k = 0; k < length && same; k++) same = bytes[aStart + k] === bytes[bStart + k]

			if (same) return true
		}
	}

	return false
}

// Parses one part into an XmlDocument. Structure is checked strictly (every tag closed by its own name, one
// root element, nothing but whitespace, comments and processing instructions around it); a start tag whose
// attributes do not read as name="value" pairs is recorded (XmlDocument.isMalformed) rather than refused;
// names, characters and references are checked by assertWellFormed, as only written parts need them.
// A DOCTYPE with an internal subset is refused, so no entity can ever expand; an external one is kept as an
// opaque node and never read. UTF-16 parts are transcoded: the document's bytes are then UTF-8.
export function parseXml(input: Uint8Array): XmlDocument {
	const encoding = detectEncoding(input)

	if (encoding !== "utf-8" && input.length % 2 !== 0) throw new XmlError("encoding", "Truncated UTF-16 part", input.length - 1)

	const bom = encoding !== "utf-8" && (input[0] === 0xff || input[0] === 0xfe) ? 2 : 0
	const bytes = encoding === "utf-8" ? input : transcodeUtf16(input, bom, encoding === "utf-16le")
	const length = bytes.length
	const prologStart = length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0

	// OOXML averages 35-40 bytes per node; starting near that grows the arrays once or twice at most.
	let capacity = (length >> 5) + 64
	let kinds = new Uint8Array(capacity)
	let starts = new Int32Array(capacity)
	let openEnds = new Int32Array(capacity)
	let ends = new Int32Array(capacity)
	let parents = new Int32Array(capacity)
	let firstChildren = new Int32Array(capacity)
	let nextSiblings = new Int32Array(capacity)
	let nameIds = new Int32Array(capacity)
	const names = new NameTable(bytes)
	// Last child per open depth, so siblings link without a per-node lastChild array.
	let lastChild = new Int32Array(64)
	const malformed: number[] = []
	let depth = 0
	let count = 0

	const grow = (): void => {
		capacity = Math.ceil(capacity * 1.5)

		const widen = <T extends Uint8Array | Int32Array>(array: T, make: (size: number) => T): T => {
			const next = make(capacity)

			next.set(array)

			return next
		}

		kinds = widen(kinds, size => new Uint8Array(size))
		starts = widen(starts, size => new Int32Array(size))
		openEnds = widen(openEnds, size => new Int32Array(size))
		ends = widen(ends, size => new Int32Array(size))
		parents = widen(parents, size => new Int32Array(size))
		firstChildren = widen(firstChildren, size => new Int32Array(size))
		nextSiblings = widen(nextSiblings, size => new Int32Array(size))
		nameIds = widen(nameIds, size => new Int32Array(size))
	}

	const add = (kind: number, start: number, end: number, parent: number, nameId: number): number => {
		if (count === capacity) grow()

		const node = count++

		kinds[node] = kind
		starts[node] = start
		openEnds[node] = end
		ends[node] = end
		parents[node] = parent
		firstChildren[node] = -1
		nextSiblings[node] = -1
		nameIds[node] = nameId

		if (parent >= 0) {
			const previous = lastChild[depth] ?? -1

			if (previous < 0) {
				firstChildren[parent] = node
			} else {
				nextSiblings[previous] = node
			}

			lastChild[depth] = node
		}

		return node
	}

	add(NodeKind.Document, 0, length, -1, -1)
	lastChild[0] = -1

	let current = 0
	let rootElements = 0
	let doctype = false
	let i = prologStart

	while (i < length) {
		if (bytes[i] !== LT) {
			const lt = bytes.indexOf(LT, i)
			const end = lt < 0 ? length : lt

			if (current === 0 && !isWhitespaceOnly(bytes, i, end)) {
				throw new XmlError("contentOutsideRoot", "Text outside the root element", i)
			}

			add(NodeKind.Text, i, end, current, -1)
			i = end

			continue
		}

		const next = bytes[i + 1]

		if (next === SLASH) {
			if (current === 0) throw new XmlError("unexpectedCloseTag", "Close tag without an open element", i)

			// The open element's name sits right after its "<": compare in place, then expect ">".
			const openName = (starts[current] ?? 0) + 1
			const nameLength = names.lengths[nameIds[current] ?? 0] ?? 0
			let gt = i + 2 + nameLength
			let same = gt < length

			for (let k = 0; k < nameLength && same; k++) same = bytes[i + 2 + k] === bytes[openName + k]

			while (gt < length && isSpace(bytes[gt])) gt++

			if (gt >= length) throw new XmlError("unterminated", "Unterminated close tag", i)

			if (!same || bytes[gt] !== GT) {
				throw new XmlError("mismatchedTag", `Close tag does not match <${names.names[nameIds[current] ?? -1] ?? ""}>`, i)
			}

			ends[current] = gt + 1
			current = parents[current] ?? 0
			depth--
			i = gt + 1

			continue
		}

		if (next === BANG) {
			if (startsWith(bytes, i, "<!--")) {
				const close = indexOfSequence(bytes, "-->", i + 4)

				if (close < 0) throw new XmlError("unterminated", "Unterminated comment", i)

				add(NodeKind.Comment, i, close + 3, current, -1)
				i = close + 3
			} else if (startsWith(bytes, i, "<![CDATA[")) {
				if (current === 0) throw new XmlError("contentOutsideRoot", "CDATA outside the root element", i)

				const close = indexOfSequence(bytes, "]]>", i + 9)

				if (close < 0) throw new XmlError("unterminated", "Unterminated CDATA section", i)

				add(NodeKind.CData, i, close + 3, current, -1)
				i = close + 3
			} else if (startsWith(bytes, i, "<!DOCTYPE")) {
				if (current !== 0 || rootElements > 0 || doctype) throw new XmlError("doctype", "Misplaced document type declaration", i)

				let j = i + 9

				while (j < length && bytes[j] !== GT) {
					const b = bytes[j]

					if (b === 91) throw new XmlError("doctype", "Internal DTD subsets are not allowed", j)
					if (b === LT) throw new XmlError("doctype", "Unterminated document type declaration", j)

					if (b === DQUOTE || b === SQUOTE) {
						const close = bytes.indexOf(b, j + 1)

						j = close < 0 ? length : close + 1
					} else {
						j++
					}
				}

				if (j >= length) throw new XmlError("unterminated", "Unterminated document type declaration", i)

				add(NodeKind.DocumentType, i, j + 1, current, -1)
				doctype = true
				i = j + 1
			} else {
				throw new XmlError("invalidName", "Unknown markup declaration", i)
			}

			continue
		}

		if (next === QUESTION) {
			const close = indexOfSequence(bytes, "?>", i + 2)

			if (close < 0) throw new XmlError("unterminated", "Unterminated processing instruction", i)

			let targetEnd = i + 2

			while (targetEnd < close && !isSpace(bytes[targetEnd])) targetEnd++

			if (targetEnd === i + 2) throw new XmlError("invalidName", "Processing instruction without a target", i)

			const target = names.intern(i + 2, targetEnd)

			// "<?xml" is the declaration, which may only open the part.
			if (i !== prologStart && targetEnd - i === 5 && (names.names[target] ?? "").toLowerCase() === "xml") {
				throw new XmlError("invalidName", "XML declaration not at the start", i)
			}

			add(NodeKind.ProcessingInstruction, i, close + 2, current, target)
			i = close + 2

			continue
		}

		let j = i + 1
		let hash = 0

		while (j < length) {
			const b = bytes[j] ?? 0

			if (b === GT || b === SLASH || b === 32 || b === 10 || b === 9 || b === 13) break

			hash = (Math.imul(hash, 31) + b) | 0
			j++
		}

		if (j === i + 1 || j >= length) {
			throw new XmlError(j >= length ? "unterminated" : "invalidName", "Invalid start tag", i)
		}

		const nameId = names.intern(i + 1, j, hash)
		// Attributes are checked on the way to the tag's end: whitespace, name, "=", quoted value, repeated.
		// The first irregularity marks the element malformed and hands the rest to the permissive scan below,
		// which only finds the end.
		let malformedTag = false
		let attributes = 0

		for (;;) {
			const separated = isSpace(bytes[j])

			while (j < length && isSpace(bytes[j])) j++

			const b = bytes[j]

			if (b === GT || (b === SLASH && bytes[j + 1] === GT)) break

			if (!separated || b === SLASH || j >= length) {
				malformedTag = j < length
				break
			}

			const nameStart = j

			while (j < length) {
				const c = bytes[j]

				if (c === 61 || c === GT || c === SLASH || c === LT || c === 32 || c === 10 || c === 9 || c === 13) break

				j++
			}

			const nameEnd = j

			while (j < length && isSpace(bytes[j])) j++

			if (bytes[j] !== 61) {
				malformedTag = true
				break
			}

			j++

			while (j < length && isSpace(bytes[j])) j++

			const quote = bytes[j]

			if (quote !== DQUOTE && quote !== SQUOTE) {
				malformedTag = true
				break
			}

			const close = bytes.indexOf(quote, j + 1)

			if (close < 0) throw new XmlError("unterminated", "Unterminated attribute value", j)

			if (attributes * 2 + 2 > attributeNames.length) {
				const grown = new Int32Array(attributeNames.length * 2)

				grown.set(attributeNames)
				attributeNames = grown
			}

			attributeNames[attributes * 2] = nameStart
			attributeNames[attributes * 2 + 1] = nameEnd
			attributes++
			j = close + 1
		}

		if (!malformedTag && attributes > 1) malformedTag = repeatsName(bytes, attributeNames, attributes)

		// The ">" ending the start tag is the first one outside a quoted value. A "<" inside a quoted value is
		// left to attribute decoding; one outside means the tag never ended.
		while (j < length) {
			const b = bytes[j]

			if (b === GT) break
			if (b === LT) throw new XmlError("invalidAttribute", '"<" inside a start tag', j)

			if (b === DQUOTE || b === SQUOTE) {
				const close = bytes.indexOf(b, j + 1)

				if (close < 0) throw new XmlError("unterminated", "Unterminated attribute value", j)

				j = close + 1
			} else {
				j++
			}
		}

		if (j >= length) throw new XmlError("unterminated", "Unterminated start tag", i)

		if (current === 0 && ++rootElements > 1) {
			throw new XmlError("contentOutsideRoot", "More than one root element", i)
		}

		const element = add(NodeKind.Element, i, j + 1, current, nameId)

		if (malformedTag) malformed.push(element)

		i = j + 1

		if (bytes[j - 1] !== SLASH) {
			current = element
			depth++

			if (depth === lastChild.length) {
				const deeper = new Int32Array(depth * 2)

				deeper.set(lastChild)
				lastChild = deeper
			}

			lastChild[depth] = -1
		}
	}

	if (current !== 0) {
		throw new XmlError(
			"unclosedElement",
			`Element <${names.names[nameIds[current] ?? -1] ?? ""}> is never closed`,
			starts[current] ?? 0
		)
	}

	if (rootElements === 0) {
		throw new XmlError("contentOutsideRoot", "No root element", prologStart)
	}

	// The first guess and growth leave slack; trees live as long as the document, so copy it away.
	if (capacity - count > count >> 5) {
		kinds = kinds.slice(0, count)
		starts = starts.slice(0, count)
		openEnds = openEnds.slice(0, count)
		ends = ends.slice(0, count)
		parents = parents.slice(0, count)
		firstChildren = firstChildren.slice(0, count)
		nextSiblings = nextSiblings.slice(0, count)
		nameIds = nameIds.slice(0, count)
	}

	return new XmlDocument(bytes, {
		count,
		kinds,
		starts,
		openEnds,
		ends,
		parents,
		firstChildren,
		nextSiblings,
		nameIds,
		names: names.names,
		nameLengths: names.lengths,
		encoding,
		malformed: malformed.length === 0 ? undefined : Int32Array.from(malformed)
	})
}
