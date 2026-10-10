import { Namespaces, NodeKind, decodeCharacterData, decodeUtf8, type XmlDocument } from "../../xml/index"
import { CODES_BY_GROUP, El, Ns, URIS_OF_GROUP, groupOf } from "./namespaces"

const UNSET = -1

// The first reader built for a part, so every layer reading it shares one namespace map and one set of
// per-name caches.
const shared = new WeakMap<XmlDocument, PartReader>()

export function partReader(document: XmlDocument): PartReader {
	return shared.get(document) ?? new PartReader(document)
}

function isSpace(b: number | undefined): boolean {
	return b === 32 || b === 9 || b === 10 || b === 13
}

// Element classification and attribute reads for one part. When every namespace is declared on the root
// (Word always does this) an element's code depends only on its interned name, so it is computed once per
// distinct name; otherwise per node. Attributes are matched against every prefix bound to the group in
// scope, scanned in place, and decoded only when asked for.
interface PrefixSet {
	readonly list: readonly string[]
	readonly set: ReadonlySet<string>
}

const NO_PREFIX_SET: PrefixSet = { list: [], set: new Set() }
const XML_PREFIX_SET: PrefixSet = { list: ["xml"], set: new Set(["xml"]) }
// Up to this many aliases are compared byte for byte; more go through a set lookup.
const LINEAR_PREFIXES = 4
// Up to this many attributes are checked for duplicates pairwise; more through a set.
const PAIRWISE_NAMES = 8

// Element classification and attribute reads for one part. When every namespace is declared on the root
// (Word always does this) an element's code depends only on its interned name, so it is computed once per
// distinct name; otherwise per node. An attribute lookup scans the start tag once, matching each
// attribute's local name and then its prefix against every alias bound to the group in scope. Nothing here
// throws on malformed markup: a start tag that cannot be read is recorded in `malformedNodes` and yields
// no attributes, and values with bad references decode as their raw text.
export class PartReader {
	readonly document: XmlDocument
	readonly namespaces: Namespaces
	readonly bytes: Uint8Array
	readonly malformedNodes = new Set<number>()
	private readonly flat: boolean
	private readonly codeByName: Int16Array
	private readonly groupByName: Int8Array
	private readonly prefixCache = new Map<ReadonlyMap<string, string>, PrefixSet[]>()
	private readonly span = new Int32Array(2)
	private swept = false

	constructor(document: XmlDocument, namespaces: Namespaces = new Namespaces(document, { lenient: true })) {
		this.document = document
		this.namespaces = namespaces
		this.bytes = document.bytes

		for (const node of namespaces.malformed) this.malformedNodes.add(node)

		const root = document.documentElement()
		let flat = true

		for (const node of namespaces.declarations().keys()) {
			if (node !== root) {
				flat = false

				break
			}
		}

		this.flat = flat
		this.codeByName = new Int16Array(document.names.length).fill(UNSET)
		this.groupByName = new Int8Array(document.names.length).fill(UNSET)

		if (!shared.has(document)) shared.set(document, this)
	}

	code(node: number): El {
		const document = this.document

		if (document.kinds[node] !== NodeKind.Element) return El.Unknown

		const id = document.nameIds[node] ?? 0

		if (this.flat) {
			const cached = this.codeByName[id] ?? UNSET

			if (cached !== UNSET) return cached as El
		}

		const code = CODES_BY_GROUP[this.group(node)]?.get(document.localName(node)) ?? El.Unknown

		if (this.flat) this.codeByName[id] = code

		return code
	}

	group(node: number): Ns {
		const document = this.document

		if (document.kinds[node] !== NodeKind.Element) return Ns.None

		const id = document.nameIds[node] ?? 0

		if (this.flat) {
			const cached = this.groupByName[id] ?? UNSET

			if (cached !== UNSET) return cached as Ns
		}

		const group = groupOf(this.namespaces.elementUri(node))

		if (this.flat) this.groupByName[id] = group

		return group
	}

	is(node: number, group: Ns, localName: string): boolean {
		return this.group(node) === group && this.document.localName(node) === localName
	}

	// First child element that is {group}localName, or -1.
	child(node: number, group: Ns, localName: string): number {
		const document = this.document

		for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) {
			if (this.is(child, group, localName)) return child
		}

		return -1
	}

	// First descendant element that is {group}localName, in document order, or -1.
	descendant(node: number, group: Ns, localName: string): number {
		const document = this.document
		const end = document.end(node)

		for (let n = node + 1; n < document.count && document.start(n) < end; n++) {
			if (document.kinds[n] === NodeKind.Element && this.is(n, group, localName)) return n
		}

		return -1
	}

	// Whether the element's start tag reads as name="value" pairs with distinct names. A malformed one is
	// recorded.
	wellFormed(node: number): boolean {
		if (this.swept) return !this.malformedNodes.has(node)

		const document = this.document

		if (document.kinds[node] !== NodeKind.Element) return true

		if (malformedTag(this.bytes, document.starts[node] ?? 0, (document.openEnds[node] ?? 0) - 1)) {
			this.malformedNodes.add(node)

			return false
		}

		return true
	}

	// wellFormed over an element and everything below it, reading the tree arrays directly: property
	// containers are swept whole during the build.
	sweep(node: number): void {
		if (this.swept) return

		const { kinds, starts, openEnds, count } = this.document
		const bytes = this.bytes
		const end = this.document.end(node)

		for (let n = node; n < count; n++) {
			const start = starts[n] ?? 0

			if (start >= end) break

			if (kinds[n] === NodeKind.Element && malformedTag(bytes, start, (openEnds[n] ?? 0) - 1)) this.malformedNodes.add(n)
		}
	}

	// Checks every start tag of the part once (for small parts read leniently, like settings and styles);
	// afterwards wellFormed answers from the record. Returns whether anything was malformed.
	sweepAll(): boolean {
		if (!this.swept) {
			const root = this.document.documentElement()

			if (root >= 0) this.sweep(root)

			this.swept = true
		}

		return this.malformedNodes.size > 0
	}

	// Writes the raw value span of the attribute {group}localName into `out`; whether it exists. `group`
	// Ns.None reads an unprefixed attribute.
	find(node: number, group: Ns, localName: string, out: Int32Array): boolean {
		if (this.document.kinds[node] !== NodeKind.Element) return false

		const prefixes = group === Ns.None ? NO_PREFIX_SET : group === Ns.Xml ? XML_PREFIX_SET : this.prefixes(node, group)

		if (group !== Ns.None && prefixes.list.length === 0) return false

		return this.scan(node, group, localName, prefixes, out) === FOUND
	}

	attr(node: number, group: Ns, localName: string): string | undefined {
		const span = this.span

		if (!this.find(node, group, localName, span)) return undefined

		const start = span[0] ?? 0
		const end = span[1] ?? 0

		try {
			return decodeCharacterData(this.bytes, start, end, true)
		} catch {
			this.malformedNodes.add(node)

			return decodeLenient(this.bytes, start, end, true)
		}
	}

	has(node: number, group: Ns, localName: string): boolean {
		return this.find(node, group, localName, this.span)
	}

	// A decimal integer attribute, or `fallback` when absent or not an integer.
	int(node: number, group: Ns, localName: string, fallback: number): number {
		const span = this.span

		return this.find(node, group, localName, span) ? parseInteger(this.bytes, span[0] ?? 0, span[1] ?? 0, fallback) : fallback
	}

	// A number attribute (ST_DecimalNumber, or a decimal fraction some producers write), or `fallback`.
	number(node: number, group: Ns, localName: string, fallback: number): number {
		const span = this.span

		if (!this.find(node, group, localName, span)) return fallback

		const start = span[0] ?? 0
		const end = span[1] ?? 0
		const integer = parseInteger(this.bytes, start, end, Number.NaN)

		if (!Number.isNaN(integer)) return integer

		const text = decodeUtf8(this.bytes, start, end)

		return /^\s*[-+]?[0-9]*\.[0-9]+\s*$/.test(text) ? Number.parseFloat(text) : fallback
	}

	// A length (ST_TwipsMeasure, ST_HpsMeasure...): a bare number is in the target unit, an
	// ST_UniversalMeasure ("0.5in", "12pt") converts with `perPoint` target units per point. Rounded and
	// clamped to int32; `fallback` when absent or unreadable.
	measure(node: number, group: Ns, localName: string, perPoint: number, fallback: number): number {
		const plain = this.number(node, group, localName, Number.NaN)

		if (!Number.isNaN(plain)) return clampInt32(plain)

		const text = this.attr(node, group, localName)
		const measured = text === undefined ? undefined : parseUniversalMeasure(text, perPoint)

		return measured === undefined ? fallback : clampInt32(measured)
	}

	hex(node: number, group: Ns, localName: string, fallback: number): number {
		const span = this.span

		return this.find(node, group, localName, span) ? parseHex(this.bytes, span[0] ?? 0, span[1] ?? 0, fallback) : fallback
	}

	// Whether the attribute's value equals `ascii` exactly, without decoding it.
	equals(node: number, group: Ns, localName: string, ascii: string): boolean {
		const span = this.span

		if (!this.find(node, group, localName, span)) return false

		const start = span[0] ?? 0

		if ((span[1] ?? 0) - start !== ascii.length) return false

		for (let k = 0; k < ascii.length; k++) {
			if (this.bytes[start + k] !== ascii.charCodeAt(k)) return false
		}

		return true
	}

	// ST_OnOff on an element's w:val: absent means on.
	onOff(node: number): boolean {
		const span = this.span

		if (!this.find(node, Ns.W, "val", span)) return true

		return isOn(this.bytes, span[0] ?? 0, span[1] ?? 0)
	}

	// An ST_OnOff attribute: on when present and not "0", "false" or "off".
	flag(node: number, group: Ns, localName: string): boolean {
		const span = this.span

		return this.find(node, group, localName, span) && isOn(this.bytes, span[0] ?? 0, span[1] ?? 0)
	}

	// Decoded character data of an element's text and CDATA children; bad references stay as written.
	text(node: number): string {
		const document = this.document
		let out = ""

		for (let child = document.firstChild(node); child >= 0; child = document.nextSibling(child)) {
			const kind = document.kinds[child]

			if (kind !== NodeKind.Text && kind !== NodeKind.CData) continue

			try {
				out += document.text(child)
			} catch {
				this.malformedNodes.add(node)
				out += decodeLenient(this.bytes, document.start(child), document.end(child), false)
			}
		}

		return out
	}

	// The branch of an mc:AlternateContent to read (ECMA-376 Part 3, §10.2): the first mc:Choice whose
	// Requires prefixes all resolve to namespaces `understands` accepts, else mc:Fallback, else -1.
	selectBranch(node: number, understands: (uri: string) => boolean): number {
		const document = this.document
		let fallback = -1

		for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) {
			if (this.group(child) !== Ns.MC) continue

			const local = document.localName(child)

			if (local === "Choice") {
				const requires = (this.attr(child, Ns.None, "Requires") ?? "").split(/\s+/).filter(prefix => prefix !== "")

				if (
					requires.length > 0 &&
					requires.every(prefix => {
						const uri = this.namespaces.uri(child, prefix)

						return uri !== undefined && understands(uri)
					})
				) {
					return child
				}
			} else if (local === "Fallback" && fallback < 0) {
				fallback = child
			}
		}

		return fallback
	}

	// One pass over a start tag, stopping at the first match.
	private scan(node: number, group: Ns, localName: string, prefixes: PrefixSet, out: Int32Array): number {
		const document = this.document
		const bytes = this.bytes
		const tagEnd = document.openEnd(node) - 1
		let i = document.start(node) + 1

		while (i < tagEnd && !isSpace(bytes[i]) && bytes[i] !== 47) i++

		for (;;) {
			i = skipSpaces(bytes, i, tagEnd)

			if (i >= tagEnd || (bytes[i] === 47 && i === tagEnd - 1)) return NOT_FOUND

			const nameStart = i
			let colon = -1

			while (i < tagEnd && !isSpace(bytes[i]) && bytes[i] !== 61 && bytes[i] !== 47) {
				if (bytes[i] === 58 && colon < 0) colon = i
				i++
			}

			const nameEnd = i

			i = skipSpaces(bytes, i, tagEnd)

			if (nameEnd === nameStart || bytes[i] !== 61) return this.malformed(node)

			i = skipSpaces(bytes, i + 1, tagEnd)

			const quote = bytes[i]

			if (quote !== 34 && quote !== 39) return this.malformed(node)

			const valueStart = i + 1
			const valueEnd = bytes.indexOf(quote, valueStart)

			if (valueEnd < 0 || valueEnd >= tagEnd) return this.malformed(node)

			const matches =
				group === Ns.None
					? colon < 0 && sameAscii(bytes, nameStart, nameEnd, localName)
					: colon > nameStart && sameAscii(bytes, colon + 1, nameEnd, localName) && this.prefixMatches(nameStart, colon, prefixes)

			if (matches) {
				out[0] = valueStart
				out[1] = valueEnd

				return FOUND
			}

			i = valueEnd + 1
		}
	}

	private malformed(node: number): number {
		this.malformedNodes.add(node)

		return MALFORMED
	}

	private prefixMatches(start: number, end: number, prefixes: PrefixSet): boolean {
		if (prefixes.list.length <= LINEAR_PREFIXES) {
			for (const prefix of prefixes.list) {
				if (sameAscii(this.bytes, start, end, prefix)) return true
			}

			return false
		}

		return prefixes.set.has(decodeUtf8(this.bytes, start, end))
	}

	private prefixes(node: number, group: Ns): PrefixSet {
		const scope = this.namespaces.inScope(node)
		let byGroup = this.prefixCache.get(scope)

		if (byGroup === undefined) {
			byGroup = []
			this.prefixCache.set(scope, byGroup)
		}

		let prefixes = byGroup[group]

		if (prefixes === undefined) {
			const uris = URIS_OF_GROUP[group] ?? []
			const found: string[] = []

			for (const [prefix, uri] of scope) {
				if (prefix !== "" && uri !== "" && uris.includes(uri)) found.push(prefix)
			}

			prefixes = found.length === 0 ? NO_PREFIX_SET : { list: found, set: new Set(found) }
			byGroup[group] = prefixes
		}

		return prefixes
	}
}

// The first node after `node`'s subtree, staying inside `root` (the node count when nothing follows).
export function afterSubtree(document: XmlDocument, node: number, root: number): number {
	for (let n = node; n !== root && n >= 0; n = document.parent(n)) {
		const next = document.nextSibling(n)

		if (next >= 0) return next
	}

	return document.count
}

// Character data with every valid reference resolved and invalid ones kept as written.
function decodeLenient(bytes: Uint8Array, start: number, end: number, attribute: boolean): string {
	let out = ""
	let from = start

	for (let amp = bytes.indexOf(38, start); amp >= 0 && amp < end; amp = bytes.indexOf(38, from)) {
		const semicolon = bytes.indexOf(59, amp)
		const referenceEnd = semicolon >= 0 && semicolon < end && semicolon - amp <= 12 ? semicolon + 1 : amp + 1

		out += decodeCharacterData(bytes, from, amp, attribute)

		try {
			out += decodeCharacterData(bytes, amp, referenceEnd, attribute)
		} catch {
			out += decodeUtf8(bytes, amp, referenceEnd)
		}

		from = referenceEnd
	}

	return out + decodeCharacterData(bytes, from, end, attribute)
}

let nameSpans = new Int32Array(32)

// Whether a start tag ("<" at `start`, ">" at `tagEnd`) fails to read as name="value" pairs with distinct
// names. Runs over every property element during the build, so it is a flat byte loop.
function malformedTag(bytes: Uint8Array, start: number, tagEnd: number): boolean {
	let i = start + 1

	while (i < tagEnd) {
		const b = bytes[i]

		if (b === 32 || b === 9 || b === 10 || b === 13 || b === 47) break

		i++
	}

	let names = 0
	let spans = nameSpans

	for (;;) {
		while (i < tagEnd) {
			const b = bytes[i]

			if (b !== 32 && b !== 9 && b !== 10 && b !== 13) break

			i++
		}

		if (i >= tagEnd || (bytes[i] === 47 && i === tagEnd - 1)) return names > 1 && duplicateNames(bytes, spans, names)

		const nameStart = i

		while (i < tagEnd) {
			const b = bytes[i]

			if (b === 61 || b === 47 || b === 32 || b === 9 || b === 10 || b === 13) break

			i++
		}

		const nameEnd = i

		while (i < tagEnd) {
			const b = bytes[i]

			if (b !== 32 && b !== 9 && b !== 10 && b !== 13) break

			i++
		}

		if (nameEnd === nameStart || bytes[i] !== 61) return true

		i++

		while (i < tagEnd) {
			const b = bytes[i]

			if (b !== 32 && b !== 9 && b !== 10 && b !== 13) break

			i++
		}

		const quote = bytes[i]

		if (quote !== 34 && quote !== 39) return true

		// Values are short: a plain loop beats indexOf's call overhead.
		i++

		while (i < tagEnd && bytes[i] !== quote) i++

		if (i >= tagEnd) return true

		if (names * 2 + 2 > spans.length) {
			const grown = new Int32Array(spans.length * 2)

			grown.set(spans)
			spans = grown
			nameSpans = grown
		}

		spans[names * 2] = nameStart
		spans[names * 2 + 1] = nameEnd
		names++
		i++
	}
}

function duplicateNames(bytes: Uint8Array, spans: Int32Array, count: number): boolean {
	if (count > PAIRWISE_NAMES) {
		const seen = new Set<string>()

		for (let k = 0; k < count; k++) {
			const name = decodeUtf8(bytes, spans[k * 2] ?? 0, spans[k * 2 + 1] ?? 0)

			if (seen.has(name)) return true

			seen.add(name)
		}

		return false
	}

	for (let a = 1; a < count; a++) {
		const aStart = spans[a * 2] ?? 0
		const length = (spans[a * 2 + 1] ?? 0) - aStart

		for (let b = 0; b < a; b++) {
			const bStart = spans[b * 2] ?? 0

			if (length !== (spans[b * 2 + 1] ?? 0) - bStart || bytes[aStart] !== bytes[bStart]) continue

			let same = true

			for (let k = 1; k < length && same; k++) same = bytes[aStart + k] === bytes[bStart + k]

			if (same) return true
		}
	}

	return false
}

const FOUND = 1
const NOT_FOUND = 0
const MALFORMED = -1

function skipSpaces(bytes: Uint8Array, from: number, end: number): number {
	let i = from

	while (i < end && isSpace(bytes[i])) i++

	return i
}

// Byte comparison against a string; non-ASCII strings compare through decoding.
function sameAscii(bytes: Uint8Array, start: number, end: number, text: string): boolean {
	if (end - start !== text.length) return end - start > text.length && decodeUtf8(bytes, start, end) === text

	for (let k = 0; k < text.length; k++) {
		const code = text.charCodeAt(k)

		if (code >= 0x80) return decodeUtf8(bytes, start, end) === text
		if (bytes[start + k] !== code) return false
	}

	return true
}

// Twips per unit of an ST_UniversalMeasure suffix.
const TWIPS_PER_UNIT: Readonly<Record<string, number>> = {
	mm: 1440 / 25.4,
	cm: 1440 / 2.54,
	in: 1440,
	pt: 20,
	pc: 240,
	pi: 240
}

export function parseUniversalMeasure(text: string, perPoint: number): number | undefined {
	const match = /^\s*(-?[0-9]*\.?[0-9]+)\s*(mm|cm|in|pt|pc|pi)\s*$/.exec(text)

	if (match === null) return undefined

	const value = Number.parseFloat(match[1] ?? "")
	const twipsPerUnit = TWIPS_PER_UNIT[match[2] ?? ""]

	if (!Number.isFinite(value) || twipsPerUnit === undefined) return undefined

	return Math.round((value * twipsPerUnit * perPoint) / 20)
}

const INT32_MIN = -2147483648
const INT32_MAX = 2147483647

// OOXML's decimal types live in int32: clamping keeps hostile values out of later arithmetic.
export function clampInt32(value: number): number {
	const rounded = Math.round(value)

	return rounded < INT32_MIN ? INT32_MIN : rounded > INT32_MAX ? INT32_MAX : rounded
}

export function parseInteger(bytes: Uint8Array, start: number, end: number, fallback: number): number {
	while (start < end && isSpace(bytes[start])) start++
	while (end > start && isSpace(bytes[end - 1])) end--

	let negative = false

	if (bytes[start] === 45 || bytes[start] === 43) {
		negative = bytes[start] === 45
		start++
	}

	if (start >= end) return fallback

	let value = 0

	for (let i = start; i < end; i++) {
		const b = (bytes[i] ?? 0) - 48

		if (b < 0 || b > 9) return fallback

		value = value * 10 + b
	}

	if (!Number.isSafeInteger(value)) return fallback

	return negative ? -value : value
}

export function parseHex(bytes: Uint8Array, start: number, end: number, fallback: number): number {
	while (start < end && isSpace(bytes[start])) start++
	while (end > start && isSpace(bytes[end - 1])) end--

	if (start >= end || end - start > 8) return fallback

	let value = 0

	for (let i = start; i < end; i++) {
		const b = bytes[i] ?? 0
		let digit: number

		if (b >= 48 && b <= 57) {
			digit = b - 48
		} else if (b >= 97 && b <= 102) {
			digit = b - 87
		} else if (b >= 65 && b <= 70) {
			digit = b - 55
		} else {
			return fallback
		}

		value = value * 16 + digit
	}

	return value
}

// ST_OnOff values: "0", "false" and "off" are off, anything else is on.
function isOn(bytes: Uint8Array, start: number, end: number): boolean {
	const length = end - start
	const first = bytes[start]

	if (length === 1) return first !== 48
	if (length === 3 && first === 111 && bytes[start + 1] === 102 && bytes[start + 2] === 102) return false

	return !(length === 5 && first === 102 && bytes[start + 1] === 97 && bytes[start + 2] === 108 && bytes[start + 3] === 115)
}
