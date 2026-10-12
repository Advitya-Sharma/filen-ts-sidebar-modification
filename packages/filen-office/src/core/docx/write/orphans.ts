import type { BytePieces, SliceClock } from "../../opc/index"
import { decodeCharacterData, encodeUtf8, type XmlDocument } from "../../xml/index"
import type { StoryPart } from "../model/index"
import { partReader } from "../model/reader"
import type { DocxPackage } from "../package"

const R_URIS = [
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
	"http://purl.oclc.org/ooxml/officeDocument/relationships"
]
const HYPERLINK_TYPES = new Set([
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
	"http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink"
])
// Bytes scanned between looks at the clock.
const SCAN_STEP = 4 * 1024 * 1024
// Bytes of markup scanned again at the start of the next window (an attribute split by a window's end).
const SEAM = 512

function isSpace(byte: number | undefined): boolean {
	return byte === 32 || byte === 9 || byte === 10 || byte === 13
}

// Calls `visit` with the decoded value of every attribute of the relationships namespace (one of `prefixes`,
// each "prefix:" as bytes) whose name starts in `bytes` [from, to). Raw bytes, no parse: one tight pass finds
// a whitespace byte followed by a prefix, and a false match only keeps a relationship.
function eachRelationshipValue(
	bytes: Uint8Array,
	from: number,
	to: number,
	prefixes: readonly Uint8Array[],
	visit: (value: string, pos: number) => void
): void {
	const firsts = new Uint8Array(256)

	for (const prefix of prefixes) firsts[prefix[0] ?? 0] = 1

	const limit = Math.min(to, bytes.length)

	for (let start = from; start < limit; start++) {
		if (firsts[bytes[start] ?? 0] === 0 || (start > 0 && !isSpace(bytes[start - 1]))) continue

		let prefix: Uint8Array | undefined

		for (const candidate of prefixes) {
			let k = 0

			while (k < candidate.length && bytes[start + k] === candidate[k]) k++

			if (k === candidate.length) {
				prefix = candidate

				break
			}
		}

		if (prefix === undefined) continue

		let i = start + prefix.length

		while (i < bytes.length && bytes[i] !== 0x3d && !isSpace(bytes[i]) && bytes[i] !== 0x3e && bytes[i] !== 0x2f) i++
		while (isSpace(bytes[i])) i++

		if (bytes[i] !== 0x3d) continue

		i++

		while (isSpace(bytes[i])) i++

		const quote = bytes[i]

		if (quote !== 0x22 && quote !== 0x27) continue

		const end = bytes.indexOf(quote, i + 1)

		if (end >= 0) visit(decodeCharacterData(bytes, i + 1, end, true), start)
	}
}

interface Use {
	readonly pos: number
	readonly value: string
}

// The relationship ids a source part uses and where, read once per part (its source never changes).
const usesBySource = new WeakMap<XmlDocument, readonly Use[]>()

async function sourceUses(document: XmlDocument, prefixes: readonly Uint8Array[], pause: () => Promise<void>): Promise<readonly Use[]> {
	const known = usesBySource.get(document)

	if (known !== undefined) return known

	const source = document.bytes
	const uses: Use[] = []

	for (let at = 0; at < source.length; at += SCAN_STEP) {
		eachRelationshipValue(source, at, Math.min(source.length, at + SCAN_STEP), prefixes, (value, pos) => uses.push({ pos, value }))
		await pause()
	}

	usesBySource.set(document, uses)

	return uses
}

// The "prefix:" names that may stand for the relationships namespace in `document` and in markup written for
// it.
function relationshipPrefixes(document: XmlDocument): Uint8Array[] {
	const out = new Set<string>(["r"])

	for (const bindings of partReader(document).namespaces.declarations().values()) {
		for (const [prefix, uri] of bindings) if (prefix !== "" && R_URIS.includes(uri)) out.add(prefix)
	}

	return [...out].map(prefix => encodeUtf8(`${prefix}:`))
}

// External hyperlink relationships of `part` that source markup the save rewrote used and its new content
// (`bytes`: views of the source and new markup) does not: an edit took out the last link to them, and Word
// drops such relationships on its next save. One the source did not use is never among them. Uses are
// attributes of the relationships namespace, values decoded (an id written with character references counts).
export async function orphanedHyperlinks(
	pkg: DocxPackage,
	part: StoryPart,
	bytes: BytePieces,
	clock: SliceClock | undefined
): Promise<Set<string>> {
	const candidates = new Set(
		(await pkg.relationships(part.name)).all
			.filter(relationship => relationship.external && HYPERLINK_TYPES.has(relationship.type))
			.map(relationship => relationship.id)
	)

	if (candidates.size === 0) return candidates

	const document = part.document
	const source = document.bytes
	// The new content as source ranges (views) and new markup.
	const views: [number, number][] = []
	const markup: Uint8Array[] = []

	bytes.visit(0, bytes.length, (piece, from, count) => {
		const at = piece.byteOffset - source.byteOffset + from

		if (piece.buffer === source.buffer && at >= 0 && at + count <= source.length) views.push([at, at + count])
		else markup.push(piece.subarray(from, from + count))

		return true
	})

	const prefixes = relationshipPrefixes(document)
	const pause = async (): Promise<void> => {
		if (clock?.due === true) await clock.pause()
	}

	// The source's uses of candidates, one pass: those inside a view stay; the others were in what the save
	// rewrote.
	const merged: [number, number][] = []

	for (const [from, to] of views.sort((a, b) => a[0] - b[0])) {
		const last = merged[merged.length - 1]

		if (last !== undefined && from <= last[1]) last[1] = Math.max(last[1], to)
		else merged.push([from, to])
	}

	const inView = (pos: number): boolean => {
		let low = 0
		let high = merged.length

		while (low < high) {
			const middle = (low + high) >> 1

			if ((merged[middle]?.[1] ?? 0) <= pos) low = middle + 1
			else high = middle
		}

		return low < merged.length && (merged[low]?.[0] ?? 0) <= pos
	}
	const missing = new Set<string>()
	const kept = new Set<string>()

	for (const use of await sourceUses(document, prefixes, pause))
		if (candidates.has(use.value)) (inView(use.pos) ? kept : missing).add(use.value)

	for (const id of kept) missing.delete(id)

	// What the new markup still uses: its pieces joined into windows (an attribute may straddle two), each
	// window starting with the end of the one before.
	const window = new Uint8Array(SCAN_STEP + SEAM)
	let length = 0
	const flush = async (): Promise<void> => {
		eachRelationshipValue(window, 0, length, prefixes, value => missing.delete(value))
		await pause()

		const tail = Math.min(SEAM, length)

		window.copyWithin(0, length - tail, length)
		length = tail
	}

	for (const piece of markup) {
		for (let from = 0; from < piece.length && missing.size > 0;) {
			const count = Math.min(piece.length - from, window.length - length)

			window.set(piece.subarray(from, from + count), length)
			length += count
			from += count

			if (length === window.length) await flush()
		}

		if (missing.size === 0) break
	}

	if (missing.size > 0 && length > 0)
		eachRelationshipValue(window.subarray(0, length), 0, length, prefixes, value => missing.delete(value))

	return missing
}
