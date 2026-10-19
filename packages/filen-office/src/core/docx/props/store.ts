import type { XmlDocument } from "../../xml/index"
import { Ns } from "../model/namespaces"
import {
	CHANGE_ELEMENTS,
	EMPTY_CELL,
	EMPTY_PARAGRAPH,
	EMPTY_ROW,
	EMPTY_RUN,
	EMPTY_SECTION,
	EMPTY_TABLE,
	parseParagraphProps,
	parseRunProps,
	parseSectionProps,
	parseTableCellProps,
	parseTableProps,
	parseTableRowProps
} from "./parse"
import { PropsReader } from "./reader"
import type { ParagraphProps, RevisionMark, RunProps, SectionProps, TableCellProps, TableProps, TableRowProps, XmlNodeRef } from "./types"

interface Entry<T> {
	readonly signature: string
	readonly bytes: Uint8Array
	readonly value: T
}

// Tracked-change stamps in a property container: the change element (who changed the properties, when, what
// they were) and a revision mark's w:id and w:date. Every tracked edit writes new ones, so what is drawn
// (`runLook` and the like) interns a container by its bytes without them, and holds one object per formatting
// however many edits were tracked.
const REVISION_MARKS: ReadonlySet<string> = new Set(["ins", "del", "moveFrom", "moveTo", "cellIns", "cellDel", "cellMerge"])

// Whether a part may hold stamps at all: a name of a change element or revision mark among its element names
// (by local name, whatever the prefix binds). Most parts have none, and their containers are drawn as they are.
const stampedParts = new WeakMap<XmlDocument, boolean>()

function mayHoldStamps(document: XmlDocument): boolean {
	let found = stampedParts.get(document)

	if (found === undefined) {
		found = document.names.some(name => {
			const local = name.slice(name.indexOf(":") + 1)

			return CHANGE_ELEMENTS.has(local) || REVISION_MARKS.has(local)
		})
		stampedParts.set(document, found)
	}

	return found
}

// Appends the byte ranges of the stamps in `node` (a paragraph's mark w:rPr included), ascending, as start/end
// pairs; whether any is an attribute value (the rest are whole change elements).
function stampRanges(r: PropsReader, node: number, out: number[], span: Int32Array): boolean {
	const document = r.document
	let values = false

	for (let c = document.firstChildElement(node); c >= 0; c = document.nextSiblingElement(c)) {
		const local = r.local(c)

		if (CHANGE_ELEMENTS.has(local)) {
			out.push(document.start(c), document.end(c))
		} else if (REVISION_MARKS.has(local)) {
			const at = out.length

			if (r.model.find(c, Ns.W, "id", span)) out.push(span[0] ?? 0, span[1] ?? 0)
			if (r.model.find(c, Ns.W, "date", span)) out.push(span[0] ?? 0, span[1] ?? 0)

			// The attributes in either order.
			if (out.length - at === 4 && (out[at] ?? 0) > (out[at + 2] ?? 0)) out.push(...out.splice(at, 2))

			values ||= out.length > at
		} else if (local === "rPr") {
			values = stampRanges(r, c, out, span) || values
		}
	}

	return values
}

// The bytes from `start` to `end` without `ranges`, into a buffer reused from call to call; their length.
let masked = new Uint8Array(512)

function mask(bytes: Uint8Array, start: number, end: number, ranges: readonly number[]): number {
	if (masked.length < end - start) masked = new Uint8Array(Math.max(end - start, masked.length * 2))

	const out = masked
	let length = 0
	let at = start

	for (let k = 0; k <= ranges.length; k += 2) {
		const to = k < ranges.length ? (ranges[k] ?? end) : end

		for (let i = at; i < to; i++) out[length++] = bytes[i] ?? 0

		at = ranges[k + 1] ?? end
	}

	return length
}

const lookRevisions = (revisions: readonly RevisionMark[]): RevisionMark[] =>
	revisions.map(({ kind, author }) => (author === undefined ? { kind } : { kind, author }))

function runLook(props: RunProps): RunProps {
	const { change: _change, revisions, ...rest } = props

	return revisions === undefined ? rest : { ...rest, revisions: lookRevisions(revisions) }
}

function paragraphLook(props: ParagraphProps): ParagraphProps {
	const { change: _change, markRun, ...rest } = props

	return markRun === undefined ? rest : { ...rest, markRun: runLook(markRun) }
}

function tableLook(props: TableProps): TableProps {
	const { change: _change, ...rest } = props

	return rest
}

function rowLook(props: TableRowProps): TableRowProps {
	const { change: _change, revisions, ...rest } = props

	return revisions === undefined ? rest : { ...rest, revisions: lookRevisions(revisions) }
}

function cellLook(props: TableCellProps): TableCellProps {
	const { change: _change, revisions, ...rest } = props

	return revisions === undefined ? rest : { ...rest, revisions: lookRevisions(revisions) }
}

// Parsed property containers interned by their source bytes: Word repeats the same w:rPr thousands of times,
// so a document holds one object per distinct container and resolution caches can key on identity.
class InternTable<T extends object> {
	private readonly buckets = new Map<number, Entry<T>[]>()
	// Containers whose sampled hash a full bucket holds (many that differ only where sampling skips: tracked
	// property changes differing in their w:id), by a hash of every byte.
	private readonly crowded = new Map<number, Entry<T>[]>()
	private readonly parse: (reader: PropsReader, node: number) => T
	private readonly empty: T
	private readonly malformed: XmlNodeRef[]
	// Containers without their stamps, by those bytes (see stampRanges); undefined: not kept for this kind.
	private readonly looks: InternTable<T> | undefined
	private readonly strip: ((value: T) => T) | undefined
	private readonly ranges: number[] = []
	private readonly span = new Int32Array(2)
	count = 0

	constructor(parse: (reader: PropsReader, node: number) => T, empty: T, malformed: XmlNodeRef[], strip?: (value: T) => T) {
		this.parse = parse
		this.empty = empty
		this.malformed = malformed
		this.strip = strip
		this.looks = strip === undefined ? undefined : new InternTable(parse, empty, malformed)
	}

	get looksCount(): number {
		return this.looks?.count ?? 0
	}

	get(document: XmlDocument, node: number): T {
		if (node < 0 || document.isSelfClosing(node)) return this.empty

		const start = document.openEnd(node)
		const end = document.contentEnd(node)

		if (end <= start) return this.empty

		const reader = PropsReader.for(document)
		const signature = reader.signature(node)
		const bytes = document.bytes

		return this.find(bytes, start, end, signature) ?? this.add(bytes, start, end, signature, this.parseChecked(reader, node))
	}

	// The container as drawn: without its stamps, one object for every container equal but for them. One with
	// none (or with malformed markup, which reads as a whole) is the container itself.
	look(document: XmlDocument, node: number): T {
		const looks = this.looks
		const strip = this.strip

		if (looks === undefined || strip === undefined || node < 0 || document.isSelfClosing(node) || !mayHoldStamps(document))
			return this.get(document, node)

		const reader = PropsReader.for(document)
		const ranges = this.ranges

		ranges.length = 0

		const values = stampRanges(reader, node, ranges, this.span)

		if (ranges.length === 0 || firstMalformed(document, node) >= 0) return this.get(document, node)

		const length = mask(document.bytes, document.openEnd(node), document.contentEnd(node), ranges)
		const signature = reader.signature(node)
		const found = looks.find(masked, 0, length, signature)

		if (found !== undefined) return found

		// Without its change elements a container is one without stamps, which may be known already (the same
		// formatting untracked): that object is shared too.
		const twin = values ? undefined : this.find(masked, 0, length, signature)

		return looks.add(masked, 0, length, signature, twin ?? strip(this.parseChecked(reader, node)))
	}

	private find(bytes: Uint8Array, start: number, end: number, signature: string): T | undefined {
		const hash = sampledHash(bytes, start, end)
		const bucket = this.buckets.get(hash)

		if (bucket === undefined) return undefined

		const found = match(bucket, bytes, start, end, signature)

		if (found !== undefined || bucket.length < BUCKET_LIMIT) return found

		const crowded = this.crowded.get(fullHash(bytes, start, end, hash))

		return crowded === undefined ? undefined : match(crowded, bytes, start, end, signature)
	}

	private add(bytes: Uint8Array, start: number, end: number, signature: string, value: T): T {
		const hash = sampledHash(bytes, start, end)
		let bucket = this.buckets.get(hash)

		if (bucket === undefined) {
			bucket = []
			this.buckets.set(hash, bucket)
		} else if (bucket.length >= BUCKET_LIMIT) {
			const full = fullHash(bytes, start, end, hash)

			bucket = this.crowded.get(full)

			if (bucket === undefined) {
				bucket = []
				this.crowded.set(full, bucket)
			}
		}

		// A copy: a view would keep the whole part's buffer alive for as long as the store lives.
		bucket.push({ signature, bytes: bytes.slice(start, end), value })
		this.count++

		return value
	}

	// A container with an unreadable start tag anywhere inside reads as empty and is reported, rather than
	// half-parsed: its attributes could not be trusted.
	private parseChecked(reader: PropsReader, node: number): T {
		if (firstMalformed(reader.document, node) >= 0) return this.reject(reader.document, node)

		try {
			return this.parse(reader, node)
		} catch {
			return this.reject(reader.document, node)
		}
	}

	private reject(document: XmlDocument, node: number): T {
		this.malformed.push({ document, node })

		return this.empty
	}
}

function match<T>(bucket: readonly Entry<T>[], bytes: Uint8Array, start: number, end: number, signature: string): T | undefined {
	for (const entry of bucket) {
		if (entry.signature === signature && sameBytes(entry.bytes, bytes, start, end)) return entry.value
	}

	return undefined
}

// The first element with a malformed start tag in `node`'s subtree, -1 for none. The parser lists them in node
// order; the subtree is the nodes from `node` up to the first that starts after it ends.
function firstMalformed(document: XmlDocument, node: number): number {
	const malformed = document.malformed

	if (malformed.length === 0) return -1

	let low = 0
	let high = malformed.length

	while (low < high) {
		const middle = (low + high) >> 1

		if ((malformed[middle] ?? 0) < node) low = middle + 1
		else high = middle
	}

	const first = malformed[low] ?? -1

	return first >= 0 && document.start(first) < document.end(node) ? first : -1
}

// FNV-1a over the content's first and last 64 bytes and every fourth word between, plus its length: a
// hit is confirmed by a full comparison anyway, so hashing every byte would read the container twice.
function sampledHash(bytes: Uint8Array, start: number, end: number): number {
	let hash = Math.imul(0x811c9dc5 ^ (end - start), 0x01000193)
	const word = (i: number): number =>
		(bytes[i] ?? 0) | ((bytes[i + 1] ?? 0) << 8) | ((bytes[i + 2] ?? 0) << 16) | ((bytes[i + 3] ?? 0) << 24)

	if (end - start <= 160) {
		let i = start

		for (; i + 4 <= end; i += 4) hash = Math.imul(hash ^ word(i), 0x01000193)
		for (; i < end; i++) hash = Math.imul(hash ^ (bytes[i] ?? 0), 0x01000193)

		return hash
	}

	for (let i = start; i < start + 64; i += 4) hash = Math.imul(hash ^ word(i), 0x01000193)
	for (let i = start + 64; i + 4 <= end - 64; i += 16) hash = Math.imul(hash ^ word(i), 0x01000193)
	let i = end - 64

	for (; i + 4 <= end; i += 4) hash = Math.imul(hash ^ word(i), 0x01000193)
	for (; i < end; i++) hash = Math.imul(hash ^ (bytes[i] ?? 0), 0x01000193)

	return hash
}

// Entries a sampled-hash bucket takes before later ones go by a hash of every byte.
const BUCKET_LIMIT = 8

function fullHash(bytes: Uint8Array, start: number, end: number, seed: number): number {
	let hash = seed
	let i = start

	for (; i + 4 <= end; i += 4)
		hash = Math.imul(
			hash ^ ((bytes[i] ?? 0) | ((bytes[i + 1] ?? 0) << 8) | ((bytes[i + 2] ?? 0) << 16) | ((bytes[i + 3] ?? 0) << 24)),
			0x01000193
		)
	for (; i < end; i++) hash = Math.imul(hash ^ (bytes[i] ?? 0), 0x01000193)

	return hash
}

function sameBytes(a: Uint8Array, bytes: Uint8Array, start: number, end: number): boolean {
	if (a.length !== end - start) return false

	for (let i = 0; i < a.length; i++) {
		if (a[i] !== bytes[start + i]) return false
	}

	return true
}

// One per open document. Every accessor takes the part and the container element (w:rPr, w:pPr, w:tblPr or
// w:tblPrEx, w:trPr, w:tcPr, w:sectPr); -1 gives the shared empty object. The `*Look` accessors give what the
// layout draws: the properties without tracked-change stamps (no `change`, revision marks without id and
// date), shared by every container that differs from another only in its stamps.
export class PropsStore {
	// Containers read as empty because their markup was malformed (each distinct container once).
	readonly malformed: XmlNodeRef[] = []
	private readonly runs = new InternTable(parseRunProps, EMPTY_RUN, this.malformed, runLook)
	private readonly paragraphs = new InternTable(parseParagraphProps, EMPTY_PARAGRAPH, this.malformed, paragraphLook)
	private readonly tables = new InternTable(parseTableProps, EMPTY_TABLE, this.malformed, tableLook)
	private readonly rows = new InternTable(parseTableRowProps, EMPTY_ROW, this.malformed, rowLook)
	private readonly cells = new InternTable(parseTableCellProps, EMPTY_CELL, this.malformed, cellLook)
	private readonly sections = new InternTable(parseSectionProps, EMPTY_SECTION, this.malformed)

	run(document: XmlDocument, node: number): RunProps {
		return this.runs.get(document, node)
	}

	runLook(document: XmlDocument, node: number): RunProps {
		return this.runs.look(document, node)
	}

	paragraph(document: XmlDocument, node: number): ParagraphProps {
		return this.paragraphs.get(document, node)
	}

	paragraphLook(document: XmlDocument, node: number): ParagraphProps {
		return this.paragraphs.look(document, node)
	}

	table(document: XmlDocument, node: number): TableProps {
		return this.tables.get(document, node)
	}

	tableLook(document: XmlDocument, node: number): TableProps {
		return this.tables.look(document, node)
	}

	row(document: XmlDocument, node: number): TableRowProps {
		return this.rows.get(document, node)
	}

	rowLook(document: XmlDocument, node: number): TableRowProps {
		return this.rows.look(document, node)
	}

	cell(document: XmlDocument, node: number): TableCellProps {
		return this.cells.get(document, node)
	}

	cellLook(document: XmlDocument, node: number): TableCellProps {
		return this.cells.look(document, node)
	}

	section(document: XmlDocument, node: number): SectionProps {
		return this.sections.get(document, node)
	}

	// Distinct containers parsed so far, per kind (`looks`: those without their stamps, of every kind).
	stats(): { runs: number; paragraphs: number; tables: number; rows: number; cells: number; sections: number; looks: number } {
		const tables = [this.runs, this.paragraphs, this.tables, this.rows, this.cells]

		return {
			runs: this.runs.count,
			paragraphs: this.paragraphs.count,
			tables: this.tables.count,
			rows: this.rows.count,
			cells: this.cells.count,
			sections: this.sections.count,
			looks: tables.reduce((sum, table) => sum + table.looksCount, 0)
		}
	}
}
