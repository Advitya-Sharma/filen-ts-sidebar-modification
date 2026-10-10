import type { XmlDocument } from "../../xml/index"
import {
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
import type { ParagraphProps, RunProps, SectionProps, TableCellProps, TableProps, TableRowProps, XmlNodeRef } from "./types"

interface Entry<T> {
	readonly signature: string
	readonly bytes: Uint8Array
	readonly value: T
}

// Parsed property containers interned by their source bytes: Word repeats the same w:rPr thousands of
// times, so a document holds one object per distinct container and resolution caches can key on identity.
class InternTable<T extends object> {
	private readonly buckets = new Map<number, Entry<T>[]>()
	private readonly parse: (reader: PropsReader, node: number) => T
	private readonly empty: T
	private readonly malformed: XmlNodeRef[]
	count = 0

	constructor(parse: (reader: PropsReader, node: number) => T, empty: T, malformed: XmlNodeRef[]) {
		this.parse = parse
		this.empty = empty
		this.malformed = malformed
	}

	get(document: XmlDocument, node: number): T {
		if (node < 0 || document.isSelfClosing(node)) return this.empty

		const bytes = document.bytes
		const start = document.openEnd(node)
		const end = document.contentEnd(node)

		if (end <= start) return this.empty

		const hash = sampledHash(bytes, start, end)

		const reader = PropsReader.for(document)
		const signature = reader.signature(node)
		let bucket = this.buckets.get(hash)

		if (bucket !== undefined) {
			for (const entry of bucket) {
				if (entry.signature === signature && sameBytes(entry.bytes, bytes, start, end)) return entry.value
			}
		} else {
			bucket = []
			this.buckets.set(hash, bucket)
		}

		const value = this.parseChecked(reader, node)

		// A copy: a view would keep the whole part's buffer alive for as long as the store lives.
		bucket.push({ signature, bytes: bytes.slice(start, end), value })
		this.count++

		return value
	}

	// A container with an unreadable start tag anywhere inside reads as empty and is reported, rather than
	// half-parsed: its attributes could not be trusted.
	private parseChecked(reader: PropsReader, node: number): T {
		const document = reader.document
		const malformed = document.malformed

		if (malformed.length > 0) {
			// The parser lists malformed elements in node order; the subtree is the nodes from `node` up to
			// the first that starts after it ends.
			let low = 0
			let high = malformed.length

			while (low < high) {
				const middle = (low + high) >> 1

				if ((malformed[middle] ?? 0) < node) low = middle + 1
				else high = middle
			}

			const first = malformed[low] ?? -1

			if (first >= 0 && document.start(first) < document.end(node)) return this.reject(document, node)
		}

		try {
			return this.parse(reader, node)
		} catch {
			return this.reject(document, node)
		}
	}

	private reject(document: XmlDocument, node: number): T {
		this.malformed.push({ document, node })

		return this.empty
	}
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

function sameBytes(a: Uint8Array, bytes: Uint8Array, start: number, end: number): boolean {
	if (a.length !== end - start) return false

	for (let i = 0; i < a.length; i++) {
		if (a[i] !== bytes[start + i]) return false
	}

	return true
}

// One per open document. Every accessor takes the part and the container element (w:rPr, w:pPr, w:tblPr or
// w:tblPrEx, w:trPr, w:tcPr, w:sectPr); -1 gives the shared empty object.
export class PropsStore {
	// Containers read as empty because their markup was malformed (each distinct container once).
	readonly malformed: XmlNodeRef[] = []
	private readonly runs = new InternTable(parseRunProps, EMPTY_RUN, this.malformed)
	private readonly paragraphs = new InternTable(parseParagraphProps, EMPTY_PARAGRAPH, this.malformed)
	private readonly tables = new InternTable(parseTableProps, EMPTY_TABLE, this.malformed)
	private readonly rows = new InternTable(parseTableRowProps, EMPTY_ROW, this.malformed)
	private readonly cells = new InternTable(parseTableCellProps, EMPTY_CELL, this.malformed)
	private readonly sections = new InternTable(parseSectionProps, EMPTY_SECTION, this.malformed)

	run(document: XmlDocument, node: number): RunProps {
		return this.runs.get(document, node)
	}

	paragraph(document: XmlDocument, node: number): ParagraphProps {
		return this.paragraphs.get(document, node)
	}

	table(document: XmlDocument, node: number): TableProps {
		return this.tables.get(document, node)
	}

	row(document: XmlDocument, node: number): TableRowProps {
		return this.rows.get(document, node)
	}

	cell(document: XmlDocument, node: number): TableCellProps {
		return this.cells.get(document, node)
	}

	section(document: XmlDocument, node: number): SectionProps {
		return this.sections.get(document, node)
	}

	// Distinct containers parsed so far, per kind.
	stats(): { runs: number; paragraphs: number; tables: number; rows: number; cells: number; sections: number } {
		return {
			runs: this.runs.count,
			paragraphs: this.paragraphs.count,
			tables: this.tables.count,
			rows: this.rows.count,
			cells: this.cells.count,
			sections: this.sections.count
		}
	}
}
