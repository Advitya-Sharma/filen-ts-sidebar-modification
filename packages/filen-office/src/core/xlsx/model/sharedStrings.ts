import type { XlsxRichRun, XlsxSharedStrings } from "./types"

// Entries pooled into one string each: a large table costs a few strings and an offset array, not a string per
// entry. text() slices (V8 shares the pool's characters for longer slices).
export const POOL_ENTRIES = 4096
const POOL_SHIFT = 12

// Parses a rich entry's runs from its <si> bytes (read/sharedStrings.ts), so this module needs no parser.
export type RunsParser = (si: Uint8Array) => readonly XlsxRichRun[] | null

export class SharedStrings implements XlsxSharedStrings {
	readonly sourceCount: number
	readonly sourceLength: number
	referenceDelta = 0
	private readonly pools: readonly string[]
	// Start of each entry in its pool.
	private readonly starts: Uint32Array
	// <si> bytes of the rich source entries, parsed on demand.
	private readonly richSource: ReadonlyMap<number, Uint8Array>
	private readonly parseRuns: RunsParser
	private readonly richParsed = new Map<number, readonly XlsxRichRun[] | null>()
	private readonly appended: string[] = []
	private readonly appendedRuns = new Map<number, readonly XlsxRichRun[]>()
	private readonly appendedPlain = new Map<string, number>()
	// Open-addressed hash of the plain source entries (index + 1), built on the first find.
	private table: Int32Array | null = null

	constructor(
		sourceCount: number,
		pools: readonly string[],
		starts: Uint32Array,
		length: number,
		richSource: ReadonlyMap<number, Uint8Array>,
		parseRuns: RunsParser
	) {
		this.sourceCount = sourceCount
		this.sourceLength = length
		this.pools = pools
		this.starts = starts
		this.richSource = richSource
		this.parseRuns = parseRuns
	}

	static empty(parseRuns: RunsParser): SharedStrings {
		return new SharedStrings(-1, [], new Uint32Array(0), 0, new Map(), parseRuns)
	}

	get length(): number {
		return this.sourceLength + this.appended.length
	}

	// Whether the source has a sharedStrings part (count -1: none, or no count attribute).
	get hasSourceCount(): boolean {
		return this.sourceCount >= 0
	}

	text(index: number): string {
		if (index < 0) return ""
		if (index >= this.sourceLength) return this.appended[index - this.sourceLength] ?? ""

		const pool = this.pools[index >>> POOL_SHIFT] ?? ""
		const next = index + 1

		return pool.slice(
			this.starts[index] ?? 0,
			(next & (POOL_ENTRIES - 1)) === 0 || next === this.sourceLength ? pool.length : (this.starts[next] ?? 0)
		)
	}

	isRich(index: number): boolean {
		return index < this.sourceLength ? this.richSource.has(index) : this.appendedRuns.has(index)
	}

	runs(index: number): readonly XlsxRichRun[] | null {
		if (index >= this.sourceLength) return this.appendedRuns.get(index) ?? null

		const bytes = this.richSource.get(index)

		if (bytes === undefined) return null

		let runs = this.richParsed.get(index)

		if (runs === undefined) {
			runs = this.parseRuns(bytes)
			this.richParsed.set(index, runs)
		}

		return runs
	}

	find(text: string): number {
		const appended = this.appendedPlain.get(text)

		if (appended !== undefined) return appended

		const table = (this.table ??= this.buildTable())
		const mask = table.length - 1

		for (let slot = hashString(text) & mask; ; slot = (slot + 1) & mask) {
			const entry = table[slot] ?? 0

			if (entry === 0) return -1
			if (this.text(entry - 1) === text) return entry - 1
		}
	}

	append(text: string, runs?: readonly XlsxRichRun[]): number {
		if (runs === undefined || runs.length === 0) {
			const existing = this.appendedPlain.get(text)

			if (existing !== undefined) return existing
		}

		const index = this.sourceLength + this.appended.length

		this.appended.push(text)

		if (runs !== undefined && runs.length > 0) this.appendedRuns.set(index, runs)
		else this.appendedPlain.set(text, index)

		return index
	}

	adjustReferences(delta: number): void {
		this.referenceDelta += delta
	}

	private buildTable(): Int32Array {
		let size = 16

		while (size < this.sourceLength * 2) size *= 2

		const table = new Int32Array(size)
		const mask = size - 1

		for (let index = 0; index < this.sourceLength; index++) {
			if (this.richSource.has(index)) continue

			const pool = this.pools[index >>> POOL_SHIFT] ?? ""
			const next = index + 1
			const start = this.starts[index] ?? 0
			const end = (next & (POOL_ENTRIES - 1)) === 0 || next === this.sourceLength ? pool.length : (this.starts[next] ?? 0)
			let hash = 0x811c9dc5

			for (let k = start; k < end; k++) hash = Math.imul(hash ^ pool.charCodeAt(k), 0x01000193)

			let slot = hash & mask

			while ((table[slot] ?? 0) !== 0) {
				// The first of equal entries wins, as a reader looking it up would find it.
				if (
					this.text((table[slot] ?? 0) - 1).length === end - start &&
					this.text((table[slot] ?? 0) - 1) === pool.slice(start, end)
				)
					break

				slot = (slot + 1) & mask
			}

			if ((table[slot] ?? 0) === 0) table[slot] = index + 1
		}

		return table
	}

	estimateBytes(): number {
		let bytes = this.starts.byteLength + (this.table?.byteLength ?? 0)

		for (const pool of this.pools) bytes += pool.length * 2

		for (const si of this.richSource.values()) bytes += si.byteLength + 64

		for (const text of this.appended) bytes += text.length * 2 + 32

		return bytes
	}
}

function hashString(text: string): number {
	let hash = 0x811c9dc5

	for (let k = 0; k < text.length; k++) hash = Math.imul(hash ^ text.charCodeAt(k), 0x01000193)

	return hash
}
