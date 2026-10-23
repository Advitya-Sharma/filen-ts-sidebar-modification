// OpenType layout common tables (OpenType 1.9 "OpenType Layout common table formats"): script and feature
// selection, lookup headers, Coverage and ClassDef, plus GDEF glyph classes. Reads go through a DataView;
// a malformed table throws a RangeError, which the shaper catches (the run is then drawn unshaped).

export const LookupFlag = {
	RIGHT_TO_LEFT: 0x0001,
	IGNORE_BASE_GLYPHS: 0x0002,
	IGNORE_LIGATURES: 0x0004,
	IGNORE_MARKS: 0x0008,
	USE_MARK_FILTERING_SET: 0x0010,
	MARK_ATTACHMENT_TYPE: 0xff00
} as const

export const GlyphClass = { UNKNOWN: 0, BASE: 1, LIGATURE: 2, MARK: 3, COMPONENT: 4 } as const

export interface Lookup {
	readonly type: number
	readonly flag: number
	readonly markFilteringSet: number
	// Subtable offsets, extension subtables already resolved.
	readonly subtables: readonly number[]
}

const DEFAULT_LANGUAGE = "dflt"

export function tagAt(view: DataView, offset: number): string {
	return String.fromCharCode(view.getUint8(offset), view.getUint8(offset + 1), view.getUint8(offset + 2), view.getUint8(offset + 3))
}

// Coverage index of `glyph`, or -1.
export function coverageIndex(view: DataView, offset: number, glyph: number): number {
	const format = view.getUint16(offset)

	if (format === 1) {
		let low = 0
		let high = view.getUint16(offset + 2) - 1

		while (low <= high) {
			const middle = (low + high) >>> 1
			const value = view.getUint16(offset + 4 + middle * 2)

			if (value === glyph) return middle
			if (value < glyph) low = middle + 1
			else high = middle - 1
		}

		return -1
	}

	if (format === 2) {
		let low = 0
		let high = view.getUint16(offset + 2) - 1

		while (low <= high) {
			const middle = (low + high) >>> 1
			const record = offset + 4 + middle * 6
			const start = view.getUint16(record)

			if (glyph < start) high = middle - 1
			else if (glyph > view.getUint16(record + 2)) low = middle + 1
			else return view.getUint16(record + 4) + glyph - start
		}
	}

	return -1
}

export function classOf(view: DataView, offset: number, glyph: number): number {
	const format = view.getUint16(offset)

	if (format === 1) {
		const start = view.getUint16(offset + 2)
		const count = view.getUint16(offset + 4)

		return glyph >= start && glyph < start + count ? view.getUint16(offset + 6 + (glyph - start) * 2) : 0
	}

	if (format === 2) {
		let low = 0
		let high = view.getUint16(offset + 2) - 1

		while (low <= high) {
			const middle = (low + high) >>> 1
			const record = offset + 4 + middle * 6

			if (glyph < view.getUint16(record)) high = middle - 1
			else if (glyph > view.getUint16(record + 2)) low = middle + 1
			else return view.getUint16(record + 4)
		}
	}

	return 0
}

// GSUB or GPOS header. `extensionType` is the lookup type that wraps others (GSUB 7, GPOS 9).
export class LayoutTable {
	readonly view: DataView
	private readonly scriptList: number
	private readonly featureList: number
	private readonly lookupList: number
	private readonly extensionType: number
	private readonly lookups = new Map<number, Lookup>()
	private readonly digests = new Map<number, Uint32Array | null>()

	constructor(view: DataView, offset: number, extensionType: number) {
		this.view = view
		this.extensionType = extensionType

		const major = view.getUint16(offset)

		if (major !== 1) throw new RangeError("Unsupported layout table version")

		this.scriptList = offset + view.getUint16(offset + 4)
		this.featureList = offset + view.getUint16(offset + 6)
		this.lookupList = offset + view.getUint16(offset + 8)
	}

	// The first script tag the table has, of `tags` in preference order, else DFLT, else none.
	scriptOffset(tags: readonly string[]): { readonly tag: string; readonly offset: number } | undefined {
		const view = this.view
		const count = view.getUint16(this.scriptList)
		const found = new Map<string, number>()

		for (let i = 0; i < count; i++) {
			const record = this.scriptList + 2 + i * 6

			found.set(tagAt(view, record), this.scriptList + view.getUint16(record + 4))
		}

		for (const tag of [...tags, "DFLT", "dflt", "latn"]) {
			const offset = found.get(tag)

			if (offset !== undefined) return { tag, offset }
		}

		return undefined
	}

	// Lookup indices per feature tag for a script and language system, each list ascending.
	featureLookups(scriptTags: readonly string[], language: string | undefined): Map<string, number[]> {
		const out = new Map<string, number[]>()
		const script = this.scriptOffset(scriptTags)

		if (script === undefined) return out

		const view = this.view
		let langSys = view.getUint16(script.offset) === 0 ? 0 : script.offset + view.getUint16(script.offset)
		const langCount = view.getUint16(script.offset + 2)

		if (language !== undefined && language !== DEFAULT_LANGUAGE) {
			for (let i = 0; i < langCount; i++) {
				const record = script.offset + 4 + i * 6

				if (tagAt(view, record) === language) {
					langSys = script.offset + view.getUint16(record + 4)
					break
				}
			}
		}

		if (langSys === 0) return out

		const required = view.getUint16(langSys + 2)
		const featureCount = view.getUint16(langSys + 4)
		const indices: number[] = []

		if (required !== 0xffff) indices.push(required)

		for (let i = 0; i < featureCount; i++) indices.push(view.getUint16(langSys + 6 + i * 2))

		const totalFeatures = view.getUint16(this.featureList)

		for (const index of indices) {
			if (index >= totalFeatures) continue

			const record = this.featureList + 2 + index * 6
			const tag = tagAt(view, record)
			const feature = this.featureList + view.getUint16(record + 4)
			const lookupCount = view.getUint16(feature + 2)
			let list = out.get(tag)

			if (list === undefined) {
				list = []
				out.set(tag, list)
			}

			for (let k = 0; k < lookupCount; k++) list.push(view.getUint16(feature + 4 + k * 2))
		}

		for (const list of out.values()) list.sort((a, b) => a - b)

		return out
	}

	get lookupCount(): number {
		return this.view.getUint16(this.lookupList)
	}

	digest(index: number): Uint32Array | null {
		let digest = this.digests.get(index)

		if (digest !== undefined) return digest

		const lookup = this.lookup(index)
		const bits: number[] = []

		digest = null

		if (lookup !== undefined && lookup.type !== 0) {
			const gsub = this.extensionType === 7
			let complete = true

			for (const subtable of lookup.subtables) {
				const coverage = firstCoverage(this.view, lookup.type, subtable, gsub)

				if (coverage < 0) {
					complete = false
					break
				}

				addCoverage(this.view, coverage, bits)
			}

			if (complete) digest = Uint32Array.from(bits)
		}

		this.digests.set(index, digest)

		return digest
	}

	lookup(index: number): Lookup | undefined {
		let lookup = this.lookups.get(index)

		if (lookup !== undefined) return lookup
		if (index >= this.lookupCount) return undefined

		const view = this.view
		const offset = this.lookupList + view.getUint16(this.lookupList + 2 + index * 2)
		const declared = view.getUint16(offset)
		let type = declared
		const flag = view.getUint16(offset + 2)
		const count = view.getUint16(offset + 4)
		const subtables: number[] = []

		for (let i = 0; i < count; i++) {
			let subtable = offset + view.getUint16(offset + 6 + i * 2)

			if (declared === this.extensionType) {
				// Every extension subtable of a lookup wraps the same type.
				const wrapped = view.getUint16(subtable + 2)

				subtable += view.getUint32(subtable + 4)

				if (subtables.length === 0 && type === declared) type = wrapped
				else if (wrapped !== type) continue
			}

			subtables.push(subtable)
		}

		if (type === this.extensionType) type = 0

		lookup = {
			type,
			flag,
			markFilteringSet: flag & LookupFlag.USE_MARK_FILTERING_SET ? view.getUint16(offset + 6 + count * 2) : 0,
			subtables
		}
		this.lookups.set(index, lookup)

		return lookup
	}
}

// The glyphs a lookup can start at (the union of its subtables' first coverage), as a bit set, so most glyphs
// are rejected without a coverage search. Null when a subtable's first glyph is not given by a coverage table.
function firstCoverage(view: DataView, type: number, offset: number, gsub: boolean): number {
	const context = gsub ? type === 5 || type === 6 : type === 7 || type === 8

	if (!context) return offset + view.getUint16(offset + 2)

	const format = view.getUint16(offset)

	if (format !== 3) return offset + view.getUint16(offset + 2)

	const chained = gsub ? type === 6 : type === 8

	if (!chained) return view.getUint16(offset + 2) === 0 ? -1 : offset + view.getUint16(offset + 6)

	const backtrack = view.getUint16(offset + 2)
	const inputCountAt = offset + 4 + backtrack * 2

	return view.getUint16(inputCountAt) === 0 ? -1 : offset + view.getUint16(inputCountAt + 2)
}

function addCoverage(view: DataView, offset: number, bits: number[]): void {
	const set = (glyph: number): void => {
		const word = glyph >>> 5

		while (bits.length <= word) bits.push(0)

		bits[word] = (bits[word] ?? 0) | (1 << (glyph & 31))
	}
	const format = view.getUint16(offset)
	const count = view.getUint16(offset + 2)

	if (format === 1) {
		for (let i = 0; i < count; i++) set(view.getUint16(offset + 4 + i * 2))
	} else if (format === 2) {
		for (let i = 0; i < count; i++) {
			const record = offset + 4 + i * 6
			const last = view.getUint16(record + 2)

			for (let glyph = view.getUint16(record); glyph <= last; glyph++) set(glyph)
		}
	}
}

// GDEF glyph classes, mark attachment classes and mark glyph sets, expanded per glyph on first use.
export class GlyphDefinitions {
	private readonly view: DataView | null
	private readonly glyphClassDef: number
	private readonly markClassDef: number
	private readonly markSets: number
	private classes: Uint8Array | null = null
	private markClasses: Uint8Array | null = null
	private readonly numGlyphs: number

	constructor(view: DataView | null, offset: number, numGlyphs: number) {
		this.numGlyphs = numGlyphs

		if (view === null) {
			this.view = null
			this.glyphClassDef = 0
			this.markClassDef = 0
			this.markSets = 0

			return
		}

		this.view = view

		const minor = view.getUint16(offset + 2)

		this.glyphClassDef = view.getUint16(offset + 4) === 0 ? 0 : offset + view.getUint16(offset + 4)
		this.markClassDef = view.getUint16(offset + 10) === 0 ? 0 : offset + view.getUint16(offset + 10)
		this.markSets = minor >= 2 && view.getUint16(offset + 12) !== 0 ? offset + view.getUint16(offset + 12) : 0
	}

	get hasClasses(): boolean {
		return this.glyphClassDef !== 0
	}

	glyphClass(glyph: number): number {
		if (this.view === null || this.glyphClassDef === 0) return GlyphClass.UNKNOWN

		this.classes ??= this.expand(this.glyphClassDef)

		return this.classes[glyph] ?? 0
	}

	markClass(glyph: number): number {
		if (this.view === null || this.markClassDef === 0) return 0

		this.markClasses ??= this.expand(this.markClassDef)

		return this.markClasses[glyph] ?? 0
	}

	inMarkSet(set: number, glyph: number): boolean {
		const view = this.view

		if (view === null || this.markSets === 0 || set >= view.getUint16(this.markSets + 2)) return false

		return coverageIndex(view, this.markSets + view.getUint32(this.markSets + 4 + set * 4), glyph) >= 0
	}

	private expand(offset: number): Uint8Array {
		const out = new Uint8Array(this.numGlyphs)
		const view = this.view

		if (view === null) return out

		const format = view.getUint16(offset)

		if (format === 1) {
			const start = view.getUint16(offset + 2)
			const count = view.getUint16(offset + 4)

			for (let i = 0; i < count && start + i < out.length; i++) out[start + i] = view.getUint16(offset + 6 + i * 2)
		} else if (format === 2) {
			const count = view.getUint16(offset + 2)

			for (let i = 0; i < count; i++) {
				const record = offset + 4 + i * 6

				out.fill(view.getUint16(record + 4), view.getUint16(record), Math.min(out.length, view.getUint16(record + 2) + 1))
			}
		}

		return out
	}
}
