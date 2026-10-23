// Hostile OpenType layout tables for the PDF font tests: a GSUB built of given lookups, put into a real font.

// A GSUB whose `ccmp` (DFLT and latn) runs the given lookups in order, each [type, its one subtable's bytes].
export function gsubTable(lookups: readonly (readonly [number, (lookupIndex: number) => number[]])[]): Uint8Array {
	const out: number[] = []
	const u16 = (...values: number[]): void => {
		for (const value of values) out.push((value >>> 8) & 0xff, value & 0xff)
	}
	const set = (offset: number, value: number): void => {
		out[offset] = (value >>> 8) & 0xff
		out[offset + 1] = value & 0xff
	}
	const tag = (value: string): void => {
		for (let k = 0; k < 4; k++) out.push(value.charCodeAt(k))
	}

	u16(1, 0, 0, 0, 0)
	set(4, out.length)
	u16(2)
	tag("DFLT")
	u16(14)
	tag("latn")
	u16(14)
	u16(4, 0)
	u16(0, 0xffff, 1, 0)
	set(6, out.length)
	u16(1)
	tag("ccmp")
	u16(8)
	u16(0, lookups.length, ...lookups.map((_, k) => k))
	set(8, out.length)

	const lookupList = out.length

	u16(lookups.length, ...lookups.map(() => 0))
	lookups.forEach(([type, subtable], k) => {
		set(lookupList + 2 + k * 2, out.length - lookupList)
		u16(type, 0, 1, 8)
		out.push(...subtable(k))
	})

	return Uint8Array.from(out)
}

function words(...values: number[]): number[] {
	return values.flatMap(value => [(value >>> 8) & 0xff, value & 0xff])
}

// MultipleSubst format 1: `glyph` becomes `count` times itself.
export function multiply(glyph: number, count: number): () => number[] {
	return () => words(1, 8, 1, 14, 1, 1, glyph, count, ...new Array<number>(count).fill(glyph))
}

// SequenceContext format 3 matching `glyph`, running the given lookup `records` times at it (itself: recursion).
export function recurse(glyph: number, records: number): (self: number) => number[] {
	return self => words(3, 1, records, 6 + 2 + records * 4, ...new Array<number>(records).fill(0).flatMap(() => [0, self]), 1, 1, glyph)
}

export function directoryRecord(bytes: Uint8Array, tag: string): number {
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)

	for (let t = 0; t < view.getUint16(4); t++) {
		const record = 12 + t * 16

		if (String.fromCharCode(...bytes.subarray(record, record + 4)) === tag) return record
	}

	throw new Error(`No ${tag}`)
}

// A font with its `tag` table replaced (the new table appended, its directory record pointed at it).
export function withTable(font: Uint8Array, tag: string, table: Uint8Array): Uint8Array {
	const start = (font.length + 3) & ~3
	const out = new Uint8Array(start + table.length)

	out.set(font)
	out.set(table, start)

	const view = new DataView(out.buffer)
	const record = directoryRecord(out, tag)

	view.setUint32(record + 8, start)
	view.setUint32(record + 12, table.length)

	return out
}
