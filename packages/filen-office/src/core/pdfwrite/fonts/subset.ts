import type { SfntFont } from "../../fonts/sfnt"

// TrueType subsetting for PDF embedding (ISO 32000-1 §9.9). The glyphs used (and the components of composite
// ones) are kept and numbered densely from 0 (.notdef), the content stream's codes reaching them through the
// CIDToGIDMap; outlines are copied as they are, without hinting instructions, so the hinting programs (cvt,
// fpgm, prep: needed only by instructions) are left out, as other PDF producers do. name, OS/2 and post (format
// 3) are kept so the program is a complete font for viewers that insist, and the copyright notice travels with it.

const KEEP = ["name", "OS/2"]

const ARG_1_AND_2_ARE_WORDS = 0x0001
const WE_HAVE_A_SCALE = 0x0008
const MORE_COMPONENTS = 0x0020
const WE_HAVE_AN_X_AND_Y_SCALE = 0x0040
const WE_HAVE_A_TWO_BY_TWO = 0x0080
const WE_HAVE_INSTRUCTIONS = 0x0100
const MAX_COMPONENT_DEPTH = 16
// A glyph id the subset does not have.
export const NO_GLYPH = 0xffff

export class SubsetError extends Error {}

export interface TrueTypeSubset {
	readonly program: Uint8Array
	// The subset's glyph id of each original glyph id (NO_GLYPH when left out).
	readonly glyphs: Uint16Array
}

interface Table {
	readonly tag: string
	readonly data: Uint8Array
}

function table(font: SfntFont, tag: string): Uint8Array | undefined {
	const range = font.table(tag)

	return range === undefined ? undefined : font.bytes.subarray(range.offset, range.offset + range.length)
}

function required(font: SfntFont, tag: string): Uint8Array {
	const data = table(font, tag)

	if (data === undefined) throw new SubsetError(`The font has no ${tag} table`)

	return data
}

// Glyph data ranges [start, end) within glyf, from loca.
function glyphRanges(font: SfntFont, head: Uint8Array, loca: Uint8Array, glyfLength: number): Uint32Array {
	const long = new DataView(head.buffer, head.byteOffset, head.byteLength).getInt16(50) === 1
	const count = font.numGlyphs
	const view = new DataView(loca.buffer, loca.byteOffset, loca.byteLength)
	const ranges = new Uint32Array(count + 1)

	if (loca.length < (count + 1) * (long ? 4 : 2)) throw new SubsetError("loca is too short")

	for (let g = 0; g <= count; g++) {
		const offset = long ? view.getUint32(g * 4) : view.getUint16(g * 2) * 2

		ranges[g] = Math.min(offset, glyfLength)
	}

	return ranges
}

// Calls `visit` with the offset of each component record of a composite glyph (glyf[start, end)), and returns
// where the records end, or -1 for a simple or malformed glyph.
function components(view: DataView, start: number, end: number, visit: (at: number, flags: number) => void): number {
	if (end - start < 10 || view.getInt16(start) >= 0) return -1

	let at = start + 10

	for (;;) {
		if (at + 4 > end) return -1

		const flags = view.getUint16(at)

		visit(at, flags)
		at += 4 + (flags & ARG_1_AND_2_ARE_WORDS ? 4 : 2)

		if (flags & WE_HAVE_A_SCALE) at += 2
		else if (flags & WE_HAVE_AN_X_AND_Y_SCALE) at += 4
		else if (flags & WE_HAVE_A_TWO_BY_TWO) at += 8

		if (!(flags & MORE_COMPONENTS)) return Math.min(at, end)
	}
}

// Adds the components of composite glyphs to `used` (glyph ids index the bit set).
function closeComposites(used: Uint8Array, glyf: Uint8Array, ranges: Uint32Array, numGlyphs: number): void {
	const view = new DataView(glyf.buffer, glyf.byteOffset, glyf.byteLength)
	const stack: number[] = []

	for (let g = 0; g < numGlyphs; g++) if (used[g] === 1) stack.push(g)

	let guard = numGlyphs * MAX_COMPONENT_DEPTH

	while (stack.length > 0 && guard-- > 0) {
		const glyph = stack.pop() ?? 0

		components(view, ranges[glyph] ?? 0, ranges[glyph + 1] ?? 0, at => {
			const component = view.getUint16(at + 2)

			if (component < numGlyphs && used[component] !== 1) {
				used[component] = 1
				stack.push(component)
			}
		})
	}
}

// A glyph's outline without its instructions, composite components renumbered.
function copyGlyph(glyf: Uint8Array, start: number, end: number, glyphs: Uint16Array): Uint8Array {
	const view = new DataView(glyf.buffer, glyf.byteOffset, glyf.byteLength)

	if (end - start < 10) return glyf.slice(start, end)

	const contours = view.getInt16(start)

	if (contours >= 0) {
		const lengthAt = start + 10 + contours * 2

		if (lengthAt + 2 > end) return glyf.slice(start, end)

		const instructions = view.getUint16(lengthAt)
		const rest = Math.min(end, lengthAt + 2 + instructions)
		const out = new Uint8Array(lengthAt + 2 - start + (end - rest))

		out.set(glyf.subarray(start, lengthAt + 2))
		out[lengthAt - start] = 0
		out[lengthAt - start + 1] = 0
		out.set(glyf.subarray(rest, end), lengthAt + 2 - start)

		return out
	}

	let last = -1
	const recordsEnd = components(view, start, end, at => {
		last = at
	})

	if (recordsEnd < 0) return glyf.slice(start, end)

	const out = glyf.slice(start, recordsEnd)
	const outView = new DataView(out.buffer)

	components(outView, 0, out.length, (at, flags) => {
		outView.setUint16(at + 2, glyphs[outView.getUint16(at + 2)] ?? 0)

		if (at === last - start && (flags & WE_HAVE_INSTRUCTIONS) !== 0) outView.setUint16(at, flags & ~WE_HAVE_INSTRUCTIONS)
	})

	return out
}

// `used` has one byte per glyph id (1: used); glyph 0 is always kept.
// Throws SubsetError unless the font has what subsetting reads: head (indexToLocFormat at 50), hhea
// (numberOfHMetrics at 34), maxp (numGlyphs at 4), hmtx, glyf and a loca covering every glyph. Checked when
// the font loads, so a face that cannot be subset gives way to its substitute before anything is drawn.
export function checkTrueTypeSubset(font: SfntFont): void {
	if (font.cff) throw new SubsetError("CFF outlines are not subset here")

	const head = required(font, "head")

	if (head.length < 54) throw new SubsetError("head is too short")
	if (required(font, "hhea").length < 36) throw new SubsetError("hhea is too short")
	if (required(font, "maxp").length < 6) throw new SubsetError("maxp is too short")

	required(font, "hmtx")
	required(font, "glyf")

	const long = new DataView(head.buffer, head.byteOffset, head.byteLength).getInt16(50) === 1

	if (required(font, "loca").length < (font.numGlyphs + 1) * (long ? 4 : 2)) throw new SubsetError("loca is too short")
}

export function subsetTrueType(font: SfntFont, used: Uint8Array): TrueTypeSubset {
	checkTrueTypeSubset(font)

	const head = required(font, "head")
	const hhea = required(font, "hhea")
	const maxp = required(font, "maxp")
	const hmtx = required(font, "hmtx")
	const loca = required(font, "loca")
	const glyf = required(font, "glyf")
	const numGlyphs = font.numGlyphs
	const keep = new Uint8Array(numGlyphs)

	keep.set(used.subarray(0, numGlyphs))
	keep[0] = 1

	const ranges = glyphRanges(font, head, loca, glyf.length)

	closeComposites(keep, glyf, ranges, numGlyphs)

	const glyphs = new Uint16Array(numGlyphs).fill(NO_GLYPH)
	const kept: number[] = []

	for (let g = 0; g < numGlyphs; g++) {
		if (keep[g] !== 1) continue

		glyphs[g] = kept.length
		kept.push(g)
	}

	const count = kept.length
	const outlines = kept.map(g => copyGlyph(glyf, ranges[g] ?? 0, Math.max(ranges[g] ?? 0, ranges[g + 1] ?? 0), glyphs))
	let size = 0

	for (const outline of outlines) size += pad4(outline.length)

	const shortLoca = size < 0x20000
	const newGlyf = new Uint8Array(size)
	const newLoca = new Uint8Array((count + 1) * (shortLoca ? 2 : 4))
	const locaView = new DataView(newLoca.buffer)
	let offset = 0

	for (let g = 0; g <= count; g++) {
		if (shortLoca) locaView.setUint16(g * 2, offset / 2)
		else locaView.setUint32(g * 4, offset)

		const outline = outlines[g]

		if (outline === undefined) continue

		newGlyf.set(outline, offset)
		offset += pad4(outline.length)
	}

	const newHead = head.slice()
	const headView = new DataView(newHead.buffer)

	headView.setUint32(8, 0)
	headView.setInt16(50, shortLoca ? 0 : 1)

	const newMaxp = maxp.slice()
	const maxpView = new DataView(newMaxp.buffer)

	maxpView.setUint16(4, count)

	// Version 1.0 maxp: no instructions are left.
	if (newMaxp.length >= 32) maxpView.setUint16(26, 0)

	// hmtx: a long metric (advance, left side bearing) per glyph kept.
	const hheaView = new DataView(hhea.buffer, hhea.byteOffset, hhea.byteLength)
	const oldLong = Math.max(1, hheaView.getUint16(34))
	const hmtxView = new DataView(hmtx.buffer, hmtx.byteOffset, hmtx.byteLength)
	const newHmtx = new Uint8Array(count * 4)
	const newHmtxView = new DataView(newHmtx.buffer)
	const read = (at: number): number => (at + 2 <= hmtx.length ? hmtxView.getUint16(at) : 0)

	kept.forEach((g, index) => {
		const advance = g < oldLong ? read(g * 4) : read((oldLong - 1) * 4)
		const bearing = g < oldLong ? read(g * 4 + 2) : read(oldLong * 4 + (g - oldLong) * 2)

		newHmtxView.setUint16(index * 4, advance)
		newHmtxView.setUint16(index * 4 + 2, bearing)
	})

	const newHhea = hhea.slice()

	new DataView(newHhea.buffer).setUint16(34, count)

	const tables: Table[] = [
		{ tag: "head", data: newHead },
		{ tag: "hhea", data: newHhea },
		{ tag: "maxp", data: newMaxp },
		{ tag: "hmtx", data: newHmtx },
		{ tag: "loca", data: newLoca },
		{ tag: "glyf", data: newGlyf },
		{ tag: "cmap", data: EMPTY_CMAP },
		{ tag: "post", data: post3(table(font, "post")) }
	]

	for (const tag of KEEP) {
		const data = table(font, tag)

		if (data !== undefined) tables.push({ tag, data })
	}

	return { program: assemble(tables), glyphs }
}

function pad4(value: number): number {
	return (value + 3) & ~3
}

// A format 4 cmap with only the closing segment: glyphs are reached by CID, never by code point.
const EMPTY_CMAP = Uint8Array.from([
	0, 0, 0, 1, 0, 3, 0, 1, 0, 0, 0, 12, 0, 4, 0, 24, 0, 0, 0, 2, 0, 2, 0, 0, 0, 0, 255, 255, 0, 0, 255, 255, 0, 1, 0, 0
])

// post format 3 (no glyph names), keeping the italic angle and underline metrics.
function post3(post: Uint8Array | undefined): Uint8Array {
	const out = new Uint8Array(32)

	if (post !== undefined && post.length >= 32) out.set(post.subarray(0, 32))

	new DataView(out.buffer).setUint32(0, 0x00030000)

	return out
}

function checksum(data: Uint8Array): number {
	let sum = 0
	const view = new DataView(data.buffer, data.byteOffset, data.byteLength)
	const whole = data.length & ~3

	for (let i = 0; i < whole; i += 4) sum = (sum + view.getUint32(i)) >>> 0

	if (whole < data.length) {
		let last = 0

		for (let i = whole; i < whole + 4; i++) last = (last << 8) | (data[i] ?? 0)

		sum = (sum + (last >>> 0)) >>> 0
	}

	return sum
}

function assemble(tables: Table[]): Uint8Array {
	tables.sort((a, b) => (a.tag < b.tag ? -1 : a.tag > b.tag ? 1 : 0))

	const count = tables.length
	let size = 12 + count * 16

	for (const t of tables) size += pad4(t.data.length)

	const out = new Uint8Array(size)
	const view = new DataView(out.buffer)
	let power = 1
	let log = 0

	while (power * 2 <= count) {
		power *= 2
		log++
	}

	view.setUint32(0, 0x00010000)
	view.setUint16(4, count)
	view.setUint16(6, power * 16)
	view.setUint16(8, log)
	view.setUint16(10, count * 16 - power * 16)

	let offset = 12 + count * 16
	let headOffset = -1

	tables.forEach((t, index) => {
		const record = 12 + index * 16

		for (let k = 0; k < 4; k++) out[record + k] = t.tag.charCodeAt(k)

		view.setUint32(record + 4, checksum(t.data))
		view.setUint32(record + 8, offset)
		view.setUint32(record + 12, t.data.length)
		out.set(t.data, offset)

		if (t.tag === "head") headOffset = offset

		offset += pad4(t.data.length)
	})

	if (headOffset >= 0) view.setUint32(headOffset + 8, (0xb1b0afba - checksum(out)) >>> 0)

	return out
}
