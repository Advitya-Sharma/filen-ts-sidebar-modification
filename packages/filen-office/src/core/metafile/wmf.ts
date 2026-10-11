import type { Bytes } from "./bytes"
import { packedDib } from "./dib"
import { Gdi, rgb, solidPen, type GdiObject } from "./gdi"
import { MAX_REGION_RECTS, MetafileLimit } from "./limits"
import { Output } from "./output"
import { decodeAnsi, metafileFont } from "./text"
import type { MetafileConversion } from "./types"

// Windows Metafiles ([MS-WMF]): the optional placeable header, the META_HEADER, then records of 16-bit
// words. Parameters of many records are stored last-first, as the original API pushed them.

const PLACEABLE_KEY = 0x9ac6cdd7
const WMFC = 0x43464d57

const ETO_OPAQUE = 0x2
const ETO_CLIPPED = 0x4

// The picture frame a placeable header declares: its bounding box in logical units and their units per inch.
interface Placement {
	readonly left: number
	readonly top: number
	readonly right: number
	readonly bottom: number
	readonly inch: number
}

export function isWmf(bytes: Bytes): boolean {
	if (bytes.u32(0) === PLACEABLE_KEY) return true

	const type = bytes.u16(0)

	return (type === 1 || type === 2) && bytes.u16(2) === 9 && (bytes.u16(4) === 0x300 || bytes.u16(4) === 0x100)
}

// The first window origin and extent the records set (for files without a placeable header).
function scanWindow(bytes: Bytes, start: number): { org: [number, number]; ext: [number, number] | undefined } {
	const org: [number, number] = [0, 0]
	let ext: [number, number] | undefined
	let at = start
	let orgSeen = false

	for (let count = 0; at + 6 <= bytes.length && count < 100_000; count++) {
		const size = bytes.u32(at)
		const fn = bytes.u16(at + 4)

		if (size < 3 || fn === 0) break

		if (fn === 0x20b && !orgSeen) {
			org[0] = bytes.i16(at + 8)
			org[1] = bytes.i16(at + 6)
			orgSeen = true
		} else if (fn === 0x20c) {
			ext = [bytes.i16(at + 8), bytes.i16(at + 6)]
			break
		}

		at += size * 2
	}

	return { org, ext }
}

// An EMF carried in META_ESCAPE_ENHANCED_METAFILE comments ("WMFC" chunks), which players draw instead of the
// WMF records; undefined when there is none or its chunks do not add up.
function embeddedEmf(bytes: Bytes, start: number): Uint8Array | undefined {
	const parts: Uint8Array[] = []
	let declared = -1
	let expected = -1
	let received = 0
	let at = start

	for (let count = 0; at + 6 <= bytes.length && count < 1_000_000; count++) {
		const size = bytes.u32(at)
		const fn = bytes.u16(at + 4)

		if (size < 3 || fn === 0 || at + size * 2 > bytes.length) break

		const p = at + 6

		if (fn === 0x626 && bytes.u16(p) === 0x000f && bytes.u32(p + 4) === WMFC && bytes.u32(p + 8) === 1) {
			const total = bytes.u32(p + 34)
			const chunk = bytes.u32(p + 26)

			if (declared < 0) {
				// A declared size beyond the input cannot be filled by it.
				if (total === 0 || total > bytes.length) return undefined

				declared = total
				expected = bytes.u32(p + 22)
			}

			if (total !== declared || p + 38 + chunk > at + size * 2) return undefined

			parts.push(bytes.sub(p + 38, chunk))
			received += chunk

			if (parts.length === expected) break
		}

		at += size * 2
	}

	if (declared < 0 || received !== declared || parts.length !== expected) return undefined

	const out = new Uint8Array(declared)
	let offset = 0

	for (const part of parts) {
		out.set(part, offset)
		offset += part.length
	}

	return out
}

export function convertWmf(bytes: Bytes, nested: (bytes: Uint8Array) => MetafileConversion): MetafileConversion {
	let at = 0
	let placement: Placement | undefined

	if (bytes.u32(0) === PLACEABLE_KEY) {
		placement = { left: bytes.i16(6), top: bytes.i16(8), right: bytes.i16(10), bottom: bytes.i16(12), inch: bytes.u16(14) || 1440 }
		at = 22
	}

	const headerWords = bytes.u16(at + 2)

	if (headerWords < 9 || !bytes.has(at, 18)) return { ok: false, reason: "malformed" }

	const emf = embeddedEmf(bytes, at + headerWords * 2)

	if (emf !== undefined) {
		const conversion = nested(emf)

		if (conversion.ok) return conversion
	}

	at += headerWords * 2

	let width: number
	let height: number
	let windowOrg: [number, number]
	let windowExt: [number, number]
	let boundsFrame = false

	if (placement !== undefined && placement.right !== placement.left && placement.bottom !== placement.top) {
		width = (Math.abs(placement.right - placement.left) * 96) / placement.inch
		height = (Math.abs(placement.bottom - placement.top) * 96) / placement.inch
		windowOrg = [placement.left, placement.top]
		windowExt = [placement.right - placement.left, placement.bottom - placement.top]
	} else {
		const window = scanWindow(bytes, at)

		windowOrg = window.org

		if (window.ext !== undefined && window.ext[0] !== 0 && window.ext[1] !== 0) {
			// Without a placeable header the units are unknown: large extents are taken as twips.
			const unit = Math.max(Math.abs(window.ext[0]), Math.abs(window.ext[1])) > 4000 ? 15 : 1

			windowExt = window.ext
			width = Math.abs(window.ext[0]) / unit
			height = Math.abs(window.ext[1]) / unit
		} else {
			windowExt = [1, 1]
			width = 1
			height = 1
			boundsFrame = true
		}
	}

	const out = new Output(width, height)
	const gdi = new Gdi(out, { kind: "wmf", frameWidth: width, frameHeight: height }, [1, 0, 0, 1, 0, 0], windowOrg, windowExt)
	let truncated = false

	try {
		truncated = play(bytes, at, gdi)
	} catch (error) {
		if (!(error instanceof MetafileLimit)) throw error

		truncated = true
	}

	let frame: [number, number, number, number] = [0, 0, width, height]

	if (boundsFrame) {
		const bounds = out.list.bounds()

		frame = [
			bounds[0] ?? 0,
			bounds[1] ?? 0,
			Math.max(1, (bounds[2] ?? 0) - (bounds[0] ?? 0)),
			Math.max(1, (bounds[3] ?? 0) - (bounds[1] ?? 0))
		]
		out.list.width = frame[2]
		out.list.height = frame[3]
	}

	return { ok: true, format: "wmf", list: out.list, frame, unsupported: out.unsupported, truncated }
}

function points(bytes: Bytes, at: number, count: number): Float64Array {
	const out = new Float64Array(Math.max(0, count) * 2)

	for (let k = 0; k < count * 2; k++) out[k] = bytes.i16(at + k * 2)

	return out
}

// A Region object's scans as rectangles (left, top, right, bottom).
function regionRects(bytes: Bytes, at: number, end: number): number[] {
	const scans = bytes.i16(at + 10)
	const rects: number[] = []
	let p = at + 22

	for (let s = 0; s < scans && p + 6 <= end && rects.length < MAX_REGION_RECTS * 4; s++) {
		const count = bytes.u16(p)
		const top = bytes.i16(p + 2)
		const bottom = bytes.i16(p + 4)

		for (let k = 0; k + 1 < count && p + 6 + k * 2 + 4 <= end; k += 2)
			rects.push(bytes.i16(p + 6 + k * 2), top, bytes.i16(p + 8 + k * 2), bottom)

		p += 6 + count * 2 + 2
	}

	return rects
}

// Plays records from `start`; true when the stream ended without its EOF record.
function play(bytes: Bytes, start: number, gdi: Gdi): boolean {
	const out = gdi.out
	let at = start

	const create = (object: GdiObject): void => gdi.setObject(gdi.firstFree(), object)

	while (at + 6 <= bytes.length) {
		const size = bytes.u32(at)
		const fn = bytes.u16(at + 4)

		if (fn === 0) return false
		if (size < 3 || at + size * 2 > bytes.length) return true

		out.record()

		const p = at + 6
		const end = at + size * 2
		const i16 = (index: number): number => bytes.i16(p + index * 2)
		const u16 = (index: number): number => bytes.u16(p + index * 2)

		switch (fn) {
			case 0x201:
				gdi.setBkColor(gdi.color(bytes.u32(p)))
				break
			case 0x102:
				gdi.setBkMode(u16(0))
				break
			case 0x103:
				gdi.setMapMode(u16(0))
				break
			case 0x104:
				gdi.setRop2(u16(0))
				break
			case 0x106:
				gdi.setPolyFillMode(u16(0))
				break
			case 0x108:
				gdi.setCharExtra(i16(0))
				break
			case 0x209:
				gdi.setTextColor(gdi.color(bytes.u32(p)))
				break
			case 0x12e:
				gdi.setTextAlign(u16(0))
				break
			case 0x20b:
				gdi.setWindowOrg(i16(1), i16(0))
				break
			case 0x20c:
				gdi.setWindowExt(i16(1), i16(0))
				break
			case 0x20d:
				gdi.setViewportOrg(i16(1), i16(0))
				break
			case 0x20e:
				gdi.setViewportExt(i16(1), i16(0))
				break
			case 0x20f:
				gdi.offsetWindowOrg(i16(1), i16(0))
				break
			case 0x211:
				gdi.offsetViewportOrg(i16(1), i16(0))
				break
			case 0x410:
				gdi.scaleWindowExt(i16(3), i16(2), i16(1), i16(0))
				break
			case 0x412:
				gdi.scaleViewportExt(i16(3), i16(2), i16(1), i16(0))
				break
			case 0x213:
				gdi.lineTo(i16(1), i16(0))
				break
			case 0x214:
				gdi.moveTo(i16(1), i16(0))
				break
			case 0x415:
				gdi.excludeClipRect(i16(3), i16(2), i16(1), i16(0))
				break
			case 0x416:
				gdi.intersectClipRect(i16(3), i16(2), i16(1), i16(0))
				break
			case 0x220:
				gdi.offsetClip(i16(1), i16(0))
				break
			case 0x817:
			case 0x81a:
			case 0x830:
				gdi.arc(
					fn === 0x817 ? "arc" : fn === 0x81a ? "pie" : "chord",
					i16(7),
					i16(6),
					i16(5),
					i16(4),
					i16(3),
					i16(2),
					i16(1),
					i16(0)
				)
				break
			case 0x418:
				gdi.ellipse(i16(3), i16(2), i16(1), i16(0))
				break
			case 0x41b:
				gdi.rectangle(i16(3), i16(2), i16(1), i16(0))
				break
			case 0x61c:
				gdi.roundRect(i16(5), i16(4), i16(3), i16(2), i16(1), i16(0))
				break
			case 0x41f: {
				const x = bytes.i16(p + 6)
				const y = bytes.i16(p + 4)

				gdi.fillRect(x, y, x + 1, y + 1, gdi.color(bytes.u32(p)))
				break
			}
			case 0x324:
			case 0x325: {
				const count = Math.min(u16(0), Math.floor((end - p - 2) / 4))

				if (fn === 0x324) gdi.polygon(points(bytes, p + 2, count), count)
				else gdi.polyline(points(bytes, p + 2, count), count)

				break
			}
			case 0x538: {
				const polygons = u16(0)
				const counts: number[] = []
				let total = 0

				for (let k = 0; k < polygons && p + 2 + k * 2 < end; k++) {
					const count = u16(1 + k)

					counts.push(count)
					total += count
				}

				const first = p + 2 + polygons * 2

				if (first + total * 4 <= end) gdi.polyPolygon(points(bytes, first, total), counts)

				break
			}
			case 0x521: {
				const length = Math.min(u16(0), end - p - 2)
				const stringBytes = bytes.sub(p + 2, length)
				const after = p + 2 + length + (length & 1)

				gdi.text(bytes.i16(after + 2), bytes.i16(after), decodeAnsi(stringBytes, gdi.dc.font.charset), null, 1, 0, null)
				break
			}
			case 0xa32: {
				const y = i16(0)
				const x = i16(1)
				const length = Math.max(0, i16(2))
				const options = u16(3)
				let q = p + 8
				let rect: number[] | null = null

				if ((options & (ETO_OPAQUE | ETO_CLIPPED)) !== 0) {
					rect = [bytes.i16(q), bytes.i16(q + 2), bytes.i16(q + 4), bytes.i16(q + 6)]
					q += 8
				}

				const count = Math.min(length, end - q)
				const stringBytes = bytes.sub(q, count)
				const dxAt = q + count + (count & 1)
				let dx: Int16Array | null = null

				if (dxAt + count * 2 <= end && count > 0) {
					dx = new Int16Array(count)

					for (let k = 0; k < count; k++) dx[k] = bytes.i16(dxAt + k * 2)
				}

				const charset = gdi.dc.font.charset
				const text = decodeAnsi(stringBytes, charset)

				// A multi-byte code page decodes to fewer units than bytes: advances then follow the bytes.
				gdi.text(x, y, text, dx !== null && text.length !== count ? multiByteAdvances(stringBytes, text, dx) : dx, 1, options, rect)
				break
			}
			case 0x1e:
				gdi.saveDc()
				break
			case 0x127:
				gdi.restoreDc(i16(0))
				break
			case 0x12d:
				gdi.selectObject(gdi.objects.get(u16(0)))
				break
			case 0x1f0:
				gdi.deleteObject(u16(0))
				break
			case 0x12c: {
				const region = gdi.objects.get(u16(0))

				if (region?.kind === "region") gdi.selectLogicalRects(region.rects, 5)
				else gdi.resetClip()

				break
			}
			case 0x234:
				gdi.selectObject(gdi.objects.get(u16(0)))
				break
			case 0x2fa: {
				const style = u16(0)

				create(solidPen(style, Math.abs(i16(1)), gdi.color(bytes.u32(p + 6))))
				break
			}
			case 0x2fc:
				create(gdi.brush(u16(0), gdi.color(bytes.u32(p + 2)), u16(3)))
				break
			case 0x2fb: {
				let faceLength = 0

				while (faceLength < 32 && bytes.u8(p + 18 + faceLength) !== 0) faceLength++

				const charset = bytes.u8(p + 13)

				create({
					kind: "font",
					font: metafileFont(
						i16(0),
						i16(1),
						i16(2),
						i16(4),
						bytes.u8(p + 10) !== 0,
						bytes.u8(p + 11) !== 0,
						bytes.u8(p + 12) !== 0,
						charset,
						decodeAnsi(bytes.sub(p + 18, faceLength), charset === 2 ? 0 : charset)
					)
				})
				break
			}
			case 0xf7: {
				const count = Math.min(u16(1), Math.floor((end - p - 4) / 4))
				const colors = []

				// PALETTEENTRY: red, green, blue, flags.
				for (let k = 0; k < count; k++) colors.push(rgb(bytes.u8(p + 4 + k * 4), bytes.u8(p + 5 + k * 4), bytes.u8(p + 6 + k * 4)))

				create({ kind: "palette", colors })
				break
			}
			case 0x6ff:
				create({ kind: "region", rects: regionRects(bytes, p, end) })
				break
			case 0x142:
				// BS_PATTERN names a monochrome bitmap, drawn in the text and background colours.
				create(gdi.patternBrush(packedDib(bytes, p + 4, end - p - 4, u16(1)), u16(0) === 3))
				break
			case 0x1f9: {
				// A monochrome Bitmap16 (its 14-byte header, 18 reserved bytes, then rows of WidthBytes).
				const width = i16(1)
				const height = i16(2)
				const stride = i16(3)
				const bitsAt = p + 32

				if (
					bytes.u8(p + 8) !== 1 ||
					bytes.u8(p + 9) !== 1 ||
					width <= 0 ||
					height <= 0 ||
					stride * 8 < width ||
					bitsAt + stride * height > end ||
					!gdi.out.spend(width * height)
				) {
					gdi.out.unsupported++
					create(gdi.brush(1, 0, 0))
					break
				}

				const bits = new Uint8Array(width * height)

				for (let y = 0; y < height; y++)
					for (let x = 0; x < width; x++) bits[y * width + x] = (bytes.u8(bitsAt + y * stride + (x >> 3)) >> (7 - (x & 7))) & 1

				create(gdi.monoBrush(bits, width, height))
				break
			}
			case 0x228:
			case 0x12b: {
				const region = gdi.objects.get(u16(0))
				const brush = fn === 0x228 ? gdi.objects.get(u16(1)) : gdi.dc.brush

				if (region?.kind === "region" && brush?.kind === "brush") gdi.fillLogicalRects(region.rects, brush)

				break
			}
			case 0x61d:
				gdi.patBlt(i16(5), i16(4), i16(3), i16(2), bytes.u32(p))
				break
			case 0x940: {
				const rop = bytes.u32(p)

				if (size === (fn >>> 8) + 3) {
					gdi.patBlt(i16(8), i16(7), i16(6), i16(5), rop)
				} else {
					const dib = packedDib(bytes, p + 16, end - p - 16, 0)

					if (dib === undefined) out.unsupported++
					else gdi.bitmap(dib, `${p}`, i16(3), i16(2), i16(5), i16(4), i16(7), i16(6), i16(5), i16(4), rop)
				}

				break
			}
			case 0xb41: {
				const rop = bytes.u32(p)

				if (size === (fn >>> 8) + 3) {
					gdi.patBlt(i16(10), i16(9), i16(8), i16(7), rop)
				} else {
					const dib = packedDib(bytes, p + 20, end - p - 20, 0)

					if (dib === undefined) out.unsupported++
					else gdi.bitmap(dib, `${p}`, i16(5), i16(4), i16(3), i16(2), i16(9), i16(8), i16(7), i16(6), rop)
				}

				break
			}
			case 0xf43: {
				const dib = packedDib(bytes, p + 22, end - p - 22, u16(2))

				if (dib === undefined) out.unsupported++
				else
					gdi.bitmap(
						dib,
						`${p}`,
						i16(6),
						i16(5),
						i16(4),
						i16(3),
						i16(10),
						i16(9),
						i16(8),
						i16(7),
						bytes.u32(p),
						undefined,
						1,
						true
					)

				break
			}
			case 0xd33: {
				const dib = packedDib(bytes, p + 18, end - p - 18, u16(0))

				if (dib === undefined) out.unsupported++
				else gdi.bitmap(dib, `${p}`, i16(4), i16(3), i16(6), i16(5), i16(8), i16(7), i16(6), i16(5), 0x00cc0020, undefined, 1, true)

				break
			}
			case 0x922:
			case 0xb23:
				// BitBlt and StretchBlt: only their bitmap-less (pattern) form is drawn.
				if (fn === 0x922 && size === (fn >>> 8) + 3) gdi.patBlt(i16(8), i16(7), i16(6), i16(5), bytes.u32(p))
				else if (fn === 0xb23 && size === (fn >>> 8) + 3) gdi.patBlt(i16(10), i16(9), i16(8), i16(7), bytes.u32(p))
				else out.unsupported++

				break
			// Escapes (MathType's equation data among them), palette realisation, mapper flags, stretch mode,
			// layout and text justification change nothing that is drawn here.
			case 0x626:
			case 0x35:
			case 0x37:
			case 0x139:
			case 0x436:
			case 0x231:
			case 0x107:
			case 0x105:
			case 0x14c:
			case 0x20a:
				break
			default:
				out.unsupported++
				break
		}

		at = end
	}

	return true
}

// Advances per UTF-16 unit for multi-byte text: each decoded character takes the advances of its bytes.
function multiByteAdvances(stringBytes: Uint8Array, text: string, dx: Int16Array): Float64Array {
	const out = new Float64Array(text.length)
	let byte = 0

	for (let unit = 0; unit < text.length; unit++) {
		const code = text.charCodeAt(unit)
		// Lead bytes of the East Asian code pages are >= 0x81; single bytes are ASCII or half-width.
		const width = code > 0xff && (stringBytes[byte] ?? 0) >= 0x81 ? 2 : 1

		for (let k = 0; k < width && byte < dx.length; k++) out[unit] = (out[unit] ?? 0) + (dx[byte++] ?? 0)
	}

	return out
}
