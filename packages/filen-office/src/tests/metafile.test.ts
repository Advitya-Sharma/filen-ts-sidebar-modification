import { gzipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { DrawingList, IDENTITY, ItemFloat, ItemInt, ItemKind, type DrawingImage } from "../core/drawingml/displayList"
import {
	convertMetafile,
	expandMetafiles,
	hasMetafileImages,
	inflateMetafile,
	metafileFormat,
	type MetafileConversion
} from "../core/metafile/index"
import {
	ByteWriter,
	EmfBuilder,
	bounds,
	brush,
	converted,
	dib24,
	emf,
	glyphs,
	plusFill,
	plusHeader,
	plusRecord,
	rectangle,
	wmf,
	wmfFontParams,
	wmfTextParams
} from "./metafileFixtures"

function close(actual: ArrayLike<number>, expected: readonly number[], digits: number = 2): void {
	expect(Array.from(actual).map(value => Number(value.toFixed(digits)))).toEqual(expected.map(value => Number(value.toFixed(digits))))
}

describe("WMF", () => {
	it("maps the window onto the placeable frame and fills with the selected brush", () => {
		const bytes = wmf()
			.recordBytes(0x2fc, new ByteWriter().u16(0).u32(0x0000ff).u16(0).done())
			.record(0x12d, [0])
			.recordBytes(0x2fa, new ByteWriter().u16(5).i16(0).i16(0).u32(0).done())
			.record(0x12d, [1])
			.record(0x41b, [360, 720, 0, 0])
			.build([0, 0, 1440, 720])
		const result = converted(bytes)
		const list = result.list

		expect(metafileFormat(bytes)).toBe("wmf")
		expect(result.format).toBe("wmf")
		close(result.frame, [0, 0, 96, 48])
		expect(list.itemCount).toBe(1)
		expect(bounds(list, 0)).toEqual([0, 0, 48, 24])
		expect(list.paints[list.itemInt(0, ItemInt.FILL)]).toEqual({ kind: "solid", color: 0xff0000ff })
		expect(list.itemInt(0, ItemInt.STROKE)).toBe(-1)
	})

	it("draws Symbol-font text as the Unicode characters it shows, at its advances", () => {
		const bytes = wmf()
			.recordBytes(0x2fb, wmfFontParams(-240, "Symbol", 2))
			.record(0x12d, [0])
			.record(0x12e, [24])
			.recordBytes(0xa32, wmfTextParams(150, 600, "ab\xb3", [120, 120, 120]))
			.build([0, 0, 1440, 720])
		const list = converted(bytes).list
		const [text] = glyphs(list)

		expect(text?.run.text).toBe("αβ≥")
		expect(text?.run.font).toBe("Times New Roman")
		expect(text?.run.size).toBeCloseTo(12)
		// Each glyph centred in its cell.
		close(text?.positions ?? [], [4, 12, 20])
		expect(text?.anchor).toBe("middle")
		expect(list.itemFloat(0, ItemFloat.E)).toBeCloseTo(10)
		expect(list.itemFloat(0, ItemFloat.F)).toBeCloseTo(40)
	})

	it("maps MT Extra and drops its pieces without a Unicode equivalent", () => {
		const bytes = wmf()
			.recordBytes(0x2fb, wmfFontParams(-240, "MT Extra", 2))
			.record(0x12d, [0])
			.recordBytes(0xa32, wmfTextParams(0, 0, "L1h", [100, 100, 100]))
			.build([0, 0, 1440, 720])

		expect(glyphs(converted(bytes).list)[0]?.run.text).toBe("⋯ℏ")
	})

	it("decodes 8-bit text through the font's code page and anchors text without advances", () => {
		const bytes = wmf()
			.recordBytes(0x2fb, wmfFontParams(-240, "Arial"))
			.record(0x12d, [0])
			.record(0x12e, [6])
			.recordBytes(0xa32, wmfTextParams(720, 360, "\x80uro", null))
			.build([0, 0, 1440, 720])
		const [text] = glyphs(converted(bytes).list)

		expect(text?.run.text).toBe("€uro")
		expect(text?.positions).toBeNull()
		expect(text?.anchor).toBe("middle")
	})

	it("clips to IntersectClipRect and restores the clip with RestoreDC", () => {
		const bytes = wmf()
			.record(0x1e, [])
			.record(0x416, [360, 720, 0, 0])
			.record(0x41b, [720, 1440, 0, 0])
			.record(0x127, [-1])
			.record(0x41b, [720, 1440, 0, 0])
			.build([0, 0, 1440, 720])
		const list = converted(bytes).list

		expect(list.itemCount).toBe(2)
		expect(list.itemClip(0)).toBe(0)
		expect(list.itemClip(1)).toBe(-1)
		expect(bounds(list, 0)).toEqual([0, 0, 48, 24])
	})

	it("fills pattern blits with the brush and embeds DIB blits as BMP images", () => {
		const { info, bits } = dib24(2, 2, 0x00ff00)
		const dib = new ByteWriter().raw(info).raw(bits).done()
		// DIBSTRETCHBLT: rop, src h, w, y, x, dest h, w, y, x, the DIB.
		const blit = new ByteWriter().u32(0x00cc0020).i16(2).i16(2).i16(0).i16(0).i16(360).i16(720).i16(0).i16(0).raw(dib).done()
		const bytes = wmf()
			.recordBytes(0x2fc, new ByteWriter().u16(0).u32(0x00ff0000).u16(0).done())
			.record(0x12d, [0])
			.record(0x61d, [0x0021, 0x00f0, 720, 1440, 0, 0])
			.recordBytes(0xb41, blit)
			.build([0, 0, 1440, 720])
		const list = converted(bytes).list

		expect(list.kind(0)).toBe(ItemKind.PATH)
		expect(list.paints[list.itemInt(0, ItemInt.FILL)]).toEqual({ kind: "solid", color: 0x0000ffff })
		expect(list.kind(1)).toBe(ItemKind.IMAGE)

		const image = list.images[list.itemInt(1, ItemInt.REF)]

		expect(image?.format).toBe("bmp")
		expect(image?.partName).toBeUndefined()
		expect(Array.from(image?.bytes?.subarray(0, 2) ?? [])).toEqual([0x42, 0x4d])
		expect(bounds(list, 1)).toEqual([0, 0, 48, 24])
	})
})

describe("EMF", () => {
	it("maps logical units through the mapping mode onto the frame", () => {
		const result = converted(rectangle(brush(emf(), 1, 0x0000ff), 0, 0, 500, 250).build())

		expect(result.format).toBe("emf")
		close(result.frame, [0, 0, 96, 96])
		expect(bounds(result.list, 0)).toEqual([0, 0, 48, 24])
	})

	it("applies the world transform", () => {
		const builder = brush(emf(), 1, 0x0000ff).record(36, writer => writer.f32(1).f32(0).f32(0).f32(1).f32(500).f32(0).u32(4))

		expect(bounds(converted(rectangle(builder, 0, 0, 100, 100).build()).list, 0)).toEqual([48, 0, 57.6, 9.6])
	})

	it("fills a path bracket without outlining it", () => {
		const builder = brush(emf(), 1, 0x00ff00)
			.record(59)
			.record(27, writer => writer.i32(0).i32(0))
			.record(54, writer => writer.i32(1000).i32(0))
			.record(54, writer => writer.i32(1000).i32(1000))
			.record(61)
			.record(60)
			.record(62, writer => writer.i32(0).i32(0).i32(0).i32(0))
		const list = converted(builder.build()).list

		expect(list.itemCount).toBe(1)
		expect(list.itemInt(0, ItemInt.STROKE)).toBe(-1)
		expect(list.paints[list.itemInt(0, ItemInt.FILL)]).toEqual({ kind: "solid", color: 0x00ff00ff })
		expect(bounds(list, 0)).toEqual([0, 0, 96, 96])
	})

	it("intersects clips, restores them with the device context and clears them with an empty region copy", () => {
		const builder = brush(emf(), 1, 0x0000ff)
			.record(30, writer => writer.i32(0).i32(0).i32(500).i32(500))
			.record(33)
			.record(30, writer => writer.i32(0).i32(0).i32(250).i32(1000))
		rectangle(builder, 0, 0, 1000, 1000)
		builder.record(34, writer => writer.i32(-1))
		rectangle(builder, 0, 0, 1000, 1000)
		builder.record(75, writer => writer.u32(0).u32(5))
		rectangle(builder, 0, 0, 1000, 1000)

		const list = converted(builder.build()).list

		expect(list.itemCount).toBe(3)
		// Rectangles intersect into one region.
		expect(list.clips[list.itemClip(0)]?.parent).toBe(-1)
		expect(bounds(list, 0)).toEqual([0, 0, 24, 48])
		expect(bounds(list, 1)).toEqual([0, 0, 48, 48])
		expect(list.itemClip(2)).toBe(-1)
	})

	it("chains clips that are not rectangles", () => {
		const builder = brush(emf(), 1, 0x0000ff)
			.record(59)
			.record(42, writer => writer.i32(0).i32(0).i32(1000).i32(1000))
			.record(60)
			.record(67, writer => writer.u32(5))
			.record(30, writer => writer.i32(0).i32(0).i32(500).i32(1000))
		rectangle(builder, 0, 0, 1000, 1000)

		const list = converted(builder.build()).list
		const clip = list.clips[list.itemClip(0)]

		expect(clip?.parent).toBeGreaterThanOrEqual(0)
		expect(list.clips[clip?.parent ?? -1]?.commandCount).toBeGreaterThan(5)
		expect(bounds(list, 0)).toEqual([0, 0, 48, 96])
	})

	it("draws ExtTextOutW with its advances, aligned to the top of the cell", () => {
		const builder = emf()
			.record(82, writer => writer.u32(1).i32(-250).i32(0).i32(0).i32(0).i32(700).u8(1).u8(1).u8(0).u8(0).u32(0).utf16("Arial", 32))
			.record(37, writer => writer.u32(1))
			.record(24, writer => writer.u32(0x0000ff))
			.record(84, writer =>
				writer
					.i32(0)
					.i32(0)
					.i32(0)
					.i32(0)
					.u32(1)
					.f32(0)
					.f32(0)
					.i32(100)
					.i32(100)
					.u32(2)
					.u32(76)
					.u32(0)
					.i32(0)
					.i32(0)
					.i32(0)
					.i32(0)
					.u32(80)
					.utf16("Hi")
					.i32(125)
					.i32(125)
			)
		const list = converted(builder.build()).list
		const [text] = glyphs(list)

		expect(text?.run).toMatchObject({ text: "Hi", bold: true, italic: true, underline: "sng", color: 0xff0000ff })
		expect(text?.run.size).toBeCloseTo(18)
		close(text?.positions ?? [], [6, 18])
		expect(list.itemFloat(0, ItemFloat.E)).toBeCloseTo(9.6)
		// The baseline sits the face's ascent below the reference point.
		expect(list.itemFloat(0, ItemFloat.F)).toBeGreaterThan(9.6 + 0.85 * 24)
	})

	it("skips glyph-index text and counts it as not drawn", () => {
		const builder = emf().record(84, writer =>
			writer
				.i32(0)
				.i32(0)
				.i32(0)
				.i32(0)
				.u32(1)
				.f32(0)
				.f32(0)
				.i32(0)
				.i32(0)
				.u32(1)
				.u32(76)
				.u32(0x10)
				.i32(0)
				.i32(0)
				.i32(0)
				.i32(0)
				.u32(0)
				.utf16("A")
		)
		const result = converted(builder.build())

		expect(result.list.itemCount).toBe(0)
		expect(result.unsupported).toBe(1)
	})

	it("embeds StretchDIBits bitmaps, as PNG when the pixels need decoding", () => {
		const { info, bits } = dib24(2, 2, 0xff0000)
		const stretch =
			(dibInfo: Uint8Array, dibBits: Uint8Array): ((writer: ByteWriter) => void) =>
			writer =>
				writer
					.i32(0)
					.i32(0)
					.i32(0)
					.i32(0)
					.i32(0)
					.i32(0)
					.i32(0)
					.i32(0)
					.i32(2)
					.i32(2)
					.u32(80)
					.u32(dibInfo.length)
					.u32(80 + dibInfo.length)
					.u32(dibBits.length)
					.u32(0)
					.u32(0x00cc0020)
					.i32(500)
					.i32(500)
					.raw(dibInfo)
					.raw(dibBits)
		const info32 = new ByteWriter().u32(40).i32(2).i32(2).u16(1).u16(32).u32(0).u32(0).u32(0).u32(0).u32(0).u32(0).done()
		const list = converted(
			emf()
				.record(81, stretch(info, bits))
				.record(81, stretch(info32, new Uint8Array(16).fill(0x80)))
				.build()
		).list

		expect(list.images.map(image => image.format)).toEqual(["bmp", "png"])
		expect(Array.from(list.images[1]?.bytes?.subarray(1, 4) ?? [])).toEqual([0x50, 0x4e, 0x47])
		expect(bounds(list, 0)).toEqual([0, 0, 48, 48])
	})

	it("fills BitBlt without a bitmap with the brush (PATCOPY)", () => {
		const builder = brush(emf(), 1, 0x00ffff).record(76, writer => {
			writer.i32(0).i32(0).i32(0).i32(0).i32(0).i32(0).i32(1000).i32(1000).u32(0x00f00021).i32(0).i32(0)

			for (let k = 0; k < 6; k++) writer.f32(k === 0 || k === 3 ? 1 : 0)

			writer.u32(0).u32(0).u32(0).u32(0).u32(0).u32(0)
		})
		const list = converted(builder.build()).list

		expect(list.paints[list.itemInt(0, ItemInt.FILL)]).toEqual({ kind: "solid", color: 0xffff00ff })
		expect(bounds(list, 0)).toEqual([0, 0, 96, 96])
	})

	it("draws hatched brushes as 8-pixel tiles at the brush origin over the background mode", () => {
		const builder = emf()
			.record(39, writer => writer.u32(1).u32(2).u32(0x0000ff).u32(5))
			.record(37, writer => writer.u32(1))
			.record(18, writer => writer.u32(2))
			.record(25, writer => writer.u32(0x00ff00))
			.record(13, writer => writer.i32(3).i32(4))
		const list = converted(rectangle(builder, 0, 0, 100, 100).build()).list
		const paint = list.paints[list.itemInt(0, ItemInt.FILL)]

		expect(paint?.kind).toBe("image")

		if (paint?.kind !== "image") return

		expect(paint.tile).toMatchObject({ width: 8, height: 8, matrix: [1, 0, 0, 1, 3, 4] })
		expect(list.images[paint.image]?.format).toBe("png")
	})
})

describe("EMF+", () => {
	it("draws a dual picture from its EMF+ records and skips the EMF ones", () => {
		const builder = new EmfBuilder().plus([plusHeader(true)]).plus([plusFill(0xff112233)])

		rectangle(
			brush(
				builder.record(17, writer => writer.u32(1)),
				1,
				0x0000ff
			),
			0,
			0,
			96,
			96
		)

		const result = converted(builder.build())

		expect(result.format).toBe("emf+")
		expect(result.list.itemCount).toBe(1)
		expect(result.list.paints[result.list.itemInt(0, ItemInt.FILL)]).toEqual({ kind: "solid", color: 0x112233ff })
		expect(bounds(result.list, 0)).toEqual([0, 0, 48, 48])
	})

	it("falls back to the EMF records when an EMF+ record cannot be drawn", () => {
		const builder = new EmfBuilder().plus([plusHeader(true)]).plus([plusFill(0xff112233), plusRecord(0x40ff, 0)])

		rectangle(brush(builder, 1, 0x0000ff), 0, 0, 96, 96)

		const result = converted(builder.build())

		expect(result.format).toBe("emf")
		expect(result.list.itemCount).toBe(1)
		expect(result.list.paints[result.list.itemInt(0, ItemInt.FILL)]).toEqual({ kind: "solid", color: 0xff0000ff })
	})

	it("draws the EMF records that follow GetDC", () => {
		const builder = new EmfBuilder().plus([plusHeader(true)]).plus([plusRecord(0x4004, 0)])

		rectangle(brush(builder, 1, 0x0000ff), 0, 0, 96, 96)

		const result = converted(builder.build())

		expect(result.format).toBe("emf+")
		expect(result.list.itemCount).toBe(1)
	})

	it("fills and strokes paths through object records and the world transform", () => {
		const path = plusRecord(0x4008, 0x0300, writer =>
			writer.u32(0xdbc01002).u32(3).u32(0).f32(0).f32(0).f32(10).f32(0).f32(10).f32(10).u8(0).u8(1).u8(0x81).pad(4)
		)
		const pen = plusRecord(0x4008, 0x0201, writer =>
			writer.u32(0xdbc01002).u32(0).u32(0).u32(0).f32(1).u32(0xdbc01002).u32(0).u32(0xff00ff00)
		)
		const builder = new EmfBuilder()
			.plus([plusHeader(false)])
			.plus([
				path,
				pen,
				plusRecord(0x402a, 0, writer => writer.f32(2).f32(0).f32(0).f32(2).f32(4).f32(4)),
				plusRecord(0x4014, 0x8000, writer => writer.u32(0xffff0000)),
				plusRecord(0x4015, 0, writer => writer.u32(1))
			])
		const list = converted(builder.build()).list

		expect(list.itemCount).toBe(2)
		expect(bounds(list, 0)).toEqual([4, 4, 24, 24])
		expect(list.strokes[list.itemInt(1, ItemInt.STROKE)]?.width).toBeCloseTo(2)
	})

	it("draws driver strings with Unicode lookup at their glyph positions", () => {
		const font = plusRecord(0x4008, 0x0602, writer => writer.u32(0xdbc01002).f32(12).u32(2).i32(0).u32(0).u32(6).utf16("Symbol"))
		const text = plusRecord(0x4036, 0x8002, writer =>
			writer.u32(0xff000000).u32(1).u32(0).u32(2).u16(0xf061).u16(0xf062).f32(10).f32(20).f32(18).f32(20)
		)
		const list = converted(
			new EmfBuilder()
				.plus([plusHeader(false)])
				.plus([font, text])
				.build()
		).list
		const [run] = glyphs(list)

		expect(run?.run.text).toBe("αβ")
		close(run?.positions ?? [], [4, 12])
	})
})

describe("metafile robustness", () => {
	it("refuses what is not a metafile and never throws on truncated or mutated input", () => {
		expect(convertMetafile(new Uint8Array(0))).toEqual({ ok: false, reason: "malformed" })
		expect(convertMetafile(new Uint8Array(64).fill(7))).toEqual({ ok: false, reason: "malformed" })

		const sources = [
			rectangle(brush(emf(), 1, 0x0000ff), 0, 0, 500, 250).build(),
			wmf()
				.recordBytes(0x2fb, wmfFontParams(-240, "Symbol", 2))
				.record(0x12d, [0])
				.recordBytes(0xa32, wmfTextParams(150, 600, "ab", [120, 120]))
				.build([0, 0, 1440, 720]),
			new EmfBuilder()
				.plus([plusHeader(true)])
				.plus([plusFill(0xff112233)])
				.build()
		]
		let seed = 12345
		const random = (): number => {
			seed = (seed * 1103515245 + 12345) & 0x7fffffff

			return seed / 0x7fffffff
		}

		for (const source of sources) {
			for (let cut = 0; cut < source.length; cut += 7) expect(() => convertMetafile(source.subarray(0, cut))).not.toThrow()

			for (let round = 0; round < 300; round++) {
				const mutated = source.slice()

				for (let k = 0; k < 4; k++) mutated[Math.floor(random() * mutated.length)] = Math.floor(random() * 256)

				const result = convertMetafile(mutated)

				expect(result.ok || result.reason !== "failed").toBe(true)
			}
		}
	})

	it("marks a stream without its EOF record as truncated", () => {
		const bytes = rectangle(brush(emf(), 1, 0x0000ff), 0, 0, 500, 250).build()
		const result = converted(bytes.subarray(0, bytes.length - 20))

		expect(result.truncated).toBe(true)
		expect(result.list.itemCount).toBe(1)
	})

	it("inflates gzip-compressed metafiles (.emz, .wmz) and refuses corrupt ones", () => {
		const bytes = rectangle(brush(emf(), 1, 0x0000ff), 0, 0, 500, 250).build()
		const compressed = gzipSync(bytes)

		expect(inflateMetafile(compressed)).toEqual(bytes)
		expect(converted(compressed).list.itemCount).toBe(1)
		expect(convertMetafile(compressed.subarray(0, 20))).toEqual({ ok: false, reason: "malformed" })
	})
})

function metafileImage(crop = { left: 0, top: 0, right: 0, bottom: 0 }): DrawingImage {
	return {
		partName: "/word/media/image1.wmf",
		format: "wmf",
		handling: "metafile",
		svgPartName: undefined,
		crop,
		alpha: 1,
		grayscale: false,
		biLevel: -1,
		duotone: null,
		brightness: 0,
		contrast: 0,
		colorChange: null
	}
}

describe("expandMetafiles", () => {
	const picture = (): Extract<MetafileConversion, { ok: true }> =>
		converted(
			wmf()
				.recordBytes(0x2fc, new ByteWriter().u16(0).u32(0x0000ff).u16(0).done())
				.record(0x12d, [0])
				.recordBytes(0x2fa, new ByteWriter().u16(5).i16(0).i16(0).u32(0).done())
				.record(0x12d, [1])
				.record(0x41b, [720, 1440, 0, 720])
				.build([0, 0, 1440, 720])
		)

	it("replaces each metafile image item by the picture, mapped onto its box and clipped to it", async () => {
		const host = new DrawingList()

		host.width = 200
		host.height = 100
		host.addItem(ItemKind.IMAGE, [1, 0, 0, 1, 10, 20], 0, 0, 100, 50, 0, 0, -1, -1, host.image(metafileImage()), -1, 0, 7)

		expect(hasMetafileImages(host)).toBe(true)

		const conversion = picture()
		const out = await expandMetafiles(host, async () => conversion)

		expect(out).not.toBe(host)
		expect(out.itemCount).toBe(1)
		expect(out.kind(0)).toBe(ItemKind.PATH)
		expect(out.itemInt(0, ItemInt.SOURCE)).toBe(7)
		// The right half of the picture, drawn into the right half of the box.
		expect(bounds(out, 0)).toEqual([60, 20, 110, 70])
		expect(out.itemClip(0)).toBeGreaterThanOrEqual(0)
		expect(hasMetafileImages(out)).toBe(false)
	})

	it("crops the picture as the image crops it", async () => {
		const host = new DrawingList()

		host.addItem(
			ItemKind.IMAGE,
			IDENTITY,
			0,
			0,
			100,
			50,
			0,
			0,
			-1,
			-1,
			host.image(metafileImage({ left: 0.5, top: 0, right: 0, bottom: 0 })),
			-1,
			0,
			0
		)

		const out = await expandMetafiles(host, async () => picture())

		// The cropped right half fills the whole box.
		expect(bounds(out, 0)).toEqual([0, 0, 100, 50])
	})

	it("keeps the list when nothing converts", async () => {
		const host = new DrawingList()

		host.addItem(ItemKind.IMAGE, IDENTITY, 0, 0, 100, 50, 0, 0, -1, -1, host.image(metafileImage()), -1, 0, 0)

		expect(await expandMetafiles(host, async () => ({ ok: false, reason: "malformed" }))).toBe(host)
	})
})

describe("DrawingList clips", () => {
	it("intersects clip bounds with their parents and item bounds with their clip, and compacts them", () => {
		const list = new DrawingList()
		const square = (x: number, y: number, size: number): [number, number] => {
			const command = list.commandCount
			const coord = list.coordCount

			list.moveTo(x, y)
			list.lineTo(x + size, y)
			list.lineTo(x + size, y + size)
			list.lineTo(x, y + size)
			list.close()

			return [command, coord]
		}
		const [c0, k0] = square(0, 0, 50)
		const outer = list.clip(IDENTITY, c0, k0, false, -1)
		const [c1, k1] = square(25, 25, 50)
		const inner = list.clip([2, 0, 0, 2, 0, 0], c1, k1, false, outer)
		const [c2, k2] = square(0, 0, 100)
		const item = list.addItem(ItemKind.PATH, IDENTITY, 0, 0, 100, 100, c2, k2, -1, -1, -1, -1, 0, 0)

		expect(Array.from(list.clipBounds.subarray(4, 8))).toEqual([50, 50, 50, 50])

		list.setItemClip(item, outer)

		expect(Array.from(list.itemBounds.subarray(0, 4))).toEqual([0, 0, 50, 50])
		expect(list.clips[inner]?.parent).toBe(outer)
		expect(list.compact().itemClips).toEqual(Int32Array.of(outer))
		expect(list.clip(IDENTITY, 0, 0, false, 99)).toBe(2)
		expect(list.clips[2]?.parent).toBe(-1)
	})
})
