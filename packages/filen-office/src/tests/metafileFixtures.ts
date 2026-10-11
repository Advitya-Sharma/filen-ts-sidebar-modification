import type { DrawingList, DrawingText } from "../core/drawingml/displayList"
import { convertMetafile, type MetafileConversion } from "../core/metafile/index"

// Builders for synthetic EMF, EMF+ and WMF files.

export class ByteWriter {
	private bytes = new Uint8Array(256)
	length = 0

	private ensure(size: number): void {
		if (this.length + size <= this.bytes.length) return

		const grown = new Uint8Array(Math.max(this.bytes.length * 2, this.length + size))

		grown.set(this.bytes)
		this.bytes = grown
	}

	u8(value: number): this {
		this.ensure(1)
		this.bytes[this.length++] = value & 255

		return this
	}

	u16(value: number): this {
		this.ensure(2)
		new DataView(this.bytes.buffer).setUint16(this.length, value & 0xffff, true)
		this.length += 2

		return this
	}

	i16(value: number): this {
		return this.u16(value < 0 ? value + 0x10000 : value)
	}

	u32(value: number): this {
		this.ensure(4)
		new DataView(this.bytes.buffer).setUint32(this.length, value >>> 0, true)
		this.length += 4

		return this
	}

	i32(value: number): this {
		return this.u32(value >>> 0)
	}

	f32(value: number): this {
		this.ensure(4)
		new DataView(this.bytes.buffer).setFloat32(this.length, value, true)
		this.length += 4

		return this
	}

	raw(bytes: ArrayLike<number>): this {
		this.ensure(bytes.length)

		for (let k = 0; k < bytes.length; k++) this.bytes[this.length++] = (bytes[k] ?? 0) & 255

		return this
	}

	utf16(text: string, units: number = text.length): this {
		for (let k = 0; k < units; k++) this.u16(k < text.length ? text.charCodeAt(k) : 0)

		return this
	}

	ascii(text: string, size: number = text.length): this {
		for (let k = 0; k < size; k++) this.u8(k < text.length ? text.charCodeAt(k) : 0)

		return this
	}

	pad(multiple: number): this {
		while (this.length % multiple !== 0) this.u8(0)

		return this
	}

	set(at: number, value: number): void {
		new DataView(this.bytes.buffer).setUint32(at, value >>> 0, true)
	}

	done(): Uint8Array {
		return this.bytes.slice(0, this.length)
	}
}

// A placeable WMF: records are a function number and its 16-bit parameters (or raw parameter bytes).
export class WmfBuilder {
	private readonly records: Uint8Array[] = []

	record(fn: number, words: readonly number[]): this {
		const writer = new ByteWriter()

		writer.u32(3 + words.length).u16(fn)

		for (const word of words) writer.i16(word)

		this.records.push(writer.done())

		return this
	}

	recordBytes(fn: number, params: Uint8Array): this {
		const writer = new ByteWriter()
		const padded = params.length + (params.length & 1)

		writer
			.u32(3 + padded / 2)
			.u16(fn)
			.raw(params)
			.pad(2)
		this.records.push(writer.done())

		return this
	}

	build(bbox: readonly [number, number, number, number] | null, inch: number = 1440): Uint8Array {
		const writer = new ByteWriter()

		if (bbox !== null) {
			writer.u32(0x9ac6cdd7).u16(0).i16(bbox[0]).i16(bbox[1]).i16(bbox[2]).i16(bbox[3]).u16(inch).u32(0).u16(0)
		}

		const total = this.records.reduce((sum, record) => sum + record.length, 0) + 6

		writer
			.u16(1)
			.u16(9)
			.u16(0x300)
			.u32((18 + total) / 2)
			.u16(16)
			.u32(0)
			.u16(0)

		for (const record of this.records) writer.raw(record)

		writer.u32(3).u16(0)

		return writer.done()
	}
}

export function wmfFontParams(height: number, face: string, charset: number = 0, weight: number = 400): Uint8Array {
	const writer = new ByteWriter()

	writer.i16(height).i16(0).i16(0).i16(0).i16(weight).u8(0).u8(0).u8(0).u8(charset).u8(0).u8(0).u8(0).u8(0).ascii(face, 32)

	return writer.done()
}

export function wmfTextParams(x: number, y: number, text: string, dx: readonly number[] | null, options: number = 0): Uint8Array {
	const writer = new ByteWriter()

	writer.i16(y).i16(x).i16(text.length).u16(options).ascii(text).pad(2)

	if (dx !== null) for (const value of dx) writer.i16(value)

	return writer.done()
}

// An EMF: records are a type and a body writer; the header names the frame (.01 mm) and a 96 dpi device.
export class EmfBuilder {
	private readonly records: Uint8Array[] = []

	record(type: number, body: (writer: ByteWriter) => void = () => undefined): this {
		const writer = new ByteWriter()

		writer.u32(type).u32(0)
		body(writer)
		writer.pad(4)
		writer.set(4, writer.length)
		this.records.push(writer.done())

		return this
	}

	// An EMR_COMMENT carrying EMF+ records.
	plus(records: readonly Uint8Array[]): this {
		return this.record(70, writer => {
			const size = records.reduce((sum, record) => sum + record.length, 0) + 4

			writer.u32(size).u32(0x2b464d45)

			for (const record of records) writer.raw(record)
		})
	}

	build(frame: readonly [number, number, number, number] = [0, 0, 2540, 2540]): Uint8Array {
		const writer = new ByteWriter()
		const eof = new ByteWriter().u32(14).u32(20).u32(0).u32(0).u32(20).done()
		const total = 108 + this.records.reduce((sum, record) => sum + record.length, 0) + eof.length

		writer
			.u32(1)
			.u32(108)
			.i32(0)
			.i32(0)
			.i32(95)
			.i32(95)
			.i32(frame[0])
			.i32(frame[1])
			.i32(frame[2])
			.i32(frame[3])
			.u32(0x464d4520)
			.u32(0x10000)
			.u32(total)
			.u32(this.records.length + 2)
			.u16(16)
			.u16(0)
			.u32(0)
			.u32(0)
			.u32(0)
			.i32(960)
			.i32(960)
			.i32(254)
			.i32(254)
			.u32(0)
			.u32(0)
			.u32(0)
			.i32(254_000)
			.i32(254_000)

		for (const record of this.records) writer.raw(record)

		writer.raw(eof)

		return writer.done()
	}
}

// One EMF+ record.
export function plusRecord(type: number, flags: number, body: (writer: ByteWriter) => void = () => undefined): Uint8Array {
	const data = new ByteWriter()

	body(data)

	const bytes = data.done()

	return new ByteWriter()
		.u16(type)
		.u16(flags)
		.u32(12 + bytes.length)
		.u32(bytes.length)
		.raw(bytes)
		.done()
}

// A packed 24-bit bottom-up DIB of `width` x `height` with every pixel `rgb` (0xRRGGBB).
export function dib24(width: number, height: number, rgb: number): { info: Uint8Array; bits: Uint8Array } {
	const info = new ByteWriter().u32(40).i32(width).i32(height).u16(1).u16(24).u32(0).u32(0).u32(0).u32(0).u32(0).u32(0).done()
	const stride = Math.ceil((width * 3) / 4) * 4
	const bits = new Uint8Array(stride * height)

	for (let y = 0; y < height; y++) {
		for (let x = 0; x < width; x++) {
			bits[y * stride + x * 3] = rgb & 255
			bits[y * stride + x * 3 + 1] = (rgb >>> 8) & 255
			bits[y * stride + x * 3 + 2] = (rgb >>> 16) & 255
		}
	}

	return { info, bits }
}

export function converted(bytes: Uint8Array): Extract<MetafileConversion, { ok: true }> {
	const result = convertMetafile(bytes)

	if (!result.ok) throw new Error(`conversion failed: ${result.reason}`)

	return result
}

export function bounds(list: DrawingList, item: number): number[] {
	return Array.from(list.itemBounds.subarray(item * 4, item * 4 + 4), value => Math.round(value * 100) / 100)
}

export function glyphs(list: DrawingList): Extract<DrawingText, { kind: "glyphs" }>[] {
	return list.texts.filter((text): text is Extract<DrawingText, { kind: "glyphs" }> => text.kind === "glyphs")
}

// A 1440-unit-per-inch WMF of 1 x 0.5 inch with a black text colour and transparent backgrounds.
export function wmf(): WmfBuilder {
	return new WmfBuilder().record(0x20b, [0, 0]).record(0x20c, [720, 1440]).record(0x102, [1])
}

// An EMF of one inch at 96 device pixels per inch, logical units 1/1000 inch.
export function emf(): EmfBuilder {
	return new EmfBuilder()
		.record(17, writer => writer.u32(8))
		.record(9, writer => writer.i32(1000).i32(1000))
		.record(11, writer => writer.i32(96).i32(96))
		.record(18, writer => writer.u32(1))
}

export function brush(builder: EmfBuilder, index: number, color: number): EmfBuilder {
	return builder
		.record(39, writer => writer.u32(index).u32(0).u32(color).u32(0))
		.record(37, writer => writer.u32(index))
		.record(37, writer => writer.u32(0x80000008))
}

export function rectangle(builder: EmfBuilder, left: number, top: number, right: number, bottom: number): EmfBuilder {
	return builder.record(43, writer => writer.i32(left).i32(top).i32(right).i32(bottom))
}

export function plusHeader(dual: boolean): Uint8Array {
	return plusRecord(0x4001, dual ? 1 : 0, writer => writer.u32(0xdbc01002).u32(1).u32(96).u32(96))
}

export function plusFill(color: number): Uint8Array {
	return plusRecord(0x400a, 0x8000, writer => writer.u32(color).u32(1).f32(0).f32(0).f32(48).f32(48))
}
