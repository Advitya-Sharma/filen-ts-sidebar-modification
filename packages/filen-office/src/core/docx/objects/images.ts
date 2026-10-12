import { decodeUtf8 } from "../../xml/index"

// What a picture's bytes are, read from their header only (nothing is decoded): the format, the pixel size
// and the resolution a new picture is sized by. Only formats Word reads as pictures are taken; WebP and
// anything unknown are refused rather than written into a file Word may not show.

export type DocxImageFormat = "png" | "jpeg" | "gif" | "bmp" | "tiff" | "svg" | "emf" | "wmf"

export interface DocxImageInfo {
	readonly format: DocxImageFormat
	readonly contentType: string
	// The extension of the media part (Word's own names).
	readonly extension: string
	// Pixels, or for vector formats the size at 96 dpi.
	readonly width: number
	readonly height: number
	// Dots per inch the file asks for (96 where it names none).
	readonly dpiX: number
	readonly dpiY: number
	// Vector formats give their size in EMUs themselves.
	readonly emuWidth?: number
	readonly emuHeight?: number
}

export type DocxImageRefusal = "format" | "webp" | "corrupt" | "tooLarge" | "pixels" | "fallback"

// Limits of a new picture: larger ones are refused, never downscaled (that would lose detail).
export const MAX_IMAGE_BYTES = 25 * 1024 * 1024
export const MAX_IMAGE_PIXELS = 20_000_000

const DEFAULT_DPI = 96
export const EMU_PER_INCH = 914400

const TYPES: Readonly<Record<DocxImageFormat, readonly [string, string]>> = {
	png: ["image/png", "png"],
	jpeg: ["image/jpeg", "jpeg"],
	gif: ["image/gif", "gif"],
	bmp: ["image/bmp", "bmp"],
	tiff: ["image/tiff", "tiff"],
	svg: ["image/svg+xml", "svg"],
	emf: ["image/x-emf", "emf"],
	wmf: ["image/x-wmf", "wmf"]
}

function u16be(b: Uint8Array, at: number): number {
	return ((b[at] ?? 0) << 8) | (b[at + 1] ?? 0)
}

function u16le(b: Uint8Array, at: number): number {
	return (b[at] ?? 0) | ((b[at + 1] ?? 0) << 8)
}

function u32be(b: Uint8Array, at: number): number {
	return (((b[at] ?? 0) << 24) | ((b[at + 1] ?? 0) << 16) | ((b[at + 2] ?? 0) << 8) | (b[at + 3] ?? 0)) >>> 0
}

function u32le(b: Uint8Array, at: number): number {
	return ((b[at] ?? 0) | ((b[at + 1] ?? 0) << 8) | ((b[at + 2] ?? 0) << 16) | ((b[at + 3] ?? 0) << 24)) >>> 0
}

function i32le(b: Uint8Array, at: number): number {
	return u32le(b, at) | 0
}

function i16le(b: Uint8Array, at: number): number {
	const value = u16le(b, at)

	return value >= 0x8000 ? value - 0x10000 : value
}

function starts(b: Uint8Array, at: number, ascii: string): boolean {
	for (let k = 0; k < ascii.length; k++) if (b[at + k] !== ascii.charCodeAt(k)) return false

	return true
}

// A resolution in dots per inch, or 96 when it is none a picture would be sized by.
function dpi(value: number): number {
	return Number.isFinite(value) && value >= 1 && value <= 10_000 ? value : DEFAULT_DPI
}

function info(format: DocxImageFormat, width: number, height: number, dpiX = DEFAULT_DPI, dpiY = DEFAULT_DPI): DocxImageInfo {
	const [contentType, extension] = TYPES[format]

	return { format, contentType, extension, width, height, dpiX: dpi(dpiX), dpiY: dpi(dpiY) }
}

function png(b: Uint8Array): DocxImageInfo | undefined {
	if (b.length < 24 || !starts(b, 12, "IHDR")) return undefined

	const width = u32be(b, 16)
	const height = u32be(b, 20)
	let dpiX = DEFAULT_DPI
	let dpiY = DEFAULT_DPI

	// pHYs comes before the first IDAT.
	for (let at = 8; at + 8 <= b.length;) {
		const length = u32be(b, at)

		if (starts(b, at + 4, "IDAT") || starts(b, at + 4, "IEND")) break
		if (starts(b, at + 4, "pHYs") && length >= 9 && at + 17 <= b.length && b[at + 16] === 1) {
			dpiX = u32be(b, at + 8) * 0.0254
			dpiY = u32be(b, at + 12) * 0.0254
		}

		at += 12 + length
	}

	return info("png", width, height, dpiX, dpiY)
}

function jpeg(b: Uint8Array): DocxImageInfo | undefined {
	let dpiX = DEFAULT_DPI
	let dpiY = DEFAULT_DPI

	for (let at = 2; at + 4 <= b.length;) {
		if (b[at] !== 0xff) return undefined

		const marker = b[at + 1] ?? 0

		// Fill bytes and markers without a length.
		if (marker === 0xff) {
			at++

			continue
		}

		if (marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd7) || marker === 0x01) {
			at += 2

			continue
		}

		const length = u16be(b, at + 2)

		if (length < 2) return undefined

		// APP0 JFIF: units 1 dots per inch, 2 per centimetre.
		if (marker === 0xe0 && starts(b, at + 4, "JFIF\0") && length >= 14) {
			const units = b[at + 11] ?? 0
			const x = u16be(b, at + 12)
			const y = u16be(b, at + 14)

			if (units === 1) {
				dpiX = x
				dpiY = y
			} else if (units === 2) {
				dpiX = x * 2.54
				dpiY = y * 2.54
			}
		}

		// Start of frame (any but DHT, JPG and DAC): height, then width.
		if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
			if (at + 9 > b.length) return undefined

			return info("jpeg", u16be(b, at + 7), u16be(b, at + 5), dpiX, dpiY)
		}

		if (marker === 0xda || marker === 0xd9) return undefined

		at += 2 + length
	}

	return undefined
}

function gif(b: Uint8Array): DocxImageInfo | undefined {
	return b.length < 10 ? undefined : info("gif", u16le(b, 6), u16le(b, 8))
}

function bmp(b: Uint8Array): DocxImageInfo | undefined {
	if (b.length < 26) return undefined

	const header = u32le(b, 14)

	// BITMAPCOREHEADER: 16-bit sizes, no resolution.
	if (header === 12) return info("bmp", u16le(b, 18), u16le(b, 20))
	if (header < 40 || b.length < 46) return undefined

	const perMeter = (at: number): number => (u32le(b, at) > 0 ? u32le(b, at) * 0.0254 : DEFAULT_DPI)

	return info("bmp", Math.abs(i32le(b, 18)), Math.abs(i32le(b, 22)), perMeter(38), perMeter(42))
}

function tiff(b: Uint8Array): DocxImageInfo | undefined {
	const little = b[0] === 0x49
	const u16 = (at: number): number => (little ? u16le(b, at) : u16be(b, at))
	const u32 = (at: number): number => (little ? u32le(b, at) : u32be(b, at))
	const ifd = u32(4)

	if (ifd < 8 || ifd + 2 > b.length) return undefined

	const count = u16(ifd)
	let width = 0
	let height = 0
	let resolutionX = 0
	let resolutionY = 0
	let unit = 2

	for (let k = 0; k < count; k++) {
		const entry = ifd + 2 + k * 12

		if (entry + 12 > b.length) return undefined

		const tag = u16(entry)
		const type = u16(entry + 2)
		const value = type === 3 ? u16(entry + 8) : u32(entry + 8)
		const rational = (): number => {
			const at = u32(entry + 8)

			return at + 8 <= b.length && u32(at + 4) > 0 ? u32(at) / u32(at + 4) : 0
		}

		if (tag === 256) width = value
		else if (tag === 257) height = value
		else if (tag === 282 && type === 5) resolutionX = rational()
		else if (tag === 283 && type === 5) resolutionY = rational()
		else if (tag === 296) unit = value
	}

	const scale = unit === 3 ? 2.54 : unit === 2 ? 1 : 0

	return info("tiff", width, height, resolutionX * scale, resolutionY * scale)
}

// EMF: the frame (hundredths of a millimetre) is its size; the bounds (device pixels) its pixels.
function emf(b: Uint8Array): DocxImageInfo | undefined {
	if (b.length < 88) return undefined

	const pixelsX = i32le(b, 16) - i32le(b, 8)
	const pixelsY = i32le(b, 20) - i32le(b, 12)
	const frameX = i32le(b, 32) - i32le(b, 24)
	const frameY = i32le(b, 36) - i32le(b, 28)

	if (frameX <= 0 || frameY <= 0) return undefined

	const emuWidth = Math.round(frameX * 360)
	const emuHeight = Math.round(frameY * 360)

	return {
		...info("emf", Math.max(1, pixelsX + 1), Math.max(1, pixelsY + 1)),
		emuWidth,
		emuHeight
	}
}

// Placeable WMF: a bounding box in units of which `inch` make an inch.
function wmf(b: Uint8Array): DocxImageInfo | undefined {
	if (b.length < 22) return undefined

	const inch = u16le(b, 14)
	const width = i16le(b, 10) - i16le(b, 6)
	const height = i16le(b, 12) - i16le(b, 8)

	if (inch <= 0 || width <= 0 || height <= 0) return undefined

	const emuWidth = Math.round((width / inch) * EMU_PER_INCH)
	const emuHeight = Math.round((height / inch) * EMU_PER_INCH)

	return {
		...info("wmf", Math.max(1, Math.round((width / inch) * DEFAULT_DPI)), Math.max(1, Math.round((height / inch) * DEFAULT_DPI))),
		emuWidth,
		emuHeight
	}
}

const SVG_LENGTH = /^\s*([0-9]*\.?[0-9]+(?:e[-+]?[0-9]+)?)\s*(px|pt|pc|in|cm|mm|)\s*$/i
const PX_PER_UNIT: Readonly<Record<string, number>> = { "": 1, px: 1, pt: 96 / 72, pc: 16, in: 96, cm: 96 / 2.54, mm: 96 / 25.4 }

function svgLength(tag: string, name: string): number | undefined {
	const value = new RegExp(`\\s${name}\\s*=\\s*("([^"]*)"|'([^']*)')`).exec(tag)
	const text = value?.[2] ?? value?.[3]
	const length = text === undefined ? null : SVG_LENGTH.exec(text)

	if (length === null) return undefined

	const px = Number(length[1]) * (PX_PER_UNIT[(length[2] ?? "").toLowerCase()] ?? 1)

	return Number.isFinite(px) && px > 0 ? px : undefined
}

// SVG: its width and height (in CSS units), else its view box, else 300 x 150 (a browser's default).
function svg(b: Uint8Array): DocxImageInfo | undefined {
	const head = decodeUtf8(b, 0, Math.min(b.length, 64 * 1024))
	const open = /<svg[\s>]/.exec(head)

	if (open === null || /<(?!\?|!|svg)[A-Za-z]/.test(head.slice(0, open.index))) return undefined

	const end = head.indexOf(">", open.index)
	const tag = end < 0 ? head.slice(open.index) : head.slice(open.index, end)
	const box = /\sviewBox\s*=\s*["']\s*[-0-9.e]+[\s,]+[-0-9.e]+[\s,]+([0-9.e]+)[\s,]+([0-9.e]+)\s*["']/i.exec(tag)
	const boxWidth = box === null ? undefined : Number(box[1])
	const boxHeight = box === null ? undefined : Number(box[2])
	let width = svgLength(tag, "width")
	let height = svgLength(tag, "height")

	if (width === undefined && height !== undefined && boxWidth !== undefined && boxHeight !== undefined && boxHeight > 0)
		width = (height * boxWidth) / boxHeight
	if (height === undefined && width !== undefined && boxWidth !== undefined && boxHeight !== undefined && boxWidth > 0)
		height = (width * boxHeight) / boxWidth

	width ??= boxWidth !== undefined && boxWidth > 0 ? boxWidth : 300
	height ??= boxHeight !== undefined && boxHeight > 0 ? boxHeight : 150

	return info("svg", Math.max(1, Math.round(width)), Math.max(1, Math.round(height)))
}

// The picture `bytes` hold, or why it cannot be inserted.
export function sniffImage(bytes: Uint8Array): DocxImageInfo | DocxImageRefusal {
	if (bytes.length > MAX_IMAGE_BYTES) return "tooLarge"

	let found: DocxImageInfo | undefined

	if (starts(bytes, 0, "\x89PNG\r\n\x1a\n")) found = png(bytes)
	else if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) found = jpeg(bytes)
	else if (starts(bytes, 0, "GIF87a") || starts(bytes, 0, "GIF89a")) found = gif(bytes)
	else if (starts(bytes, 0, "BM")) found = bmp(bytes)
	else if (starts(bytes, 0, "II*\0") || starts(bytes, 0, "MM\0*")) found = tiff(bytes)
	else if (starts(bytes, 0, "RIFF") && starts(bytes, 8, "WEBP")) return "webp"
	else if (u32le(bytes, 0) === 1 && starts(bytes, 40, " EMF")) found = emf(bytes)
	else if (u32le(bytes, 0) === 0x9ac6cdd7) found = wmf(bytes)
	else if (bytes.length > 0 && (bytes[0] === 0x3c || bytes[0] === 0xef || bytes[0] === 0x20 || bytes[0] === 0x0a || bytes[0] === 0x0d))
		found = svg(bytes)
	else return "format"

	if (found === undefined) return "corrupt"
	if (found.width < 1 || found.height < 1) return "corrupt"
	if (found.width * found.height > MAX_IMAGE_PIXELS) return "pixels"

	return found
}

// The size (EMUs) Word gives a picture inserted at its own resolution: pixels at the file's dpi, a vector
// picture's own size.
export function naturalExtent(image: DocxImageInfo): { cx: number; cy: number } {
	if (image.emuWidth !== undefined && image.emuHeight !== undefined) return { cx: image.emuWidth, cy: image.emuHeight }

	return {
		cx: Math.max(1, Math.round((image.width * EMU_PER_INCH) / image.dpiX)),
		cy: Math.max(1, Math.round((image.height * EMU_PER_INCH) / image.dpiY))
	}
}
