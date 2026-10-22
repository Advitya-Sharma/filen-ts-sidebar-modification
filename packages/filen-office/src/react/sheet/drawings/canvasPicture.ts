import { ItemFloat, ItemInt, ItemKind, ITEM_FLOATS, ITEM_INTS, type DrawingImage } from "../../../core/drawingml/displayList"
import type { SheetDrawingData } from "../../../core/xlsx/drawings/render"

// A plain raster picture drawn on a canvas of its box's device pixels: decoded at that size (createImageBitmap's
// resize), drawn, and the decode closed at once. The canvas holds only the part of the box near the view (a region), and
// its pixels go when the object is out of view. Chromium keeps
// an SVG <image>'s decode for the document's life, so pictures drawn this way cost memory only while they show.
//
// Only a drawing that is exactly one picture over its whole frame qualifies: no crop, recolouring, transparency,
// outline, shadow, clip or transform, and a PNG, JPEG or BMP (a GIF may be animated, which a canvas would freeze).
// Everything else keeps the SVG path, which draws all of it.

const RASTER = new Set(["png", "jpeg", "bmp"])
const EPSILON = 1e-3

export type CanvasPicture =
	| { readonly kind: "part"; readonly partName: string; readonly type: string }
	| { readonly kind: "bytes"; readonly bytes: Uint8Array; readonly type: string }

const TYPES: Readonly<Record<string, string>> = { png: "image/png", jpeg: "image/jpeg", bmp: "image/bmp" }

function plain(image: DrawingImage): boolean {
	const { crop } = image

	return (
		RASTER.has(image.format) &&
		image.svgPartName === undefined &&
		crop.left === 0 &&
		crop.top === 0 &&
		crop.right === 0 &&
		crop.bottom === 0 &&
		image.alpha === 1 &&
		!image.grayscale &&
		image.biLevel < 0 &&
		image.duotone === null &&
		image.brightness === 0 &&
		image.contrast === 0 &&
		image.colorChange === null
	)
}

function near(a: number, b: number): boolean {
	return Math.abs(a - b) < EPSILON
}

// The picture a drawing is, when the canvas path draws it exactly; null otherwise.
export function canvasPicture(drawing: SheetDrawingData): CanvasPicture | null {
	if (drawing.itemCount !== 1 || drawing.itemClips !== undefined) return null
	if (drawing.itemInts.length < ITEM_INTS || drawing.itemFloats.length < ITEM_FLOATS) return null

	const ints = drawing.itemInts
	const floats = drawing.itemFloats

	if (
		ints[ItemInt.KIND] !== ItemKind.IMAGE ||
		ints[ItemInt.COMMAND_COUNT] !== 0 ||
		ints[ItemInt.FILL] !== -1 ||
		ints[ItemInt.STROKE] !== -1 ||
		ints[ItemInt.SHADOW] !== -1 ||
		ints[ItemInt.FLAGS] !== 0
	)
		return null

	const identity =
		near(floats[ItemFloat.A] ?? 0, 1) &&
		near(floats[ItemFloat.B] ?? 1, 0) &&
		near(floats[ItemFloat.C] ?? 1, 0) &&
		near(floats[ItemFloat.D] ?? 0, 1) &&
		near(floats[ItemFloat.E] ?? 1, 0) &&
		near(floats[ItemFloat.F] ?? 1, 0)
	const whole =
		near(floats[ItemFloat.X] ?? 1, 0) &&
		near(floats[ItemFloat.Y] ?? 1, 0) &&
		near(floats[ItemFloat.WIDTH] ?? 0, drawing.width) &&
		near(floats[ItemFloat.HEIGHT] ?? 0, drawing.height)

	if (!identity || !whole) return null

	const image = drawing.images[ints[ItemInt.REF] ?? -1]

	if (image === undefined || !plain(image)) return null

	const type = TYPES[image.format] ?? ""

	if (image.bytes !== undefined) return image.partName === undefined ? { kind: "bytes", bytes: image.bytes, type } : null

	return image.partName === undefined ? null : { kind: "part", partName: image.partName, type }
}

const DECODES_AT_ONCE = 2
let decoding = 0
const waiting: (() => void)[] = []

function decodeTurn(): Promise<void> {
	if (decoding < DECODES_AT_ONCE) {
		decoding++

		return Promise.resolve()
	}

	return new Promise(resolve => {
		waiting.push(resolve)
	})
}

function decodeDone(): void {
	const next = waiting.shift()

	if (next === undefined) decoding--
	else next()
}

// A part of an object's box, CSS px from its top left: what of a picture its canvas holds.
export interface PictureRegion {
	readonly x: number
	readonly y: number
	readonly width: number
	readonly height: number
}

// Regions are snapped out to this many CSS px, so a scroll changes them (and decodes again) only now and then.
const REGION_STEP = 128

// What of a box at (`x`, `y`) on screen, `width` x `height`, lies within the view widened by `margin` times its size,
// snapped out to steps and kept inside the box; null when none of it does.
export function regionIn(
	x: number,
	y: number,
	width: number,
	height: number,
	view: { readonly width: number; readonly height: number },
	margin: number
): PictureRegion | null {
	const left = Math.max(x, -view.width * margin)
	const top = Math.max(y, -view.height * margin)
	const right = Math.min(x + width, view.width * (1 + margin))
	const bottom = Math.min(y + height, view.height * (1 + margin))

	if (right <= left || bottom <= top) return null

	const x0 = Math.max(0, Math.floor((left - x) / REGION_STEP) * REGION_STEP)
	const y0 = Math.max(0, Math.floor((top - y) / REGION_STEP) * REGION_STEP)
	const x1 = Math.min(width, Math.ceil((right - x) / REGION_STEP) * REGION_STEP)
	const y1 = Math.min(height, Math.ceil((bottom - y) / REGION_STEP) * REGION_STEP)

	return { x: x0, y: y0, width: x1 - x0, height: y1 - y0 }
}

export function regionContains(outer: PictureRegion, inner: PictureRegion): boolean {
	return (
		inner.x >= outer.x &&
		inner.y >= outer.y &&
		inner.x + inner.width <= outer.x + outer.width &&
		inner.y + inner.height <= outer.y + outer.height
	)
}

export function sameRegion(a: PictureRegion | undefined, b: PictureRegion | undefined): boolean {
	return a === b || (a !== undefined && b !== undefined && a.x === b.x && a.y === b.y && a.width === b.width && a.height === b.height)
}

// The source pixels of a `natural`-sized picture stretched over a `box` that `region` shows, snapped out to whole
// pixels, and the region those map back to exactly (so the canvas lands on the pixels the whole picture would).
export function sourceOf(
	region: PictureRegion,
	box: { readonly width: number; readonly height: number },
	natural: { readonly width: number; readonly height: number }
): { readonly sx: number; readonly sy: number; readonly sw: number; readonly sh: number; readonly region: PictureRegion } {
	const kx = natural.width / box.width
	const ky = natural.height / box.height
	const sx = Math.max(0, Math.floor(region.x * kx))
	const sy = Math.max(0, Math.floor(region.y * ky))
	const sw = Math.min(natural.width, Math.ceil((region.x + region.width) * kx)) - sx
	const sh = Math.min(natural.height, Math.ceil((region.y + region.height) * ky)) - sy

	return { sx, sy, sw, sh, region: { x: sx / kx, y: sy / ky, width: sw / kx, height: sh / ky } }
}

// Decodes `blob` at `width` x `height` device pixels, draws it over `canvas` and closes the decode at once (a bitmap
// renderer would keep the transferred decode until garbage collection, long after the canvas let it go); false when
// the platform cannot or `stale` says the drawing is no longer wanted. The canvas keeps what it showed until then.
// `source`: the part of the picture to decode (source pixels), all of it without.
export async function drawPicture(
	canvas: HTMLCanvasElement,
	blob: Blob,
	width: number,
	height: number,
	stale: () => boolean,
	source?: { readonly sx: number; readonly sy: number; readonly sw: number; readonly sh: number }
): Promise<boolean> {
	if (typeof createImageBitmap !== "function") return false

	// Decodes take turns, a few at a time: each holds a whole decoded picture until it is cut and scaled, and a scroll
	// over many pictures would otherwise hold all of theirs at once.
	await decodeTurn()

	if (stale()) {
		decodeDone()

		return false
	}

	let bitmap: ImageBitmap

	try {
		const options: ImageBitmapOptions = { resizeWidth: width, resizeHeight: height, resizeQuality: "high" }
		bitmap =
			source === undefined
				? await createImageBitmap(blob, options)
				: await createImageBitmap(blob, source.sx, source.sy, source.sw, source.sh, options)
	} finally {
		decodeDone()
	}

	try {
		if (stale()) return false

		canvas.width = width
		canvas.height = height

		const context = canvas.getContext("2d")

		if (context === null) return false

		context.imageSmoothingQuality = "high"
		context.drawImage(bitmap, 0, 0, width, height)

		return true
	} finally {
		bitmap.close()
	}
}

// The canvas's pixels released (an object out of view).
export function releaseCanvas(canvas: HTMLCanvasElement): void {
	canvas.width = 0
	canvas.height = 0
}
