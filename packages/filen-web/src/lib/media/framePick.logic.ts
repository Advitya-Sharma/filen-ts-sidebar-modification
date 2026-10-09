// Picking and sizing a representative frame of a video: where to look, whether what was found is too dark
// to show, its dominant colour, and its display size. Plain and DOM-free.

// Where a representative frame is looked for, as fractions of the duration: the first position, then the
// others in order only while the frame found is too dark (an opening fade, a night scene).
export const FRAME_POSITIONS: readonly number[] = [0.18, 0.28, 0.1]

// Mean luma (0-255) below which a frame counts as black. A heuristic; HDR frames read washed out anyway.
export const DARK_LUMA = 16

// The downsample luma and colour are measured on: small enough to cost nothing, large enough to be stable.
export const SAMPLE_WIDTH = 64
export const SAMPLE_HEIGHT = 36

export type Rgb = [number, number, number]

// Mean Rec. 709 luma of RGBA pixels.
export function meanLuma(rgba: Uint8ClampedArray): number {
	const pixels = Math.floor(rgba.length / 4)

	if (pixels === 0) {
		return 0
	}

	let sum = 0

	for (let at = 0; at + 3 < rgba.length; at += 4) {
		sum += 0.2126 * (rgba[at] ?? 0) + 0.7152 * (rgba[at + 1] ?? 0) + 0.0722 * (rgba[at + 2] ?? 0)
	}

	return sum / pixels
}

export function isDark(luma: number): boolean {
	return luma < DARK_LUMA
}

// The mean colour of the most populated coarse colour bucket (3 bits per channel), so a poster's large
// background wins over a mean that would blend every colour in it into grey.
export function dominantColour(rgba: Uint8ClampedArray): Rgb {
	const counts = new Uint32Array(512)
	const sums = new Float64Array(512 * 3)

	for (let at = 0; at + 3 < rgba.length; at += 4) {
		const r = rgba[at] ?? 0
		const g = rgba[at + 1] ?? 0
		const b = rgba[at + 2] ?? 0
		const bucket = ((r >> 5) << 6) | ((g >> 5) << 3) | (b >> 5)

		counts[bucket] = (counts[bucket] ?? 0) + 1
		sums[bucket * 3] = (sums[bucket * 3] ?? 0) + r
		sums[bucket * 3 + 1] = (sums[bucket * 3 + 1] ?? 0) + g
		sums[bucket * 3 + 2] = (sums[bucket * 3 + 2] ?? 0) + b
	}

	let best = 0

	for (let bucket = 1; bucket < 512; bucket++) {
		if ((counts[bucket] ?? 0) > (counts[best] ?? 0)) {
			best = bucket
		}
	}

	const count = counts[best] ?? 0

	if (count === 0) {
		return [0, 0, 0]
	}

	return [
		Math.round((sums[best * 3] ?? 0) / count),
		Math.round((sums[best * 3 + 1] ?? 0) / count),
		Math.round((sums[best * 3 + 2] ?? 0) / count)
	]
}

export interface FrameSample {
	luma: number
	dominant: Rgb
}

export function sampleOf(rgba: Uint8ClampedArray): FrameSample {
	return { luma: meanLuma(rgba), dominant: dominantColour(rgba) }
}

// The next position to try after a frame measured `luma` at attempt `attempt` (0-based), or null to
// stop: a bright enough frame ends the search, and so does running out of positions.
export function nextFramePosition(attempt: number, luma: number): number | null {
	if (!isDark(luma)) {
		return null
	}

	return FRAME_POSITIONS[attempt + 1] ?? null
}

export interface Size {
	width: number
	height: number
}

// The picture as displayed: coded pixels stretched by the pixel aspect ratio, then turned by `rotation`
// (clockwise degrees, a right angle).
export function displaySize({ width, height, par, rotation }: { width: number; height: number; par: number; rotation: number }): Size {
	const stretched = { width: Math.max(1, Math.round(width * (par > 0 ? par : 1))), height: Math.max(1, height) }

	return rotation % 180 === 0 ? stretched : { width: stretched.height, height: stretched.width }
}

// `size` scaled down to fit inside `box`, never up.
export function fitInBox(size: Size, box: Size): Size {
	const scale = Math.min(1, box.width / size.width, box.height / size.height)

	return { width: Math.max(1, Math.round(size.width * scale)), height: Math.max(1, Math.round(size.height * scale)) }
}

// Whether a picture of `size` fills `box` once cropped by at most 5%: cover it then, else show it whole.
export function fitMode(size: Size, box: Size): "cover" | "contain" {
	const ratio = size.width / size.height / (box.width / box.height)

	return ratio > 0.95 && ratio < 1.0526 ? "cover" : "contain"
}
