import { describe, expect, it } from "vitest"
import {
	DARK_LUMA,
	FRAME_POSITIONS,
	displaySize,
	dominantColour,
	fitInBox,
	fitMode,
	meanLuma,
	nextFramePosition,
	sampleOf
} from "@/lib/media/framePick.logic"

function pixels(...colours: [number, number, number][]): Uint8ClampedArray {
	return Uint8ClampedArray.from(colours.flatMap(([r, g, b]) => [r, g, b, 255]))
}

describe("frame luma", () => {
	it("weighs channels by Rec. 709", () => {
		expect(meanLuma(pixels([255, 255, 255]))).toBeCloseTo(255)
		expect(meanLuma(pixels([0, 0, 0], [255, 255, 255]))).toBeCloseTo(127.5)
		expect(meanLuma(pixels([0, 255, 0]))).toBeCloseTo(0.7152 * 255)
		expect(meanLuma(new Uint8ClampedArray())).toBe(0)
	})

	it("tries 18%, then 28% and 10% only while the frame stays dark", () => {
		expect(FRAME_POSITIONS).toEqual([0.18, 0.28, 0.1])
		expect(nextFramePosition(0, DARK_LUMA)).toBeNull()
		expect(nextFramePosition(0, DARK_LUMA - 1)).toBe(0.28)
		expect(nextFramePosition(1, 3)).toBe(0.1)
		expect(nextFramePosition(2, 3)).toBeNull()
	})
})

describe("dominantColour", () => {
	it("takes the most populated colour bucket rather than the mean", () => {
		const image = pixels([200, 30, 30], [204, 28, 26], [198, 31, 31], [10, 10, 240], [240, 240, 240])

		expect(dominantColour(image)).toEqual([201, 30, 29])
		expect(sampleOf(image).dominant).toEqual([201, 30, 29])
	})

	it("answers black for no pixels", () => {
		expect(dominantColour(new Uint8ClampedArray())).toEqual([0, 0, 0])
	})
})

describe("frame geometry", () => {
	it("stretches by the pixel aspect ratio, then turns", () => {
		expect(displaySize({ width: 720, height: 576, par: 64 / 45, rotation: 0 })).toEqual({ width: 1024, height: 576 })
		expect(displaySize({ width: 1920, height: 1080, par: 1, rotation: 90 })).toEqual({ width: 1080, height: 1920 })
		expect(displaySize({ width: 1920, height: 1080, par: 0, rotation: 180 })).toEqual({ width: 1920, height: 1080 })
	})

	it("fits inside a box without upscaling", () => {
		expect(fitInBox({ width: 1920, height: 1080 }, { width: 480, height: 270 })).toEqual({ width: 480, height: 270 })
		expect(fitInBox({ width: 1080, height: 1920 }, { width: 480, height: 270 })).toEqual({ width: 152, height: 270 })
		expect(fitInBox({ width: 320, height: 180 }, { width: 480, height: 270 })).toEqual({ width: 320, height: 180 })
	})

	it("covers a box only when cropping takes at most about 5%", () => {
		expect(fitMode({ width: 480, height: 270 }, { width: 480, height: 270 })).toBe("cover")
		expect(fitMode({ width: 1920, height: 800 }, { width: 480, height: 270 })).toBe("contain")
		expect(fitMode({ width: 342, height: 500 }, { width: 342, height: 513 })).toBe("cover")
	})
})
