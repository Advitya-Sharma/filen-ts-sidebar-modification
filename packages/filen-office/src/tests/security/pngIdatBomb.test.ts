import { describe, expect, it } from "vitest"
import { zlibSync } from "fflate"
import type { DrawingImage } from "@filen/office/core/drawingml"
import { effectImage } from "../../core/metafile/effects"
import { png, zlibZeros } from "./generate"
import { budget } from "../timing"

// A picture recoloured by a colour change decodes its PNG: the inflate stops at the size the header
// declares, however far the compressed data would expand.

function recolour(bytes: Uint8Array): DrawingImage {
	const image = { partName: undefined, format: "png", bytes, colorChange: null } as unknown as DrawingImage
	const picture = {
		colorChange: [0x000000ff, 0xff0000ff],
		grayscale: false,
		biLevel: -1,
		duotone: null,
		brightness: 0,
		contrast: 0
	} as unknown as DrawingImage

	return effectImage(image, picture, { spend: () => true })
}

describe("PNG decoding in metafile pictures", () => {
	it("stops inflating IDAT data past the declared size", () => {
		// A 1x1 picture whose IDAT inflates to 1 GiB.
		const started = performance.now()
		const out = recolour(png(1, 1, zlibZeros(1024)))

		expect(performance.now() - started).toBeLessThan(budget(500))
		expect(out.format).toBe("png")
	})

	it("still recolours a well-formed picture", () => {
		// 2x1 RGBA: one black pixel, one white; filter byte 0.
		const out = recolour(png(2, 1, zlibSync(Uint8Array.of(0, 0, 0, 0, 255, 255, 255, 255, 255))))

		expect(out.bytes).toBeDefined()
		expect(out.bytes).not.toEqual(png(2, 1, zlibSync(Uint8Array.of(0, 0, 0, 0, 255, 255, 255, 255, 255))))
	})
})
