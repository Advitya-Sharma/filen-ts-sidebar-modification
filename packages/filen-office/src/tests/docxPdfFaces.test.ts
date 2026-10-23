import { describe, expect, it } from "vitest"
import { FontManager } from "../core/fonts/index"

describe("PDF faces for installed fonts", () => {
	it("draws an installed font without a substitute with the bundled family its font-table hints pick", async () => {
		const manager = new FontManager(
			{ addBytes: () => Promise.resolve(true), addLocal: () => Promise.resolve(true), remove: () => undefined },
			{ read: () => Promise.resolve(new Uint8Array(4)) },
			undefined,
			"pdffaces"
		)

		manager.declare([
			{ name: "Arial Black", altNames: [], hints: { fontClass: "swiss", pitch: "variable", panose: undefined, charset: 0 } }
		])

		const request = { family: "Arial Black", bold: false, italic: false, slot: "ascii" } as const

		await manager.prepare([request])

		expect(manager.resolve(request).source).toBe("local")
		expect(manager.pdfFace(request)).toMatchObject({ kind: "bundled", family: "arimo" })
	})
})
