import { readFileSync } from "node:fs"
import { gunzipSync } from "node:zlib"
import { describe, expect, it } from "vitest"
import type { FontAssets, FontFaceDescriptor, FontHost } from "../core/fonts/host"
import { FontManager, type FontRequest } from "../core/fonts/manager"

// FontManager.pdfFace: what a PDF export embeds for what layout drew.

class Host implements FontHost {
	private readonly installed: readonly string[]

	constructor(installed: readonly string[] = []) {
		this.installed = installed
	}

	addBytes(_face: FontFaceDescriptor, _data: Uint8Array): Promise<boolean> {
		return Promise.resolve(true)
	}

	addLocal(_face: FontFaceDescriptor, name: string): Promise<boolean> {
		return Promise.resolve(this.installed.includes(name))
	}

	remove(): void {}
}

const assets: FontAssets = { read: () => Promise.resolve(new Uint8Array(4)) }

function caladea(fsType: number): Uint8Array {
	const bytes = new Uint8Array(gunzipSync(readFileSync(new URL("../../assets/fonts/caladea/Caladea-Regular.ttf.gz", import.meta.url))))
	const view = new DataView(bytes.buffer)

	for (let t = 0; t < view.getUint16(4); t++) {
		if (String.fromCharCode(...bytes.subarray(12 + t * 16, 16 + t * 16)) === "OS/2")
			view.setUint16(view.getUint32(20 + t * 16) + 8, fsType)
	}

	return bytes
}

function request(family: string, bold = false, italic = false): FontRequest {
	return { family, bold, italic, slot: "hAnsi" }
}

describe("FontManager.pdfFace", () => {
	it("names the bundled substitute a family is drawn with, in the drawn weight and style", async () => {
		const manager = new FontManager(new Host(), assets)

		await manager.prepare([request("Calibri", true), request("Times New Roman", false, true)])

		expect(manager.pdfFace(request("Calibri", true))).toEqual({ kind: "bundled", family: "carlito", weight: 700, italic: false })
		expect(manager.pdfFace(request("Times New Roman", false, true))).toEqual({
			kind: "bundled",
			family: "tinos",
			weight: 400,
			italic: true
		})
	})

	it("hands over an embedded face's reader with its fsType and the substitute for when it may not be embedded", async () => {
		const manager = new FontManager(new Host(), assets)
		const data = caladea(0x0002)

		manager.embed("Cambria", false, false, () => Promise.resolve(data))
		await manager.prepare([request("Cambria")])

		const face = manager.pdfFace(request("Cambria"))

		expect(face.kind).toBe("embedded")

		if (face.kind !== "embedded") return

		expect(face.fsType).toBe(0x0002)
		expect(face.substitute).toBe("caladea")
		expect(await face.read(1 << 20)).toBe(data)
	})

	it("embeds the bundled substitute in place of an installed font", async () => {
		const manager = new FontManager(new Host(["Aptos"]), assets)

		await manager.prepare([request("Aptos")])

		expect(manager.pdfFace(request("Aptos"))).toEqual({ kind: "bundled", family: "carlito", weight: 400, italic: false })
	})

	it("falls back by the generic family for fonts it knows nothing about", async () => {
		const manager = new FontManager(new Host(), assets)

		await manager.prepare([request("Courier Prime")])

		expect(manager.pdfFace(request("Courier Prime")).kind).toBe("bundled")
		expect(manager.pdfFace(request("Courier Prime"))).toMatchObject({ family: "cousine" })
	})
})
