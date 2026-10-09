import { describe, expect, it } from "vitest"
import { isInitSegment, scanBoxes, splitForAppend } from "@/lib/media/fmp4Split"

function box(type: string, payload: number): Uint8Array {
	const bytes = new Uint8Array(8 + payload)
	const view = new DataView(bytes.buffer)

	view.setUint32(0, bytes.length)

	for (let index = 0; index < 4; index++) {
		bytes[4 + index] = type.charCodeAt(index)
	}

	return bytes
}

function concat(...parts: Uint8Array[]): Uint8Array {
	const out = new Uint8Array(parts.reduce((total, part) => total + part.length, 0))
	let offset = 0

	for (const part of parts) {
		out.set(part, offset)
		offset += part.length
	}

	return out
}

function lengths(pieces: Uint8Array[]): number[] {
	return pieces.map(piece => piece.length)
}

describe("scanBoxes", () => {
	it("lists top-level boxes, 64-bit sizes included", () => {
		const large = box("mdat", 8)
		const view = new DataView(large.buffer)

		view.setUint32(0, 1)
		view.setBigUint64(8, BigInt(large.length))

		const bytes = concat(box("moof", 100), large)

		expect(scanBoxes(bytes)).toEqual([
			{ type: "moof", start: 0, size: 108 },
			{ type: "mdat", start: 108, size: 16 }
		])
	})

	it("stops at a truncated header", () => {
		expect(scanBoxes(concat(box("moof", 10), new Uint8Array(5)))).toHaveLength(1)
	})
})

describe("isInitSegment", () => {
	it("tells an init segment from a media segment", () => {
		expect(isInitSegment(concat(box("ftyp", 16), box("moov", 100)))).toBe(true)
		expect(isInitSegment(concat(box("moof", 16), box("mdat", 100)))).toBe(false)
	})
})

describe("splitForAppend", () => {
	it("keeps bytes that fit one piece whole", () => {
		const bytes = concat(box("moof", 100), box("mdat", 1000))

		expect(splitForAppend(bytes, 2000)).toEqual([bytes])
	})

	it("cuts at box boundaries where the boxes allow it", () => {
		const bytes = concat(box("moof", 92), box("mdat", 892), box("moof", 92), box("mdat", 892))

		expect(lengths(splitForAppend(bytes, 1200))).toEqual([1100, 900])
	})

	it("cuts a box larger than a piece into near-equal pieces, without copying", () => {
		const bytes = concat(box("moof", 92), box("mdat", 2492))
		const pieces = splitForAppend(bytes, 1000)

		expect(lengths(pieces)).toEqual([100, 834, 834, 832])
		expect(pieces.every(piece => piece.buffer === bytes.buffer)).toBe(true)
		expect(concat(...pieces)).toEqual(bytes)
	})

	it("covers trailing bytes that are no whole box", () => {
		const bytes = concat(box("moof", 2992), new Uint8Array(3))
		const pieces = splitForAppend(bytes, 1000)

		expect(concat(...pieces)).toEqual(bytes)
		expect(pieces.every(piece => piece.length <= 1000)).toBe(true)
	})
})
