import { describe, expect, it } from "vitest"
import { BlockSource, TS_SNIFF_BYTES, bytesReadRange, transportStreamOfHead, walkContainer } from "@filen/shared/videos"

const MIB = 1024 * 1024

// A PAT section naming `programs` programs (program 0, the network PID, counts as none).
function pat(programs: number[]): number[] {
	const entries = programs.flatMap(program => [program >> 8, program & 0xff, 0xe0 | 0x01, 0x00])
	const length = 5 + entries.length + 4

	return [0x00, 0xb0 | (length >> 8), length & 0xff, 0x00, 0x01, 0xc1, 0x00, 0x00, ...entries, 0, 0, 0, 0]
}

function packet(pid: number, payload: number[], unitStart: boolean): number[] {
	const bytes = new Array<number>(188).fill(0xff)

	bytes[0] = 0x47
	bytes[1] = (unitStart ? 0x40 : 0) | (pid >> 8)
	bytes[2] = pid & 0xff
	bytes[3] = 0x10

	const body = unitStart ? [0x00, ...payload] : payload

	body.forEach((value, index) => {
		bytes[4 + index] = value
	})

	return bytes
}

function stream(stride: number, prefix: number, first: number[], count = 8): Uint8Array {
	const bytes = new Uint8Array(prefix + stride * count)

	for (let k = 0; k < count; k++) {
		bytes.set(k === 0 ? first : packet(0x1fff, [], false), prefix + k * stride)
	}

	return bytes
}

const NULL_PACKET = packet(0x1fff, [], false)

function head(bytes: Uint8Array): Uint8Array {
	return bytes.subarray(0, TS_SNIFF_BYTES)
}

describe("transportStreamOfHead", () => {
	it("reads the three packet layouts", () => {
		expect(transportStreamOfHead(head(stream(188, 0, NULL_PACKET)), MIB)).toEqual({ stride: 188, prefix: 0 })
		expect(transportStreamOfHead(head(stream(192, 4, NULL_PACKET)), MIB)).toEqual({ stride: 192, prefix: 4 })
		expect(transportStreamOfHead(head(stream(204, 0, NULL_PACKET)), MIB)).toEqual({ stride: 204, prefix: 0 })
	})

	it("accepts a PAT naming one program and rejects one naming several, as the walk does", async () => {
		const single = stream(188, 0, packet(0, pat([0, 1]), true))
		const several = stream(188, 0, packet(0, pat([1, 2]), true))

		expect(transportStreamOfHead(head(single), single.length)).toEqual({ stride: 188, prefix: 0 })
		expect(transportStreamOfHead(head(several), several.length)).toBeNull()

		const walked = await walkContainer(new BlockSource(several.length, bytesReadRange(several)), "card")

		expect(walked).toBeNull()
	})

	it("is never TypeScript, a Matroska or an ISO BMFF file", () => {
		const source = new TextEncoder().encode("export const x = 1\n".repeat(60))
		const matroska = stream(192, 4, NULL_PACKET)
		const mp4 = stream(188, 0, NULL_PACKET)

		matroska.set([0x1a, 0x45, 0xdf, 0xa3])
		mp4.set(new TextEncoder().encode("ftyp"), 4)

		expect(transportStreamOfHead(source, MIB)).toBeNull()
		expect(transportStreamOfHead(head(matroska), MIB)).toBeNull()
		expect(transportStreamOfHead(head(mp4), MIB)).toBeNull()
	})

	it("needs a sync byte on every one of the first packets, and at least three of them", () => {
		const broken = stream(188, 0, NULL_PACKET)

		broken[188 * 2] = 0x48

		expect(transportStreamOfHead(head(broken), MIB)).toBeNull()
		expect(transportStreamOfHead(stream(188, 0, NULL_PACKET, 3), 188 * 3)).toEqual({ stride: 188, prefix: 0 })
		expect(transportStreamOfHead(stream(188, 0, NULL_PACKET, 2), 188 * 2)).toBeNull()
		expect(transportStreamOfHead(new Uint8Array(7), 7)).toBeNull()
	})

	it("reads its decision from the first TS_SNIFF_BYTES", () => {
		expect(TS_SNIFF_BYTES).toBe(817)
	})
})
