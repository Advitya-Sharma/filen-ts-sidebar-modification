import { describe, expect, it } from "vitest"
import { BLOCK_BYTES } from "@filen/shared/videos"
import { LARGE_CLOSED_RANGE_BYTES, parseRange, previewPieceEnd, previewServedEnd, PREVIEW_PIECE_BYTES } from "@/sw/range"

describe("parseRange", () => {
	it("marks only a start-only range as open-ended", () => {
		expect(parseRange("bytes=100-", 1000)).toEqual({ start: 100, end: 999, openEnded: true, suffix: false })
		expect(parseRange("bytes=100-199", 1000)).toEqual({ start: 100, end: 199, openEnded: false, suffix: false })
		expect(parseRange("bytes=-100", 1000)).toEqual({ start: 900, end: 999, openEnded: false, suffix: true })
	})

	it("rejects a malformed range or one that starts at or past the end", () => {
		expect(parseRange("bytes=-", 1000)).toBeNull()
		expect(parseRange("bytes=1000-", 1000)).toBeNull()
		expect(parseRange("bytes=5-2", 1000)).toBeNull()
		expect(parseRange("items=0-1", 1000)).toBeNull()
		expect(parseRange("bytes=0-0", 0)).toBeNull()
		expect(parseRange("bytes=-5", 0)).toBeNull()
	})

	it("clamps a last byte past the end to the last byte, as RFC 9110 asks", () => {
		expect(parseRange("bytes=100-5000", 1000)).toEqual({ start: 100, end: 999, openEnded: false, suffix: false })
		expect(parseRange(`bytes=0-${"9".repeat(400)}`, 1000)).toEqual({ start: 0, end: 999, openEnded: false, suffix: false })
	})

	it("answers a suffix longer than the file with the whole file", () => {
		expect(parseRange("bytes=-5000", 1000)).toEqual({ start: 0, end: 999, openEnded: false, suffix: true })
	})

	it("serves a one-byte file", () => {
		expect(parseRange("bytes=0-0", 1)).toEqual({ start: 0, end: 0, openEnded: false, suffix: false })
		expect(parseRange("bytes=0-1", 1)).toEqual({ start: 0, end: 0, openEnded: false, suffix: false })
		expect(parseRange("bytes=0-", 1)).toEqual({ start: 0, end: 0, openEnded: true, suffix: false })
		expect(parseRange("bytes=-1", 1)).toEqual({ start: 0, end: 0, openEnded: false, suffix: true })
		expect(parseRange("bytes=1-1", 1)).toBeNull()
	})
})

describe("previewPieceEnd", () => {
	const total = 20 * BLOCK_BYTES

	it("ends a piece on an SDK chunk boundary, so the next one starts on one", () => {
		expect(previewPieceEnd(0, total - 1)).toBe(PREVIEW_PIECE_BYTES - 1)
		expect(previewPieceEnd(PREVIEW_PIECE_BYTES, total - 1)).toBe(2 * PREVIEW_PIECE_BYTES - 1)
		// A seek lands mid-chunk: the piece still ends on the boundary, a little short.
		expect(previewPieceEnd(BLOCK_BYTES + 5, total - 1)).toBe(BLOCK_BYTES + PREVIEW_PIECE_BYTES - 1)
	})

	it("never runs past what was asked for", () => {
		expect(previewPieceEnd(total - 10, total - 1)).toBe(total - 1)
	})
})

describe("previewServedEnd", () => {
	const total = 40 * BLOCK_BYTES
	const range = (header: string) => {
		const parsed = parseRange(header, total)

		if (parsed === null) {
			throw new Error(`unparseable ${header}`)
		}

		return parsed
	}

	it("pieces an open-ended range", () => {
		expect(previewServedEnd(range("bytes=0-"), total)).toBe(PREVIEW_PIECE_BYTES - 1)
	})

	it("pieces a closed range that runs to the last byte, as WebKit asks for the rest of the file", () => {
		expect(previewServedEnd(range(`bytes=0-${String(total - 1)}`), total)).toBe(PREVIEW_PIECE_BYTES - 1)
		expect(previewServedEnd(range(`bytes=${String(2 * BLOCK_BYTES)}-${String(total - 1)}`), total)).toBe(
			2 * BLOCK_BYTES + PREVIEW_PIECE_BYTES - 1
		)
	})

	it("pieces any other closed range longer than the large-range bound", () => {
		expect(previewServedEnd(range(`bytes=0-${String(LARGE_CLOSED_RANGE_BYTES)}`), total)).toBe(PREVIEW_PIECE_BYTES - 1)
	})

	it("answers a small closed read exactly", () => {
		expect(previewServedEnd(range("bytes=0-1"), total)).toBe(1)
		expect(previewServedEnd(range(`bytes=100-${String(LARGE_CLOSED_RANGE_BYTES + 99)}`), total)).toBe(LARGE_CLOSED_RANGE_BYTES + 99)
		// A suffix range ends at the last byte but is never pieced: piecing would cut off the tail it asked for.
		expect(previewServedEnd(range("bytes=-500"), total)).toBe(total - 1)
		expect(previewServedEnd(range(`bytes=-${String(10 * BLOCK_BYTES + 7)}`), total)).toBe(total - 1)
	})
})
