import { describe, expect, it } from "vitest"
import { gzipSync } from "fflate"
import { MAX_METAFILE_BYTES, convertMetafile, inflateMetafile } from "../../core/metafile/index"
import { gzipZeros } from "./generate"

// .emz / .wmz pictures are gzip inside the zip: their inflate is bounded by the metafile cap as it runs,
// not after the whole stream is in memory.

describe("compressed metafiles", () => {
	it("stops inflating a gzip bomb at the metafile cap", () => {
		// 1 MiB of gzip that inflates to 1 GiB.
		const bomb = gzipZeros(1024)
		const before = process.memoryUsage().arrayBuffers
		const started = performance.now()

		expect(inflateMetafile(bomb)).toBeUndefined()
		expect(convertMetafile(bomb)).toEqual({ ok: false, reason: "malformed" })

		const elapsed = performance.now() - started
		const grown = process.memoryUsage().arrayBuffers - before

		expect(elapsed).toBeLessThan(1000)
		expect(grown).toBeLessThan(4 * MAX_METAFILE_BYTES)
	})

	it("still inflates a metafile below the cap", () => {
		const bytes = new Uint8Array(3 * 1024 * 1024).map((_, k) => (k * 7) & 0xff)
		const inflated = inflateMetafile(gzipSync(bytes, { mtime: 0 }))

		expect(inflated !== undefined && Buffer.from(inflated).equals(bytes)).toBe(true)
	})
})
