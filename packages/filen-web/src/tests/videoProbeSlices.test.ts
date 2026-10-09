import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { BlockSource } from "@/lib/media/blockSource"
import { initialBlockReads, probeVideo } from "@/lib/media/probe/probe"
import type { ProbeTier } from "@/lib/media/probe/types"
import {
	CARD_BUDGET,
	MEDIA_DIR,
	SLICES_DIR,
	TIERS,
	fileReadRange,
	loadSlices,
	measureProbe,
	mergeRanges,
	sliceReadRange,
	type TierExpectation,
	type VideoSlice
} from "@/tests/support/videoProbe"

// The probe replayed against sparse slices of real muxer output (see support/videoProbe.ts). Regenerate
// with UPDATE_VIDEO_SLICES=1 after a parser change, once the media fixtures exist locally.

// Multi-hour and multi-gigabit files add size, not coverage the short ones lack.
const NOT_SLICED = new Set(["long-faststart.mp4", "long-tailmoov.mp4", "long.mkv", "h264-longgop-1080p.mkv"])

describe.skipIf(process.env["UPDATE_VIDEO_SLICES"] !== "1" || !existsSync(join(MEDIA_DIR, "manifest.json")))(
	"video probe slice writer",
	() => {
		it("records what every tier reads of each fixture", async () => {
			const manifest = JSON.parse(readFileSync(join(MEDIA_DIR, "manifest.json"), "utf8")) as { name: string }[]

			mkdirSync(SLICES_DIR, { recursive: true })

			for (const { name } of manifest) {
				if (NOT_SLICED.has(name)) {
					continue
				}

				const file = await fileReadRange(join(MEDIA_DIR, name))

				try {
					const consumed: [number, number][] = []
					const measure = async (tier: ProbeTier): Promise<TierExpectation> => {
						const measured = await measureProbe(file.size, file.readRange, tier)

						consumed.push(...measured.ranges)

						return { record: measured.record, blockReads: measured.blockReads }
					}
					const tiers = { card: await measure("card"), full: await measure("full"), cover: await measure("cover") }

					const ranges: [number, string][] = []

					for (const [offset, length] of mergeRanges(consumed)) {
						ranges.push([offset, Buffer.from(await file.readRange(offset, offset + length)).toString("base64")])
					}

					const slice: VideoSlice = { name, size: file.size, ranges, tiers }

					writeFileSync(join(SLICES_DIR, `${name}.json`), `${JSON.stringify(slice, null, "\t")}\n`)
				} finally {
					await file.close()
				}
			}
		})
	}
)

describe("video probe over recorded slices", () => {
	for (const slice of loadSlices()) {
		for (const tier of TIERS) {
			it(`${slice.name}: ${tier}`, async () => {
				const expected = slice.tiers[tier]
				const source = new BlockSource(slice.size, sliceReadRange(slice), { maxBlockReads: initialBlockReads(tier) })

				expect(await probeVideo(source, tier)).toEqual(expected.record)
				expect(source.blockReads).toBe(expected.blockReads)

				if (tier === "card") {
					expect(source.blockReads).toBeLessThanOrEqual(CARD_BUDGET)
				}
			})
		}
	}
})
