import { readFileSync, readdirSync } from "node:fs"
import { open } from "node:fs/promises"
import { join } from "node:path"
import {
	BLOCK_BYTES,
	BlockSource,
	type BlockSourceOptions,
	type ReadRange,
	initialBlockReads,
	probeVideo,
	type ProbeTier,
	type VideoProbe
} from "@filen/shared/videos"

// Sparse slices: the exact byte ranges a probe consumed from a real file, so committed tests replay the
// probe against real muxer output without committing the media. Everything outside the ranges reads as
// zeros, which the parser never looks at, so its reads (and block downloads) replay exactly.

export const TIERS: readonly ProbeTier[] = ["card", "full", "cover"]

export const SLICES_DIR = join(__dirname, "..", "fixtures", "videos")

// Generated media fixtures (pnpm run fixtures:videos), gitignored.
export const MEDIA_DIR = join(__dirname, "..", "..", "..", "fixtures", "videos")

export interface TierExpectation {
	record: VideoProbe | null
	blockReads: number
}

export interface VideoSlice {
	name: string
	size: number
	// [offset, base64 bytes], sorted and non-overlapping.
	ranges: [number, string][]
	tiers: Record<ProbeTier, TierExpectation>
}

// A block source that logs each range the parser reads.
export class RecordingBlockSource extends BlockSource {
	public readonly ranges: [number, number][] = []

	public constructor(size: number, readRange: ReadRange, options?: BlockSourceOptions) {
		super(size, readRange, options)
	}

	public override async read(target: Uint8Array, position: number, length: number): Promise<void> {
		this.ranges.push([position, length])

		await super.read(target, position, length)
	}
}

export interface Measured {
	record: VideoProbe | null
	blockReads: number
	downloaded: number
	consumed: number
	ranges: [number, number][]
}

// Probes one tier through a recording source, counting what was downloaded and what was parsed.
export async function measureProbe(size: number, readRange: ReadRange, tier: ProbeTier): Promise<Measured> {
	let downloaded = 0
	const counting: ReadRange = async (start, end) => {
		downloaded += end - start

		return readRange(start, end)
	}
	const source = new RecordingBlockSource(size, counting, { maxBlockReads: initialBlockReads(tier) })
	const record = await probeVideo(source, tier)

	return {
		record,
		blockReads: source.blockReads,
		downloaded,
		consumed: mergeRanges(source.ranges).reduce((sum, [, length]) => sum + length, 0),
		ranges: source.ranges
	}
}

// Sorted, coalesced [offset, length] ranges.
export function mergeRanges(ranges: readonly (readonly [number, number])[]): [number, number][] {
	const sorted = [...ranges].filter(([, length]) => length > 0).sort((a, b) => a[0] - b[0])
	const merged: [number, number][] = []

	for (const [offset, length] of sorted) {
		const last = merged[merged.length - 1]

		if (last !== undefined && offset <= last[0] + last[1]) {
			last[1] = Math.max(last[1], offset + length - last[0])
		} else {
			merged.push([offset, length])
		}
	}

	return merged
}

// A file read by range from disk.
export async function fileReadRange(path: string): Promise<{ size: number; readRange: ReadRange; close: () => Promise<void> }> {
	const handle = await open(path, "r")
	const { size } = await handle.stat()

	return {
		size,
		readRange: async (start, end) => {
			const bytes = new Uint8Array(end - start)

			await handle.read(bytes, 0, bytes.length, start)

			return bytes
		},
		close: () => handle.close()
	}
}

// The whole sparse file a slice stands for, materialised lazily one block at a time.
export function sliceReadRange(slice: Pick<VideoSlice, "ranges">): ReadRange {
	const ranges = slice.ranges.map(([offset, data]) => [offset, Uint8Array.from(Buffer.from(data, "base64"))] as const)

	return (start, end) => {
		const out = new Uint8Array(end - start)

		for (const [offset, bytes] of ranges) {
			const from = Math.max(start, offset)
			const to = Math.min(end, offset + bytes.length)

			if (from < to) {
				out.set(bytes.subarray(from - offset, to - offset), from - start)
			}
		}

		return Promise.resolve(out)
	}
}

export function loadSlices(): VideoSlice[] {
	return readdirSync(SLICES_DIR)
		.filter(name => name.endsWith(".json"))
		.sort()
		.map(name => JSON.parse(readFileSync(join(SLICES_DIR, name), "utf8")) as VideoSlice)
}

// The block-aligned budget bound a tier's reads must stay within: card is fixed, the others are raised
// by the walk itself and are checked against the recorded count instead.
export const CARD_BUDGET = 4

export function blocksOf(size: number): number {
	return Math.ceil(size / BLOCK_BYTES)
}
