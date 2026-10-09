import { describe, expect, it } from "vitest"
import { BLOCK_BYTES, BlockReadError, ReadBudgetError, type ReadRange } from "@/lib/media/blockSource"
import { MAX_PROBE_BLOCK_READS } from "@/lib/media/probe/bytes"
import { CARD_BLOCK_READS, initialBlockReads, probeVideo } from "@/lib/media/probe/probe"
import type { ProbeTier, VideoProbe } from "@/lib/media/probe/types"
import { deflateSync } from "node:zlib"
import {
	FTYP,
	audioEntry,
	box,
	concat,
	ebmlHeader,
	element,
	esds,
	fullBox,
	hdlr,
	mdhd,
	mvhd,
	segment,
	stringElement,
	text,
	tkhd,
	trak,
	u32,
	u64,
	uintElement,
	unknownSize,
	zeros
} from "@/tests/support/containerBytes"
import { RecordingBlockSource, TIERS, loadSlices, type VideoSlice } from "@/tests/support/videoProbe"

// Seeded mutations of the recorded slices: whatever the bytes, a probe resolves (a record or null) or
// fails with one of the source's own errors, never reads past the file or its tier's budget, and never
// asks for more than one block's worth of bytes at once.

// mulberry32: small, fast, and the same sequence on every run.
function random(seed: number): () => number {
	let state = seed >>> 0

	return () => {
		state = (state + 0x6d2b79f5) >>> 0

		let t = state

		t = Math.imul(t ^ (t >>> 15), t | 1)
		t ^= t + Math.imul(t ^ (t >>> 7), t | 61)

		return ((t ^ (t >>> 14)) >>> 0) / 4294967296
	}
}

interface Sparse {
	size: number
	ranges: [number, Uint8Array][]
}

function decode(slice: VideoSlice): Sparse {
	return { size: slice.size, ranges: slice.ranges.map(([offset, data]) => [offset, Uint8Array.from(Buffer.from(data, "base64"))]) }
}

function readRangeOf(sparse: Sparse): ReadRange {
	return (start, end) => {
		if (start < 0 || end > sparse.size || start > end) {
			return Promise.reject(new Error(`read outside the file: ${String(start)}-${String(end)} of ${String(sparse.size)}`))
		}

		const out = new Uint8Array(end - start)

		for (const [offset, bytes] of sparse.ranges) {
			const from = Math.max(start, offset)
			const to = Math.min(end, offset + bytes.length)

			if (from < to) {
				out.set(bytes.subarray(from - offset, to - offset), from - start)
			}
		}

		return Promise.resolve(out)
	}
}

type Mutation = (sparse: Sparse, next: () => number) => Sparse

function pick<T>(items: readonly T[], next: () => number): T {
	const item = items[Math.floor(next() * items.length)]

	if (item === undefined) {
		throw new Error("empty pick")
	}

	return item
}

function mutateBytes(sparse: Sparse, next: () => number, write: (bytes: Uint8Array, at: number) => void): Sparse {
	const ranges = sparse.ranges.map(([offset, bytes]) => [offset, bytes.slice()] as [number, Uint8Array])
	const [, bytes] = pick(ranges, next)

	write(bytes, Math.floor(next() * bytes.length))

	return { size: sparse.size, ranges }
}

const MUTATIONS: readonly Mutation[] = [
	// Bit flips.
	(sparse, next) =>
		mutateBytes(sparse, next, (bytes, at) => {
			for (let i = 0; i < 1 + Math.floor(next() * 4); i++) {
				const index = Math.min(bytes.length - 1, at + Math.floor(next() * 8))

				bytes[index] = (bytes[index] ?? 0) ^ (1 << Math.floor(next() * 8))
			}
		}),
	// A random byte run.
	(sparse, next) =>
		mutateBytes(sparse, next, (bytes, at) => {
			for (let i = at; i < Math.min(bytes.length, at + 1 + Math.floor(next() * 16)); i++) {
				bytes[i] = Math.floor(next() * 256)
			}
		}),
	// A size field blown up: a 32-bit box size, or an EBML vint of 8 bytes all ones (unknown) or huge.
	(sparse, next) =>
		mutateBytes(sparse, next, (bytes, at) => {
			const patterns = [
				[0xff, 0xff, 0xff, 0xff],
				[0x00, 0x00, 0x00, 0x01],
				[0x00, 0x00, 0x00, 0x00],
				[0x01, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff],
				[0x01, 0x1f, 0xff, 0xff, 0xff, 0xff, 0xff, 0xf0],
				[0x10, 0xff, 0xff, 0xff]
			]

			bytes.set(pick(patterns, next).slice(0, bytes.length - at), at)
		}),
	// Truncation.
	(sparse, next) => ({ size: Math.max(1, Math.floor(next() * sparse.size)), ranges: sparse.ranges }),
	// A zeroed run.
	(sparse, next) =>
		mutateBytes(sparse, next, (bytes, at) => {
			bytes.fill(0, at, at + 1 + Math.floor(next() * 64))
		})
]

interface Outcome {
	record: VideoProbe | null
	blockReads: number
	largestRead: number
}

// A probe that hangs fails here rather than at the suite's timeout.
async function probeGuarded(sparse: Sparse, tier: ProbeTier): Promise<Outcome> {
	const source = new RecordingBlockSource(sparse.size, readRangeOf(sparse), { maxBlockReads: initialBlockReads(tier) })
	let timer: ReturnType<typeof setTimeout> | undefined
	const timeout = new Promise<never>((_resolve, reject) => {
		timer = setTimeout(() => {
			reject(new Error("probe did not settle"))
		}, 2000)
	})

	try {
		const record = await Promise.race([probeVideo(source, tier), timeout])

		return { record, blockReads: source.blockReads, largestRead: Math.max(0, ...source.ranges.map(([, length]) => length)) }
	} catch (error) {
		if (error instanceof ReadBudgetError) {
			return { record: null, blockReads: source.blockReads, largestRead: 0 }
		}

		if (error instanceof BlockReadError) {
			// The probe never reads past the file, so the reader's guard must not have fired.
			throw error.cause
		}

		throw error
	} finally {
		clearTimeout(timer)
	}
}

function checkOutcome(outcome: Outcome, tier: ProbeTier): void {
	expect(outcome.blockReads).toBeLessThanOrEqual(tier === "card" ? CARD_BLOCK_READS : MAX_PROBE_BLOCK_READS)
	expect(outcome.largestRead).toBeLessThanOrEqual(BLOCK_BYTES)

	if (outcome.record !== null) {
		expect(outcome.record.chapters.length).toBeLessThanOrEqual(64)
		expect(JSON.parse(JSON.stringify(outcome.record))).toEqual(outcome.record)
	}
}

const ITERATIONS = 40

// Paths the recorded slices do not reach: compressed CodecPrivate, chapters, attachments, chpl and covr.
function syntheticCorpus(): Sparse[] {
	const compressed = element(
		0xae,
		uintElement(0xd7, 1),
		uintElement(0x83, 2),
		stringElement(0x86, "A_AAC"),
		element(0x63a2, deflateSync(Buffer.from([0x11, 0x90]))),
		element(0x6d80, element(0x6240, uintElement(0x5032, 2), element(0x5034, uintElement(0x4254, 0))))
	)
	const chapters = element(
		0x1043a770,
		element(
			0x45b9,
			uintElement(0x45db, 1),
			element(0xb6, uintElement(0x91, 0), uintElement(0x4588, 2), element(0x80, stringElement(0x85, "Intro")))
		)
	)
	const attachments = element(0x1941a469, element(0x61a7, stringElement(0x466e, "cover.jpg"), element(0x465c, [0xff, 0xd8, 0xff])))
	const mkv = concat(
		ebmlHeader(),
		segment(element(0x1549a966, uintElement(0x2ad7b1, 1_000_000)), element(0x1654ae6b, compressed), chapters, attachments)
	)
	const chpl = fullBox("chpl", 1, 0, zeros(4), [1], u64(5e7), [5], text("Intro"))
	const covr = box("covr", box("data", u32(13), u32(0), [0xff, 0xd8]))
	const mp4 = concat(
		FTYP,
		box(
			"moov",
			mvhd({ duration: 1000 }),
			trak(tkhd({ id: 1 }), [mdhd({}), hdlr("soun")], audioEntry("mp4a", {}, esds(0x40, [0x11, 0x90]))),
			box("udta", chpl, fullBox("meta", 0, 0, box("ilst", covr)))
		)
	)

	// Corrupt headers that each look like a Cluster, for resync; a compressed CodecPrivate that inflates
	// far past its cap.
	const resync = concat(
		ebmlHeader(),
		segment(new Uint8Array([0x00]), concat(...Array.from({ length: 2000 }, () => [0x1f, 0x43, 0xb6, 0x75, 0x00])))
	)
	const bombTrack = element(
		0xae,
		uintElement(0xd7, 1),
		uintElement(0x83, 1),
		stringElement(0x86, "V_MPEG4/ISO/AVC"),
		element(0x63a2, deflateSync(Buffer.alloc(16 * 1024 * 1024))),
		element(0x6d80, element(0x6240, uintElement(0x5032, 2), element(0x5034, uintElement(0x4254, 0))))
	)
	const bomb = concat(ebmlHeader(), segment(element(0x1654ae6b, bombTrack)))

	return [mkv, mp4, resync, bomb].map(bytes => ({ size: bytes.length, ranges: [[0, bytes]] }))
}

describe("video probe fuzz", () => {
	for (const slice of loadSlices()) {
		it(`survives mutations of ${slice.name}`, async () => {
			const original = decode(slice)
			const next = random(slice.size)

			for (let i = 0; i < ITERATIONS; i++) {
				const mutated = pick(MUTATIONS, next)(original, next)

				for (const tier of TIERS) {
					checkOutcome(await probeGuarded(mutated, tier), tier)
				}
			}
		})
	}

	it("survives mutations of the synthetic corpus", async () => {
		const next = random(7)

		for (const original of syntheticCorpus()) {
			for (let i = 0; i < ITERATIONS * 10; i++) {
				const mutated = pick(MUTATIONS, next)(original, next)

				for (const tier of TIERS) {
					checkOutcome(await probeGuarded(mutated, tier), tier)
				}
			}
		}
	})

	it("survives deep nesting and unknown sizes where none are allowed", async () => {
		let nested = element(0x1654ae6b)

		for (let i = 0; i < 200; i++) {
			nested = element(0xae, nested)
		}

		const files = [
			concat(ebmlHeader(), segment(element(0x1654ae6b, nested))),
			concat(ebmlHeader(), segment(unknownSize(0x1549a966, unknownSize(0x1654ae6b, unknownSize(0xae))))),
			concat(ebmlHeader(), unknownSize(0x18538067, unknownSize(0x18538067, unknownSize(0x1f43b675)))),
			concat(ebmlHeader(), segment(element(0x1043a770, ...Array.from({ length: 5000 }, () => element(0x45b9, element(0xb6))))))
		]

		for (const bytes of files) {
			for (const tier of TIERS) {
				const sparse: Sparse = { size: bytes.length, ranges: [[0, bytes]] }

				checkOutcome(await probeGuarded(sparse, tier), tier)
			}
		}
	})
})
