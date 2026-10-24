import { existsSync, readFileSync, readdirSync } from "node:fs"
import { join } from "node:path"
import { unzipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { convertMetafile, type MetafileConversion } from "../core/metafile/index"
import { CORPUS_DIR as CORPUS } from "./corpusDir"
import { ByteWriter, EmfBuilder, WmfBuilder, plusHeader, plusRecord } from "./metafileFixtures"
import { budget } from "./timing"

// Adversarial input: each case must end quickly with bounded memory (playback stops at a limit and keeps
// what it drew) and never throw.

function timed(bytes: Uint8Array): { result: MetafileConversion; ms: number; heap: number } {
	const heap = process.memoryUsage().heapUsed
	const start = performance.now()
	const result = convertMetafile(bytes)

	return { result, ms: performance.now() - start, heap: process.memoryUsage().heapUsed - heap }
}

// EMF+ records grouped into comments of at most 1000 records.
function plusFile(records: readonly Uint8Array[]): Uint8Array {
	const builder = new EmfBuilder().plus([plusHeader(false)])

	for (let k = 0; k < records.length; k += 1000) builder.plus(records.slice(k, k + 1000))

	return builder.build()
}

function floatPath(points: number): Uint8Array {
	return plusRecord(0x4008, 0x0300, writer => {
		writer.u32(0xdbc01002).u32(points).u32(0)

		for (let k = 0; k < points; k++) writer.f32(k % 100).f32((k * 7) % 100)
		for (let k = 0; k < points; k++) writer.u8(k === 0 ? 0 : 1)

		writer.pad(4)
	})
}

describe("metafile limits", () => {
	it("bounds clip work on a large path set as the clip over and over", { timeout: 60_000 }, () => {
		for (const mode of [0, 1, 2, 3, 4, 5]) {
			const records = [floatPath(50_000)]

			for (let k = 0; k < 1000; k++) records.push(plusRecord(0x4033, mode << 8))

			const { result, ms, heap } = timed(plusFile(records))

			expect(result.ok).toBe(true)
			expect(ms).toBeLessThan(budget(5000))
			expect(heap).toBeLessThan(256 * 1024 * 1024)
		}
	})

	it("bounds repeated clip rectangles and their unions", { timeout: 60_000 }, () => {
		const builder = new WmfBuilder().record(0x20c, [1000, 1000])

		for (let k = 0; k < 200_000; k++) builder.record(k % 2 === 0 ? 0x416 : 0x415, [1000 - (k % 7), 1000 - (k % 5), k % 3, k % 11])

		builder.record(0x41b, [10, 10, 0, 0])

		const { result, ms } = timed(builder.build([0, 0, 1000, 1000]))

		expect(result.ok).toBe(true)
		expect(ms).toBeLessThan(budget(5000))
	})

	it("allocates nothing from declared sizes: continued objects and path point counts", () => {
		const continued: Uint8Array[] = []

		for (let k = 0; k < 20_000; k++) continued.push(plusRecord(0x4008, 0x8000 | (1 << 8) | (k & 1), writer => writer.u32(3_000_000)))

		const declaredPaths: Uint8Array[] = []

		for (let k = 0; k < 2000; k++)
			declaredPaths.push(plusRecord(0x4008, 0x0300 | (k & 63), writer => writer.u32(0xdbc01002).u32(1_000_000).u32(0x0800).u32(0)))

		for (const records of [continued, declaredPaths]) {
			const { result, ms, heap } = timed(plusFile(records))

			expect(result.ok).toBe(true)
			expect(ms).toBeLessThan(budget(2000))
			expect(heap).toBeLessThan(128 * 1024 * 1024)
		}
	})

	it("refuses bitmaps a browser would decode into a huge canvas", () => {
		const info = new ByteWriter()
			.u32(40)
			.i32(40_000)
			.i32(40_000)
			.u16(1)
			.u16(8)
			.u32(1)
			.u32(4)
			.u32(0)
			.u32(0)
			.u32(2)
			.u32(0)
			.u32(0xff0000)
			.u32(0xff00)
			.done()
		const bits = Uint8Array.of(0, 1)
		const bytes = new EmfBuilder()
			.record(81, writer =>
				writer
					.i32(0)
					.i32(0)
					.i32(100)
					.i32(100)
					.i32(0)
					.i32(0)
					.i32(0)
					.i32(0)
					.i32(40_000)
					.i32(40_000)
					.u32(80)
					.u32(info.length)
					.u32(80 + info.length)
					.u32(bits.length)
					.u32(0)
					.u32(0x00cc0020)
					.i32(96)
					.i32(96)
					.raw(info)
					.raw(bits)
			)
			.build()
		const result = convertMetafile(bytes)

		expect(result.ok && result.list.images.length).toBe(0)
	})
})

// Seeds from the corpus, mutated: random bytes and runs of 0xFF (huge counts and sizes) in random places.
function seeds(): Uint8Array[] {
	const out: Uint8Array[] = []

	for (const extension of ["docx", "docm", "dotx", "dotm"]) {
		const directory = join(CORPUS, extension)

		if (!existsSync(directory)) continue

		for (const name of readdirSync(directory).sort()) {
			try {
				const parts = unzipSync(new Uint8Array(readFileSync(join(directory, name))), {
					filter: entry => /\.(emf|wmf)$/i.test(entry.name)
				})

				for (const bytes of Object.values(parts)) if (bytes.length < 400_000) out.push(bytes)
			} catch {
				continue
			}
		}
	}

	return out
}

describe.skipIf(!existsSync(CORPUS))("metafile fuzzing over corpus seeds", () => {
	it("converts mutated corpus metafiles without failing, quickly and within memory", { timeout: 600_000 }, () => {
		const corpus = seeds()
		let seed = 99
		const random = (): number => {
			seed = (seed * 1103515245 + 12345) & 0x7fffffff

			return seed / 0x7fffffff
		}
		const failures: string[] = []
		let worst = 0
		let worstHeap = 0

		for (let round = 0; round < 3000 && corpus.length > 0; round++) {
			const source = corpus[Math.floor(random() * corpus.length)] ?? new Uint8Array(0)
			const mutated = source.slice()
			const changes = 1 + Math.floor(random() * 8)

			for (let k = 0; k < changes; k++) {
				const at = Math.floor(random() * mutated.length)

				if (random() < 0.5) {
					mutated[at] = Math.floor(random() * 256)
				} else {
					mutated[at] = 0xff
					mutated[at + 1] = 0xff
					mutated[at + 2] = random() < 0.5 ? 0xff : 0x0f
				}
			}

			const { result, ms, heap } = timed(mutated)

			if (!result.ok && result.reason === "failed") failures.push(`round ${round} (seed of ${source.length} bytes)`)

			worst = Math.max(worst, ms)
			worstHeap = Math.max(worstHeap, heap)
		}

		console.log(
			`metafile fuzz: ${corpus.length} seeds, worst ${worst.toFixed(0)} ms, worst heap growth ${(worstHeap / 1048576).toFixed(0)} MiB`
		)

		expect(failures).toEqual([])
		expect(worst).toBeLessThan(budget(3000))
		expect(worstHeap).toBeLessThan(512 * 1024 * 1024)
	})
})
