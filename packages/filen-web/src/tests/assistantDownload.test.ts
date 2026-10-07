import { createHash, type Hash } from "node:crypto"
import { afterEach, describe, expect, it, vi } from "vitest"
import {
	downloadModelShards,
	downloadShard,
	ModelDownloadError,
	type ShardFile,
	type ShardStore,
	type StreamingHasher
} from "@/features/assistant/runtime/download"
import type { AssistantModelFile, AssistantModelManifest } from "@/features/assistant/runtime/manifest"

class MemoryFile implements ShardFile {
	bytes = new Uint8Array(0)
	flushes = 0
	quotaAfter: number | null = null

	getSize(): number {
		return this.bytes.length
	}

	truncate(size: number): void {
		this.bytes = this.bytes.slice(0, size)
	}

	read(buffer: Uint8Array, { at }: { at: number }): number {
		const view = this.bytes.subarray(at, at + buffer.length)

		buffer.set(view)

		return view.length
	}

	write(buffer: Uint8Array, { at }: { at: number }): number {
		if (this.quotaAfter !== null && at + buffer.length > this.quotaAfter) {
			throw new DOMException("full", "QuotaExceededError")
		}

		const next = new Uint8Array(Math.max(this.bytes.length, at + buffer.length))

		next.set(this.bytes)
		next.set(buffer, at)
		this.bytes = next

		return buffer.length
	}

	flush(): void {
		this.flushes++
	}
}

function hasher(): StreamingHasher {
	let hash: Hash = createHash("sha256")

	return {
		init: () => {
			hash = createHash("sha256")
		},
		update: data => hash.update(data),
		digest: () => hash.digest("hex")
	}
}

function content(size: number, seed = 1): Uint8Array {
	return Uint8Array.from({ length: size }, (_, index) => (index * 31 + seed) & 0xff)
}

function describeShard(file: string, bytes: Uint8Array): AssistantModelFile {
	return { file, size: bytes.length, sha256: createHash("sha256").update(bytes).digest("hex") }
}

// Serves `bytes` in 1000-byte chunks, honoring "bytes=N-" unless `ignoreRange`.
function host(files: Record<string, Uint8Array>, { ignoreRange = false } = {}) {
	const requests: { url: string; range: string | null }[] = []
	const fetchShard = vi.fn((url: string, init: RequestInit) => {
		const name = decodeURIComponent(url.slice(url.lastIndexOf("/") + 1))
		const bytes = files[name]
		const range = init.headers === undefined ? null : (new Headers(init.headers).get("Range") ?? null)

		requests.push({ url, range })

		if (bytes === undefined) {
			return Promise.resolve(new Response(null, { status: 404 }))
		}

		const start = !ignoreRange && range !== null ? Number(/bytes=(\d+)-/.exec(range)?.[1] ?? 0) : 0
		const body = bytes.slice(start)
		const stream = new ReadableStream<Uint8Array>({
			start(controller) {
				for (let offset = 0; offset < body.length; offset += 1000) {
					controller.enqueue(body.slice(offset, offset + 1000))
				}

				controller.close()
			}
		})

		return Promise.resolve(new Response(stream, { status: start > 0 ? 206 : 200, headers: { "Content-Length": String(body.length) } }))
	})

	return { fetchShard, requests }
}

describe("downloadShard", () => {
	it("downloads and verifies a shard, sending no header but Range", async () => {
		const bytes = content(4500)
		const shard = describeShard("a.gguf", bytes)
		const file = new MemoryFile()
		const { fetchShard, requests } = host({ "a.gguf": bytes })
		let received = 0

		await downloadShard({
			shard,
			url: "https://cdn.example/a.gguf",
			file,
			hasher: hasher(),
			fetchShard,
			signal: new AbortController().signal,
			onBytes: bytes => (received += bytes)
		})

		expect(file.bytes).toEqual(bytes)
		expect(received).toBe(4500)
		expect(requests).toEqual([{ url: "https://cdn.example/a.gguf", range: null }])
		expect(fetchShard.mock.calls[0]?.[1]).toMatchObject({ credentials: "omit", cache: "no-store" })
	})

	it("resumes from the bytes on disk with a single range", async () => {
		const bytes = content(4500)
		const shard = describeShard("a.gguf", bytes)
		const file = new MemoryFile()
		const { fetchShard, requests } = host({ "a.gguf": bytes })

		file.write(bytes.slice(0, 2000), { at: 0 })

		let received = 0

		await downloadShard({
			shard,
			url: "https://cdn.example/a.gguf",
			file,
			hasher: hasher(),
			fetchShard,
			signal: new AbortController().signal,
			onBytes: bytes => (received += bytes)
		})

		expect(requests[0]?.range).toBe("bytes=2000-")
		expect(file.bytes).toEqual(bytes)
		expect(received).toBe(4500)
	})

	it("starts over when the host ignores the range", async () => {
		const bytes = content(3000)
		const shard = describeShard("a.gguf", bytes)
		const file = new MemoryFile()
		const { fetchShard } = host({ "a.gguf": bytes }, { ignoreRange: true })

		file.write(bytes.slice(0, 1000), { at: 0 })

		let received = 0

		await downloadShard({
			shard,
			url: "https://cdn.example/a.gguf",
			file,
			hasher: hasher(),
			fetchShard,
			signal: new AbortController().signal,
			onBytes: bytes => (received += bytes)
		})

		expect(file.bytes).toEqual(bytes)
		expect(received).toBe(3000)
	})

	it("empties a shard whose hash does not match", async () => {
		const bytes = content(2000)
		const shard = { ...describeShard("a.gguf", bytes), sha256: "0".repeat(64) }
		const file = new MemoryFile()
		const { fetchShard } = host({ "a.gguf": bytes })

		await expect(
			downloadShard({
				shard,
				url: "https://cdn.example/a.gguf",
				file,
				hasher: hasher(),
				fetchShard,
				signal: new AbortController().signal,
				onBytes: () => undefined
			})
		).rejects.toMatchObject({ kind: "integrity" })
		expect(file.getSize()).toBe(0)
	})

	it("fails on an HTTP error", async () => {
		await expect(
			downloadShard({
				shard: describeShard("missing.gguf", content(10)),
				url: "https://cdn.example/missing.gguf",
				file: new MemoryFile(),
				hasher: hasher(),
				fetchShard: host({}).fetchShard,
				signal: new AbortController().signal,
				onBytes: () => undefined
			})
		).rejects.toBeInstanceOf(ModelDownloadError)
	})
})

// Answers like a real fetch whose stream errors when its signal aborts: `deliver` decides, per request,
// how many bytes arrive before the response goes silent (null: the whole rest, then the stream closes).
function silentHost(bytes: Uint8Array, deliver: (request: number) => number | null) {
	const ranges: (string | null)[] = []
	const fetchShard = vi.fn((_url: string, init: RequestInit) => {
		const range = init.headers === undefined ? null : (new Headers(init.headers).get("Range") ?? null)
		const start = range === null ? 0 : Number(/bytes=(\d+)-/.exec(range)?.[1] ?? 0)
		const body = bytes.slice(start)
		const count = deliver(ranges.length)

		ranges.push(range)

		const stream = new ReadableStream<Uint8Array>({
			start(controller) {
				if (count === null) {
					controller.enqueue(body)
					controller.close()

					return
				}

				if (count > 0) {
					controller.enqueue(body.slice(0, count))
				}

				init.signal?.addEventListener("abort", () => {
					controller.error(new DOMException("aborted", "AbortError"))
				})
			}
		})

		return Promise.resolve(new Response(stream, { status: start > 0 ? 206 : 200, headers: { "Content-Length": String(body.length) } }))
	})

	return { fetchShard, ranges }
}

describe("downloadShard when a response goes silent", () => {
	afterEach(() => {
		vi.useRealTimers()
	})

	it("abandons the silent response and resumes from the bytes it delivered", async () => {
		vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] })

		const bytes = content(4500)
		const file = new MemoryFile()
		const { fetchShard, ranges } = silentHost(bytes, request => (request === 0 ? 2000 : null))
		let received = 0
		const run = downloadShard({
			shard: describeShard("a.gguf", bytes),
			url: "https://cdn.example/a.gguf",
			file,
			hasher: hasher(),
			fetchShard,
			signal: new AbortController().signal,
			onBytes: bytes => (received += bytes)
		})

		await vi.advanceTimersByTimeAsync(30_000)
		await run

		expect(ranges).toEqual([null, "bytes=2000-"])
		expect(file.bytes).toEqual(bytes)
		expect(received).toBe(4500)
	})

	it("fails after five silent responses in a row that brought nothing", async () => {
		vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] })

		const bytes = content(4500)
		const { fetchShard, ranges } = silentHost(bytes, () => 0)
		const run = downloadShard({
			shard: describeShard("a.gguf", bytes),
			url: "https://cdn.example/a.gguf",
			file: new MemoryFile(),
			hasher: hasher(),
			fetchShard,
			signal: new AbortController().signal,
			onBytes: () => undefined
		}).then(
			() => null,
			(e: unknown) => e
		)

		await vi.advanceTimersByTimeAsync(5 * 30_000)

		const error = await run

		expect(error).toBeInstanceOf(ModelDownloadError)
		expect((error as ModelDownloadError).kind).toBe("network")
		expect(ranges).toHaveLength(5)
	})
})

describe("downloadModelShards", () => {
	function memoryStore(verified = new Set<string>()): ShardStore & { files: Map<string, MemoryFile> } {
		const files = new Map<string, MemoryFile>()

		return {
			files,
			isVerified: shard => verified.has(shard.file),
			markVerified: shard => {
				verified.add(shard.file)

				return Promise.resolve()
			},
			withFile: (shard, run) => {
				let file = files.get(shard.file)

				if (file === undefined) {
					file = new MemoryFile()
					files.set(shard.file, file)
				}

				return run(file)
			}
		}
	}

	it("skips verified shards and reports bytes on disk", async () => {
		const first = content(1200, 1)
		const second = content(800, 2)
		const third = content(500, 3)
		const manifest: AssistantModelManifest = {
			id: "m",
			label: "M",
			license: { name: "MIT", url: "" },
			source: { repo: "r", revision: "x" },
			shards: [
				describeShard("m-00001-of-00003.gguf", first),
				describeShard("m-00002-of-00003.gguf", second),
				describeShard("m-00003-of-00003.gguf", third)
			]
		}
		const store = memoryStore(new Set(["m-00001-of-00003.gguf"]))
		const { fetchShard, requests } = host({ "m-00002-of-00003.gguf": second, "m-00003-of-00003.gguf": third })
		const progress: number[] = []

		await downloadModelShards({
			manifest,
			baseUrl: "https://cdn.example/gguf/",
			store,
			createHasher: () => Promise.resolve(hasher()),
			fetchShard,
			signal: new AbortController().signal,
			onProgress: bytes => progress.push(bytes)
		})

		expect(requests.map(request => request.url).sort()).toEqual([
			"https://cdn.example/gguf/m-00002-of-00003.gguf",
			"https://cdn.example/gguf/m-00003-of-00003.gguf"
		])
		expect(progress[0]).toBe(1200)
		expect(progress.at(-1)).toBe(2500)
		expect(manifest.shards.every(shard => store.isVerified(shard))).toBe(true)
	})

	it("reports a full disk as storage-full", async () => {
		const bytes = content(3000)
		const store = memoryStore()
		const file = new MemoryFile()

		file.quotaAfter = 1500
		store.files.set("a.gguf", file)

		await expect(
			downloadModelShards({
				manifest: {
					id: "m",
					label: "M",
					license: { name: "MIT", url: "" },
					source: { repo: "r", revision: "x" },
					shards: [describeShard("a.gguf", bytes)]
				},
				baseUrl: "https://cdn.example/",
				store,
				createHasher: () => Promise.resolve(hasher()),
				fetchShard: host({ "a.gguf": bytes }).fetchShard,
				signal: new AbortController().signal,
				onProgress: () => undefined
			})
		).rejects.toMatchObject({ kind: "storage-full" })
	})

	it("refetches a corrupted shard once, then fails", async () => {
		const bytes = content(900)
		const manifest: AssistantModelManifest = {
			id: "m",
			label: "M",
			license: { name: "MIT", url: "" },
			source: { repo: "r", revision: "x" },
			shards: [{ ...describeShard("a.gguf", bytes), sha256: "1".repeat(64) }]
		}
		const { fetchShard } = host({ "a.gguf": bytes })

		await expect(
			downloadModelShards({
				manifest,
				baseUrl: "https://cdn.example/",
				store: memoryStore(),
				createHasher: () => Promise.resolve(hasher()),
				fetchShard,
				signal: new AbortController().signal,
				onProgress: () => undefined
			})
		).rejects.toMatchObject({ kind: "integrity" })
		expect(fetchShard).toHaveBeenCalledTimes(2)
	})
})
