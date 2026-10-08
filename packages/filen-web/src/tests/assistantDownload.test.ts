import { createHash } from "node:crypto"
import { createSHA256 } from "hash-wasm"
import { afterEach, describe, expect, it, vi } from "vitest"
import {
	downloadModelShards,
	downloadShard,
	globalConnectivity,
	ModelDownloadError,
	retryDelayMs,
	STALL_MS,
	type Connectivity,
	type ModelDownloadProgress,
	type ShardFile,
	type ShardStore,
	type StreamingHasher
} from "@/features/assistant/runtime/download"
import type { AssistantModelFile, AssistantModelManifest } from "@/features/assistant/runtime/manifest"

class MemoryFile implements ShardFile {
	bytes = new Uint8Array(0)
	flushes = 0
	readBytes = 0
	quotaAfter: number | null = null

	getSize(): number {
		return this.bytes.length
	}

	truncate(size: number): void {
		const next = new Uint8Array(size)

		next.set(this.bytes.subarray(0, size))
		this.bytes = next
	}

	read(buffer: Uint8Array, { at }: { at: number }): number {
		const view = this.bytes.subarray(at, at + buffer.length)

		buffer.set(view)
		this.readBytes += view.length

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

// The real hasher: its save/load is what a resume depends on.
function hasher(): Promise<StreamingHasher> {
	return createSHA256()
}

function content(size: number, seed = 1): Uint8Array<ArrayBuffer> {
	return Uint8Array.from({ length: size }, (_, index) => (index * 31 + seed) & 0xff)
}

function describeShard(file: string, bytes: Uint8Array): AssistantModelFile {
	return { file, size: bytes.length, sha256: createHash("sha256").update(bytes).digest("hex") }
}

function manifestOf(shards: AssistantModelFile[]): AssistantModelManifest {
	return { id: "m", label: "M", license: { name: "MIT", url: "" }, source: { repo: "r", revision: "x" }, shards }
}

const ONLINE: Connectivity = {
	isOnline: () => true,
	waitOnline: () => Promise.resolve()
}

function requestRange(init: RequestInit): string | null {
	return init.headers === undefined ? null : new Headers(init.headers).get("Range")
}

// How one request is answered: a status, then `deliver` bytes of the requested range (null: all of it),
// then the stream closes, errors, or goes silent until its signal aborts. "hang" after the whole range is
// a response whose end-of-stream never arrives.
interface Answer {
	status?: number
	deliver?: number | null
	then?: "close" | "error" | "hang"
	ignoreRange?: boolean
}

function scriptedHost(files: Record<string, Uint8Array>, script: (request: number, file: string) => Answer = () => ({})) {
	const requests: { file: string; range: string | null }[] = []
	let inFlight = 0
	let maxInFlight = 0
	const fetchShard = vi.fn((url: string, init: RequestInit) => {
		const name = decodeURIComponent(url.slice(url.lastIndexOf("/") + 1))
		const bytes = files[name]
		const range = requestRange(init)
		const answer = script(requests.length, name)

		requests.push({ file: name, range })

		if (bytes === undefined) {
			return Promise.resolve(new Response(null, { status: 404 }))
		}

		if (answer.status !== undefined) {
			return Promise.resolve(new Response("error", { status: answer.status }))
		}

		const start = answer.ignoreRange !== true && range !== null ? Number(/bytes=(\d+)-/.exec(range)?.[1] ?? 0) : 0
		const body = bytes.slice(start)
		const count = answer.deliver ?? body.length
		const then = answer.then ?? "close"

		inFlight++
		maxInFlight = Math.max(maxInFlight, inFlight)

		let offset = 0
		let finished = false
		const finish = () => {
			if (!finished) {
				finished = true
				inFlight--
			}
		}
		// Pulled chunk by chunk like a network body: what is delivered is read before the stream ends.
		const stream = new ReadableStream<Uint8Array>(
			{
				pull(controller) {
					if (offset < count) {
						const end = Math.min(count, offset + 1000)

						controller.enqueue(body.slice(offset, end))
						offset = end

						return
					}

					if (then === "close") {
						finish()
						controller.close()
					} else if (then === "error") {
						finish()
						controller.error(new TypeError("network error"))
					} else {
						init.signal?.addEventListener("abort", () => {
							finish()
							controller.error(new DOMException("aborted", "AbortError"))
						})

						return new Promise<void>(() => undefined)
					}
				},
				cancel() {
					finish()
				}
			},
			{ highWaterMark: 0 }
		)

		return Promise.resolve(new Response(stream, { status: start > 0 ? 206 : 200, headers: { "Content-Length": String(body.length) } }))
	})

	return {
		fetchShard,
		requests,
		maxInFlight: () => maxInFlight
	}
}

async function runShard({
	shard,
	file = new MemoryFile(),
	checkpoint = new MemoryFile(),
	fetchShard,
	signal = new AbortController().signal,
	connectivity = ONLINE,
	onWaiting = () => undefined
}: {
	shard: AssistantModelFile
	file?: MemoryFile
	checkpoint?: MemoryFile
	fetchShard: (url: string, init: RequestInit) => Promise<Response>
	signal?: AbortSignal
	connectivity?: Connectivity
	onWaiting?: (waiting: boolean) => void
}): Promise<{ sizes: number[]; file: MemoryFile; checkpoint: MemoryFile }> {
	const sizes: number[] = []

	await downloadShard({
		shard,
		url: `https://cdn.example/${shard.file}`,
		file,
		checkpoint,
		hasher: await hasher(),
		fetchShard,
		signal,
		connectivity,
		onSize: size => sizes.push(size),
		onWaiting,
		random: () => 0.5
	})

	return { sizes, file, checkpoint }
}

// A checkpoint of `bytes` up to `offset`, as the download writes it.
async function checkpointAt(bytes: Uint8Array, offset: number): Promise<MemoryFile> {
	const h = await hasher()

	h.init()
	h.update(bytes.subarray(0, offset))

	const state = h.save()
	const record = new Uint8Array(16 + state.length)
	const view = new DataView(record.buffer)

	view.setFloat64(0, offset, true)
	record.set(state, 8)
	view.setFloat64(record.length - 8, offset, true)

	const file = new MemoryFile()

	file.write(record, { at: 0 })

	return file
}

afterEach(() => {
	vi.useRealTimers()
})

// Runs fake time forward until `run` settles: the code under test awaits real wasm work in between.
async function drain<T>(run: Promise<T>, limitMs = 15 * 60_000): Promise<T> {
	const state = { settled: false }

	void run.then(
		() => (state.settled = true),
		() => (state.settled = true)
	)

	for (let elapsed = 0; !state.settled && elapsed < limitMs; elapsed += 250) {
		await vi.advanceTimersByTimeAsync(250)
	}

	return run
}

describe("downloadShard", () => {
	it("downloads and verifies a shard, sending no header but Range", async () => {
		const bytes = content(4500)
		const shard = describeShard("a.gguf", bytes)
		const { fetchShard, requests } = scriptedHost({ "a.gguf": bytes })
		const { sizes, file, checkpoint } = await runShard({ shard, fetchShard })

		expect(file.bytes).toEqual(bytes)
		expect(sizes[0]).toBe(0)
		expect(sizes.at(-1)).toBe(4500)
		expect(Math.max(...sizes)).toBe(4500)
		expect(requests).toEqual([{ file: "a.gguf", range: null }])
		expect(fetchShard.mock.calls[0]?.[1]).toMatchObject({ credentials: "omit", cache: "no-store" })
		// A verified shard keeps no checkpoint.
		expect(checkpoint.getSize()).toBe(0)
	})

	it("resumes from the checkpoint, re-reading only the bytes written after it", async () => {
		const bytes = content(4500)
		const shard = describeShard("a.gguf", bytes)
		const file = new MemoryFile()
		const { fetchShard, requests } = scriptedHost({ "a.gguf": bytes })

		file.write(bytes.slice(0, 2000), { at: 0 })

		const { sizes } = await runShard({ shard, file, checkpoint: await checkpointAt(bytes, 1500), fetchShard })

		expect(file.readBytes).toBe(500)
		expect(requests[0]?.range).toBe("bytes=2000-")
		expect(file.bytes).toEqual(bytes)
		expect(sizes[0]).toBe(2000)
		expect(sizes.at(-1)).toBe(4500)
	})

	it("hashes the whole prefix when the checkpoint is missing or torn", async () => {
		const bytes = content(4500)
		const shard = describeShard("a.gguf", bytes)

		for (const checkpoint of [new MemoryFile(), await checkpointAt(bytes, 1500)]) {
			const file = new MemoryFile()

			file.write(bytes.slice(0, 2000), { at: 0 })

			if (checkpoint.getSize() > 0) {
				// A crash mid-write: the trailing offset still holds the previous record's.
				new DataView(checkpoint.bytes.buffer).setFloat64(checkpoint.getSize() - 8, 1000, true)
			}

			await runShard({ shard, file, checkpoint, fetchShard: scriptedHost({ "a.gguf": bytes }).fetchShard })

			expect(file.readBytes).toBe(2000)
			expect(file.bytes).toEqual(bytes)
		}
	})

	it("ignores a checkpoint ahead of the bytes on disk", async () => {
		const bytes = content(4500)
		const shard = describeShard("a.gguf", bytes)
		const file = new MemoryFile()

		file.write(bytes.slice(0, 1000), { at: 0 })

		await runShard({
			shard,
			file,
			checkpoint: await checkpointAt(bytes, 3000),
			fetchShard: scriptedHost({ "a.gguf": bytes }).fetchShard
		})

		expect(file.readBytes).toBe(1000)
		expect(file.bytes).toEqual(bytes)
	})

	it("cancelling leaves a checkpoint at the bytes on disk, so the resume reads nothing back", async () => {
		const bytes = content(4500)
		const shard = describeShard("a.gguf", bytes)
		const file = new MemoryFile()
		const checkpoint = new MemoryFile()
		const controller = new AbortController()
		const first = scriptedHost({ "a.gguf": bytes }, () => ({ deliver: 3000, then: "hang" }))
		const run = runShard({
			shard,
			file,
			checkpoint,
			fetchShard: first.fetchShard,
			signal: controller.signal
		}).catch((e: unknown) => e)

		await vi.waitFor(() => {
			expect(file.getSize()).toBe(3000)
		})
		controller.abort()

		expect(await run).toMatchObject({ name: "AbortError" })
		expect(checkpoint.getSize()).toBeGreaterThan(0)

		await runShard({ shard, file, checkpoint, fetchShard: scriptedHost({ "a.gguf": bytes }).fetchShard })

		expect(file.readBytes).toBe(0)
		expect(file.bytes).toEqual(bytes)
	})

	it("starts over when the host ignores the range", async () => {
		const bytes = content(3000)
		const shard = describeShard("a.gguf", bytes)
		const file = new MemoryFile()

		file.write(bytes.slice(0, 1000), { at: 0 })

		const { sizes } = await runShard({
			shard,
			file,
			fetchShard: scriptedHost({ "a.gguf": bytes }, () => ({ ignoreRange: true })).fetchShard
		})

		expect(file.bytes).toEqual(bytes)
		expect(sizes).toContain(0)
		expect(sizes.at(-1)).toBe(3000)
		expect(Math.max(...sizes)).toBe(3000)
	})

	it("empties a shard and its checkpoint when the hash does not match", async () => {
		const bytes = content(2000)
		const shard = { ...describeShard("a.gguf", bytes), sha256: "0".repeat(64) }
		const file = new MemoryFile()
		const checkpoint = new MemoryFile()
		const sizes: number[] = []

		await expect(
			downloadShard({
				shard,
				url: "https://cdn.example/a.gguf",
				file,
				checkpoint,
				hasher: await hasher(),
				fetchShard: scriptedHost({ "a.gguf": bytes }).fetchShard,
				signal: new AbortController().signal,
				connectivity: ONLINE,
				onSize: size => sizes.push(size),
				onWaiting: () => undefined
			})
		).rejects.toMatchObject({ kind: "integrity" })
		expect(file.getSize()).toBe(0)
		expect(checkpoint.getSize()).toBe(0)
		expect(sizes.at(-1)).toBe(0)
	})

	it("fails at once on a client error", async () => {
		const { fetchShard } = scriptedHost({})

		await expect(runShard({ shard: describeShard("missing.gguf", content(10)), fetchShard })).rejects.toBeInstanceOf(ModelDownloadError)
		expect(fetchShard).toHaveBeenCalledTimes(1)
	})

	it("retries a server error after a backoff and continues with a range", async () => {
		vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] })

		const bytes = content(4500)
		const { fetchShard, requests } = scriptedHost({ "a.gguf": bytes }, request =>
			request === 0 ? { deliver: 2000, then: "error" } : request === 1 ? { status: 503 } : {}
		)
		const waiting: boolean[] = []
		const { file } = await drain(runShard({ shard: describeShard("a.gguf", bytes), fetchShard, onWaiting: next => waiting.push(next) }))

		expect(requests.map(request => request.range)).toEqual([null, "bytes=2000-", "bytes=2000-"])
		expect(file.bytes).toEqual(bytes)
		expect(waiting).toEqual([true, false, false])
	})

	it("resumes at once after a dropped connection that brought bytes", async () => {
		const bytes = content(4500)
		const { fetchShard, requests } = scriptedHost({ "a.gguf": bytes }, request =>
			request === 0 ? { deliver: 3000, then: "error" } : {}
		)
		const { file } = await runShard({ shard: describeShard("a.gguf", bytes), fetchShard })

		expect(requests.map(request => request.range)).toEqual([null, "bytes=3000-"])
		expect(file.bytes).toEqual(bytes)
	})

	it("finishes at the published size without waiting for the end of the stream", async () => {
		vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] })

		const bytes = content(4500)
		const { file } = await runShard({
			shard: describeShard("a.gguf", bytes),
			fetchShard: scriptedHost({ "a.gguf": bytes }, () => ({ then: "hang" })).fetchShard
		})

		expect(file.bytes).toEqual(bytes)
		expect(vi.getTimerCount()).toBe(0)
	})

	it("abandons a silent response and resumes from the bytes it delivered", async () => {
		vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout", "Date"] })

		const bytes = content(4500)
		const { fetchShard, requests } = scriptedHost({ "a.gguf": bytes }, request =>
			request === 0 ? { deliver: 2000, then: "hang" } : {}
		)
		const started = Date.now()
		const { file, sizes } = await drain(runShard({ shard: describeShard("a.gguf", bytes), fetchShard }))

		expect(Date.now() - started).toBeGreaterThanOrEqual(STALL_MS)
		expect(Date.now() - started).toBeLessThan(STALL_MS + 1000)

		expect(requests.map(request => request.range)).toEqual([null, "bytes=2000-"])
		expect(file.bytes).toEqual(bytes)
		expect(sizes.at(-1)).toBe(4500)
	})

	it("fails after eight attempts in a row that brought nothing", async () => {
		vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] })

		const bytes = content(4500)
		const { fetchShard, requests } = scriptedHost({ "a.gguf": bytes }, () => ({ deliver: 0, then: "hang" }))
		const error = await drain(
			runShard({ shard: describeShard("a.gguf", bytes), fetchShard }).then(
				() => null,
				(e: unknown) => e
			)
		)

		expect(error).toBeInstanceOf(ModelDownloadError)
		expect((error as ModelDownloadError).kind).toBe("network")
		expect(requests).toHaveLength(8)
	})

	it("waits for the connection while offline, without spending attempts", async () => {
		const bytes = content(4500)
		let online = false
		let comeOnline: () => void = () => undefined
		const connectivity: Connectivity = {
			isOnline: () => online,
			waitOnline: () =>
				new Promise(resolve => {
					comeOnline = () => {
						online = true
						resolve()
					}
				})
		}
		const { fetchShard } = scriptedHost({ "a.gguf": bytes })
		const waiting: boolean[] = []
		const run = runShard({ shard: describeShard("a.gguf", bytes), fetchShard, connectivity, onWaiting: next => waiting.push(next) })

		await vi.waitFor(() => {
			expect(waiting).toEqual([true])
		})
		expect(fetchShard).not.toHaveBeenCalled()

		comeOnline()

		const { file } = await run

		expect(file.bytes).toEqual(bytes)
		expect(waiting).toEqual([true, false, false])
	})

	it("going offline mid-transfer does not count against the shard", async () => {
		const bytes = content(4500)
		let online = true
		const connectivity: Connectivity = {
			isOnline: () => online,
			waitOnline: () => {
				online = true

				return Promise.resolve()
			}
		}
		// Ten drops with nothing delivered, each while offline: more than the attempt limit.
		const { fetchShard } = scriptedHost({ "a.gguf": bytes }, request => {
			if (request < 10) {
				online = false

				return { deliver: 0, then: "error" }
			}

			return {}
		})
		const { file } = await runShard({ shard: describeShard("a.gguf", bytes), fetchShard, connectivity })

		expect(file.bytes).toEqual(bytes)
		expect(fetchShard).toHaveBeenCalledTimes(11)
	})
})

describe("retryDelayMs", () => {
	it("doubles from a second up to thirty, with at least half of each step", () => {
		expect(retryDelayMs(1, () => 0)).toBe(500)
		expect(retryDelayMs(1, () => 1)).toBe(1000)
		expect(retryDelayMs(3, () => 0)).toBe(2000)
		expect(retryDelayMs(20, () => 1)).toBe(30_000)
		expect(retryDelayMs(20, () => 0)).toBe(15_000)
	})
})

describe("globalConnectivity", () => {
	it("reads onLine and settles on the online event", async () => {
		const scope = Object.assign(new EventTarget(), { navigator: { onLine: false } })
		const connectivity = globalConnectivity(scope)
		const waited = connectivity.waitOnline(new AbortController().signal)

		expect(connectivity.isOnline()).toBe(false)
		scope.navigator.onLine = true
		scope.dispatchEvent(new Event("online"))

		await expect(waited).resolves.toBeUndefined()
	})

	it("rejects when the download is cancelled", async () => {
		const scope = Object.assign(new EventTarget(), { navigator: { onLine: false } })
		const controller = new AbortController()
		const waited = globalConnectivity(scope).waitOnline(controller.signal)

		controller.abort()

		await expect(waited).rejects.toMatchObject({ name: "AbortError" })
	})
})

describe("downloadModelShards", () => {
	function memoryStore(verified = new Set<string>()): ShardStore & { files: Map<string, MemoryFile> } {
		const files = new Map<string, MemoryFile>()
		const checkpoints = new Map<string, MemoryFile>()
		const open = (map: Map<string, MemoryFile>, name: string) => {
			let file = map.get(name)

			if (file === undefined) {
				file = new MemoryFile()
				map.set(name, file)
			}

			return file
		}

		return {
			files,
			isVerified: shard => verified.has(shard.file),
			sizeOnDisk: shard => files.get(shard.file)?.getSize() ?? 0,
			markVerified: shard => {
				verified.add(shard.file)

				return Promise.resolve()
			},
			withFile: (shard, run) => run(open(files, shard.file), open(checkpoints, shard.file))
		}
	}

	function run({
		manifest,
		store,
		fetchShard,
		connectivity = ONLINE,
		onProgress = () => undefined
	}: {
		manifest: AssistantModelManifest
		store: ShardStore
		fetchShard: (url: string, init: RequestInit) => Promise<Response>
		connectivity?: Connectivity
		onProgress?: (progress: ModelDownloadProgress) => void
	}): Promise<void> {
		return downloadModelShards({
			manifest,
			baseUrl: "https://cdn.example/gguf/",
			store,
			createHasher: hasher,
			fetchShard,
			connectivity,
			signal: new AbortController().signal,
			onProgress,
			random: () => 0.5
		})
	}

	it("skips verified shards and reports exact bytes on disk", async () => {
		const first = content(1200, 1)
		const second = content(800, 2)
		const third = content(500, 3)
		const manifest = manifestOf([
			describeShard("m-00001-of-00003.gguf", first),
			describeShard("m-00002-of-00003.gguf", second),
			describeShard("m-00003-of-00003.gguf", third)
		])
		const store = memoryStore(new Set(["m-00001-of-00003.gguf"]))
		const { fetchShard, requests } = scriptedHost({ "m-00002-of-00003.gguf": second, "m-00003-of-00003.gguf": third })
		const progress: number[] = []

		await run({ manifest, store, fetchShard, onProgress: ({ receivedBytes }) => progress.push(receivedBytes) })

		expect(requests.map(request => request.file).sort()).toEqual(["m-00002-of-00003.gguf", "m-00003-of-00003.gguf"])
		expect(progress[0]).toBe(1200)
		expect(progress.at(-1)).toBe(2500)
		expect(progress.every((bytes, index) => index === 0 || bytes >= (progress[index - 1] ?? 0))).toBe(true)
		expect(manifest.shards.every(shard => store.isVerified(shard))).toBe(true)
	})

	it("starts the longest shards first, three at a time", async () => {
		const sizes = [100, 5000, 4000, 300, 4500]
		const files = Object.fromEntries(sizes.map((size, index) => [`m-0000${String(index + 1)}-of-00005.gguf`, content(size, index)]))
		const manifest = manifestOf(Object.entries(files).map(([name, bytes]) => describeShard(name, bytes)))
		const { fetchShard, requests, maxInFlight } = scriptedHost(files)

		await run({ manifest, store: memoryStore(), fetchShard })

		expect(requests.slice(0, 3).map(request => request.file)).toEqual([
			"m-00002-of-00005.gguf",
			"m-00005-of-00005.gguf",
			"m-00003-of-00005.gguf"
		])
		expect(maxInFlight()).toBe(3)
	})

	it("never counts a byte twice across resumes, drops and stalls in parallel shards", async () => {
		vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] })

		const files = {
			"m-00001-of-00003.gguf": content(5000, 1),
			"m-00002-of-00003.gguf": content(6000, 2),
			"m-00003-of-00003.gguf": content(7000, 3)
		}
		const manifest = manifestOf(Object.entries(files).map(([name, bytes]) => describeShard(name, bytes)))
		const store = memoryStore()

		// A partial first shard from an earlier session.
		store.files.set("m-00001-of-00003.gguf", Object.assign(new MemoryFile(), { bytes: files["m-00001-of-00003.gguf"].slice(0, 1500) }))

		const answered = new Map<string, number>()
		const { fetchShard } = scriptedHost(files, (_request, file) => {
			const count = answered.get(file) ?? 0

			answered.set(file, count + 1)

			return count === 0 ? { deliver: 2000, then: file.includes("00002") ? "hang" : "error" } : {}
		})
		const progress: number[] = []
		await drain(run({ manifest, store, fetchShard, onProgress: ({ receivedBytes }) => progress.push(receivedBytes) }))

		expect(progress[0]).toBe(1500)
		expect(Math.max(...progress)).toBe(18_000)
		expect(progress.at(-1)).toBe(18_000)
		expect(progress.every((bytes, index) => index === 0 || bytes >= (progress[index - 1] ?? 0))).toBe(true)
	})

	it("reports waiting only while no shard in flight is moving", async () => {
		const bytes = content(3000)
		let online = false
		let comeOnline: () => void = () => undefined
		const connectivity: Connectivity = {
			isOnline: () => online,
			waitOnline: () =>
				new Promise(resolve => {
					comeOnline = () => {
						online = true
						resolve()
					}
				})
		}
		const waiting: boolean[] = []
		const done = run({
			manifest: manifestOf([describeShard("a.gguf", bytes)]),
			store: memoryStore(),
			fetchShard: scriptedHost({ "a.gguf": bytes }).fetchShard,
			connectivity,
			onProgress: progress => waiting.push(progress.waiting)
		})

		await vi.waitFor(() => {
			expect(waiting).toContain(true)
		})
		comeOnline()
		await done

		expect(waiting.at(-1)).toBe(false)
	})

	it("reports a full disk as storage-full", async () => {
		const bytes = content(3000)
		const store = memoryStore()
		const file = new MemoryFile()

		file.quotaAfter = 1500
		store.files.set("a.gguf", file)

		await expect(
			run({ manifest: manifestOf([describeShard("a.gguf", bytes)]), store, fetchShard: scriptedHost({ "a.gguf": bytes }).fetchShard })
		).rejects.toMatchObject({ kind: "storage-full" })
	})

	it("refetches a corrupted shard once from scratch", async () => {
		const bytes = content(900)
		const corrupt = bytes.slice()

		corrupt[10] = (corrupt[10] ?? 0) ^ 0xff

		let served = 0
		const fetchShard = vi.fn(() => {
			served++

			return Promise.resolve(new Response(served === 1 ? corrupt : bytes, { status: 200, headers: { "Content-Length": "900" } }))
		})
		const progress: number[] = []
		const store = memoryStore()

		await run({
			manifest: manifestOf([describeShard("a.gguf", bytes)]),
			store,
			fetchShard,
			onProgress: ({ receivedBytes }) => progress.push(receivedBytes)
		})

		expect(fetchShard).toHaveBeenCalledTimes(2)
		expect(store.files.get("a.gguf")?.bytes).toEqual(bytes)
		expect(progress).toContain(0)
		expect(Math.max(...progress)).toBe(900)
	})

	it("fails when the refetched shard is still corrupt", async () => {
		const bytes = content(900)
		const { fetchShard } = scriptedHost({ "a.gguf": bytes })

		await expect(
			run({ manifest: manifestOf([{ ...describeShard("a.gguf", bytes), sha256: "1".repeat(64) }]), store: memoryStore(), fetchShard })
		).rejects.toMatchObject({ kind: "integrity" })
		expect(fetchShard).toHaveBeenCalledTimes(2)
	})
})
