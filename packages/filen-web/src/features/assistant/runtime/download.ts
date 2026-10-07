import { modelFileUrl, type AssistantModelFile, type AssistantModelManifest } from "@/features/assistant/runtime/manifest"

// The resumable, verified shard download, written against small interfaces so it runs on OPFS sync
// access handles in the worker and on in-memory fakes in tests.
//
// The model host only allows simple CORS requests (no preflight: it answers OPTIONS with 405) and does not
// expose Content-Range. So the only request header ever sent is a single "Range: bytes=N-", which is
// CORS-safelisted, and a resumed response is recognized by status 206 and its Content-Length alone.

// The subset of FileSystemSyncAccessHandle the download uses.
export interface ShardFile {
	getSize(): number
	truncate(size: number): void
	read(buffer: Uint8Array, options: { at: number }): number
	write(buffer: Uint8Array, options: { at: number }): number
	flush(): void
}

export interface StreamingHasher {
	init(): unknown
	update(data: Uint8Array): unknown
	digest(outputType: "hex"): string
}

export type ModelDownloadErrorKind = "network" | "integrity" | "storage-full"

export class ModelDownloadError extends Error {
	readonly kind: ModelDownloadErrorKind

	constructor(kind: ModelDownloadErrorKind, message: string) {
		super(message)
		this.name = `ModelDownloadError:${kind}`
		this.kind = kind
	}
}

const READ_CHUNK = 8 * 1024 * 1024
// Durability checkpoint: what a crash can lose of a shard.
const FLUSH_EVERY = 32 * 1024 * 1024

// Chromium raises it from write() or only from the next flush(), and as its own DOMException subclass.
function isQuotaError(e: unknown): boolean {
	return e instanceof Error && e.name === "QuotaExceededError"
}

// Hashes what a previous attempt left on disk, so the download continues where it stopped.
function hashExisting({ file, size, hasher }: { file: ShardFile; size: number; hasher: StreamingHasher }): void {
	const buffer = new Uint8Array(Math.min(READ_CHUNK, size))

	for (let at = 0; at < size; at += buffer.length) {
		const view = buffer.subarray(0, Math.min(buffer.length, size - at))

		if (file.read(view, { at }) !== view.length) {
			throw new ModelDownloadError("network", "Short read of a partial shard")
		}

		hasher.update(view)
	}
}

// Downloads one shard into `file` (resuming from its current size), verifying the sha256 over the whole
// file. Leaves a complete, verified shard, or throws; on a hash mismatch the file is emptied.
export async function downloadShard({
	shard,
	url,
	file,
	hasher,
	fetchShard,
	signal,
	onBytes
}: {
	shard: AssistantModelFile
	url: string
	file: ShardFile
	hasher: StreamingHasher
	fetchShard: (url: string, init: RequestInit) => Promise<Response>
	signal: AbortSignal
	// Bytes newly accounted for, including a resumed prefix.
	onBytes: (bytes: number) => void
}): Promise<void> {
	let size = file.getSize()

	if (size > shard.size) {
		file.truncate(0)
		size = 0
	}

	hasher.init()
	hashExisting({ file, size, hasher })
	onBytes(size)

	if (size < shard.size) {
		const response = await fetchShard(url, {
			signal,
			credentials: "omit",
			// The shard lands in OPFS; a second copy in the HTTP cache would only take disk.
			cache: "no-store",
			...(size > 0 ? { headers: { Range: `bytes=${String(size)}-` } } : {})
		}).catch((e: unknown) => {
			throw signal.aborted ? e : new ModelDownloadError("network", `Fetching ${shard.file} failed`)
		})

		if (size > 0 && response.status === 200) {
			// The host ignored the range: start the shard over.
			onBytes(-size)
			file.truncate(0)
			size = 0
			hasher.init()
		} else if (response.status !== (size > 0 ? 206 : 200)) {
			throw new ModelDownloadError("network", `Fetching ${shard.file} answered ${String(response.status)}`)
		}

		const length = Number(response.headers.get("Content-Length") ?? Number.NaN)

		if (!Number.isNaN(length) && length !== shard.size - size) {
			throw new ModelDownloadError("network", `${shard.file} has an unexpected length`)
		}

		if (response.body === null) {
			throw new ModelDownloadError("network", `${shard.file} has no body`)
		}

		const reader = response.body.getReader()
		let unflushed = 0

		try {
			for (;;) {
				const { done, value } = await reader.read()

				if (done) {
					break
				}

				if (size + value.length > shard.size) {
					throw new ModelDownloadError("integrity", `${shard.file} is longer than published`)
				}

				file.write(value, { at: size })
				hasher.update(value)
				size += value.length
				unflushed += value.length
				onBytes(value.length)

				if (unflushed >= FLUSH_EVERY) {
					file.flush()
					unflushed = 0
				}
			}
		} finally {
			reader.releaseLock()
		}

		file.flush()
	}

	if (size !== shard.size) {
		throw new ModelDownloadError("network", `${shard.file} ended early`)
	}

	if (hasher.digest("hex") !== shard.sha256) {
		file.truncate(0)
		file.flush()
		onBytes(-size)

		throw new ModelDownloadError("integrity", `${shard.file} does not match its published hash`)
	}
}

export interface ShardStore {
	isVerified(shard: AssistantModelFile): boolean
	markVerified(shard: AssistantModelFile): Promise<void>
	// Opens the shard's file for the duration of `run`.
	withFile<T>(shard: AssistantModelFile, run: (file: ShardFile) => Promise<T>): Promise<T>
}

const PARALLEL_SHARDS = 2

// Downloads every shard not yet verified, two at a time, reporting the bytes on disk as they grow. A shard
// whose hash mismatches is fetched once more from scratch before the download fails.
export async function downloadModelShards({
	manifest,
	baseUrl,
	store,
	createHasher,
	fetchShard,
	signal,
	onProgress
}: {
	manifest: AssistantModelManifest
	baseUrl: string
	store: ShardStore
	createHasher: () => Promise<StreamingHasher>
	fetchShard: (url: string, init: RequestInit) => Promise<Response>
	signal: AbortSignal
	onProgress: (receivedBytes: number) => void
}): Promise<void> {
	let received = manifest.shards.filter(shard => store.isVerified(shard)).reduce((sum, shard) => sum + shard.size, 0)
	const pending = manifest.shards.filter(shard => !store.isVerified(shard))

	onProgress(received)

	// One failed shard stops the other in flight.
	const failed = new AbortController()
	const shardSignal = AbortSignal.any([signal, failed.signal])
	const worker = async () => {
		for (let shard = pending.shift(); shard !== undefined; shard = pending.shift()) {
			const current = shard
			const hasher = await createHasher()

			for (let attempt = 0; ; attempt++) {
				shardSignal.throwIfAborted()

				try {
					await store.withFile(current, file =>
						downloadShard({
							shard: current,
							url: modelFileUrl({ baseUrl, file: current }),
							file,
							hasher,
							fetchShard,
							signal: shardSignal,
							onBytes: bytes => {
								received += bytes
								onProgress(received)
							}
						})
					)

					break
				} catch (e) {
					if (isQuotaError(e)) {
						throw new ModelDownloadError("storage-full", "Storage is full")
					}

					if (attempt > 0 || !(e instanceof ModelDownloadError) || e.kind !== "integrity") {
						throw e
					}
				}
			}

			await store.markVerified(current).catch((e: unknown) => {
				throw isQuotaError(e) ? new ModelDownloadError("storage-full", "Storage is full") : e
			})
		}
	}

	await Promise.all(
		Array.from({ length: Math.min(PARALLEL_SHARDS, pending.length) }, () =>
			worker().catch((e: unknown) => {
				failed.abort()

				throw e
			})
		)
	)
}
