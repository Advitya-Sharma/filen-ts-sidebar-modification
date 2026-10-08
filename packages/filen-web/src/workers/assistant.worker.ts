import * as Comlink from "comlink"
import { createSHA256 } from "hash-wasm"
import { LogLevel, Wllama, WllamaAbortError, WllamaError, type ChatCompletionTool, type WllamaLogger } from "@wllama/wllama/esm/index.js"
// Emitted by Vite as a hashed asset of this build: the engine is self-hosted, never fetched from a CDN.
import wllamaWasmUrl from "@wllama/wllama/esm/wasm/wllama.wasm?url"
import type { AssistantGenerateRequest, AssistantGenerateResult } from "@filen/shared/assistant"
import type { AssistantModelFile, AssistantModelManifest } from "@/features/assistant/runtime/manifest"
import {
	accumulateChunk,
	ENGINE_CRASHED,
	toEngineMessage,
	type GenerationState,
	type GenerationUpdate
} from "@/features/assistant/runtime/engineMessages"
import type { AssistantBackend } from "@/features/assistant/runtime/deviceCheck"
import { downloadModelShards, globalConnectivity, type ModelDownloadProgress, type ShardStore } from "@/features/assistant/runtime/download"
import { blobGgufRead, readGgufSpecialTokens } from "@/features/assistant/runtime/gguf"
import { modelDirectory, readVerified, VERIFIED_FILE } from "@/features/assistant/runtime/modelStore"
import { log } from "@/lib/log"
import { PROGRESS_THROTTLE_MS, throttle } from "@/lib/throttle"

// Runs the on-device model: its own module worker (and its own bundle, so the engine's ~380 KB never
// reaches the page) that downloads the model into OPFS and drives wllama. wllama starts its own nested
// worker for the wasm runtime and, on the CPU path, a pthread pool; those are blob: workers, which
// worker-src already allows. Only one tab at a time talks to this worker (runtime/client.ts holds a Web
// Lock), so the OPFS files are never written by two contexts.

export interface AssistantLoadOptions {
	backend: AssistantBackend
	threads: number
	contextTokens: number
}

export interface AssistantLoadResult {
	// The model's own special-token strings, for sanitizeUntrustedText.
	specialTokens: string[]
}

// Errors cross Comlink as { name, message }; the client reads the kind from the name.
export type AssistantLoadErrorKind = "missing" | "out-of-memory" | "load-failed"

function loadError(kind: AssistantLoadErrorKind, message: string): Error {
	const error = new Error(message)

	error.name = `ModelLoadError:${kind}`

	return error
}

function crashError(message: string): Error {
	const error = new Error(message)

	error.name = ENGINE_CRASHED

	return error
}

// The engine's own warning and error lines, also kept (the newest few) to tell why a load failed: llama.cpp
// reports a failed load only in its log.
const ENGINE_LINES_KEPT = 8

function engineLogger(lines: string[]): WllamaLogger {
	const keep = (args: unknown[]) => {
		lines.push(args.map(String).join(" "))

		if (lines.length > ENGINE_LINES_KEPT) {
			lines.shift()
		}
	}

	return {
		debug: () => undefined,
		log: () => undefined,
		warn: (...args: unknown[]) => {
			keep(args)
			log.warn("assistant.engine", ...args)
		},
		error: (...args: unknown[]) => {
			keep(args)
			log.error("assistant.engine", ...args)
		}
	}
}

let engine: Wllama | null = null
let download: { controller: AbortController; done: Promise<void> } | null = null
// Each generation's controller and its settling, so an unload can let a cancelled one wind down first.
const generations = new Map<number, { controller: AbortController; done: Promise<unknown> }>()
// Generations run one at a time: the client rejects a stopped one at once, and the next must not start on
// the engine while that one is still winding down.
let generationTail: Promise<unknown> = Promise.resolve()

// How long an unload waits for cancelled generations: wllama checks the signal once per engine step.
const SETTLE_GRACE_MS = 500

// Read through a call: TypeScript would keep `download` narrowed across the awaits before it.
function clearDownload(controller: AbortController): void {
	if (download?.controller === controller) {
		download = null
	}
}

// The hasher state of a partial shard (download.ts), beside it; gone once the shard is verified.
function checkpointName(shard: AssistantModelFile): string {
	return `${shard.file}.sha256-state`
}

// verified.json through a sync access handle (createWritable is missing in some engines). Written over the
// old record and then cut to length, never emptied first; a torn record reads as empty, which costs one
// hash of the complete shards at the next load, never a shard wrongly trusted.
async function writeVerified({
	directory,
	verified
}: {
	directory: FileSystemDirectoryHandle
	verified: ReadonlyMap<string, string>
}): Promise<void> {
	const access = await (await directory.getFileHandle(VERIFIED_FILE, { create: true })).createSyncAccessHandle()

	try {
		const bytes = new TextEncoder().encode(JSON.stringify(Object.fromEntries(verified)))

		access.write(bytes, { at: 0 })
		access.truncate(bytes.length)
		access.flush()
	} finally {
		access.close()
	}
}

interface WorkerShardStore extends ShardStore {
	// Drops a shard that no longer matches: its file and its entry in the record.
	discard(shard: AssistantModelFile): Promise<void>
}

async function shardStore(manifest: AssistantModelManifest): Promise<{ store: WorkerShardStore; directory: FileSystemDirectoryHandle }> {
	const directory = await modelDirectory(manifest.id)
	const verified = await readVerified(directory)
	const sizes = new Map<string, number>()

	for (const shard of manifest.shards) {
		sizes.set(
			shard.file,
			await directory
				.getFileHandle(shard.file)
				.then(handle => handle.getFile())
				.then(
					file => file.size,
					() => 0
				)
		)
	}

	let writing = Promise.resolve()
	const record = () => {
		writing = writing.then(() => writeVerified({ directory, verified }))

		return writing
	}

	return {
		directory,
		store: {
			isVerified: shard => sizes.get(shard.file) === shard.size && verified.get(shard.file) === shard.sha256,
			sizeOnDisk: shard => sizes.get(shard.file) ?? 0,
			markVerified: shard => {
				verified.set(shard.file, shard.sha256)
				sizes.set(shard.file, shard.size)

				return record()
			},
			discard: async shard => {
				verified.delete(shard.file)
				sizes.set(shard.file, 0)
				await record()
				await directory.removeEntry(shard.file).catch(() => undefined)
			},
			withFile: async (shard, run) => {
				const data = await (await directory.getFileHandle(shard.file, { create: true })).createSyncAccessHandle()
				let checkpoint: FileSystemSyncAccessHandle

				try {
					checkpoint = await (await directory.getFileHandle(checkpointName(shard), { create: true })).createSyncAccessHandle()
				} catch (e) {
					data.close()

					throw e
				}

				let complete = false

				try {
					const result = await run(data, checkpoint)

					complete = checkpoint.getSize() === 0

					return result
				} finally {
					data.close()
					checkpoint.close()

					if (complete) {
						await directory.removeEntry(checkpointName(shard)).catch(() => undefined)
					}
				}
			}
		}
	}
}

async function hashFile(file: File): Promise<string> {
	const hasher = await createSHA256()
	const reader = file.stream().getReader()

	hasher.init()

	for (;;) {
		const { done, value } = await reader.read()

		if (done) {
			return hasher.digest("hex")
		}

		hasher.update(value)
	}
}

// The model's shards as Files, in order. A shard is trusted on its size and the verified record, which
// the download writes only after the shard's hash matched; one at full size but missing from the record
// (a crash between the last write and the record) is hashed now, once, and dropped if it does not match.
async function openShards({
	manifest,
	store,
	directory,
	onVerifying
}: {
	manifest: AssistantModelManifest
	store: WorkerShardStore
	directory: FileSystemDirectoryHandle
	onVerifying: () => void
}): Promise<File[]> {
	const files: File[] = []

	for (const shard of manifest.shards) {
		const file = await directory
			.getFileHandle(shard.file)
			.then(handle => handle.getFile())
			.catch(() => null)

		if (file?.size !== shard.size) {
			throw loadError("missing", `${shard.file} is not downloaded`)
		}

		if (!store.isVerified(shard)) {
			onVerifying()

			if ((await hashFile(file)) !== shard.sha256) {
				await store.discard(shard)

				throw loadError("missing", `${shard.file} does not match its published hash`)
			}

			await store.markVerified(shard)
		}

		files.push(file)
	}

	return files
}

// After a load that failed for a reason other than memory: a shard damaged on disk since it was verified
// is one cause the record cannot show. Hashes every shard once and drops those that no longer match, so the
// client downloads just them again. Whether any were dropped.
async function discardDamagedShards({
	manifest,
	store,
	files
}: {
	manifest: AssistantModelManifest
	store: WorkerShardStore
	files: readonly File[]
}): Promise<boolean> {
	let damaged = false

	for (const [index, shard] of manifest.shards.entries()) {
		const file = files[index]

		if (file !== undefined && (await hashFile(file)) !== shard.sha256) {
			await store.discard(shard)
			damaged = true
		}
	}

	return damaged
}

// A callback the page passed with Comlink.proxy holds a MessagePort until released.
function releaseCallback(callback: unknown): void {
	const release = (callback as Record<symbol, unknown> | undefined)?.[Comlink.releaseProxy]

	if (typeof release === "function") {
		;(release as () => void)()
	}
}

// Streamed text reaches the page at most this often; the last state always arrives before the result.
const TEXT_THROTTLE_MS = 50

function isOutOfMemory(text: string): boolean {
	return /out of memory|cannot allocate|allocation failed|failed to allocate|unable to allocate|reserve memory|\boom\b/i.test(text)
}

// A generation failure that leaves the engine unusable, as opposed to a stop or a request it refused: wllama
// rejects everything with a RuntimeError once its wasm aborted, and answers later calls with a bare string.
function isEngineCrash(e: unknown, aborted: boolean): boolean {
	if (aborted || e instanceof WllamaAbortError || (e instanceof Error && e.name === "AbortError")) {
		return false
	}

	if (e instanceof WllamaError) {
		return e.type !== "inference_error" && e.type !== "kv_cache_full"
	}

	return !(e instanceof Error) || e.name === "RuntimeError" || /crashed/i.test(e.message)
}

function delay(ms: number): Promise<void> {
	return new Promise(resolve => {
		setTimeout(resolve, ms)
	})
}

const api = {
	// Resumable: shards already verified are skipped, a partial shard continues where it stopped.
	async download(
		manifest: AssistantModelManifest,
		baseUrl: string,
		onProgress: (progress: ModelDownloadProgress) => void
	): Promise<void> {
		const previous = download
		const controller = new AbortController()
		const progress = throttle(onProgress, PROGRESS_THROTTLE_MS)
		const run = (async () => {
			// Its files stay open until it has wound down; a second sync access handle on them would fail.
			await previous?.done.catch(() => undefined)
			controller.signal.throwIfAborted()

			const { store } = await shardStore(manifest)

			await downloadModelShards({
				manifest,
				baseUrl,
				store,
				createHasher: createSHA256,
				fetchShard: (url, init) => fetch(url, init),
				connectivity: globalConnectivity(globalThis),
				signal: controller.signal,
				onProgress: progress
			})
		})()

		previous?.controller.abort()
		download = { controller, done: run }

		try {
			await run
		} finally {
			progress.flush()
			releaseCallback(onProgress)

			clearDownload(controller)
		}
	},

	cancelDownload(): void {
		download?.controller.abort()
	},

	// `onVerifying` fires when shards have to be hashed before the model can start.
	async load(manifest: AssistantModelManifest, options: AssistantLoadOptions, onVerifying?: () => void): Promise<AssistantLoadResult> {
		try {
			if (engine !== null) {
				throw loadError("load-failed", "A model is already loaded")
			}

			const { store, directory } = await shardStore(manifest)
			const files = await openShards({ manifest, store, directory, onVerifying: () => onVerifying?.() })
			const first = files[0]

			if (first === undefined) {
				throw loadError("missing", "The model has no shards")
			}

			const lines: string[] = []
			const instance = new Wllama({ default: wllamaWasmUrl }, { suppressNativeLog: false, logger: engineLogger(lines) })
			// Independent of the engine start, so it runs while the engine reads the weights.
			const specialTokens = readGgufSpecialTokens({ read: blobGgufRead(first), size: first.size })

			// Never the CDN-hosted compat build: a browser that would need it is refused by the device check.
			instance.setCompat(null)

			try {
				await instance.loadModel(files, {
					n_ctx: options.contextTokens,
					n_threads: options.threads,
					n_gpu_layers: options.backend === "webgpu" ? 999 : 0,
					jinja: true,
					// Non-thinking mode: the small Qwen3.5 models loop when allowed to think.
					reasoning: false,
					// Prompt progress is reported once per decoded batch: at the micro-batch size the compute is the
					// same as the 2048 default, and a long prompt reports in steps instead of one jump.
					n_batch: 512,
					n_ubatch: 512,
					// Router, main loop and quarantine alternate; a slot each keeps their prompt prefixes cached. The
					// server's default would copy every idle slot to a host-memory prompt cache and clear it on each
					// new request, which grows the heap and re-reads every prompt.
					n_parallel: 3,
					kv_unified: true,
					cache_idle_slots: false,
					log_level: LogLevel.ERROR
				})

				// wllama does not check the engine's own load result: a model that failed to load (memory, a damaged
				// file) still resolves, with an empty vocabulary and context.
				const info = instance.getLoadedContextInfo()

				if (!(info.n_vocab > 0 && info.n_ctx > 0)) {
					throw new Error(lines.join("\n") || "The engine did not load the model")
				}
			} catch (e) {
				await instance.exit().catch(() => undefined)
				await specialTokens.catch(() => undefined)

				const message = [e instanceof Error ? e.message : String(e), ...lines].join("\n")

				if (isOutOfMemory(message)) {
					throw loadError("out-of-memory", message)
				}

				// A failure on the GPU is the GPU's: the client tries the CPU path next, and only a failure there
				// (or on a CPU-only device) reads every shard again for damage.
				if (options.backend === "cpu") {
					onVerifying?.()

					if (await discardDamagedShards({ manifest, store, files })) {
						throw loadError("missing", "A downloaded shard was damaged")
					}
				}

				throw loadError("load-failed", message)
			}

			engine = instance

			return { specialTokens: await specialTokens }
		} finally {
			releaseCallback(onVerifying)
		}
	},

	// `onUpdate` gets the streamed text (throttled), the prompt-reading progress and each tool call as soon
	// as its name is known; the text written before a call always arrives before the call.
	async generate(
		id: number,
		request: AssistantGenerateRequest,
		onUpdate?: (update: GenerationUpdate) => void
	): Promise<AssistantGenerateResult> {
		if (engine === null) {
			throw loadError("missing", "No model is loaded")
		}

		const controller = new AbortController()
		const state: GenerationState = { content: "", calls: [] }
		const tools: ChatCompletionTool[] | undefined = request.tools?.map(tool => ({
			type: "function",
			function: { name: tool.name, description: tool.description, parameters: tool.parameters }
		}))

		const text =
			onUpdate === undefined
				? null
				: throttle((content: string) => {
						onUpdate({ type: "text", text: content })
					}, TEXT_THROTTLE_MS)

		const current = engine
		const previous = generationTail
		let finish: () => void = () => undefined
		const done = new Promise<void>(resolve => {
			finish = resolve
		})

		generationTail = done
		generations.set(id, { controller, done })

		try {
			await previous

			if (controller.signal.aborted) {
				throw new WllamaAbortError()
			}

			await current.createChatCompletion({
				messages: request.messages.map(toEngineMessage),
				...(tools === undefined || tools.length === 0 ? {} : { tools, tool_choice: "auto" as const }),
				...(request.grammar === undefined ? {} : { grammar: request.grammar }),
				max_tokens: request.maxTokens,
				temperature: 0,
				chat_template_kwargs: { enable_thinking: false },
				abortSignal: controller.signal,
				stream: true,
				return_progress: onUpdate !== undefined,
				onData: chunk => {
					const changes = accumulateChunk({ state, chunk })

					if (onUpdate === undefined) {
						return
					}

					if (changes.progress !== null) {
						onUpdate({ type: "progress", ...changes.progress })
					}

					if (changes.text) {
						text?.(state.content)
					}

					if (changes.named.length > 0) {
						text?.flush()

						for (const call of changes.named) {
							onUpdate({ type: "call", ...call })
						}
					}
				}
			})
		} catch (e) {
			if (isEngineCrash(e, controller.signal.aborted)) {
				// Released here, so the client finds no engine to reuse; its files stay on disk.
				if (engine === current) {
					engine = null
				}

				await current.exit().catch(() => undefined)

				throw crashError(e instanceof Error ? e.message : String(e))
			}

			throw e
		} finally {
			text?.flush()
			releaseCallback(onUpdate)
			generations.delete(id)
			finish()
		}

		return {
			content: state.content,
			toolCalls: state.calls.filter(call => call.name.length > 0),
			...(state.truncated ? { truncated: true } : {})
		}
	},

	cancel(id: number): void {
		generations.get(id)?.controller.abort()
	},

	// Cancelled generations get a short while to settle first: exiting under one leaves it waiting forever on a
	// worker that is gone.
	async unload(): Promise<void> {
		const running = [...generations.values()]

		for (const { controller } of running) {
			controller.abort()
		}

		const current = engine

		engine = null
		await Promise.race([Promise.allSettled(running.map(({ done }) => done)), delay(SETTLE_GRACE_MS)])
		await current?.exit()
	}
}

export type AssistantWorkerApi = typeof api

Comlink.expose(api)
