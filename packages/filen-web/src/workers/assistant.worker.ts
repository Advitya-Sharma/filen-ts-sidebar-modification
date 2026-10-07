import * as Comlink from "comlink"
import { createSHA256 } from "hash-wasm"
import { Wllama, type ChatCompletionTool, type WllamaLogger } from "@wllama/wllama/esm/index.js"
// Emitted by Vite as a hashed asset of this build: the engine is self-hosted, never fetched from a CDN.
import wllamaWasmUrl from "@wllama/wllama/esm/wasm/wllama.wasm?url"
import type { AssistantGenerateRequest, AssistantGenerateResult } from "@filen/shared/assistant"
import type { AssistantModelManifest } from "@/features/assistant/runtime/manifest"
import { accumulateChunk, toEngineMessage, type GenerationState } from "@/features/assistant/runtime/engineMessages"
import type { AssistantBackend } from "@/features/assistant/runtime/deviceCheck"
import { downloadModelShards, type ShardStore } from "@/features/assistant/runtime/download"
import { blobGgufRead, readGgufSpecialTokens } from "@/features/assistant/runtime/gguf"
import { modelDirectory, readVerified } from "@/features/assistant/runtime/modelStore"
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

const quietLogger: WllamaLogger = {
	debug: () => undefined,
	log: () => undefined,
	warn: (...args: unknown[]) => {
		log.warn("assistant.engine", ...args)
	},
	error: (...args: unknown[]) => {
		log.error("assistant.engine", ...args)
	}
}

let engine: Wllama | null = null
let download: AbortController | null = null
const generations = new Map<number, AbortController>()

// verified.json through a sync access handle (createWritable is missing in some engines); a torn write
// only means the shards are hashed again on the next load.
async function writeVerified({
	directory,
	verified
}: {
	directory: FileSystemDirectoryHandle
	verified: ReadonlyMap<string, string>
}): Promise<void> {
	const access = await (await directory.getFileHandle("verified.json", { create: true })).createSyncAccessHandle()

	try {
		access.truncate(0)
		access.write(new TextEncoder().encode(JSON.stringify(Object.fromEntries(verified))), { at: 0 })
		access.flush()
	} finally {
		access.close()
	}
}

async function shardStore(manifest: AssistantModelManifest): Promise<{ store: ShardStore; directory: FileSystemDirectoryHandle }> {
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

	return {
		directory,
		store: {
			isVerified: shard => sizes.get(shard.file) === shard.size && verified.get(shard.file) === shard.sha256,
			markVerified: shard => {
				verified.set(shard.file, shard.sha256)
				sizes.set(shard.file, shard.size)
				writing = writing.then(() => writeVerified({ directory, verified }))

				return writing
			},
			withFile: async (shard, run) => {
				const access = await (await directory.getFileHandle(shard.file, { create: true })).createSyncAccessHandle()

				try {
					return await run(access)
				} finally {
					access.close()
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

// The model's shards as Files, in order. A shard on disk at full size but missing from the verified record
// (a crash between the last write and the record) is hashed now, once.
async function openShards(manifest: AssistantModelManifest): Promise<File[]> {
	const { store, directory } = await shardStore(manifest)
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
			if ((await hashFile(file)) !== shard.sha256) {
				throw loadError("missing", `${shard.file} does not match its published hash`)
			}

			await store.markVerified(shard)
		}

		files.push(file)
	}

	return files
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

function isOutOfMemory(e: unknown): boolean {
	return /out of memory|cannot allocate|allocation failed|oom/i.test(e instanceof Error ? e.message : String(e))
}

const api = {
	// Resumable: shards already verified are skipped, a partial shard continues where it stopped.
	async download(manifest: AssistantModelManifest, baseUrl: string, onProgress: (receivedBytes: number) => void): Promise<void> {
		download?.abort()

		const controller = new AbortController()

		download = controller

		const progress = throttle(onProgress, PROGRESS_THROTTLE_MS)

		try {
			const { store } = await shardStore(manifest)

			await downloadModelShards({
				manifest,
				baseUrl,
				store,
				createHasher: createSHA256,
				fetchShard: (url, init) => fetch(url, init),
				signal: controller.signal,
				onProgress: progress
			})
		} finally {
			progress.flush()
			releaseCallback(onProgress)

			if (download === controller) {
				download = null
			}
		}
	},

	cancelDownload(): void {
		download?.abort()
		download = null
	},

	async load(manifest: AssistantModelManifest, options: AssistantLoadOptions): Promise<AssistantLoadResult> {
		if (engine !== null) {
			throw loadError("load-failed", "A model is already loaded")
		}

		const files = await openShards(manifest)
		const first = files[0]

		if (first === undefined) {
			throw loadError("missing", "The model has no shards")
		}

		const specialTokens = await readGgufSpecialTokens({ read: blobGgufRead(first), size: first.size })
		const instance = new Wllama({ default: wllamaWasmUrl }, { suppressNativeLog: true, logger: quietLogger })

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
				// Router, main loop and quarantine alternate; a slot each keeps their prompt prefixes cached.
				n_parallel: 3,
				kv_unified: true
			})
		} catch (e) {
			await instance.exit().catch(() => undefined)

			throw loadError(isOutOfMemory(e) ? "out-of-memory" : "load-failed", e instanceof Error ? e.message : String(e))
		}

		engine = instance

		return { specialTokens }
	},

	async generate(id: number, request: AssistantGenerateRequest, onText?: (text: string) => void): Promise<AssistantGenerateResult> {
		if (engine === null) {
			throw loadError("missing", "No model is loaded")
		}

		const controller = new AbortController()
		const state: GenerationState = { content: "", calls: [] }
		const tools: ChatCompletionTool[] | undefined = request.tools?.map(tool => ({
			type: "function",
			function: { name: tool.name, description: tool.description, parameters: tool.parameters }
		}))

		const text = onText === undefined ? null : throttle(onText, TEXT_THROTTLE_MS)

		generations.set(id, controller)

		try {
			await engine.createChatCompletion({
				messages: request.messages.map(toEngineMessage),
				...(tools === undefined || tools.length === 0 ? {} : { tools, tool_choice: "auto" as const }),
				...(request.grammar === undefined ? {} : { grammar: request.grammar }),
				max_tokens: request.maxTokens,
				temperature: 0,
				chat_template_kwargs: { enable_thinking: false },
				abortSignal: controller.signal,
				stream: true,
				onData: chunk => {
					if (accumulateChunk({ state, chunk })) {
						text?.(state.content)
					}
				}
			})
		} finally {
			text?.flush()
			releaseCallback(onText)
			generations.delete(id)
		}

		return {
			content: state.content,
			toolCalls: state.calls.filter(call => call.name.length > 0)
		}
	},

	cancel(id: number): void {
		generations.get(id)?.abort()
	},

	async unload(): Promise<void> {
		for (const controller of generations.values()) {
			controller.abort()
		}

		const current = engine

		engine = null
		await current?.exit()
	}
}

export type AssistantWorkerApi = typeof api

Comlink.expose(api)
