import { createHash } from "node:crypto"
import { beforeEach, describe, expect, it, vi } from "vitest"
import type { AssistantWorkerApi } from "@/workers/assistant.worker"
import type { AssistantModelManifest } from "@/features/assistant/runtime/manifest"
import type { GenerationUpdate } from "@/features/assistant/runtime/engineMessages"

// The worker module exposes its API through Comlink at import; the engine, its wasm URL and OPFS are
// replaced so the contract between the page and the engine can be checked under node.
const { exposed, engines, shardFiles, verified, loadFailure, loadOutcome, completion } = vi.hoisted(() => ({
	exposed: new Map<"api", unknown>(),
	engines: [] as {
		loadModel: ReturnType<typeof vi.fn>
		createChatCompletion: ReturnType<typeof vi.fn>
		setCompat: ReturnType<typeof vi.fn>
		exit: ReturnType<typeof vi.fn>
	}[],
	// The model directory: file name to contents.
	shardFiles: new Map<string, Uint8Array<ArrayBuffer>>(),
	verified: new Map<string, string>(),
	loadFailure: { error: null as Error | null },
	// What a load that "succeeded" reports, and the engine log line it wrote.
	loadOutcome: { vocab: 10, log: null as string | null },
	// Replaces the scripted completion when set.
	completion: { run: null as ((options: { abortSignal: AbortSignal; onData: (chunk: unknown) => void }) => Promise<void>) | null }
}))

vi.mock("comlink", () => ({
	expose: (api: unknown) => {
		exposed.set("api", api)
	},
	releaseProxy: Symbol("releaseProxy")
}))

vi.mock("@wllama/wllama/esm/wasm/wllama.wasm?url", () => ({ default: "/assets/wllama-hash.wasm" }))

vi.mock("@wllama/wllama/esm/index.js", () => ({
	LogLevel: { DEBUG: 1, INFO: 2, WARN: 3, ERROR: 4 },
	WllamaAbortError: class extends Error {
		override name = "AbortError"
	},
	WllamaError: class extends Error {
		type: string

		constructor(message: string, type = "unknown_error") {
			super(message)
			this.type = type
		}
	},
	Wllama: class {
		readonly config: { logger: { error: (...args: unknown[]) => void } }
		readonly loadModel = vi.fn(() => {
			if (loadOutcome.log !== null) {
				this.config.logger.error(loadOutcome.log)
			}

			return loadFailure.error === null ? Promise.resolve() : Promise.reject(loadFailure.error)
		})
		readonly getLoadedContextInfo = vi.fn(() => ({ n_vocab: loadOutcome.vocab, n_ctx: loadOutcome.vocab > 0 ? 8192 : 0 }))
		readonly setCompat = vi.fn()
		readonly exit = vi.fn(() => Promise.resolve())
		readonly createChatCompletion = vi.fn((options: { onData: (chunk: unknown) => void; abortSignal: AbortSignal }): Promise<void> => {
			if (completion.run !== null) {
				return completion.run(options)
			}

			const chunk = (delta: unknown) => ({ choices: [{ index: 0, delta, finish_reason: null, logprobs: null }] })

			options.onData({
				...chunk({ role: "assistant", content: null }),
				prompt_progress: { total: 40, cache: 8, processed: 8, time_ms: 0 }
			})
			options.onData({
				...chunk({ role: "assistant", content: null }),
				prompt_progress: { total: 40, cache: 8, processed: 40, time_ms: 5 }
			})
			options.onData(chunk({ content: "Hi" }))
			options.onData(chunk({ tool_calls: [{ index: 0, id: "t", function: { name: "search_items", arguments: '{"query":' } }] }))
			options.onData(chunk({ tool_calls: [{ index: 0, function: { arguments: '"x"}' } }] }))

			return Promise.resolve()
		})

		constructor(_paths: unknown, config: { logger: { error: (...args: unknown[]) => void } }) {
			this.config = config
			engines.push(this)
		}
	}
}))

// OPFS with sync access handles over shardFiles.
function syncHandle(name: string) {
	const bytes = () => shardFiles.get(name) ?? new Uint8Array(0)

	return {
		getSize: () => bytes().length,
		truncate: (size: number) => {
			const next = new Uint8Array(size)

			next.set(bytes().subarray(0, size))
			shardFiles.set(name, next)
		},
		read: (buffer: Uint8Array, { at }: { at: number }) => {
			const view = bytes().subarray(at, at + buffer.length)

			buffer.set(view)

			return view.length
		},
		write: (buffer: Uint8Array, { at }: { at: number }) => {
			const next = new Uint8Array(Math.max(bytes().length, at + buffer.length))

			next.set(bytes())
			next.set(buffer, at)
			shardFiles.set(name, next)

			return buffer.length
		},
		flush: () => undefined,
		close: () => undefined
	}
}

vi.mock("@/features/assistant/runtime/modelStore", () => ({
	VERIFIED_FILE: "verified.json",
	modelDirectory: () =>
		Promise.resolve({
			getFileHandle: (name: string, options?: { create?: boolean }) => {
				if (!shardFiles.has(name)) {
					if (options?.create !== true) {
						return Promise.reject(new DOMException("missing", "NotFoundError"))
					}

					shardFiles.set(name, new Uint8Array(0))
				}

				return Promise.resolve({
					getFile: () => Promise.resolve(new File([shardFiles.get(name) ?? new Uint8Array(0)], name)),
					createSyncAccessHandle: () => Promise.resolve(syncHandle(name))
				})
			},
			removeEntry: (name: string) => {
				shardFiles.delete(name)

				return Promise.resolve()
			}
		}),
	readVerified: () => Promise.resolve(verified)
}))

vi.mock("@/features/assistant/runtime/gguf", () => ({
	blobGgufRead: () => () => Promise.resolve(new Uint8Array()),
	readGgufSpecialTokens: () => Promise.resolve(["<|im_start|>", "<tool_call>"])
}))

await import("@/workers/assistant.worker")

const api = exposed.get("api") as AssistantWorkerApi

const CONTENTS = [new Uint8Array([1, 2, 3]), new Uint8Array([4, 5])]

function sha256(bytes: Uint8Array): string {
	return createHash("sha256").update(bytes).digest("hex")
}

const MANIFEST: AssistantModelManifest = {
	id: "m",
	label: "M",
	license: { name: "MIT", url: "" },
	source: { repo: "r", revision: "x" },
	shards: CONTENTS.map((bytes, index) => ({
		file: `m-0000${String(index + 1)}-of-00002.gguf`,
		size: bytes.length,
		sha256: sha256(bytes)
	}))
}

const OPTIONS = { backend: "cpu", threads: 4, contextTokens: 8192 } as const

describe("assistant worker API", () => {
	beforeEach(async () => {
		await api.unload()
		engines.length = 0
		shardFiles.clear()
		verified.clear()
		loadFailure.error = null
		loadOutcome.vocab = 10
		loadOutcome.log = null
		completion.run = null

		MANIFEST.shards.forEach((shard, index) => {
			shardFiles.set(shard.file, CONTENTS[index] ?? new Uint8Array(0))
			verified.set(shard.file, shard.sha256)
		})
	})

	it("loads the verified shards in order, never through the CDN compat build", async () => {
		const result = await api.load(MANIFEST, OPTIONS)
		const engine = engines[0]

		expect(result).toEqual({ specialTokens: ["<|im_start|>", "<tool_call>"] })
		expect(engine?.setCompat).toHaveBeenCalledWith(null)

		const [files, params] = (engine?.loadModel.mock.calls[0] ?? []) as [File[]?, Record<string, unknown>?]

		expect(files?.map(file => file.name)).toEqual(["m-00001-of-00002.gguf", "m-00002-of-00002.gguf"])
		expect(params).toMatchObject({
			n_ctx: 8192,
			n_threads: 4,
			n_gpu_layers: 0,
			jinja: true,
			reasoning: false,
			n_batch: 512,
			n_ubatch: 512
		})
	})

	it("offloads to the GPU on the WebGPU backend", async () => {
		await api.load(MANIFEST, { backend: "webgpu", threads: 4, contextTokens: 8192 })

		expect(engines[0]?.loadModel.mock.calls[0]?.[1]).toMatchObject({ n_gpu_layers: 999 })
	})

	it("trusts the verified record without reading the shards", async () => {
		const onVerifying = vi.fn()

		// Contents that would not match: a verified shard of the right size is never hashed again.
		shardFiles.set("m-00001-of-00002.gguf", new Uint8Array([9, 9, 9]))
		await api.load(MANIFEST, OPTIONS, onVerifying)

		expect(onVerifying).not.toHaveBeenCalled()
		expect(engines[0]?.loadModel).toHaveBeenCalled()
	})

	it("hashes a complete shard missing from the record once, and records it", async () => {
		const onVerifying = vi.fn()

		verified.delete("m-00002-of-00002.gguf")
		await api.load(MANIFEST, OPTIONS, onVerifying)

		expect(onVerifying).toHaveBeenCalled()
		expect(verified.get("m-00002-of-00002.gguf")).toBe(MANIFEST.shards[1]?.sha256)
		expect(JSON.parse(new TextDecoder().decode(shardFiles.get("verified.json")))).toEqual(Object.fromEntries(verified))
	})

	it("drops an unrecorded shard that does not match, so it is downloaded again", async () => {
		verified.delete("m-00002-of-00002.gguf")
		shardFiles.set("m-00002-of-00002.gguf", new Uint8Array([0, 0]))

		await expect(api.load(MANIFEST, OPTIONS)).rejects.toMatchObject({ name: "ModelLoadError:missing" })
		expect(shardFiles.has("m-00002-of-00002.gguf")).toBe(false)
		expect(engines).toHaveLength(0)
	})

	it("after a failed start, drops the shards damaged on disk so only they are downloaded again", async () => {
		const onVerifying = vi.fn()

		loadFailure.error = new Error("invalid tensor data")
		shardFiles.set("m-00001-of-00002.gguf", new Uint8Array([1, 2, 0]))

		await expect(api.load(MANIFEST, OPTIONS, onVerifying)).rejects.toMatchObject({ name: "ModelLoadError:missing" })
		expect(onVerifying).toHaveBeenCalled()
		expect(shardFiles.has("m-00001-of-00002.gguf")).toBe(false)
		expect(verified.has("m-00001-of-00002.gguf")).toBe(false)
		expect(shardFiles.has("m-00002-of-00002.gguf")).toBe(true)
		expect(verified.has("m-00002-of-00002.gguf")).toBe(true)
	})

	it("reports a failed start with intact shards as load-failed, and running out of memory without hashing", async () => {
		loadFailure.error = new Error("invalid tensor data")
		await expect(api.load(MANIFEST, OPTIONS)).rejects.toMatchObject({ name: "ModelLoadError:load-failed" })
		expect(shardFiles.size).toBe(2)

		const onVerifying = vi.fn()

		loadFailure.error = new Error("Out of memory")
		await expect(api.load(MANIFEST, OPTIONS, onVerifying)).rejects.toMatchObject({ name: "ModelLoadError:out-of-memory" })
		expect(onVerifying).not.toHaveBeenCalled()
	})

	it("downloads the missing shards, keeping no checkpoint once they are verified", async () => {
		vi.stubGlobal("navigator", { onLine: true })
		verified.clear()
		shardFiles.clear()
		// A partial first shard from an earlier session.
		shardFiles.set("m-00001-of-00002.gguf", new Uint8Array([1]))

		const fetchShard = vi.fn((url: string, init: RequestInit) => {
			const index = url.includes("00001") ? 0 : 1
			const range = init.headers === undefined ? null : new Headers(init.headers).get("Range")
			const from = range === null ? 0 : Number(/bytes=(\d+)-/.exec(range)?.[1] ?? 0)
			const body = (CONTENTS[index] ?? new Uint8Array(0)).slice(from)

			return Promise.resolve(new Response(body, { status: from > 0 ? 206 : 200, headers: { "Content-Length": String(body.length) } }))
		})

		vi.stubGlobal("fetch", fetchShard)

		const progress: number[] = []

		await api.download(MANIFEST, "https://cdn.example/", ({ receivedBytes }) => progress.push(receivedBytes))

		expect(progress.at(-1)).toBe(5)
		expect(Math.max(...progress)).toBe(5)
		expect(shardFiles.get("m-00001-of-00002.gguf")).toEqual(CONTENTS[0])
		expect(shardFiles.get("m-00002-of-00002.gguf")).toEqual(CONTENTS[1])
		expect([...shardFiles.keys()].filter(name => name.endsWith(".sha256-state"))).toEqual([])
		expect(Object.fromEntries(verified)).toEqual(Object.fromEntries(MANIFEST.shards.map(shard => [shard.file, shard.sha256])))
		expect(
			fetchShard.mock.calls.map(([url, init]) => [
				url.slice(-21),
				init.headers === undefined ? null : new Headers(init.headers).get("Range")
			])
		).toEqual(
			expect.arrayContaining([
				["m-00001-of-00002.gguf", "bytes=1-"],
				["m-00002-of-00002.gguf", null]
			])
		)
	})

	it("refuses to load a model that is not fully downloaded", async () => {
		shardFiles.delete("m-00002-of-00002.gguf")

		await expect(api.load(MANIFEST, OPTIONS)).rejects.toMatchObject({
			name: "ModelLoadError:missing"
		})
		expect(engines).toHaveLength(0)
	})

	it("generates with the model's own chat template, streaming progress, text and tool calls in order", async () => {
		await api.load(MANIFEST, OPTIONS)

		const updates: GenerationUpdate[] = []
		const result = await api.generate(
			1,
			{
				messages: [
					{ role: "system", content: "s" },
					{ role: "user", content: "find x" }
				],
				tools: [
					{
						name: "search_items",
						description: "Search",
						parameters: { type: "object", properties: { query: { type: "string" } }, additionalProperties: false }
					}
				],
				maxTokens: 384
			},
			update => updates.push(update)
		)

		expect(result).toEqual({ content: "Hi", toolCalls: [{ id: "t", name: "search_items", arguments: '{"query":"x"}' }] })
		// The text before the call is flushed ahead of it, not left to the throttle.
		expect(updates).toEqual([
			{ type: "progress", processed: 0, total: 32 },
			{ type: "progress", processed: 32, total: 32 },
			{ type: "text", text: "Hi" },
			{ type: "call", index: 0, name: "search_items" }
		])

		const options = engines[0]?.createChatCompletion.mock.calls[0]?.[0] as Record<string, unknown>

		expect(options).toMatchObject({
			max_tokens: 384,
			temperature: 0,
			tool_choice: "auto",
			stream: true,
			return_progress: true,
			chat_template_kwargs: { enable_thinking: false }
		})
		expect(options["grammar"]).toBeUndefined()
	})

	it("passes a grammar instead of tools for constrained decoding", async () => {
		await api.load(MANIFEST, OPTIONS)
		await api.generate(2, { messages: [{ role: "user", content: "x" }], grammar: 'root ::= "a"', maxTokens: 8 })

		const options = engines[0]?.createChatCompletion.mock.calls[0]?.[0] as Record<string, unknown>

		expect(options["grammar"]).toBe('root ::= "a"')
		expect(options["tools"]).toBeUndefined()
		// Nobody listens: no progress chunks to send.
		expect(options["return_progress"]).toBe(false)
	})

	it("treats a load the engine silently failed as failed, and its allocation errors as out of memory", async () => {
		loadOutcome.vocab = 0
		await expect(api.load(MANIFEST, OPTIONS)).rejects.toMatchObject({ name: "ModelLoadError:load-failed" })
		expect(engines[0]?.exit).toHaveBeenCalled()

		loadOutcome.log = "llama_kv_cache: failed to allocate buffer for kv cache"
		await expect(api.load(MANIFEST, OPTIONS)).rejects.toMatchObject({ name: "ModelLoadError:out-of-memory" })

		// Nothing stays loaded: a later load starts clean.
		loadOutcome.vocab = 10
		loadOutcome.log = null
		await expect(api.load(MANIFEST, OPTIONS)).resolves.toEqual({ specialTokens: ["<|im_start|>", "<tool_call>"] })
	})

	it("leaves the shards alone when the GPU failed to load them: the CPU path is tried next", async () => {
		const onVerifying = vi.fn()

		loadFailure.error = new Error("failed to load model")
		shardFiles.set("m-00001-of-00002.gguf", new Uint8Array([1, 2, 0]))

		await expect(api.load(MANIFEST, { ...OPTIONS, backend: "webgpu" }, onVerifying)).rejects.toMatchObject({
			name: "ModelLoadError:load-failed"
		})
		expect(onVerifying).not.toHaveBeenCalled()
		expect(shardFiles.has("m-00001-of-00002.gguf")).toBe(true)
	})

	it("releases an engine that crashed under a generation, and keeps one a stop ended", async () => {
		await api.load(MANIFEST, OPTIONS)

		const engine = engines[0]

		completion.run = ({ abortSignal }) =>
			new Promise((_resolve, reject) => {
				abortSignal.addEventListener("abort", () => {
					reject(Object.assign(new Error("Operation aborted"), { name: "AbortError" }))
				})
			})

		const stopped = api.generate(1, { messages: [], maxTokens: 1 })

		api.cancel(1)
		await expect(stopped).rejects.toMatchObject({ name: "AbortError" })
		expect(engine?.exit).not.toHaveBeenCalled()

		// wllama rejects with a RuntimeError once its wasm aborted, then answers every call with a bare string.
		const answer = "wllama has crashed, please reload the module" as unknown as Error

		for (const crash of [Object.assign(new Error("memory access out of bounds"), { name: "RuntimeError" }), answer]) {
			await api.unload()
			await api.load(MANIFEST, OPTIONS)
			completion.run = () => Promise.reject(crash)

			await expect(api.generate(2, { messages: [], maxTokens: 1 })).rejects.toMatchObject({ name: "ModelRuntimeError:crashed" })
			expect(engines.at(-1)?.exit).toHaveBeenCalled()
			await expect(api.generate(3, { messages: [], maxTokens: 1 })).rejects.toMatchObject({ name: "ModelLoadError:missing" })
		}
	})

	it("lets a cancelled generation wind down before the engine exits on unload", async () => {
		await api.load(MANIFEST, OPTIONS)

		const engine = engines[0]
		const order: string[] = []

		completion.run = ({ abortSignal }) =>
			new Promise((_resolve, reject) => {
				abortSignal.addEventListener("abort", () => {
					setTimeout(() => {
						order.push(engine?.exit.mock.calls.length === 0 ? "settled" : "settled after exit")
						reject(Object.assign(new Error("aborted"), { name: "AbortError" }))
					}, 20)
				})
			})
		const running = api.generate(1, { messages: [], maxTokens: 1 }).catch(() => undefined)

		await vi.waitFor(() => {
			expect(engine?.createChatCompletion).toHaveBeenCalled()
		})
		await api.unload()
		await running

		expect(order).toEqual(["settled"])
		expect(engine?.exit).toHaveBeenCalledTimes(1)
	})

	it("says when a generation stopped at its length limit", async () => {
		await api.load(MANIFEST, OPTIONS)
		completion.run = ({ onData }) => {
			onData({ choices: [{ index: 0, delta: { content: "cut" }, finish_reason: "length", logprobs: null }] })

			return Promise.resolve()
		}

		await expect(api.generate(1, { messages: [], maxTokens: 1 })).resolves.toEqual({ content: "cut", toolCalls: [], truncated: true })
	})

	it("unloads the engine", async () => {
		await api.load(MANIFEST, OPTIONS)

		const engine = engines[0]

		await api.unload()

		expect(engine?.exit).toHaveBeenCalled()
		await expect(api.generate(3, { messages: [], maxTokens: 1 })).rejects.toMatchObject({ name: "ModelLoadError:missing" })
	})
})
