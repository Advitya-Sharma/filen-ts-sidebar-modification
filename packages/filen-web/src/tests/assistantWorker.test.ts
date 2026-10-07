import { beforeEach, describe, expect, it, vi } from "vitest"
import type { AssistantWorkerApi } from "@/workers/assistant.worker"
import type { AssistantModelManifest } from "@/features/assistant/runtime/manifest"

// The worker module exposes its API through Comlink at import; the engine, its wasm URL and OPFS are
// replaced so the contract between the page and the engine can be checked under node.
const { exposed, engines, shardFiles, verified } = vi.hoisted(() => ({
	exposed: new Map<"api", unknown>(),
	engines: [] as {
		loadModel: ReturnType<typeof vi.fn>
		createChatCompletion: ReturnType<typeof vi.fn>
		setCompat: ReturnType<typeof vi.fn>
		exit: ReturnType<typeof vi.fn>
	}[],
	shardFiles: new Map<string, File>(),
	verified: new Map<string, string>()
}))

vi.mock("comlink", () => ({
	expose: (api: unknown) => {
		exposed.set("api", api)
	},
	releaseProxy: Symbol("releaseProxy")
}))

vi.mock("@wllama/wllama/esm/wasm/wllama.wasm?url", () => ({ default: "/assets/wllama-hash.wasm" }))

vi.mock("@wllama/wllama/esm/index.js", () => ({
	Wllama: class {
		readonly loadModel = vi.fn(() => Promise.resolve())
		readonly setCompat = vi.fn()
		readonly exit = vi.fn(() => Promise.resolve())
		readonly createChatCompletion = vi.fn((options: { onData: (chunk: unknown) => void; abortSignal: AbortSignal }): Promise<void> => {
			const chunk = (delta: unknown) => ({ choices: [{ index: 0, delta, finish_reason: null, logprobs: null }] })

			options.onData(chunk({ content: "Hi" }))
			options.onData(chunk({ tool_calls: [{ index: 0, id: "t", function: { name: "search_items", arguments: '{"query":' } }] }))
			options.onData(chunk({ tool_calls: [{ index: 0, function: { arguments: '"x"}' } }] }))

			return Promise.resolve()
		})

		constructor() {
			engines.push(this)
		}
	}
}))

vi.mock("@/features/assistant/runtime/modelStore", () => ({
	modelDirectory: () =>
		Promise.resolve({
			getFileHandle: (name: string) => {
				const file = shardFiles.get(name)

				return file === undefined
					? Promise.reject(new DOMException("missing", "NotFoundError"))
					: Promise.resolve({ getFile: () => Promise.resolve(file) })
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

const MANIFEST: AssistantModelManifest = {
	id: "m",
	label: "M",
	license: { name: "MIT", url: "" },
	source: { repo: "r", revision: "x" },
	shards: [
		{ file: "m-00001-of-00002.gguf", size: 3, sha256: "a".repeat(64) },
		{ file: "m-00002-of-00002.gguf", size: 2, sha256: "b".repeat(64) }
	]
}

describe("assistant worker API", () => {
	beforeEach(async () => {
		await api.unload()
		engines.length = 0
		shardFiles.clear()
		verified.clear()

		for (const shard of MANIFEST.shards) {
			shardFiles.set(shard.file, new File([new Uint8Array(shard.size)], shard.file))
			verified.set(shard.file, shard.sha256)
		}
	})

	it("loads the verified shards in order, never through the CDN compat build", async () => {
		const result = await api.load(MANIFEST, { backend: "cpu", threads: 4, contextTokens: 8192 })
		const engine = engines[0]

		expect(result).toEqual({ specialTokens: ["<|im_start|>", "<tool_call>"] })
		expect(engine?.setCompat).toHaveBeenCalledWith(null)

		const [files, params] = (engine?.loadModel.mock.calls[0] ?? []) as [File[]?, Record<string, unknown>?]

		expect(files?.map(file => file.name)).toEqual(["m-00001-of-00002.gguf", "m-00002-of-00002.gguf"])
		expect(params).toMatchObject({ n_ctx: 8192, n_threads: 4, n_gpu_layers: 0, jinja: true, reasoning: false })
	})

	it("offloads to the GPU on the WebGPU backend", async () => {
		await api.load(MANIFEST, { backend: "webgpu", threads: 4, contextTokens: 8192 })

		expect(engines[0]?.loadModel.mock.calls[0]?.[1]).toMatchObject({ n_gpu_layers: 999 })
	})

	it("refuses to load a model that is not fully downloaded", async () => {
		shardFiles.delete("m-00002-of-00002.gguf")

		await expect(api.load(MANIFEST, { backend: "cpu", threads: 4, contextTokens: 8192 })).rejects.toMatchObject({
			name: "ModelLoadError:missing"
		})
		expect(engines).toHaveLength(0)
	})

	it("generates with the model's own chat template, streaming text and collecting tool calls", async () => {
		await api.load(MANIFEST, { backend: "cpu", threads: 4, contextTokens: 8192 })

		const texts: string[] = []
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
			text => texts.push(text)
		)

		expect(result).toEqual({ content: "Hi", toolCalls: [{ id: "t", name: "search_items", arguments: '{"query":"x"}' }] })
		expect(texts.at(-1)).toBe("Hi")

		const options = engines[0]?.createChatCompletion.mock.calls[0]?.[0] as Record<string, unknown>

		expect(options).toMatchObject({
			max_tokens: 384,
			temperature: 0,
			tool_choice: "auto",
			stream: true,
			chat_template_kwargs: { enable_thinking: false }
		})
		expect(options["grammar"]).toBeUndefined()
	})

	it("passes a grammar instead of tools for constrained decoding", async () => {
		await api.load(MANIFEST, { backend: "cpu", threads: 4, contextTokens: 8192 })
		await api.generate(2, { messages: [{ role: "user", content: "x" }], grammar: 'root ::= "a"', maxTokens: 8 })

		const options = engines[0]?.createChatCompletion.mock.calls[0]?.[0] as Record<string, unknown>

		expect(options["grammar"]).toBe('root ::= "a"')
		expect(options["tools"]).toBeUndefined()
	})

	it("unloads the engine", async () => {
		await api.load(MANIFEST, { backend: "cpu", threads: 4, contextTokens: 8192 })

		const engine = engines[0]

		await api.unload()

		expect(engine?.exit).toHaveBeenCalled()
		await expect(api.generate(3, { messages: [], maxTokens: 1 })).rejects.toMatchObject({ name: "ModelLoadError:missing" })
	})
})
