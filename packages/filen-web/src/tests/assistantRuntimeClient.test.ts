import { beforeEach, describe, expect, it, vi } from "vitest"
import type { ModelDownloadProgress } from "@/features/assistant/runtime/download"
import type { ModelInspection } from "@/features/assistant/runtime/modelStore"
import type { AssistantModelState } from "@/features/assistant/types"

// The page side of the runtime against a scripted worker: the states the panel sees across a download, a
// start that finds a damaged shard, and the download of just that shard which follows.
const { remote, inspections, releaseProxy } = vi.hoisted(() => ({
	releaseProxy: Symbol("releaseProxy"),
	remote: {
		download: vi.fn<(manifest: unknown, baseUrl: string, onProgress: (progress: ModelDownloadProgress) => void) => Promise<void>>(),
		load: vi.fn<(manifest: unknown, options: unknown, onVerifying: () => void) => Promise<{ specialTokens: string[] }>>(),
		cancelDownload: vi.fn(),
		unload: vi.fn(() => Promise.resolve())
	},
	inspections: [] as ModelInspection[]
}))

vi.mock("@/workers/assistant.worker.ts?worker", () => ({
	default: class {
		readonly addEventListener = vi.fn()
		readonly terminate = vi.fn()
	}
}))

vi.mock("comlink", () => ({
	wrap: () => Object.assign({ [releaseProxy]: () => undefined }, remote),
	proxy: <T>(value: T) => value,
	releaseProxy
}))

vi.mock("@/lib/appUpdate", () => ({ recoverIfNewerBuild: () => undefined }))

vi.mock("@/features/assistant/runtime/deviceCheck", () => ({
	probeDevice: () => Promise.resolve({}),
	recommendRuntime: () => ({ supported: true, backend: "webgpu", threads: 4, estimatedSpeed: "fast" }),
	storageVerdict: () => ({ ok: true })
}))

vi.mock("@/features/assistant/runtime/modelStore", () => ({
	inspectModel: () => Promise.resolve(inspections.shift() ?? { presentBytes: 0, verifiedBytes: 0, complete: false }),
	deleteModels: () => Promise.resolve()
}))

const { AssistantRuntime } = await import("@/features/assistant/runtime/client")
const { QWEN35_4B } = await import("@/features/assistant/runtime/models")
const { manifestBytes } = await import("@/features/assistant/runtime/manifest")

const TOTAL = manifestBytes(QWEN35_4B)

function stubBrowser(): void {
	vi.stubGlobal("navigator", {
		locks: {
			query: () => Promise.resolve({ held: [] }),
			request: (_name: string, _options: unknown, callback: (lock: unknown) => Promise<void>) => {
				void callback({})

				return Promise.resolve()
			}
		},
		storage: {
			persist: () => Promise.resolve(true),
			estimate: () => Promise.resolve({ quota: 100 * TOTAL, usage: 0 })
		}
	})
}

function settled(runtime: InstanceType<typeof AssistantRuntime>, kind: AssistantModelState["kind"]): Promise<void> {
	return vi.waitFor(() => {
		expect(runtime.getState().kind).toBe(kind)
	})
}

describe("AssistantRuntime", () => {
	beforeEach(() => {
		stubBrowser()
		inspections.length = 0
		remote.download.mockReset()
		remote.load.mockReset()
	})

	it("shows exact progress, never past the total, and leaves the download state once the bytes are in", async () => {
		const states: AssistantModelState[] = []

		remote.download.mockImplementation((_manifest, _baseUrl, onProgress) => {
			onProgress({ receivedBytes: 1000, waiting: false })
			onProgress({ receivedBytes: 1000, waiting: true })
			// Never sent by the worker; the page holds the line anyway.
			onProgress({ receivedBytes: TOTAL + 50_000_000, waiting: false })

			return Promise.resolve()
		})
		remote.load.mockResolvedValue({ specialTokens: [] })

		const runtime = new AssistantRuntime()

		runtime.subscribe(state => states.push(state))
		runtime.setPanelOpen(true)
		runtime.download()
		await settled(runtime, "ready")

		const downloading = states.flatMap(state => (state.kind === "downloading" ? [state] : []))

		expect(downloading.map(state => state.receivedBytes)).toEqual([0, 1000, 1000, TOTAL])
		expect(downloading.map(state => state.waiting)).toEqual([false, false, true, false])
		expect(states.map(state => state.kind).slice(-3)).toEqual(["downloaded", "loading", "ready"])
	})

	it("reports a start that checks the files, and downloads again just the shards found damaged", async () => {
		const states: AssistantModelState[] = []

		inspections.push(
			// init: complete on disk
			{ presentBytes: TOTAL, verifiedBytes: TOTAL, complete: true },
			// after the failed start: one shard dropped, the rest still verified
			{ presentBytes: TOTAL - 500, verifiedBytes: TOTAL - 500, complete: false },
			// the download that follows
			{ presentBytes: TOTAL - 500, verifiedBytes: TOTAL - 500, complete: false }
		)
		remote.load
			.mockImplementationOnce((_manifest, _options, onVerifying) => {
				onVerifying()

				return Promise.reject(Object.assign(new Error("damaged"), { name: "ModelLoadError:missing" }))
			})
			.mockResolvedValue({ specialTokens: [] })
		remote.download.mockImplementation((_manifest, _baseUrl, onProgress) => {
			onProgress({ receivedBytes: TOTAL, waiting: false })

			return Promise.resolve()
		})

		const runtime = new AssistantRuntime()

		runtime.subscribe(state => states.push(state))
		runtime.setPanelOpen(true)
		await vi.waitFor(() => {
			expect(remote.load).toHaveBeenCalledTimes(2)
		})
		await settled(runtime, "ready")

		expect(states).toContainEqual(expect.objectContaining({ kind: "loading", verifying: true }))
		expect(remote.download).toHaveBeenCalledTimes(1)
		expect(states.find(state => state.kind === "downloading")).toMatchObject({ receivedBytes: TOTAL - 500 })
	})
})
