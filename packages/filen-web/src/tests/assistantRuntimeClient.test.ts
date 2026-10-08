import { afterEach, beforeEach, describe, expect, it, vi, type Mock } from "vitest"
import type { ModelDownloadProgress } from "@/features/assistant/runtime/download"
import type { ModelInspection } from "@/features/assistant/runtime/modelStore"
import type { AssistantModelState } from "@/features/assistant/types"

// The page side of the runtime against a scripted worker: the states the panel sees across a download, an
// explicit load, a start that finds a damaged shard, and the idle unload that frees the worker again.
const { remote, inspections, releaseProxy, workers, locks, prune } = vi.hoisted(() => ({
	releaseProxy: Symbol("releaseProxy"),
	remote: {
		download: vi.fn<(manifest: unknown, baseUrl: string, onProgress: (progress: ModelDownloadProgress) => void) => Promise<void>>(),
		load: vi.fn<(manifest: unknown, options: unknown, onVerifying: () => void) => Promise<{ specialTokens: string[] }>>(),
		cancelDownload: vi.fn(),
		unload: vi.fn<() => Promise<void>>(),
		generate: vi.fn(() => Promise.resolve({ content: "", toolCalls: [] }))
	},
	// An Error fails that inspection.
	inspections: [] as (ModelInspection | Error)[],
	workers: [] as { terminate: Mock<() => void>; released: Mock<() => void> }[],
	// Model locks currently held by this "tab", how often it was granted one, and whether another tab holds it.
	locks: { held: 0, granted: 0, elsewhere: false },
	// The pruning of older model versions, when a test holds it open.
	prune: { pending: null as Promise<void> | null }
}))

vi.mock("@/workers/assistant.worker.ts?worker", () => ({
	default: class {
		readonly addEventListener = vi.fn()
		readonly terminate = vi.fn<() => void>()
		readonly released = vi.fn<() => void>()

		constructor() {
			workers.push(this)
		}
	}
}))

vi.mock("comlink", () => ({
	wrap: () => {
		const worker = workers.at(-1)

		return Object.assign({ [releaseProxy]: () => worker?.released() }, remote)
	},
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
	inspectModel: () => {
		const next = inspections.shift() ?? { presentBytes: 0, verifiedBytes: 0, complete: false }

		return next instanceof Error ? Promise.reject(next) : Promise.resolve(next)
	},
	deleteModels: () => prune.pending ?? Promise.resolve()
}))

const { AssistantRuntime, IDLE_UNLOAD_MS } = await import("@/features/assistant/runtime/client")
const { QWEN35_4B } = await import("@/features/assistant/runtime/models")
const { manifestBytes } = await import("@/features/assistant/runtime/manifest")

type Runtime = InstanceType<typeof AssistantRuntime>

const TOTAL = manifestBytes(QWEN35_4B)
const COMPLETE: ModelInspection = { presentBytes: TOTAL, verifiedBytes: TOTAL, complete: true }
const MINUTE = 60_000

function stubBrowser(): void {
	vi.stubGlobal("navigator", {
		locks: {
			query: () => Promise.resolve({ held: locks.elsewhere ? [{ name: "filen-assistant-model" }] : [] }),
			request: (
				_name: string,
				options: { ifAvailable?: boolean } | ((lock: unknown) => unknown),
				maybeCallback?: (lock: unknown) => unknown
			) => {
				const callback = typeof options === "function" ? options : maybeCallback
				const ifAvailable = typeof options === "object" && options.ifAvailable === true

				if (callback === undefined) {
					return Promise.reject(new Error("no callback"))
				}

				if (locks.elsewhere) {
					// Queued behind the other tab, which never lets go here.
					return ifAvailable ? Promise.resolve(callback(null)) : new Promise(() => undefined)
				}

				locks.held++
				locks.granted++
				void Promise.resolve(callback({})).then(() => {
					locks.held--
				})

				return Promise.resolve()
			}
		},
		storage: {
			persist: () => Promise.resolve(true),
			estimate: () => Promise.resolve({ quota: 100 * TOTAL, usage: 0 })
		}
	})
}

function settled(runtime: Runtime, kind: AssistantModelState["kind"]): Promise<void> {
	return vi.waitFor(() => {
		expect(runtime.getState().kind).toBe(kind)
	})
}

// Under fake timers: lets the promise chains run without moving the clock.
async function until(check: () => boolean): Promise<void> {
	for (let i = 0; i < 50 && !check(); i++) {
		await vi.advanceTimersByTimeAsync(0)
	}

	expect(check()).toBe(true)
}

beforeEach(() => {
	stubBrowser()
	inspections.length = 0
	workers.length = 0
	locks.held = 0
	locks.granted = 0
	locks.elsewhere = false
	prune.pending = null
	remote.download.mockReset()
	remote.load.mockReset()
	remote.unload.mockReset()
	remote.unload.mockResolvedValue(undefined)
})

afterEach(() => {
	vi.useRealTimers()
})

describe("AssistantRuntime download", () => {
	it("shows exact progress, never past the total, and stops at downloaded without loading", async () => {
		const states: AssistantModelState[] = []

		remote.download.mockImplementation((_manifest, _baseUrl, onProgress) => {
			onProgress({ receivedBytes: 1000, waiting: false })
			onProgress({ receivedBytes: 1000, waiting: true })
			// Never sent by the worker; the page holds the line anyway.
			onProgress({ receivedBytes: TOTAL + 50_000_000, waiting: false })

			return Promise.resolve()
		})

		const runtime = new AssistantRuntime()

		runtime.subscribe(state => states.push(state))
		runtime.setPanelOpen(true)
		runtime.download()
		await settled(runtime, "downloaded")
		await vi.waitFor(() => {
			expect(workers[0]?.terminate).toHaveBeenCalled()
		})

		const downloading = states.flatMap(state => (state.kind === "downloading" ? [state] : []))

		expect(downloading.map(state => state.receivedBytes)).toEqual([0, 1000, 1000, TOTAL])
		expect(downloading.map(state => state.waiting)).toEqual([false, false, true, false])
		expect(states.at(-1)).toMatchObject({ kind: "downloaded", idleUnloaded: false })
		expect(remote.load).not.toHaveBeenCalled()
		expect(locks.held).toBe(0)
	})

	it("reports a start that checks the files, and downloads again just the shards found damaged", async () => {
		const states: AssistantModelState[] = []

		inspections.push(
			COMPLETE,
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
		await runtime.init()
		await runtime.load()
		await vi.waitFor(() => {
			expect(remote.download).toHaveBeenCalledTimes(1)
		})
		await settled(runtime, "downloaded")

		expect(states).toContainEqual(expect.objectContaining({ kind: "loading", verifying: true }))
		expect(states.find(state => state.kind === "downloading")).toMatchObject({ receivedBytes: TOTAL - 500 })

		await runtime.load()

		expect(runtime.getState().kind).toBe("ready")
	})

	it("shows downloaded only once its worker and lock are gone, so a load right after starts a fresh worker", async () => {
		let finishPrune: () => void = () => undefined

		prune.pending = new Promise(resolve => {
			finishPrune = () => {
				resolve()
			}
		})
		remote.download.mockImplementation((_manifest, _baseUrl, onProgress) => {
			onProgress({ receivedBytes: TOTAL - 1, waiting: true })

			return Promise.resolve()
		})
		remote.load.mockResolvedValue({ specialTokens: [] })

		const runtime = new AssistantRuntime()

		runtime.download()
		await vi.waitFor(() => {
			expect(remote.download).toHaveBeenCalledTimes(1)
		})
		await new Promise(resolve => setTimeout(resolve, 10))

		// Older versions are still being removed: the bytes are in, so it reads as finishing, not waiting.
		expect(runtime.getState()).toMatchObject({ kind: "downloading", receivedBytes: TOTAL, waiting: false })

		await runtime.load()

		expect(remote.load).not.toHaveBeenCalled()

		finishPrune()
		await settled(runtime, "downloaded")

		expect(workers[0]?.terminate).toHaveBeenCalled()

		await runtime.load()

		expect(runtime.getState().kind).toBe("ready")
		expect(workers).toHaveLength(2)
		expect(workers[1]?.terminate).not.toHaveBeenCalled()
		await vi.waitFor(() => {
			expect(locks.held).toBe(1)
		})
	})

	it("downloads from a click made before the startup check has settled", async () => {
		const persist = vi.spyOn(navigator.storage, "persist")

		remote.download.mockResolvedValue(undefined)

		const runtime = new AssistantRuntime()

		runtime.retry()

		expect(persist).toHaveBeenCalledTimes(1)

		await settled(runtime, "downloaded")

		expect(remote.download).toHaveBeenCalledTimes(1)
	})

	it("retries a failed startup check by running the check again, not by downloading", async () => {
		inspections.push(new Error("storage unreadable"), COMPLETE)

		const runtime = new AssistantRuntime()

		await runtime.init()

		expect(runtime.getState()).toMatchObject({ kind: "error", error: "load-failed" })

		runtime.retry()
		await settled(runtime, "downloaded")

		expect(remote.download).not.toHaveBeenCalled()
		expect(remote.load).not.toHaveBeenCalled()
	})
})

describe("AssistantRuntime explicit load", () => {
	it("never loads on the startup check or a panel restored open after a reload", async () => {
		inspections.push(COMPLETE)

		const runtime = new AssistantRuntime()

		runtime.setPanelOpen(true)
		await runtime.init()
		runtime.setPanelOpen(false)
		runtime.setPanelOpen(true)
		await new Promise(resolve => setTimeout(resolve, 10))

		expect(runtime.getState()).toMatchObject({ kind: "downloaded", idleUnloaded: false })
		expect(remote.load).not.toHaveBeenCalled()
		expect(workers).toHaveLength(0)
		expect(locks.held).toBe(0)
	})

	it("loads on request, once for concurrent requests, holding the lock while loaded", async () => {
		inspections.push(COMPLETE)
		remote.load.mockResolvedValue({ specialTokens: ["<|im_end|>"] })

		const runtime = new AssistantRuntime()

		await runtime.init()
		await Promise.all([runtime.load(), runtime.load()])

		expect(runtime.getState().kind).toBe("ready")
		expect(runtime.specialTokens).toEqual(["<|im_end|>"])
		expect(remote.load).toHaveBeenCalledTimes(1)
		expect(workers).toHaveLength(1)
		expect(locks.held).toBe(1)
	})

	it("ignores a download asked for while the model is loaded, leaving it running", async () => {
		inspections.push(COMPLETE)
		remote.load.mockResolvedValue({ specialTokens: [] })

		const runtime = new AssistantRuntime()

		await runtime.init()
		await runtime.load()
		runtime.download()
		await new Promise(resolve => setTimeout(resolve, 10))

		expect(runtime.getState().kind).toBe("ready")
		expect(workers).toHaveLength(1)
		expect(workers[0]?.terminate).not.toHaveBeenCalled()
		expect(locks.held).toBe(1)
	})

	it("reads as in use elsewhere once the panel opens after another tab loaded the model", async () => {
		inspections.push(COMPLETE)

		const runtime = new AssistantRuntime()

		await runtime.init()

		expect(runtime.getState().kind).toBe("downloaded")

		locks.elsewhere = true
		runtime.setPanelOpen(true)
		await settled(runtime, "unavailable")

		expect(runtime.getState()).toMatchObject({ reason: "in-use-elsewhere" })
		expect(remote.load).not.toHaveBeenCalled()
	})

	it("re-checks the other tabs when the page with the panel open becomes visible again", async () => {
		const page = Object.assign(new EventTarget(), { visibilityState: "visible" })

		vi.stubGlobal("document", page)
		inspections.push(COMPLETE)

		const runtime = new AssistantRuntime()

		runtime.setPanelOpen(true)
		await settled(runtime, "downloaded")
		locks.elsewhere = true
		page.dispatchEvent(new Event("visibilitychange"))
		await settled(runtime, "unavailable")

		expect(runtime.getState()).toMatchObject({ reason: "in-use-elsewhere" })
	})

	it("loads nothing before the model is on disk", async () => {
		const runtime = new AssistantRuntime()

		await runtime.init()
		await runtime.load()

		expect(runtime.getState().kind).toBe("not-downloaded")
		expect(workers).toHaveLength(0)
		expect(locks.held).toBe(0)
	})
})

describe("AssistantRuntime idle unload", () => {
	async function loaded(): Promise<Runtime> {
		vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout", "Date"] })
		inspections.push(COMPLETE)
		remote.load.mockResolvedValue({ specialTokens: [] })

		const runtime = new AssistantRuntime()

		void runtime.init()
		await until(() => runtime.getState().kind === "downloaded")
		void runtime.load()
		await until(() => runtime.getState().kind === "ready")

		return runtime
	}

	it("unloads after the idle period with the panel open, freeing the worker and the lock", async () => {
		const runtime = await loaded()
		const worker = workers[0]

		runtime.setPanelOpen(true)
		await vi.advanceTimersByTimeAsync(IDLE_UNLOAD_MS - 1)

		expect(runtime.getState().kind).toBe("ready")

		await vi.advanceTimersByTimeAsync(1)
		await until(() => locks.held === 0)

		expect(runtime.getState()).toMatchObject({ kind: "downloaded", idleUnloaded: true })
		expect(remote.unload).toHaveBeenCalledTimes(1)
		expect(worker?.released).toHaveBeenCalled()
		expect(worker?.terminate).toHaveBeenCalled()
		expect(runtime.specialTokens).toEqual([])
		await expect(runtime.generate({ messages: [], maxTokens: 1 }, { signal: new AbortController().signal })).rejects.toThrow(
			"not loaded"
		)
	})

	it("terminates the worker even when the engine never answers the unload", async () => {
		const runtime = await loaded()

		remote.unload.mockReturnValue(new Promise(() => undefined))
		await vi.advanceTimersByTimeAsync(IDLE_UNLOAD_MS)

		expect(runtime.getState().kind).toBe("downloaded")

		await vi.advanceTimersByTimeAsync(1000)
		await until(() => locks.held === 0)

		expect(workers[0]?.terminate).toHaveBeenCalled()
	})

	it("postpones the unload while the user keeps typing", async () => {
		const runtime = await loaded()

		await vi.advanceTimersByTimeAsync(9 * MINUTE)
		runtime.touch()
		await vi.advanceTimersByTimeAsync(9 * MINUTE)
		runtime.touch()
		await vi.advanceTimersByTimeAsync(9 * MINUTE)

		expect(runtime.getState().kind).toBe("ready")

		await vi.advanceTimersByTimeAsync(MINUTE)

		expect(runtime.getState().kind).toBe("downloaded")
	})

	it("never unloads while a turn, a waiting card or a confirmed action holds it, and counts from the release", async () => {
		const runtime = await loaded()
		const turn = runtime.hold()
		const action = runtime.hold()

		await vi.advanceTimersByTimeAsync(3 * IDLE_UNLOAD_MS)
		turn()
		turn()
		await vi.advanceTimersByTimeAsync(3 * IDLE_UNLOAD_MS)

		expect(runtime.getState().kind).toBe("ready")

		action()
		await vi.advanceTimersByTimeAsync(IDLE_UNLOAD_MS - 1)

		expect(runtime.getState().kind).toBe("ready")

		await vi.advanceTimersByTimeAsync(1)

		expect(runtime.getState().kind).toBe("downloaded")
	})

	it("has no countdown during a download, so a long one is never cut off", async () => {
		vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout", "Date"] })
		remote.download.mockReturnValue(new Promise(() => undefined))

		const runtime = new AssistantRuntime()

		runtime.download()
		await until(() => runtime.getState().kind === "downloading")
		await vi.advanceTimersByTimeAsync(3 * IDLE_UNLOAD_MS)

		expect(runtime.getState().kind).toBe("downloading")
		expect(workers[0]?.terminate).not.toHaveBeenCalled()
		expect(remote.unload).not.toHaveBeenCalled()
	})

	it("checks the deadline when the page becomes visible again, as after a sleep that held the timer back", async () => {
		const page = new EventTarget()

		vi.stubGlobal("document", page)

		const runtime = await loaded()

		vi.setSystemTime(Date.now() + IDLE_UNLOAD_MS + MINUTE)
		page.dispatchEvent(new Event("visibilitychange"))
		await until(() => runtime.getState().kind === "downloaded")

		expect(workers[0]?.terminate).toHaveBeenCalled()

		// The listener left with the model.
		const remove = vi.spyOn(page, "removeEventListener")

		void runtime.load()
		await until(() => runtime.getState().kind === "ready")
		await runtime.unload()

		expect(remove).toHaveBeenCalledWith("visibilitychange", expect.any(Function))
	})

	it("loads a fresh worker only after an unload in flight has let the old one go", async () => {
		const runtime = await loaded()
		let finishUnload: () => void = () => undefined

		remote.unload.mockReturnValue(
			new Promise(resolve => {
				finishUnload = () => {
					resolve()
				}
			})
		)
		await vi.advanceTimersByTimeAsync(IDLE_UNLOAD_MS)
		void runtime.load()
		await vi.advanceTimersByTimeAsync(0)

		expect(workers).toHaveLength(1)

		finishUnload()
		await until(() => runtime.getState().kind === "ready")

		expect(workers[0]?.terminate).toHaveBeenCalled()
		expect(workers).toHaveLength(2)
		// The lock passed straight from the old worker to the new one: never let go, so never taken again.
		expect(locks.held).toBe(1)
		expect(locks.granted).toBe(1)
	})

	it("an explicit unload does not read as an idle one", async () => {
		const runtime = await loaded()

		await runtime.unload()

		expect(runtime.getState()).toMatchObject({ kind: "downloaded", idleUnloaded: false })
	})
})
