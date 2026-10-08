import * as Comlink from "comlink"
import AssistantWorker from "@/workers/assistant.worker.ts?worker"
import type { AssistantWorkerApi } from "@/workers/assistant.worker"
import {
	AssistantAbortError,
	type AssistantGenerateRequest,
	type AssistantGenerateResult,
	type AssistantModel,
	type AssistantSignal
} from "@filen/shared/assistant"
import type { AssistantModelInfo, AssistantModelState } from "@/features/assistant/types"
import { DEFAULT_MODEL_BASE_URL, manifestBytes, parseManifest, type AssistantModelManifest } from "@/features/assistant/runtime/manifest"
import { QWEN35_4B } from "@/features/assistant/runtime/models"
import { probeDevice, recommendRuntime, storageVerdict, type RuntimeRecommendation } from "@/features/assistant/runtime/deviceCheck"
import { deleteModels, inspectModel } from "@/features/assistant/runtime/modelStore"
import type { ModelDownloadProgress } from "@/features/assistant/runtime/download"
import { recoverIfNewerBuild } from "@/lib/appUpdate"
import { log } from "@/lib/log"

// The page side of the on-device model: decides the model state, owns the worker (spawned only to
// download or to run the model, terminated otherwise, so an idle app holds none of its memory), and holds
// a Web Lock while it does either, so only one tab ever downloads or loads the model.

export const ASSISTANT_CONTEXT_TOKENS = 8192

const LOCK = "filen-assistant-model"

// The panel closed this long ago: the model leaves memory.
const IDLE_UNLOAD_MS = 5 * 60_000

type Settled = AssistantModelState["kind"]

interface WorkerHandle {
	worker: Worker
	remote: Comlink.Remote<AssistantWorkerApi>
}

// Read through a call: TypeScript would keep `aborted` narrowed across the awaits in between.
function isAborted(signal: AssistantSignal): boolean {
	return signal.aborted
}

function errorName(e: unknown): string {
	return e instanceof Error ? e.name : ""
}

// A dev build without a configured host serves the GGUF named by FILEN_ASSISTANT_DEV_GGUF (vite/assistant-model.ts).
async function resolveModel(): Promise<{ manifest: AssistantModelManifest; baseUrl: string }> {
	const configured = import.meta.env.VITE_ASSISTANT_MODEL_BASE_URL

	if (import.meta.env.DEV && (configured === undefined || configured === "")) {
		try {
			const response = await fetch("/__assistant-model/manifest.json")

			// Without the plugin the dev server answers with the app's index.html.
			if (response.ok && (response.headers.get("Content-Type") ?? "").includes("application/json")) {
				return {
					manifest: parseManifest(await response.json()),
					baseUrl: new URL("/__assistant-model/", location.origin).href
				}
			}
		} catch (e) {
			log.warn("assistant", "dev model manifest unavailable", e)
		}
	}

	return { manifest: QWEN35_4B, baseUrl: configured === undefined || configured === "" ? DEFAULT_MODEL_BASE_URL : configured }
}

// Also the AssistantModel the loop drives: generate() runs in the worker while the state is "ready".
export class AssistantRuntime implements AssistantModel {
	readonly contextTokens = ASSISTANT_CONTEXT_TOKENS

	private state: AssistantModelState
	private readonly listeners = new Set<(state: AssistantModelState) => void>()
	private manifest: AssistantModelManifest = QWEN35_4B
	private baseUrl = DEFAULT_MODEL_BASE_URL
	private recommendation: RuntimeRecommendation | null = null
	private ready: Promise<void> | null = null
	private handle: WorkerHandle | null = null
	private releaseLock: (() => void) | null = null
	private watchingLock = false
	private tokenList: readonly string[] = []
	private panelOpen = false
	private idleTimer: ReturnType<typeof setTimeout> | undefined
	private busy = 0
	private downloading: Promise<void> | null = null
	private loading: Promise<void> | null = null
	private redownload = false
	private nextGeneration = 1

	constructor() {
		this.state = { kind: "not-downloaded", model: this.info() }
	}

	get specialTokens(): readonly string[] {
		return this.tokenList
	}

	get model(): AssistantModel {
		return this
	}

	getState(): AssistantModelState {
		return this.state
	}

	subscribe(listener: (state: AssistantModelState) => void): () => void {
		this.listeners.add(listener)

		return () => {
			this.listeners.delete(listener)
		}
	}

	private info(): AssistantModelInfo {
		return { id: this.manifest.id, name: this.manifest.label, sizeBytes: manifestBytes(this.manifest) }
	}

	private setState(state: AssistantModelState): void {
		this.state = state

		for (const listener of this.listeners) {
			listener(state)
		}
	}

	private settle(kind: Exclude<Settled, "unavailable" | "error" | "downloading" | "loading">): void {
		this.setState({ kind, model: this.info() })
	}

	// The cheap startup check: device probe, OPFS sizes and the verified record; no shard content is read.
	init(): Promise<void> {
		this.ready ??= this.refresh().catch((e: unknown) => {
			log.error("assistant", "model check failed", e)
			this.setState({ kind: "error", error: "load-failed", model: this.info() })
		})

		return this.ready
	}

	private async refresh(): Promise<void> {
		const resolved = await resolveModel()

		this.manifest = resolved.manifest
		this.baseUrl = resolved.baseUrl
		// At startup only whether the model can run at all, which threads alone decide on every engine this
		// app runs in (it is cross-origin isolated); the GPU is asked only when there are none, or at load.
		if (this.recommendation === null) {
			const withoutGpu = recommendRuntime(await probeDevice({ gpu: false }))

			this.recommendation = withoutGpu.supported ? withoutGpu : recommendRuntime(await probeDevice({ gpu: true }))
		}

		if (!this.recommendation.supported) {
			this.setState({ kind: "unavailable", reason: this.recommendation.reason, model: this.info() })

			return
		}

		if (this.releaseLock === null && (await this.lockHeldElsewhere())) {
			this.inUseElsewhere()

			return
		}

		const inspection = await inspectModel(this.manifest)

		this.settle(inspection.complete ? "downloaded" : "not-downloaded")

		if (inspection.complete && this.panelOpen) {
			void this.load()
		}
	}

	private async lockHeldElsewhere(): Promise<boolean> {
		const { held = [] } = await navigator.locks.query()

		return held.some(lock => lock.name === LOCK)
	}

	private tryLock(): Promise<boolean> {
		if (this.releaseLock !== null) {
			return Promise.resolve(true)
		}

		return new Promise((resolve, reject) => {
			navigator.locks
				.request(LOCK, { ifAvailable: true }, lock => {
					if (lock === null) {
						resolve(false)

						return
					}

					resolve(true)

					return new Promise<void>(release => {
						this.releaseLock = release
					})
				})
				.catch(reject)
		})
	}

	private unlock(): void {
		this.releaseLock?.()
		this.releaseLock = null
	}

	// Another tab downloads or runs the model. Waits in the lock's queue and re-checks once that tab lets go.
	private inUseElsewhere(): void {
		this.setState({ kind: "unavailable", reason: "in-use-elsewhere", model: this.info() })

		if (this.watchingLock) {
			return
		}

		this.watchingLock = true
		void navigator.locks
			.request(LOCK, () => undefined)
			.then(() => {
				this.watchingLock = false
				this.ready = null

				return this.init()
			})
	}

	private worker(): WorkerHandle {
		if (this.handle === null) {
			const worker = new AssistantWorker()

			// A deploy that changed this worker removes its old file (appUpdate.ts).
			worker.addEventListener("error", () => {
				recoverIfNewerBuild()
			})
			this.handle = { worker, remote: Comlink.wrap<AssistantWorkerApi>(worker) }
		}

		return this.handle
	}

	// Everything the worker held (engine, wasm memory, GPU buffers) goes with it.
	private stopWorker(): void {
		const handle = this.handle

		this.handle = null

		if (handle !== null) {
			handle.remote[Comlink.releaseProxy]()
			handle.worker.terminate()
		}
	}

	setPanelOpen(open: boolean): void {
		this.panelOpen = open
		clearTimeout(this.idleTimer)
		this.idleTimer = undefined

		if (open) {
			void this.init().then(() => {
				if (this.panelOpen && this.state.kind === "downloaded") {
					void this.load()
				}
			})

			return
		}

		this.scheduleUnload()
	}

	// Turns and confirmed actions keep the model; the idle timer starts once the last one ends.
	hold(): () => void {
		this.busy++
		clearTimeout(this.idleTimer)

		let released = false

		return () => {
			if (released) {
				return
			}

			released = true
			this.busy--
			this.scheduleUnload()
		}
	}

	private scheduleUnload(): void {
		clearTimeout(this.idleTimer)
		this.idleTimer = undefined

		if (this.panelOpen || this.busy > 0 || (this.state.kind !== "ready" && this.state.kind !== "loading")) {
			return
		}

		this.idleTimer = setTimeout(() => {
			this.idleTimer = undefined

			if (!this.panelOpen && this.busy === 0) {
				void this.unload()
			}
		}, IDLE_UNLOAD_MS)
	}

	// From a click: the opt-in download, then the load. From "error" it retries whichever step failed.
	download(): void {
		// Asked first, while the click's user activation lasts (Firefox prompts for it). Without it the
		// browser may evict the 2.7 GB under storage pressure; with or without, the download proceeds.
		void navigator.storage.persist().catch(() => false)

		if (this.downloading !== null || this.loading !== null) {
			return
		}

		this.downloading = this.runDownload().finally(() => {
			this.downloading = null
		})
	}

	private async runDownload(): Promise<void> {
		await this.init()

		if (this.recommendation?.supported !== true) {
			return
		}

		if (!(await this.tryLock())) {
			this.inUseElsewhere()

			return
		}

		try {
			const before = await inspectModel(this.manifest)

			if (!before.complete) {
				const { quota = 0, usage = 0 } = await navigator.storage.estimate()
				const verdict = storageVerdict({ quota, usage, neededBytes: manifestBytes(this.manifest) - before.presentBytes })

				if (!verdict.ok) {
					this.setState({ kind: "unavailable", reason: "insufficient-storage", model: this.info() })
					this.unlock()

					return
				}

				const sizeBytes = manifestBytes(this.manifest)

				this.setState({ kind: "downloading", model: this.info(), receivedBytes: before.presentBytes, waiting: false })
				await this.worker().remote.download(
					this.manifest,
					this.baseUrl,
					Comlink.proxy(({ receivedBytes, waiting }: ModelDownloadProgress) => {
						if (this.state.kind === "downloading") {
							this.setState({
								kind: "downloading",
								model: this.info(),
								receivedBytes: Math.min(sizeBytes, Math.max(0, receivedBytes)),
								waiting
							})
						}
					})
				)
				this.settle("downloaded")
				// An older model version is dead weight once its successor is complete.
				await deleteModels(this.manifest.id).catch((e: unknown) => {
					log.warn("assistant", "pruning old models failed", e)
				})
			}
		} catch (e) {
			this.stopWorker()
			this.unlock()

			const name = errorName(e)

			if (name === "AbortError") {
				this.settle("not-downloaded")
			} else {
				log.error("assistant", "model download failed", e)
				this.setState({
					kind: "error",
					error: name === "ModelDownloadError:storage-full" || name === "QuotaExceededError" ? "storage-full" : "download-failed",
					model: this.info()
				})
			}

			return
		}

		if (this.state.kind !== "downloaded") {
			this.settle("downloaded")
		}

		if (this.panelOpen) {
			await this.load()
		} else {
			this.stopWorker()
			this.unlock()
		}
	}

	cancelDownload(): void {
		void this.handle?.remote.cancelDownload()
	}

	private load(): Promise<void> {
		this.loading ??= this.runLoad().finally(() => {
			this.loading = null

			if (this.redownload) {
				this.redownload = false
				this.download()
			}
		})

		return this.loading
	}

	private async runLoad(): Promise<void> {
		if (this.recommendation?.supported !== true || this.state.kind === "ready") {
			return
		}

		// The backend: WebGPU when the adapter is up to it.
		const recommendation = recommendRuntime(await probeDevice({ gpu: true }))

		if (!recommendation.supported) {
			this.setState({ kind: "unavailable", reason: recommendation.reason, model: this.info() })

			return
		}

		if (!(await this.tryLock())) {
			this.inUseElsewhere()

			return
		}

		this.setState({ kind: "loading", model: this.info(), verifying: false })

		try {
			const { specialTokens } = await this.worker().remote.load(
				this.manifest,
				{
					backend: recommendation.backend,
					threads: recommendation.threads,
					contextTokens: ASSISTANT_CONTEXT_TOKENS
				},
				Comlink.proxy(() => {
					if (this.state.kind === "loading") {
						this.setState({ kind: "loading", model: this.info(), verifying: true })
					}
				})
			)

			this.tokenList = specialTokens
			this.settle("ready")
			this.scheduleUnload()
		} catch (e) {
			this.stopWorker()
			this.unlock()
			log.error("assistant", "model load failed", e)

			const name = errorName(e)

			if (name === "ModelLoadError:missing") {
				const inspection = await inspectModel(this.manifest)

				this.settle(inspection.complete ? "downloaded" : "not-downloaded")
				// Some shards are gone or were damaged, the rest are still verified: the user opted in already, so
				// only the missing ones are fetched again, right after this load has wound down.
				this.redownload = !inspection.complete && inspection.verifiedBytes > 0
			} else {
				this.setState({
					kind: "error",
					error: name === "ModelLoadError:out-of-memory" ? "out-of-memory" : "load-failed",
					model: this.info()
				})
			}
		}
	}

	// Retry from the error state: the step that failed again.
	retry(): void {
		if (this.state.kind === "error" && this.state.error !== "download-failed" && this.state.error !== "storage-full") {
			void this.load()

			return
		}

		this.download()
	}

	async unload(): Promise<void> {
		if (this.state.kind !== "ready" && this.state.kind !== "loading") {
			return
		}

		await this.loading
		await this.handle?.remote.unload().catch((e: unknown) => {
			log.warn("assistant", "model unload failed", e)
		})
		this.stopWorker()
		this.unlock()
		this.tokenList = []
		this.settle("downloaded")
	}

	// Cancels a download, unloads, and removes every downloaded model from OPFS.
	async deleteModel(): Promise<void> {
		this.cancelDownload()
		await this.downloading
		await this.unload()

		if (!(await this.tryLock())) {
			this.inUseElsewhere()

			throw new Error("The model is in use in another tab")
		}

		try {
			await deleteModels()
		} finally {
			this.stopWorker()
			this.unlock()
		}

		this.settle("not-downloaded")
	}

	async generate(
		request: AssistantGenerateRequest,
		{ signal, onText }: { signal: AssistantSignal; onText?: (text: string) => void }
	): Promise<AssistantGenerateResult> {
		const handle = this.handle

		if (handle === null || this.state.kind !== "ready") {
			throw new Error("The model is not loaded")
		}

		if (signal.aborted) {
			throw new AssistantAbortError()
		}

		const id = this.nextGeneration++
		const cancel = () => {
			void handle.remote.cancel(id)
		}

		signal.addEventListener("abort", cancel)

		try {
			return await handle.remote.generate(id, request, onText === undefined ? undefined : Comlink.proxy(onText))
		} catch (e) {
			throw errorName(e) === "AbortError" || isAborted(signal) ? new AssistantAbortError() : e
		} finally {
			signal.removeEventListener("abort", cancel)
		}
	}
}
