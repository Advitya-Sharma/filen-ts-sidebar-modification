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
// a Web Lock while it does either, so only one tab ever downloads or loads the model. The model loads only
// on an explicit request (a sent message or the Load action), never because the panel shows.

export const ASSISTANT_CONTEXT_TOKENS = 8192

const LOCK = "filen-assistant-model"

// Unused this long, the model leaves memory, panel open or not. It holds ~4 GB of RAM/VRAM while a reload
// costs ~4-6 s: 10 min keeps it through a working session's follow-ups, where 5 min would reload between
// ordinary pauses and 50 min would hold a quarter of a 16 GB machine for most of an hour after one question.
export const IDLE_UNLOAD_MS = 10 * 60_000

// Long enough for the engine to exit cleanly; terminating the worker frees everything either way.
const EXIT_GRACE_MS = 1000

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

function delay(ms: number): Promise<void> {
	return new Promise(resolve => {
		setTimeout(resolve, ms)
	})
}

// Everything the worker held goes with it: the engine's own worker and its thread pool are nested workers,
// so their wasm memory and the WebGPU device are released too.
function stopWorker(handle: WorkerHandle): void {
	handle.remote[Comlink.releaseProxy]()
	handle.worker.terminate()
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
	// The startup check itself failed (not its verdict): Retry runs it again.
	private checkFailed = false
	private handle: WorkerHandle | null = null
	private releaseLock: (() => void) | null = null
	private watchingLock = false
	private tokenList: readonly string[] = []
	private idleTimer: ReturnType<typeof setTimeout> | undefined
	// Wall-clock time of the last activity: it keeps counting while the machine sleeps, so the first check
	// after waking unloads a model that sat unused across the sleep.
	private lastActivity = 0
	private watchingVisibility = false
	private panelOpen = false
	private busy = 0
	private downloading: Promise<void> | null = null
	private loading: Promise<void> | null = null
	private unloading: Promise<void> | null = null
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

	private settle(kind: Exclude<Settled, "unavailable" | "error" | "downloading" | "loading" | "downloaded">): void {
		this.setState({ kind, model: this.info() })
	}

	private settleDownloaded(idleUnloaded = false): void {
		this.setState({ kind: "downloaded", model: this.info(), idleUnloaded })
	}

	// The cheap startup check: device probe, OPFS sizes and the verified record; no shard content is read.
	init(): Promise<void> {
		this.ready ??= this.refresh().catch((e: unknown) => {
			log.error("assistant", "model check failed", e)
			this.checkFailed = true
			this.setState({ kind: "error", error: "load-failed", model: this.info() })
		})

		return this.ready
	}

	private async refresh(): Promise<void> {
		this.checkFailed = false

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

		if (inspection.complete) {
			this.settleDownloaded()
		} else {
			this.settle("not-downloaded")
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

	private stopWorker(): void {
		const handle = this.handle

		this.handle = null

		if (handle !== null) {
			stopWorker(handle)
		}
	}

	// Opening the panel runs the cheap check early and counts as activity; it never loads the model.
	setPanelOpen(open: boolean): void {
		this.panelOpen = open
		this.syncVisibilityWatch()

		if (!open) {
			return
		}

		this.touch()

		// The first check asks the lock itself; a later opening only re-asks whether another tab took the model.
		if (this.ready === null) {
			void this.init()
		} else {
			void this.ready.then(() => this.checkElsewhere())
		}
	}

	// A model on disk reads as loadable only while no other tab runs it: the startup check saw the lock once,
	// and another tab may have loaded the model since.
	private async checkElsewhere(): Promise<void> {
		if (!this.idleOnDisk()) {
			return
		}

		if ((await this.lockHeldElsewhere()) && this.idleOnDisk()) {
			this.inUseElsewhere()
		}
	}

	private idleOnDisk(): boolean {
		return (
			this.state.kind === "downloaded" &&
			this.releaseLock === null &&
			this.loading === null &&
			this.unloading === null &&
			this.downloading === null
		)
	}

	// The user is doing something with the assistant (typing). Called per keystroke, so it only moves the
	// deadline; the pending timer reads it when it fires.
	touch(): void {
		this.lastActivity = Date.now()
	}

	// Turns (including one waiting on a confirmation card) and confirmed actions keep the model; the idle
	// countdown starts once the last one ends.
	hold(): () => void {
		this.busy++
		clearTimeout(this.idleTimer)
		this.idleTimer = undefined

		let released = false

		return () => {
			if (released) {
				return
			}

			released = true
			this.busy--
			this.touch()
			this.armIdle()
		}
	}

	// Hidden tabs fire timers late (at most once a minute in Chrome) and a sleeping machine not at all, so the
	// timer only prompts a check of the wall-clock deadline, and a visibility change (waking, switching back)
	// checks it too. Coming back to an open panel also re-checks whether another tab took the model meanwhile.
	private readonly onVisibilityChange = () => {
		this.armIdle()

		if (this.panelOpen && document.visibilityState === "visible") {
			void this.checkElsewhere()
		}
	}

	private syncVisibilityWatch(): void {
		this.watchVisibility(this.panelOpen || this.state.kind === "ready")
	}

	private watchVisibility(on: boolean): void {
		if (on === this.watchingVisibility || typeof document === "undefined") {
			return
		}

		this.watchingVisibility = on

		if (on) {
			document.addEventListener("visibilitychange", this.onVisibilityChange)
		} else {
			document.removeEventListener("visibilitychange", this.onVisibilityChange)
		}
	}

	private armIdle(): void {
		clearTimeout(this.idleTimer)
		this.idleTimer = undefined

		if (this.state.kind !== "ready" || this.busy > 0) {
			return
		}

		const now = Date.now()

		// A clock set back must not postpone the unload past one full period.
		this.lastActivity = Math.min(this.lastActivity, now)

		const remaining = this.lastActivity + IDLE_UNLOAD_MS - now

		if (remaining <= 0) {
			void this.unloadModel(true)

			return
		}

		this.idleTimer = setTimeout(() => {
			this.idleTimer = undefined
			this.armIdle()
		}, remaining)
	}

	// From a click: the opt-in download. A loaded model is complete on disk, so there is nothing to fetch, and
	// winding the download down would take the model's worker with it.
	download(): void {
		// Asked first, while the click's user activation lasts (Firefox prompts for it). Without it the
		// browser may evict the 2.7 GB under storage pressure; with or without, the download proceeds.
		void navigator.storage.persist().catch(() => false)

		if (this.downloading !== null || this.loading !== null || this.unloading !== null || this.state.kind === "ready") {
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

				if (this.state.kind === "downloading" && (this.state.receivedBytes < sizeBytes || this.state.waiting)) {
					this.setState({ kind: "downloading", model: this.info(), receivedBytes: sizeBytes, waiting: false })
				}

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

		// Downloading is not a request to run it: the first message (or the Load action) loads it. The worker
		// and the lock go before "downloaded" shows (pruning reads as the download finishing), so a load can
		// never start in this worker and be cut off with it.
		this.stopWorker()
		this.unlock()
		this.settleDownloaded()
	}

	cancelDownload(): void {
		void this.handle?.remote.cancelDownload()
	}

	// From the Load action, or a turn that found the model unloaded. Settles in "ready" or in the state that
	// says why not.
	load(): Promise<void> {
		this.loading ??= (async () => {
			// An unload in flight finishes first: the old worker must be gone before a new one starts.
			await this.unloading
			await this.runLoad()
		})().finally(() => {
			this.loading = null

			// A load that never started a worker (refused, or nothing to load) may still hold the lock an unload
			// kept for it; the lock is only ever held alongside a worker.
			if (this.handle === null && this.downloading === null) {
				this.unlock()
			}

			if (this.redownload) {
				this.redownload = false
				this.download()
			}
		})

		return this.loading
	}

	private async runLoad(): Promise<void> {
		// Only a model on disk, or one whose load failed: nothing to load mid-download or in another tab's hands.
		if (this.recommendation?.supported !== true || (this.state.kind !== "downloaded" && this.state.kind !== "error")) {
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
			this.touch()
			this.syncVisibilityWatch()
			this.armIdle()
		} catch (e) {
			this.stopWorker()
			this.unlock()
			log.error("assistant", "model load failed", e)

			const name = errorName(e)

			if (name === "ModelLoadError:missing") {
				const inspection = await inspectModel(this.manifest)

				if (inspection.complete) {
					this.settleDownloaded()
				} else {
					this.settle("not-downloaded")
				}

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

	// Retry from the error state: the step that failed again. A failed startup check runs again first; one
	// still running is awaited by the download itself.
	retry(): void {
		if (this.checkFailed) {
			this.ready = null
			void this.init()

			return
		}

		if (this.state.kind === "error" && this.state.error !== "download-failed" && this.state.error !== "storage-full") {
			void this.load()

			return
		}

		this.download()
	}

	unload(): Promise<void> {
		return this.unloadModel(false)
	}

	// `idle`: the countdown ran out, which the panel mentions.
	private unloadModel(idle: boolean): Promise<void> {
		this.unloading ??= this.runUnload(idle).finally(() => {
			this.unloading = null
		})

		return this.unloading
	}

	private async runUnload(idle: boolean): Promise<void> {
		await this.loading

		if (this.state.kind !== "ready") {
			return
		}

		const handle = this.handle

		// Detached before anything awaits, so no turn can reach the old engine and nothing here keeps it alive.
		this.handle = null
		this.tokenList = []
		clearTimeout(this.idleTimer)
		this.idleTimer = undefined
		this.settleDownloaded(idle)
		this.syncVisibilityWatch()

		if (handle !== null) {
			await Promise.race([
				handle.remote.unload().catch((e: unknown) => {
					log.warn("assistant", "model unload failed", e)
				}),
				delay(EXIT_GRACE_MS)
			])
			stopWorker(handle)
		}

		// A load requested meanwhile keeps the lock, so no other tab takes the model in between.
		if (this.loading === null) {
			this.unlock()
		}
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
