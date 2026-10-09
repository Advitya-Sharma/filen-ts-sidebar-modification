import * as Comlink from "comlink"
import AssistantWorker from "@/workers/assistant.worker.ts?worker"
import type { AssistantWorkerApi } from "@/workers/assistant.worker"
import {
	AssistantAbortError,
	type AssistantGenerateRequest,
	type AssistantGenerateResult,
	type AssistantModel,
	type AssistantPromptProgress,
	type AssistantSignal,
	QUARANTINE_CONTEXT_TOKENS,
	ROUTER_CONTEXT_TOKENS,
	ROUTER_SYSTEM_PROMPT
} from "@filen/shared/assistant"
import type { AssistantModelInfo, AssistantModelState } from "@/features/assistant/types"
import { DEFAULT_MODEL_BASE_URL, manifestBytes, parseManifest, type AssistantModelManifest } from "@/features/assistant/runtime/manifest"
import { QWEN35_4B } from "@/features/assistant/runtime/models"
import {
	probeDevice,
	recommendRuntime,
	storageVerdict,
	type AssistantBackend,
	type RuntimeRecommendation
} from "@/features/assistant/runtime/deviceCheck"
import { deleteModels, inspectModel, otherModelPresent } from "@/features/assistant/runtime/modelStore"
import { autoStartIntent, readDownloadIntent, writeDownloadIntent } from "@/features/assistant/runtime/downloadIntent"
import type { ModelDownloadProgress } from "@/features/assistant/runtime/download"
import { ENGINE_CRASHED, type GenerationUpdate } from "@/features/assistant/runtime/engineMessages"
import { recoverIfNewerBuild } from "@/lib/appUpdate"
import { ACCOUNT_QUERY_KEY, accountQueryGet, fetchAccount } from "@/queries/account"
import { queryClient } from "@/queries/client"
import { log } from "@/lib/log"

// The page side of the on-device model: decides the model state, owns the worker (spawned only to
// download or to run the model, terminated otherwise, so an idle app holds none of its memory), and holds
// a Web Lock while it does either, so only one tab ever downloads or loads the model. The model loads only
// on an explicit request (a sent message or the Load action), never because the panel shows.

export const ASSISTANT_CONTEXT_TOKENS = 8192

// The engine's cache holds the router's and the content extraction's prompts beside the main loop's, so it
// is sized to all three: one that overflows clears idle slots and loses what they cached (~32 KiB a token).
export const ENGINE_CONTEXT_TOKENS = ASSISTANT_CONTEXT_TOKENS + ROUTER_CONTEXT_TOKENS + QUARANTINE_CONTEXT_TOKENS

const LOCK = "filen-assistant-model"

// Unused this long, the model leaves memory, panel open or not. It holds ~4 GB of RAM/VRAM while a reload
// costs ~4-6 s: 10 min keeps it through a working session's follow-ups, where 5 min would reload between
// ordinary pauses and 50 min would hold a quarter of a 16 GB machine for most of an hour after one question.
export const IDLE_UNLOAD_MS = 10 * 60_000

// Long enough for the engine to exit cleanly; terminating the worker frees everything either way.
const EXIT_GRACE_MS = 1000

// After Stop, how long a generation may take to wind down before its engine counts as stuck and the worker
// is replaced: a stop lands between prompt batches, a fraction of a second on the GPU and up to several
// seconds on the CPU path even at its smaller batch (two threads), where replacing a worker that would have
// settled costs a full reload.
const STOP_WATCHDOG_MS: Readonly<Record<AssistantBackend, number>> = { webgpu: 10_000, cpu: 60_000 }

type Settled = AssistantModelState["kind"]

interface WorkerHandle {
	worker: Worker
	remote: Comlink.Remote<AssistantWorkerApi>
	// Aborted when the worker is stopped or fails: a call to it can never be answered after that.
	closed: AbortController
}

const WORKER_CLOSED = "WorkerClosedError"

function workerClosedError(): Error {
	const error = new Error("The assistant worker stopped")

	error.name = WORKER_CLOSED

	return error
}

// A call to the worker, rejected once the worker is gone: Comlink drops a call's listener when its proxy is
// released, and a worker that failed to start never answers at all.
function untilClosed<T>(handle: WorkerHandle, call: Promise<T>): Promise<T> {
	const { signal } = handle.closed

	if (signal.aborted) {
		return Promise.reject(workerClosedError())
	}

	return new Promise((resolve, reject) => {
		const onClosed = () => {
			reject(workerClosedError())
		}

		signal.addEventListener("abort", onClosed, { once: true })
		void call.then(resolve, reject).finally(() => {
			signal.removeEventListener("abort", onClosed)
		})
	})
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
	handle.closed.abort()
	handle.remote[Comlink.releaseProxy]()
	handle.worker.terminate()
}

// The signed-in account the download's opt-in belongs to. The shell reads the account anyway, so this
// shares its cached or in-flight read.
async function signedInAccount(): Promise<string | null> {
	const account =
		accountQueryGet() ??
		(await queryClient.query({ queryKey: ACCOUNT_QUERY_KEY, queryFn: fetchAccount, staleTime: "static" }).catch(() => undefined))

	return account === undefined ? null : String(account.id)
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
	// What the loaded model runs on.
	private backend: AssistantBackend = "webgpu"
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
	// Whether an interrupted download continues is decided once, at this page's start.
	private resumeDecided = false
	private nextGeneration = 1
	// WebGPU failed to load the model on this page: later loads go straight to the CPU path.
	private gpuFailed = false
	// Signed out: nothing starts again on this page.
	private shutDown = false

	// Nothing is known until the startup check ran: the panel shows neither the opt-in nor the conversation.
	constructor() {
		this.state = { kind: "checking", model: this.info() }
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

	// Not on disk, or only part of it: what a stopped or failed download left, which Remove can still free.
	private settleNotDownloaded(presentBytes = 0): void {
		this.setState({ kind: "not-downloaded", model: this.info(), ...(presentBytes > 0 ? { partialBytes: presentBytes } : {}) })
	}

	private async presentBytes(): Promise<number> {
		return (await inspectModel(this.manifest).catch(() => null))?.presentBytes ?? 0
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
		// Only a verdict a GPU could change asks for an adapter; a mobile or low-memory device is refused either way.
		if (this.recommendation === null) {
			const withoutGpu = recommendRuntime(await probeDevice({ gpu: false }))

			this.recommendation =
				withoutGpu.supported || !withoutGpu.gpuCouldHelp ? withoutGpu : recommendRuntime(await probeDevice({ gpu: true }))
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
			// Cut off after its last byte (pruning older versions): there is nothing left to continue.
			if (readDownloadIntent()?.modelId === this.manifest.id) {
				writeDownloadIntent(null)
			}

			this.settleDownloaded()
		} else {
			this.settleNotDownloaded(inspection.presentBytes)

			if (inspection.presentBytes === 0) {
				await this.checkOutdated()
			}
		}
	}

	// None of this build's model is on disk but another model is: a newer build in another tab replaced it.
	// Downloading the old one again would remove that tab's model in turn, so this tab updates instead (a no-op
	// while this build is still the deployed one).
	private async checkOutdated(): Promise<void> {
		if (await otherModelPresent(this.manifest.id).catch(() => false)) {
			recoverIfNewerBuild()
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
			const handle: WorkerHandle = { worker, remote: Comlink.wrap<AssistantWorkerApi>(worker), closed: new AbortController() }
			// A worker that failed to start (a deploy removed its old file, appUpdate.ts) or threw at the top level
			// answers nothing more: whatever it was doing fails now, rather than waiting forever with the lock held.
			const failed = () => {
				recoverIfNewerBuild()
				this.workerLost(handle)
			}

			worker.addEventListener("error", failed)
			worker.addEventListener("messageerror", failed)
			this.handle = handle
		}

		return this.handle
	}

	// The worker died, crashed its engine or got stuck: a pending call rejects (its own catch settles the state);
	// a loaded model is gone, and the next request loads it again from disk.
	private workerLost(handle: WorkerHandle): void {
		if (this.handle !== handle) {
			return
		}

		this.stopWorker()

		if (this.state.kind === "ready") {
			this.tokenList = []
			clearTimeout(this.idleTimer)
			this.idleTimer = undefined
			this.unlock()
			this.settleDownloaded()
			this.syncVisibilityWatch()
		}
	}

	// Signing out: the worker and its memory go now, the lock with it, and nothing starts again on this page.
	shutdown(): void {
		this.shutDown = true
		writeDownloadIntent(null)
		this.stopWorker()
		this.unlock()
		clearTimeout(this.idleTimer)
		this.idleTimer = undefined
		this.tokenList = []
		this.watchVisibility(false)
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

			// Between turns: the router's cached prompt is read again if a request found it gone.
			if (this.busy === 0 && this.state.kind === "ready") {
				void this.handle?.remote.reprime().catch(() => undefined)
			}

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
		this.startDownload(false)
	}

	// From the signed-in app once it shows, never from the logged-out page: continues a download the user started
	// that a reload, crash or closed tab cut off. Asks nothing and shows nothing by itself; the state says so.
	// Without a recorded opt-in (nearly every start) it costs one synchronous read and leaves the startup
	// check to the idle callback.
	resumeInterruptedDownload(): void {
		if (this.resumeDecided || this.shutDown) {
			return
		}

		this.resumeDecided = true

		if (readDownloadIntent() === null) {
			return
		}

		void this.init().then(() => {
			this.maybeResume()
		})
	}

	// Only the download the user asked for, only while part of it is on disk; the account, the room for the
	// rest and the count of fruitless starts are checked under the lock as it starts. A lock held elsewhere
	// with an opt-in recorded is that tab's download: it stays there, and closing that tab never restarts it
	// here before this page's next start. Persistence is not required: the click's download runs without it
	// too, and the room check is the same.
	private maybeResume(): void {
		if (this.shutDown || this.state.kind !== "not-downloaded" || (this.state.partialBytes ?? 0) === 0) {
			return
		}

		if (readDownloadIntent()?.modelId === this.manifest.id) {
			this.startDownload(true)
		}
	}

	private startDownload(resumed: boolean): void {
		if (this.shutDown || this.downloading !== null || this.loading !== null || this.unloading !== null || this.state.kind === "ready") {
			return
		}

		this.downloading = this.runDownload(resumed).finally(() => {
			this.downloading = null
		})
	}

	private async runDownload(resumed: boolean): Promise<void> {
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

				const account = await signedInAccount()
				const intent = resumed
					? autoStartIntent(readDownloadIntent(), { modelId: this.manifest.id, account, presentBytes: before.presentBytes })
					: account === null
						? null
						: { modelId: this.manifest.id, account, autoStarts: 0, presentBytes: before.presentBytes }

				// Unasked, nothing starts without the room for the rest, once the part on disk is gone, for another
				// account, or after starts that kept finding no progress (the opt-in goes; a click records it anew).
				// An account that could not be read (offline) keeps the opt-in for the next start.
				if (resumed && (intent === null || !verdict.ok || before.presentBytes === 0)) {
					if ((intent === null && account !== null) || before.presentBytes === 0) {
						writeDownloadIntent(null)
					}

					this.unlock()
					this.settleNotDownloaded(before.presentBytes)

					return
				}

				if (!verdict.ok) {
					this.setState({ kind: "unavailable", reason: "insufficient-storage", model: this.info() })
					this.unlock()

					return
				}

				const sizeBytes = manifestBytes(this.manifest)
				const handle = this.worker()

				writeDownloadIntent(intent)
				this.setState({ kind: "downloading", model: this.info(), receivedBytes: before.presentBytes, waiting: false, resumed })
				await untilClosed(
					handle,
					handle.remote.download(
						this.manifest,
						this.baseUrl,
						Comlink.proxy(({ receivedBytes, waiting }: ModelDownloadProgress) => {
							if (this.state.kind === "downloading") {
								this.setState({
									kind: "downloading",
									model: this.info(),
									receivedBytes: Math.min(sizeBytes, Math.max(0, receivedBytes)),
									waiting,
									resumed
								})
							}
						})
					)
				)

				if (this.state.kind === "downloading" && (this.state.receivedBytes < sizeBytes || this.state.waiting)) {
					this.setState({ kind: "downloading", model: this.info(), receivedBytes: sizeBytes, waiting: false, resumed })
				}

				// An older model version is dead weight once its successor is complete.
				await deleteModels(this.manifest.id).catch((e: unknown) => {
					log.warn("assistant", "pruning old models failed", e)
				})
			}
		} catch (e) {
			// Stopped or failed, it is over: only a download the page died under continues by itself.
			writeDownloadIntent(null)
			this.stopWorker()
			this.unlock()

			const name = errorName(e)

			const partial = await this.presentBytes()

			if (name === "AbortError") {
				this.settleNotDownloaded(partial)
			} else {
				log.error("assistant", "model download failed", e)
				this.setState({
					kind: "error",
					error: name === "ModelDownloadError:storage-full" || name === "QuotaExceededError" ? "storage-full" : "download-failed",
					model: this.info(),
					...(partial > 0 ? { partialBytes: partial } : {})
				})
			}

			return
		}

		// Downloading is not a request to run it: the first message (or the Load action) loads it. The worker
		// and the lock go before "downloaded" shows (pruning reads as the download finishing), so a load can
		// never start in this worker and be cut off with it.
		writeDownloadIntent(null)
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

	private async loadIn(recommendation: Extract<RuntimeRecommendation, { supported: true }>): Promise<readonly string[]> {
		const handle = this.worker()
		const { specialTokens } = await untilClosed(
			handle,
			handle.remote.load(
				this.manifest,
				{
					backend: recommendation.backend,
					threads: recommendation.threads,
					contextTokens: ENGINE_CONTEXT_TOKENS,
					primeSystem: ROUTER_SYSTEM_PROMPT
				},
				Comlink.proxy(() => {
					if (this.state.kind === "loading") {
						this.setState({ kind: "loading", model: this.info(), verifying: true })
					}
				})
			)
		)

		this.backend = recommendation.backend

		return specialTokens
	}

	// WebGPU first; when the GPU cannot take the model (its memory, its driver), the CPU path in a fresh
	// worker, once, and from then on for this page.
	private async loadWithFallback(recommendation: Extract<RuntimeRecommendation, { supported: true }>): Promise<readonly string[]> {
		try {
			return await this.loadIn(recommendation)
		} catch (e) {
			const name = errorName(e)

			if (recommendation.backend !== "webgpu" || name === "ModelLoadError:missing" || name === WORKER_CLOSED) {
				throw e
			}

			const cpu = recommendRuntime(await probeDevice({ gpu: false }))

			if (!cpu.supported) {
				throw e
			}

			log.warn("assistant", "WebGPU load failed, trying the CPU path", e)
			this.gpuFailed = true
			this.stopWorker()

			return this.loadIn(cpu)
		}
	}

	private async runLoad(): Promise<void> {
		// Only a model on disk, or one whose load failed: nothing to load mid-download or in another tab's hands.
		if (this.shutDown || this.recommendation?.supported !== true || (this.state.kind !== "downloaded" && this.state.kind !== "error")) {
			return
		}

		// The backend: WebGPU when the adapter is up to it, unless it already failed here.
		const recommendation = recommendRuntime(await probeDevice({ gpu: !this.gpuFailed }))

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
			this.tokenList = await this.loadWithFallback(recommendation)
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
					this.settleNotDownloaded(inspection.presentBytes)
				}

				// Some shards are gone or were damaged, the rest are still verified: the user opted in already, so
				// only the missing ones are fetched again, right after this load has wound down.
				this.redownload = !inspection.complete && inspection.verifiedBytes > 0

				if (inspection.presentBytes === 0) {
					await this.checkOutdated()
				}
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
		writeDownloadIntent(null)
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
		{
			signal,
			onText,
			onProgress,
			onToolCall
		}: {
			signal: AssistantSignal
			onText?: (text: string) => void
			onProgress?: (progress: AssistantPromptProgress) => void
			onToolCall?: (call: { index: number; name: string }) => void
		}
	): Promise<AssistantGenerateResult> {
		const handle = this.handle

		if (handle === null || this.state.kind !== "ready") {
			throw new Error("The model is not loaded")
		}

		if (signal.aborted) {
			throw new AssistantAbortError()
		}

		const id = this.nextGeneration++
		let stuck: ReturnType<typeof setTimeout> | undefined
		const cancel = () => {
			void handle.remote.cancel(id).catch(() => undefined)
			// The worker answers a stop only between engine steps; one that never answers is replaced.
			stuck = setTimeout(() => {
				log.warn("assistant", "a stopped generation did not wind down")
				this.workerLost(handle)
			}, STOP_WATCHDOG_MS[this.backend])
		}

		// One port for every kind of update, opened only when someone listens. Updates travel on that port and
		// the result on the worker's own, in no guaranteed order: one that lands after the result is stale.
		let settled = false
		const onUpdate =
			onText === undefined && onProgress === undefined && onToolCall === undefined
				? undefined
				: Comlink.proxy((update: GenerationUpdate) => {
						if (settled) {
							return
						}

						switch (update.type) {
							case "text":
								onText?.(update.text)
								break
							case "progress":
								onProgress?.({ processed: update.processed, total: update.total })
								break
							case "call":
								onToolCall?.({ index: update.index, name: update.name })
								break
						}
					})

		signal.addEventListener("abort", cancel)

		const call = untilClosed(handle, handle.remote.generate(id, request, onUpdate))

		// Stop ends the turn at once; the worker winds the generation down on its own, and runs the next one only
		// after it (the worker serializes them).
		void call.then(
			() => {
				clearTimeout(stuck)
			},
			() => {
				clearTimeout(stuck)
			}
		)

		try {
			return await new Promise<AssistantGenerateResult>((resolve, reject) => {
				const onAbort = () => {
					reject(new AssistantAbortError())
				}

				signal.addEventListener("abort", onAbort, { once: true })
				void call.then(resolve, reject).finally(() => {
					signal.removeEventListener("abort", onAbort)
				})
			})
		} catch (e) {
			if (errorName(e) === ENGINE_CRASHED) {
				log.error("assistant", "the engine crashed", e)
				this.workerLost(handle)
			}

			throw errorName(e) === "AbortError" || isAborted(signal) ? new AssistantAbortError() : e
		} finally {
			settled = true
			signal.removeEventListener("abort", cancel)
		}
	}
}
