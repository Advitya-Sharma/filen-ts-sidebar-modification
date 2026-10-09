import * as Comlink from "comlink"
import { activeLibraryRoots, type LibraryConfigDoc } from "@filen/shared/videos"
import type { File as SdkFile } from "@filen/sdk-rs"
import { sdkApi } from "@/lib/sdk/client"
import { log } from "@/lib/log"
import { createListenerSet } from "@/lib/listenerSet"
import { onSignOutPending, readSignOutPending } from "@/lib/signOutPending"
import { onAuthBroadcast } from "@/lib/sdk/session"
import {
	cachedLibraryConfig,
	ensureLibraryConfig,
	holdLibraryConfig,
	libraryConfigNeedsRead,
	persistLibraryConfigWatermark,
	setLibraryAccount
} from "@/features/videos/lib/library"
import { queryClient } from "@/queries/client"
import type { IndexRootConfig } from "@/workers/videoIndex/engine"
import type {
	TitleArtCandidates,
	VideoFileRef,
	VideoHomeItem,
	VideoIndexUpdate,
	VideoRootStatus,
	VideoTitleView
} from "@/workers/videoIndex/view.logic"
import type { ArtFile } from "@/features/videos/lib/artwork"
import type { ArtCandidate } from "@/features/videos/lib/artwork.logic"
import type { ArtSlotCode } from "@/features/videos/lib/artStore.logic"

// The main thread's side of the library index: the titles and root statuses the SDK worker pushes, held as one
// immutable snapshot per change for useSyncExternalStore. Nothing here holds key material; a playable file is
// asked for when it is about to be played.

export const VIDEO_TITLE_DETAIL_KEY_PREFIX = ["videos", "title"] as const

export interface VideoLibrarySnapshot {
	titles: readonly VideoTitleView[]
	byId: ReadonlyMap<string, VideoTitleView>
	roots: readonly VideoRootStatus[]
	// Whether the worker has classified anything yet this session.
	loaded: boolean
}

const EMPTY: VideoLibrarySnapshot = {
	titles: [],
	byId: new Map(),
	roots: [],
	loaded: false
}

export function indexRootsOf(doc: LibraryConfigDoc): IndexRootConfig[] {
	return activeLibraryRoots(doc).map(root => ({
		uuid: root.dirUuid,
		type: root.type,
		enabled: root.enabled
	}))
}

export interface VideoArtCandidates {
	candidates: ArtCandidate[]
	// The keyless uuid of a candidate's file: all a stored-art or drive-thumbnail lookup needs.
	fileUuid: (file: number) => string | null
	// The keyed file, rebuilt in the worker, only for a read that is about to happen.
	resolveFile: (file: number) => Promise<ArtFile | null>
}

// Home rows, grids and search unmounted this long stop the worker's video list (a screen switch remounts them).
const HOME_RELEASE_MS = 30_000
// Retries of a failed subscription to the worker's video list, doubling from the first delay.
const HOME_RETRY_LIMIT = 5
const HOME_RETRY_MS = 1_000

class VideoLibraryClient {
	private snapshot = EMPTY
	private titles = new Map<string, VideoTitleView>()
	private readonly changes = createListenerSet("videos.library")
	private attached = false
	private detach: (() => void) | null = null
	private opening: Promise<void> | null = null
	// Keyed files asked for and not resolved yet; nothing keyed is kept once handed over.
	private readonly artFiles = new Map<string, Promise<ArtFile | null>>()
	private readonly artBatch = new Map<string, { ref: VideoFileRef; resolve: (file: ArtFile | null) => void }>()
	// The library's videos (home rows, watch filters, search) by stable uuid, kept current by the worker's diffs
	// while a screen that reads them is mounted.
	private home = new Map<string, VideoHomeItem>()
	private homeItems: readonly VideoHomeItem[] | null = null
	private homeConsumers = 0
	private homeWanted = false
	private homeEnabled = false
	private homeAttempts = 0
	private homeTimer: ReturnType<typeof setTimeout> | null = null
	private homeRetryArmed = false

	public readonly subscribe = (listener: () => void): (() => void) => this.changes.subscribe(listener)

	public readonly getSnapshot = (): VideoLibrarySnapshot => this.snapshot

	public readonly getHomeItems = (): readonly VideoHomeItem[] | null => this.homeItems

	// A screen that reads the library's videos mounted; the returned release unmounts it. While any is mounted the
	// worker sends the list once and then only what changed with each classification; a while after the last one
	// went, it stops and both sides drop the list.
	public retainHome(): () => void {
		this.homeConsumers++

		if (this.homeTimer !== null) {
			clearTimeout(this.homeTimer)
			this.homeTimer = null
		}

		this.setHomeWanted(true)

		let released = false

		return () => {
			if (released) {
				return
			}

			released = true
			this.homeConsumers--

			if (this.homeConsumers === 0) {
				this.homeTimer = setTimeout(() => {
					this.homeTimer = null
					this.setHomeWanted(false)
				}, HOME_RELEASE_MS)
			}
		}
	}

	private setHomeWanted(wanted: boolean): void {
		if (this.homeWanted === wanted) {
			return
		}

		this.homeWanted = wanted
		this.homeAttempts = 0

		if (!wanted) {
			this.home = new Map()
			this.homeItems = null
			this.changes.emit()
		}

		this.syncHome()
	}

	// Brings the worker's subscription in line with what the screens want; a failure is retried with backoff (a few
	// times), and again when the page comes back online or into focus.
	private syncHome(): void {
		const wanted = this.homeWanted

		// Nothing to stop in a worker this session never subscribed to (or one a sign-out reset).
		if (!wanted && !this.attached) {
			this.homeEnabled = false

			return
		}

		// Attaching subscribes the index first, and syncs on its own.
		this.attach()

		if (wanted === this.homeEnabled) {
			return
		}

		this.homeEnabled = wanted

		sdkApi.videoHomeSubscribe(wanted).catch((e: unknown) => {
			log.warn("videos.library", "subscribing to the library's videos failed", e)

			if (this.homeEnabled !== wanted) {
				return
			}

			this.homeEnabled = !wanted
			this.retryHome()
		})
	}

	private retryHome(): void {
		if (this.homeAttempts >= HOME_RETRY_LIMIT) {
			this.armHomeRetry()

			return
		}

		const delay = HOME_RETRY_MS * 2 ** this.homeAttempts

		this.homeAttempts++
		setTimeout(() => {
			this.syncHome()
		}, delay)
	}

	private armHomeRetry(): void {
		if (this.homeRetryArmed) {
			return
		}

		this.homeRetryArmed = true

		const retry = (): void => {
			window.removeEventListener("online", retry)
			window.removeEventListener("focus", retry)
			this.homeRetryArmed = false
			this.homeAttempts = 0
			this.syncHome()
		}

		window.addEventListener("online", retry)
		window.addEventListener("focus", retry)
	}

	private applyHome(update: Extract<VideoIndexUpdate, { type: "home" }>): void {
		if (!this.homeWanted) {
			return
		}

		const home = update.reset ? new Map<string, VideoHomeItem>() : new Map(this.home)

		for (const item of update.upserts) {
			home.set(item.stableUuid, item)
		}

		for (const stableUuid of update.removed) {
			home.delete(stableUuid)
		}

		this.home = home
		this.homeItems = [...home.values()]
		this.changes.emit()
	}

	private apply(update: VideoIndexUpdate): void {
		if (update.type === "home") {
			this.applyHome(update)

			return
		}

		if (update.type === "roots") {
			this.snapshot = {
				...this.snapshot,
				roots: update.roots
			}
			this.changes.emit()

			return
		}

		const titles = update.reset ? new Map<string, VideoTitleView>() : new Map(this.titles)

		for (const title of update.upserts) {
			titles.set(title.id, title)
		}

		for (const id of update.removed) {
			titles.delete(id)
		}

		this.titles = titles
		this.snapshot = {
			...this.snapshot,
			titles: [...titles.values()],
			byId: titles,
			loaded: true
		}

		// A title's details follow its row: one pass over the cached detail queries.
		const changed = new Set(update.reset ? [] : [...update.upserts.map(title => title.id), ...update.removed])

		if (changed.size > 0) {
			void queryClient.invalidateQueries({
				queryKey: VIDEO_TITLE_DETAIL_KEY_PREFIX,
				predicate: query => typeof query.queryKey[2] === "string" && changed.has(query.queryKey[2])
			})
		}

		this.changes.emit()
	}

	// Subscribes to the worker once per session. The tab with the library open writes its snapshots (the worker
	// serializes writes across tabs), now and when the page goes away, until a sign-out starts here or in any other
	// tab: from then on nothing of the account is persisted.
	private attach(): void {
		if (this.attached) {
			return
		}

		this.attached = true

		void sdkApi.videoIndexSubscribe(
			Comlink.proxy((update: VideoIndexUpdate) => {
				this.apply(update)
			})
		)

		const flush = (): void => {
			if (readSignOutPending()) {
				this.hold()

				return
			}

			persistLibraryConfigWatermark()
			void sdkApi.videoIndexFlush()
		}
		const onVisibility = (): void => {
			if (document.visibilityState === "hidden") {
				flush()
			}
		}
		const hold = (): void => {
			this.hold()
		}

		void sdkApi.videoIndexSetWriter(!readSignOutPending())

		const unwatchSignOut = onSignOutPending(hold)
		const unwatchAuth = onAuthBroadcast(message => {
			if (message.kind === "logout") {
				hold()
			}
		})

		document.addEventListener("visibilitychange", onVisibility)
		window.addEventListener("pagehide", flush)

		// A screen may want the video list from before the subscription (re)started.
		this.syncHome()

		this.detach = () => {
			unwatchSignOut()
			unwatchAuth()
			document.removeEventListener("visibilitychange", onVisibility)
			window.removeEventListener("pagehide", flush)
		}
	}

	// A sign-out started: snapshot writes stop (one under way is abandoned) and the configuration persists nothing.
	public hold(): void {
		holdLibraryConfig()
		void sdkApi.videoIndexSetWriter(false)
	}

	// Brings the library current: the cached roots render from their snapshots first; then, when the session
	// hasn't checked yet, one request for the drive's latest event id decides what still needs reading.
	public open(): Promise<void> {
		this.opening ??= this.runOpen().finally(() => {
			this.opening = null
		})

		return this.opening
	}

	private async runOpen(): Promise<void> {
		this.attach()
		setLibraryAccount(await sdkApi.videoIndexAccount())

		const cachedConfig = cachedLibraryConfig()
		const indexNeedsCheck = cachedConfig === undefined ? true : await sdkApi.videoIndexConfigure(indexRootsOf(cachedConfig.doc))

		if (!indexNeedsCheck && !libraryConfigNeedsRead()) {
			return
		}

		const watermark = await sdkApi.getLastDriveEventId()
		const config = await ensureLibraryConfig(watermark)

		await sdkApi.videoIndexConfigure(indexRootsOf(config.doc))
		await sdkApi.videoIndexCheck(watermark)
	}

	// Applies a changed root list (an edit of the configuration) to the index. Resolves once the index holds the
	// new roots; `scanned` settles when the scans that started for them do.
	public async configure(doc: LibraryConfigDoc): Promise<{ scanned: Promise<void> }> {
		this.attach()

		return {
			scanned: (await sdkApi.videoIndexConfigure(indexRootsOf(doc))) ? sdkApi.videoIndexCheck() : Promise.resolve()
		}
	}

	public async rescan(rootUuid: string): Promise<VideoRootStatus | null> {
		this.attach()

		return sdkApi.listVideoRoot(rootUuid, {
			rescan: true
		})
	}

	// Roots whose scan never started (applying the configuration or the first check failed): the configuration is
	// applied again and they are walked now.
	public async retry(rootUuids: readonly string[]): Promise<void> {
		const config = cachedLibraryConfig()

		if (config !== undefined) {
			await (
				await this.configure(config.doc)
			).scanned
		}

		await Promise.all(rootUuids.map(uuid => this.rescan(uuid)))
	}

	// Stops showing a root's scan; the index keeps what it had.
	public async stopScan(rootUuid: string): Promise<void> {
		await sdkApi.stopVideoRootWalk(rootUuid)
	}

	// Drive files to play or probe, for titles the index holds.
	public async playableFiles(stableUuids: string[]): Promise<SdkFile[]> {
		try {
			return await sdkApi.playableVideoFiles(stableUuids)
		} catch (e) {
			log.warn("videos.library", "playable files unavailable", e)

			return []
		}
	}

	// The file behind a keyless ref, as the art loader takes it right before a read. Refs asked for in the same task
	// go to the worker as one call (rebuilt from the index there, no request); nothing keyed stays here afterwards.
	public resolveArtFile(ref: VideoFileRef): Promise<ArtFile | null> {
		const key = `${ref.stableUuid}:${ref.uuid}`
		let resolved = this.artFiles.get(key)

		if (resolved !== undefined) {
			return resolved
		}

		resolved = new Promise<ArtFile | null>(resolve => {
			this.artBatch.set(ref.stableUuid, { ref, resolve })
		}).finally(() => {
			this.artFiles.delete(key)
		})
		this.artFiles.set(key, resolved)

		if (this.artBatch.size === 1) {
			queueMicrotask(() => {
				void this.flushArtBatch()
			})
		}

		return resolved
	}

	private async flushArtBatch(): Promise<void> {
		const batch = [...this.artBatch.values()]

		this.artBatch.clear()

		const files = await this.playableFiles(batch.map(entry => entry.ref.stableUuid))
		const byUuid = new Map<string, SdkFile>(files.map(file => [file.uuid, file]))

		for (const { ref, resolve } of batch) {
			// Gone, or a newer version since the ref was handed out: null.
			const file = byUuid.get(ref.uuid)

			resolve(
				file === undefined
					? null
					: {
							uuid: file.uuid,
							size: ref.size,
							file
						}
			)
		}
	}

	// A slot's art candidates for a title, with a resolver for the files they name.
	public async titleArt(id: string, slot: ArtSlotCode): Promise<VideoArtCandidates | null> {
		return this.artCandidates(await sdkApi.videoTitleArt(id, slot))
	}

	// One video's art: an episode's own still, else its title's.
	public async itemArt(titleId: string, stableUuid: string, slot: ArtSlotCode): Promise<VideoArtCandidates | null> {
		return this.artCandidates(await sdkApi.videoItemArt(titleId, stableUuid, slot))
	}

	private artCandidates(art: TitleArtCandidates | null): VideoArtCandidates | null {
		return art === null
			? null
			: {
					candidates: art.candidates,
					fileUuid: file => art.files[file]?.uuid ?? null,
					resolveFile: file => {
						const ref = art.files[file]

						return ref === undefined ? Promise.resolve(null) : this.resolveArtFile(ref)
					}
				}
	}

	// Logout: the worker drops its index with the client; nothing of it stays here either.
	public reset(): void {
		this.detach?.()
		this.detach = null
		this.attached = false
		this.opening = null
		this.titles = new Map()
		this.home = new Map()
		this.homeItems = null
		this.homeEnabled = false
		this.homeWanted = this.homeConsumers > 0

		if (this.homeTimer !== null) {
			clearTimeout(this.homeTimer)
			this.homeTimer = null
		}

		this.artFiles.clear()
		this.snapshot = EMPTY
		this.changes.emit()
	}
}

export const videoLibrary = new VideoLibraryClient()
