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
import type { VideoFileRef, VideoIndexUpdate, VideoRootStatus, VideoTitleView } from "@/workers/videoIndex/view.logic"
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

class VideoLibraryClient {
	private snapshot = EMPTY
	private titles = new Map<string, VideoTitleView>()
	private readonly changes = createListenerSet("videos.library")
	private attached = false
	private detach: (() => void) | null = null
	private opening: Promise<void> | null = null
	private readonly artFiles = new Map<string, Promise<ArtFile | null>>()
	private readonly artBatch = new Map<string, { ref: VideoFileRef; resolve: (file: ArtFile | null) => void }>()

	public readonly subscribe = (listener: () => void): (() => void) => this.changes.subscribe(listener)

	public readonly getSnapshot = (): VideoLibrarySnapshot => this.snapshot

	private apply(update: VideoIndexUpdate): void {
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

	// Applies a changed root list (an edit of the configuration) to the index.
	public async configure(doc: LibraryConfigDoc): Promise<void> {
		this.attach()

		if (await sdkApi.videoIndexConfigure(indexRootsOf(doc))) {
			await sdkApi.videoIndexCheck()
		}
	}

	public async rescan(rootUuid: string): Promise<VideoRootStatus | null> {
		this.attach()

		return sdkApi.listVideoRoot(rootUuid, {
			rescan: true
		})
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

	// The file behind a keyless ref, as the art loader takes it. Refs asked for in the same task go to the worker
	// as one call (rebuilt from the index there, no request), and each version is resolved once per session.
	public resolveArtFile(ref: VideoFileRef): Promise<ArtFile | null> {
		const key = `${ref.stableUuid}:${ref.uuid}`
		let resolved = this.artFiles.get(key)

		if (resolved !== undefined) {
			return resolved
		}

		resolved = new Promise<ArtFile | null>(resolve => {
			this.artBatch.set(ref.stableUuid, { ref, resolve })
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
			const file = byUuid.get(ref.uuid)

			// Gone, or a newer version since the ref was handed out: asked again with the next ref.
			if (file === undefined) {
				this.artFiles.delete(`${ref.stableUuid}:${ref.uuid}`)
			}

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
	public async titleArt(
		id: string,
		slot: ArtSlotCode
	): Promise<{ candidates: ArtCandidate[]; resolveFile: (file: number) => Promise<ArtFile | null> } | null> {
		const art = await sdkApi.videoTitleArt(id, slot)

		return art === null
			? null
			: {
					candidates: art.candidates,
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
		this.artFiles.clear()
		this.snapshot = EMPTY
		this.changes.emit()
	}
}

export const videoLibrary = new VideoLibraryClient()
