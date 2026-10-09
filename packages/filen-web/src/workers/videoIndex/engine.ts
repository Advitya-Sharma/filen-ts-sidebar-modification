import {
	applyLibraryOverrides,
	classifyVideoLibrarySteps,
	splitMergeKeysOf,
	type LibraryOverride,
	type LibraryRootType,
	type LibraryScanDir,
	type VideoLibrary
} from "@filen/shared/videos"
import type { DirsAndFiles, DriveEvent, File, SocketEvent, UuidStr } from "@filen/sdk-rs"
import { log } from "@/lib/log"
import type { ArtSlotCode } from "@/features/videos/lib/artStore.logic"
import { rootSize } from "@/features/videos/lib/library.logic"
import {
	applyDriveEvent,
	applyLookedUpFile,
	applyScopedWalk,
	applyWalk,
	emptyRootState,
	fromRecords,
	liveFiles,
	projectWalk,
	type IndexFile,
	type PatchResult,
	type RootState
} from "@/workers/videoIndex/rootIndex.logic"
import { encodeSnapshot, nextSlot, pickSnapshot, snapshotIsCurrent } from "@/workers/videoIndex/snapshot.logic"
import {
	LibraryFiles,
	buildTitleDetail,
	buildTitles,
	buildHomeItems,
	diffRecords,
	diffTitles,
	itemArtCandidates,
	titleArtCandidates,
	type TitleArtCandidates,
	type VideoHomeItem,
	type VideoIndexUpdate,
	type VideoRootStatus,
	type VideoTitleDetail,
	type VideoTitleView
} from "@/workers/videoIndex/view.logic"

// The Videos library index, inside the SDK worker: per-root snapshots loaded from OPFS, kept current from the
// socket's drive events, re-walked only when the event stream can't vouch for them, and classified into the
// keyless view the main thread renders.

// Patches to the classification wait for a quiet spell, but never longer than the cap.
export const PATCH_DEBOUNCE_MS = 1_500
export const PATCH_MAX_WAIT_MS = 5_000
// Directories that came into a root are re-listed together after this; past the count, the root is walked.
export const SCOPED_WALK_COALESCE_MS = 2_000
export const SCOPED_WALK_MAX_PENDING = 8
// At most one automatic walk per root in this window; a manual rescan isn't held to it.
export const WALK_FLOOR_MS = 30 * 60_000
// Snapshot writes: a change is written after this, and no sooner than the interval after the previous write.
export const WRITE_DEBOUNCE_MS = 10_000
export const WRITE_MIN_INTERVAL_MS = 60_000
// Classification runs in slices of about this much, so other worker calls aren't held up.
export const SLICE_MS = 8
// Events held for a root while it can't apply them; past this, it is walked instead.
export const EVENT_BUFFER_CAP = 5_000

export interface IndexRootConfig {
	uuid: string
	type: LibraryRootType
	enabled: boolean
}

export interface SnapshotStore {
	// Both slots, null for a missing one.
	read: (rootUuid: string) => Promise<[Uint8Array | null, Uint8Array | null]>
	// Each slot's header generation, null for a missing or unreadable one.
	readGenerations: (rootUuid: string) => Promise<[bigint | null, bigint | null]>
	// Abandons the write, leaving the slot as it was, when `cancelled` says so before it completes.
	write: (rootUuid: string, slot: 0 | 1, bytes: Uint8Array<ArrayBuffer>, cancelled: () => boolean) => Promise<void>
	remove: (rootUuid: string) => Promise<void>
}

// `ancestors`: the root's ancestor directories as far as known (trashing one takes the root away too).
export type WalkOutcome = { type: "ok"; name: string; raw: DirsAndFiles; ancestors: string[] } | { type: "gone" }

export interface VideoIndexDeps {
	// A root's recursive listing and its current name. The listing may be shared with another walker of the same
	// directory, but only one that started after `startedAfter` (a walkMark taken before the watermark was read).
	walkRoot: (rootUuid: string, startedAfter: number) => Promise<WalkOutcome>
	walkMark: () => number
	// A directory's recursive listing, for one that came into a root.
	walkDir: (dirUuid: string) => Promise<DirsAndFiles | null>
	lastDriveEventId: () => Promise<bigint>
	// The signed-in account's drive root: a root that is the drive root leaves Filen's own directory out.
	driveRoot: () => string | null
	getFile: (uuid: string) => Promise<File | undefined>
	store: SnapshotStore
	now: () => number
	// Ends the current task, so a long classification yields between slices.
	yieldSlice: () => Promise<void>
}

type RootMode =
	// Loaded, not yet checked against the drive: events are held.
	| "pending"
	// Current: every drive event since the watermark was applied, in order.
	| "live"
	// Known or possibly behind: events are applied as they come, without a watermark, until a walk.
	| "stale"

// manual: asked for by the user; event: the event stream changed the root in a way only a walk catches up with;
// auto: the index may be behind. Only auto is held to the floor.
type WalkReason = "manual" | "event" | "auto"

const WALK_REASON_RANK: Record<WalkReason, number> = {
	auto: 0,
	event: 1,
	manual: 2
}

// Consecutive failed snapshot writes after which writing stops for the session (storage unusable here).
const MAX_WRITE_FAILURES = 3

interface BufferedEvent {
	id: bigint
	// null for an event that couldn't be decoded.
	inner: DriveEvent | null
}

interface RootRuntime {
	config: IndexRootConfig
	state: RootState | null
	loading: Promise<void> | null
	status: VideoRootStatus["state"]
	mode: RootMode
	buffer: BufferedEvent[] | null
	bufferOverflow: boolean
	walking: Promise<void> | null
	// The user stopped the walk under way: its result is dropped when it lands.
	stopped: boolean
	// The user stopped the root's scan: nothing walks it again until a rescan is asked for.
	userStopped: boolean
	// The index doesn't reflect a complete walk: the latest one failed, came back partial or was stopped.
	incomplete: boolean
	walkAgain: WalkReason | null
	lastWalkAt: number
	deferredWalk: ReturnType<typeof setTimeout> | null
	scoped: Set<string>
	scopedTimer: ReturnType<typeof setTimeout> | null
	moved: boolean
	// The root's ancestors as far as known.
	ancestors: Set<string>
	// The snapshot was written under other classification or filtering rules: walk, whatever the floor.
	outdated: boolean
	// Video files the latest classification found in this root.
	videoCount: number
	// Bumped when a walk's result is applied, and copied when a classification that includes it is done: a scan
	// reads as under way until its count is known.
	walkSeq: number
	countedSeq: number
	version: number
	writtenVersion: number
	writeTimer: ReturnType<typeof setTimeout> | null
	writing: Promise<void> | null
	lastWriteAt: number
	// Only the watermark changed since the last write: written with the next one, or when the page goes away.
	watermarkDirty: boolean
}

// One classification and what it was made from: reused while neither the roots' contents nor the split corrections
// change (`seq`, `split`).
interface ClassifiedInput {
	seq: number
	split: string
	library: VideoLibrary
	files: IndexFile[]
	loaded: RootRuntime[]
	seqs: number[]
	ends: number[]
}

interface CurrentLibrary {
	files: LibraryFiles
	titles: Map<string, VideoTitleView>
}

function isUuidStr(value: string): value is UuidStr {
	return value.split("-").length >= 4
}

async function drainSteps<T>(steps: Generator<void, T>, yieldSlice: () => Promise<void>): Promise<T> {
	let sliceStart = performance.now()

	for (;;) {
		const step = steps.next()

		if (step.done === true) {
			return step.value
		}

		if (performance.now() - sliceStart >= SLICE_MS) {
			await yieldSlice()

			sliceStart = performance.now()
		}
	}
}

export class VideoIndexEngine {
	private readonly deps: VideoIndexDeps
	private readonly roots = new Map<string, RootRuntime>()
	private readonly listeners = new Set<(update: VideoIndexUpdate) => void>()
	private readonly signatures = new Map<string, string>()
	// The video list's last sent state (stable uuid -> serialized item) while the main thread wants it.
	private homeSignatures: Map<string, string> | null = null
	private homeSent = false
	private current: CurrentLibrary | null = null
	// Bumped by reset: async work started before it drops its result.
	private token = 0
	private writer = false
	// Bumped whenever writing stops, so a write under way is abandoned.
	private writeEpoch = 0
	private writeFailures = 0
	// The last drive event id delivered, and whether the socket is up.
	private lastEventId: bigint | null = null
	private socketLive = false
	private sawReconnecting = false
	private rebuilding: Promise<void> | null = null
	private rebuildAgain = false
	private patchTimer: ReturnType<typeof setTimeout> | null = null
	private firstPatchAt = 0
	// Drive events applied while a lookup or scoped walk is out, re-applied over its result.
	private eventSeq = 0
	private asyncPatches = 0
	private patchLog: { seq: number; root: string; inner: DriveEvent }[] = []
	private readonly lookups = new Map<string, Promise<void>>()
	// The library configuration's manual corrections, by key, and their serialized form (to skip an unchanged set).
	private overrides: Readonly<Record<string, LibraryOverride>> = {}
	private overridesSignature = "{}"
	// Bumped whenever what the classifier reads changes (roots, their files, the split corrections); a rebuild for
	// any other correction reuses the last classification.
	private inputSeq = 0
	private classified: ClassifiedInput | null = null
	private notTransportStream: ReadonlySet<string> = new Set()
	private notTransportStreamSignature = ""

	public constructor(deps: VideoIndexDeps) {
		this.deps = deps
	}

	// Sets the `.ts`/`.mts` files a read of their content found to be no transport stream (TypeScript): the
	// library leaves them out. A changed set classifies again (no network).
	public setNotTransportStreams(uuids: readonly string[]): void {
		const signature = [...uuids].sort().join(",")

		if (signature === this.notTransportStreamSignature) {
			return
		}

		this.notTransportStream = new Set(uuids)
		this.notTransportStreamSignature = signature

		if (this.current !== null || this.rebuilding !== null) {
			this.scheduleRebuild("now")
		} else {
			this.inputSeq++
		}
	}

	// Sets the manual corrections; a changed set applies them again (no network), classifying again only when the
	// split corrections changed.
	public setOverrides(overrides: Readonly<Record<string, LibraryOverride>>): void {
		const signature = JSON.stringify(overrides)

		if (signature === this.overridesSignature) {
			return
		}

		const splitChanged = splitSignature(overrides) !== splitSignature(this.overrides)

		this.overrides = overrides
		this.overridesSignature = signature

		// A classification under way read the previous set: it runs again once done.
		if (this.current !== null || this.rebuilding !== null) {
			this.scheduleRebuild("now", splitChanged)
		}
	}

	// ── Configuration ──

	// Sets the roots. New ones load their snapshot (no network); a disabled root is unloaded and keeps its
	// snapshot; one no longer listed loses it. Resolves whether a check against the drive is needed.
	public async configure(roots: readonly IndexRootConfig[]): Promise<boolean> {
		const listed = new Set(roots.map(root => root.uuid))
		let changed = false

		for (const [uuid, rt] of this.roots) {
			const config = roots.find(root => root.uuid === uuid)

			if (config?.enabled !== true) {
				this.unload(rt, !listed.has(uuid))
				changed = true
			}
		}

		const loads: Promise<void>[] = []

		for (const config of roots) {
			if (!config.enabled) {
				continue
			}

			const existing = this.roots.get(config.uuid)

			if (existing !== undefined) {
				changed ||= existing.config.type !== config.type
				existing.config = config

				if (existing.loading !== null) {
					loads.push(existing.loading)
				}

				continue
			}

			const rt = this.createRuntime(config)

			this.roots.set(config.uuid, rt)
			rt.loading = this.load(rt)
			loads.push(rt.loading)
			changed = true
		}

		await Promise.all(loads)

		// The first configuration classifies even an empty library, so a subscriber learns there is nothing.
		if (changed || this.current === null) {
			this.pushRoots()
			this.scheduleRebuild("now")
		}

		return this.needsCheck()
	}

	private createRuntime(config: IndexRootConfig): RootRuntime {
		return {
			config,
			state: null,
			loading: null,
			status: "loading",
			mode: "pending",
			buffer: [],
			bufferOverflow: false,
			walking: null,
			stopped: false,
			userStopped: false,
			incomplete: false,
			walkAgain: null,
			lastWalkAt: 0,
			deferredWalk: null,
			scoped: new Set(),
			scopedTimer: null,
			moved: false,
			ancestors: new Set(),
			outdated: false,
			videoCount: 0,
			walkSeq: 0,
			countedSeq: 0,
			version: 0,
			writtenVersion: 0,
			writeTimer: null,
			writing: null,
			lastWriteAt: 0,
			watermarkDirty: false
		}
	}

	private async load(rt: RootRuntime): Promise<void> {
		const token = this.token

		try {
			const decoded = pickSnapshot(await this.deps.store.read(rt.config.uuid), rt.config.uuid)

			if (token !== this.token || this.roots.get(rt.config.uuid) !== rt) {
				return
			}

			if (decoded !== null) {
				rt.state = fromRecords(decoded)
				rt.status = "ready"
				rt.outdated = !snapshotIsCurrent(decoded)
				// Only partial walks were ever stored.
				rt.incomplete = rt.state.completeAt === 0
			}
		} catch (e) {
			log.warn("videos.index", "snapshot unreadable; the root is walked instead", rt.config.uuid, e)
		} finally {
			rt.loading = null
		}
	}

	private unload(rt: RootRuntime, removeSnapshot: boolean): void {
		const uuid = rt.config.uuid

		this.clearTimers(rt)
		this.roots.delete(uuid)

		const flushed = this.writer && rt.version > rt.writtenVersion && !removeSnapshot ? this.writeNow(rt) : Promise.resolve()

		if (removeSnapshot) {
			void flushed
				.then(() => this.deps.store.remove(uuid))
				.catch((e: unknown) => {
					log.warn("videos.index", "removing a root's snapshot failed", uuid, e)
				})
		}
	}

	private clearTimers(rt: RootRuntime): void {
		for (const timer of [rt.deferredWalk, rt.scopedTimer, rt.writeTimer]) {
			if (timer !== null) {
				clearTimeout(timer)
			}
		}

		rt.deferredWalk = null
		rt.scopedTimer = null
		rt.writeTimer = null
	}

	// Not walking, and unchecked against the drive's event stream or never walked; a root whose scan the user
	// stopped waits for a rescan.
	private unchecked(rt: RootRuntime): boolean {
		return rt.walking === null && !rt.userStopped && (rt.state === null || rt.mode === "pending")
	}

	// Some loaded root hasn't been checked against the drive's event stream, or hasn't been walked at all.
	public needsCheck(): boolean {
		for (const rt of this.roots.values()) {
			if (this.unchecked(rt)) {
				return true
			}
		}

		return false
	}

	// Compares each unchecked root's watermark with the drive's latest event id (`watermark`, one request the
	// caller made): equal means nothing changed anywhere in the drive, so no walk. Unwalked roots are walked.
	public async check(watermark?: bigint): Promise<void> {
		const pending = [...this.roots.values()].filter(rt => this.unchecked(rt))

		if (pending.length === 0) {
			return
		}

		// Taken before the watermark is read (the caller read one before calling): a listing begun earlier may not
		// reflect what that watermark covers.
		const mark = this.deps.walkMark()
		let latest = watermark ?? (this.socketLive ? this.lastEventId : null)

		if (latest === null) {
			try {
				latest = await this.deps.lastDriveEventId()
			} catch (e) {
				// Never-walked roots are walked anyway (a walk reads its own and reports its failure); the others stay
				// unchecked until the next check.
				log.warn("videos.index", "latest drive event id unavailable", e)
			}
		}

		const walks: Promise<void>[] = []

		for (const rt of pending) {
			if (rt.state === null) {
				walks.push(this.walk(rt, "manual", latest === null ? undefined : { watermark: latest, mark }))

				continue
			}

			if (latest === null || rt.mode !== "pending" || rt.walking !== null) {
				continue
			}

			if (rt.outdated) {
				walks.push(this.walk(rt, "event", { watermark: latest, mark }))

				continue
			}

			if (rt.state.watermark !== null && this.settle(rt, rt.state.watermark, latest)) {
				continue
			}

			walks.push(this.walk(rt, "auto", { watermark: latest, mark }))
		}

		this.pushRoots()

		await Promise.all(walks)
	}

	// A user-asked rescan: walks now, whatever the floor.
	public rescan(rootUuid: string): Promise<void> {
		const rt = this.roots.get(rootUuid)

		if (rt === undefined) {
			return Promise.resolve()
		}

		rt.userStopped = false

		return this.walk(rt, "manual")
	}

	// The user stopped a root's scan. A walk can't be aborted, so it finishes unseen and its result is dropped:
	// the index keeps what it had, flagged incomplete at once, and nothing walks it again until a rescan.
	public stopWalk(rootUuid: string): void {
		const rt = this.roots.get(rootUuid)

		if (rt === undefined) {
			return
		}

		if (rt.walking === null || rt.stopped) {
			return
		}

		rt.stopped = true
		rt.userStopped = true
		rt.incomplete = true
		rt.walkAgain = null

		if (rt.state === null) {
			rt.status = "error"
		}

		this.pushRoots()
	}

	// Applies the events held since `watermark`, which the state is current to, when they run without a gap through
	// `through` (the drive's latest event id): then nothing is missing. False, changing nothing, otherwise.
	private settle(rt: RootRuntime, watermark: bigint, through: bigint): boolean {
		const state = rt.state

		if (state === null) {
			return false
		}

		const buffer = rt.buffer ?? []
		let expected = watermark + 1n

		if (rt.bufferOverflow) {
			return false
		}

		for (const event of buffer) {
			if (event.id < expected) {
				continue
			}

			if (event.id !== expected || event.inner === null) {
				return false
			}

			expected++
		}

		if (expected - 1n < through) {
			return false
		}

		rt.buffer = null
		rt.mode = "live"

		for (const event of buffer) {
			if (event.id > watermark && event.inner !== null) {
				this.apply(rt, event.inner, null)
			}
		}

		this.setWatermark(rt, expected - 1n)

		return true
	}

	// Gives up on holding events: applies what is held as best it can and drops the watermark until a walk.
	private toStale(rt: RootRuntime): void {
		const buffer = rt.buffer

		rt.buffer = null
		rt.bufferOverflow = false
		rt.mode = "stale"

		if (rt.state === null) {
			return
		}

		for (const event of buffer ?? []) {
			if (event.inner !== null) {
				this.apply(rt, event.inner, null)
			}
		}

		this.setWatermark(rt, null)
	}

	// A watermark-only change (every drive event anywhere moves it) isn't worth a snapshot write of its own: it goes
	// with the next content write, or when the page goes away.
	private setWatermark(rt: RootRuntime, watermark: bigint | null): void {
		if (rt.state === null || rt.state.watermark === watermark) {
			return
		}

		rt.state.watermark = watermark
		rt.watermarkDirty = true
	}

	// ── Walks ──

	private floorBase(rt: RootRuntime): number {
		return Math.max(rt.lastWalkAt, rt.state?.completeAt ?? 0)
	}

	// Single flight per root: a walk asked for while one runs follows it (the strongest reason wins).
	private walk(rt: RootRuntime, reason: WalkReason, known?: { watermark: bigint; mark: number }): Promise<void> {
		// Asked for again after a stop: the walk under way counts again rather than starting another.
		if (rt.walking !== null && rt.stopped && reason === "manual") {
			rt.stopped = false
			this.pushRoots()

			return rt.walking
		}

		if (rt.walking !== null) {
			rt.walkAgain = rt.walkAgain === null || WALK_REASON_RANK[reason] > WALK_REASON_RANK[rt.walkAgain] ? reason : rt.walkAgain

			return rt.walking
		}

		const now = this.deps.now()

		if (reason === "auto" && rt.state !== null && now - this.floorBase(rt) < WALK_FLOOR_MS) {
			this.deferWalk(rt)
			this.toStale(rt)
			this.pushRoots()

			return Promise.resolve()
		}

		if (rt.deferredWalk !== null) {
			clearTimeout(rt.deferredWalk)
			rt.deferredWalk = null
		}

		// A live root is current to its watermark, so walking from there needs no lookup.
		const base = rt.mode === "live" ? (rt.state?.watermark ?? null) : null
		const mark = known?.mark ?? this.deps.walkMark()

		rt.lastWalkAt = now
		rt.buffer ??= []

		const walking = this.runWalk(rt, base ?? known?.watermark ?? null, mark).finally(() => {
			rt.walking = null
			this.pushRoots()

			const again = rt.walkAgain

			rt.walkAgain = null

			if (again !== null && this.roots.get(rt.config.uuid) === rt) {
				void this.walk(rt, again)
			}
		})

		rt.walking = walking
		this.pushRoots()

		return walking
	}

	// A stopped walk landed: a root that was current stays current when the events held meanwhile run without a
	// gap through the latest one delivered; otherwise it is stale until the next walk.
	private keepAfterStop(rt: RootRuntime): void {
		const watermark = rt.state?.watermark ?? null

		if (
			rt.mode === "live" &&
			watermark !== null &&
			this.socketLive &&
			this.lastEventId !== null &&
			this.settle(rt, watermark, this.lastEventId)
		) {
			return
		}

		this.toStale(rt)
	}

	private deferWalk(rt: RootRuntime): void {
		if (rt.deferredWalk !== null) {
			return
		}

		const delay = Math.max(0, this.floorBase(rt) + WALK_FLOOR_MS - this.deps.now())

		rt.deferredWalk = setTimeout(() => {
			rt.deferredWalk = null

			if (this.roots.get(rt.config.uuid) === rt) {
				void this.walk(rt, "auto")
			}
		}, delay)
	}

	private async runWalk(rt: RootRuntime, knownWatermark: bigint | null, mark: number): Promise<void> {
		const token = this.token
		const uuid = rt.config.uuid

		try {
			const watermark = knownWatermark ?? (this.socketLive ? this.lastEventId : null) ?? (await this.deps.lastDriveEventId())
			const startedAt = this.deps.now()
			const outcome = await this.deps.walkRoot(uuid, mark)

			if (token !== this.token || this.roots.get(uuid) !== rt) {
				return
			}

			if (outcome.type === "gone") {
				rt.stopped = false
				rt.status = "missing"
				this.toStale(rt)

				return
			}

			if (rt.stopped) {
				rt.stopped = false
				this.keepAfterStop(rt)

				return
			}

			const projection = projectWalk(uuid, outcome.raw, uuid === this.deps.driveRoot())
			const state = rt.state ?? emptyRootState(uuid, outcome.name)

			state.name = outcome.name
			rt.ancestors = new Set(outcome.ancestors)
			applyWalk({
				state,
				walk: projection,
				startedAt,
				now: this.deps.now()
			})
			rt.state = state
			rt.status = "ready"
			rt.walkSeq++

			const buffer = rt.buffer ?? []
			const overflow = rt.bufferOverflow
			const walked = new Set(projection.dirs.map(dir => dir.uuid))
			let expected = watermark + 1n
			let contiguous = projection.complete && !overflow

			rt.buffer = null
			rt.bufferOverflow = false

			// Everything delivered since the walk began is applied over it in order; an event the walk already
			// reflects sets what is already there.
			for (const event of buffer) {
				if (event.id >= expected) {
					contiguous &&= event.id === expected && event.inner !== null
					expected = event.id + 1n
				}

				if (event.inner !== null) {
					this.apply(rt, event.inner, walked)
				}
			}

			rt.mode = contiguous ? "live" : "stale"
			this.markPersist(rt)
			this.setWatermark(rt, contiguous ? expected - 1n : null)
			this.scheduleRebuild("now")

			rt.incomplete = !projection.complete

			if (projection.complete) {
				rt.outdated = false
			}

			// Events were missed while it ran: once more (the floor permitting). A partial walk (orphans in the listing)
			// is tried again once the floor passes, so its removals and tombstone purges aren't put off for good.
			if (!contiguous && projection.complete) {
				rt.walkAgain ??= "auto"
			} else if (!projection.complete) {
				this.deferWalk(rt)
			}
		} catch (e) {
			if (token !== this.token) {
				return
			}

			log.warn("videos.index", "root walk failed; the index keeps what it had", uuid, e)

			rt.stopped = false
			rt.incomplete = true
			rt.status = rt.state === null ? "error" : "ready"
			this.toStale(rt)
			this.deferWalk(rt)
		}
	}

	// ── Socket ──

	public onSocketEvent(event: SocketEvent): void {
		switch (event.type) {
			case "drive":
				this.onDriveEvent(event.driveMessageId, event.inner)

				break

			case "driveMalformed":
				this.onDriveEvent(event.driveMessageId, null)

				break

			case "reconnecting":
				this.onReconnecting()

				break

			case "authSuccess":
				this.onAuthSuccess()

				break

			default:
				break
		}
	}

	private onReconnecting(): void {
		this.socketLive = false
		this.sawReconnecting = true
		// What was delivered before the drop says nothing about what the drive holds after it.
		this.lastEventId = null

		// What arrives after the drop may follow a gap: held until checked again.
		for (const rt of this.roots.values()) {
			if (rt.mode === "live") {
				rt.mode = "pending"
				rt.buffer ??= []
			}
		}
	}

	private onAuthSuccess(): void {
		this.socketLive = true

		if (!this.sawReconnecting) {
			return
		}

		this.sawReconnecting = false

		if ([...this.roots.values()].some(rt => rt.state !== null && rt.mode === "pending")) {
			void this.check().catch((e: unknown) => {
				log.warn("videos.index", "check after reconnect failed", e)
			})
		}
	}

	private onDriveEvent(id: bigint, inner: DriveEvent | null): void {
		this.lastEventId = id
		this.eventSeq++

		for (const rt of this.roots.values()) {
			if (rt.walking !== null || rt.mode === "pending" || rt.state === null) {
				this.hold(rt, id, inner)

				continue
			}

			if (rt.mode === "live") {
				const watermark = rt.state.watermark ?? -1n

				if (id <= watermark) {
					continue
				}

				if (id !== watermark + 1n || inner === null) {
					this.toStale(rt)
					void this.walk(rt, "auto")
				} else {
					this.setWatermark(rt, id)
				}
			} else if (inner === null) {
				void this.walk(rt, "auto")
			}

			if (inner !== null) {
				this.apply(rt, inner, null)
			}
		}
	}

	private hold(rt: RootRuntime, id: bigint, inner: DriveEvent | null): void {
		rt.buffer ??= []

		if (rt.buffer.length >= EVENT_BUFFER_CAP) {
			rt.bufferOverflow = true

			return
		}

		rt.buffer.push({
			id,
			inner
		})
	}

	// One event against one root, with its follow-ups. `walked` names directories a just-applied walk listed, so a
	// replayed move into one of them needs no re-listing.
	private apply(rt: RootRuntime, inner: DriveEvent, walked: ReadonlySet<string> | null): void {
		const state = rt.state

		if (state === null) {
			return
		}

		this.followAncestors(rt, inner)

		const result = applyDriveEvent({
			state,
			inner,
			now: this.deps.now(),
			driveRoot: rt.config.uuid === this.deps.driveRoot()
		})

		if (this.asyncPatches > 0) {
			this.patchLog.push({
				seq: this.eventSeq,
				root: rt.config.uuid,
				inner
			})
		}

		this.afterPatch(rt, result, walked)
	}

	private afterPatch(rt: RootRuntime, result: PatchResult, walked: ReadonlySet<string> | null): void {
		if (result.dirty) {
			this.markPersist(rt)
		}

		if (result.changed) {
			this.scheduleRebuild("patch")
		}

		const followUp = result.followUp

		switch (followUp.type) {
			case "none":
				return

			case "scopedWalk":
				if (walked?.has(followUp.dirUuid) !== true) {
					this.queueScoped(rt, followUp.dirUuid)
				}

				return

			case "fullWalk":
				void this.walk(rt, "event")

				return

			// Known to another root: placed there by the rename itself, nothing to look up.
			case "getFile":
				if (![...this.roots.values()].some(other => other.state?.byUuid.has(followUp.uuid) === true)) {
					this.lookUp(followUp.uuid)
				}

				return

			case "rootGone":
				rt.status = "missing"
				this.pushRoots()

				return

			case "rootRestored":
				rt.status = "ready"
				void this.walk(rt, "event")

				return

			case "rootMoved":
				rt.moved = true
				this.pushRoots()

				return
		}
	}

	// A directory above the root trashed or deleted takes the root with it (its index is kept, flagged); one
	// restored brings it back; one moved may have nested it under another root.
	private followAncestors(rt: RootRuntime, inner: DriveEvent): void {
		switch (inner.type) {
			case "folderTrash":
			case "folderDeletedPermanent":
				if (rt.ancestors.has(inner.uuid)) {
					rt.status = "missing"
					this.pushRoots()
				}

				return

			case "folderRestore":
				if (rt.ancestors.has(inner.dir.uuid) && rt.status === "missing") {
					rt.status = "ready"
					void this.walk(rt, "event")
				}

				return

			case "folderMove":
				if (rt.ancestors.has(inner.dir.uuid)) {
					rt.moved = true
					this.pushRoots()
				}

				return

			default:
				return
		}
	}

	// Starts an out-of-band patch: events applied until it lands are re-applied over its result.
	private beginAsyncPatch(): number {
		this.asyncPatches++

		return this.eventSeq
	}

	// `walked`: directories the patch re-listed, whose own move or restore needs no second listing.
	private endAsyncPatch(startSeq: number, root: RootRuntime | null, walked: ReadonlySet<string> | null): void {
		for (const entry of this.patchLog) {
			const rt = this.roots.get(entry.root)

			if (entry.seq > startSeq && rt !== undefined && (root === null || rt === root)) {
				this.afterPatch(
					rt,
					rt.state === null
						? {
								changed: false,
								dirty: false,
								followUp: {
									type: "none"
								}
							}
						: applyDriveEvent({
								state: rt.state,
								inner: entry.inner,
								now: this.deps.now(),
								driveRoot: rt.config.uuid === this.deps.driveRoot()
							}),
					walked
				)
			}
		}

		this.asyncPatches--

		if (this.asyncPatches === 0) {
			this.patchLog = []
		}
	}

	// A file renamed into a video name, located once and placed in whichever root holds it.
	private lookUp(uuid: string): void {
		if (this.lookups.has(uuid)) {
			return
		}

		const token = this.token
		const startSeq = this.beginAsyncPatch()
		const lookup = this.deps
			.getFile(uuid)
			.then(
				file => {
					if (token !== this.token || file === undefined) {
						return
					}

					for (const rt of this.roots.values()) {
						if (rt.state !== null && rt.walking === null && rt.mode !== "pending") {
							this.afterPatch(
								rt,
								applyLookedUpFile({
									state: rt.state,
									file,
									now: this.deps.now()
								}),
								null
							)
						}
					}
				},
				(e: unknown) => {
					log.warn("videos.index", "renamed file lookup failed", uuid, e)
				}
			)
			.finally(() => {
				this.lookups.delete(uuid)

				if (token === this.token) {
					this.endAsyncPatch(startSeq, null, null)
				}
			})

		this.lookups.set(uuid, lookup)
	}

	private queueScoped(rt: RootRuntime, dirUuid: string): void {
		rt.scoped.add(dirUuid)

		if (rt.scoped.size > SCOPED_WALK_MAX_PENDING) {
			rt.scoped.clear()

			if (rt.scopedTimer !== null) {
				clearTimeout(rt.scopedTimer)
				rt.scopedTimer = null
			}

			void this.walk(rt, "event")

			return
		}

		rt.scopedTimer ??= setTimeout(() => {
			rt.scopedTimer = null
			void this.runScoped(rt)
		}, SCOPED_WALK_COALESCE_MS)
	}

	private async runScoped(rt: RootRuntime): Promise<void> {
		const token = this.token
		const dirs = [...rt.scoped]

		rt.scoped.clear()

		// A walk under way or due covers them.
		if (rt.walking !== null || rt.mode === "pending" || rt.state === null) {
			return
		}

		const startSeq = this.beginAsyncPatch()

		try {
			for (const dirUuid of dirs) {
				const raw = await this.deps.walkDir(dirUuid)
				const state = rt.state

				if (token !== this.token || this.roots.get(rt.config.uuid) !== rt) {
					return
				}

				// Gone again, or moved out meanwhile.
				if (raw === null || !state.dirs.has(dirUuid)) {
					continue
				}

				if (applyScopedWalk({ state, walk: projectWalk(dirUuid, raw) })) {
					this.markPersist(rt)
					this.scheduleRebuild("patch")
				}
			}
		} catch (e) {
			log.warn("videos.index", "scoped directory walk failed; the root is walked instead", rt.config.uuid, e)

			this.toStale(rt)
			void this.walk(rt, "event")
		} finally {
			if (token === this.token) {
				this.endAsyncPatch(startSeq, rt, new Set(dirs))
			}
		}
	}

	// ── Classification ──

	private scheduleRebuild(kind: "now" | "patch", inputChanged = true): void {
		const now = this.deps.now()

		if (inputChanged) {
			this.inputSeq++
		}

		if (this.patchTimer !== null) {
			clearTimeout(this.patchTimer)
			this.patchTimer = null
		}

		if (kind === "now") {
			this.firstPatchAt = 0
			void this.rebuild()

			return
		}

		if (this.firstPatchAt === 0) {
			this.firstPatchAt = now
		}

		const delay = Math.max(0, Math.min(PATCH_DEBOUNCE_MS, this.firstPatchAt + PATCH_MAX_WAIT_MS - now))

		this.patchTimer = setTimeout(() => {
			this.patchTimer = null
			this.firstPatchAt = 0
			void this.rebuild()
		}, delay)
	}

	// Single flight: a change landing mid-run runs it once more after.
	private rebuild(): Promise<void> {
		if (this.rebuilding !== null) {
			this.rebuildAgain = true

			return this.rebuilding
		}

		this.rebuilding = this.runRebuild()
			.catch((e: unknown) => {
				log.error("videos.index", "classification failed", e)

				// No count is coming: the scans end with what is known.
				for (const rt of this.roots.values()) {
					rt.countedSeq = rt.walkSeq
				}

				this.pushRoots()
			})
			.finally(() => {
				this.rebuilding = null

				if (this.rebuildAgain) {
					this.rebuildAgain = false
					void this.rebuild()
				}
			})

		return this.rebuilding
	}

	private async runRebuild(): Promise<void> {
		const token = this.token
		const overrides = this.overrides
		const split = splitSignature(overrides)
		const cached = this.classified
		const input =
			cached !== null && cached.seq === this.inputSeq && cached.split === split
				? cached
				: await this.classify(this.inputSeq, overrides, split)

		if (input === null || token !== this.token) {
			return
		}

		this.classified = input

		const { library, files, loaded } = input
		const corrected = applyLibraryOverrides({
			library,
			overrides,
			stableUuidOf: file => files[file]?.stableUuid
		})
		const libraryFiles = new LibraryFiles({
			library: corrected.library,
			files,
			rootUuids: loaded.map(rt => rt.config.uuid),
			hidden: corrected.hidden,
			titleKeys: corrected.titleKeys,
			episodeKeys: corrected.episodeKeys,
			originals: corrected.originals,
			autoMerged: corrected.autoMerged
		})

		if (input !== cached) {
			this.countVideos(input)
		}

		this.publish(libraryFiles)
	}

	// The classifier's view of every loaded root, in slices; null when a reset came in meanwhile.
	private async classify(
		seq: number,
		overrides: Readonly<Record<string, LibraryOverride>>,
		split: string
	): Promise<ClassifiedInput | null> {
		const token = this.token
		const loaded = [...this.roots.values()].filter(rt => rt.state !== null)
		const seqs = loaded.map(rt => rt.walkSeq)
		const dirs: LibraryScanDir[] = []
		const files: IndexFile[] = []

		// Each root's files sit together in `files`, ending at its offset.
		const ends: number[] = []

		for (const rt of loaded) {
			for (const dir of rt.state?.dirs.values() ?? []) {
				dirs.push(dir)
			}

			for (const file of rt.state === null ? [] : liveFiles(rt.state)) {
				files.push(file)
			}

			ends.push(files.length)
		}

		const library = await drainSteps(
			classifyVideoLibrarySteps({
				roots: loaded.map(rt => ({
					uuid: rt.config.uuid,
					name: rt.state?.name ?? "",
					type: rt.config.type
				})),
				dirs,
				files,
				splitMergeKeys: splitMergeKeysOf(overrides),
				notTransportStream: this.notTransportStream
			}),
			this.deps.yieldSlice
		)

		return token === this.token ? { seq, split, library, files, loaded, seqs, ends } : null
	}

	private countVideos({ library, loaded, seqs, ends }: ClassifiedInput): void {
		const counts = new Array<number>(loaded.length).fill(0)

		for (const video of library.videos) {
			const root = ends.findIndex(end => video.file < end)

			counts[root] = (counts[root] ?? 0) + 1
		}

		loaded.forEach((rt, index) => {
			rt.videoCount = counts[index] ?? 0
			rt.countedSeq = seqs[index] ?? rt.countedSeq
		})
	}

	private publish(libraryFiles: LibraryFiles): void {
		const titles = buildTitles(libraryFiles)
		const reset = this.current === null
		const { upserts, removed } = diffTitles(this.signatures, titles)

		this.current = {
			files: libraryFiles,
			titles: new Map(titles.map(title => [title.id, title]))
		}

		if (reset || upserts.length > 0 || removed.length > 0) {
			this.emit({
				type: "titles",
				reset,
				upserts: reset ? titles : upserts,
				removed: reset ? [] : removed
			})
		}

		this.pushHome()
		this.pushRoots()
	}

	// ── Subscriptions and reads ──

	// The main thread's video list (home rows, grids, search): sent while a screen that reads it is mounted, all of
	// it first and then only what each classification changed. Off, nothing of it is held here.
	public setHomeEnabled(enabled: boolean): void {
		this.homeSignatures = enabled ? new Map() : null
		this.homeSent = false
		this.pushHome()
	}

	private pushHome(): void {
		const signatures = this.homeSignatures

		if (signatures === null || this.current === null) {
			return
		}

		const items = this.homeItems()
		const reset = !this.homeSent
		const { upserts, removed } = diffRecords(signatures, items, item => item.stableUuid)

		this.homeSent = true

		if (reset || upserts.length > 0 || removed.length > 0) {
			this.emit({
				type: "home",
				reset,
				upserts: reset ? items : upserts,
				removed: reset ? [] : removed
			})
		}
	}

	// The listener gets everything known now, then every change.
	public subscribe(listener: (update: VideoIndexUpdate) => void): () => void {
		this.listeners.add(listener)

		if (this.current !== null) {
			listener({
				type: "titles",
				reset: true,
				upserts: [...this.current.titles.values()],
				removed: []
			})
		}

		if (this.homeSignatures !== null && this.current !== null && this.homeSent) {
			listener({
				type: "home",
				reset: true,
				upserts: this.homeItems(),
				removed: []
			})
		}

		listener({
			type: "roots",
			roots: this.rootStatuses()
		})

		return () => {
			this.listeners.delete(listener)
		}
	}

	private emit(update: VideoIndexUpdate): void {
		for (const listener of this.listeners) {
			listener(update)
		}
	}

	private pushRoots(): void {
		this.emit({
			type: "roots",
			roots: this.rootStatuses()
		})
	}

	public rootStatuses(): VideoRootStatus[] {
		return [...this.roots.values()].map(rt => {
			const state = rt.state

			return {
				uuid: rt.config.uuid,
				name: state?.name ?? "",
				type: rt.config.type,
				state: rt.status,
				// Until the walk's result is classified too, so the count read when it ends is the new one.
				walking: (rt.walking !== null && !rt.stopped) || rt.countedSeq !== rt.walkSeq,
				current: rt.mode === "live",
				completeAt: state?.completeAt ?? 0,
				walkedEntries: state?.walkedEntries ?? 0,
				size: rootSize(state?.walkedEntries ?? 0),
				videos: rt.videoCount,
				moved: rt.moved,
				incomplete: rt.incomplete
			}
		})
	}

	public titleArt(id: string, slot: ArtSlotCode): TitleArtCandidates | null {
		const current = this.current
		const title = current?.titles.get(id)

		return current === null || title === undefined ? null : titleArtCandidates(current.files, title, slot)
	}

	// One video's art: an episode's own still, else its title's.
	public itemArt(titleId: string, stableUuid: string, slot: ArtSlotCode): TitleArtCandidates | null {
		const current = this.current
		const title = current?.titles.get(titleId)

		return current === null || title === undefined ? null : itemArtCandidates(current.files, title, stableUuid, slot)
	}

	// Every playable video of the classification.
	public homeItems(): VideoHomeItem[] {
		const current = this.current

		return current === null ? [] : buildHomeItems(current.files, parent => this.dirName(parent))
	}

	private dirName(uuid: string): string {
		for (const rt of this.roots.values()) {
			const state = rt.state

			if (state === null) {
				continue
			}

			if (rt.config.uuid === uuid) {
				return state.name
			}

			const dir = state.dirs.get(uuid)

			if (dir !== undefined) {
				return dir.name
			}
		}

		return ""
	}

	public titleDetail(id: string): VideoTitleDetail | null {
		const current = this.current
		const title = current?.titles.get(id)

		return current === null || title === undefined ? null : buildTitleDetail(current.files, title)
	}

	// Playable drive files rebuilt from the index, for the ones it holds live. These carry key material and are
	// built only on request, a few at a time.
	public playableFiles(stableUuids: readonly string[]): File[] {
		const out: File[] = []

		for (const stableUuid of stableUuids) {
			for (const rt of this.roots.values()) {
				const file = rt.state?.files.get(stableUuid)

				if (file?.tombstoneAt === 0) {
					const built = toSdkFile(file)

					if (built !== null) {
						out.push(built)
					}

					break
				}
			}
		}

		return out
	}

	// ── Persistence ──

	// Only the storage leader tab writes snapshots; the others keep theirs in memory.
	// Whether this tab writes snapshots: on while it has the library open, off once a sign-out starts (here or in
	// another tab), which also abandons a write under way.
	public setWriter(writer: boolean): void {
		this.writer = writer && this.writeFailures < MAX_WRITE_FAILURES

		if (!this.writer) {
			this.writeEpoch++

			for (const rt of this.roots.values()) {
				if (rt.writeTimer !== null) {
					clearTimeout(rt.writeTimer)
					rt.writeTimer = null
				}
			}

			return
		}

		for (const rt of this.roots.values()) {
			this.scheduleWrite(rt)
		}
	}

	private markPersist(rt: RootRuntime): void {
		rt.version++
		this.scheduleWrite(rt)
	}

	private scheduleWrite(rt: RootRuntime): void {
		if (!this.writer || rt.state === null || rt.version === rt.writtenVersion || rt.writeTimer !== null) {
			return
		}

		const delay = Math.max(WRITE_DEBOUNCE_MS, rt.lastWriteAt + WRITE_MIN_INTERVAL_MS - this.deps.now())

		rt.writeTimer = setTimeout(() => {
			rt.writeTimer = null
			void this.writeNow(rt)
		}, delay)
	}

	private writeNow(rt: RootRuntime): Promise<void> {
		if (rt.writing !== null) {
			return rt.writing.then(() => this.writeNow(rt))
		}

		const token = this.token
		const writing = this.runWrite(rt).finally(() => {
			rt.writing = null

			if (token === this.token && this.roots.get(rt.config.uuid) === rt) {
				this.scheduleWrite(rt)
			}
		})

		rt.writing = writing

		return writing
	}

	private async runWrite(rt: RootRuntime): Promise<void> {
		const state = rt.state
		const uuid = rt.config.uuid

		if (state === null || !this.writer || (rt.version === rt.writtenVersion && !rt.watermarkDirty)) {
			return
		}

		const version = rt.version
		const token = this.token
		const epoch = this.writeEpoch
		const cancelled = (): boolean => token !== this.token || epoch !== this.writeEpoch

		rt.watermarkDirty = false

		try {
			const { slot, generation } = nextSlot(await this.deps.store.readGenerations(uuid))

			if (cancelled()) {
				return
			}

			state.generation = generation

			await this.deps.store.write(uuid, slot, encodeSnapshot(state), cancelled)

			if (!cancelled()) {
				rt.writtenVersion = version
				rt.lastWriteAt = this.deps.now()
				this.writeFailures = 0
			}
		} catch (e) {
			rt.watermarkDirty = true
			rt.lastWriteAt = this.deps.now()
			this.writeFailures++

			// Storage that keeps failing here (no OPFS write support, quota) isn't tried for the rest of the session.
			if (this.writeFailures >= MAX_WRITE_FAILURES) {
				log.warn("videos.index", "snapshot writes keep failing; the index stays in memory this session", uuid, e)
				this.setWriter(false)
			} else {
				log.warn("videos.index", "snapshot write failed", uuid, e)
			}
		}
	}

	// Page hidden or closing: every pending snapshot write now, watermark-only changes included.
	public async flush(): Promise<void> {
		if (!this.writer) {
			return
		}

		const writes: Promise<void>[] = []

		for (const rt of this.roots.values()) {
			if (rt.writeTimer !== null) {
				clearTimeout(rt.writeTimer)
				rt.writeTimer = null
			}

			if (rt.state !== null && (rt.version !== rt.writtenVersion || rt.watermarkDirty)) {
				writes.push(this.writeNow(rt))
			}
		}

		await Promise.all(writes)
	}

	// Logout: everything of the signed-out account goes; work under way drops its result.
	public reset(): void {
		this.token++

		for (const rt of this.roots.values()) {
			this.clearTimers(rt)
		}

		if (this.patchTimer !== null) {
			clearTimeout(this.patchTimer)
			this.patchTimer = null
		}

		this.roots.clear()
		this.listeners.clear()
		this.signatures.clear()
		this.homeSignatures = null
		this.homeSent = false
		this.lookups.clear()
		this.patchLog = []
		this.asyncPatches = 0
		this.current = null
		this.writer = false
		this.writeEpoch++
		this.lastEventId = null
		this.socketLive = false
		this.sawReconnecting = false
		this.rebuildAgain = false
		this.firstPatchAt = 0
		this.overrides = {}
		this.overridesSignature = "{}"
		this.classified = null
	}
}

// The split corrections alone: the only ones the classifier itself reads.
function splitSignature(overrides: Readonly<Record<string, LibraryOverride>>): string {
	return [...splitMergeKeysOf(overrides)].sort().join("\n")
}

function toSdkFile(file: IndexFile): File | null {
	if (!isUuidStr(file.uuid) || !isUuidStr(file.stableUuid) || !isUuidStr(file.parent)) {
		return null
	}

	return {
		uuid: file.uuid,
		stableUUID: file.stableUuid,
		parent: file.parent,
		size: BigInt(file.size),
		favorited: file.favorited,
		region: file.region,
		bucket: file.bucket,
		timestamp: BigInt(file.timestamp),
		chunks: BigInt(file.chunks),
		canMakeThumbnail: file.canMakeThumbnail,
		meta: {
			type: "decoded",
			data: {
				name: file.name,
				mime: file.mime,
				...(Number.isNaN(file.created) ? {} : { created: BigInt(file.created) }),
				modified: BigInt(file.modified),
				size: BigInt(file.metaSize),
				key: file.key,
				version: file.keyVersion
			}
		}
	}
}
