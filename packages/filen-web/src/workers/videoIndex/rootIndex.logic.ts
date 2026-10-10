import { extensionDot, videoFileKind, type VideoFileKind } from "@filen/shared/videos"
import type { DirMeta, DirsAndFiles, DriveEvent, File, FileEncryptionVersion, FileMeta } from "@filen/sdk-rs"
import { DOT_FILEN_DIR_NAME } from "@/lib/dotFilen.logic"

// One library root's slim index: every directory below it and every video-related file (videos, sidecars,
// `.ignore` markers), keyed by stable uuid. Holds what classification needs and what rebuilding a playable
// File needs, nothing else. Pure: walks and socket events go in, effects come out.

export interface IndexDir {
	uuid: string
	parent: string
	// Empty for a directory whose name can't be decrypted.
	name: string
}

export interface IndexFile {
	uuid: string
	stableUuid: string
	parent: string
	name: string
	size: number
	metaSize: number
	chunks: number
	timestamp: number
	// NaN when the metadata carries none.
	created: number
	modified: number
	// When the file joined the library; never changes for a known stable uuid.
	dateAdded: number
	// 0 for a live file; else when it went missing. A tombstone is hidden and keeps dateAdded.
	tombstoneAt: number
	// Complete walks that confirmed a tombstoned file missing.
	tombstoneWalks: number
	key: string
	keyVersion: FileEncryptionVersion
	mime: string
	region: string
	bucket: string
	canMakeThumbnail: boolean
	favorited: boolean
}

export interface RootState {
	uuid: string
	name: string
	dirs: Map<string, IndexDir>
	// By stable uuid, tombstones included.
	files: Map<string, IndexFile>
	// Head uuid -> stable uuid, live files only.
	byUuid: Map<string, string>
	// Drive event id the index is current to; null when unknown.
	watermark: bigint | null
	generation: bigint
	// Start of the latest complete walk (0 before the first) and of the one before it.
	completeAt: number
	prevCompleteAt: number
	// Directories and files the latest walk listed, all types: the walk's size, for the large-root warning.
	walkedEntries: number
}

export const DAY_MS = 86_400_000
// A tombstone is purged once it is this old and two complete walks confirmed it missing.
export const TOMBSTONE_MS = 30 * DAY_MS
const TOMBSTONE_PURGE_WALKS = 2

export function emptyRootState(uuid: string, name: string): RootState {
	return {
		uuid,
		name,
		dirs: new Map(),
		files: new Map(),
		byUuid: new Map(),
		watermark: null,
		generation: 0n,
		completeAt: 0,
		prevCompleteAt: 0,
		walkedEntries: 0
	}
}

export function dirName(meta: DirMeta): string {
	return meta.type === "decoded" ? meta.data.name : ""
}

// The library kind of a file name, or null for a file the index doesn't keep. `.ignore` markers are kept for the
// classifier; their kind is null too.
export function indexedKind(name: string, size: number): VideoFileKind | "marker" | null {
	if (name === ".ignore") {
		return "marker"
	}

	const dot = extensionDot(name)

	if (dot < 0) {
		return null
	}

	return videoFileKind(name.slice(dot + 1).toLowerCase(), size, name.slice(0, dot).toLowerCase())
}

function isPlayableKind(kind: VideoFileKind | "marker" | null): boolean {
	return kind === "video" || kind === "unsupportedVideo"
}

// The indexed fields of a decrypted drive file, or null for one the index doesn't keep (undecryptable or not
// video-related). `dateAdded` and tombstone fields are left for the caller.
function slimFile(file: File, meta: FileMeta): Omit<IndexFile, "dateAdded" | "tombstoneAt" | "tombstoneWalks"> | null {
	if (meta.type !== "decoded") {
		return null
	}

	const data = meta.data
	const size = Number(file.size)

	if (indexedKind(data.name, size) === null) {
		return null
	}

	return {
		uuid: file.uuid,
		stableUuid: file.stableUUID ?? file.uuid,
		parent: file.parent,
		name: data.name,
		size,
		metaSize: Number(data.size),
		chunks: Number(file.chunks),
		timestamp: Number(file.timestamp),
		created: data.created === undefined ? Number.NaN : Number(data.created),
		modified: Number(data.modified),
		key: data.key,
		keyVersion: data.version,
		mime: data.mime,
		region: file.region,
		bucket: file.bucket,
		canMakeThumbnail: file.canMakeThumbnail,
		favorited: file.favorited
	}
}

export interface WalkProjection {
	dirs: IndexDir[]
	files: Omit<IndexFile, "dateAdded" | "tombstoneAt" | "tombstoneWalks">[]
	entries: number
	// Whether every directory and kept file hangs off the root through listed directories.
	complete: boolean
}

// Filen's own directory directly below the drive root, when the root is the drive root: never part of a library.
export function isSystemDir({
	parent,
	meta,
	rootUuid,
	driveRoot
}: {
	parent: string
	meta: DirMeta
	rootUuid: string
	driveRoot: boolean
}): boolean {
	return driveRoot && parent === rootUuid && meta.type === "decoded" && meta.data.name.toLowerCase() === DOT_FILEN_DIR_NAME
}

// The system directory's subtree in a walk of the drive root, to leave out.
function systemSubtree(rootUuid: string, raw: DirsAndFiles): Set<string> {
	const excluded = new Set<string>()
	const children = new Map<string, string[]>()

	for (const dir of raw.dirs) {
		if (dir.type !== "normal") {
			continue
		}

		if (isSystemDir({ parent: dir.parent, meta: dir.meta, rootUuid, driveRoot: true })) {
			excluded.add(dir.uuid)
		}

		const list = children.get(dir.parent)

		if (list === undefined) {
			children.set(dir.parent, [dir.uuid])
		} else {
			list.push(dir.uuid)
		}
	}

	const stack = [...excluded]

	for (let next = stack.pop(); next !== undefined; next = stack.pop()) {
		for (const child of children.get(next) ?? []) {
			if (!excluded.has(child)) {
				excluded.add(child)
				stack.push(child)
			}
		}
	}

	return excluded
}

// Keeps every directory (socket events need them to place a file) and only the video-related files. Raw files,
// which hold key material, never leave the walk's caller. A walk of the drive root leaves Filen's own directory out.
export function projectWalk(rootUuid: string, raw: DirsAndFiles, driveRoot = false): WalkProjection {
	const dirs: IndexDir[] = []
	const known = new Set<string>([rootUuid])
	const excluded = driveRoot ? systemSubtree(rootUuid, raw) : null

	for (const dir of raw.dirs) {
		if (dir.type !== "normal" || excluded?.has(dir.uuid) === true) {
			continue
		}

		dirs.push({
			uuid: dir.uuid,
			parent: dir.parent,
			name: dirName(dir.meta)
		})
		known.add(dir.uuid)
	}

	let complete = true

	for (const dir of dirs) {
		if (!known.has(dir.parent)) {
			complete = false
		}
	}

	const files: WalkProjection["files"] = []

	for (const file of raw.files) {
		const slim = slimFile(file, file.meta)

		if (slim === null || excluded?.has(slim.parent) === true) {
			continue
		}

		if (!known.has(slim.parent)) {
			complete = false
		}

		files.push(slim)
	}

	return {
		dirs,
		files,
		entries: raw.dirs.length + raw.files.length,
		complete
	}
}

function rebuildByUuid(state: RootState): void {
	state.byUuid.clear()

	for (const file of state.files.values()) {
		if (file.tombstoneAt === 0) {
			state.byUuid.set(file.uuid, file.stableUuid)
		}
	}
}

// A file the index hasn't held before joins at the latest of its own timestamp and the previous complete walk: a
// file moved or copied in with an old timestamp counts as added when the library first could have seen it.
function joinedAt(timestamp: number, state: RootState): number {
	return Math.max(timestamp, state.completeAt)
}

// Upserts by stable uuid. A known one (a tombstone included) keeps its dateAdded.
function upsertFile(state: RootState, slim: Omit<IndexFile, "dateAdded" | "tombstoneAt" | "tombstoneWalks">): void {
	const existing = state.files.get(slim.stableUuid)
	const index = treeIndexes.get(state)

	if (index !== undefined && existing?.parent !== slim.parent) {
		if (existing !== undefined) {
			index.files.get(existing.parent)?.delete(slim.stableUuid)
		}

		addTo(index.files, slim.parent, slim.stableUuid)
	}

	if (existing?.tombstoneAt === 0) {
		state.byUuid.delete(existing.uuid)
	}

	state.files.set(slim.stableUuid, {
		...slim,
		dateAdded: existing?.dateAdded ?? joinedAt(slim.timestamp, state),
		tombstoneAt: 0,
		tombstoneWalks: 0
	})
	state.byUuid.set(slim.uuid, slim.stableUuid)
}

function tombstone(state: RootState, file: IndexFile, now: number): void {
	if (file.tombstoneAt !== 0) {
		return
	}

	state.byUuid.delete(file.uuid)
	file.tombstoneAt = now
	file.tombstoneWalks = 0
}

function removeFile(state: RootState, file: IndexFile): void {
	treeIndexes.get(state)?.files.get(file.parent)?.delete(file.stableUuid)
	state.files.delete(file.stableUuid)

	if (state.byUuid.get(file.uuid) === file.stableUuid) {
		state.byUuid.delete(file.uuid)
	}
}

// Applies a walk's result. A complete walk is the authority on presence: directories it didn't list go, files it
// didn't list become tombstones, and old tombstones it confirms missing again are purged. A partial one only
// inserts and updates.
export function applyWalk({
	state,
	walk,
	startedAt,
	now
}: {
	state: RootState
	walk: WalkProjection
	startedAt: number
	now: number
}): void {
	forgetChildIndex(state)

	if (walk.complete) {
		state.dirs.clear()
	}

	for (const dir of walk.dirs) {
		state.dirs.set(dir.uuid, dir)
	}

	const seen = walk.complete ? new Set<string>() : null

	for (const slim of walk.files) {
		upsertFile(state, slim)
		seen?.add(slim.stableUuid)
	}

	state.walkedEntries = walk.entries

	if (seen === null) {
		return
	}

	for (const file of [...state.files.values()]) {
		if (seen.has(file.stableUuid)) {
			continue
		}

		if (file.tombstoneAt === 0) {
			tombstone(state, file, now)
			file.tombstoneWalks = 1

			continue
		}

		file.tombstoneWalks = Math.min(255, file.tombstoneWalks + 1)

		if (file.tombstoneWalks >= TOMBSTONE_PURGE_WALKS && now - file.tombstoneAt >= TOMBSTONE_MS) {
			removeFile(state, file)
		}
	}

	state.prevCompleteAt = state.completeAt
	state.completeAt = startedAt
	rebuildByUuid(state)
}

// What the engine does after a patch: nothing more, re-list a directory that came into the root, walk the whole
// root, look a file up by uuid, or flag the root itself.
export type PatchFollowUp =
	| { type: "none" }
	| { type: "scopedWalk"; dirUuid: string }
	| { type: "fullWalk" }
	| { type: "getFile"; uuid: string }
	| { type: "rootGone" }
	| { type: "rootRestored" }
	| { type: "rootMoved" }

export interface PatchResult {
	// The classification input changed.
	changed: boolean
	// Something persisted changed (also flags and other fields classification ignores).
	dirty: boolean
	followUp: PatchFollowUp
}

const NONE: PatchResult = {
	changed: false,
	dirty: false,
	followUp: {
		type: "none"
	}
}

const CHANGED: PatchResult = {
	changed: true,
	dirty: true,
	followUp: {
		type: "none"
	}
}

// Persisted, but nothing the classification reads: a directory change with no live file below it.
const DIRTY: PatchResult = {
	changed: false,
	dirty: true,
	followUp: {
		type: "none"
	}
}

function followUp(next: PatchFollowUp, changed = false): PatchResult {
	return {
		changed,
		dirty: changed,
		followUp: next
	}
}

export function inRoot(state: RootState, parent: string): boolean {
	return parent === state.uuid || state.dirs.has(parent)
}

function headOf(state: RootState, uuid: string): IndexFile | undefined {
	const stable = state.byUuid.get(uuid)

	return stable === undefined ? undefined : state.files.get(stable)
}

// Each directory's child directories and files (stable uuids, tombstones included), built when a subtree is first
// asked for and then kept current in place by every patch, so a burst of folder moves, renames or trashes costs one
// build and then only each subtree. A walk applied replaces it.
interface TreeIndex {
	children: Map<string, Set<string>>
	files: Map<string, Set<string>>
}

const treeIndexes = new WeakMap<RootState, TreeIndex>()

function addTo(map: Map<string, Set<string>>, key: string, value: string): void {
	const set = map.get(key)

	if (set === undefined) {
		map.set(key, new Set([value]))
	} else {
		set.add(value)
	}
}

function treeIndex(state: RootState): TreeIndex {
	let index = treeIndexes.get(state)

	if (index === undefined) {
		index = { children: new Map(), files: new Map() }

		for (const dir of state.dirs.values()) {
			addTo(index.children, dir.parent, dir.uuid)
		}

		for (const file of state.files.values()) {
			addTo(index.files, file.parent, file.stableUuid)
		}

		treeIndexes.set(state, index)
	}

	return index
}

function forgetChildIndex(state: RootState): void {
	treeIndexes.delete(state)
}

// Directories at and below `dirUuid`. A parent cycle (events replayed out of order) still ends.
function subtree(state: RootState, dirUuid: string): Set<string> {
	const { children } = treeIndex(state)
	const out = new Set<string>([dirUuid])
	const stack = [dirUuid]

	for (let next = stack.pop(); next !== undefined; next = stack.pop()) {
		for (const child of children.get(next) ?? []) {
			if (!out.has(child)) {
				out.add(child)
				stack.push(child)
			}
		}
	}

	return out
}

// Whether `parent` is `dirUuid` or lies below it: placing `dirUuid` there would close a cycle. Walks up from
// `parent`, so it costs the depth, not the subtree.
function wouldCycle(state: RootState, dirUuid: string, parent: string): boolean {
	let current: string | undefined = parent

	for (let steps = 0; current !== undefined && current !== state.uuid && steps <= state.dirs.size; steps++) {
		if (current === dirUuid) {
			return true
		}

		current = state.dirs.get(current)?.parent
	}

	return current === dirUuid
}

// Sets a directory, keeping the child index current.
function placeDir(state: RootState, dir: IndexDir): void {
	const index = treeIndexes.get(state)

	if (index !== undefined) {
		const previous = state.dirs.get(dir.uuid)

		if (previous !== undefined && previous.parent !== dir.parent) {
			index.children.get(previous.parent)?.delete(dir.uuid)
		}

		addTo(index.children, dir.parent, dir.uuid)
	}

	state.dirs.set(dir.uuid, dir)
}

// Events applied out of order (replayed over a newer walk, or after a gap) asked for a directory to go below
// itself: the index can't be trusted, so the root is walked and the event changes nothing.
const CYCLE: PatchResult = {
	changed: false,
	dirty: false,
	followUp: {
		type: "fullWalk"
	}
}

// The files lying directly in one of `dirs`.
function* filesIn(state: RootState, dirs: ReadonlySet<string>): Generator<IndexFile> {
	const { files } = treeIndex(state)

	for (const dir of dirs) {
		for (const stable of files.get(dir) ?? []) {
			const file = state.files.get(stable)

			if (file?.parent === dir) {
				yield file
			}
		}
	}
}

// Whether a live file lies directly in one of `dirs`.
function liveFileIn(state: RootState, dirs: ReadonlySet<string>): boolean {
	for (const file of filesIn(state, dirs)) {
		if (file.tombstoneAt === 0) {
			return true
		}
	}

	return false
}

// A directory left the root (moved out, trashed, deleted): its directories go and its files become tombstones,
// which a later return revives with their dateAdded. The classification changes only when a live file went.
function dropSubtree(state: RootState, dirUuid: string, now: number): PatchResult {
	const dirs = subtree(state, dirUuid)
	const { children } = treeIndex(state)
	const parent = state.dirs.get(dirUuid)?.parent
	let changed = false

	// Tombstoning keeps each file where it lies, so the index stays valid.
	for (const file of filesIn(state, dirs)) {
		if (file.tombstoneAt === 0) {
			tombstone(state, file, now)
			changed = true
		}
	}

	for (const uuid of dirs) {
		state.dirs.delete(uuid)
		children.delete(uuid)
	}

	if (parent !== undefined) {
		children.get(parent)?.delete(dirUuid)
	}

	return changed ? CHANGED : DIRTY
}

function upsertFromEvent(state: RootState, file: File, now: number): PatchResult {
	const slim = slimFile(file, file.meta)
	const stable = file.stableUUID ?? file.uuid
	const existing = state.files.get(stable)

	if (slim !== null && inRoot(state, file.parent)) {
		upsertFile(state, slim)

		return CHANGED
	}

	// Moved out of the root, renamed past recognition or undecryptable now.
	if (existing?.tombstoneAt === 0) {
		tombstone(state, existing, now)

		return CHANGED
	}

	return NONE
}

// One drive event against one root. Ids and ordering are the engine's business; this only maps an event's
// payload onto the index. `driveRoot`: the root is the drive root, whose system directory stays out.
export function applyDriveEvent({
	state,
	inner,
	now,
	driveRoot = false
}: {
	state: RootState
	inner: DriveEvent
	now: number
	driveRoot?: boolean
}): PatchResult {
	switch (inner.type) {
		// A version restore (fileArchiveRestored) makes an older uuid the head again; the stable uuid, and so
		// dateAdded, stays.
		case "fileNew":
		case "fileRestore":
		case "fileMove":
		case "fileArchiveRestored":
			return upsertFromEvent(state, inner.file, now)

		case "fileTrash": {
			// With newUUID it retires a version on a versioning-disabled account; the successor's fileNew follows.
			const head = inner.newUUID === undefined ? headOf(state, inner.uuid) : undefined

			if (head === undefined) {
				return NONE
			}

			tombstone(state, head, now)

			return CHANGED
		}

		case "fileArchived": {
			// Without newUUID the lineage was replaced by another file moved over it: gone for good.
			const head = inner.newUUID === undefined ? headOf(state, inner.uuid) : undefined

			if (head === undefined) {
				return NONE
			}

			removeFile(state, head)

			return CHANGED
		}

		case "fileDeletedPermanent": {
			// The stable uuid is present for version-only deletes too: only a delete of the indexed head removes.
			const head = headOf(state, inner.uuid)

			if (head === undefined) {
				return NONE
			}

			removeFile(state, head)

			return CHANGED
		}

		case "fileMetadataChanged": {
			const head = headOf(state, inner.uuid)

			if (head === undefined) {
				// A file renamed into a video name may belong to the root; only a lookup tells where it lives.
				if (
					inner.metadata.type === "decoded" &&
					isPlayableKind(indexedKind(inner.metadata.data.name, Number(inner.metadata.data.size)))
				) {
					return followUp({
						type: "getFile",
						uuid: inner.uuid
					})
				}

				return NONE
			}

			if (inner.metadata.type !== "decoded" || indexedKind(inner.metadata.data.name, head.size) === null) {
				tombstone(state, head, now)

				return CHANGED
			}

			const data = inner.metadata.data

			head.name = data.name
			head.mime = data.mime
			head.key = data.key
			head.keyVersion = data.version
			head.modified = Number(data.modified)
			head.created = data.created === undefined ? Number.NaN : Number(data.created)

			return CHANGED
		}

		case "folderSubCreated":
			if (
				!inRoot(state, inner.dir.parent) ||
				isSystemDir({ parent: inner.dir.parent, meta: inner.dir.meta, rootUuid: state.uuid, driveRoot })
			) {
				return NONE
			}

			if (wouldCycle(state, inner.dir.uuid, inner.dir.parent)) {
				return CYCLE
			}

			placeDir(state, {
				uuid: inner.dir.uuid,
				parent: inner.dir.parent,
				name: dirName(inner.dir.meta)
			})

			// New, so empty: what it will hold comes with events of its own.
			return DIRTY

		case "folderMove": {
			const dir = inner.dir

			if (dir.uuid === state.uuid) {
				return followUp({
					type: "rootMoved"
				})
			}

			const known = state.dirs.get(dir.uuid)
			const system = isSystemDir({ parent: dir.parent, meta: dir.meta, rootUuid: state.uuid, driveRoot })

			if (inRoot(state, dir.parent) && !system) {
				if (wouldCycle(state, dir.uuid, dir.parent)) {
					return CYCLE
				}

				if (known !== undefined) {
					const holdsFiles = liveFileIn(state, subtree(state, dir.uuid))

					placeDir(state, {
						uuid: dir.uuid,
						parent: dir.parent,
						name: dirName(dir.meta)
					})

					return holdsFiles ? CHANGED : DIRTY
				}

				placeDir(state, {
					uuid: dir.uuid,
					parent: dir.parent,
					name: dirName(dir.meta)
				})

				return followUp(
					{
						type: "scopedWalk",
						dirUuid: dir.uuid
					},
					true
				)
			}

			return known === undefined ? NONE : dropSubtree(state, dir.uuid, now)
		}

		case "folderTrash":
		case "folderDeletedPermanent":
			if (inner.uuid === state.uuid) {
				return followUp({
					type: "rootGone"
				})
			}

			// Files below a permanently deleted directory stay resolvable for a while with a trash parent; they
			// are gone all the same.
			return state.dirs.has(inner.uuid) ? dropSubtree(state, inner.uuid, now) : NONE

		case "folderRestore": {
			const dir = inner.dir

			if (dir.uuid === state.uuid) {
				return followUp({
					type: "rootRestored"
				})
			}

			if (!inRoot(state, dir.parent) || isSystemDir({ parent: dir.parent, meta: dir.meta, rootUuid: state.uuid, driveRoot })) {
				return NONE
			}

			if (wouldCycle(state, dir.uuid, dir.parent)) {
				return CYCLE
			}

			placeDir(state, {
				uuid: dir.uuid,
				parent: dir.parent,
				name: dirName(dir.meta)
			})

			return followUp(
				{
					type: "scopedWalk",
					dirUuid: dir.uuid
				},
				true
			)
		}

		case "folderMetadataChanged": {
			if (inner.uuid === state.uuid) {
				state.name = dirName(inner.meta)

				return CHANGED
			}

			const known = state.dirs.get(inner.uuid)

			if (known === undefined) {
				return NONE
			}

			// Renamed into Filen's own directory name: it is that directory now.
			if (isSystemDir({ parent: known.parent, meta: inner.meta, rootUuid: state.uuid, driveRoot })) {
				return dropSubtree(state, known.uuid, now)
			}

			known.name = dirName(inner.meta)

			return liveFileIn(state, subtree(state, known.uuid)) ? CHANGED : DIRTY
		}

		case "itemFavorite": {
			const item = inner.item

			if (item.type !== "file") {
				return NONE
			}

			const head = headOf(state, item.uuid)

			if (head === undefined || head.favorited === item.favorited) {
				return NONE
			}

			head.favorited = item.favorited

			return {
				changed: false,
				dirty: true,
				followUp: {
					type: "none"
				}
			}
		}

		// The whole drive was emptied: nothing indexed survives, and a walk confirms what (if anything) is there now.
		case "deleteAll":
			forgetChildIndex(state)
			state.dirs.clear()
			state.files.clear()
			state.byUuid.clear()

			return followUp(
				{
					type: "fullWalk"
				},
				true
			)

		case "folderColorChanged":
		case "trashEmpty":
		case "deleteVersioned":
			return NONE
	}
}

// A file looked up after a rename event, placed if it lives in the root.
export function applyLookedUpFile({ state, file, now }: { state: RootState; file: File; now: number }): PatchResult {
	return upsertFromEvent(state, file, now)
}

// Applies a scoped walk of a directory that came into the root: inserts and updates only.
export function applyScopedWalk({ state, walk }: { state: RootState; walk: WalkProjection }): boolean | "cycle" {
	let changed = false

	for (const dir of walk.dirs) {
		placeDir(state, dir)
		changed = true
	}

	// The listing is consistent in itself; with what the index held of the tree meanwhile it may close a cycle, which
	// only a walk of the root sorts out.
	for (const dir of walk.dirs) {
		if (wouldCycle(state, dir.uuid, dir.parent)) {
			return "cycle"
		}
	}

	for (const slim of walk.files) {
		if (inRoot(state, slim.parent)) {
			upsertFile(state, slim)
			changed = true
		}
	}

	return changed
}

// Live (non-tombstoned) files.
// Records materialized between yields of fromRecordsSteps.
const RECORD_STEP = 4096

export function fromRecords(records: {
	header: Omit<RootState, "dirs" | "files" | "byUuid">
	dirs: readonly IndexDir[]
	files: readonly IndexFile[]
}): RootState {
	const steps = fromRecordsSteps(records)

	for (;;) {
		const step = steps.next()

		if (step.done === true) {
			return step.value
		}
	}
}

// fromRecords in steps (yields between them), for a large root loaded while the worker serves other calls.
export function* fromRecordsSteps({
	header,
	dirs,
	files
}: {
	header: Omit<RootState, "dirs" | "files" | "byUuid">
	dirs: readonly IndexDir[]
	files: readonly IndexFile[]
}): Generator<void, RootState> {
	const state: RootState = {
		...header,
		dirs: new Map(),
		files: new Map(),
		byUuid: new Map()
	}

	let step = 0

	for (const dir of dirs) {
		state.dirs.set(dir.uuid, dir)

		if (++step % RECORD_STEP === 0) {
			yield
		}
	}

	for (const file of files) {
		state.files.set(file.stableUuid, file)

		if (++step % RECORD_STEP === 0) {
			yield
		}
	}

	for (const file of state.files.values()) {
		if (file.tombstoneAt === 0) {
			state.byUuid.set(file.uuid, file.stableUuid)
		}

		if (++step % RECORD_STEP === 0) {
			yield
		}
	}

	return state
}
