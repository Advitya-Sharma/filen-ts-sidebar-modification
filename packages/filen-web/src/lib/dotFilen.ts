import type { DirMeta, SocketEvent } from "@filen/sdk-rs"
import { sdkApi } from "@/lib/sdk/client"
import { runOp } from "@/lib/actions/outcome"
import { registerSocketHandler } from "@/lib/sdk/socket"
import { narrowItem } from "@/features/drive/lib/item"
import { isPhotoItem } from "@/features/photos/lib/predicate"
import { DOT_FILEN_DIR_NAME } from "@/lib/dotFilen.logic"

// The app's own directories under `.filen` at the drive root (Playlists, Chat Uploads, Videos), resolved once
// per session and shared by every feature. createDirectory is find-or-create, idempotent and case-insensitive
// server-side, so a lookup is one call per level and needs no listing. A read that must not create anything finds
// instead: one existence check per level, then the same find-or-create, which returns the existing directory.

export { DOT_FILEN_DIR_NAME } from "@/lib/dotFilen.logic"

// A lookup in flight can't match a directory event by uuid yet, so any event that could concern it counts while
// it is pending. Forgetting bumps the generation: a lookup started under an older one still answers its caller
// but isn't remembered.
interface Memo {
	// Resolved or being resolved, found or created.
	promise: Promise<string> | null
	// A find in flight; a find that comes back empty leaves nothing behind, so the next find or create asks again.
	finding: Promise<string | null> | null
	uuid: string | null
	pending: number
	generation: number
}

function emptyMemo(): Memo {
	return {
		promise: null,
		finding: null,
		uuid: null,
		pending: 0,
		generation: 0
	}
}

const dotFilen = emptyMemo()
const subdirs = new Map<string, Memo>()
const forgetListeners = new Map<string, Set<() => void>>()

function subdirMemo(name: string): Memo {
	let memo = subdirs.get(name)

	if (memo === undefined) {
		memo = emptyMemo()
		subdirs.set(name, memo)
	}

	return memo
}

// A rejected lookup clears its memo before propagating, so the next caller tries again instead of every later
// use failing for the session on one transient error.
function memoize(memo: Memo, resolve: () => Promise<string>): Promise<string> {
	if (memo.promise !== null) {
		return memo.promise
	}

	const generation = memo.generation

	memo.pending++

	const promise = resolve().then(
		uuid => {
			memo.pending--

			if (generation === memo.generation) {
				memo.uuid = uuid
			}

			return uuid
		},
		(error: unknown) => {
			memo.pending--

			// Only its own memo: a lookup forgotten mid-flight may already have a successor.
			if (memo.promise === promise) {
				memo.promise = null
			}

			throw error
		}
	)

	memo.promise = promise

	return promise
}

// A found directory is remembered like a created one; absence never is.
function memoizeFind(memo: Memo, find: () => Promise<string | null>): Promise<string | null> {
	if (memo.promise !== null) {
		return memo.promise
	}

	if (memo.finding !== null) {
		return memo.finding
	}

	const generation = memo.generation

	memo.pending++

	const finding = find()
		.then(uuid => {
			if (uuid !== null && generation === memo.generation && memo.promise === null) {
				memo.uuid = uuid
				memo.promise = Promise.resolve(uuid)
			}

			return uuid
		})
		.finally(() => {
			memo.pending--

			if (memo.finding === finding) {
				memo.finding = null
			}
		})

	memo.finding = finding

	return finding
}

// Creates when no find is out; after one that came back empty, creates.
function memoizeCreate(memo: Memo, create: () => Promise<string>): Promise<string> {
	const finding = memo.promise === null ? memo.finding : null

	return finding === null ? memoize(memo, create) : finding.then(found => found ?? memoize(memo, create))
}

// `name` under `parent` (null: the drive root) when it exists, creating nothing.
async function findLevel(parent: string | null, name: string): Promise<string | null> {
	if (!(await runOp(sdkApi.directoryExists(parent, name)))) {
		return null
	}

	return (await runOp(sdkApi.createDirectory(parent, name))).uuid
}

function dotFilenUuid(): Promise<string> {
	return memoizeCreate(dotFilen, async () => (await runOp(sdkApi.createDirectory(null, DOT_FILEN_DIR_NAME))).uuid)
}

// `.filen/<name>`, created when missing, memoized until forgotten.
export function getDotFilenSubdir(name: string): Promise<string> {
	return memoizeCreate(subdirMemo(name), async () => {
		const parent = await dotFilenUuid()

		return (await runOp(sdkApi.createDirectory(parent, name))).uuid
	})
}

// `.filen/<name>` when it exists, else null; never creates either level. For reads: a feature's directory is
// created by its first write.
export function findDotFilenSubdir(name: string): Promise<string | null> {
	return memoizeFind(subdirMemo(name), async () => {
		const parent = await memoizeFind(dotFilen, () => findLevel(null, DOT_FILEN_DIR_NAME))

		return parent === null ? null : findLevel(parent, name)
	})
}

// The resolved uuid of `.filen/<name>`, or null when it isn't known (yet).
export function dotFilenSubdirUuid(name: string): string | null {
	return subdirs.get(name)?.uuid ?? null
}

// Whether `uuid` is `.filen` or one of its resolved subdirectories. Synchronous, for socket handlers.
export function isDotFilenSystemDir(uuid: string): boolean {
	if (uuid === dotFilen.uuid) {
		return true
	}

	for (const memo of subdirs.values()) {
		if (memo.uuid === uuid) {
			return true
		}
	}

	return false
}

function clearMemo(memo: Memo): void {
	memo.generation++
	memo.promise = null
	memo.finding = null
	memo.uuid = null
}

function notifyForgotten(name: string): void {
	for (const listener of forgetListeners.get(name) ?? []) {
		listener()
	}
}

// Whether a subdirectory other than `name` is resolved or being resolved under the current `.filen`.
function hasSibling(name: string): boolean {
	for (const [other, memo] of subdirs) {
		if (other !== name && (memo.uuid !== null || memo.pending > 0)) {
			return true
		}
	}

	return false
}

// The subdirectory was trashed, deleted, moved or renamed, or an op found it gone. With no sibling resolved under
// it, `.filen` is looked up again too (a gone subdirectory may mean a gone parent); a sibling vouches for it.
export function forgetDotFilenSubdir(name: string): void {
	clearMemo(subdirMemo(name))

	if (!hasSibling(name)) {
		clearMemo(dotFilen)
	}

	notifyForgotten(name)
}

// `.filen` itself went: every subdirectory with it, and every feature is told.
function forgetAll(): void {
	clearMemo(dotFilen)

	for (const name of new Set([...subdirs.keys(), ...forgetListeners.keys()])) {
		clearMemo(subdirMemo(name))
		notifyForgotten(name)
	}
}

// Called whenever `.filen/<name>` is forgotten, so a feature drops what it derived from it.
export function onDotFilenSubdirForgotten(name: string, listener: () => void): () => void {
	let listeners = forgetListeners.get(name)

	if (listeners === undefined) {
		listeners = new Set()
		forgetListeners.set(name, listeners)
	}

	listeners.add(listener)

	return () => {
		listeners.delete(listener)
	}
}

function keepsName(meta: DirMeta, name: string): boolean {
	return meta.type === "decoded" && meta.data.name.toLowerCase() === name.toLowerCase()
}

// Mid-lookup, a directory renamed to one of the names being looked up is still ours if it was.
function keepsPendingName(meta: DirMeta): boolean {
	if (keepsName(meta, DOT_FILEN_DIR_NAME)) {
		return true
	}

	for (const [name, memo] of subdirs) {
		if (memo.pending > 0 && keepsName(meta, name)) {
			return true
		}
	}

	return false
}

// A directory event against the memos. `parent` is the directory it left or sits in, when the event says;
// `renamed` its new metadata for a rename. A lookup still out can't be matched by uuid, so an event that could be
// the directory it resolves forgets that lookup (and only that one).
function forgetDirectory({ uuid, parent, renamed }: { uuid: string; parent: string | null; renamed: DirMeta | null }): void {
	if (uuid === dotFilen.uuid) {
		if (renamed === null || !keepsName(renamed, DOT_FILEN_DIR_NAME)) {
			forgetAll()
		}

		return
	}

	const uncertain = renamed === null || !keepsPendingName(renamed)

	if (dotFilen.pending > 0 && uncertain) {
		clearMemo(dotFilen)

		for (const [name, memo] of subdirs) {
			if (memo.pending > 0) {
				clearMemo(memo)
				notifyForgotten(name)
			}
		}

		return
	}

	for (const [name, memo] of subdirs) {
		const isIt = uuid === memo.uuid && (renamed === null || !keepsName(renamed, name))
		const mayBeIt = memo.pending > 0 && uncertain && (parent === null || parent === dotFilen.uuid)

		if (isIt || mayBeIt) {
			forgetDotFilenSubdir(name)
		}
	}
}

function anythingKnown(): boolean {
	return dotFilen.uuid !== null || dotFilen.pending > 0 || [...subdirs.values()].some(memo => memo.uuid !== null || memo.pending > 0)
}

// Files seen landing in a system directory this session that can't be photos, so a later event naming only the
// file's uuid (a new version retiring it, a trash, a delete) is known not to concern a photos listing. A file that
// leaves, is renamed or is trashed is dropped again. Bounded: each save retires one uuid and adds the next.
const MAX_SYSTEM_FILES = 512
const systemFiles = new Set<string>()

export function noteDotFilenFile(uuid: string): void {
	systemFiles.delete(uuid)
	systemFiles.add(uuid)

	if (systemFiles.size > MAX_SYSTEM_FILES) {
		const oldest = systemFiles.values().next()

		if (oldest.done !== true) {
			systemFiles.delete(oldest.value)
		}
	}
}

export function isDotFilenSystemFile(uuid: string): boolean {
	return systemFiles.has(uuid)
}

// The only handler of `.filen` directory events: features learn of them through onDotFilenSubdirForgotten, so the
// order handlers are registered in doesn't matter.
export function handleDotFilenDriveEvent(event: Extract<SocketEvent, { type: "drive" }>): void {
	const inner = event.inner

	switch (inner.type) {
		case "fileNew":
		case "fileRestore":
		case "fileMove":
		case "fileArchiveRestored":
			if (isDotFilenSystemDir(inner.file.parent) && !isPhotoItem(narrowItem(inner.file))) {
				noteDotFilenFile(inner.file.uuid)
			} else {
				systemFiles.delete(inner.file.uuid)
			}

			return

		case "fileTrash":
		case "fileMetadataChanged":
			systemFiles.delete(inner.uuid)

			return

		case "folderTrash":
			forgetDirectory({ uuid: inner.uuid, parent: inner.parent, renamed: null })

			return

		case "folderDeletedPermanent":
			forgetDirectory({ uuid: inner.uuid, parent: null, renamed: null })

			return

		case "folderMove":
			forgetDirectory({ uuid: inner.dir.uuid, parent: null, renamed: null })

			return

		case "folderMetadataChanged":
			forgetDirectory({ uuid: inner.uuid, parent: null, renamed: inner.meta })

			return

		case "deleteAll":
			if (anythingKnown()) {
				forgetAll()
			}

			return

		default:
			return
	}
}

// Called once by the authed shell's socket host.
export function registerDotFilenSocketHandlers(): () => void {
	return registerSocketHandler("drive", handleDotFilenDriveEvent)
}

// Logout: nothing resolved under the signed-out account survives, and a lookup still in flight answers its own
// callers only.
export function resetDotFilen(): void {
	clearMemo(dotFilen)

	for (const memo of subdirs.values()) {
		clearMemo(memo)
	}

	systemFiles.clear()
}
