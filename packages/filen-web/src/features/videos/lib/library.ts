import {
	LIBRARY_CONFIG_FILE_NAME,
	VIDEOS_DIR_NAME,
	activeLibraryRoots,
	emptyLibraryConfig,
	hlcObserve,
	hlcTick,
	libraryConfigMaxStamp,
	mergeLibraryConfig,
	parseLibraryConfigText,
	serializeLibraryConfig,
	validateLibraryRoot,
	type Hlc,
	type LibraryConfigDoc,
	type LibraryRootValidation,
	type Stamp
} from "@filen/shared/videos"
import { settleHeldRevisions } from "@filen/shared"
import type { File as SdkFile, SocketEvent } from "@filen/sdk-rs"
import { sdkApi } from "@/lib/sdk/client"
import { runOp } from "@/lib/actions/outcome"
import { log } from "@/lib/log"
import { registerSocketHandler } from "@/lib/sdk/socket"
import { readSignOutPending } from "@/lib/signOutPending"
import { DOT_FILEN_DIR_NAME, dotFilenSubdirUuid, getDotFilenSubdir, noteDotFilenFile, onDotFilenSubdirForgotten } from "@/lib/dotFilen"
import { queryClient } from "@/queries/client"
import { persister } from "@/queries/persist"
import { versionAsFile } from "@/features/videos/lib/library.logic"

// The synced library configuration, `.filen/Videos/library.v1.json`: read once per session (none at all when the
// drive hasn't changed since the cached copy was read), written only when an edit changes it. Each upload is a
// new file version, six requests, and comes back to every session of the account as a socket event; this tab's
// own can arrive before its upload resolves. Reads, edits and uploads run one at a time, each on what the previous
// left, and every upload is the merge of this tab's copy with the newest remote one it knows.

export const LIBRARY_CONFIG_QUERY_KEY = ["videos", "library"] as const

export interface LibraryConfigState {
	// The account (its drive root uuid) the state was read under; a cached state of another account is ignored.
	account: string
	doc: LibraryConfigDoc
	// `.filen/Videos` as of the last read, so a session served from the cache still recognizes saves made
	// elsewhere without looking the directory up.
	dirUuid: string | null
	// Drive uuid of the file the document was read from or saved as; null when none exists yet.
	headUuid: string | null
	// Drive event id the read was current to; null when unknown (after this tab's own save).
	watermark: bigint | null
}

// The signed-in account, set when the library opens; null before, and after logout.
let account: string | null = null

export function setLibraryAccount(next: string): void {
	account = next
}

// This tab's writer id and clock for stamping edits: one writer per tab, so two tabs never issue the same stamp.
const writerId = crypto.randomUUID().slice(0, 8)
let clock: Hlc = 0

function nextStamp(): Stamp {
	clock = hlcTick({
		clock,
		now: Date.now()
	})

	return {
		t: clock,
		d: writerId
	}
}

function observe(doc: LibraryConfigDoc): void {
	clock = hlcObserve({
		clock,
		seen: libraryConfigMaxStamp(doc),
		now: Date.now()
	})
}

// Another session changed the file: the next open reads it again.
let unsynced = false
// The drive event id the cached document is known current to this session: set by a read, advanced by every
// event the socket delivers in order (none can change the file unseen), dropped on a drop. Written into the
// cached state when the page goes away, so the next session's open compares against it.
let currentTo: bigint | null = null
// Heads this tab saved, and the saves echoed while one of its own was in flight (judged once it settles).
const ownUuids = new Set<string>()
let held: SdkFile[] | null = null
// The newest remote document this tab knows: what it read, or what it saved since.
let lastRemote: LibraryConfigDoc | null = null
// A sign-out started (here or in another tab): nothing more is persisted.
let closed = false

export function markLibraryConfigUnsynced(): void {
	unsynced = true
}

onDotFilenSubdirForgotten(VIDEOS_DIR_NAME, markLibraryConfigUnsynced)

function isConfigName(name: string): boolean {
	return name.toLowerCase() === LIBRARY_CONFIG_FILE_NAME
}

function cached(): LibraryConfigState | undefined {
	const state = queryClient.getQueryData<LibraryConfigState>(LIBRARY_CONFIG_QUERY_KEY)

	return state !== undefined && account !== null && state.account === account ? state : undefined
}

// Into the query cache, and its persisted row (never once a sign-out started).
function store(state: LibraryConfigState): LibraryConfigState {
	queryClient.setQueryData(LIBRARY_CONFIG_QUERY_KEY, state)

	const query = queryClient.getQueryCache().find({
		queryKey: LIBRARY_CONFIG_QUERY_KEY,
		exact: true
	})

	if (query !== undefined && !closed && !readSignOutPending()) {
		void persister.persistQuery(query)
	}

	return state
}

function requireAccount(): string {
	if (account === null) {
		throw new Error("the library is not open")
	}

	return account
}

function parseBytes(bytes: Uint8Array): LibraryConfigDoc | null {
	return parseLibraryConfigText(new TextDecoder().decode(bytes))
}

async function readFileDoc(file: SdkFile): Promise<LibraryConfigDoc | null> {
	return parseBytes(await runOp(sdkApi.downloadFileBytes(file, crypto.randomUUID())))
}

// A head that doesn't parse (a torn foreign write) is never taken for an empty document: the newest version that
// does parse stands in for it.
const CORRUPT_HEAD_VERSIONS = 3

async function readHead(head: SdkFile): Promise<LibraryConfigDoc> {
	const doc = await readFileDoc(head)

	if (doc !== null) {
		return doc
	}

	log.warn("videos.library", "library file doesn't parse; reading its previous versions", head.uuid)

	const versions = await runOp(sdkApi.listFileVersionsOp(head))

	for (const version of versions.filter(item => item.uuid !== head.uuid).slice(0, CORRUPT_HEAD_VERSIONS)) {
		const older = await readFileDoc(versionAsFile(head, version))

		if (older !== null) {
			return older
		}
	}

	return emptyLibraryConfig()
}

async function readRemote(): Promise<{ doc: LibraryConfigDoc; head: SdkFile | null; dirUuid: string }> {
	const dirUuid = await getDotFilenSubdir(VIDEOS_DIR_NAME)
	const { files } = await runOp(sdkApi.listDirectory({ kind: "uuid", uuid: dirUuid }))
	const head = files.find(file => file.meta.type === "decoded" && isConfigName(file.meta.data.name)) ?? null

	if (head === null) {
		return {
			doc: emptyLibraryConfig(),
			head: null,
			dirUuid
		}
	}

	noteDotFilenFile(head.uuid)

	return {
		doc: await readHead(head),
		head,
		dirUuid
	}
}

// One read, edit or upload at a time.
let queue: Promise<unknown> = Promise.resolve()

function serial<T>(task: () => Promise<T>): Promise<T> {
	const run = queue.then(task)

	queue = run.catch(() => undefined)

	return run
}

// Uploads `doc` and resolves the new head's uuid plus the saves from elsewhere this one went over (echoed before
// its own). One saved after it marks the configuration for a read.
async function upload(doc: LibraryConfigDoc): Promise<{ headUuid: string; replaced: SdkFile[] }> {
	const dirUuid = await getDotFilenSubdir(VIDEOS_DIR_NAME)
	const bytes = new TextEncoder().encode(JSON.stringify(doc))

	held = []

	const settle = (savedUuid: string | null): SdkFile[] => {
		const echoes = held ?? []

		held = null

		const { newer } = settleHeldRevisions(echoes, savedUuid, echo => echo.uuid)

		if (newer.length > 0) {
			unsynced = true
		}

		return echoes.filter(echo => echo.uuid !== savedUuid && !newer.includes(echo))
	}

	let saved: SdkFile

	try {
		saved = await runOp(sdkApi.uploadFileBytes(dirUuid, bytes, LIBRARY_CONFIG_FILE_NAME, "application/json"))
	} catch (e) {
		settle(null)

		throw e
	}

	ownUuids.add(saved.uuid)
	noteDotFilenFile(saved.uuid)

	return {
		headUuid: saved.uuid,
		replaced: settle(saved.uuid)
	}
}

// Saves to bring the remote file up to `state.doc` merged with the newest remote one known, when that merge holds
// something the remote lacks. A save that went over another session's (both landed at once) folds that one in
// and saves again.
const RECONCILE_ROUNDS = 3

async function reconcile(state: LibraryConfigState): Promise<LibraryConfigState> {
	let next = state

	for (let round = 0; round < RECONCILE_ROUNDS; round++) {
		const merge = mergeLibraryConfig({ local: next.doc, remote: lastRemote ?? emptyLibraryConfig() })

		if (!merge.localContributed && lastRemote !== null) {
			return {
				...next,
				doc: merge.doc
			}
		}

		const { headUuid, replaced } = await upload(merge.doc)

		lastRemote = merge.doc
		next = {
			...next,
			doc: merge.doc,
			dirUuid: dotFilenSubdirUuid(VIDEOS_DIR_NAME) ?? next.dirUuid,
			headUuid,
			watermark: null
		}

		if (replaced.length === 0) {
			return next
		}

		let doc = next.doc

		for (const file of replaced) {
			const theirs = await readFileDoc(file).catch(() => null)

			if (theirs !== null) {
				observe(theirs)
				doc = mergeLibraryConfig({ local: doc, remote: theirs }).doc
			}
		}

		next = {
			...next,
			doc
		}
	}

	unsynced = true

	return next
}

// A read, merged with the cached copy (an edit another session's save went over is restored and saved back).
async function readAndReconcile(watermark: bigint | null): Promise<LibraryConfigState> {
	const owner = requireAccount()
	const previous = cached()

	unsynced = false

	const remote = await readRemote()

	currentTo = watermark
	observe(remote.doc)
	lastRemote = remote.doc

	const read: LibraryConfigState = {
		account: owner,
		doc: remote.doc,
		dirUuid: remote.dirUuid,
		headUuid: remote.head?.uuid ?? null,
		watermark
	}

	return store(
		previous === undefined
			? read
			: await reconcile({
					...read,
					doc: mergeLibraryConfig({ local: previous.doc, remote: remote.doc }).doc
				})
	)
}

// The configuration, current to `watermark` (the drive's latest event id, when the caller has one): the cached
// copy as it is when nothing in the drive changed since it was read, else read again.
export function fetchLibraryConfig(watermark: bigint | null): Promise<LibraryConfigState> {
	return serial(async () => {
		const previous = cached()

		if (watermark !== null && !unsynced && previous !== undefined && (currentTo ?? previous.watermark) === watermark) {
			currentTo = watermark
			lastRemote ??= previous.doc

			return previous
		}

		return readAndReconcile(watermark)
	})
}

// Page hidden or closing: records how far the cached document is known current, so the next session's open
// needs no read when nothing changed in the drive since. Never while a save is out (its echo may still be
// held), after a change was seen, or once a sign-out started.
export function persistLibraryConfigWatermark(): void {
	const state = cached()

	if (
		state === undefined ||
		currentTo === null ||
		state.watermark === currentTo ||
		unsynced ||
		held !== null ||
		closed ||
		readSignOutPending()
	) {
		return
	}

	store({
		...state,
		watermark: currentTo
	})
}

// Brings the cached configuration current to `watermark` and returns it. Not through the query's own fetch: the
// query's reader opens the library, which lands here.
export function ensureLibraryConfig(watermark: bigint | null): Promise<LibraryConfigState> {
	return fetchLibraryConfig(watermark)
}

// The cached configuration of the signed-in account, if any.
export function cachedLibraryConfig(): LibraryConfigState | undefined {
	return cached()
}

// Whether the cached configuration may be behind: never read, a change was seen since, or the event stream
// can't vouch for it this session (not checked yet, or a gap since).
export function libraryConfigNeedsRead(): boolean {
	return unsynced || currentTo === null || cached() === undefined
}

// Applies an edit (a function of the current document and a fresh stamp) and saves the result, unless it
// changes nothing. Resolves the saved state.
export function updateLibraryConfig(edit: (doc: LibraryConfigDoc, stamp: Stamp) => LibraryConfigDoc): Promise<LibraryConfigState> {
	return serial(async () => {
		const known = cached()
		const current = libraryConfigNeedsRead() || known === undefined ? await readAndReconcile(currentTo) : known

		lastRemote ??= current.doc

		const next = edit(current.doc, nextStamp())

		if (serializeLibraryConfig(next) === serializeLibraryConfig(current.doc)) {
			return current
		}

		return store(
			await reconcile({
				...current,
				doc: next
			})
		)
	})
}

// ── Root validation ──

// Whether `uuid` can become a library root, given the current roots: an own-drive directory outside `.filen`
// that neither holds nor sits inside another root (the overlapped ones are offered for replacement). `ownDrive` is
// false for a shared-in or public-link directory. Ancestors usually come from the directory cache, no request.
export async function checkLibraryRoot({
	uuid,
	ownDrive
}: {
	uuid: string
	ownDrive: boolean
}): Promise<LibraryRootValidation | { ok: false; reason: "notFound" }> {
	if (!ownDrive) {
		return {
			ok: false,
			reason: "notOwnDrive"
		}
	}

	const chain = await runOp(sdkApi.videoRootChain(uuid))

	if (chain === null) {
		return {
			ok: false,
			reason: "notFound"
		}
	}

	const roots = await Promise.all(
		activeLibraryRoots(cached()?.doc ?? emptyLibraryConfig()).map(async root => {
			const rootChain = await runOp(sdkApi.videoRootChain(root.dirUuid))

			return rootChain === null
				? null
				: {
						uuid: root.dirUuid,
						ancestors: rootChain.slice(0, -1).map(entry => entry.uuid)
					}
		})
	)
	// `.filen` sits directly under the drive root.
	const topLevel = chain[1]

	return validateLibraryRoot({
		candidate: {
			uuid,
			ancestors: chain.slice(0, -1).map(entry => entry.uuid),
			ownDrive
		},
		roots: roots.filter(root => root !== null),
		dotFilenUuid: topLevel?.name.toLowerCase() === DOT_FILEN_DIR_NAME ? topLevel.uuid : null
	})
}

// ── Socket ──

type DriveSocketEvent = Extract<SocketEvent, { type: "drive" }>

// Another session's save of the file marks the configuration for a read on the next open; this tab's own save
// doesn't, whether its echo comes before the upload resolves or after.
export function handleLibraryConfigDriveEvent(event: DriveSocketEvent): void {
	const dirUuid = dotFilenSubdirUuid(VIDEOS_DIR_NAME) ?? cached()?.dirUuid ?? null
	const inner = event.inner

	currentTo = currentTo !== null && event.driveMessageId === currentTo + 1n ? event.driveMessageId : null

	if (dirUuid === null) {
		return
	}

	switch (inner.type) {
		// The directory itself went or changed: it is looked up afresh with the next read.
		case "folderTrash":
		case "folderDeletedPermanent":
		case "folderMetadataChanged":
			if (inner.uuid === dirUuid) {
				unsynced = true
			}

			return

		case "folderMove":
			if (inner.dir.uuid === dirUuid) {
				unsynced = true
			}

			return

		case "fileNew":
		case "fileRestore":
		case "fileMove":
		case "fileArchiveRestored": {
			const file = inner.file

			if (file.parent !== dirUuid || file.meta.type !== "decoded" || !isConfigName(file.meta.data.name)) {
				return
			}

			if (held !== null) {
				held.push(file)
			} else if (!ownUuids.has(file.uuid)) {
				unsynced = true
			}

			return
		}

		case "fileTrash":
		case "fileDeletedPermanent":
			if (inner.uuid === cached()?.headUuid) {
				unsynced = true
			}

			return

		default:
			return
	}
}

function loseTrack(): void {
	currentTo = null
}

// Called once by the authed shell's socket host. An undecodable event or a drop may hide a save from elsewhere.
export function registerLibraryConfigSocketHandlers(): () => void {
	const unregisterDrive = registerSocketHandler("drive", handleLibraryConfigDriveEvent)
	const unregisterMalformed = registerSocketHandler("driveMalformed", loseTrack)
	const unregisterReconnecting = registerSocketHandler("reconnecting", loseTrack)

	return () => {
		unregisterDrive()
		unregisterMalformed()
		unregisterReconnecting()
	}
}

// A sign-out started, here or in another tab: nothing of this account is persisted any more.
export function holdLibraryConfig(): void {
	closed = true
}

// Logout (then held, see holdLibraryConfig): nothing of this account's saves is remembered.
export function resetLibraryConfig(): void {
	closed = false
	account = null
	unsynced = false
	currentTo = null
	held = null
	lastRemote = null
	ownUuids.clear()
	clock = 0
	queue = Promise.resolve()
}
