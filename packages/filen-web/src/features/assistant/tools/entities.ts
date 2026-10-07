import { contactDisplayName, deriveBlockedUsers, driveItemName, EMPTY_BLOCKED_USERS, type BlockedUsers, type Playlist } from "@filen/shared"
import type { Chat, Note } from "@filen/sdk-rs"
import type { AssistantSignal } from "@filen/shared/assistant"
import { i18n } from "@/lib/i18n"
import { queryClient } from "@/queries/client"
import { accountQueryGet } from "@/queries/account"
import { contactsQueryGet } from "@/features/contacts/queries/contacts"
import { cachedQueriesWithPrefix } from "@/queries/patch"
import {
	DRIVE_LISTING_KEY_PREFIX,
	driveListingQueryOptions,
	findCachedListingItem,
	type DriveListingParams
} from "@/features/drive/queries/drive"
import { isDirectoryItem, type DriveItem } from "@/features/drive/lib/item"
import { notesQueryFind, notesQueryGet } from "@/features/notes/queries/notes"
import { isNoteUndecryptable } from "@/features/notes/lib/sort"
import { chatsQueryGet } from "@/features/chats/queries/chats"
import { chatTitle } from "@/features/chats/lib/sort"
import { playlistsQueryGet } from "@/features/audio/queries/playlists"
import { AssistantToolError, matchScore, sameName, searchTokens } from "@/features/assistant/tools/format"
import type { AssistantEntity, AssistantEntityKind, AssistantHandleRegistry } from "@/features/assistant/tools/handles"
import type { AssistantHandleRef, AssistantUiEffect } from "@/features/assistant/runtime/executor"

// What an adapter gets for one call. `show` registers an item and records that the result mentions it,
// which is what the guard learns handles from.
export interface ToolScope {
	registry: AssistantHandleRegistry
	signal: AssistantSignal | undefined
	show: (entity: AssistantEntity, name: string) => string
}

export interface ToolOutput {
	value: Record<string, unknown>
	uiEffect?: AssistantUiEffect
}

export type ToolAdapter = (input: Record<string, unknown>, scope: ToolScope) => Promise<ToolOutput>

// The loop's signal as a DOM AbortSignal, for the download it may stop.
export function abortSignalOf(signal: AssistantSignal | undefined): AbortSignal | undefined {
	if (signal === undefined || signal instanceof AbortSignal) {
		return signal
	}

	const controller = new AbortController()

	if (signal.aborted) {
		controller.abort()
	} else {
		signal.addEventListener(
			"abort",
			() => {
				controller.abort()
			},
			{ once: true }
		)
	}

	return controller.signal
}

export function optionalString(input: Record<string, unknown>, key: string): string | undefined {
	const value = input[key]

	return typeof value === "string" && value.trim().length > 0 ? value : undefined
}

export function requiredString(input: Record<string, unknown>, key: string): string {
	return optionalString(input, key) ?? ""
}

export function currentUserId(): bigint | undefined {
	return accountQueryGet()?.id
}

// The blocked set the sidebars filter by, as the contacts cache holds it; fail-open like useBlockedUsers.
export function blockedUsers(): BlockedUsers {
	const blocked = contactsQueryGet()?.blocked

	return blocked === undefined || blocked.length === 0 ? EMPTY_BLOCKED_USERS : deriveBlockedUsers(blocked)
}

export function driveName(item: DriveItem): string {
	return item.data.undecryptable ? i18n.t("common:cannotDecryptTitle") : driveItemName(item)
}

export function noteTitle(note: Note): string {
	if (isNoteUndecryptable(note)) {
		return i18n.t("common:cannotDecryptTitle")
	}

	return note.title !== undefined && note.title.length > 0 ? note.title : i18n.t("notes:noteUntitled")
}

export function chatName(chat: Chat): string {
	return chatTitle(chat, currentUserId(), i18n.t("common:cannotDecryptTitle"), i18n.t("chats:chatJustYou"))
}

export function participantName(participant: { email: string; nickName: string | undefined }): string {
	return contactDisplayName(participant)
}

// The live rows behind a snapshot: a rename or a socket patch since the model saw the handle wins.
export function freshDriveItem(item: DriveItem): DriveItem {
	return findCachedListingItem(item.data.uuid) ?? item
}

export function freshNote(note: Note): Note {
	return notesQueryFind(note.uuid) ?? note
}

export function freshChat(chat: Chat): Chat {
	return chatsQueryGet()?.find(candidate => candidate.uuid === chat.uuid) ?? chat
}

export function freshPlaylist(playlist: Playlist): Playlist {
	const entry = playlistsQueryGet()?.find(candidate => candidate.status === "ok" && candidate.playlist.uuid === playlist.uuid)

	return entry?.status === "ok" ? entry.playlist : playlist
}

export function rootDirUuid(): string | undefined {
	return accountQueryGet()?.rootDirUuid
}

// A directory listing through the same query, key and freshness rule the drive screen uses.
export function readDriveListing(uuid: string | null): Promise<DriveItem[]> {
	return queryClient.query(driveListingQueryOptions("drive", uuid === rootDirUuid() ? null : uuid))
}

export function joinPath(parent: string, name: string): string {
	return parent === "/" || parent === "" ? `/${name}` : `${parent}/${name}`
}

// An absolute path ("/Documents/Invoices") walked through the listings the drive screen would read.
export async function resolveDrivePath(path: string): Promise<AssistantEntity> {
	const segments = path.split("/").filter(segment => segment.trim().length > 0)
	let entity: AssistantEntity = { kind: "root" }
	let uuid: string | null = null
	let walked = "/"

	for (const segment of segments) {
		const items = await readDriveListing(uuid)
		const found = items.find(item => item.type === "directory" && !item.data.undecryptable && sameName(driveItemName(item), segment))

		if (found === undefined) {
			throw new AssistantToolError(`No directory at "${path}"`)
		}

		walked = joinPath(walked, driveItemName(found))
		uuid = found.data.uuid
		entity = { kind: "drive", item: found, path: walked }
	}

	return entity
}

const KIND_LABELS: Record<AssistantEntityKind, string> = {
	root: "the drive root",
	drive: "a file or directory",
	note: "a note",
	chat: "a chat",
	playlist: "a playlist",
	track: "a track"
}

const LOOKUP_TOOLS: Record<AssistantEntityKind, string> = {
	root: "list_directory",
	drive: "search_items or list_directory",
	note: "list_notes",
	chat: "list_chats",
	playlist: "list_playlists",
	track: "open_playlist"
}

export function entityName(entity: AssistantEntity): string {
	switch (entity.kind) {
		case "root":
			return "/"
		case "drive":
			return driveName(entity.item)
		case "note":
			return noteTitle(entity.note)
		case "chat":
			return chatName(entity.chat)
		case "playlist":
			return entity.playlist.name
		case "track":
			return entity.track.name
	}
}

function entityKindLabel(entity: AssistantEntity): string {
	switch (entity.kind) {
		case "root":
			return "directory"
		case "drive":
			return driveKind(entity.item)
		default:
			return entity.kind
	}
}

function entityUuid(entity: AssistantEntity): string {
	switch (entity.kind) {
		case "root":
			return ""
		case "drive":
			return entity.item.data.uuid
		case "note":
			return entity.note.uuid
		case "chat":
			return entity.chat.uuid
		case "playlist":
			return entity.playlist.uuid
		case "track":
			return entity.track.uuid
	}
}

// Every My Drive row the cached listings hold, with its path where every ancestor is cached too.
export function cachedDriveRows(): { item: DriveItem; path: string | undefined }[] {
	const rows = new Map<string, DriveItem>()

	for (const query of cachedQueriesWithPrefix(DRIVE_LISTING_KEY_PREFIX)) {
		const params = query.queryKey[2] as DriveListingParams | undefined

		if (params?.variant !== "drive") {
			continue
		}

		for (const item of (query.state.data as DriveItem[] | undefined) ?? []) {
			rows.set(item.data.uuid, item)
		}
	}

	const root = rootDirUuid()
	const paths = new Map<string, string | undefined>()

	function pathOf(uuid: string): string | undefined {
		if (uuid === root) {
			return "/"
		}

		if (paths.has(uuid)) {
			return paths.get(uuid)
		}

		// Marked first, so a cyclic cache can't recurse forever.
		paths.set(uuid, undefined)

		const row = rows.get(uuid)
		const parent = row === undefined ? undefined : pathOf(row.data.parent)
		const path = row === undefined || parent === undefined ? undefined : joinPath(parent, driveItemName(row))

		paths.set(uuid, path)

		return path
	}

	return [...rows.values()].map(item => ({ item, path: pathOf(item.data.uuid) }))
}

// What a name could mean: the items this conversation has seen, then what the caches already hold.
function candidatePool(scope: ToolScope, kinds: readonly AssistantEntityKind[]): AssistantEntity[] {
	const pool: AssistantEntity[] = scope.registry
		.entries()
		.map(entry => entry.entity)
		.filter(entity => kinds.includes(entity.kind))

	if (kinds.includes("drive")) {
		for (const { item, path } of cachedDriveRows()) {
			pool.push(path === undefined ? { kind: "drive", item } : { kind: "drive", item, path })
		}
	}

	if (kinds.includes("note")) {
		for (const note of notesQueryGet() ?? []) {
			if (!note.trash) {
				pool.push({ kind: "note", note })
			}
		}
	}

	if (kinds.includes("chat")) {
		for (const chat of chatsQueryGet() ?? []) {
			pool.push({ kind: "chat", chat })
		}
	}

	if (kinds.includes("playlist")) {
		for (const entry of playlistsQueryGet() ?? []) {
			if (entry.status === "ok") {
				pool.push({ kind: "playlist", playlist: entry.playlist })
			}
		}
	}

	return pool
}

const CANDIDATE_LIMIT = 5

// The item behind a handle the model passed, if it is one of `kinds`. A path stands in for a drive
// handle; a name where a handle belongs answers with the handles whose names match, so the model can
// call again with one.
export async function resolveHandle(scope: ToolScope, handle: string, kinds: readonly AssistantEntityKind[]): Promise<AssistantEntity> {
	const wanted = handle.trim()
	const expected = kinds.map(kind => KIND_LABELS[kind]).join(" or ")

	if (wanted.startsWith("/") && kinds.includes("drive")) {
		const entity = await resolveDrivePath(wanted)

		if (!kinds.includes(entity.kind)) {
			throw new AssistantToolError(`"${wanted}" is ${KIND_LABELS[entity.kind]}, not ${expected}`, { invalid: true })
		}

		return entity
	}

	const entry = scope.registry.get(wanted)

	if (entry !== undefined) {
		if (!kinds.includes(entry.entity.kind)) {
			throw new AssistantToolError(`"${wanted}" is ${KIND_LABELS[entry.entity.kind]}, not ${expected}`, { invalid: true })
		}

		return entry.entity
	}

	const lookup = LOOKUP_TOOLS[kinds.find(kind => kind !== "root") ?? "drive"]
	const tokens = searchTokens(wanted)

	if (/^[a-z]\d+$/i.test(wanted) || tokens.length === 0) {
		throw new AssistantToolError(`Unknown handle "${wanted}". Find it first with ${lookup}.`, { invalid: true })
	}

	const seen = new Set<string>()
	const candidates: { ref: AssistantHandleRef; kind: string }[] = []

	for (const entity of candidatePool(scope, kinds)) {
		const uuid = entityUuid(entity)
		const name = entityName(entity)

		if (candidates.length >= CANDIDATE_LIMIT) {
			break
		}

		if (seen.has(uuid) || matchScore(name, tokens) < tokens.length) {
			continue
		}

		seen.add(uuid)
		candidates.push({ ref: { handle: scope.show(entity, name), name }, kind: entityKindLabel(entity) })
	}

	throw new AssistantToolError(
		candidates.length > 0
			? `"${wanted}" is a name, not a handle. Matching handles: ${candidates.map(({ ref, kind }) => `${ref.handle} "${ref.name}" (${kind})`).join(", ")}. Call again with the handle.`
			: `No item "${wanted}". Find it first with ${lookup}.`,
		{ invalid: true, candidates: candidates.map(({ ref }) => ref) }
	)
}

export function showDrive(scope: ToolScope, item: DriveItem, path?: string): string {
	return scope.show(path === undefined ? { kind: "drive", item } : { kind: "drive", item, path }, driveName(item))
}

export function driveKind(item: DriveItem): "file" | "directory" {
	return isDirectoryItem(item) ? "directory" : "file"
}
