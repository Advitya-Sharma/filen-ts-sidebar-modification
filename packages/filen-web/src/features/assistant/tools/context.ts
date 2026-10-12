import type { Chat, Note } from "@filen/sdk-rs"
import { splatToUuids } from "@/features/drive/lib/navigate"
import { driveNamesQueryKey, findCachedListingItem, type DirectoryNameScope } from "@/features/drive/queries/drive"
import { queryClient } from "@/queries/client"
import { useDriveStore } from "@/features/drive/store/useDriveStore"
import type { DriveItem } from "@/features/drive/lib/item"
import { selectedNoteUuidFromPath } from "@/features/notes/lib/route"
import { notesQueryFind } from "@/features/notes/queries/notes"
import { selectedChatUuidFromPath } from "@/features/chats/components/chatsSidebar.logic"
import { chatsQueryGet } from "@/features/chats/queries/chats"
import { useAudioStore } from "@/features/audio/store/useAudioStore"
import type { AudioPlaybackStatus, QueueTrack } from "@/features/audio/store/audioQueue"
import { assistantRouter } from "@/features/assistant/tools/navigation"
import { screenOfPath, type AssistantRouteName } from "@/features/assistant/tools/routes"
import { driveName, joinPath } from "@/features/assistant/tools/entities"

// What the user is looking at, read from the router and the stores the screens themselves render from.
export interface AssistantAppContext {
	pathname: string
	screen: AssistantRouteName | "other"
	// The open directory: on My Drive (null uuid is the root), or one opened on Shared with me or Shared by me,
	// whose `chain` is the route's uuid chain from that screen's root (Shared with me lists by it). `path` is a
	// My Drive path only; unresolved names (an ancestor not in any cached listing) leave it undefined, and
	// `name` is the directory's own name where only the breadcrumb resolved it.
	directory: {
		variant: DirectoryNameScope
		uuid: string | null
		chain: readonly string[]
		item: DriveItem | undefined
		path: string | undefined
		name: string | undefined
	} | null
	selection: DriveItem[]
	note: Note | undefined
	chat: Chat | undefined
	playing: { track: QueueTrack; status: AudioPlaybackStatus } | null
}

// Screens whose listing owns the drive selection; elsewhere a leftover selection means nothing.
const SELECTION_SCREENS: ReadonlySet<string> = new Set([
	"drive",
	"recents",
	"favorites",
	"shared_with_me",
	"shared_by_me",
	"public_links",
	"trash"
])

// The router has decoded the path already; a stray "%" it left in place must not throw.
function decodedSplat(splat: string): string {
	try {
		return decodeURIComponent(splat)
	} catch {
		return splat
	}
}

function cachedName(scope: DirectoryNameScope, uuid: string): string | undefined {
	const item = findCachedListingItem(uuid)

	return item === undefined ? (queryClient.getQueryData<string | null>(driveNamesQueryKey(scope, uuid)) ?? undefined) : driveName(item)
}

const SHARED_SCREENS: readonly (readonly [RegExp, DirectoryNameScope])[] = [
	[/^\/shared-in(?:\/(.*))?$/, "sharedIn"],
	[/^\/shared-out(?:\/(.*))?$/, "sharedOut"]
]

// A directory opened on Shared with me or Shared by me; the screen's root lists shares, not a directory.
function sharedDirectory(pathname: string): AssistantAppContext["directory"] {
	for (const [pattern, variant] of SHARED_SCREENS) {
		const match = pattern.exec(pathname)

		if (match === null) {
			continue
		}

		const chain = splatToUuids(decodedSplat(match[1] ?? "").replace(/\/+$/, ""))
		const uuid = chain.at(-1)

		return uuid === undefined
			? null
			: { variant, uuid, chain, item: findCachedListingItem(uuid), path: undefined, name: cachedName(variant, uuid) }
	}

	return null
}

function openDirectory(pathname: string): AssistantAppContext["directory"] {
	const match = /^\/drive(?:\/(.*))?$/.exec(pathname)

	if (match === null) {
		return sharedDirectory(pathname)
	}

	const uuids = splatToUuids(decodedSplat(match[1] ?? "").replace(/\/+$/, ""))
	const uuid = uuids.at(-1) ?? null
	let path: string | undefined = "/"

	// A directory opened by a link or a reveal often has no cached parent listing; the breadcrumb has
	// already resolved its name.
	for (const segment of uuids) {
		const name = cachedName("drive", segment)

		path = name === undefined || path === undefined ? undefined : joinPath(path, name)
	}

	return {
		variant: "drive",
		uuid,
		chain: uuids,
		item: uuid === null ? undefined : findCachedListingItem(uuid),
		path,
		name: uuid === null ? "/" : cachedName("drive", uuid)
	}
}

export function readAssistantContext(): AssistantAppContext {
	const pathname = assistantRouter()?.state.location.pathname ?? ""
	const screen = screenOfPath(pathname)
	const noteUuid = selectedNoteUuidFromPath(pathname)
	const chatUuid = selectedChatUuidFromPath(pathname)
	const audio = useAudioStore.getState()
	const track = audio.queue[audio.currentIndex]

	return {
		pathname,
		screen,
		directory: openDirectory(pathname),
		selection: SELECTION_SCREENS.has(screen) ? useDriveStore.getState().selectedItems : [],
		note: noteUuid === "" ? undefined : notesQueryFind(noteUuid),
		chat: chatUuid === "" ? undefined : chatsQueryGet()?.find(chat => chat.uuid === chatUuid),
		playing: track === undefined ? null : { track, status: audio.status }
	}
}
