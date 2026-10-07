import type { Chat, Note } from "@filen/sdk-rs"
import { splatToUuids } from "@/features/drive/lib/navigate"
import { findCachedListingItem } from "@/features/drive/queries/drive"
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
	// The open My Drive directory: null uuid is the root. Unresolved names (an ancestor not in any
	// cached listing) leave `path` undefined.
	directory: { uuid: string | null; item: DriveItem | undefined; path: string | undefined } | null
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

function myDriveDirectory(pathname: string): AssistantAppContext["directory"] {
	const match = /^\/drive(?:\/(.*))?$/.exec(pathname)

	if (match === null) {
		return null
	}

	const uuids = splatToUuids(decodeURIComponent(match[1] ?? "").replace(/\/+$/, ""))
	const uuid = uuids.at(-1) ?? null
	let path: string | undefined = "/"

	for (const segment of uuids) {
		const ancestor = findCachedListingItem(segment)

		path = ancestor === undefined || path === undefined ? undefined : joinPath(path, driveName(ancestor))
	}

	return { uuid, item: uuid === null ? undefined : findCachedListingItem(uuid), path }
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
		directory: myDriveDirectory(pathname),
		selection: SELECTION_SCREENS.has(screen) ? useDriveStore.getState().selectedItems : [],
		note: noteUuid === "" ? undefined : notesQueryFind(noteUuid),
		chat: chatUuid === "" ? undefined : chatsQueryGet()?.find(chat => chat.uuid === chatUuid),
		playing: track === undefined ? null : { track, status: audio.status }
	}
}
