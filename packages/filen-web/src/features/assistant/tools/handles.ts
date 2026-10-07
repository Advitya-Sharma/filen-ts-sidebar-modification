import type { Chat, Note } from "@filen/sdk-rs"
import type { Playlist } from "@filen/shared"
import { isDirectoryItem, type DriveItem } from "@/features/drive/lib/item"
import type { QueueTrack } from "@/features/audio/store/audioQueue"
import type { AssistantHandleRef } from "@/features/assistant/runtime/executor"

// What a handle stands for. Snapshots as of when the model saw them; adapters re-read the live row
// from the caches before acting. `path` is the drive item's absolute path when a lookup knew it.
export type AssistantEntity =
	| { kind: "root" }
	| { kind: "drive"; item: DriveItem; path?: string }
	| { kind: "note"; note: Note }
	| { kind: "chat"; chat: Chat }
	| { kind: "playlist"; playlist: Playlist }
	| { kind: "track"; track: QueueTrack }

export type AssistantEntityKind = AssistantEntity["kind"]

export interface AssistantHandleEntry {
	ref: AssistantHandleRef
	entity: AssistantEntity
}

export const ROOT_HANDLE = "d0"

const ROOT_NAME = "/"

function entityKey(entity: AssistantEntity): string {
	switch (entity.kind) {
		case "root":
			return "root"
		case "drive":
			return `drive:${entity.item.data.uuid}`
		case "note":
			return `note:${entity.note.uuid}`
		case "chat":
			return `chat:${entity.chat.uuid}`
		case "playlist":
			return `playlist:${entity.playlist.uuid}`
		case "track":
			return `track:${entity.track.uuid}`
	}
}

function handlePrefix(entity: AssistantEntity): string {
	switch (entity.kind) {
		case "root":
			return "d"
		case "drive":
			return isDirectoryItem(entity.item) ? "d" : "f"
		case "note":
			return "n"
		case "chat":
			return "c"
		case "playlist":
			return "p"
		case "track":
			return "t"
	}
}

// Short handles for one conversation: a model copies "f3" reliably and mangles a UUID. One item keeps
// its handle for the whole conversation, so a handle from an earlier turn still names the same item.
export class AssistantHandleRegistry {
	private readonly byHandle = new Map<string, AssistantHandleEntry>()
	private readonly byKey = new Map<string, string>()
	private readonly counters = new Map<string, number>()

	constructor() {
		this.clear()
	}

	register(entity: AssistantEntity, name: string): AssistantHandleRef {
		const key = entityKey(entity)
		const known = this.byKey.get(key)

		if (known !== undefined) {
			const previous = this.byHandle.get(known)?.entity
			// A later sighting without a path keeps the one an earlier lookup found.
			const kept =
				entity.kind === "drive" && entity.path === undefined && previous?.kind === "drive" && previous.path !== undefined
					? { ...entity, path: previous.path }
					: entity
			const ref = { handle: known, name }

			this.byHandle.set(known, { ref, entity: kept })

			return ref
		}

		const prefix = handlePrefix(entity)
		const next = (this.counters.get(prefix) ?? 0) + 1
		const ref = { handle: `${prefix}${String(next)}`, name }

		this.counters.set(prefix, next)
		this.byKey.set(key, ref.handle)
		this.byHandle.set(ref.handle, { ref, entity })

		return ref
	}

	get(handle: string): AssistantHandleEntry | undefined {
		return this.byHandle.get(handle.trim())
	}

	entries(): AssistantHandleEntry[] {
		return [...this.byHandle.values()]
	}

	clear(): void {
		this.byHandle.clear()
		this.byKey.clear()
		this.counters.clear()
		// The drive root is always "d0", as the model is told.
		this.byKey.set("root", ROOT_HANDLE)
		this.byHandle.set(ROOT_HANDLE, { ref: { handle: ROOT_HANDLE, name: ROOT_NAME }, entity: { kind: "root" } })
	}
}
