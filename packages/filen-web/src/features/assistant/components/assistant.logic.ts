import { formatBytesFixed } from "@filen/shared"
import type { AssistantKey } from "@/lib/i18n"
import type { TranscriptEntry } from "@/features/assistant/store/transcript"
import type { AssistantModelState, AssistantTurnError } from "@/features/assistant/types"

export type TranscriptRow = TranscriptEntry | { kind: "thinking"; id: "thinking" }

export function downloadPercent(receivedBytes: number, sizeBytes: number): number {
	if (sizeBytes <= 0) {
		return 0
	}

	return Math.min(100, Math.max(0, Math.floor((receivedBytes / sizeBytes) * 100)))
}

// Both figures in one format, the received part capped at the total: two precisions (one decimal against
// two) read as "2.6 GiB of 2.55 GiB" in the last few megabytes.
export function downloadByteFigures(receivedBytes: number, sizeBytes: number): { received: string; total: string } {
	return {
		received: formatBytesFixed(Math.min(sizeBytes, Math.max(0, receivedBytes))),
		total: formatBytesFixed(sizeBytes)
	}
}

// Files on disk: what "Remove downloaded model" can delete. A failed load still left them there; a
// failed or full download is the runtime's to clean up.
export function canRemoveModel(state: AssistantModelState): boolean {
	switch (state.kind) {
		case "downloaded":
		case "ready":
			return true
		case "error":
			return state.error === "load-failed" || state.error === "out-of-memory"
		default:
			return false
	}
}

// The waiting bubble shows while a turn has nothing visibly in progress: after the user's message, or
// between a finished step and the next one. A streaming reply or a running activity is its own signal.
export function transcriptRows(entries: readonly TranscriptEntry[], generating: boolean): readonly TranscriptRow[] {
	if (!generating) {
		return entries
	}

	const last = entries.at(-1)

	if ((last?.kind === "reply" && !last.final) || (last?.kind === "activity" && last.status === "running")) {
		return entries
	}

	return [...entries, { kind: "thinking", id: "thinking" }]
}

export const TURN_END_KEYS: Record<"cancelled" | AssistantTurnError, AssistantKey> = {
	cancelled: "assistantTurnCancelled",
	"model-unavailable": "assistantTurnErrorModelUnavailable",
	"context-full": "assistantTurnErrorContextFull",
	failed: "assistantTurnErrorFailed"
}

// Keyed by the registry names in @filen/shared/assistant (tools.ts). A Map, so a name the model or a
// newer runtime makes up never resolves through the object prototype.
const TOOL_ACTIVITY_KEYS = new Map<string, AssistantKey>([
	["get_context", "assistantToolGetContext"],
	["search_items", "assistantToolSearchItems"],
	["list_directory", "assistantToolListDirectory"],
	["get_item_info", "assistantToolGetItemInfo"],
	["read_text_file", "assistantToolReadTextFile"],
	["reveal_item", "assistantToolRevealItem"],
	["set_favorite", "assistantToolSetFavorite"],
	["create_directory", "assistantToolCreateDirectory"],
	["rename_item", "assistantToolRenameItem"],
	["move_items", "assistantToolMoveItems"],
	["copy_items", "assistantToolCopyItems"],
	["trash_items", "assistantToolTrashItems"],
	["list_notes", "assistantToolListNotes"],
	["read_note", "assistantToolReadNote"],
	["open_note", "assistantToolOpenNote"],
	["create_note", "assistantToolCreateNote"],
	["set_note_flags", "assistantToolSetNoteFlags"],
	["list_chats", "assistantToolListChats"],
	["read_chat", "assistantToolReadChat"],
	["open_chat", "assistantToolOpenChat"],
	["draft_chat_message", "assistantToolDraftChatMessage"],
	["list_playlists", "assistantToolListPlaylists"],
	["open_playlist", "assistantToolOpenPlaylist"],
	["play", "assistantToolPlay"],
	["audio_control", "assistantToolAudioControl"],
	["add_tracks_to_playlist", "assistantToolAddTracksToPlaylist"],
	["open_route", "assistantToolOpenRoute"]
])

export function toolActivityKey(tool: string): AssistantKey {
	return TOOL_ACTIVITY_KEYS.get(tool) ?? "assistantToolGeneric"
}

export type ProposalTargetKind = "file" | "directory" | "note" | "chat" | "playlist" | "track"

const TARGET_KINDS = new Map<string, ProposalTargetKind>([
	["f", "file"],
	["d", "directory"],
	["n", "note"],
	["c", "chat"],
	["p", "playlist"],
	["t", "track"]
])

// The handle's leading letter is its kind (runtime/executor.ts); only the icon depends on it, so an
// unexpected handle still renders, as a file.
export function proposalTargetKind(handle: string): ProposalTargetKind {
	return TARGET_KINDS.get(handle.charAt(0)) ?? "file"
}

// Markdown punctuation a screen reader would otherwise spell out.
export function spokenText(markdown: string): string {
	return markdown.replace(/[*`]/g, "").replace(/^[#>]+\s*/gm, "")
}
