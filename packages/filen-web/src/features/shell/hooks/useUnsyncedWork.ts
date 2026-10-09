import { useSyncExternalStore } from "react"
import { useNotesInflightStore } from "@/features/notes/store/useNotesInflight"
import { useChatsInflightStore } from "@/features/chats/store/useChatsInflight"
import { hasUnsyncedNotes, hasUnsyncedChatSends } from "@/features/shell/hooks/useUnsyncedWork.logic"
import { watchState } from "@/features/videos/lib/watchState"

function watchStateDirty(): boolean {
	return watchState.getSnapshot().dirty
}

// True while either durable outbox, or the Videos watch-state journal, still holds something the server has
// never seen. Sign-out cancels them and wipes their storage (performLogout), so the confirm reads this to warn
// instead of promising the content syncs back. Boolean-collapsed selectors, so typing in a note or a composer
// re-renders the subscriber only on the has/has-not edge.
export function useHasUnsyncedWork(): boolean {
	const notes = useNotesInflightStore(state => hasUnsyncedNotes(state.inflightContent))
	const chats = useChatsInflightStore(state => hasUnsyncedChatSends(state.inflightMessages, state.inflightErrors))

	const watch = useSyncExternalStore(watchState.subscribe, watchStateDirty)

	return notes || chats || watch
}
