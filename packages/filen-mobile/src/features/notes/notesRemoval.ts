import { noteContentQueryKey } from "@/features/notes/queries/useNoteContent.query"
import { removeQueryEverywhere } from "@/queries/client"
import { notesQueryUpdate } from "@/features/notes/queries/useNotesQuery"
import useNotesStore from "@/features/notes/store/useNotes.store"
import useAppStore from "@/stores/useApp.store"

// Upper bound on how long a note left open in its editor stays listed after a local delete/leave.
const OPEN_NOTE_REMOVAL_DELAY_MS = 3000

// Drops a note this account no longer has from local state. The selection purge is always
// immediate so the header's selection count stays right and bulk ops can't target a ghost (#42).
export function dropNoteLocally(uuid: string, { deferListRemoval }: { deferListRemoval: boolean }): void {
	useNotesStore.getState().setSelectedNotes(prev => prev.filter(n => n.uuid !== uuid))

	const drop = () => {
		notesQueryUpdate({
			updater: prev => prev.filter(n => n.uuid !== uuid)
		})

		// removeQueryEverywhere, not `updater: () => undefined`: query-core's setQueryData bails out
		// when the updater yields undefined (`if (data === void 0) return`), so the old form was a
		// silent no-op and this note's decrypted body survived in memory and on disk. This also drops
		// the persisted row, which removeQueries alone would leave behind.
		removeQueryEverywhere(noteContentQueryKey({ uuid }))
	}

	const route = `/note/${uuid}`
	const onRoute = (pathname: string) => pathname.startsWith(route)

	// Only a local delete/leave from the note's own editor waits: removing the note at once fires that
	// route's redirect while confirmedAction is still popping it, which is janky and messes with the
	// navigation stack. Anywhere else (the list) the row goes at once, so it can't be acted on again.
	if (!deferListRemoval || !onRoute(useAppStore.getState().pathname)) {
		drop()

		return
	}

	// Dropped as soon as the editor is left, or after the delay if it never is.
	let unsubscribe: (() => void) | null = null
	const timer = setTimeout(() => {
		unsubscribe?.()
		drop()
	}, OPEN_NOTE_REMOVAL_DELAY_MS)

	unsubscribe = useAppStore.subscribe(state => {
		if (onRoute(state.pathname)) {
			return
		}

		unsubscribe?.()
		clearTimeout(timer)
		drop()
	})
}
