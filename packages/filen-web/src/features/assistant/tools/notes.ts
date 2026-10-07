import type { Note } from "@filen/sdk-rs"
import { queryClient } from "@/queries/client"
import { fetchNotes, NOTES_QUERY_KEY } from "@/features/notes/queries/notes"
import { noteContentQueryRead } from "@/features/notes/queries/noteContent"
import { isNoteUndecryptable, sortNotes, tagDisplayName } from "@/features/notes/lib/sort"
import { filterNotesByBlockedOwner } from "@/features/notes/components/notesSidebar.logic"
import { selectedNoteUuidFromPath } from "@/features/notes/lib/route"
import { noteIsEditing, useNotesInflightStore } from "@/features/notes/store/useNotesInflight"
import { checklistToMarkdown } from "@/features/notes/lib/export.logic"
import { tabNoteContent } from "@/features/notes/lib/tabEditors"
import { assistantRouter } from "@/features/assistant/tools/navigation"
import {
	blockedUsers,
	freshNote,
	noteTitle,
	optionalString,
	requiredString,
	resolveHandle,
	type ToolAdapter,
	type ToolScope
} from "@/features/assistant/tools/entities"
import { AssistantToolError, clipText, isoDate, normalizeText, searchTokens, tokenMatches } from "@/features/assistant/tools/format"

const LIST_LIMIT = 40
export const NOTE_MAX_CHARS = 4_000

// The list the notes sidebar renders from, cached first: socket events keep it current.
function readNotes(): Promise<Note[]> {
	return queryClient.query({ queryKey: NOTES_QUERY_KEY, queryFn: fetchNotes, staleTime: "static" })
}

async function resolveNote(scope: ToolScope, handle: string): Promise<Note> {
	const entity = await resolveHandle(scope, handle, ["note"])

	if (entity.kind !== "note") {
		throw new AssistantToolError("Not a note", { invalid: true })
	}

	const note = freshNote(entity.note)

	if (note.trash) {
		throw new AssistantToolError(`"${noteTitle(note)}" is in the trash`)
	}

	return note
}

// Block breaks become line breaks, then the markup goes: what the reader would show, as plain text.
function richTextToPlain(html: string): string {
	const marked = html.replace(/<\/(p|li|h[1-6]|div|blockquote|pre)>|<br\s*\/?>/gi, "$&\n")

	return new DOMParser()
		.parseFromString(marked, "text/html")
		.body.textContent.replace(/\n{3,}/g, "\n\n")
		.trim()
}

function plainContent(note: Note, content: string): string {
	switch (note.noteType) {
		case "rich":
			return richTextToPlain(content)
		case "checklist":
			return checklistToMarkdown(content)
		case "text":
		case "md":
		case "code":
			return content
	}
}

export const listNotesTool: ToolAdapter = async (input, scope) => {
	const tokens = searchTokens(optionalString(input, "query") ?? "")
	const notes = sortNotes(filterNotesByBlockedOwner(await readNotes(), blockedUsers())).filter(note => {
		if (note.trash) {
			return false
		}

		if (tokens.length === 0) {
			return true
		}

		const haystack = normalizeText(`${noteTitle(note)} ${note.tags.map(tagDisplayName).join(" ")}`)

		return tokens.some(token => tokenMatches(haystack, token))
	})

	return {
		value: {
			notes: notes.slice(0, LIST_LIMIT).map(note => {
				const title = noteTitle(note)
				const tags = note.tags.map(tagDisplayName)

				return {
					handle: scope.show({ kind: "note", note }, title),
					title,
					type: note.noteType,
					edited: isoDate(note.editedTimestamp),
					...(tags.length > 0 ? { tags } : {}),
					...(note.pinned ? { pinned: true } : {}),
					...(note.favorite ? { favorite: true } : {}),
					...(note.archive ? { archived: true } : {})
				}
			}),
			...(notes.length > LIST_LIMIT ? { more: notes.length - LIST_LIMIT } : {})
		}
	}
}

// A note open in the editor or holding unsynced edits is read from what this tab knows, the editor's
// typed text first (tabNoteContent), never fetched: a read landing would advance the editor's remount
// key under the caret, and the server copy is older than what the user typed. Reading writes nothing.
export const readNoteTool: ToolAdapter = async (input, scope) => {
	const note = await resolveNote(scope, requiredString(input, "handle"))
	const title = noteTitle(note)
	const pathname = assistantRouter()?.state.location.pathname ?? ""

	if (isNoteUndecryptable(note)) {
		throw new AssistantToolError("This note cannot be decrypted")
	}

	const live = selectedNoteUuidFromPath(pathname) === note.uuid || noteIsEditing(useNotesInflightStore.getState(), note.uuid)
	const content = live ? tabNoteContent(note.uuid) : await noteContentQueryRead(note)

	if (content === undefined) {
		throw new AssistantToolError(`"${title}" is still loading in the editor; try again in a moment`)
	}

	const { text, truncated } = clipText(plainContent(note, content), NOTE_MAX_CHARS)

	return {
		value: {
			handle: scope.show({ kind: "note", note }, title),
			title,
			content: text,
			...(truncated ? { truncated: true } : {})
		}
	}
}

export const openNoteTool: ToolAdapter = async (input, scope) => {
	const note = await resolveNote(scope, requiredString(input, "handle"))
	const router = assistantRouter()

	if (router === null) {
		throw new AssistantToolError("The app cannot navigate right now")
	}

	await router.navigate({ to: "/notes/$uuid", params: { uuid: note.uuid } })

	return { value: { opened: noteTitle(note) }, uiEffect: { type: "navigated", pathname: `/notes/${note.uuid}` } }
}
