import type { Node } from "prosemirror-model"
import type { DocxEditContext } from "../edit/context"
import type { DocxEditorSession } from "../edit/session"

// The editing sessions of each document (made and not disposed), and the host making the sessions of a
// document's other stories (a paged editor: what it makes, it lays out and saves). Features changing one story
// because of an edit of another (controls bound to one node) find the other story's session here.

const live = new WeakMap<DocxEditContext, Set<DocxEditorSession>>()
const hosts = new WeakMap<DocxEditContext, (story: number) => DocxEditorSession | undefined>()
// The session each state's document is the document of (a find and replace or a repair gets a state only).
const ofDoc = new WeakMap<Node, DocxEditorSession>()

export function sessionMade(context: DocxEditContext, session: DocxEditorSession): void {
	let sessions = live.get(context)

	if (sessions === undefined) {
		sessions = new Set()
		live.set(context, sessions)
	}

	sessions.add(session)
	ofDoc.set(session.state.doc, session)
	session.onApplied((_transactions, state) => ofDoc.set(state.doc, session))
}

// The open editing sessions of a document.
export function liveSessions(context: DocxEditContext): ReadonlySet<DocxEditorSession> {
	return live.get(context) ?? NONE
}

const NONE: ReadonlySet<DocxEditorSession> = new Set()

// The session whose state has `doc` (made or applied by it), or undefined.
export function sessionOfDoc(doc: Node): DocxEditorSession | undefined {
	return ofDoc.get(doc)
}

export function sessionGone(context: DocxEditContext, session: DocxEditorSession): void {
	live.get(context)?.delete(session)
}

// `host` makes the editing session of a story by index (made on first use, kept); returns the undo.
export function hostStories(context: DocxEditContext, host: (story: number) => DocxEditorSession | undefined): () => void {
	hosts.set(context, host)

	return () => {
		if (hosts.get(context) === host) hosts.delete(context)
	}
}

// The editing session of story `story` of a document: the host's (made if needed), else a live one of it.
export function storySessionOf(context: DocxEditContext, story: number): DocxEditorSession | undefined {
	const host = hosts.get(context)

	if (host !== undefined) return host(story)

	for (const session of live.get(context) ?? []) if (session.story.story.index === story) return session

	return undefined
}
