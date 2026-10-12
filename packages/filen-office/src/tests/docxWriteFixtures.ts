import type { StoryKind } from "@filen/office/core/docx"
import { DocxEditorSession } from "@filen/office/core/docx/edit"
import { writerProof } from "@filen/office/core/docx/write"
import type { Opened } from "./docxEditFixtures"
import { scriptOps, type ScriptTarget } from "./fuzz/ops/index"

// Scripted edits for writer tests: the deterministic edit of every op that has one (fuzz/ops), each on its
// own paragraph, in script order. Edits that find no suitable paragraph are skipped.

export interface EditScript {
	// The body's session.
	readonly session: DocxEditorSession
	// Every session the script edited, the body's first: what a save writes.
	readonly sessions: readonly DocxEditorSession[]
	// Names of the ops that applied.
	readonly done: string[]
}

export function runEditScript(opened: Opened): EditScript {
	const proof = writerProof(opened.context)
	const session = new DocxEditorSession(opened.context, { proof })
	const stories = new Map<StoryKind, DocxEditorSession | undefined>()
	const target: ScriptTarget = {
		session,
		story(kind) {
			if (!stories.has(kind)) {
				const story = opened.model.stories.find(candidate => candidate.kind === kind)

				stories.set(kind, story === undefined ? undefined : new DocxEditorSession(opened.context, { story, proof }))
			}

			return stories.get(kind)
		}
	}
	const done: string[] = []

	for (const op of scriptOps()) if (op.script?.run(target) === true) done.push(op.name)

	const sessions = [session]

	for (const other of stories.values()) if (other !== undefined) sessions.push(other)

	return { session, sessions, done }
}
