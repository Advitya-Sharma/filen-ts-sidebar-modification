import type { EditorState, Transaction } from "prosemirror-state"
import type { PagedEditor } from "../paged/editor"
import type { DocxCommandOf, DocxEditorCommandType } from "./commands"
import type { DocxEditContext } from "./context"
import { FEATURE_HANDLERS, FEATURE_TRANSFORMERS } from "./handlers/index"
import type { DocxEditorSession } from "./session"
import type { RevertibleUnits } from "./unsavable"

// Where feature commands run. A story command (scope "story") becomes a transaction of the session it is
// sent to; a document command (scope "document") runs on the paged editor, which may touch several stories.
// Each feature registers its handlers in handlers/<feature>.ts; a command without a handler is disabled and
// refused.

export interface StoryCommandContext {
	readonly session: DocxEditorSession
	readonly state: EditorState
	readonly context: DocxEditContext
}

// A transaction to apply, true when the command ran without one, null when it cannot run.
export type StoryCommandResult = Transaction | true | null

export interface StoryCommandHandler<T extends DocxEditorCommandType = DocxEditorCommandType> {
	// Whether the command can run in this state (snapshot.enabled): called with every snapshot, so cheap.
	enabled(context: StoryCommandContext): boolean
	run(context: StoryCommandContext, command: DocxCommandOf<T>): StoryCommandResult
}

export interface DocumentCommandContext {
	readonly editor: PagedEditor
	// The story being edited.
	readonly session: DocxEditorSession
	readonly context: DocxEditContext
}

export interface DocumentCommandHandler<T extends DocxEditorCommandType = DocxEditorCommandType> {
	enabled(context: DocumentCommandContext): boolean
	// Whether it ran; a promise for async commands (DocxCommandInfo.async).
	run(context: DocumentCommandContext, command: DocxCommandOf<T>): boolean | Promise<boolean>
}

export type StoryHandlers = { readonly [T in DocxEditorCommandType]?: StoryCommandHandler<T> }
export type DocumentHandlers = { readonly [T in DocxEditorCommandType]?: DocumentCommandHandler<T> }

export interface FeatureHandlers {
	readonly story?: StoryHandlers
	readonly document?: DocumentHandlers
	// Units a failed save takes back one by one: a comment, a note, a section's properties.
	readonly units?: RevertibleUnits
}

// Turns a user transaction into what is applied instead (tracked changes record edits this way). Called
// for every transaction that changes the document and is neither the session's own nor an undo or redo.
export type DocxTransactionTransformer = (tr: Transaction, session: DocxEditorSession) => Transaction

const storyHandlers = new Map<string, StoryCommandHandler>()
const documentHandlers = new Map<string, DocumentCommandHandler>()

for (const feature of FEATURE_HANDLERS) {
	for (const [type, handler] of Object.entries(feature.story ?? {})) storyHandlers.set(type, handler as StoryCommandHandler)
	for (const [type, handler] of Object.entries(feature.document ?? {})) documentHandlers.set(type, handler as DocumentCommandHandler)
}

export function storyHandler<T extends DocxEditorCommandType>(type: T): StoryCommandHandler<T> | undefined {
	return storyHandlers.get(type) as StoryCommandHandler<T> | undefined
}

export function documentHandler<T extends DocxEditorCommandType>(type: T): DocumentCommandHandler<T> | undefined {
	return documentHandlers.get(type) as DocumentCommandHandler<T> | undefined
}

// The registered transformers applied one after the other, or undefined without any.
export function registeredTransformer(): DocxTransactionTransformer | undefined {
	if (FEATURE_TRANSFORMERS.length === 0) return undefined

	return (tr, session) => FEATURE_TRANSFORMERS.reduce((current, transform) => transform(current, session), tr)
}
