import { useState, useSyncExternalStore } from "react"
import type { DocxEditorCommand, DocxEditorCommandType, DocxEditorSnapshot } from "../../core/docx/edit/index"
import type { CaretBox, PagedEditor, PagedSelection } from "../../core/docx/paged/index"

// The headless state of a paged editor for an app's toolbar, menus and touch UI: the session's snapshot
// (formatting, enablement, undo, read-only), the selection with its two ends as drawn (where selection handles
// and a cut/copy/paste menu go), composition and layout progress; plus the session's commands. The same data
// crosses a worker or WebView bridge unchanged.

export interface DocxEditorState {
	readonly snapshot: DocxEditorSnapshot
	readonly selection: PagedSelection
	// The caret as drawn (an empty selection), in page coordinates.
	readonly caret: CaretBox | undefined
	readonly ends: { readonly start: CaretBox | undefined; readonly end: CaretBox | undefined }
	readonly composing: boolean
	readonly pageCount: number
	// Layout has caught up with every edit.
	readonly settled: boolean
}

export interface DocxEditorControls {
	readonly state: DocxEditorState | undefined
	execute(command: DocxEditorCommand): boolean
	canExecute(type: DocxEditorCommandType): boolean
}

class EditorStateStore {
	private readonly editor: PagedEditor
	private version = 0
	private cached: { version: number; state: DocxEditorState } | undefined = undefined

	constructor(editor: PagedEditor) {
		this.editor = editor
	}

	subscribe = (listener: () => void): (() => void) => {
		const bump = (): void => {
			this.version++
			listener()
		}
		const stopEditor = this.editor.subscribe(bump)
		const stopSession = this.editor.session.subscribe(bump)

		return () => {
			stopEditor()
			stopSession()
		}
	}

	getSnapshot = (): DocxEditorState => {
		const cached = this.cached

		if (cached?.version === this.version) return cached.state

		const editor = this.editor
		const selection = editor.selection()
		const state: DocxEditorState = {
			snapshot: editor.session.snapshot(),
			selection,
			caret: selection.empty ? editor.caret() : undefined,
			ends: editor.ends(),
			composing: editor.composition !== undefined,
			pageCount: editor.pageCount,
			settled: editor.settled
		}

		this.cached = { version: this.version, state }

		return state
	}
}

function subscribeNothing(): () => void {
	return () => undefined
}

function nothing(): undefined {
	return undefined
}

export function useDocxEditor(editor: PagedEditor | undefined): DocxEditorControls {
	const [entry, setEntry] = useState(() => ({ editor, store: editor === undefined ? undefined : new EditorStateStore(editor) }))
	let current = entry

	if (entry.editor !== editor) {
		current = { editor, store: editor === undefined ? undefined : new EditorStateStore(editor) }
		setEntry(current)
	}

	const store = current.store
	const state = useSyncExternalStore<DocxEditorState | undefined>(
		store?.subscribe ?? subscribeNothing,
		store?.getSnapshot ?? nothing,
		store?.getSnapshot ?? nothing
	)

	return {
		state,
		execute: command => editor?.session.execute(command) ?? false,
		canExecute: type => editor?.session.canExecute(type) ?? false
	}
}
