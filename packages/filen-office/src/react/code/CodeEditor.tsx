import { useEffect, useEffectEvent, useLayoutEffect, useRef, useState } from "react"
import { Compartment, EditorState, type Extension } from "@codemirror/state"
import { EditorView } from "@codemirror/view"

export interface CodeEditorProps {
	// The document the editor opens with, read once at mount: a different document remounts the editor
	// (key it by the document's identity). The view's own doc is the buffer of record from then on.
	doc: string
	// The whole configuration (dom/code's setup, theme, language, keymaps). A new value reconfigures the
	// live view; the doc, selection and history stay.
	extensions: Extension
	// Names the editable content for assistive technology.
	label?: string
	autoFocus?: boolean
	className?: string
	// The view, once, right after it is created.
	onCreate?: (view: EditorView) => void
}

function labelled(extensions: Extension, label: string | undefined): Extension {
	return label === undefined ? extensions : [extensions, EditorView.contentAttributes.of({ "aria-label": label })]
}

// A headless CodeMirror editor: no styling, no value syncing. It holds one EditorView for its lifetime.
export function CodeEditor({ doc, extensions, label, autoFocus, className, onCreate }: CodeEditorProps) {
	const hostRef = useRef<HTMLDivElement>(null)
	const viewRef = useRef<EditorView | null>(null)
	// What the view holds now, so a commit that changes neither reconfigures nothing.
	const appliedRef = useRef<{ extensions: Extension; label: string | undefined } | null>(null)
	const [compartment] = useState(() => new Compartment())
	const [seed] = useState(doc)
	const mountConfiguration = useEffectEvent(() => ({ extensions, label }))
	const created = useEffectEvent((view: EditorView) => {
		onCreate?.(view)
	})

	useLayoutEffect(() => {
		const parent = hostRef.current

		if (parent === null) {
			return undefined
		}

		const initial = mountConfiguration()
		const view = new EditorView({
			state: EditorState.create({ doc: seed, extensions: compartment.of(labelled(initial.extensions, initial.label)) }),
			parent
		})

		viewRef.current = view
		appliedRef.current = initial
		created(view)

		return () => {
			viewRef.current = null
			appliedRef.current = null
			view.destroy()
		}
	}, [compartment, seed])

	useLayoutEffect(() => {
		const view = viewRef.current

		const applied = appliedRef.current

		if (view !== null && (applied === null || applied.extensions !== extensions || applied.label !== label)) {
			appliedRef.current = { extensions, label }
			view.dispatch({ effects: compartment.reconfigure(labelled(extensions, label)) })
		}
	}, [compartment, extensions, label])

	useEffect(() => {
		if (autoFocus === true) {
			viewRef.current?.focus()
		}
	}, [autoFocus])

	return (
		<div
			ref={hostRef}
			className={className}
		/>
	)
}
