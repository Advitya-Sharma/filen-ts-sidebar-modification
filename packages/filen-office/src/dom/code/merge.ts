import { MergeView, unifiedMergeView } from "@codemirror/merge"
import { EditorState, type Extension } from "@codemirror/state"
import { EditorView } from "@codemirror/view"

// Unchanged stretches fold away so the differences stay in view, and a diff that runs long falls back to
// a coarser one rather than holding the main thread on two large, very different files.
const COLLAPSE_UNCHANGED = { margin: 3, minSize: 4 }
const DIFF_CONFIG = { scanLimit: 10_000, timeout: 500 }

export interface CodeCompareOptions {
	readonly parent: HTMLElement
	// The other version, and the one being kept.
	readonly theirs: string
	readonly mine: string
	// Two columns, or one with both versions interleaved.
	readonly sideBySide: boolean
	// Theme and language for both sides.
	readonly extensions: Extension
}

// A read-only comparison mounted into `parent`; the returned function destroys it.
export function mountCodeCompare({ parent, theirs, mine, sideBySide, extensions }: CodeCompareOptions): () => void {
	const shared: Extension = [EditorState.readOnly.of(true), EditorView.editable.of(false), EditorView.lineWrapping, extensions]

	if (sideBySide) {
		const view = new MergeView({
			a: { doc: theirs, extensions: shared },
			b: { doc: mine, extensions: shared },
			parent,
			gutter: true,
			collapseUnchanged: COLLAPSE_UNCHANGED,
			diffConfig: DIFF_CONFIG
		})

		return () => {
			view.destroy()
		}
	}

	const view = new EditorView({
		doc: mine,
		extensions: [
			shared,
			unifiedMergeView({
				original: theirs,
				mergeControls: false,
				gutter: true,
				collapseUnchanged: COLLAPSE_UNCHANGED,
				diffConfig: DIFF_CONFIG
			})
		],
		parent
	})

	return () => {
		view.destroy()
	}
}
