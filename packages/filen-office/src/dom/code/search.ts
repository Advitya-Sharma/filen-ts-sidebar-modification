import { openSearchPanel, searchKeymap } from "@codemirror/search"
import { Prec, type Extension } from "@codemirror/state"
import { keymap, type Command, type KeyBinding } from "@codemirror/view"

// The search panel with its replace field focused; a read-only editor's panel has none and keeps the
// search field.
export const openReplacePanel: Command = view => {
	openSearchPanel(view)

	const replaceField = view.dom.querySelector<HTMLInputElement>(".cm-search input[name=replace]")

	replaceField?.focus()
	replaceField?.select()

	return true
}

export { openSearchPanel }

// CodeMirror's own search bindings but Mod-f, which an app binds to its own (rebindable) find.
const SEARCH_BINDINGS: readonly KeyBinding[] = searchKeymap.filter(binding => binding.key !== "Mod-f")

// Bindings as CodeMirror key names; a scope limits one to those parts of the editor.
export interface CodeBinding {
	readonly keys: readonly string[]
	readonly run: Command
	readonly scope?: string
}

// The editor scope, and the editor plus its search panel's own fields (find and replace answer there too).
export const CODE_EDITOR_SCOPE = "editor"
export const CODE_SEARCH_SCOPE = "editor search-panel"

// The app's bindings ahead of CodeMirror's own, followed by the search panel's.
export function codeKeymap(bindings: readonly CodeBinding[]): Extension {
	const keyBindings: KeyBinding[] = []

	for (const binding of bindings) {
		for (const key of binding.keys) {
			keyBindings.push(binding.scope === undefined ? { key, run: binding.run } : { key, run: binding.run, scope: binding.scope })
		}
	}

	keyBindings.push(...SEARCH_BINDINGS)

	return Prec.high(keymap.of(keyBindings))
}
