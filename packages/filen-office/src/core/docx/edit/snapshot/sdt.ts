import type { SnapshotInput } from "./index"

// The content control the selection is in.

export interface DocxControlContext {
	readonly type: "checkbox" | "dropDownList" | "comboBox" | "date" | "plainText" | "richText" | "picture" | "other"
	readonly checked: boolean | null
	// List entries: display text and value.
	readonly items: readonly { readonly text: string; readonly value: string }[]
	readonly value: string | null
	// ISO date of a date control.
	readonly date: string | null
	readonly placeholder: boolean
	// Why it cannot be changed here (bound to custom XML the editor cannot update...), or null.
	readonly locked: string | null
}

export function controlOf(_input: SnapshotInput): DocxControlContext | null {
	return null
}
