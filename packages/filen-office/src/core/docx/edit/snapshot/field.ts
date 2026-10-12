import type { SnapshotInput } from "./index"

// The field the selection is in or on.

export interface DocxFieldContext {
	// The field type (PAGE, REF, TOC...).
	readonly type: string
	readonly updatable: boolean
}

export function fieldOf(_input: SnapshotInput): DocxFieldContext | null {
	return null
}
