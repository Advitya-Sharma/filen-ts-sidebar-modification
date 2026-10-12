import type { SnapshotInput } from "./index"

// The equation the selection is.

export interface DocxEquationContext {
	readonly latex: string
	// Editing it as LaTeX loses formatting the equation has.
	readonly lossy: boolean
}

export function equationOf(_input: SnapshotInput): DocxEquationContext | null {
	return null
}
