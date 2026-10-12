import type { SnapshotInput } from "./index"

// The table the selection is in: its size, the cells the selection touches and what table commands need.

export interface DocxTableContext {
	readonly rows: number
	readonly columns: number
	// The selection's cells as grid rectangle (rows and grid columns, end exclusive).
	readonly fromRow: number
	readonly toRow: number
	readonly fromColumn: number
	readonly toColumn: number
	// The selected cells can be merged, or the cell split.
	readonly mergeable: boolean
	readonly splittable: boolean
	readonly headerRow: boolean
	readonly styleId: string | null
}

export function tableOf(_input: SnapshotInput): DocxTableContext | null {
	return null
}
