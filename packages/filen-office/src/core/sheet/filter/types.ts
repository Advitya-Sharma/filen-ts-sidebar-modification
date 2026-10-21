import type { SheetRange } from "../protocol"
import type { SheetFilterCriteria } from "./criteria"

// What the filter feature shows a client: a sheet's filter ranges with their buttons (the "filters" layout
// extension) and a column's values list (the "filter.values" query).

export const FilterButton = {
	// The column shows a button (hiddenButton and showButton allow it).
	SHOWN: 1 << 0,
	FILTERED: 1 << 1,
	ASCENDING: 1 << 2,
	DESCENDING: 1 << 3
} as const

export interface SheetFilterRange {
	// The sheet's own autofilter, or a table's.
	readonly table: string | null
	// Its first row is the header row the buttons are on.
	readonly range: SheetRange
	// FilterButton bits per column of the range.
	readonly buttons: Uint8Array
}

export interface SheetFilterLayout {
	readonly ranges: readonly SheetFilterRange[]
}

export interface SheetFilterItem {
	// As the cells show it.
	readonly text: string
	readonly checked: boolean
	// A date's day, for a list that names dates by day (dateGroupItem).
	readonly date?: { readonly year: number; readonly month: number; readonly day: number }
}

export interface SheetFilterValues {
	readonly table: string | null
	readonly range: SheetRange
	// The header cell's text.
	readonly header: string
	readonly items: readonly SheetFilterItem[]
	// The column holds blank cells (in the rows other filters let through): whether they are let through.
	readonly blank: { readonly checked: boolean } | null
	// More values than `limit`.
	readonly more: boolean
	readonly criteria: SheetFilterCriteria | null
	// Criteria kept as the file has them that this app does not work out.
	readonly unknown: boolean
	// What most of the column holds: which conditions the menu offers (text, number or date filters).
	readonly kind: "text" | "number" | "date"
	// The column's sort direction in the filter's sort state, if any.
	readonly sorted: "ascending" | "descending" | null
}

declare module "../protocol" {
	interface SheetCommandMap {
		// Turns the sheet's autofilter off (its rows shown again) when it has one; else turns a table's filter buttons
		// on or off when `active` is in a table; else puts an autofilter on `range` (its first row the header).
		"filter.toggle": {
			readonly sheet: SheetId
			readonly range: SheetRange
			readonly active: { readonly row: number; readonly col: number }
		}
		// The criteria of column `col` of the filter holding the cell (null clears them), then the range filtered again.
		"filter.column": {
			readonly sheet: SheetId
			readonly row: number
			readonly col: number
			readonly criteria: SheetFilterCriteria | null
		}
		// Every criteria of the filter holding the cell cleared (its buttons kept), its rows shown.
		"filter.clear": { readonly sheet: SheetId; readonly row: number; readonly col: number }
		// The filter holding the cell applied again to the cells as they now are.
		"filter.reapply": { readonly sheet: SheetId; readonly row: number; readonly col: number }
	}

	interface SheetQueryMap {
		// The values of column `col` of the filter holding the cell, as its list offers them: the rows the range's
		// other filters let through, each value once, in sort order, at most `limit`.
		"filter.values": {
			request: { readonly sheet: SheetId; readonly row: number; readonly col: number; readonly limit: number }
			result: SheetFilterValues | null
		}
	}

	interface SheetLayoutExtensionMap {
		// The sheet's filter ranges, with a button state per column.
		readonly filters?: SheetFilterLayout
	}

	interface SheetRefusalReasons {
		// Criteria this app cannot work out (icons, colours under conditional formatting, a newer kind): the filter
		// is kept as the file has it.
		filterCriteria: true
	}
}
