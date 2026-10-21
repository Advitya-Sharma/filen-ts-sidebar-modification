import { useSyncExternalStore } from "react"
import type { SheetFilterCriteria, SheetRange } from "@filen/office/core/sheet"
import type { SheetGridModel } from "@filen/office/react/sheet"
import type { DialogOperator, FilterKind } from "@/features/spreadsheet/next/data/sortFilter.logic"

// What the sort and filter UI has open on a grid, shared by what opens it (a header's button, the toolbar's menu,
// the grid's menu, the shortcuts) and the extension that renders it: the filter menu of a header, the custom sort
// dialog, a condition dialog or the Top 10 dialog. One at a time.

export type SortFilterOpen =
	| { readonly type: "menu"; readonly row: number; readonly col: number }
	| { readonly type: "sort"; readonly range: SheetRange; readonly header: boolean; readonly key: number }
	| {
			readonly type: "condition"
			readonly row: number
			readonly col: number
			readonly header: string
			readonly kind: FilterKind
			readonly first: DialogOperator
			readonly second?: DialogOperator
			readonly criteria: SheetFilterCriteria | null
	  }
	| {
			readonly type: "top10"
			readonly row: number
			readonly col: number
			readonly header: string
			readonly criteria: SheetFilterCriteria | null
	  }

interface Store {
	open: SortFilterOpen | null
	readonly listeners: Set<() => void>
}

const stores = new WeakMap<SheetGridModel, Store>()

function storeOf(model: SheetGridModel): Store {
	let store = stores.get(model)

	if (store === undefined) {
		store = { open: null, listeners: new Set() }
		stores.set(model, store)
	}

	return store
}

export function setSortFilterOpen(model: SheetGridModel, open: SortFilterOpen | null): void {
	const store = storeOf(model)

	if (store.open === open) {
		return
	}

	store.open = open

	for (const listener of store.listeners) {
		listener()
	}
}

export function useSortFilterOpen(model: SheetGridModel): SortFilterOpen | null {
	const store = storeOf(model)

	return useSyncExternalStore(
		listener => {
			store.listeners.add(listener)

			return () => {
				store.listeners.delete(listener)
			}
		},
		() => store.open,
		() => store.open
	)
}
