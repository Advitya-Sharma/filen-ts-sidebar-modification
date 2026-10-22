import { useSyncExternalStore } from "react"
import type { SheetId } from "@filen/office/core/sheet"
import type { SheetGridModel } from "@filen/office/react/sheet"
import type { SheetDrawingObject } from "@filen/office/react/sheet/drawings"

// The drawing object selected in a grid (instead of cells), shared by the objects overlay and what acts on it (the
// grid's menu, a chart's side panel). Named by sheet and z-order key; the object itself is read from the sheet's
// layout as it now is.

export interface ObjectSelection {
	readonly sheet: SheetId
	readonly key: number
}

interface SelectionStore {
	selection: ObjectSelection | null
	readonly listeners: Set<() => void>
}

const stores = new WeakMap<SheetGridModel, SelectionStore>()

function storeOf(model: SheetGridModel): SelectionStore {
	let store = stores.get(model)

	if (store === undefined) {
		store = { selection: null, listeners: new Set() }
		stores.set(model, store)
	}

	return store
}

export function selectObject(model: SheetGridModel, selection: ObjectSelection | null): void {
	const store = storeOf(model)

	if (store.selection?.sheet === selection?.sheet && store.selection?.key === selection?.key) {
		return
	}

	store.selection = selection

	for (const listener of store.listeners) {
		listener()
	}
}

export function selectedObjectKey(model: SheetGridModel): ObjectSelection | null {
	return storeOf(model).selection
}

export function useObjectSelection(model: SheetGridModel): ObjectSelection | null {
	const store = storeOf(model)

	return useSyncExternalStore(
		listener => {
			store.listeners.add(listener)

			return () => {
				store.listeners.delete(listener)
			}
		},
		() => store.selection,
		() => store.selection
	)
}

// The selected object as the sheet on show now lists it; null when none is selected (cells are) or it is gone.
export function useSelectedObject(model: SheetGridModel): { readonly sheet: SheetId; readonly object: SheetDrawingObject } | null {
	const selection = useObjectSelection(model)
	const snapshot = useSyncExternalStore(model.subscribe, model.getSnapshot, model.getSnapshot)

	if (selection?.sheet !== snapshot.sheet) {
		return null
	}

	const object = snapshot.layout?.extensions?.drawings?.objects.find(candidate => candidate.key === selection.key)

	return object === undefined ? null : { sheet: selection.sheet, object }
}
