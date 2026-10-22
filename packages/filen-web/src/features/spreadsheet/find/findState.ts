import { useSyncExternalStore } from "react"
import type { SheetGridModel } from "@filen/office/react/sheet"

// Whether a grid's find bar is open, and how (Mod+F finds, Mod+H also replaces), shared by what opens it (the grid's
// shortcut, the toolbar's Find button) and the bar. `focus` grows each time it is asked to open, so an open bar
// takes focus again.
export interface SheetFindOpen {
	readonly mode: "find" | "replace"
	readonly focus: number
}

interface FindStore {
	open: SheetFindOpen | null
	// The open bar's Find Next and Previous: false when it has no matches to step through.
	step: ((direction: 1 | -1) => boolean) | null
	readonly listeners: Set<() => void>
}

const stores = new WeakMap<SheetGridModel, FindStore>()

function storeOf(model: SheetGridModel): FindStore {
	let store = stores.get(model)

	if (store === undefined) {
		store = { open: null, step: null, listeners: new Set() }
		stores.set(model, store)
	}

	return store
}

function set(model: SheetGridModel, open: SheetFindOpen | null): void {
	const store = storeOf(model)

	store.open = open

	for (const listener of store.listeners) {
		listener()
	}
}

export function openSheetFind(model: SheetGridModel, mode: SheetFindOpen["mode"]): void {
	const current = storeOf(model).open

	set(model, { mode: current?.mode === "replace" ? "replace" : mode, focus: (current?.focus ?? 0) + 1 })
}

// Shift+F4 and Ctrl+Shift+F4 (Excel for the web's Find Next and Previous): steps the open bar's matches, or opens the
// bar when it has none.
export function stepSheetFind(model: SheetGridModel, direction: 1 | -1): void {
	if (storeOf(model).step?.(direction) !== true) {
		openSheetFind(model, "find")
	}
}

// The open bar's stepping, while it is open; returns its removal.
export function setSheetFindStep(model: SheetGridModel, step: (direction: 1 | -1) => boolean): () => void {
	const store = storeOf(model)

	store.step = step

	return () => {
		if (store.step === step) {
			store.step = null
		}
	}
}

export function closeSheetFind(model: SheetGridModel): void {
	if (storeOf(model).open !== null) {
		set(model, null)
	}
}

export function useSheetFindOpen(model: SheetGridModel): SheetFindOpen | null {
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
