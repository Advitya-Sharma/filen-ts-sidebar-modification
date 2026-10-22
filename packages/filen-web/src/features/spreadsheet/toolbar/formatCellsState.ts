import { useSyncExternalStore } from "react"
import { borderPresetEdges, type SheetCommand, type SheetId, type SheetRange } from "@filen/office/core/sheet"
import type { SheetGridModel } from "@filen/office/react/sheet"

// Whether a grid's Format Cells dialog is open, shared by what opens it (Ctrl+1, the toolbar's menus) and the
// toolbar that holds it.

interface DialogStore {
	open: boolean
	readonly listeners: Set<() => void>
}

const stores = new WeakMap<SheetGridModel, DialogStore>()

function storeOf(model: SheetGridModel): DialogStore {
	let store = stores.get(model)

	if (store === undefined) {
		store = { open: false, listeners: new Set() }
		stores.set(model, store)
	}

	return store
}

export function setFormatCellsOpen(model: SheetGridModel, open: boolean): void {
	const store = storeOf(model)

	if (store.open === open) {
		return
	}

	store.open = open

	for (const listener of store.listeners) {
		listener()
	}
}

export function useFormatCellsOpen(model: SheetGridModel): boolean {
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

// Excel's border shortcuts: Ctrl+Shift+& puts an outline round the selection, Ctrl+Shift+_ takes its borders off
// ("Keyboard shortcuts in Excel", support.microsoft.com).
export function borderShortcut(shortcut: "borderOutline" | "borderNone", sheet: SheetId, ranges: readonly SheetRange[]): SheetCommand {
	return { type: "format.borders", sheet, ranges, edges: borderPresetEdges(shortcut === "borderOutline" ? "outside" : "none") }
}
