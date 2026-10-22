import { useEffect, useSyncExternalStore } from "react"
import type { SheetGridModel } from "@filen/office/react/sheet"
import { useLatestRef } from "@/lib/useLatestRef"

// A picture asked for from outside the objects feature (the toolbar's Insert menu): from this device's file chooser
// or the drive's picker. The objects feature takes the request and clears it.

export type ImageSource = "file" | "drive"

export interface ImageRequest {
	readonly source: ImageSource
	readonly nonce: number
}

interface RequestStore {
	request: ImageRequest | null
	nonce: number
	readonly listeners: Set<() => void>
}

const stores = new WeakMap<SheetGridModel, RequestStore>()

function storeOf(model: SheetGridModel): RequestStore {
	let store = stores.get(model)

	if (store === undefined) {
		store = { request: null, nonce: 0, listeners: new Set() }
		stores.set(model, store)
	}

	return store
}

function notify(store: RequestStore): void {
	for (const listener of store.listeners) {
		listener()
	}
}

export function requestImage(model: SheetGridModel, source: ImageSource): void {
	const store = storeOf(model)

	store.nonce++
	store.request = { source, nonce: store.nonce }
	notify(store)
}

export function clearImageRequest(model: SheetGridModel, nonce: number): void {
	const store = storeOf(model)

	if (store.request?.nonce !== nonce) {
		return
	}

	store.request = null
	notify(store)
}

export function useImageRequest(model: SheetGridModel): ImageRequest | null {
	const store = storeOf(model)

	return useSyncExternalStore(
		listener => {
			store.listeners.add(listener)

			return () => {
				store.listeners.delete(listener)
			}
		},
		() => store.request,
		() => store.request
	)
}

// A link or a note asked for from the toolbar's Insert menu: run as the keyboard's Ctrl/Cmd+K and Shift+F2 are, by the
// view's own shortcut handling (useShortcutRequests), so the menu needs nothing of the annotations feature.

export type InsertShortcut = "hyperlink" | "note"

const shortcutListeners = new WeakMap<SheetGridModel, Set<(shortcut: InsertShortcut) => void>>()

export function requestShortcut(model: SheetGridModel, shortcut: InsertShortcut): void {
	for (const listener of shortcutListeners.get(model) ?? []) {
		listener(shortcut)
	}
}

// The view's shortcut handler takes what the Insert menu asks for.
export function useShortcutRequests(model: SheetGridModel, run: (shortcut: InsertShortcut) => boolean): void {
	const latest = useLatestRef(run)

	useEffect(() => {
		let listeners = shortcutListeners.get(model)

		if (listeners === undefined) {
			listeners = new Set()
			shortcutListeners.set(model, listeners)
		}

		const listener = (shortcut: InsertShortcut): void => {
			latest.current(shortcut)
		}

		listeners.add(listener)

		return () => {
			listeners.delete(listener)
		}
	}, [model, latest])
}
