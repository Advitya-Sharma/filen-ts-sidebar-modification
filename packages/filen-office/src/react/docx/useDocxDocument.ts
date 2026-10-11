import { useState, useSyncExternalStore } from "react"
import { DocxDocumentStore, IDLE_DOCUMENT, type DocxDocument, type DocxSource } from "./store"

// The state of a document opened from `source`: loading, then pages streaming in, then ready, or an error
// with its reason (encrypted, legacyBinary, notZip, malformed, failed...). A new source object opens a new
// document; the previous one is closed (its worker terminated, its fonts unregistered).

function subscribeNothing(): () => void {
	return () => undefined
}

function idle(): DocxDocument {
	return IDLE_DOCUMENT
}

export function useDocxDocument(source: DocxSource | null): DocxDocument {
	const [entry, setEntry] = useState(() => ({ source, store: source === null ? null : new DocxDocumentStore(source) }))
	let current = entry

	// A new source replaces the store during render; the old one closes once nothing subscribes to it.
	if (entry.source !== source) {
		current = { source, store: source === null ? null : new DocxDocumentStore(source) }
		setEntry(current)
	}

	const store = current.store

	return useSyncExternalStore(store?.subscribe ?? subscribeNothing, store?.getSnapshot ?? idle, store?.getSnapshot ?? idle)
}
