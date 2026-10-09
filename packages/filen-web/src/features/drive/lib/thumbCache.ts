import { THUMB_DIR, THUMB_DIR_ROOT, THUMB_EXT } from "@/features/drive/lib/thumbnails.logic"
import { deleteStoreEntry, readStoreBlob, wipeOpfsTree } from "@/lib/storage/opfsBlobStore"

// Main-thread read side of the OPFS thumbnail store — async only (no createSyncAccessHandle, which
// is dedicated-worker-only by spec; see workers/thumbStore.ts for the worker-side write path over
// the same tree).

const THUMB_STORE = { dir: THUMB_DIR, ext: THUMB_EXT }

// A cache miss (never written, or evicted) resolves null rather than rejecting — the service's own
// generate-on-miss path treats this as the normal "not cached yet" signal, not an error.
export function readThumbnailBlob(uuid: string): Promise<Blob | null> {
	return readStoreBlob(THUMB_STORE, uuid)
}

// Called from the main thread directly (no worker round trip needed — removeEntry needs no
// exclusive lock) by the service's invalidateThumbnail, after a uuid rotation or a failed render
// that should allow a fresh regenerate. A missing entry is a clean no-op.
export function deleteThumbnail(uuid: string): Promise<void> {
	return deleteStoreEntry(THUMB_STORE, uuid)
}

// Logout: the whole thumbnail tree, every cache generation included — decrypted derivatives of the
// account's files.
export function wipeThumbnailStore(root?: Promise<FileSystemDirectoryHandle>): Promise<void> {
	return wipeOpfsTree(THUMB_DIR_ROOT, [THUMB_DIR], root)
}
