import { wipeOpfsTree } from "@/lib/storage/opfsBlobStore"
import { INDEX_DIR, INDEX_DIR_ROOT, INDEX_LOCK } from "@/features/videos/lib/indexStore.logic"

// Logout: every root's index snapshot. They hold file keys. Under the snapshot lock, so no write interleaves.
export function wipeVideoIndexStore(root?: Promise<FileSystemDirectoryHandle>): Promise<void> {
	return navigator.locks.request(INDEX_LOCK, () => wipeOpfsTree(INDEX_DIR_ROOT, [INDEX_DIR], root))
}
