import { wipeOpfsTree } from "@/lib/storage/opfsBlobStore"
import { ART_DIR, ART_DIR_ROOT } from "@/features/videos/lib/artStore.logic"

// Logout: the whole artwork tree, every generation included — decrypted derivatives of the account's files.
export function wipeArtStore(root?: Promise<FileSystemDirectoryHandle>): Promise<void> {
	return wipeOpfsTree(ART_DIR_ROOT, [ART_DIR], root)
}
