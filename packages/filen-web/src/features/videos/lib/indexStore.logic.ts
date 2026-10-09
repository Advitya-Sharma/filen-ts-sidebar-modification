// Where the library index snapshots live in OPFS: one directory per format generation, two slot files per root.
export const INDEX_DIR_ROOT = ["videos", "index"]
export const INDEX_DIR = [...INDEX_DIR_ROOT, "v1"]

// Held by every snapshot write (and shared by reads) in every tab, and by the logout wipe: a write either lands
// before the wipe or finds its directory gone.
export const INDEX_LOCK = "filen-web-videos-index"

export function snapshotSlotName(rootUuid: string, slot: 0 | 1): string {
	return `${rootUuid}.${slot === 0 ? "a" : "b"}.fvix`
}
