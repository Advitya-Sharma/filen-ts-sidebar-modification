// Where the library index lives in OPFS: one directory per format generation, two slot files per root snapshot and
// two for the catalog the screens paint from.
export const INDEX_DIR_ROOT = ["videos", "index"]
export const INDEX_DIR = [...INDEX_DIR_ROOT, "v1"]

// Held by every snapshot and catalog write (and shared by reads) in every tab, and by the logout wipe: a write either
// lands before the wipe or finds its directory gone.
export const INDEX_LOCK = "filen-web-videos-index"

export function snapshotSlotName(rootUuid: string, slot: 0 | 1): string {
	return `${rootUuid}.${slot === 0 ? "a" : "b"}.fvix`
}

export function catalogSlotName(slot: 0 | 1): string {
	return `library.${slot === 0 ? "a" : "b"}.fvct`
}
