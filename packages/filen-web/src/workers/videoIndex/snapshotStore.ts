import { opfsDirectory, isNotFoundError } from "@/lib/storage/opfs"
import { INDEX_DIR, INDEX_LOCK, snapshotSlotName } from "@/features/videos/lib/indexStore.logic"
import { SNAPSHOT_HEADER_BYTES, readSnapshotHeader } from "@/workers/videoIndex/snapshot.logic"
import type { SnapshotStore } from "@/workers/videoIndex/engine"

// The OPFS side of the index snapshots. Writes go through a sync access handle (createWritable is missing in some
// engines), over the slot in place: a torn write only spoils that slot, which its CRC then rejects in favour of the
// other. Every write holds the snapshot lock across tabs, reads share it. The directory is created once per
// session; when it is gone after that (a logout wipe in any tab), writes fail instead of re-creating it.

let directory: Promise<FileSystemDirectoryHandle> | null = null

function indexDirectory(): Promise<FileSystemDirectoryHandle> {
	directory ??= opfsDirectory(INDEX_DIR)

	return directory
}

async function slotFile(rootUuid: string, slot: 0 | 1): Promise<globalThis.File | null> {
	try {
		return await (await (await indexDirectory()).getFileHandle(snapshotSlotName(rootUuid, slot))).getFile()
	} catch (e) {
		if (isNotFoundError(e)) {
			return null
		}

		throw e
	}
}

async function readSlot(rootUuid: string, slot: 0 | 1): Promise<Uint8Array | null> {
	const file = await slotFile(rootUuid, slot)

	return file === null ? null : new Uint8Array(await file.arrayBuffer())
}

async function readGeneration(rootUuid: string, slot: 0 | 1): Promise<bigint | null> {
	const file = await slotFile(rootUuid, slot)

	if (file === null) {
		return null
	}

	return readSnapshotHeader(new Uint8Array(await file.slice(0, SNAPSHOT_HEADER_BYTES).arrayBuffer()))?.generation ?? null
}

function shared<T>(task: () => Promise<T>): Promise<T> {
	return navigator.locks.request(INDEX_LOCK, { mode: "shared" }, task)
}

export const opfsSnapshotStore: SnapshotStore = {
	read(rootUuid) {
		return shared(() => Promise.all([readSlot(rootUuid, 0), readSlot(rootUuid, 1)]))
	},
	readGenerations(rootUuid) {
		return shared(() => Promise.all([readGeneration(rootUuid, 0), readGeneration(rootUuid, 1)]))
	},
	write(rootUuid, slot, bytes, cancelled) {
		return navigator.locks.request(INDEX_LOCK, async () => {
			if (cancelled()) {
				return
			}

			const handle = await (
				await (await indexDirectory()).getFileHandle(snapshotSlotName(rootUuid, slot), { create: true })
			).createSyncAccessHandle()

			try {
				if (cancelled()) {
					return
				}

				handle.write(bytes, { at: 0 })
				handle.truncate(bytes.length)
				handle.flush()
			} finally {
				handle.close()
			}
		})
	},
	remove(rootUuid) {
		return navigator.locks.request(INDEX_LOCK, async () => {
			const dir = await indexDirectory()

			for (const slot of [0, 1] as const) {
				try {
					await dir.removeEntry(snapshotSlotName(rootUuid, slot))
				} catch (e) {
					if (!isNotFoundError(e)) {
						throw e
					}
				}
			}
		})
	}
}
