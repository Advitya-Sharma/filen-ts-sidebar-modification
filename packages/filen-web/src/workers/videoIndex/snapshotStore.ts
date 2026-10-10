import { opfsDirectory, isNotFoundError } from "@/lib/storage/opfs"
import { INDEX_DIR, INDEX_LOCK, catalogSlotName, snapshotSlotName } from "@/features/videos/lib/indexStore.logic"
import { SNAPSHOT_HEADER_BYTES, readSnapshotHeader } from "@/workers/videoIndex/snapshot.logic"
import { CATALOG_HEADER_BYTES, readCatalogHeader } from "@/workers/videoIndex/catalogFile.logic"
import type { CatalogStore, SnapshotStore } from "@/workers/videoIndex/engine"

// The OPFS side of the index snapshots and the catalog. Writes go through a sync access handle (createWritable is
// missing in some engines), over the slot in place: a torn write only spoils that slot, which its checksum then
// rejects in favour of the other. Every write holds the index lock across tabs, reads share it. The directory is
// created once per session; when it is gone after that (a logout wipe in any tab), writes fail instead of
// re-creating it.

let directory: Promise<FileSystemDirectoryHandle> | null = null

function indexDirectory(): Promise<FileSystemDirectoryHandle> {
	directory ??= opfsDirectory(INDEX_DIR)

	return directory
}

async function slotFile(name: string): Promise<globalThis.File | null> {
	try {
		return await (await (await indexDirectory()).getFileHandle(name)).getFile()
	} catch (e) {
		if (isNotFoundError(e)) {
			return null
		}

		throw e
	}
}

async function readSlot(name: string): Promise<Uint8Array | null> {
	const file = await slotFile(name)

	return file === null ? null : new Uint8Array(await file.arrayBuffer())
}

async function readGeneration(name: string, headerBytes: number, read: (header: Uint8Array) => bigint | null): Promise<bigint | null> {
	const file = await slotFile(name)

	return file === null ? null : read(new Uint8Array(await file.slice(0, headerBytes).arrayBuffer()))
}

function shared<T>(task: () => Promise<T>): Promise<T> {
	return navigator.locks.request(INDEX_LOCK, { mode: "shared" }, task)
}

function writeSlot(name: string, bytes: Uint8Array<ArrayBuffer>, cancelled: () => boolean): Promise<void> {
	return navigator.locks.request(INDEX_LOCK, async () => {
		if (cancelled()) {
			return
		}

		const handle = await (await (await indexDirectory()).getFileHandle(name, { create: true })).createSyncAccessHandle()

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
}

function snapshotGeneration(header: Uint8Array): bigint | null {
	return readSnapshotHeader(header)?.generation ?? null
}

function catalogGeneration(header: Uint8Array): bigint | null {
	return readCatalogHeader(header)?.generation ?? null
}

export const opfsSnapshotStore: SnapshotStore = {
	read(rootUuid) {
		return shared(() => Promise.all([readSlot(snapshotSlotName(rootUuid, 0)), readSlot(snapshotSlotName(rootUuid, 1))]))
	},
	readGenerations(rootUuid) {
		return shared(() =>
			Promise.all([
				readGeneration(snapshotSlotName(rootUuid, 0), SNAPSHOT_HEADER_BYTES, snapshotGeneration),
				readGeneration(snapshotSlotName(rootUuid, 1), SNAPSHOT_HEADER_BYTES, snapshotGeneration)
			])
		)
	},
	write(rootUuid, slot, bytes, cancelled) {
		return writeSlot(snapshotSlotName(rootUuid, slot), bytes, cancelled)
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

export const opfsCatalogStore: CatalogStore = {
	readSlot(slot) {
		return shared(() => readSlot(catalogSlotName(slot)))
	},
	readGenerations() {
		return shared(() =>
			Promise.all([
				readGeneration(catalogSlotName(0), CATALOG_HEADER_BYTES, catalogGeneration),
				readGeneration(catalogSlotName(1), CATALOG_HEADER_BYTES, catalogGeneration)
			])
		)
	},
	write(slot, bytes, cancelled) {
		return writeSlot(catalogSlotName(slot), bytes, cancelled)
	}
}
