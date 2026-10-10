import { opfsDirectory, isNotFoundError } from "@/lib/storage/opfs"
import { INDEX_DIR, INDEX_LOCK, catalogSlotName, snapshotSlotName } from "@/features/videos/lib/indexStore.logic"
import { SNAPSHOT_HEADER_BYTES, crc32, readSnapshotHeader } from "@/workers/videoIndex/snapshot.logic"
import { CATALOG_HEADER_BYTES, readCatalogHeader } from "@/workers/videoIndex/catalogFile.logic"
import type { CatalogStore, SnapshotStore, SnapshotWatermark } from "@/workers/videoIndex/engine"

// The OPFS side of the index snapshots, their watermarks and the catalog. Writes go through a sync access handle (createWritable is
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

// A root's watermark beside its slots (32 B): "FVWM", flags u32 (1: a watermark is set), generation u64, watermark
// u64, the snapshot's crc32 u32, and a crc32 u32 over the bytes before it.
const WATERMARK_MAGIC = 0x4d575646
const WATERMARK_BYTES = 32

function watermarkName(rootUuid: string): string {
	return `${rootUuid}.fvwm`
}

export function encodeSnapshotWatermark(mark: SnapshotWatermark): Uint8Array<ArrayBuffer> {
	const bytes = new Uint8Array(WATERMARK_BYTES)
	const view = new DataView(bytes.buffer)

	view.setUint32(0, WATERMARK_MAGIC, true)
	view.setUint32(4, mark.watermark === null ? 0 : 1, true)
	view.setBigUint64(8, mark.generation, true)
	view.setBigUint64(16, mark.watermark ?? 0n, true)
	view.setUint32(24, mark.crc, true)
	view.setUint32(28, crc32(bytes.subarray(0, 28)), true)

	return bytes
}

// null for anything but an intact record.
export function decodeSnapshotWatermark(bytes: Uint8Array): SnapshotWatermark | null {
	if (bytes.length !== WATERMARK_BYTES) {
		return null
	}

	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)

	if (view.getUint32(0, true) !== WATERMARK_MAGIC || view.getUint32(28, true) !== crc32(bytes.subarray(0, 28))) {
		return null
	}

	return {
		generation: view.getBigUint64(8, true),
		watermark: (view.getUint32(4, true) & 1) === 1 ? view.getBigUint64(16, true) : null,
		crc: view.getUint32(24, true)
	}
}

function snapshotGeneration(header: Uint8Array): bigint | null {
	return readSnapshotHeader(header)?.generation ?? null
}

function catalogGeneration(header: Uint8Array): bigint | null {
	return readCatalogHeader(header)?.generation ?? null
}

export const opfsSnapshotStore: SnapshotStore = {
	readSlot(rootUuid, slot) {
		return shared(() => readSlot(snapshotSlotName(rootUuid, slot)))
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
	async readWatermark(rootUuid) {
		const bytes = await shared(() => readSlot(watermarkName(rootUuid)))

		return bytes === null ? null : decodeSnapshotWatermark(bytes)
	},
	writeWatermark(rootUuid, mark, cancelled) {
		return writeSlot(watermarkName(rootUuid), encodeSnapshotWatermark(mark), cancelled)
	},
	remove(rootUuid) {
		return navigator.locks.request(INDEX_LOCK, async () => {
			const dir = await indexDirectory()

			for (const name of [snapshotSlotName(rootUuid, 0), snapshotSlotName(rootUuid, 1), watermarkName(rootUuid)]) {
				try {
					await dir.removeEntry(name)
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
