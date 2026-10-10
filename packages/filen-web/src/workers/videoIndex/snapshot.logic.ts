import { CLASSIFIER_VERSION, VIDEO_FILE_KINDS_SIGNATURE } from "@filen/shared/videos"
import type { FileEncryptionVersion } from "@filen/sdk-rs"
import type { IndexDir, IndexFile, RootState } from "@/workers/videoIndex/rootIndex.logic"

// The binary snapshot of one root's index, written to two alternating slots so a torn or corrupt write always
// leaves the previous one readable. Little-endian throughout.
//
// Header (96 B): "FVIX", formatVersion u16, classifierVersion u16, flags u32, crc32 u32 (over everything after
// it), generation u64, driveWatermark u64, completeAt f64, prevCompleteAt f64, dirCount u32, fileCount u32,
// stringBytes u32, walkedEntries u32, rootUuid 16 B, root name offset u32 + length u32, internCount u32, kinds
// signature u32 (what the index keeps: a snapshot kept under other rules is stale). Then the intern table (offset
// u32 + length u32 each: mime types, regions, buckets), directory records, file records, and one UTF-8 string blob
// every name and fallback field points into.
//
// A uuid is 16 raw bytes when canonical (lowercase 8-4-4-4-12), else a string reference with a flag. A key is 32
// raw bytes when it is 64 lowercase hex characters (v3) or 32 ASCII characters (v1/v2), else a string reference.

export const SNAPSHOT_MAGIC = 0x58495646 // "FVIX" read as a little-endian u32
export const SNAPSHOT_FORMAT_VERSION = 2
export const SNAPSHOT_HEADER_BYTES = 96
export const SNAPSHOT_DIR_BYTES = 29
export const SNAPSHOT_FILE_BYTES = 151

const HEADER_FLAG_WATERMARK = 1

const DIR_FLAG_UUID_STRING = 1

const FILE_FLAG_THUMBNAIL = 1
const FILE_FLAG_FAVORITED = 2
const FILE_FLAG_UUID_STRING = 4
const FILE_FLAG_STABLE_STRING = 8
const FILE_FLAG_KEY_STRING = 16
const FILE_FLAG_KEY_HEX = 32

const NO_PARENT = -1

// ── CRC32 (IEEE) ─────────────────────────────────────────────────────────────

// Slicing-by-8: eight bytes per step through eight tables (8 KiB), several times faster than a byte at a time over a
// snapshot of many megabytes, which is checked on every load and written on every save.
const CRC_TABLES = (() => {
	const tables = new Int32Array(256 * 8)

	for (let n = 0; n < 256; n++) {
		let c = n

		for (let k = 0; k < 8; k++) {
			c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
		}

		tables[n] = c
	}

	for (let n = 0; n < 256; n++) {
		let c = tables[n] ?? 0

		for (let t = 1; t < 8; t++) {
			c = (tables[c & 0xff] ?? 0) ^ (c >>> 8)
			tables[t * 256 + n] = c
		}
	}

	return tables
})()

// The running CRC (pre-inversion) over `bytes`, continuing from `crc` (-1 to start).
function crc32Update(crc: number, bytes: Uint8Array): number {
	const t = CRC_TABLES
	const tail = bytes.length - (bytes.length % 8)
	let out = crc
	let i = 0

	for (; i < tail; i += 8) {
		const low = out ^ ((bytes[i] ?? 0) | ((bytes[i + 1] ?? 0) << 8) | ((bytes[i + 2] ?? 0) << 16) | ((bytes[i + 3] ?? 0) << 24))

		out =
			(t[1792 + (low & 0xff)] ?? 0) ^
			(t[1536 + ((low >>> 8) & 0xff)] ?? 0) ^
			(t[1280 + ((low >>> 16) & 0xff)] ?? 0) ^
			(t[1024 + (low >>> 24)] ?? 0) ^
			(t[768 + (bytes[i + 4] ?? 0)] ?? 0) ^
			(t[512 + (bytes[i + 5] ?? 0)] ?? 0) ^
			(t[256 + (bytes[i + 6] ?? 0)] ?? 0) ^
			(t[bytes[i + 7] ?? 0] ?? 0)
	}

	for (; i < bytes.length; i++) {
		out = (t[(out ^ (bytes[i] ?? 0)) & 0xff] ?? 0) ^ (out >>> 8)
	}

	return out
}

export function crc32(bytes: Uint8Array): number {
	return ~crc32Update(-1, bytes) >>> 0
}

// crc32 in steps of CRC_STEP_BYTES, so a large snapshot's check yields between them.
const CRC_STEP_BYTES = 1 << 20

function* crc32Steps(bytes: Uint8Array): Generator<void, number> {
	let crc = -1

	for (let at = 0; at < bytes.length; at += CRC_STEP_BYTES) {
		crc = crc32Update(crc, bytes.subarray(at, at + CRC_STEP_BYTES))

		yield
	}

	return ~crc >>> 0
}

// What decides which files the index keeps, so a snapshot written under other rules is known stale.
export const INDEX_KINDS_SIGNATURE = crc32(new TextEncoder().encode(VIDEO_FILE_KINDS_SIGNATURE))

// Whether a decoded snapshot was classified and filtered under the rules this build runs: when not, its records
// still render, but the root is walked again.
export function snapshotIsCurrent(snapshot: DecodedSnapshot): boolean {
	return snapshot.classifierVersion === CLASSIFIER_VERSION && snapshot.kindsSignature === INDEX_KINDS_SIGNATURE
}

// ── uuids, keys ──────────────────────────────────────────────────────────────

const CANONICAL_UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/
const HEX_KEY = /^[0-9a-f]{64}$/
const ASCII_KEY = /^[\x20-\x7e]{32}$/

const HEX = Array.from({ length: 256 }, (_, byte) => byte.toString(16).padStart(2, "0"))

function hexValue(code: number): number {
	return code <= 57 ? code - 48 : code - 87
}

function writeHex(bytes: Uint8Array, at: number, hex: string): void {
	let out = at

	for (let i = 0; i < hex.length; i++) {
		const code = hex.charCodeAt(i)

		if (code === 45) {
			continue
		}

		bytes[out++] = (hexValue(code) << 4) | hexValue(hex.charCodeAt(i + 1))
		i++
	}
}

function readUuid(bytes: Uint8Array, at: number): string {
	let out = ""

	for (let i = 0; i < 16; i++) {
		out += HEX[bytes[at + i] ?? 0] ?? "00"

		if (i === 3 || i === 5 || i === 7 || i === 9) {
			out += "-"
		}
	}

	return out
}

function readHexKey(bytes: Uint8Array, at: number): string {
	let out = ""

	for (let i = 0; i < 32; i++) {
		out += HEX[bytes[at + i] ?? 0] ?? "00"
	}

	return out
}

// ── Strings ──────────────────────────────────────────────────────────────────

class StringBlob {
	public bytes = new Uint8Array(4096)
	public length = 0
	private readonly encoder = new TextEncoder()

	// Offset and length of `text` once appended.
	public add(text: string): [number, number] {
		const needed = this.length + text.length * 3

		if (needed > this.bytes.length) {
			const grown = new Uint8Array(Math.max(needed, this.bytes.length * 2))

			grown.set(this.bytes.subarray(0, this.length))
			this.bytes = grown
		}

		const offset = this.length
		const { written } = this.encoder.encodeInto(text, this.bytes.subarray(offset))

		this.length += written

		return [offset, written]
	}
}

// ── Encode ───────────────────────────────────────────────────────────────────

export function encodeSnapshot(state: RootState): Uint8Array<ArrayBuffer> {
	const steps = encodeSnapshotSteps(state)

	for (;;) {
		const step = steps.next()

		if (step.done === true) {
			return step.value
		}
	}
}

// encodeSnapshot in steps (yields between runs of records). The root's records and header are taken when it starts,
// so a patch landing meanwhile goes to the next write.
export function* encodeSnapshotSteps(state: RootState): Generator<void, Uint8Array<ArrayBuffer>> {
	const root = {
		uuid: state.uuid,
		name: state.name,
		watermark: state.watermark,
		generation: state.generation,
		completeAt: state.completeAt,
		prevCompleteAt: state.prevCompleteAt,
		walkedEntries: state.walkedEntries
	}
	const blob = new StringBlob()
	const interned: string[] = []
	const internIndex = new Map<string, number>()
	const intern = (text: string): number => {
		let index = internIndex.get(text)

		if (index === undefined) {
			index = interned.length
			interned.push(text)
			internIndex.set(text, index)
		}

		return index
	}
	const dirs = [...state.dirs.values()]
	const dirIndex = new Map<string, number>()
	const files = [...state.files.values()]

	for (let i = 0; i < dirs.length; i++) {
		dirIndex.set(dirs[i]?.uuid ?? "", i)

		if (i % (DECODE_STEP_RECORDS * 4) === DECODE_STEP_RECORDS * 4 - 1) {
			yield
		}
	}

	for (let i = 0; i < files.length; i++) {
		const file = files[i]

		if (file !== undefined) {
			intern(file.mime)
			intern(file.region)
			intern(file.bucket)
		}

		if (i % (DECODE_STEP_RECORDS * 4) === DECODE_STEP_RECORDS * 4 - 1) {
			yield
		}
	}

	const internBytes = interned.length * 8
	const dirsAt = SNAPSHOT_HEADER_BYTES + internBytes
	const filesAt = dirsAt + dirs.length * SNAPSHOT_DIR_BYTES
	const stringsAt = filesAt + files.length * SNAPSHOT_FILE_BYTES
	const head = new Uint8Array(stringsAt)
	const view = new DataView(head.buffer)
	const parentIndex = (parent: string): number => (parent === root.uuid ? NO_PARENT : (dirIndex.get(parent) ?? NO_PARENT))
	const writeString = (at: number, text: string): void => {
		const [offset, length] = blob.add(text)

		view.setUint32(at, offset, true)
		view.setUint32(at + 4, length, true)
	}
	const writeUuidField = (at: number, uuid: string): boolean => {
		if (CANONICAL_UUID.test(uuid)) {
			writeHex(head, at, uuid)

			return false
		}

		writeString(at, uuid)

		return true
	}

	for (let i = 0; i < interned.length; i++) {
		writeString(SNAPSHOT_HEADER_BYTES + i * 8, interned[i] ?? "")
	}

	for (let i = 0; i < dirs.length; i++) {
		const dir = dirs[i]

		if (dir === undefined) {
			continue
		}

		if (i % DECODE_STEP_RECORDS === DECODE_STEP_RECORDS - 1) {
			yield
		}

		const at = dirsAt + i * SNAPSHOT_DIR_BYTES
		const stringUuid = writeUuidField(at, dir.uuid)

		view.setInt32(at + 16, parentIndex(dir.parent), true)
		writeString(at + 20, dir.name)
		view.setUint8(at + 28, stringUuid ? DIR_FLAG_UUID_STRING : 0)
	}

	for (let i = 0; i < files.length; i++) {
		const file = files[i]

		if (file === undefined) {
			continue
		}

		if (i % DECODE_STEP_RECORDS === DECODE_STEP_RECORDS - 1) {
			yield
		}

		const at = filesAt + i * SNAPSHOT_FILE_BYTES
		let flags = (file.canMakeThumbnail ? FILE_FLAG_THUMBNAIL : 0) | (file.favorited ? FILE_FLAG_FAVORITED : 0)

		if (writeUuidField(at, file.uuid)) {
			flags |= FILE_FLAG_UUID_STRING
		}

		if (writeUuidField(at + 16, file.stableUuid)) {
			flags |= FILE_FLAG_STABLE_STRING
		}

		view.setInt32(at + 32, parentIndex(file.parent), true)
		writeString(at + 36, file.name)
		view.setFloat64(at + 44, file.size, true)
		view.setFloat64(at + 52, file.metaSize, true)
		view.setUint32(at + 60, file.chunks, true)
		view.setFloat64(at + 64, file.timestamp, true)
		view.setFloat64(at + 72, file.created, true)
		view.setFloat64(at + 80, file.modified, true)
		view.setFloat64(at + 88, file.dateAdded, true)
		view.setFloat64(at + 96, file.tombstoneAt, true)

		if (HEX_KEY.test(file.key)) {
			writeHex(head, at + 104, file.key)
			flags |= FILE_FLAG_KEY_HEX
		} else if (ASCII_KEY.test(file.key)) {
			for (let k = 0; k < 32; k++) {
				head[at + 104 + k] = file.key.charCodeAt(k)
			}
		} else {
			writeString(at + 104, file.key)
			flags |= FILE_FLAG_KEY_STRING
		}

		view.setUint8(at + 136, file.keyVersion)
		view.setUint8(at + 137, Math.min(255, file.tombstoneWalks))
		view.setUint32(at + 138, internIndex.get(file.mime) ?? 0, true)
		view.setUint32(at + 142, internIndex.get(file.region) ?? 0, true)
		view.setUint32(at + 146, internIndex.get(file.bucket) ?? 0, true)
		view.setUint8(at + 150, flags)
	}

	const [rootNameOffset, rootNameLength] = blob.add(root.name)
	const out = new Uint8Array(stringsAt + blob.length)

	out.set(head)
	out.set(blob.bytes.subarray(0, blob.length), stringsAt)

	const header = new DataView(out.buffer)

	header.setUint32(0, SNAPSHOT_MAGIC, true)
	header.setUint16(4, SNAPSHOT_FORMAT_VERSION, true)
	header.setUint16(6, CLASSIFIER_VERSION, true)
	header.setUint32(8, root.watermark === null ? 0 : HEADER_FLAG_WATERMARK, true)
	header.setBigUint64(16, root.generation, true)
	header.setBigUint64(24, root.watermark ?? 0n, true)
	header.setFloat64(32, root.completeAt, true)
	header.setFloat64(40, root.prevCompleteAt, true)
	header.setUint32(48, dirs.length, true)
	header.setUint32(52, files.length, true)
	header.setUint32(56, blob.length, true)
	header.setUint32(60, root.walkedEntries, true)
	writeHex(out, 64, CANONICAL_UUID.test(root.uuid) ? root.uuid : "00000000-0000-0000-0000-000000000000")
	header.setUint32(80, rootNameOffset, true)
	header.setUint32(84, rootNameLength, true)
	header.setUint32(88, interned.length, true)
	header.setUint32(92, INDEX_KINDS_SIGNATURE, true)
	header.setUint32(12, yield* crc32Steps(out.subarray(16)), true)

	return out
}

// ── Decode ───────────────────────────────────────────────────────────────────

export interface SnapshotHeader {
	generation: bigint
	rootUuid: string
}

// The header fields slot selection needs, from the first SNAPSHOT_HEADER_BYTES of a slot; null when they aren't
// a snapshot header. Checks nothing past the header.
export function readSnapshotHeader(bytes: Uint8Array): SnapshotHeader | null {
	if (bytes.length < SNAPSHOT_HEADER_BYTES) {
		return null
	}

	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)

	if (view.getUint32(0, true) !== SNAPSHOT_MAGIC || view.getUint16(4, true) !== SNAPSHOT_FORMAT_VERSION) {
		return null
	}

	return {
		generation: view.getBigUint64(16, true),
		rootUuid: readUuid(bytes, 64)
	}
}

export interface DecodedSnapshot {
	header: Omit<RootState, "dirs" | "files" | "byUuid">
	classifierVersion: number
	kindsSignature: number
	dirs: IndexDir[]
	files: IndexFile[]
}

// Records decoded between yields of decodeSnapshotSteps.
const DECODE_STEP_RECORDS = 1024

// null for anything but an intact snapshot of `rootUuid`: wrong magic or version, a size that doesn't add up, or
// a CRC mismatch.
export function decodeSnapshot(bytes: Uint8Array, rootUuid: string): DecodedSnapshot | null {
	const steps = decodeSnapshotSteps(bytes, rootUuid)

	for (;;) {
		const step = steps.next()

		if (step.done === true) {
			return step.value
		}
	}
}

// decodeSnapshot in steps (yields between them), so a large snapshot never holds the worker for long.
export function* decodeSnapshotSteps(bytes: Uint8Array, rootUuid: string): Generator<void, DecodedSnapshot | null> {
	const header = readSnapshotHeader(bytes)

	if (header === null || (CANONICAL_UUID.test(rootUuid) && header.rootUuid !== rootUuid)) {
		return null
	}

	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
	const dirCount = view.getUint32(48, true)
	const fileCount = view.getUint32(52, true)
	const stringBytes = view.getUint32(56, true)
	const internCount = view.getUint32(88, true)
	const dirsAt = SNAPSHOT_HEADER_BYTES + internCount * 8
	const filesAt = dirsAt + dirCount * SNAPSHOT_DIR_BYTES
	const stringsAt = filesAt + fileCount * SNAPSHOT_FILE_BYTES

	if (stringsAt + stringBytes !== bytes.length || view.getUint32(12, true) !== (yield* crc32Steps(bytes.subarray(16)))) {
		return null
	}

	const decoder = new TextDecoder()
	const readString = (at: number): string => {
		const offset = view.getUint32(at, true)
		const length = view.getUint32(at + 4, true)

		return decoder.decode(bytes.subarray(stringsAt + offset, stringsAt + offset + length))
	}
	const interned: string[] = []

	for (let i = 0; i < internCount; i++) {
		interned.push(readString(SNAPSHOT_HEADER_BYTES + i * 8))
	}

	const dirs: IndexDir[] = []

	for (let i = 0; i < dirCount; i++) {
		if (i % DECODE_STEP_RECORDS === DECODE_STEP_RECORDS - 1) {
			yield
		}

		const at = dirsAt + i * SNAPSHOT_DIR_BYTES
		const flags = view.getUint8(at + 28)

		dirs.push({
			uuid: flags & DIR_FLAG_UUID_STRING ? readString(at) : readUuid(bytes, at),
			parent: "",
			name: readString(at + 20)
		})
	}

	const parentOf = (index: number): string => (index === NO_PARENT ? rootUuid : (dirs[index]?.uuid ?? rootUuid))

	for (let i = 0; i < dirCount; i++) {
		const dir = dirs[i]

		if (dir !== undefined) {
			dir.parent = parentOf(view.getInt32(dirsAt + i * SNAPSHOT_DIR_BYTES + 16, true))
		}
	}

	const files: IndexFile[] = []

	for (let i = 0; i < fileCount; i++) {
		if (i % DECODE_STEP_RECORDS === DECODE_STEP_RECORDS - 1) {
			yield
		}

		const at = filesAt + i * SNAPSHOT_FILE_BYTES
		const flags = view.getUint8(at + 150)
		const keyVersion = view.getUint8(at + 136)

		files.push({
			uuid: flags & FILE_FLAG_UUID_STRING ? readString(at) : readUuid(bytes, at),
			stableUuid: flags & FILE_FLAG_STABLE_STRING ? readString(at + 16) : readUuid(bytes, at + 16),
			parent: parentOf(view.getInt32(at + 32, true)),
			name: readString(at + 36),
			size: view.getFloat64(at + 44, true),
			metaSize: view.getFloat64(at + 52, true),
			chunks: view.getUint32(at + 60, true),
			timestamp: view.getFloat64(at + 64, true),
			created: view.getFloat64(at + 72, true),
			modified: view.getFloat64(at + 80, true),
			dateAdded: view.getFloat64(at + 88, true),
			tombstoneAt: view.getFloat64(at + 96, true),
			tombstoneWalks: view.getUint8(at + 137),
			key:
				flags & FILE_FLAG_KEY_STRING
					? readString(at + 104)
					: flags & FILE_FLAG_KEY_HEX
						? readHexKey(bytes, at + 104)
						: String.fromCharCode(...bytes.subarray(at + 104, at + 136)),
			keyVersion: keyVersion === 1 || keyVersion === 2 ? keyVersion : (3 satisfies FileEncryptionVersion),
			mime: interned[view.getUint32(at + 138, true)] ?? "",
			region: interned[view.getUint32(at + 142, true)] ?? "",
			bucket: interned[view.getUint32(at + 146, true)] ?? "",
			canMakeThumbnail: (flags & FILE_FLAG_THUMBNAIL) !== 0,
			favorited: (flags & FILE_FLAG_FAVORITED) !== 0
		})
	}

	return {
		classifierVersion: view.getUint16(6, true),
		kindsSignature: view.getUint32(92, true),
		header: {
			uuid: rootUuid,
			name: decoder.decode(
				bytes.subarray(stringsAt + view.getUint32(80, true), stringsAt + view.getUint32(80, true) + view.getUint32(84, true))
			),
			watermark: view.getUint32(8, true) & HEADER_FLAG_WATERMARK ? view.getBigUint64(24, true) : null,
			generation: header.generation,
			completeAt: view.getFloat64(32, true),
			prevCompleteAt: view.getFloat64(40, true),
			walkedEntries: view.getUint32(60, true)
		},
		dirs,
		files
	}
}

// The slot to load: the intact one with the highest generation, else the other. null when neither is intact.
export function pickSnapshot(slots: readonly (Uint8Array | null)[], rootUuid: string): DecodedSnapshot | null {
	const steps = pickSnapshotSteps(slots, rootUuid)

	for (;;) {
		const step = steps.next()

		if (step.done === true) {
			return step.value
		}
	}
}

export function* pickSnapshotSteps(slots: readonly (Uint8Array | null)[], rootUuid: string): Generator<void, DecodedSnapshot | null> {
	const ordered = slots
		.filter((slot): slot is Uint8Array => slot !== null)
		.map(slot => ({
			slot,
			generation: readSnapshotHeader(slot)?.generation ?? -1n
		}))
		.sort((a, b) => (a.generation === b.generation ? 0 : a.generation > b.generation ? -1 : 1))

	for (const { slot } of ordered) {
		const decoded = yield* decodeSnapshotSteps(slot, rootUuid)

		if (decoded !== null) {
			return decoded
		}
	}

	return null
}

// Which of two slots to overwrite, given each one's header generation (null when missing or unreadable): the
// older one, so the newest intact copy survives a torn write. The new generation is one past both.
export function nextSlot(generations: readonly [bigint | null, bigint | null]): { slot: 0 | 1; generation: bigint } {
	const [a, b] = generations
	const highest = [a, b].reduce<bigint>((max, value) => (value !== null && value > max ? value : max), 0n)

	if (a === null) {
		return {
			slot: 0,
			generation: highest + 1n
		}
	}

	if (b === null) {
		return {
			slot: 1,
			generation: highest + 1n
		}
	}

	return {
		slot: a <= b ? 0 : 1,
		generation: highest + 1n
	}
}
