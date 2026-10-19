import type { ByteSource } from "../opc/io"
import { CfbError } from "./errors"

// [MS-CFB] Compound File Binary format: header §2.2, sector numbers §2.1, FAT §2.3, mini FAT §2.4, DIFAT §2.5,
// directory §2.6.

export const CFB_SIGNATURE: readonly number[] = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]
export const MAXREGSECT = 0xfffffffa
export const DIFSECT = 0xfffffffc
export const FATSECT = 0xfffffffd
export const ENDOFCHAIN = 0xfffffffe
export const FREESECT = 0xffffffff
export const NOSTREAM = 0xffffffff
export const MINI_STREAM_CUTOFF = 4096
export const MINI_SECTOR_SIZE = 64
export const HEADER_DIFAT = 109
export const ENTRY_SIZE = 128
export const TYPE_STORAGE = 1
export const TYPE_STREAM = 2
export const TYPE_ROOT = 5

export interface CfbLimits {
	// Largest file accepted.
	readonly maxSize: number
	// Most directory entries (allocated or not) the directory chain may hold.
	readonly maxEntries: number
}

export const DEFAULT_CFB_LIMITS: CfbLimits = {
	maxSize: 2 ** 31,
	maxEntries: 65_536
}

export interface CfbOpenOptions {
	// Merged over DEFAULT_CFB_LIMITS.
	readonly limits?: Partial<CfbLimits>
}

export type CfbEntryKind = "root" | "storage" | "stream"

export interface CfbEntry {
	readonly name: string
	readonly kind: CfbEntryKind
	// Names from the root's child down to this entry (empty for the root).
	readonly path: readonly string[]
	// Stream length; the mini stream's for the root; 0 for storages.
	readonly size: number
	// "{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}", or undefined when all zero.
	readonly clsid: string | undefined
	// In directory order (§2.6.4: length, then upper-cased code units).
	readonly children: readonly CfbEntry[]
}

export function isCompoundFile(head: Uint8Array): boolean {
	return head.length >= 8 && CFB_SIGNATURE.every((byte, k) => head[k] === byte)
}

function upperUnit(code: number): number {
	if (code < 0x61) return code
	if (code <= 0x7a) return code - 0x20
	if (code < 0x80) return code

	const upper = String.fromCharCode(code).toUpperCase()

	return upper.length === 1 ? upper.charCodeAt(0) : code
}

// A name upper-cased code unit by code unit (simple case mapping), as §2.6.4 compares them.
export function upperName(name: string): string {
	let out = ""

	for (let k = 0; k < name.length; k++) out += String.fromCharCode(upperUnit(name.charCodeAt(k)))

	return out
}

// §2.6.4 order: shorter names first, then by upper-cased code units.
export function compareNames(a: string, b: string): number {
	if (a.length !== b.length) return a.length - b.length

	for (let k = 0; k < a.length; k++) {
		const difference = upperUnit(a.charCodeAt(k)) - upperUnit(b.charCodeAt(k))

		if (difference !== 0) return difference
	}

	return 0
}

function hex(value: number, digits: number): string {
	return value.toString(16).toUpperCase().padStart(digits, "0")
}

// A GUID as stored (§2.1.3 of [MS-DTYP]: Data1-3 little-endian, Data4 as bytes) in registry form.
export function formatClsid(bytes: Uint8Array, at: number): string | undefined {
	let zero = true

	for (let k = 0; k < 16; k++) if (bytes[at + k] !== 0) zero = false

	if (zero) return undefined

	const data = new DataView(bytes.buffer, bytes.byteOffset + at, 16)
	let tail = ""

	for (let k = 8; k < 16; k++) tail += (k === 10 ? "-" : "") + hex(data.getUint8(k), 2)

	return `{${hex(data.getUint32(0, true), 8)}-${hex(data.getUint16(4, true), 4)}-${hex(data.getUint16(6, true), 4)}-${tail}}`
}

class Node implements CfbEntry {
	readonly name: string
	readonly kind: CfbEntryKind
	readonly size: number
	readonly clsid: string | undefined
	readonly children: Node[] = []
	readonly parent: Node | undefined
	readonly start: number
	readonly id: number
	// Upper-cased child name -> child, built on the first lookup.
	byName: Map<string, Node> | undefined

	constructor(
		id: number,
		name: string,
		kind: CfbEntryKind,
		size: number,
		clsid: string | undefined,
		start: number,
		parent: Node | undefined
	) {
		this.id = id
		this.name = name
		this.kind = kind
		this.size = size
		this.clsid = clsid
		this.start = start
		this.parent = parent
	}

	// Computed, not stored: a deep hostile nesting would otherwise cost the square of its depth.
	get path(): readonly string[] {
		if (this.parent === undefined) return []

		const names: string[] = [this.name]

		for (let node = this.parent; node.parent !== undefined; node = node.parent) names.push(node.name)

		return names.reverse()
	}
}

// Where a stream's bytes lie in the file: runs of consecutive bytes, in stream order.
interface Extents {
	// Stream offset each run starts at, and one past the last run's end.
	readonly starts: Float64Array
	readonly offsets: Float64Array
}

// Sector owner tags: a sector belongs to at most one chain, so a chain running into another (or itself) is
// caught on the spot and no stream can make the reader go over the same bytes twice.
const OWNER_FAT = 1
const OWNER_DIRECTORY = 2
const OWNER_MINI_FAT = 3
const OWNER_MINI_STREAM = 4
const OWNER_STREAM = 5

interface Mini {
	readonly fat: Uint32Array
	readonly stream: Extents
	readonly count: number
	readonly size: number
	readonly owners: Uint32Array
}

function corrupt(message: string): CfbError {
	return new CfbError("corrupt", message)
}

function view(bytes: Uint8Array): DataView {
	return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
}

// A read-only compound file over a random-access source. Opening reads the header, the FAT (only the FAT
// sectors that describe sectors of the file) and the directory; streams are read on demand. Every count in
// the file is bounded by the file size or the limits, every chain is checked for loops and for running into
// another chain, and every malformation is a CfbError.
export class CfbFile {
	readonly version: 3 | 4
	readonly sectorSize: number
	readonly root: CfbEntry
	private readonly source: ByteSource
	private readonly sectorCount: number
	private readonly fat: Uint32Array
	private readonly owners: Uint32Array
	private readonly miniFatStart: number
	private readonly extentsCache = new Map<number, Promise<Extents>>()
	private mini: Promise<Mini> | undefined

	private constructor(
		source: ByteSource,
		version: 3 | 4,
		sectorSize: number,
		sectorCount: number,
		fat: Uint32Array,
		owners: Uint32Array,
		miniFatStart: number,
		root: Node
	) {
		this.source = source
		this.version = version
		this.sectorSize = sectorSize
		this.sectorCount = sectorCount
		this.fat = fat
		this.owners = owners
		this.miniFatStart = miniFatStart
		this.root = root
	}

	static async open(source: ByteSource, options: CfbOpenOptions = {}): Promise<CfbFile> {
		const limits: CfbLimits = { ...DEFAULT_CFB_LIMITS, ...options.limits }

		if (source.size < 8 || !isCompoundFile(await source.read(0, 8))) throw new CfbError("notCompoundFile", "Not a compound file")
		if (source.size < 512) throw new CfbError("truncated", "The compound file header is cut short")
		if (source.size > limits.maxSize) throw new CfbError("tooLarge", "The compound file is larger than the limit")

		const header = view(await source.read(0, 512))
		const major = header.getUint16(0x1a, true)

		if (major !== 3 && major !== 4) throw new CfbError("unsupported", `Compound file version ${major}`)
		if (header.getUint16(0x1c, true) !== 0xfffe) throw corrupt("The compound file byte order mark is wrong")

		const shift = header.getUint16(0x1e, true)

		if (shift !== 9 && shift !== 12) throw new CfbError("unsupported", `Compound file sector shift ${shift}`)
		if (header.getUint16(0x20, true) !== 6) throw new CfbError("unsupported", "Compound file mini sectors are not 64 bytes")
		if (header.getUint32(0x38, true) !== MINI_STREAM_CUTOFF)
			throw new CfbError("unsupported", "Compound file mini stream cutoff is not 4096")

		const sectorSize = 1 << shift

		if (source.size < sectorSize) throw new CfbError("truncated", "The compound file header sector is cut short")

		// A last sector cut short (some writers do not pad) still counts; reads past the end are refused.
		const sectorCount = Math.ceil((source.size - sectorSize) / sectorSize)
		const owners = new Uint32Array(sectorCount)
		const reader = new SectorReader(source, sectorSize, sectorCount)
		const fat = await reader.fat(header, owners)
		const directory = reader.chain(header.getUint32(0x30, true), fat, owners, OWNER_DIRECTORY, "directory", {
			limit: Math.ceil((limits.maxEntries * ENTRY_SIZE) / sectorSize),
			code: "tooManyEntries"
		})
		const root = readDirectory(await reader.sectors(directory), major)

		return new CfbFile(source, major, sectorSize, sectorCount, fat, owners, header.getUint32(0x3c, true), root)
	}

	// The entry at `path` (names separated by "/", or a list of names; matched as §2.6.4 compares them).
	entry(path: string | readonly string[]): CfbEntry | undefined {
		const names = typeof path === "string" ? path.split("/").filter(name => name !== "") : path
		let node = this.root as Node

		for (const name of names) {
			if (node.byName === undefined) {
				node.byName = new Map()

				for (const child of node.children) {
					const key = upperName(child.name)

					if (!node.byName.has(key)) node.byName.set(key, child)
				}
			}

			const child = node.byName.get(upperName(name))

			if (child === undefined) return undefined

			node = child
		}

		return node
	}

	// Every entry, the root first, each storage before its children.
	*walk(): Generator<CfbEntry> {
		const stack: Node[] = [this.root as Node]

		for (let node = stack.pop(); node !== undefined; node = stack.pop()) {
			yield node

			for (let k = node.children.length - 1; k >= 0; k--) stack.push(node.children[k] as Node)
		}
	}

	// A stream as a random-access source; a read inside one run of consecutive sectors is the source's own
	// read (a view for bytes in memory), any other is gathered into a new buffer.
	stream(target: CfbEntry | string | readonly string[]): ByteSource {
		const node = this.streamNode(target)
		const extents = this.extents(node)

		return {
			size: node.size,
			read: async (offset: number, length: number): Promise<Uint8Array> => {
				if (!(offset >= 0 && length >= 0 && offset + length <= node.size))
					throw new CfbError("truncated", "Read outside the stream")

				return this.gather(await extents, offset, length)
			}
		}
	}

	// A stream's bytes: a view of the source's bytes when they lie in one run, else a new buffer.
	async read(target: CfbEntry | string | readonly string[]): Promise<Uint8Array> {
		const node = this.streamNode(target)

		return this.gather(await this.extents(node), 0, node.size)
	}

	private streamNode(target: CfbEntry | string | readonly string[]): Node {
		const entry = target instanceof Node ? target : typeof target === "string" || isNames(target) ? this.entry(target) : undefined
		let top = entry instanceof Node ? entry : undefined

		while (top?.parent !== undefined) top = top.parent

		if (!(entry instanceof Node) || top !== this.root) throw new CfbError("notFound", "No such entry")
		if (entry.kind !== "stream") throw new CfbError("notStream", `${entry.name} is not a stream`)

		return entry
	}

	private async gather(extents: Extents, offset: number, length: number): Promise<Uint8Array> {
		if (length === 0) return new Uint8Array(0)

		const { starts, offsets } = extents
		const low = findRun(extents, offset)
		const runEnd = starts[low + 1] as number

		if (offset + length <= runEnd) return this.source.read((offsets[low] as number) + offset - (starts[low] as number), length)

		const out = new Uint8Array(length)

		for (let run = low, done = 0; done < length; run++) {
			const from = offset + done - (starts[run] as number)
			const count = Math.min(length - done, (starts[run + 1] as number) - (starts[run] as number) - from)

			out.set(await this.source.read((offsets[run] as number) + from, count), done)
			done += count
		}

		return out
	}

	private extents(node: Node): Promise<Extents> {
		let extents = this.extentsCache.get(node.id)

		if (extents === undefined) {
			extents = node.size >= MINI_STREAM_CUTOFF ? Promise.resolve().then(() => this.regularExtents(node)) : this.miniExtents(node)
			// Kept even when it fails: the chain is half marked as owned, so a second walk would misreport.
			extents.catch(() => undefined)
			this.extentsCache.set(node.id, extents)
		}

		return extents
	}

	private regularExtents(node: Node): Extents {
		if (node.size > this.source.size) throw corrupt(`${node.name} is larger than the file`)

		return sectorExtents(
			walkKnown(
				node.start,
				Math.ceil(node.size / this.sectorSize),
				this.fat,
				this.sectorCount,
				this.owners,
				OWNER_STREAM + node.id,
				node.name
			),
			this.sectorSize,
			node.size,
			id => (id + 1) * this.sectorSize,
			this.source.size
		)
	}

	private async miniExtents(node: Node): Promise<Extents> {
		const mini = await this.loadMini()

		if (node.size > mini.size) throw corrupt(`${node.name} is larger than the mini stream`)
		const ids = walkKnown(
			node.start,
			Math.ceil(node.size / MINI_SECTOR_SIZE),
			mini.fat,
			mini.count,
			mini.owners,
			node.id + 1,
			node.name
		)
		const runs = sectorExtents(ids, MINI_SECTOR_SIZE, node.size, id => id * MINI_SECTOR_SIZE, mini.size)
		const starts: number[] = []
		const offsets: number[] = []

		// Each run of mini sectors is a range of the mini stream, itself in runs of file sectors.
		for (let run = 0; run < runs.offsets.length; run++) {
			const end = (runs.starts[run + 1] as number) - (runs.starts[run] as number) + (runs.offsets[run] as number)

			for (let at = runs.offsets[run] as number; at < end;) {
				if (at >= mini.size) throw new CfbError("truncated", `${node.name} runs past the end of the mini stream`)

				const piece = findRun(mini.stream, at)
				const pieceEnd = Math.min(end, mini.stream.starts[piece + 1] as number)
				const fileOffset = (mini.stream.offsets[piece] as number) + at - (mini.stream.starts[piece] as number)
				const streamOffset = (runs.starts[run] as number) + at - (runs.offsets[run] as number)
				const last = offsets.length - 1

				if (last < 0 || (offsets[last] as number) + streamOffset - (starts[last] as number) !== fileOffset) {
					starts.push(streamOffset)
					offsets.push(fileOffset)
				}

				at = pieceEnd
			}
		}

		starts.push(node.size)

		return { starts: Float64Array.from(starts), offsets: Float64Array.from(offsets) }
	}

	private loadMini(): Promise<Mini> {
		this.mini ??= this.readMini()

		return this.mini
	}

	private async readMini(): Promise<Mini> {
		const root = this.root as Node

		if (root.size > this.source.size) throw corrupt("The mini stream is larger than the file")

		const count = Math.ceil(root.size / MINI_SECTOR_SIZE)
		const reader = new SectorReader(this.source, this.sectorSize, this.sectorCount)
		// Only the mini FAT sectors that describe mini sectors of the mini stream are read.
		const ids = reader.chain(this.miniFatStart, this.fat, this.owners, OWNER_MINI_FAT, "mini FAT", {
			limit: this.sectorCount,
			code: "corrupt"
		})
		const fat = toEntries(await reader.sectors(ids.slice(0, Math.ceil((count * 4) / this.sectorSize))))
		const stream = sectorExtents(
			walkKnown(
				root.start,
				Math.ceil(root.size / this.sectorSize),
				this.fat,
				this.sectorCount,
				this.owners,
				OWNER_MINI_STREAM,
				"mini stream"
			),
			this.sectorSize,
			root.size,
			id => (id + 1) * this.sectorSize,
			this.source.size
		)

		return { fat, stream, count, size: root.size, owners: new Uint32Array(count) }
	}
}

function isNames(target: CfbEntry | readonly string[]): target is readonly string[] {
	return Array.isArray(target)
}

function findRun(extents: Extents, offset: number): number {
	let low = 0
	let high = extents.offsets.length - 1

	while (low < high) {
		const middle = (low + high + 1) >> 1

		if ((extents.starts[middle] as number) <= offset) low = middle
		else high = middle - 1
	}

	return low
}

function toEntries(bytes: Uint8Array): Uint32Array {
	const data = view(bytes)
	const out = new Uint32Array(bytes.length >> 2)

	for (let k = 0; k < out.length; k++) out[k] = data.getUint32(k * 4, true)

	return out
}

// The first `count` sectors of the chain at `start`: each must be a sector of the file not owned by any
// chain yet; a chain ending (or leaving the file) before `count` is shorter than its stream.
function walkKnown(
	start: number,
	count: number,
	table: Uint32Array,
	total: number,
	owners: Uint32Array,
	owner: number,
	name: string
): Uint32Array {
	const ids = new Uint32Array(count)
	let current = start

	for (let k = 0; k < count; k++) {
		if (current >= total) {
			throw corrupt(
				current === ENDOFCHAIN || current === FREESECT
					? `The chain of ${name} is shorter than its size`
					: `The chain of ${name} leaves the file`
			)
		}

		const was = owners[current] as number

		if (was !== 0) throw corrupt(was === owner ? `The chain of ${name} loops` : `The chain of ${name} runs into another chain`)

		owners[current] = owner
		ids[k] = current
		current = current < table.length ? (table[current] as number) : FREESECT
	}

	return ids
}

// Runs of consecutive sectors among `ids`, the last one cut to `size`; no run may reach past `end` (the file's
// last sector can be cut short, and it need not be the stream's last).
function sectorExtents(ids: Uint32Array, sectorSize: number, size: number, offsetOf: (id: number) => number, end: number): Extents {
	const starts: number[] = []
	const offsets: number[] = []

	for (let k = 0; k <= ids.length; k++) {
		if (k > 0 && (k === ids.length || ids[k] !== (ids[k - 1] as number) + 1)) {
			const last = offsets.length - 1

			if ((offsets[last] as number) + Math.min(size, k * sectorSize) - (starts[last] as number) > end) {
				throw new CfbError("truncated", "A stream runs past the end of the file")
			}
		}

		if (k < ids.length && (k === 0 || ids[k] !== (ids[k - 1] as number) + 1)) {
			starts.push(k * sectorSize)
			offsets.push(offsetOf(ids[k] as number))
		}
	}

	starts.push(size)

	return { starts: Float64Array.from(starts), offsets: Float64Array.from(offsets) }
}

// Reads whole sectors, joining consecutive ones into one source read.
class SectorReader {
	private readonly source: ByteSource
	private readonly sectorSize: number
	private readonly count: number

	constructor(source: ByteSource, sectorSize: number, count: number) {
		this.source = source
		this.sectorSize = sectorSize
		this.count = count
	}

	async sectors(ids: ArrayLike<number>): Promise<Uint8Array> {
		const size = this.sectorSize
		const out = new Uint8Array(ids.length * size)

		for (let k = 0; k < ids.length;) {
			let next = k + 1

			while (next < ids.length && ids[next] === (ids[next - 1] as number) + 1) next++

			const offset = ((ids[k] as number) + 1) * size
			const length = (next - k) * size

			if (offset >= this.source.size) throw new CfbError("truncated", "A compound file sector lies past the end")

			// Only the file's last sector can be cut short: its missing tail reads as zeros.
			out.set(await this.source.read(offset, Math.min(length, this.source.size - offset)), k * size)
			k = next
		}

		return out
	}

	// A chain of unknown length (the directory, the mini FAT): it ends with ENDOFCHAIN within `limit` sectors.
	chain(
		start: number,
		fat: Uint32Array,
		owners: Uint32Array,
		owner: number,
		what: string,
		bound: { readonly limit: number; readonly code: "tooManyEntries" | "corrupt" }
	): number[] {
		const ids: number[] = []

		for (let current = start; current !== ENDOFCHAIN; current = current < fat.length ? (fat[current] as number) : FREESECT) {
			if (current >= this.count) throw corrupt(`The ${what} chain leaves the file`)
			if (ids.length >= bound.limit) throw new CfbError(bound.code, `The ${what} chain is longer than the limit`)

			const was = owners[current] as number

			if (was !== 0) throw corrupt(was === owner ? `The ${what} chain loops` : `The ${what} chain runs into another chain`)

			owners[current] = owner
			ids.push(current)
		}

		return ids
	}

	// The FAT entries for every sector of the file: only the FAT sectors that describe existing sectors are
	// read (the header may count more); sectors no FAT sector describes are free.
	async fat(header: DataView, owners: Uint32Array): Promise<Uint32Array> {
		const perSector = this.sectorSize / 4
		const count = Math.min(header.getUint32(0x2c, true), Math.ceil(this.count / perSector))
		const ids = new Uint32Array(count)
		let filled = 0

		for (; filled < Math.min(HEADER_DIFAT, count); filled++) ids[filled] = header.getUint32(0x4c + filled * 4, true)

		for (let next = header.getUint32(0x44, true); filled < count;) {
			if (next >= this.count) throw corrupt("The DIFAT ends before listing every FAT sector")
			if (owners[next] !== 0) throw corrupt("The DIFAT chain loops")

			owners[next] = OWNER_FAT

			const difat = view(await this.sectors([next]))

			for (let k = 0; k < perSector - 1 && filled < count; k++) ids[filled++] = difat.getUint32(k * 4, true)

			next = difat.getUint32(this.sectorSize - 4, true)
		}

		for (const id of ids) {
			if (id >= this.count) throw corrupt("A FAT sector lies outside the file")
			if (owners[id] !== 0) throw corrupt("A FAT sector is listed twice")

			owners[id] = OWNER_FAT
		}

		return toEntries(await this.sectors(ids))
	}
}

// Builds the tree under the root from the directory sectors: in-order walks of each storage's red-black tree
// (the colours are not checked: §2.6.4 has readers rely only on the order), every entry reached at most once.
// Sizes are checked when a stream is read, so one damaged entry (a summary stream with a garbage size, say)
// leaves the rest of the file readable.
function readDirectory(bytes: Uint8Array, major: number): Node {
	const data = view(bytes)
	const count = bytes.length / ENTRY_SIZE
	const seen = new Uint8Array(count)
	const read32 = (id: number, field: number): number => data.getUint32(id * ENTRY_SIZE + field, true)
	// Version 3 readers ignore the high half, which old writers left uninitialised (§2.6.3).
	const size = (id: number): number => (major === 3 ? 0 : read32(id, 0x7c) * 0x1_0000_0000) + read32(id, 0x78)
	const name = (id: number): string => {
		const length = data.getUint16(id * ENTRY_SIZE + 0x40, true)

		if (length < 2 || length > 64 || length % 2 !== 0) throw corrupt(`Directory entry ${id} has a bad name length`)

		let out = ""

		for (let k = 0; k < length - 2; k += 2) out += String.fromCharCode(data.getUint16(id * ENTRY_SIZE + k, true))

		return out
	}

	if (count === 0 || data.getUint8(0x42) !== TYPE_ROOT) throw corrupt("The first directory entry is not the root")

	const root = new Node(0, name(0), "root", size(0), formatClsid(bytes, 0x50), read32(0, 0x74), undefined)
	const storages: Node[] = [root]
	const pending: number[] = []

	seen[0] = 1

	for (let storage = storages.pop(); storage !== undefined; storage = storages.pop()) {
		let current = read32(storage.id, 0x4c)

		while (current !== NOSTREAM || pending.length > 0) {
			while (current !== NOSTREAM) {
				if (current >= count) throw corrupt("A directory entry points outside the directory")
				if (seen[current] !== 0) throw corrupt("The directory tree loops")

				seen[current] = 1
				pending.push(current)
				current = read32(current, 0x44)
			}

			const id = pending.pop() as number
			const type = data.getUint8(id * ENTRY_SIZE + 0x42)

			if (type !== TYPE_STORAGE && type !== TYPE_STREAM) throw corrupt(`Directory entry ${id} in the tree is not a storage or stream`)

			const stream = type === TYPE_STREAM
			const node = new Node(
				id,
				name(id),
				stream ? "stream" : "storage",
				stream ? size(id) : 0,
				formatClsid(bytes, id * ENTRY_SIZE + 0x50),
				read32(id, 0x74),
				storage
			)

			storage.children.push(node)

			if (!stream) storages.push(node)

			current = read32(id, 0x48)
		}
	}

	return root
}
