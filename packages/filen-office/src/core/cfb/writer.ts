import type { ByteSink } from "../opc/io"
import type { PartBytes } from "../opc/pieces"
import { CfbError } from "./errors"
import {
	CFB_SIGNATURE,
	DIFSECT,
	ENDOFCHAIN,
	ENTRY_SIZE,
	FATSECT,
	FREESECT,
	HEADER_DIFAT,
	MAXREGSECT,
	MINI_SECTOR_SIZE,
	MINI_STREAM_CUTOFF,
	NOSTREAM,
	TYPE_ROOT,
	TYPE_STORAGE,
	TYPE_STREAM,
	compareNames,
	upperName
} from "./reader"

// A stream, or a storage declared on its own (to give it a CLSID, or to keep it when empty). Storages on the
// path to a stream need no entry.
export type CfbWriteEntry =
	| { readonly path: readonly string[]; readonly bytes: PartBytes }
	| { readonly path: readonly string[]; readonly storage: true; readonly clsid?: string }

export interface CfbWriteOptions {
	// 3: 512-byte sectors (the default); 4: 4096-byte sectors.
	readonly version?: 3 | 4
	readonly rootClsid?: string
}

interface PlannedEntry {
	readonly name: string
	readonly type: number
	clsid: Uint8Array | undefined
	readonly bytes: PartBytes | undefined
	readonly children: PlannedEntry[]
	id: number
	left: number
	right: number
	child: number
	// First sector (regular or mini) and its count; ENDOFCHAIN when the stream is empty or a storage.
	start: number
	sectors: number
	mini: boolean
}

// A planned file: its size is known before a byte is written (an upload can declare it up front).
export interface CfbPlan {
	readonly size: number
	readonly version: 3 | 4
}

interface Layout extends CfbPlan {
	readonly entries: readonly PlannedEntry[]
	readonly sectorSize: number
	readonly fatSectors: number
	readonly difatSectors: number
	readonly directorySectors: number
	readonly miniFatSectors: number
	readonly miniSectors: number
	readonly miniStreamSectors: number
	// First sector of each area; the regular streams follow the mini stream in directory order.
	readonly directoryStart: number
	readonly miniFatStart: number
	readonly miniStreamStart: number
}

// The layouts behind the plans planCfb handed out.
const layouts = new WeakMap<CfbPlan, Layout>()

const MAX_NAME = 31
const V3_MAX_STREAM = 0x80000000

function checkName(name: string): void {
	if (name.length === 0 || name.length > MAX_NAME || /[/\\:!]/u.test(name)) {
		throw new CfbError("invalidName", `Compound file names are 1-31 characters without / \\ : !: "${name}"`)
	}
}

export function parseClsid(clsid: string): Uint8Array {
	const match = /^\{?([0-9a-f]{8})-([0-9a-f]{4})-([0-9a-f]{4})-([0-9a-f]{4})-([0-9a-f]{12})\}?$/iu.exec(clsid)

	if (match === null) throw new CfbError("invalidClsid", `Not a CLSID: ${clsid}`)

	const out = new Uint8Array(16)
	const data = new DataView(out.buffer)
	const tail = (match[4] as string) + (match[5] as string)

	data.setUint32(0, parseInt(match[1] as string, 16), true)
	data.setUint16(4, parseInt(match[2] as string, 16), true)
	data.setUint16(6, parseInt(match[3] as string, 16), true)

	for (let k = 0; k < 8; k++) out[8 + k] = parseInt(tail.slice(k * 2, k * 2 + 2), 16)

	return out
}

function entry(name: string, type: number, bytes: PartBytes | undefined): PlannedEntry {
	return {
		name,
		type,
		clsid: undefined,
		bytes,
		children: [],
		id: 0,
		left: NOSTREAM,
		right: NOSTREAM,
		child: NOSTREAM,
		start: ENDOFCHAIN,
		sectors: 0,
		mini: false
	}
}

// Links sorted siblings [from, to) as a balanced binary tree and returns its root's id. All nodes are black,
// which §2.6.4 allows; readers rely only on the order.
function link(siblings: readonly PlannedEntry[], from: number, to: number): number {
	if (from >= to) return NOSTREAM

	const middle = (from + to) >> 1
	const node = siblings[middle] as PlannedEntry

	node.left = link(siblings, from, middle)
	node.right = link(siblings, middle + 1, to)

	return node.id
}

// Lays out a compound file holding `entries`: the directory in breadth-first order of storages (each one's
// children sorted), streams under 4096 bytes in the mini stream, the rest one after another in whole
// sectors. Nothing varies but the input: timestamps and unset CLSIDs are zero.
export function planCfb(entries: readonly CfbWriteEntry[], options: CfbWriteOptions = {}): CfbPlan {
	const version = options.version ?? 3
	const sectorSize = version === 4 ? 4096 : 512
	const root = entry("Root Entry", TYPE_ROOT, undefined)

	if (options.rootClsid !== undefined) root.clsid = parseClsid(options.rootClsid)

	const childMaps = new Map<PlannedEntry, Map<string, PlannedEntry>>([[root, new Map()]])
	const declared = new Set<PlannedEntry>()

	for (const item of entries) {
		if (item.path.length === 0) throw new CfbError("invalidName", "An entry needs a name")

		let parent = root

		for (let k = 0; k < item.path.length; k++) {
			const name = item.path[k] as string
			const last = k === item.path.length - 1
			const siblings = childMaps.get(parent) as Map<string, PlannedEntry>
			const key = upperName(name)
			const existing = siblings.get(key)

			checkName(name)

			if (last && "bytes" in item) {
				if (existing !== undefined) throw new CfbError("duplicateName", `${item.path.join("/")} is given twice`)
				if (version === 3 && item.bytes.length > V3_MAX_STREAM)
					throw new CfbError("tooLarge", `${name} is too large for a version 3 file`)

				siblings.set(key, entry(name, TYPE_STREAM, item.bytes))

				break
			}

			let node = existing

			if (node === undefined) {
				node = entry(name, TYPE_STORAGE, undefined)
				siblings.set(key, node)
				childMaps.set(node, new Map())
			} else if (node.type !== TYPE_STORAGE) {
				throw new CfbError("duplicateName", `${item.path.slice(0, k + 1).join("/")} is a stream`)
			}

			if (last) {
				if (declared.has(node)) throw new CfbError("duplicateName", `${item.path.join("/")} is given twice`)

				declared.add(node)

				if ("storage" in item && item.clsid !== undefined) node.clsid = parseClsid(item.clsid)
			}

			parent = node
		}
	}

	// Ids: the root, then each storage's children together, storages in breadth-first order.
	const ordered: PlannedEntry[] = [root]

	for (let k = 0; k < ordered.length; k++) {
		const storage = ordered[k] as PlannedEntry
		const children = [...(childMaps.get(storage)?.values() ?? [])].sort((a, b) => compareNames(a.name, b.name))

		for (const child of children) {
			child.id = ordered.length
			ordered.push(child)
			storage.children.push(child)
		}

		storage.child = link(children, 0, children.length)
	}

	let miniSectors = 0
	let regularSectors = 0

	for (const node of ordered) {
		const length = node.bytes?.length ?? 0

		if (node.type !== TYPE_STREAM || length === 0) continue

		node.mini = length < MINI_STREAM_CUTOFF
		node.sectors = Math.ceil(length / (node.mini ? MINI_SECTOR_SIZE : sectorSize))

		if (node.mini) {
			node.start = miniSectors
			miniSectors += node.sectors
		} else {
			regularSectors += node.sectors
		}
	}

	const perSector = sectorSize / 4
	const directorySectors = Math.ceil((ordered.length * ENTRY_SIZE) / sectorSize)
	const miniFatSectors = Math.ceil(miniSectors / perSector)
	const miniStreamSectors = Math.ceil((miniSectors * MINI_SECTOR_SIZE) / sectorSize)
	const dataSectors = directorySectors + miniFatSectors + miniStreamSectors + regularSectors
	let fatSectors = 0
	let difatSectors = 0

	// The FAT describes itself and the DIFAT too: grow both until they cover everything.
	for (;;) {
		const fat = Math.ceil((dataSectors + fatSectors + difatSectors) / perSector)
		const difat = fat > HEADER_DIFAT ? Math.ceil((fat - HEADER_DIFAT) / (perSector - 1)) : 0

		if (fat === fatSectors && difat === difatSectors) break

		fatSectors = fat
		difatSectors = difat
	}

	const totalSectors = dataSectors + fatSectors + difatSectors

	if (totalSectors > MAXREGSECT) throw new CfbError("tooLarge", "Too much content for a compound file")

	const directoryStart = fatSectors + difatSectors
	const miniFatStart = directoryStart + directorySectors
	const miniStreamStart = miniFatStart + miniFatSectors
	let next = miniStreamStart + miniStreamSectors

	for (const node of ordered) {
		if (node.type === TYPE_STREAM && node.sectors > 0 && !node.mini) {
			node.start = next
			next += node.sectors
		}
	}

	root.start = miniSectors > 0 ? miniStreamStart : ENDOFCHAIN

	const plan: CfbPlan = { size: (totalSectors + 1) * sectorSize, version }

	layouts.set(plan, {
		...plan,
		entries: ordered,
		sectorSize,
		fatSectors,
		difatSectors,
		directorySectors,
		miniFatSectors,
		miniSectors,
		miniStreamSectors,
		directoryStart,
		miniFatStart,
		miniStreamStart
	})

	return plan
}

// Writes are gathered into chunks of this size; larger content goes to the sink as it is.
const BATCH = 64 * 1024

class Output {
	private readonly sink: ByteSink
	private batch = new Uint8Array(BATCH)
	private filled = 0
	written = 0

	constructor(sink: ByteSink) {
		this.sink = sink
	}

	async put(bytes: Uint8Array): Promise<void> {
		this.written += bytes.length

		if (bytes.length >= BATCH) {
			await this.flush()
			await this.sink.write(bytes)

			return
		}

		if (this.filled + bytes.length > BATCH) await this.flush()

		this.batch.set(bytes, this.filled)
		this.filled += bytes.length
	}

	async zeros(count: number): Promise<void> {
		for (let left = count; left > 0;) {
			const step = Math.min(left, BATCH - this.filled)

			// A fresh batch is already zero.
			this.filled += step
			this.written += step
			left -= step

			if (this.filled === BATCH) await this.flush()
		}
	}

	async flush(): Promise<void> {
		if (this.filled === 0) return

		const out = this.batch.subarray(0, this.filled)

		this.batch = new Uint8Array(BATCH)
		this.filled = 0
		await this.sink.write(out)
	}
}

async function putContent(out: Output, bytes: PartBytes, name: string): Promise<void> {
	if (bytes instanceof Uint8Array) {
		await out.put(bytes)

		return
	}

	let total = 0

	for await (const chunk of bytes.chunks()) {
		total += chunk.length

		if (total > bytes.length) break

		await out.put(chunk)
	}

	if (total !== bytes.length) throw new CfbError("lengthMismatch", `${name} does not have the length it declares`)
}

function sectorTable(sectors: number, sectorSize: number): { readonly bytes: Uint8Array; readonly data: DataView } {
	const bytes = new Uint8Array(sectors * sectorSize).fill(0xff)

	return { bytes, data: new DataView(bytes.buffer) }
}

function header(plan: Layout): Uint8Array {
	const bytes = new Uint8Array(plan.sectorSize)
	const data = new DataView(bytes.buffer)

	bytes.set(CFB_SIGNATURE)
	data.setUint16(0x18, 0x003e, true)
	data.setUint16(0x1a, plan.version, true)
	data.setUint16(0x1c, 0xfffe, true)
	data.setUint16(0x1e, plan.version === 4 ? 12 : 9, true)
	data.setUint16(0x20, 6, true)
	data.setUint32(0x28, plan.version === 4 ? plan.directorySectors : 0, true)
	data.setUint32(0x2c, plan.fatSectors, true)
	data.setUint32(0x30, plan.directoryStart, true)
	data.setUint32(0x38, MINI_STREAM_CUTOFF, true)
	data.setUint32(0x3c, plan.miniFatSectors > 0 ? plan.miniFatStart : ENDOFCHAIN, true)
	data.setUint32(0x40, plan.miniFatSectors, true)
	data.setUint32(0x44, plan.difatSectors > 0 ? plan.fatSectors : ENDOFCHAIN, true)
	data.setUint32(0x48, plan.difatSectors, true)

	for (let k = 0; k < HEADER_DIFAT; k++) data.setUint32(0x4c + k * 4, k < plan.fatSectors ? k : FREESECT, true)

	return bytes
}

// The FAT and the DIFAT sectors after it.
function allocationTables(plan: Layout): Uint8Array {
	const { sectorSize, fatSectors, difatSectors } = plan
	const { bytes, data } = sectorTable(fatSectors + difatSectors, sectorSize)
	const set = (sector: number, value: number): void => data.setUint32(sector * 4, value, true)
	const chain = (start: number, count: number): void => {
		for (let k = 0; k < count; k++) set(start + k, k === count - 1 ? ENDOFCHAIN : start + k + 1)
	}

	for (let k = 0; k < fatSectors; k++) set(k, FATSECT)
	for (let k = 0; k < difatSectors; k++) set(fatSectors + k, DIFSECT)

	chain(plan.directoryStart, plan.directorySectors)
	chain(plan.miniFatStart, plan.miniFatSectors)
	chain(plan.miniStreamStart, plan.miniStreamSectors)

	for (const node of plan.entries) if (node.type === TYPE_STREAM && !node.mini) chain(node.start, node.sectors)

	// DIFAT sectors: the FAT sector numbers past the header's, the next DIFAT sector last.
	const perSector = sectorSize / 4
	const difatAt = fatSectors * sectorSize

	for (let k = 0; k < difatSectors; k++) {
		const at = difatAt + k * sectorSize

		for (let slot = 0; slot < perSector - 1; slot++) {
			const fat = HEADER_DIFAT + k * (perSector - 1) + slot

			if (fat < fatSectors) data.setUint32(at + slot * 4, fat, true)
		}

		data.setUint32(at + sectorSize - 4, k === difatSectors - 1 ? ENDOFCHAIN : fatSectors + k + 1, true)
	}

	return bytes
}

function directory(plan: Layout): Uint8Array {
	const bytes = new Uint8Array(plan.directorySectors * plan.sectorSize)
	const data = new DataView(bytes.buffer)

	for (let id = plan.entries.length; id < bytes.length / ENTRY_SIZE; id++) {
		data.setUint32(id * ENTRY_SIZE + 0x44, NOSTREAM, true)
		data.setUint32(id * ENTRY_SIZE + 0x48, NOSTREAM, true)
		data.setUint32(id * ENTRY_SIZE + 0x4c, NOSTREAM, true)
	}

	for (const node of plan.entries) {
		const at = node.id * ENTRY_SIZE
		const size = node.type === TYPE_ROOT ? plan.miniSectors * MINI_SECTOR_SIZE : (node.bytes?.length ?? 0)

		for (let k = 0; k < node.name.length; k++) data.setUint16(at + k * 2, node.name.charCodeAt(k), true)

		data.setUint16(at + 0x40, (node.name.length + 1) * 2, true)
		data.setUint8(at + 0x42, node.type)
		data.setUint8(at + 0x43, 1)
		data.setUint32(at + 0x44, node.left, true)
		data.setUint32(at + 0x48, node.right, true)
		data.setUint32(at + 0x4c, node.child, true)

		if (node.clsid !== undefined) bytes.set(node.clsid, at + 0x50)
		if (node.type !== TYPE_STORAGE) data.setUint32(at + 0x74, node.start, true)

		data.setUint32(at + 0x78, size % 0x1_0000_0000, true)
		data.setUint32(at + 0x7c, Math.floor(size / 0x1_0000_0000), true)
	}

	return bytes
}

function miniFat(plan: Layout): Uint8Array {
	const { bytes, data } = sectorTable(plan.miniFatSectors, plan.sectorSize)

	for (const node of plan.entries) {
		if (!node.mini) continue

		for (let k = 0; k < node.sectors; k++)
			data.setUint32((node.start + k) * 4, k === node.sectors - 1 ? ENDOFCHAIN : node.start + k + 1, true)
	}

	return bytes
}

function padding(length: number, unit: number): number {
	return (unit - (length % unit)) % unit
}

// Streams the planned file to `sink` (closing it is the caller's).
export async function writeCfb(planned: CfbPlan, sink: ByteSink): Promise<void> {
	const plan = layouts.get(planned)

	if (plan === undefined) throw new CfbError("notFound", "Not a plan from planCfb")

	const out = new Output(sink)

	await out.put(header(plan))
	await out.put(allocationTables(plan))
	await out.put(directory(plan))
	await out.put(miniFat(plan))

	for (const node of plan.entries) {
		if (!node.mini || node.bytes === undefined) continue

		await putContent(out, node.bytes, node.name)
		await out.zeros(padding(node.bytes.length, MINI_SECTOR_SIZE))
	}

	await out.zeros(padding(plan.miniSectors * MINI_SECTOR_SIZE, plan.sectorSize))

	for (const node of plan.entries) {
		if (node.mini || node.sectors === 0 || node.bytes === undefined) continue

		await putContent(out, node.bytes, node.name)
		await out.zeros(padding(node.bytes.length, plan.sectorSize))
	}

	await out.flush()

	if (out.written !== plan.size) throw new CfbError("lengthMismatch", "The compound file came out a different size than planned")
}

// The compound file holding `entries`, in one buffer.
export async function cfbBytes(entries: readonly CfbWriteEntry[], options: CfbWriteOptions = {}): Promise<Uint8Array> {
	const plan = planCfb(entries, options)
	const bytes = new Uint8Array(plan.size)
	let at = 0

	await writeCfb(plan, {
		write(chunk: Uint8Array): Promise<void> {
			if (at + chunk.length <= bytes.length) bytes.set(chunk, at)

			at += chunk.length

			return Promise.resolve()
		},
		close(): Promise<void> {
			return Promise.resolve()
		}
	})

	return bytes
}
