// A small [MS-CFB] stream reader for the encryption tests: they need the EncryptionInfo, EncryptedPackage and
// DataSpaces streams of trusted corpus files, independently of the package's own CFB reader.

interface Entry {
	readonly name: string
	readonly type: number
	readonly left: number
	readonly right: number
	readonly child: number
	readonly start: number
	readonly size: number
}

const END = 0xfffffffa
const NONE = 0xffffffff

export interface CfbStreams {
	// The stream at a path of names from the root, e.g. ["\u0006DataSpaces", "DataSpaceMap"].
	stream(...path: string[]): Uint8Array | undefined
}

export function readCfbStreams(file: Uint8Array): CfbStreams {
	const view = new DataView(file.buffer, file.byteOffset, file.byteLength)
	const sectorSize = 1 << view.getUint16(0x1e, true)
	const miniSize = 1 << view.getUint16(0x20, true)
	const fatSectors = view.getUint32(0x2c, true)
	const cutoff = view.getUint32(0x38, true)
	const sector = (index: number): Uint8Array => file.subarray((index + 1) * sectorSize, (index + 2) * sectorSize)
	const difat: number[] = []

	for (let k = 0; k < 109; k++) difat.push(view.getUint32(0x4c + 4 * k, true))

	for (let next = view.getUint32(0x44, true), guard = 0; next < END && guard < 1 << 16; guard++) {
		const s = new DataView(file.buffer, file.byteOffset + (next + 1) * sectorSize, sectorSize)

		for (let k = 0; k < sectorSize / 4 - 1; k++) difat.push(s.getUint32(4 * k, true))

		next = s.getUint32(sectorSize - 4, true)
	}

	const fat: number[] = []

	for (let k = 0; k < fatSectors; k++) {
		const s = new DataView(file.buffer, file.byteOffset + ((difat[k] ?? 0) + 1) * sectorSize, sectorSize)

		for (let j = 0; j < sectorSize / 4; j++) fat.push(s.getUint32(4 * j, true))
	}

	const chain = (start: number, table: readonly number[]): number[] => {
		const out: number[] = []

		for (let at = start; at < END && out.length <= table.length; at = table[at] ?? NONE) out.push(at)

		return out
	}
	const read = (start: number): Uint8Array => {
		const sectors = chain(start, fat)
		const out = new Uint8Array(sectors.length * sectorSize)

		sectors.forEach((s, k) => out.set(sector(s), k * sectorSize))

		return out
	}
	const directory = read(view.getUint32(0x30, true))
	const dv = new DataView(directory.buffer)
	const entries: Entry[] = []

	for (let o = 0; o + 128 <= directory.length; o += 128) {
		const length = dv.getUint16(o + 0x40, true)
		let name = ""

		for (let k = 0; k + 2 < length; k += 2) name += String.fromCharCode(dv.getUint16(o + k, true))

		entries.push({
			name,
			type: directory[o + 0x42] ?? 0,
			left: dv.getUint32(o + 0x44, true),
			right: dv.getUint32(o + 0x48, true),
			child: dv.getUint32(o + 0x4c, true),
			start: dv.getUint32(o + 0x74, true),
			size: dv.getUint32(o + 0x78, true)
		})
	}

	const root = entries[0]

	if (root === undefined) throw new Error("No root entry")

	const mini = read(root.start)
	const miniFatStart = view.getUint32(0x3c, true)
	const miniFatBytes = miniFatStart < END ? read(miniFatStart) : new Uint8Array(0)
	const miniFat = Array.from({ length: miniFatBytes.length / 4 }, (_, k) => new DataView(miniFatBytes.buffer).getUint32(4 * k, true))

	const children = (storage: Entry): Entry[] => {
		const out: Entry[] = []
		const stack = [storage.child]

		while (stack.length > 0 && out.length < entries.length) {
			const id = stack.pop() ?? NONE
			const entry = entries[id]

			if (id === NONE || entry === undefined) continue

			out.push(entry)
			stack.push(entry.left, entry.right)
		}

		return out
	}

	return {
		stream(...path) {
			let entry: Entry | undefined = root

			for (const name of path) entry = entry === undefined ? undefined : children(entry).find(c => c.name === name)

			if (entry === undefined || entry.type !== 2) return undefined

			if (entry.size < cutoff) {
				const sectors = chain(entry.start, miniFat)
				const out = new Uint8Array(sectors.length * miniSize)

				sectors.forEach((s, k) => out.set(mini.subarray(s * miniSize, (s + 1) * miniSize), k * miniSize))

				return out.slice(0, entry.size)
			}

			return read(entry.start).slice(0, entry.size)
		}
	}
}
