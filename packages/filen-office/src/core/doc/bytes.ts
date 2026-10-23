import { corrupt } from "./errors"

// Little-endian reads over one stream; every read is bounds-checked and a miss is a typed corrupt error.
export class Bytes {
	readonly data: Uint8Array
	readonly view: DataView
	readonly name: string

	constructor(data: Uint8Array, name: string) {
		this.data = data
		this.view = new DataView(data.buffer, data.byteOffset, data.byteLength)
		this.name = name
	}

	get length(): number {
		return this.data.length
	}

	has(offset: number, length: number): boolean {
		return Number.isInteger(offset) && Number.isInteger(length) && offset >= 0 && length >= 0 && offset + length <= this.data.length
	}

	check(offset: number, length: number, what: string): void {
		if (!this.has(offset, length)) throw corrupt(`${what} at ${offset} (+${length}) is outside the ${this.name} stream`)
	}

	u8(offset: number): number {
		this.check(offset, 1, "byte")

		return this.view.getUint8(offset)
	}

	u16(offset: number): number {
		this.check(offset, 2, "u16")

		return this.view.getUint16(offset, true)
	}

	i16(offset: number): number {
		this.check(offset, 2, "i16")

		return this.view.getInt16(offset, true)
	}

	u32(offset: number): number {
		this.check(offset, 4, "u32")

		return this.view.getUint32(offset, true)
	}

	i32(offset: number): number {
		this.check(offset, 4, "i32")

		return this.view.getInt32(offset, true)
	}

	// A view (no copy).
	slice(offset: number, length: number, what: string): Uint8Array {
		this.check(offset, length, what)

		return this.data.subarray(offset, offset + length)
	}
}

// UTF-16LE code units to a string, in chunks so long runs never exceed the argument limit.
export function utf16(data: Uint8Array, offset: number, count: number): string {
	const view = new DataView(data.buffer, data.byteOffset, data.byteLength)
	const parts: string[] = []
	const units = new Array<number>(Math.min(count, 4096))

	for (let done = 0; done < count;) {
		const take = Math.min(4096, count - done)

		units.length = take

		for (let k = 0; k < take; k++) units[k] = view.getUint16(offset + 2 * (done + k), true)

		parts.push(String.fromCharCode(...units))
		done += take
	}

	return parts.join("")
}

// Little-endian helpers over a Uint8Array view; the caller has bounds-checked.
export function le16(data: Uint8Array, at: number): number {
	return (data[at] ?? 0) | ((data[at + 1] ?? 0) << 8)
}

export function les16(data: Uint8Array, at: number): number {
	return (le16(data, at) << 16) >> 16
}

export function le32(data: Uint8Array, at: number): number {
	return ((data[at] ?? 0) | ((data[at + 1] ?? 0) << 8) | ((data[at + 2] ?? 0) << 16) | ((data[at + 3] ?? 0) << 24)) >>> 0
}

export function les32(data: Uint8Array, at: number): number {
	return le32(data, at) | 0
}
