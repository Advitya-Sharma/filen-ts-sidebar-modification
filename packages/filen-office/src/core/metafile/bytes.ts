// Little-endian reads that never go past the end: a read beyond it returns 0, so a truncated record yields
// zeros rather than an exception or data from a neighbouring buffer.
export class Bytes {
	readonly bytes: Uint8Array
	readonly length: number
	private readonly view: DataView

	constructor(bytes: Uint8Array) {
		this.bytes = bytes
		this.length = bytes.length
		this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
	}

	u8(at: number): number {
		return at >= 0 && at < this.length ? (this.bytes[at] ?? 0) : 0
	}

	u16(at: number): number {
		return at >= 0 && at + 2 <= this.length ? this.view.getUint16(at, true) : 0
	}

	i16(at: number): number {
		return at >= 0 && at + 2 <= this.length ? this.view.getInt16(at, true) : 0
	}

	u32(at: number): number {
		return at >= 0 && at + 4 <= this.length ? this.view.getUint32(at, true) : 0
	}

	i32(at: number): number {
		return at >= 0 && at + 4 <= this.length ? this.view.getInt32(at, true) : 0
	}

	// Non-finite floats read as 0.
	f32(at: number): number {
		if (at < 0 || at + 4 > this.length) return 0

		const value = this.view.getFloat32(at, true)

		return Number.isFinite(value) ? value : 0
	}

	// True when [at, at + size) lies inside the buffer.
	has(at: number, size: number): boolean {
		return at >= 0 && size >= 0 && at + size <= this.length
	}

	// A view of [at, at + size) clamped to the buffer.
	sub(at: number, size: number): Uint8Array {
		const start = Math.max(0, Math.min(at, this.length))

		return this.bytes.subarray(start, Math.max(start, Math.min(start + Math.max(0, size), this.length)))
	}

	// UTF-16LE text of up to `count` units at `at`, stopping at the first NUL when `nul` is set.
	utf16(at: number, count: number, nul: boolean): string {
		const end = Math.min(this.length, at + Math.max(0, count) * 2)
		let out = ""
		let chunk: number[] = []

		for (let p = Math.max(0, at); p + 2 <= end; p += 2) {
			const unit = this.view.getUint16(p, true)

			if (nul && unit === 0) break

			chunk.push(unit)

			// fromCharCode with a bounded argument list.
			if (chunk.length === 4096) {
				out += String.fromCharCode(...chunk)
				chunk = []
			}
		}

		return chunk.length === 0 ? out : out + String.fromCharCode(...chunk)
	}
}
