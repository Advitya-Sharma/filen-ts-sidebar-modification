// CRC-32 (ISO-HDLC, as zip uses), slicing by eight: eight 256-entry tables, eight bytes per step.
const TABLES = (() => {
	const tables = new Int32Array(8 * 256)

	for (let n = 0; n < 256; n++) {
		let c = n

		for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1

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

// Continue a CRC over more bytes: crc32(b, crc32(a)) === crc32(a + b).
export function crc32(bytes: Uint8Array, previous: number = 0): number {
	const t = TABLES
	let crc = ~previous
	let i = 0
	const fast = bytes.length - 8

	for (; i <= fast; i += 8) {
		const low = crc ^ ((bytes[i] ?? 0) | ((bytes[i + 1] ?? 0) << 8) | ((bytes[i + 2] ?? 0) << 16) | ((bytes[i + 3] ?? 0) << 24))
		const high = (bytes[i + 4] ?? 0) | ((bytes[i + 5] ?? 0) << 8) | ((bytes[i + 6] ?? 0) << 16) | ((bytes[i + 7] ?? 0) << 24)

		crc =
			(t[1792 + (low & 0xff)] ?? 0) ^
			(t[1536 + ((low >>> 8) & 0xff)] ?? 0) ^
			(t[1280 + ((low >>> 16) & 0xff)] ?? 0) ^
			(t[1024 + (low >>> 24)] ?? 0) ^
			(t[768 + (high & 0xff)] ?? 0) ^
			(t[512 + ((high >>> 8) & 0xff)] ?? 0) ^
			(t[256 + ((high >>> 16) & 0xff)] ?? 0) ^
			(t[high >>> 24] ?? 0)
	}

	for (; i < bytes.length; i++) crc = (t[(crc ^ (bytes[i] ?? 0)) & 0xff] ?? 0) ^ (crc >>> 8)

	return ~crc >>> 0
}
