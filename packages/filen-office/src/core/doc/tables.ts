import { type Bytes, utf16 } from "./bytes"
import { corrupt } from "./errors"

// Generic Table-stream containers: STTB (§2.2.4), PLC (§2.2.2) and Xst (§2.9.353).

export interface SttbEntry {
	readonly text: string
	// The ExtraData bytes (a view).
	readonly extra: Uint8Array
}

// Strings are capped so a hostile cch cannot build a huge string.
const MAX_STTB_ENTRIES = 1 << 16

// cData is 2 bytes unless `wideCount` (a few STTBs use a 4-byte cData).
export function readSttb(table: Bytes, fc: number, lcb: number, wideCount = false): SttbEntry[] {
	if (lcb === 0) return []

	table.check(fc, lcb, "STTB")

	const end = fc + lcb
	let at = fc
	const extended = table.u16(at) === 0xffff

	if (extended) at += 2

	const count = wideCount ? table.i32(at) : table.u16(at)

	at += wideCount ? 4 : 2

	const cbExtra = table.u16(at)

	at += 2

	if (count < 0 || count > MAX_STTB_ENTRIES) throw corrupt("STTB count is invalid")

	const out: SttbEntry[] = []

	for (let k = 0; k < count && at < end; k++) {
		const cch = extended ? table.u16(at) : table.u8(at)

		at += extended ? 2 : 1

		const size = extended ? 2 * cch : cch

		if (at + size + cbExtra > end) break

		const text = extended ? utf16(table.data, at, cch) : latin1(table.data, at, cch)

		at += size
		out.push({ text, extra: table.data.subarray(at, at + cbExtra) })
		at += cbExtra
	}

	return out
}

function latin1(data: Uint8Array, at: number, count: number): string {
	let out = ""

	for (let k = 0; k < count; k++) out += String.fromCharCode(data[at + k] ?? 0)

	return out
}

export interface Plc {
	// n + 1 positions.
	readonly cps: Uint32Array
	// Offset of the data elements in the Table stream; element k is at dataAt + k * size.
	readonly dataAt: number
	readonly size: number
	readonly count: number
}

// A PLC of elements of `size` bytes: lcb = 4(n + 1) + size·n. Undefined when absent or when lcb does not fit
// (the caller reports the structure as unreadable).
export function readPlc(table: Bytes, fc: number, lcb: number, size: number): Plc | undefined {
	if (lcb < 4 || !table.has(fc, lcb) || (lcb - 4) % (4 + size) !== 0) return undefined

	const count = (lcb - 4) / (4 + size)
	const cps = new Uint32Array(count + 1)

	for (let k = 0; k <= count; k++) cps[k] = table.u32(fc + 4 * k)

	return { cps, dataAt: fc + 4 * (count + 1), size, count }
}

// Xst (§2.9.353): cch, then cch UTF-16 code units. Returns the string and the offset after it.
export function readXst(bytes: Bytes, at: number): readonly [string, number] {
	const cch = bytes.u16(at)

	bytes.check(at + 2, 2 * cch, "Xst")

	return [utf16(bytes.data, at + 2, cch), at + 2 + 2 * cch]
}
