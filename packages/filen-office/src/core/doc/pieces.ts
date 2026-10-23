import { type Bytes, utf16 } from "./bytes"
import { DocError, corrupt } from "./errors"

// [MS-DOC] §2.9.38 Clx: RgPrc (Prc §2.9.209) then Pcdt (§2.9.178) holding PlcPcd (§2.8.35); §2.4.1 Retrieving Text.

export interface Pieces {
	// n + 1 character positions.
	readonly cps: Uint32Array
	// Byte offset of each piece's text in the WordDocument stream (FcCompressed.fc, halved when compressed).
	readonly fcs: Uint32Array
	// 8-bit cp1252 text (FcCompressed.fCompressed) or UTF-16LE.
	readonly compressed: Uint8Array
	// Pcd.prm, raw (§2.9.214).
	readonly prms: Uint16Array
	// Clx.RgPrc: each Prc's grpprl as [start, end) in the Table stream.
	readonly prcs: readonly (readonly [number, number])[]
	readonly count: number
}

// Cap on the text a file may claim: a few times its stream (pieces may reuse text) plus a margin for tiny files.
export function maxTextLength(wordLength: number): number {
	return wordLength * 4 + (1 << 20)
}

export function readPieces(table: Bytes, fcClx: number, lcbClx: number, wordLength: number): Pieces {
	if (lcbClx === 0) throw corrupt("the file has no Clx (piece table)")

	table.check(fcClx, lcbClx, "Clx")

	const end = fcClx + lcbClx
	const prcs: (readonly [number, number])[] = []
	let at = fcClx

	while (at < end) {
		const clxt = table.u8(at)

		if (clxt === 0x01) {
			const cb = table.i16(at + 1)

			if (cb < 0 || at + 3 + cb > end) throw corrupt("Prc grpprl runs past the Clx")

			prcs.push([at + 3, at + 3 + cb])
			at += 3 + cb
			continue
		}

		if (clxt !== 0x02) throw corrupt(`unknown Clx entry 0x${clxt.toString(16)}`)

		const lcb = table.u32(at + 1)
		const plc = at + 5

		// Checked against the Table stream, not lcbClx: some writers understate lcbClx.
		if (lcb < 4 || (lcb - 4) % 12 !== 0 || !table.has(plc, lcb)) throw corrupt("PlcPcd size is invalid")

		return parsePlcPcd(table, plc, lcb, prcs, wordLength)
	}

	throw corrupt("the Clx has no Pcdt")
}

function parsePlcPcd(table: Bytes, at: number, lcb: number, prcs: readonly (readonly [number, number])[], wordLength: number): Pieces {
	const count = (lcb - 4) / 12
	const cps = new Uint32Array(count + 1)
	const fcs = new Uint32Array(count)
	const compressed = new Uint8Array(count)
	const prms = new Uint16Array(count)
	const pcdAt = at + 4 * (count + 1)
	let previous = 0

	for (let k = 0; k <= count; k++) {
		const cp = table.u32(at + 4 * k)

		// aCp is sorted (§2.2.2); a decreasing entry is clamped so later lookups stay well defined.
		cps[k] = cp < previous ? previous : cp
		previous = cps[k] ?? 0
	}

	if ((cps[count] ?? 0) > maxTextLength(wordLength))
		throw new DocError("tooLarge", "the piece table claims more text than the file holds")

	for (let k = 0; k < count; k++) {
		const fcRaw = table.u32(pcdAt + 8 * k + 2)
		const isCompressed = (fcRaw & 0x40000000) !== 0
		const fc = fcRaw & 0x3fffffff

		compressed[k] = isCompressed ? 1 : 0
		fcs[k] = isCompressed ? fc >>> 1 : fc
		prms[k] = table.u16(pcdAt + 8 * k + 6)
	}

	return { cps, fcs, compressed, prms, prcs, count }
}

// FcCompressed (§2.9.73): bytes of compressed text that are not the code point of the same value.
const COMPRESSED_HIGH: Readonly<Record<number, number>> = {
	0x82: 0x201a,
	0x83: 0x0192,
	0x84: 0x201e,
	0x85: 0x2026,
	0x86: 0x2020,
	0x87: 0x2021,
	0x88: 0x02c6,
	0x89: 0x2030,
	0x8a: 0x0160,
	0x8b: 0x2039,
	0x8c: 0x0152,
	0x91: 0x2018,
	0x92: 0x2019,
	0x93: 0x201c,
	0x94: 0x201d,
	0x95: 0x2022,
	0x96: 0x2013,
	0x97: 0x2014,
	0x98: 0x02dc,
	0x99: 0x2122,
	0x9a: 0x0161,
	0x9b: 0x203a,
	0x9c: 0x0153,
	0x9f: 0x0178
}

function compressedText(data: Uint8Array, at: number, count: number): string {
	const parts: string[] = []
	const units = new Array<number>(Math.min(count, 4096))

	for (let done = 0; done < count;) {
		const take = Math.min(4096, count - done)

		units.length = take

		for (let k = 0; k < take; k++) {
			const byte = data[at + done + k] ?? 0

			units[k] = byte >= 0x82 && byte <= 0x9f ? (COMPRESSED_HIGH[byte] ?? byte) : byte
		}

		parts.push(String.fromCharCode(...units))
		done += take
	}

	return parts.join("")
}

export interface PieceText {
	readonly text: string
	// Pieces whose text lies (partly) outside the WordDocument stream; those characters read as U+FFFD.
	readonly damagedPieces: number
}

// The whole document text, every story, indexed by CP.
export function pieceText(word: Bytes, pieces: Pieces): PieceText {
	const parts: string[] = []
	let damagedPieces = 0

	for (let k = 0; k < pieces.count; k++) {
		const count = (pieces.cps[k + 1] ?? 0) - (pieces.cps[k] ?? 0)

		if (count === 0) continue

		const fc = pieces.fcs[k] ?? 0
		const width = pieces.compressed[k] === 1 ? 1 : 2
		const available = Math.max(0, Math.min(count, Math.floor((word.length - fc) / width)))

		if (available > 0) {
			parts.push(width === 1 ? compressedText(word.data, fc, available) : utf16(word.data, fc, available))
		}

		if (available < count) {
			damagedPieces++
			parts.push("\ufffd".repeat(count - available))
		}
	}

	return { text: parts.join(""), damagedPieces }
}

// The piece holding `cp`, or -1.
export function pieceAt(pieces: Pieces, cp: number): number {
	let low = 0
	let high = pieces.count - 1

	if (high < 0 || cp < (pieces.cps[0] ?? 0) || cp >= (pieces.cps[pieces.count] ?? 0)) return -1

	while (low < high) {
		const middle = (low + high + 1) >> 1

		if ((pieces.cps[middle] ?? 0) <= cp) low = middle
		else high = middle - 1
	}

	return low
}

// Prm0 (§2.9.215): isprm → sprm.
export const PRM0_SPRMS: Readonly<Record<number, number>> = {
	0x00: 0x2879,
	0x04: 0x2602,
	0x05: 0x2461,
	0x07: 0x2405,
	0x08: 0x2406,
	0x09: 0x2407,
	0x0c: 0x260a,
	0x0d: 0x2470,
	0x0e: 0x240c,
	0x0f: 0x2471,
	0x18: 0x2416,
	0x19: 0x2417,
	0x1d: 0x261b,
	0x25: 0x2423,
	0x2c: 0x242a,
	0x32: 0x2430,
	0x33: 0x2431,
	0x35: 0x2433,
	0x36: 0x2434,
	0x37: 0x2435,
	0x38: 0x2436,
	0x39: 0x2437,
	0x3a: 0x2438,
	0x41: 0x0800,
	0x42: 0x0801,
	0x43: 0x0802,
	0x47: 0x0806,
	0x4b: 0x080a,
	0x4d: 0x2a0c,
	0x4e: 0x0858,
	0x4f: 0x2859,
	0x50: 0x0811,
	0x51: 0x0818,
	0x53: 0x2a33,
	0x55: 0x0835,
	0x56: 0x0836,
	0x57: 0x0837,
	0x58: 0x0838,
	0x59: 0x0839,
	0x5a: 0x083a,
	0x5b: 0x083b,
	0x5c: 0x083c,
	0x5e: 0x2a3e,
	0x62: 0x2a42,
	0x68: 0x2a48,
	0x73: 0x2a53,
	0x74: 0x0854,
	0x75: 0x0855,
	0x76: 0x0856,
	0x78: 0x2640,
	0x7b: 0x2a90,
	0x7c: 0x2a86,
	0x7e: 0x2443
}

// A piece's own property modifiers as one grpprl (a Prm0 becomes a single synthetic Prl).
export function piecePrm(table: Bytes, pieces: Pieces, piece: number): Uint8Array | undefined {
	const prm = pieces.prms[piece] ?? 0

	if ((prm & 1) === 1) {
		const range = pieces.prcs[prm >>> 1]

		return range === undefined ? undefined : table.data.subarray(range[0], range[1])
	}

	const isprm = (prm >>> 1) & 0x7f
	const val = prm >>> 8
	const sprm = PRM0_SPRMS[isprm]

	// isprm 0 with val 0 has no effect (§2.9.215).
	if (sprm === undefined || (isprm === 0 && val === 0)) return undefined

	return Uint8Array.of(sprm & 0xff, sprm >>> 8, val)
}

export function pcdFcAt(pieces: Pieces, piece: number, cp: number): number {
	const width = pieces.compressed[piece] === 1 ? 1 : 2

	return (pieces.fcs[piece] ?? 0) + width * (cp - (pieces.cps[piece] ?? 0))
}
