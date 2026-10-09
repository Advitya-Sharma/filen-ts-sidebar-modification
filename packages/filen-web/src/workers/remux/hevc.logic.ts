// HEVC access-unit inspection for the remux pump. Firefox's MediaSource starts HEVC only at an IDR
// picture: a CRA key packet (every keyframe but the first in an open-GOP encode) buffers no video at all.

const IDR_W_RADL = 19
const IDR_N_LP = 20
// NAL unit types below this carry slice data (VCL); the rest are parameter sets, SEI, delimiters.
const FIRST_NON_VCL = 32

// The NAL length field size a configuration record (hvcC) declares, null when there is none (Annex B).
export function hevcLengthSize(description: Uint8Array | null): number | null {
	const byte = description?.[21]

	return byte === undefined ? null : (byte & 0x03) + 1
}

function nalType(data: Uint8Array, at: number): number | null {
	const header = data[at]

	return header === undefined ? null : (header >> 1) & 0x3f
}

// The type of the first slice NAL unit in a length-prefixed access unit.
function firstSliceLengthPrefixed(data: Uint8Array, lengthSize: number): number | null {
	let at = 0

	while (at + lengthSize <= data.length) {
		let length = 0

		for (let index = 0; index < lengthSize; index++) {
			length = length * 256 + (data[at + index] ?? 0)
		}

		at += lengthSize

		const type = nalType(data, at)

		if (type === null || length === 0) {
			return null
		}

		if (type < FIRST_NON_VCL) {
			return type
		}

		at += length
	}

	return null
}

// The type of the first slice NAL unit in an Annex B access unit (start-code delimited).
function firstSliceAnnexB(data: Uint8Array): number | null {
	for (let at = 0; at + 3 < data.length; at++) {
		if (data[at] !== 0 || data[at + 1] !== 0 || data[at + 2] !== 1) {
			continue
		}

		const type = nalType(data, at + 3)

		if (type !== null && type < FIRST_NON_VCL) {
			return type
		}

		at += 2
	}

	return null
}

// Whether an HEVC key packet is an IDR picture. An access unit whose slices cannot be found counts as
// one: nothing is gained by refusing a start the decoder may well take.
export function hevcKeyIsIdr(data: Uint8Array, lengthSize: number | null): boolean {
	const type = lengthSize === null ? firstSliceAnnexB(data) : firstSliceLengthPrefixed(data, lengthSize)

	return type === null || type === IDR_W_RADL || type === IDR_N_LP
}

// How far before a requested start an IDR may lie: the pump reads and the page buffers everything from
// it, so the reach is capped in seconds and in bytes at the file's average bitrate.
export const IDR_REACH_SECONDS = 30
export const IDR_REACH_BYTES = 48 * 1_048_576
// Key packets looked at before giving up, each one read where it sits.
export const IDR_REACH_KEYS = 16

export function idrReachSeconds(bitrate: number): number {
	return bitrate > 0 ? Math.min(IDR_REACH_SECONDS, (IDR_REACH_BYTES * 8) / bitrate) : IDR_REACH_SECONDS
}
