import { UnsupportedContainerError, ascii, uint16, uint32 } from "./probeBytes"

// The head of an MPEG transport stream (ISO/IEC 13818-1): its packet layout and its PAT, which decide
// whether the container probe reads the file as one. A leaf of the probe (probeMpegts.ts walks the rest),
// light enough for a runtime that only needs this decision.

// The EBML header's ID, which every Matroska/WebM file starts with.
export const EBML_ID = 0x1a45dfa3

// Box types an ISO BMFF file may open with: `ftyp` in every MP4, the others in QuickTime files that predate it.
export const LEADING_BOXES = new Set(["ftyp", "moov", "mdat", "wide", "free", "skip", "pnot"])

export const TS_SYNC = 0x47
export const TS_PACKET_BYTES = 188

// Packet k's sync byte sits at k × stride + prefix: BDAV (.m2ts) puts a 4-byte timestamp before each
// packet, and 204-byte packets carry 16 bytes of parity after it.
export interface TsLayout {
	stride: number
	prefix: number
}

const LAYOUTS: readonly TsLayout[] = [
	{ stride: 188, prefix: 0 },
	{ stride: 192, prefix: 4 },
	{ stride: 204, prefix: 0 }
]

// TypeScript sources share the extension: only a sync byte on every one of the first packets counts.
export const TS_SNIFF_PACKETS = 5
export const TS_MIN_PACKETS = 3

export const TS_SNIFF_BYTES = Math.max(...LAYOUTS.map(layout => layout.prefix + layout.stride * (TS_SNIFF_PACKETS - 1) + 1))

export function packetCount(size: number, layout: TsLayout): number {
	const room = size - layout.prefix - TS_PACKET_BYTES

	return room < 0 ? 0 : Math.floor(room / layout.stride) + 1
}

// The packet layout of a file whose first TS_SNIFF_BYTES are `head`, or null.
export function sniffTransportStream(head: Uint8Array, size: number): TsLayout | null {
	for (const layout of LAYOUTS) {
		const checks = Math.min(TS_SNIFF_PACKETS, packetCount(size, layout))
		let synced = checks >= TS_MIN_PACKETS

		for (let k = 0; k < checks && synced; k++) {
			synced = head[layout.prefix + k * layout.stride] === TS_SYNC
		}

		if (synced) {
			return layout
		}
	}

	return null
}

const MAX_SECTIONS = 8
const MAX_SECTION_BYTES = 1024

// The PSI state a walk keeps: the sections being assembled per PID, and what the PAT said.
export interface PsiWalk {
	sections: Map<number, { bytes: Uint8Array; length: number }>
	patFound: boolean
	pmtPid: number | null
	malformed: boolean
}

// Assembles one PSI section per PID across packets; returns it once complete, as a view valid until the next call.
export function feedSection(walk: PsiWalk, pid: number, payload: Uint8Array, unitStart: boolean): Uint8Array | null {
	let section = walk.sections.get(pid)

	if (unitStart) {
		const from = 1 + (payload[0] ?? 0)

		if (from >= payload.length) {
			return null
		}

		if (section === undefined) {
			if (walk.sections.size >= MAX_SECTIONS) {
				return null
			}

			section = { bytes: new Uint8Array(MAX_SECTION_BYTES), length: 0 }
			walk.sections.set(pid, section)
		}

		section.length = 0
		payload = payload.subarray(from)
	} else if (section === undefined || section.length === 0) {
		return null
	}

	const take = Math.min(payload.length, MAX_SECTION_BYTES - section.length)

	section.bytes.set(payload.subarray(0, take), section.length)
	section.length += take

	if (section.length < 3) {
		return null
	}

	const total = 3 + (uint16(section.bytes, 1) & 0x0fff)

	if (total > MAX_SECTION_BYTES || total < 12) {
		section.length = 0
		walk.malformed = true

		return null
	}

	if (section.length < total) {
		return null
	}

	section.length = 0

	return section.bytes.subarray(0, total)
}

// The remux's demuxer reads a stream with exactly one program (program 0 names the network PID, not one).
export function readPat(walk: PsiWalk, section: Uint8Array): void {
	const end = section.length - 4
	let programs = 0

	for (let at = 8; at + 4 <= end; at += 4) {
		if (uint16(section, at) !== 0) {
			walk.pmtPid = uint16(section, at + 2) & 0x1fff
			programs++
		}
	}

	if (programs !== 1) {
		throw new UnsupportedContainerError()
	}

	walk.patFound = true
}

export interface TsPacket {
	pid: number
	unitStart: boolean
	payload: Uint8Array
}

// The packet whose sync byte is at `at`, or null for one without a usable payload (lost sync, a
// transport error, scrambled, adaptation field only).
export function packetAt(bytes: Uint8Array, at: number): TsPacket | null {
	const flags = bytes[at + 1] ?? 0
	const control = bytes[at + 3] ?? 0

	if (bytes[at] !== TS_SYNC || (flags & 0x80) !== 0 || (control & 0xc0) !== 0 || (control & 0x10) === 0) {
		return null
	}

	const start = at + 4 + ((control & 0x20) !== 0 ? 1 + (bytes[at + 4] ?? 0) : 0)
	const end = at + TS_PACKET_BYTES

	if (start >= end || end > bytes.length) {
		return null
	}

	return { pid: ((flags & 0x1f) << 8) | (bytes[at + 2] ?? 0), unitStart: (flags & 0x40) !== 0, payload: bytes.subarray(start, end) }
}

// The packet layout of a file of `size` bytes whose first bytes are `head` (TS_SNIFF_BYTES, or the whole file)
// when it is a transport stream the probe reads, decided as walkContainer decides it from what the head holds:
// no Matroska or ISO BMFF header, a sync byte on each of the first packets, and a PAT, if one of the head's
// packets carries it, naming exactly one program. Null otherwise. Reads nothing past `head`.
export function transportStreamOfHead(head: Uint8Array, size: number): TsLayout | null {
	if (size < 8 || uint32(head, 0) === EBML_ID || LEADING_BOXES.has(ascii(head, 4, 8)) || (head[0] !== TS_SYNC && head[4] !== TS_SYNC)) {
		return null
	}

	const layout = sniffTransportStream(head, size)

	if (layout === null) {
		return null
	}

	const walk: PsiWalk = { sections: new Map(), patFound: false, pmtPid: null, malformed: false }

	try {
		for (let at = layout.prefix; at + TS_PACKET_BYTES <= head.length && !walk.patFound; at += layout.stride) {
			const packet = packetAt(head, at)

			if (packet?.pid === 0) {
				const section = feedSection(walk, 0, packet.payload, packet.unitStart)

				if (section?.[0] === 0x00) {
					readPat(walk, section)
				}
			}
		}
	} catch (error) {
		if (error instanceof UnsupportedContainerError) {
			return null
		}

		throw error
	}

	return layout
}
