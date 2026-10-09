// Synthetic MPEG transport streams: PSI sections, PES packets split into 188-byte packets, optionally
// in 192-byte BDAV units. CRCs are zero (the probe does not check them).

export function hexBytes(value: string): number[] {
	const bytes: number[] = []

	for (let i = 0; i + 1 < value.length; i += 2) {
		bytes.push(parseInt(value.slice(i, i + 2), 16))
	}

	return bytes
}

// x264's SPS for 1280×720 High@3.1 and x265's for 1280×720 Main 10@3.1.
export const AVC_SPS = hexBytes("6764001facd9405005bb011000000300100000030300f1831960")
export const HEVC_SPS = hexBytes("42010102200000030090000003000003005da00280802d136595952930bc05a02000000300200000030301")

// One TS packet; a payload shorter than 184 bytes is padded with adaptation-field stuffing.
export function packet(pid: number, payload: readonly number[], unitStart: boolean, counter = 0): number[] {
	const header = [0x47, (unitStart ? 0x40 : 0) | ((pid >> 8) & 0x1f), pid & 0xff]

	if (payload.length >= 184) {
		return [...header, 0x10 | (counter & 0x0f), ...payload.slice(0, 184)]
	}

	const stuffing = 183 - payload.length
	const adaptation = stuffing === 0 ? [0] : [stuffing, 0x00, ...new Array<number>(stuffing - 1).fill(0xff)]

	return [...header, 0x30 | (counter & 0x0f), ...adaptation, ...payload]
}

function section(tableId: number, body: readonly number[]): number[] {
	const length = body.length + 4

	return [tableId, 0xb0 | (length >> 8), length & 0xff, ...body, 0, 0, 0, 0]
}

// A PAT listing each [program number, PMT PID]; program 0 names the network PID.
export function patOf(programs: readonly (readonly [program: number, pid: number])[]): number[] {
	const entries = programs.flatMap(([program, pid]) => [program >> 8, program & 0xff, 0xe0 | (pid >> 8), pid & 0xff])

	return packet(0, [0, ...section(0x00, [0x00, 0x01, 0xc1, 0x00, 0x00, ...entries])], true)
}

export function pat(pmtPid: number): number[] {
	return patOf([[1, pmtPid]])
}

export interface EsEntry {
	type: number
	pid: number
	descriptors?: number[]
}

export function pmtSection(streams: readonly EsEntry[], programDescriptors: readonly number[] = []): number[] {
	const body = [
		0x00,
		0x01,
		0xc1,
		0x00,
		0x00,
		0xe1,
		0x00,
		0xf0 | (programDescriptors.length >> 8),
		programDescriptors.length & 0xff,
		...programDescriptors
	]

	for (const stream of streams) {
		const descriptors = stream.descriptors ?? []

		body.push(
			stream.type,
			0xe0 | (stream.pid >> 8),
			stream.pid & 0xff,
			0xf0 | (descriptors.length >> 8),
			descriptors.length & 0xff,
			...descriptors
		)
	}

	return section(0x02, body)
}

export function pmt(pmtPid: number, streams: readonly EsEntry[], programDescriptors: readonly number[] = []): number[] {
	return packet(pmtPid, [0, ...pmtSection(streams, programDescriptors)], true)
}

export function language(code: string): number[] {
	return [0x0a, 4, ...new TextEncoder().encode(code), 0]
}

export function registration(id: string): number[] {
	return [0x05, 4, ...new TextEncoder().encode(id)]
}

function ptsBytes(pts: number): number[] {
	const high = Math.floor(pts / 2 ** 30)
	const low = pts % 2 ** 30

	return [0x21 | ((high & 0x07) << 1), (low >>> 22) & 0xff, ((low >>> 14) & 0xfe) | 1, (low >>> 7) & 0xff, ((low << 1) & 0xfe) | 1]
}

// A PES packet with a PTS, split into TS packets.
export function pes(pid: number, streamId: number, pts: number, data: readonly number[]): number[] {
	const bytes = [0, 0, 1, streamId, 0, 0, 0x80, 0x80, 5, ...ptsBytes(pts), ...data]
	const packets: number[] = []

	for (let at = 0, counter = 0; at < bytes.length; at += 184, counter++) {
		packets.push(...packet(pid, bytes.slice(at, at + 184), at === 0, counter))
	}

	return packets
}

// An access unit delimiter, the SPS, then slice-like filler.
export function annexB(sps: readonly number[], filler = 64): number[] {
	return [0, 0, 0, 1, 0x09, 0xf0, 0, 0, 0, 1, ...sps, 0, 0, 0, 1, 0x65, ...new Array<number>(filler).fill(0x88)]
}

// ADTS header: AAC LC (profile 1), 48 kHz (index 3), stereo.
export const ADTS_LC_48K_STEREO = [0xff, 0xf1, 0x4c, 0x80, 0x01, 0x7f, 0xfc]

// AC-3 sync frame start: 48 kHz, bsid 8, acmod 7 (3/2) with LFE.
export const AC3_51 = [0x0b, 0x77, 0x00, 0x00, 0x1c, 0x40, 0xe1, 0x00]

// Null packets: the bulk of a synthetic stream.
export function nulls(count: number): number[] {
	const bytes: number[] = []

	for (let i = 0; i < count; i++) {
		bytes.push(...packet(0x1fff, new Array<number>(184).fill(0xff), false))
	}

	return bytes
}

// The same packets in 192-byte BDAV units (a 4-byte arrival timestamp before each).
export function bdav(bytes: readonly number[]): number[] {
	const units: number[] = []

	for (let at = 0; at < bytes.length; at += 188) {
		units.push(0x00, 0x00, 0x00, 0x00, ...bytes.slice(at, at + 188))
	}

	return units
}
