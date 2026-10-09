import {
	avcConfigCodecString,
	hevcConfigCodecString,
	mp4aCodecString,
	parseAudioSpecificConfig,
	parseAvcSps,
	parseHevcSps,
	simpleAudioCodecString,
	type SpsPicture
} from "@filen/shared/videos"
import { BLOCK_BYTES, type BlockSource } from "@/lib/media/blockSource"
import { UnsupportedContainerError, ascii, readBytes, seconds, transferKind, uint16 } from "@/lib/media/probe/bytes"
import { AC3_CHANNELS, AC3_RATES } from "@/lib/media/probe/isobmff"
import type { MediaCodec, ParsedContainer, ParsedTrack, ProbeTier } from "@/lib/media/probe/types"

// MPEG transport streams (ISO/IEC 13818-1) of a single program: its PMT, each stream's first parameter set or
// frame header from the head, and the timing stream's PTS span from the head and the tail. The stream
// mapping follows the remux's demuxer (Mediabunny), so a track reported here is one it can read.

const SYNC = 0x47
const PACKET_BYTES = 188

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
const SNIFF_PACKETS = 5
const MIN_PACKETS = 3

export const TS_SNIFF_BYTES = Math.max(...LAYOUTS.map(layout => layout.prefix + layout.stride * (SNIFF_PACKETS - 1) + 1))

// A card stays within CARD_BLOCK_READS: two head blocks, and at most two for a tail straddling a boundary.
const HEAD_BYTES: Readonly<Record<ProbeTier, number>> = { card: 2 * BLOCK_BYTES, cover: 2 * BLOCK_BYTES, full: 4 * BLOCK_BYTES }
const TAIL_BYTES = BLOCK_BYTES
const FULL_BLOCK_READS = 6

const CHUNK_PACKETS = 32
const MAX_STREAMS = 32
const MAX_SECTIONS = 8
const MAX_SECTION_BYTES = 1024
// Enough for the SEI an encoder writes ahead of its parameter sets.
const VIDEO_PREFIX_BYTES = 16_384
const AUDIO_PREFIX_BYTES = 256
// PES starts read from the tail: B-frames put the latest PTS a few pictures before the last one.
const TAIL_PES = 4
const PTS_MODULUS = 2 ** 33
const PTS_HZ = 90_000

function packetCount(size: number, layout: TsLayout): number {
	const room = size - layout.prefix - PACKET_BYTES

	return room < 0 ? 0 : Math.floor(room / layout.stride) + 1
}

// The packet layout of a file whose first TS_SNIFF_BYTES are `head`, or null.
export function sniffTransportStream(head: Uint8Array, size: number): TsLayout | null {
	for (const layout of LAYOUTS) {
		const checks = Math.min(SNIFF_PACKETS, packetCount(size, layout))
		let synced = checks >= MIN_PACKETS

		for (let k = 0; k < checks && synced; k++) {
			synced = head[layout.prefix + k * layout.stride] === SYNC
		}

		if (synced) {
			return layout
		}
	}

	return null
}

type Need = "avc" | "hevc" | "adts" | "dolby" | "mpeg-audio" | "dts" | null

interface StreamKind {
	kind: "video" | "audio"
	codec: MediaCodec | null
	fourcc: string
	codecString: string | null
	need: Need
}

interface Stream extends StreamKind {
	pid: number
	lang: string | null
	picture: SpsPicture | null
	ch: number | null
	rate: number | null
	// The current PES's payload while `need` is unmet.
	buffer: Uint8Array | null
	length: number
	collecting: boolean
	// Video: where the SPS starts in `buffer` (-1 not yet), and how far start codes were searched.
	spsAt: number
	searched: number
}

interface Walk {
	sections: Map<number, { bytes: Uint8Array; length: number }>
	patFound: boolean
	pmtPid: number | null
	pmtFound: boolean
	malformed: boolean
	streams: Stream[]
	byPid: Map<number, Stream>
	timingPid: number | null
	firstPts: number | null
}

function video(codec: MediaCodec | null, fourcc: string, need: Need = null): StreamKind {
	return { kind: "video", codec, fourcc, codecString: null, need }
}

function audio(codec: MediaCodec | null, codecString: string | null, need: Need): StreamKind {
	return { kind: "audio", codec, fourcc: "", codecString, need }
}

interface Descriptors {
	lang: string | null
	ac3: boolean
	eac3: boolean
	dts: boolean
}

// ISO 639-2 from a language descriptor; "und" and anything not three letters is no language.
function language(bytes: Uint8Array, at: number): string | null {
	const code = ascii(bytes, at, at + 3).toLowerCase()

	return /^[a-z]{3}$/.test(code) && code !== "und" ? code : null
}

function descriptors(bytes: Uint8Array, start: number, end: number): Descriptors {
	const found: Descriptors = { lang: null, ac3: false, eac3: false, dts: false }

	for (let at = start; at + 2 <= end;) {
		const tag = bytes[at] ?? 0
		const body = at + 2
		const next = body + (bytes[at + 1] ?? 0)

		if (next > end) {
			break
		}

		if (tag === 0x0a && next - body >= 4) {
			found.lang ??= language(bytes, body)
		} else if (tag === 0x6a) {
			found.ac3 = true
		} else if (tag === 0x7a || tag === 0xcc) {
			found.eac3 = true
		} else if (tag === 0x7b || (tag === 0x05 && next - body >= 4 && /^DTS[123]$/.test(ascii(bytes, body, body + 4)))) {
			found.dts = true
		}

		at = next
	}

	return found
}

// A program registered as Blu-ray ("HDMV", "HDPR"), whose private stream types mean what BDAV says.
function bluRayProgram(bytes: Uint8Array, start: number, end: number): boolean {
	for (let at = start; at + 2 <= end;) {
		const body = at + 2
		const next = body + (bytes[at + 1] ?? 0)

		if (next > end) {
			break
		}

		if (bytes[at] === 0x05 && next - body >= 4) {
			const id = ascii(bytes, body, body + 4)

			if (id === "HDMV" || id === "HDPR") {
				return true
			}
		}

		at = next
	}

	return false
}

function classify(type: number, found: Descriptors, bluRay: boolean): StreamKind | null {
	switch (type) {
		case 0x01:
			return video(null, "mp1v")
		case 0x02:
			return video(null, "mp2v")
		case 0x10:
			return video("mpeg4", "mp4v")
		case 0x1b:
			return video("h264", "avc1", "avc")
		case 0x24:
			return video("hevc", "hvc1", "hevc")
		case 0xea:
			return bluRay ? video(null, "vc-1") : null
		// MPEG-1/2 audio: only layer III, read from the first frame, is MP3.
		case 0x03:
		case 0x04:
			return audio(null, null, "mpeg-audio")
		case 0x0f:
			return audio("aac", null, "adts")
		// AAC in LATM: a track the remux's demuxer does not read.
		case 0x11:
			return audio(null, null, null)
		case 0x81:
			return audio("ac3", simpleAudioCodecString("ac3"), "dolby")
		case 0x87:
			return audio("eac3", simpleAudioCodecString("eac3"), "dolby")
		case 0x82:
		case 0x8a:
			return audio("dts", "dtsc", "dts")
		case 0x85:
			return bluRay ? audio("dts", "dtsh", "dts") : null
		case 0x86:
			return bluRay ? audio("dts", "dtsl", "dts") : null
		case 0xa2:
			return bluRay ? audio("dts", "dtse", "dts") : null
		// LPCM, TrueHD and E-AC-3 in BDAV: tracks the remux's demuxer does not read.
		case 0x80:
		case 0x84:
		case 0xa1:
			return bluRay ? audio(null, null, null) : null
		case 0x83:
			return bluRay ? audio("truehd", simpleAudioCodecString("truehd"), null) : null
		case 0x06:
			return found.eac3
				? audio("eac3", simpleAudioCodecString("eac3"), "dolby")
				: found.ac3
					? audio("ac3", simpleAudioCodecString("ac3"), "dolby")
					: found.dts
						? audio("dts", "dtsc", "dts")
						: null
		default:
			return null
	}
}

// Assembles one PSI section per PID across packets; returns it once complete, as a view valid until the next call.
function feedSection(walk: Walk, pid: number, payload: Uint8Array, unitStart: boolean): Uint8Array | null {
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
function readPat(walk: Walk, section: Uint8Array): void {
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

function readPmt(walk: Walk, section: Uint8Array): void {
	if (section.length < 16) {
		walk.malformed = true

		return
	}

	const end = section.length - 4
	const infoEnd = 12 + (uint16(section, 10) & 0x0fff)

	if (infoEnd > end) {
		walk.malformed = true

		return
	}

	const bluRay = bluRayProgram(section, 12, infoEnd)

	for (let at = infoEnd; at + 5 <= end && walk.streams.length < MAX_STREAMS;) {
		const pid = uint16(section, at + 1) & 0x1fff
		const next = at + 5 + (uint16(section, at + 3) & 0x0fff)

		if (next > end) {
			walk.malformed = true

			break
		}

		const found = descriptors(section, at + 5, next)
		const kind = classify(section[at] ?? 0, found, bluRay)

		if (kind !== null && !walk.byPid.has(pid)) {
			const stream: Stream = {
				...kind,
				pid,
				lang: kind.kind === "audio" ? found.lang : null,
				picture: null,
				ch: null,
				rate: null,
				buffer: null,
				length: 0,
				collecting: false,
				spsAt: -1,
				searched: 0
			}

			walk.streams.push(stream)
			walk.byPid.set(pid, stream)
		}

		at = next
	}

	walk.pmtFound = true
	walk.timingPid = (walk.streams.find(stream => stream.kind === "video") ?? walk.streams[0])?.pid ?? null
}

const NO_PES_HEADER = new Set([0xbc, 0xbe, 0xbf, 0xf0, 0xf1, 0xf2, 0xf8, 0xff])

function timestamp(bytes: Uint8Array, at: number): number {
	return (
		(((bytes[at] ?? 0) >> 1) & 0x07) * 2 ** 30 +
		((bytes[at + 1] ?? 0) << 22) +
		(((bytes[at + 2] ?? 0) >> 1) << 15) +
		((bytes[at + 3] ?? 0) << 7) +
		((bytes[at + 4] ?? 0) >> 1)
	)
}

// A PES packet starting in one TS payload: its PTS and where its data begins in the payload; null when
// the payload starts no PES or its header runs past the packet.
function pesStart(payload: Uint8Array): { pts: number | null; data: number } | null {
	if (payload.length < 9 || payload[0] !== 0 || payload[1] !== 0 || payload[2] !== 1) {
		return null
	}

	if (NO_PES_HEADER.has(payload[3] ?? 0)) {
		return { pts: null, data: 6 }
	}

	if (((payload[6] ?? 0) & 0xc0) !== 0x80) {
		return null
	}

	const data = 9 + (payload[8] ?? 0)

	if (data > payload.length) {
		return null
	}

	return { pts: ((payload[7] ?? 0) & 0x80) !== 0 && data >= 14 ? timestamp(payload, 9) : null, data }
}

function readBits(bytes: Uint8Array, position: number, count: number): number {
	let value = 0

	for (let i = position; i < position + count; i++) {
		value = value * 2 + (((bytes[i >> 3] ?? 0) >> (7 - (i & 7))) & 1)
	}

	return value
}

interface AudioFormat {
	ch: number | null
	rate: number | null
	codec?: MediaCodec | null
	codecString?: string | null
}

// ADTS: the AudioSpecificConfig the header implies (object type = profile + 1).
function adtsFormat(bytes: Uint8Array, at: number): AudioFormat | null {
	if (bytes[at] !== 0xff || ((bytes[at + 1] ?? 0) & 0xf6) !== 0xf0) {
		return null
	}

	const packed = bytes[at + 2] ?? 0
	const objectType = (packed >> 6) + 1
	const rateIndex = (packed >> 2) & 0x0f
	const channels = ((packed & 1) << 2) | ((bytes[at + 3] ?? 0) >> 6)

	if (rateIndex > 12) {
		return null
	}

	const config = Uint8Array.of((objectType << 3) | (rateIndex >> 1), ((rateIndex & 1) << 7) | (channels << 3))
	const parsed = parseAudioSpecificConfig(config)

	return {
		ch: parsed?.channels ?? null,
		rate: parsed?.sampleRate ?? null,
		codecString: mp4aCodecString({ objectTypeIndication: 0x40, audioSpecificConfig: config })
	}
}

// AC-3 (bsid ≤ 10) and E-AC-3 (11-16) sync frames: the main (or first independent) stream's channels.
function dolbyFormat(bytes: Uint8Array, at: number): AudioFormat | null {
	if (bytes[at] !== 0x0b || bytes[at + 1] !== 0x77) {
		return null
	}

	const bsid = (bytes[at + 5] ?? 0) >> 3

	if (bsid <= 10) {
		const fscod = (bytes[at + 4] ?? 0) >> 6
		const acmod = (bytes[at + 6] ?? 0) >> 5
		let position = (at + 6) * 8 + 3

		position += (acmod & 1) !== 0 && acmod !== 1 ? 2 : 0
		position += (acmod & 4) !== 0 ? 2 : 0
		position += acmod === 2 ? 2 : 0

		return fscod === 3 ? null : { ch: (AC3_CHANNELS[acmod] ?? 2) + readBits(bytes, position, 1), rate: AC3_RATES[fscod] ?? null }
	}

	const packed = bytes[at + 4] ?? 0
	const fscod = packed >> 6
	const rate = fscod === 3 ? [24000, 22050, 16000][(packed >> 4) & 0x03] : AC3_RATES[fscod]

	if (bsid > 16 || (bytes[at + 2] ?? 0) >> 6 === 1 || rate === undefined) {
		return null
	}

	return { ch: (AC3_CHANNELS[(packed >> 1) & 0x07] ?? 2) + (packed & 1), rate }
}

const MPEG_AUDIO_RATES = [44100, 48000, 32000]

// MPEG-1/2/2.5 audio frame header; layer III is MP3, layers I and II no codec this names.
function mpegAudioFormat(bytes: Uint8Array, at: number): AudioFormat | null {
	const first = bytes[at + 1] ?? 0
	const second = bytes[at + 2] ?? 0
	const version = (first >> 3) & 0x03
	const layer = (first >> 1) & 0x03
	const base = MPEG_AUDIO_RATES[(second >> 2) & 0x03]

	if (bytes[at] !== 0xff || (first & 0xe0) !== 0xe0 || version === 1 || layer === 0 || second >> 4 === 0x0f || base === undefined) {
		return null
	}

	const mp3 = layer === 1

	return {
		ch: (bytes[at + 3] ?? 0) >> 6 === 3 ? 1 : 2,
		rate: version === 3 ? base : version === 2 ? base / 2 : base / 4,
		codec: mp3 ? "mp3" : null,
		codecString: mp3 ? simpleAudioCodecString("mp3") : null
	}
}

const DTS_CHANNELS = [1, 2, 2, 2, 2, 3, 3, 4, 4, 5, 6, 6, 6, 7, 8, 8]
const DTS_RATES = [0, 8000, 16000, 32000, 0, 0, 11025, 22050, 44100, 0, 0, 12000, 24000, 48000, 0, 0]

// DTS core frame header: AMODE at bit 60, SFREQ at 66, LFF at 85.
function dtsFormat(bytes: Uint8Array, at: number): AudioFormat | null {
	if (bytes[at] !== 0x7f || bytes[at + 1] !== 0xfe || bytes[at + 2] !== 0x80 || bytes[at + 3] !== 0x01) {
		return null
	}

	const base = at * 8
	const rate = DTS_RATES[readBits(bytes, base + 66, 4)] ?? 0
	const channels = DTS_CHANNELS[readBits(bytes, base + 60, 6)]
	const lfe = readBits(bytes, base + 85, 2)

	return rate === 0 ? null : { ch: channels === undefined ? null : channels + (lfe === 1 || lfe === 2 ? 1 : 0), rate }
}

const AUDIO_HEADERS: Readonly<Record<"adts" | "dolby" | "mpeg-audio" | "dts", { bytes: number; read: typeof adtsFormat }>> = {
	adts: { bytes: 7, read: adtsFormat },
	dolby: { bytes: 8, read: dolbyFormat },
	"mpeg-audio": { bytes: 4, read: mpegAudioFormat },
	dts: { bytes: 11, read: dtsFormat }
}

function resolve(stream: Stream): void {
	stream.need = null
	stream.buffer = null
	stream.collecting = false
}

function inspectAudio(stream: Stream, header: (typeof AUDIO_HEADERS)[keyof typeof AUDIO_HEADERS], complete: boolean): void {
	const { buffer, length } = stream

	if (buffer === null) {
		return
	}

	for (let at = 0; at + header.bytes <= length; at++) {
		const format = header.read(buffer, at)

		if (format !== null) {
			stream.ch = format.ch
			stream.rate = format.rate

			if (format.codec !== undefined) {
				stream.codec = format.codec
			}

			if (format.codecString !== undefined) {
				stream.codecString = format.codecString
			}

			resolve(stream)

			return
		}
	}

	if (complete || length >= buffer.length) {
		stream.collecting = false
	}
}

function isSps(stream: Stream, header: number): boolean {
	return stream.need === "avc" ? (header & 0x1f) === 7 : ((header >> 1) & 0x3f) === 33
}

// Finds the SPS in the PES payload gathered so far, resuming where the last search stopped; it ends at
// the next start code (00 00 00 or 00 00 01, which emulation prevention keeps out of a NAL unit).
function inspectVideo(stream: Stream, complete: boolean): void {
	const { buffer, length } = stream

	if (buffer === null) {
		return
	}

	if (stream.spsAt < 0) {
		for (let at = stream.searched; at + 3 < length; at++) {
			if (buffer[at] === 0 && buffer[at + 1] === 0 && buffer[at + 2] === 1 && isSps(stream, buffer[at + 3] ?? 0)) {
				stream.spsAt = at + 3

				break
			}
		}

		if (stream.spsAt < 0) {
			stream.searched = Math.max(stream.searched, length - 3)

			if (complete || length >= buffer.length) {
				stream.collecting = false
			}

			return
		}

		stream.searched = stream.spsAt + 1
	}

	let end = -1

	for (let at = stream.searched; at + 2 < length; at++) {
		if (buffer[at] === 0 && buffer[at + 1] === 0 && (buffer[at + 2] ?? 0) <= 1) {
			end = at

			break
		}
	}

	if (end < 0) {
		stream.searched = Math.max(stream.searched, length - 2)

		if (!complete && length < buffer.length) {
			return
		}

		end = length
	}

	const nal = buffer.subarray(stream.spsAt, end)

	if (stream.need === "avc") {
		const sps = parseAvcSps(nal)

		stream.picture = sps
		stream.codecString = sps === null ? null : avcConfigCodecString(stream.fourcc, sps)
	} else {
		const sps = parseHevcSps(nal)

		stream.picture = sps
		stream.codecString = sps === null ? null : hevcConfigCodecString(stream.fourcc, sps)
	}

	if (stream.picture === null) {
		stream.collecting = false

		return
	}

	resolve(stream)
}

function inspect(stream: Stream, complete: boolean): void {
	if (stream.need === "avc" || stream.need === "hevc") {
		inspectVideo(stream, complete)
	} else if (stream.need !== null) {
		inspectAudio(stream, AUDIO_HEADERS[stream.need], complete)
	}
}

function feedStream(walk: Walk, stream: Stream, payload: Uint8Array, unitStart: boolean): void {
	if (unitStart) {
		if (stream.collecting) {
			inspect(stream, true)
		}

		stream.collecting = false

		const pes = pesStart(payload)

		if (pes === null) {
			return
		}

		if (stream.pid === walk.timingPid && pes.pts !== null && walk.firstPts === null) {
			walk.firstPts = pes.pts
		}

		if (stream.need === null) {
			return
		}

		stream.buffer ??= new Uint8Array(stream.kind === "video" ? VIDEO_PREFIX_BYTES : AUDIO_PREFIX_BYTES)
		stream.collecting = true
		stream.length = 0
		stream.spsAt = -1
		stream.searched = 0
		payload = payload.subarray(pes.data)
	} else if (!stream.collecting) {
		return
	}

	const buffer = stream.buffer

	if (buffer === null) {
		return
	}

	const take = Math.min(payload.length, buffer.length - stream.length)

	buffer.set(payload.subarray(0, take), stream.length)
	stream.length += take
	inspect(stream, false)
}

interface Packet {
	pid: number
	unitStart: boolean
	payload: Uint8Array
}

// The packet whose sync byte is at `at`, or null for one without a usable payload (lost sync, a
// transport error, scrambled, adaptation field only).
function packetAt(bytes: Uint8Array, at: number): Packet | null {
	const flags = bytes[at + 1] ?? 0
	const control = bytes[at + 3] ?? 0

	if (bytes[at] !== SYNC || (flags & 0x80) !== 0 || (control & 0xc0) !== 0 || (control & 0x10) === 0) {
		return null
	}

	const start = at + 4 + ((control & 0x20) !== 0 ? 1 + (bytes[at + 4] ?? 0) : 0)
	const end = at + PACKET_BYTES

	if (start >= end || end > bytes.length) {
		return null
	}

	return { pid: ((flags & 0x1f) << 8) | (bytes[at + 2] ?? 0), unitStart: (flags & 0x40) !== 0, payload: bytes.subarray(start, end) }
}

// A PMT on a PID no PAT named (some streams carry none): a PSI payload whose first table is a PMT.
function guessedPmt(walk: Walk, packet: Packet): boolean {
	const { pid, payload } = packet

	if (pid === 0 || pid === 0x1fff) {
		return false
	}

	if (!packet.unitStart) {
		return walk.sections.has(pid)
	}

	return !(payload[0] === 0 && payload[1] === 0 && payload[2] === 1) && payload[1 + (payload[0] ?? 0)] === 0x02
}

function handlePacket(walk: Walk, bytes: Uint8Array, at: number): void {
	const packet = packetAt(bytes, at)

	if (packet === null) {
		return
	}

	const { pid, payload, unitStart } = packet

	if (walk.pmtFound) {
		const stream = walk.byPid.get(pid)

		if (stream !== undefined) {
			feedStream(walk, stream, payload, unitStart)
		}

		return
	}

	if (pid === 0 && !walk.patFound) {
		const section = feedSection(walk, pid, payload, unitStart)

		if (section?.[0] === 0x00) {
			readPat(walk, section)
		}
	} else if (pid === walk.pmtPid || (walk.pmtPid === null && guessedPmt(walk, packet))) {
		const section = feedSection(walk, pid, payload, unitStart)

		if (section?.[0] === 0x02) {
			readPmt(walk, section)
		}
	}
}

function settled(walk: Walk): boolean {
	return walk.pmtFound && walk.streams.every(stream => stream.need === null) && (walk.timingPid === null || walk.firstPts !== null)
}

async function scanHead(source: BlockSource, walk: Walk, layout: TsLayout, limit: number): Promise<void> {
	const packets = packetCount(limit, layout)

	for (let first = 0; first < packets && !settled(walk); first += CHUNK_PACKETS) {
		const count = Math.min(CHUNK_PACKETS, packets - first)
		const bytes = await readBytes(source, first * layout.stride + layout.prefix, (count - 1) * layout.stride + PACKET_BYTES)

		for (let k = 0; k < count; k++) {
			handlePacket(walk, bytes, k * layout.stride)
		}
	}

	for (const stream of walk.streams) {
		if (stream.collecting) {
			inspect(stream, true)
		}
	}
}

function modulo(value: number, divisor: number): number {
	return ((value % divisor) + divisor) % divisor
}

// The residue (mod stride) of the sync byte positions in the tail chunk at `base`: the file's own grid,
// or the one its tail moved to (a capture cut and joined), or null when none lines up.
function tailGrid(bytes: Uint8Array, base: number, stride: number, expected: number): number | null {
	const fits = (residue: number): boolean => {
		let hits = 0

		for (let at = modulo(residue - base, stride); at + PACKET_BYTES <= bytes.length && hits < SNIFF_PACKETS; at += stride) {
			if (bytes[at] !== SYNC) {
				return false
			}

			hits++
		}

		return hits >= MIN_PACKETS
	}

	if (fits(expected)) {
		return expected
	}

	for (let residue = 0; residue < stride; residue++) {
		if (residue !== expected && fits(residue)) {
			return residue
		}
	}

	return null
}

// PTS of the timing stream's last PES starts, latest first. Walks back packet by packet reading only
// headers, and a whole packet only where the timing stream's PES starts.
async function tailTimestamps(source: BlockSource, pid: number, layout: TsLayout): Promise<number[]> {
	const { size } = source
	const { stride } = layout
	const floor = Math.max(0, size - TAIL_BYTES)
	const probeStart = Math.max(floor, size - SNIFF_PACKETS * stride - PACKET_BYTES)
	const grid = tailGrid(await readBytes(source, probeStart, size - probeStart), probeStart, stride, layout.prefix % stride)
	const found: number[] = []

	if (grid === null) {
		return found
	}

	const firstSync = probeStart + modulo(grid - probeStart, stride)

	for (
		let at = firstSync + Math.floor((size - PACKET_BYTES - firstSync) / stride) * stride;
		at >= floor && found.length < TAIL_PES;
		at -= stride
	) {
		const header = await readBytes(source, at, 4)

		if (header[0] !== SYNC || ((header[1] ?? 0) & 0x40) === 0 || ((((header[1] ?? 0) & 0x1f) << 8) | (header[2] ?? 0)) !== pid) {
			continue
		}

		const packet = packetAt(await readBytes(source, at, PACKET_BYTES), 0)
		const pts = packet === null ? null : pesStart(packet.payload)?.pts

		if (pts !== null && pts !== undefined) {
			found.push(pts)
		}
	}

	return found
}

function videoTrack(stream: Stream): ParsedTrack {
	const { picture } = stream

	return {
		kind: "video",
		enabled: true,
		video: {
			codec: stream.codec,
			fourcc: stream.fourcc,
			codecString: stream.codecString,
			w: picture?.width ?? 0,
			h: picture?.height ?? 0,
			par: picture?.par ?? 1,
			rotation: 0,
			bitDepth: picture?.bitDepth ?? null,
			transfer: transferKind(picture?.transfer ?? null)
		}
	}
}

// The duration is the timing stream's PTS span, head to tail, across one 33-bit wrap. The index is "none":
// TS has none, and the remux's demuxer seeks by bisecting PES timestamps, so no seek scans the file.
export async function walkMpegTs(source: BlockSource, tier: ProbeTier, result: ParsedContainer, layout: TsLayout): Promise<void> {
	const walk: Walk = {
		sections: new Map(),
		patFound: false,
		pmtPid: null,
		pmtFound: false,
		malformed: false,
		streams: [],
		byPid: new Map(),
		timingPid: null,
		firstPts: null
	}

	result.index = "none"

	if (tier === "full") {
		source.allowBlockReads(FULL_BLOCK_READS)
	}

	await scanHead(source, walk, layout, Math.min(source.size, HEAD_BYTES[tier]))

	let firstAudio = true

	for (const stream of walk.streams) {
		if (stream.kind === "video") {
			result.tracks.push(videoTrack(stream))

			continue
		}

		result.tracks.push({
			kind: "audio",
			enabled: true,
			audio: {
				id: stream.pid,
				codec: stream.codec,
				codecString: stream.codecString,
				ch: stream.ch,
				rate: stream.rate,
				lang: stream.lang,
				name: null,
				def: firstAudio,
				forced: false,
				commentary: false
			}
		})
		firstAudio = false
	}

	result.tracksFound = walk.pmtFound
	result.partial = walk.malformed || walk.streams.some(stream => stream.need !== null)

	if (walk.timingPid === null) {
		return
	}

	const first = walk.firstPts
	const lasts = first === null ? [] : await tailTimestamps(source, walk.timingPid, layout)
	let span = -1

	for (const last of lasts) {
		const elapsed = modulo(last - (first ?? 0), PTS_MODULUS)

		// Past half the clock's range the "last" PTS is one before the first, not a wrap.
		if (elapsed < PTS_MODULUS / 2) {
			span = Math.max(span, elapsed)
		}
	}

	if (span < 0) {
		result.partial = true
	} else {
		result.durationSec = seconds(span / PTS_HZ)
	}
}
