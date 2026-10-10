// AC-3, E-AC-3 and DTS elementary streams on their own (`.ac3`, `.eac3`, `.dts`): no container, only frames
// back to back, with no timestamps. Most are constant-bitrate (AC-3 at 44.1 kHz alternates padded frames,
// which averages out), so the average unit size read from the head gives the duration and the offset of any
// time. A variable-bitrate one (DTS-HD Master Audio's lossless extension) gets a sampled timeline instead
// (rawTimeline).

export type RawAudioCodec = "ac3" | "eac3" | "dts"

export interface RawFrame {
	codec: RawAudioCodec
	bytes: number
	samples: number
	sampleRate: number
	channels: number
	// Starts an access unit: an AC-3 frame, a DTS core frame, an E-AC-3 independent frame of substream 0
	// (its dependent frames, a 7.1 mix's extra channels, belong to the same unit).
	unitStart: boolean
}

export interface RawStreamInfo {
	codec: RawAudioCodec
	// The first access unit's offset (leading junk or tags skipped).
	start: number
	// The average unit size: a fraction where units vary.
	unitBytes: number
	unitSamples: number
	sampleRate: number
	channels: number
	units: number
	duration: number
	// Unit sizes vary beyond padding: the duration and the offset of a time need a sampled timeline.
	variable: boolean
	// That timeline, once sampled; null for a constant-size stream.
	timeline: RawTimeline | null
}

// How far into the head a first sync word is looked for.
const MAX_LEADING_BYTES = 65_536

const AC3_RATES = [48_000, 44_100, 32_000]
const EAC3_HALF_RATES = [24_000, 22_050, 16_000]
const AC3_KBPS = [32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 384, 448, 512, 576, 640]
// acmod → full-bandwidth channels.
const AC3_CHANNELS = [2, 1, 2, 3, 3, 4, 4, 5]
const EAC3_BLOCKS = [1, 2, 3, 6]
const DTS_RATES = [0, 8000, 16_000, 32_000, 0, 0, 11_025, 22_050, 44_100, 0, 0, 12_000, 24_000, 48_000, 0, 0]
// AMODE → channels (0-15 defined).
const DTS_CHANNELS = [1, 2, 2, 2, 2, 3, 3, 4, 4, 5, 6, 6, 6, 7, 8, 8]

// Reads `count` bits from `bit` on, most significant first.
function bits(bytes: Uint8Array, bit: number, count: number): number {
	let value = 0

	for (let i = 0; i < count; i++) {
		const at = bit + i

		value = value * 2 + (((bytes[at >> 3] ?? 0) >> (7 - (at & 7))) & 1)
	}

	return value
}

function ac3FrameWords(fscod: number, frmsizecod: number): number | null {
	const kbps = AC3_KBPS[frmsizecod >> 1]

	if (kbps === undefined) {
		return null
	}

	switch (fscod) {
		case 0:
			return kbps * 2
		case 1:
			return Math.floor((kbps * 1000 * 1536) / 44_100 / 16) + (frmsizecod & 1)
		case 2:
			return kbps * 3
		default:
			return null
	}
}

function ac3Channels(bytes: Uint8Array, offset: number, acmodBit: number): number {
	const acmod = bits(bytes, offset * 8 + acmodBit, 3)
	// lfeon follows acmod after the mix level fields acmod implies.
	let skip = 0

	if ((acmod & 1) !== 0 && acmod !== 1) {
		skip += 2
	}

	if ((acmod & 4) !== 0) {
		skip += 2
	}

	if (acmod === 2) {
		skip += 2
	}

	const lfe = bits(bytes, offset * 8 + acmodBit + 3 + skip, 1)

	return (AC3_CHANNELS[acmod] ?? 2) + lfe
}

// The frame starting at `offset`, or null when no valid header is there (or not all of it is in `bytes`).
export function parseRawFrame(bytes: Uint8Array, offset: number): RawFrame | null {
	if (offset + 16 > bytes.length) {
		return null
	}

	if (bytes[offset] === 0x0b && bytes[offset + 1] === 0x77) {
		const bsid = bits(bytes, offset * 8 + 40, 5)

		if (bsid <= 10) {
			const fscod = bits(bytes, offset * 8 + 32, 2)
			const words = ac3FrameWords(fscod, bits(bytes, offset * 8 + 34, 6))
			const sampleRate = AC3_RATES[fscod]

			if (words === null || sampleRate === undefined) {
				return null
			}

			return { codec: "ac3", bytes: words * 2, samples: 1536, sampleRate, channels: ac3Channels(bytes, offset, 48), unitStart: true }
		}

		if (bsid > 16) {
			return null
		}

		const strmtyp = bits(bytes, offset * 8 + 16, 2)
		const substream = bits(bytes, offset * 8 + 18, 3)
		const frameBytes = (bits(bytes, offset * 8 + 21, 11) + 1) * 2
		const fscod = bits(bytes, offset * 8 + 32, 2)
		const sampleRate = fscod === 3 ? EAC3_HALF_RATES[bits(bytes, offset * 8 + 34, 2)] : AC3_RATES[fscod]
		const blocks = fscod === 3 ? 6 : (EAC3_BLOCKS[bits(bytes, offset * 8 + 34, 2)] ?? 6)
		const acmod = bits(bytes, offset * 8 + 36, 3)
		const lfe = bits(bytes, offset * 8 + 39, 1)

		if (sampleRate === undefined || strmtyp === 3) {
			return null
		}

		return {
			codec: "eac3",
			bytes: frameBytes,
			samples: blocks * 256,
			sampleRate,
			channels: (AC3_CHANNELS[acmod] ?? 2) + lfe,
			unitStart: strmtyp !== 1 && substream === 0
		}
	}

	if (bytes[offset] === 0x7f && bytes[offset + 1] === 0xfe && bytes[offset + 2] === 0x80 && bytes[offset + 3] === 0x01) {
		const base = offset * 8 + 32
		const blocks = bits(bytes, base + 7, 7) + 1
		const frameBytes = bits(bytes, base + 14, 14) + 1
		const amode = bits(bytes, base + 28, 6)
		const sampleRate = DTS_RATES[bits(bytes, base + 34, 4)] ?? 0
		const lfe = bits(bytes, base + 53, 2) === 0 ? 0 : 1

		if (blocks < 6 || frameBytes < 96 || sampleRate === 0) {
			return null
		}

		return {
			codec: "dts",
			bytes: frameBytes,
			samples: blocks * 32,
			sampleRate,
			channels: (DTS_CHANNELS[amode] ?? 2) + lfe,
			unitStart: true
		}
	}

	return null
}

// The first offset at or after `from` where a frame of `codec` (any of them when null) starts, -1 for none.
export function findRawSync(bytes: Uint8Array, from: number, codec: RawAudioCodec | null = null, limit: number = bytes.length): number {
	const end = Math.min(limit, bytes.length - 16)

	for (let at = Math.max(0, from); at <= end; at++) {
		if (bytes[at] !== 0x0b && bytes[at] !== 0x7f) {
			continue
		}

		const frame = parseRawFrame(bytes, at)

		if (frame !== null && (codec === null || frame.codec === codec || (codec !== "dts" && frame.codec !== "dts"))) {
			return at
		}
	}

	return -1
}

// The bytes a unit may span: a DTS-HD extension substream sits between core frames.
export const MAX_UNIT_BYTES = 65_536
// Units measured in the head for the average unit size.
const MEASURED_UNITS = 64

function sameFamily(a: RawAudioCodec, b: RawAudioCodec): boolean {
	return a === b || (a !== "dts" && b !== "dts")
}

// The length of the access unit of `codec` starting at `at`: its frames walked to where the next unit starts
// (which confirms it). `atEnd` says `bytes` ends where the stream does, so its last unit needs no successor.
// Null when no unit starts there or its end is not within `bytes`.
export function rawUnitLength(bytes: Uint8Array, at: number, codec: RawAudioCodec, atEnd: boolean): number | null {
	const first = parseRawFrame(bytes, at)

	if (first?.unitStart !== true || !sameFamily(first.codec, codec)) {
		return null
	}

	let next = at + first.bytes

	for (;;) {
		if (next + 16 > bytes.length) {
			return atEnd && next <= bytes.length ? next - at : null
		}

		const frame = parseRawFrame(bytes, next)

		if (frame === null) {
			const resync = codec === "dts" ? findRawSync(bytes, next, "dts", at + MAX_UNIT_BYTES) : -1

			if (resync === -1) {
				return atEnd && codec === "dts" && bytes.length - at <= MAX_UNIT_BYTES ? bytes.length - at : null
			}

			next = resync
			continue
		}

		if (!sameFamily(frame.codec, codec)) {
			return null
		}

		if (frame.unitStart) {
			return next - at
		}

		next += frame.bytes
	}
}

// The first confirmed unit start at or after `from`: a sync word inside other data or a dependent frame's is
// passed over. -1 for none.
export function nextRawUnit(bytes: Uint8Array, from: number, codec: RawAudioCodec, atEnd: boolean): number {
	for (let at = findRawSync(bytes, from, codec); at !== -1; at = findRawSync(bytes, at + 1, codec)) {
		if (rawUnitLength(bytes, at, codec, atEnd) !== null) {
			return at
		}
	}

	return -1
}

// The units from the first confirmed one at or after `from` on, walked up to `max` of them (or to the end):
// where the first starts, how many were walked and the bytes they span. Null when none starts there.
export function walkRawUnits(
	bytes: Uint8Array,
	from: number,
	codec: RawAudioCodec,
	atEnd: boolean,
	max: number
): { first: number; units: number; bytes: number } | null {
	const first = nextRawUnit(bytes, from, codec, atEnd)

	if (first === -1) {
		return null
	}

	let at = first
	let units = 0

	while (units < max) {
		const length = rawUnitLength(bytes, at, codec, atEnd)

		if (length === null) {
			break
		}

		at += length
		units++
	}

	return units === 0 ? null : { first, units, bytes: at - first }
}

// The stream in a file of `size` bytes whose head is `head`: its first confirmed access unit, and the average
// unit size over the head's units (AC-3 at 44.1 kHz alternates padded frames; DTS-HD extensions vary), which
// maps time to offset. Null when the head is not one.
export function measureRawStream(head: Uint8Array, size: number): RawStreamInfo | null {
	const atEnd = head.length >= size
	let start = -1
	let first: RawFrame | null = null

	for (let at = findRawSync(head, 0, null, MAX_LEADING_BYTES); at !== -1; at = findRawSync(head, at + 1, null, MAX_LEADING_BYTES)) {
		const frame = parseRawFrame(head, at)

		if (frame?.unitStart === true && rawUnitLength(head, at, frame.codec, atEnd) !== null) {
			start = at
			first = frame
			break
		}
	}

	if (first === null) {
		return null
	}

	const lengths: number[] = []

	for (let at = start; lengths.length < MEASURED_UNITS;) {
		const length = rawUnitLength(head, at, first.codec, atEnd)

		if (length === null) {
			break
		}

		lengths.push(length)
		at += length
	}

	const unitBytes = lengths.reduce((sum, length) => sum + length, 0) / lengths.length
	const units = Math.floor((size - start) / unitBytes)
	// Padding moves a unit by a few bytes; more than this is a variable bitrate.
	const variable = Math.max(...lengths) - Math.min(...lengths) > unitBytes * VARIABLE_SPREAD

	if (units <= 0) {
		return null
	}

	return {
		codec: first.codec === "ac3" && parseRawFrame(head, start + first.bytes)?.codec === "eac3" ? "eac3" : first.codec,
		start,
		unitBytes,
		unitSamples: first.samples,
		sampleRate: first.sampleRate,
		channels: first.channels,
		units,
		duration: (units * first.samples) / first.sampleRate,
		variable,
		timeline: null
	}
}

const VARIABLE_SPREAD = 0.01

// A variable-bitrate stream's map from time to bytes, from the average unit size measured at a few places
// (the head, the middle, ...) and an exact count of the units in a window at the tail. Time is cut into
// segments of whole units; each starts at the unit the map names for its first time and plays units in order
// from there until the next segment's start, whose own unit it then jumps to. So the audio at any time is the
// same whether reached by a seek or by playing on, and the tail plays to its last unit. What stays
// approximate is where each segment starts: by however far the map misses the local bitrate, a segment
// boundary repeats or skips that much audio.
export interface RawTimeline {
	// Offsets in increasing order, with their times; the last is the tail's first unit.
	knots: { offset: number; time: number }[]
	unitSeconds: number
	// Units per segment.
	segmentUnits: number
	// The unit index the tail starts at, and the stream's total.
	tailIndex: number
	units: number
}

// A segment's length, before it is rounded to whole units.
const SEGMENT_SECONDS = 5

export function rawTimeline(params: {
	start: number
	unitSeconds: number
	// Average unit sizes measured at offsets (the head's at `start`), in increasing order of offset.
	samples: { offset: number; unitBytes: number }[]
	// The tail's first unit, and the units and bytes from it to the end, counted.
	tail: { offset: number; units: number; bytes: number }
}): RawTimeline {
	const { unitSeconds, tail } = params
	const points = [...params.samples.filter(sample => sample.offset < tail.offset)]
	const knots = [{ offset: params.start, time: 0 }]
	let previous = points[0] ?? { offset: params.start, unitBytes: 1 }

	for (const point of [...points.slice(1), { offset: tail.offset, unitBytes: tail.bytes / Math.max(1, tail.units) }]) {
		const last = knots[knots.length - 1] ?? { offset: params.start, time: 0 }
		// The trapezoid between two measurements: units per byte changes linearly in between.
		const perByte = (unitSeconds * (1 / previous.unitBytes + 1 / point.unitBytes)) / 2

		knots.push({ offset: point.offset, time: last.time + (point.offset - last.offset) * perByte })
		previous = point
	}

	const tailKnot = knots[knots.length - 1] ?? { offset: tail.offset, time: 0 }
	const tailIndex = Math.round(tailKnot.time / unitSeconds)

	tailKnot.time = tailIndex * unitSeconds

	return {
		knots,
		unitSeconds,
		segmentUnits: Math.max(1, Math.round(SEGMENT_SECONDS / unitSeconds)),
		tailIndex,
		units: tailIndex + tail.units
	}
}

// The offset the timeline maps `time` to; the tail's first unit from its time on.
export function rawTimelineOffset(timeline: RawTimeline, time: number): number {
	const { knots } = timeline

	for (let i = 1; i < knots.length; i++) {
		const a = knots[i - 1]
		const b = knots[i]

		if (a !== undefined && b !== undefined && time < b.time) {
			return Math.round(a.offset + ((Math.max(time, a.time) - a.time) / (b.time - a.time)) * (b.offset - a.offset))
		}
	}

	return knots[knots.length - 1]?.offset ?? 0
}

// The segment holding unit `index`: its first unit and the first of the next (the stream's total for the last).
export function rawSegment(timeline: RawTimeline, index: number): { start: number; end: number; tail: boolean } {
	if (index >= timeline.tailIndex) {
		return { start: timeline.tailIndex, end: timeline.units, tail: true }
	}

	const start = Math.floor(index / timeline.segmentUnits) * timeline.segmentUnits

	return { start, end: Math.min(start + timeline.segmentUnits, timeline.tailIndex), tail: false }
}

// Where the unit playback of `time` starts in lies (about: unit sizes vary), and its index.
export function rawUnitAt(info: RawStreamInfo, time: number): { index: number; offset: number } {
	const unitSeconds = info.unitSamples / info.sampleRate
	const index = Math.max(0, Math.min(info.units - 1, Math.floor(time / unitSeconds)))

	return { index, offset: info.start + Math.round(index * info.unitBytes) }
}

// The index of the unit starting at `offset`.
export function rawUnitIndex(info: RawStreamInfo, offset: number): number {
	return Math.max(0, Math.round((offset - info.start) / info.unitBytes))
}
