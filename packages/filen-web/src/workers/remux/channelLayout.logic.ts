// Channel layouts of decoded surround audio and their stereo downmix. The FFmpeg-based decoders hand out
// planes in FFmpeg's native order for the stream's layout (by speaker position: FL FR FC LFE BL BR FLC
// FRC BC SL SR), which the channel count alone does not pin down: 3/1 and 2/2 are both four channels,
// 2/1 and 3/0 both three. The layout is read from the bitstream's own header.

export type Speaker = "FL" | "FR" | "FC" | "LFE" | "BL" | "BR" | "FLC" | "FRC" | "BC" | "SL" | "SR"

// FFmpeg's channel bit order, which is the order of the decoded planes.
const NATIVE_ORDER: readonly Speaker[] = ["FL", "FR", "FC", "LFE", "BL", "BR", "FLC", "FRC", "BC", "SL", "SR"]

function native(speakers: readonly Speaker[]): Speaker[] {
	return NATIVE_ORDER.filter(speaker => speakers.includes(speaker))
}

// AC-3 / E-AC-3 audio coding mode (acmod) → speakers, before LFE. 0 is dual mono, played as stereo.
const ACMOD: readonly (readonly Speaker[])[] = [
	["FL", "FR"],
	["FC"],
	["FL", "FR"],
	["FL", "FR", "FC"],
	["FL", "FR", "BC"],
	["FL", "FR", "FC", "BC"],
	["FL", "FR", "SL", "SR"],
	["FL", "FR", "FC", "SL", "SR"]
]

// DTS core AMODE → speakers, before LFE and the XCh back centre. Values above 9 are user-defined.
const DTS_AMODE: readonly (readonly Speaker[])[] = [
	["FC"],
	["FL", "FR"],
	["FL", "FR"],
	["FL", "FR"],
	["FL", "FR"],
	["FL", "FR", "FC"],
	["FL", "FR", "BC"],
	["FL", "FR", "FC", "BC"],
	["FL", "FR", "SL", "SR"],
	["FL", "FR", "FC", "SL", "SR"]
]

class BitReader {
	private readonly bytes: Uint8Array
	private position: number

	public constructor(bytes: Uint8Array, start: number) {
		this.bytes = bytes
		this.position = start * 8
	}

	public read(bits: number): number | null {
		let value = 0

		for (let index = 0; index < bits; index++) {
			const byte = this.bytes[this.position >> 3]

			if (byte === undefined) {
				return null
			}

			value = (value << 1) | ((byte >> (7 - (this.position & 7))) & 1)
			this.position++
		}

		return value
	}

	public skip(bits: number): void {
		this.position += bits
	}
}

// Speakers of an AC-3 or E-AC-3 syncframe (the independent substream; dependent substreams' extra
// channels come on top, see layoutFor). Null when the bytes are no syncframe.
export function ac3Layout(frame: Uint8Array): Speaker[] | null {
	if (frame[0] !== 0x0b || frame[1] !== 0x77 || frame.length < 8) {
		return null
	}

	const bsid = (frame[5] ?? 0) >> 3

	if (bsid > 10) {
		// E-AC-3: fscod (2), numblkscod or fscod2 (2), acmod (3), lfeon (1) in byte 4.
		const byte = frame[4] ?? 0
		const speakers = ACMOD[(byte >> 1) & 7]

		return speakers === undefined ? null : native([...speakers, ...((byte & 1) === 1 ? (["LFE"] as const) : [])])
	}

	const reader = new BitReader(frame, 6)
	const acmod = reader.read(3)

	if (acmod === null) {
		return null
	}

	if ((acmod & 1) === 1 && acmod !== 1) {
		reader.skip(2)
	}

	if ((acmod & 4) === 4) {
		reader.skip(2)
	}

	if (acmod === 2) {
		reader.skip(2)
	}

	const lfeon = reader.read(1)
	const speakers = ACMOD[acmod]

	return speakers === undefined || lfeon === null ? null : native([...speakers, ...(lfeon === 1 ? (["LFE"] as const) : [])])
}

const DTS_SYNC = 0x7ffe8001
// The XCh extension: a back centre channel (DTS-ES 6.1).
const DTS_EXT_XCH = 0

// Speakers of a DTS core frame (big-endian 16-bit sync). Null when the bytes are no core frame or the
// mode is user-defined.
export function dtsLayout(frame: Uint8Array): Speaker[] | null {
	const reader = new BitReader(frame, 0)

	if (reader.read(32) !== DTS_SYNC) {
		return null
	}

	// FTYPE, SHORT, CPF, NBLKS, FSIZE.
	reader.skip(1 + 5 + 1 + 7 + 14)

	const amode = reader.read(6)

	// SFREQ, RATE, DOWNMIX, DYNF, TIMEF, AUXF, HDCD.
	reader.skip(4 + 5 + 1 + 1 + 1 + 1 + 1)

	const extensionId = reader.read(3)
	const extension = reader.read(1)

	// ASPF.
	reader.skip(1)

	const lff = reader.read(2)
	const speakers = amode === null ? undefined : DTS_AMODE[amode]

	if (speakers === undefined || lff === null || extension === null || extensionId === null) {
		return null
	}

	return native([
		...speakers,
		...(lff === 1 || lff === 2 ? (["LFE"] as const) : []),
		...(extension === 1 && extensionId === DTS_EXT_XCH ? (["BC"] as const) : [])
	])
}

// FFmpeg's default layouts by channel count, for a stream whose header says nothing usable.
const DEFAULT_LAYOUTS: Record<number, readonly Speaker[]> = {
	1: ["FC"],
	2: ["FL", "FR"],
	3: ["FL", "FR", "FC"],
	4: ["FL", "FR", "SL", "SR"],
	5: ["FL", "FR", "FC", "SL", "SR"],
	6: ["FL", "FR", "FC", "LFE", "SL", "SR"],
	7: ["FL", "FR", "FC", "LFE", "BC", "SL", "SR"],
	8: ["FL", "FR", "FC", "LFE", "BL", "BR", "SL", "SR"]
}

// Speakers added, in this order, when the decoder hands out more channels than the header named (an
// E-AC-3 dependent substream carrying the back pair of a 7.1 mix).
const EXTRA_SPEAKERS: readonly Speaker[] = ["BL", "BR", "BC", "FLC", "FRC"]

// The layout of `channels` decoded planes, from the header's layout when it fits.
export function layoutFor(header: readonly Speaker[] | null, channels: number): Speaker[] {
	if (header !== null && header.length === channels) {
		return [...header]
	}

	// Only on top of a full 5.1 core: a header naming fewer channels than were decoded is simply wrong.
	if (header !== null && header.length >= 6 && header.length < channels) {
		const extra = EXTRA_SPEAKERS.filter(speaker => !header.includes(speaker)).slice(0, channels - header.length)

		if (header.length + extra.length === channels) {
			return native([...header, ...extra])
		}
	}

	return [...(DEFAULT_LAYOUTS[channels] ?? NATIVE_ORDER.slice(0, channels))]
}

// -3 dB.
const MINUS_3_DB = Math.SQRT1_2
// A single surround channel splits into both sides at -3 dB each, -6 dB in all.
const MINUS_6_DB = 0.5

// Each speaker's gain into the left and right output (ITU-R BS.775 / ATSC A/52 Lo/Ro with centre and
// surround levels at -3 dB). LFE is not part of the downmix.
const GAINS: Record<Speaker, readonly [number, number]> = {
	FL: [1, 0],
	FR: [0, 1],
	FC: [MINUS_3_DB, MINUS_3_DB],
	LFE: [0, 0],
	BL: [MINUS_3_DB, 0],
	BR: [0, MINUS_3_DB],
	FLC: [1, 0],
	FRC: [0, 1],
	BC: [MINUS_6_DB, MINUS_6_DB],
	SL: [MINUS_3_DB, 0],
	SR: [0, MINUS_3_DB]
}

// Above this, samples are compressed smoothly towards full scale instead of clipping: the -3 dB mix
// of loud surround content can exceed 1, while dialogue keeps its level.
const LIMIT_KNEE = 0.9

function limit(value: number): number {
	const magnitude = Math.abs(value)

	if (magnitude <= LIMIT_KNEE) {
		return value
	}

	const limited = LIMIT_KNEE + (1 - LIMIT_KNEE) * Math.tanh((magnitude - LIMIT_KNEE) / (1 - LIMIT_KNEE))

	return value < 0 ? -limited : limited
}

// Stereo planes from planar input laid out as `layout`. Mono plays at full level on both sides.
export function downmixToStereo(planes: readonly Float32Array[], layout: readonly Speaker[]): [Float32Array, Float32Array] {
	const frames = planes[0]?.length ?? 0

	if (planes.length === 2 && layout[0] === "FL" && layout[1] === "FR") {
		return [planes[0] ?? new Float32Array(frames), planes[1] ?? new Float32Array(frames)]
	}

	const left = new Float32Array(frames)
	const right = new Float32Array(frames)
	const mono = layout.length === 1

	for (let channel = 0; channel < planes.length; channel++) {
		const plane = planes[channel]
		const speaker = layout[channel]

		if (plane === undefined || speaker === undefined) {
			continue
		}

		const [toLeft, toRight] = mono ? [1, 1] : GAINS[speaker]

		for (let frame = 0; frame < frames; frame++) {
			const sample = plane[frame] ?? 0

			left[frame] = (left[frame] ?? 0) + sample * toLeft
			right[frame] = (right[frame] ?? 0) + sample * toRight
		}
	}

	for (let frame = 0; frame < frames; frame++) {
		left[frame] = limit(left[frame] ?? 0)
		right[frame] = limit(right[frame] ?? 0)
	}

	return [left, right]
}
