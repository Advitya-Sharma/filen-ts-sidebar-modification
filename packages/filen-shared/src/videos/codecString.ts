// RFC 6381 `codecs` values built from the raw codec configuration records a container carries (MP4 sample
// entry children, Matroska CodecPrivate). A record that cannot be read gives null; nothing here throws.

function hex(value: number, digits: number): string {
	return value.toString(16).padStart(digits, "0")
}

function dec(value: number, digits: number): string {
	return String(value).padStart(digits, "0")
}

// MSB-first bit reader over a byte array; null once a read would run past the end.
class BitReader {
	private position = 0
	private readonly bytes: Uint8Array

	public constructor(bytes: Uint8Array) {
		this.bytes = bytes
	}

	public get left(): number {
		return this.bytes.length * 8 - this.position
	}

	public read(count: number): number | null {
		if (count > this.left) {
			return null
		}

		let value = 0

		for (let i = 0; i < count; i++) {
			const byte = this.bytes[this.position >> 3] ?? 0

			value = value * 2 + ((byte >> (7 - (this.position & 7))) & 1)
			this.position++
		}

		return value
	}
}

// ── Video ────────────────────────────────────────────────────────────────────

export interface AvcConfig {
	profile: number
	compatibility: number
	level: number
	bitDepth: number
}

// High-family profiles whose avcC may carry the chroma/bit-depth extension after the parameter sets.
const AVC_HIGH_PROFILES = new Set([100, 110, 122, 244, 44, 83, 86, 118, 128, 138, 139, 134, 135])

// AVCDecoderConfigurationRecord (the avcC box payload, or a Matroska V_MPEG4/ISO/AVC CodecPrivate).
export function parseAvcC(avcC: Uint8Array): AvcConfig | null {
	const version = avcC[0]
	const profile = avcC[1]
	const compatibility = avcC[2]
	const level = avcC[3]

	if (version !== 1 || profile === undefined || compatibility === undefined || level === undefined) {
		return null
	}

	// High 10 and up imply more than 8 bits even when the extension is absent.
	let bitDepth = profile === 110 || profile === 122 || profile === 244 ? 10 : 8

	if (AVC_HIGH_PROFILES.has(profile)) {
		let at = 5
		const spsCount = (avcC[at] ?? 0) & 0x1f

		at++

		for (let i = 0; i < spsCount && at + 2 <= avcC.length; i++) {
			at += 2 + ((avcC[at] ?? 0) << 8) + (avcC[at + 1] ?? 0)
		}

		const ppsCount = avcC[at] ?? 0

		at++

		for (let i = 0; i < ppsCount && at + 2 <= avcC.length; i++) {
			at += 2 + ((avcC[at] ?? 0) << 8) + (avcC[at + 1] ?? 0)
		}

		// chroma_format(2), bit_depth_luma_minus8(3), bit_depth_chroma_minus8(3), each after reserved bits.
		const luma = avcC[at + 1]

		if (at + 4 <= avcC.length && luma !== undefined) {
			bitDepth = (luma & 0x07) + 8
		}
	}

	return { profile, compatibility, level, bitDepth }
}

// `fourcc` is the sample entry type (avc1, avc3); anything else is written as avc1.
export function avcCodecString({ fourcc, avcC }: { fourcc: string; avcC: Uint8Array }): string | null {
	const config = parseAvcC(avcC)

	if (config === null) {
		return null
	}

	return `${fourcc === "avc3" ? "avc3" : "avc1"}.${hex(config.profile, 2)}${hex(config.compatibility, 2)}${hex(config.level, 2)}`
}

export interface HevcConfig {
	profileSpace: number
	tier: number
	profile: number
	compatibility: number
	constraints: number[]
	level: number
	bitDepth: number
}

// HEVCDecoderConfigurationRecord (the hvcC box payload, or a Matroska V_MPEGH/ISO/HEVC CodecPrivate).
export function parseHvcC(hvcC: Uint8Array): HevcConfig | null {
	if (hvcC.length < 23 || hvcC[0] !== 1) {
		return null
	}

	const first = hvcC[1] ?? 0
	const constraints: number[] = []

	for (let i = 6; i < 12; i++) {
		constraints.push(hvcC[i] ?? 0)
	}

	return {
		profileSpace: first >> 6,
		tier: (first >> 5) & 1,
		profile: first & 0x1f,
		compatibility: (hvcC[2] ?? 0) * 0x1000000 + ((hvcC[3] ?? 0) << 16) + ((hvcC[4] ?? 0) << 8) + (hvcC[5] ?? 0),
		constraints,
		level: hvcC[12] ?? 0,
		bitDepth: ((hvcC[17] ?? 0) & 0x07) + 8
	}
}

function reverseBits32(value: number): number {
	let reversed = 0

	for (let i = 0; i < 32; i++) {
		reversed = reversed * 2 + (Math.floor(value / 2 ** i) % 2)
	}

	return reversed
}

// `fourcc` is the sample entry type (hvc1, hev1); anything else is written as hvc1, which is what a remux writes.
export function hevcCodecString({ fourcc, hvcC }: { fourcc: string; hvcC: Uint8Array }): string | null {
	const config = parseHvcC(hvcC)

	if (config === null) {
		return null
	}

	const space = ["", "A", "B", "C"][config.profileSpace] ?? ""
	const constraints = [...config.constraints]

	while (constraints.length > 0 && constraints[constraints.length - 1] === 0) {
		constraints.pop()
	}

	const parts = [
		fourcc === "hev1" ? "hev1" : "hvc1",
		`${space}${String(config.profile)}`,
		reverseBits32(config.compatibility).toString(16).toUpperCase(),
		`${config.tier === 1 ? "H" : "L"}${String(config.level)}`,
		...constraints.map(byte => byte.toString(16).toUpperCase())
	]

	return parts.join(".")
}

// ISO/IEC 23091-2 colour description; `fullRange` false is limited range.
export interface ColourInfo {
	primaries: number
	transfer: number
	matrix: number
	fullRange: boolean
}

// 2 means "unspecified", which says nothing beyond the BT.709 default.
function knownColour(colour: ColourInfo | null | undefined): ColourInfo | null {
	if (colour === null || colour === undefined) {
		return null
	}

	return {
		primaries: colour.primaries === 2 ? 1 : colour.primaries,
		transfer: colour.transfer === 2 ? 1 : colour.transfer,
		matrix: colour.matrix === 2 ? 1 : colour.matrix,
		fullRange: colour.fullRange
	}
}

function isDefaultColour(colour: ColourInfo): boolean {
	return colour.primaries === 1 && colour.transfer === 1 && colour.matrix === 1 && !colour.fullRange
}

export interface Av1Config {
	profile: number
	level: number
	tier: number
	bitDepth: number
	monochrome: boolean
	subsamplingX: number
	subsamplingY: number
	chromaSamplePosition: number
}

// AV1CodecConfigurationRecord (the av1C box payload, or a Matroska V_AV1 CodecPrivate).
export function parseAv1C(av1C: Uint8Array): Av1Config | null {
	const marker = av1C[0]
	const first = av1C[1]
	const second = av1C[2]

	if (marker !== 0x81 || first === undefined || second === undefined) {
		return null
	}

	const profile = first >> 5
	const highBitDepth = (second >> 6) & 1
	const twelveBit = (second >> 5) & 1

	return {
		profile,
		level: first & 0x1f,
		tier: second >> 7,
		bitDepth: profile === 2 && highBitDepth === 1 ? (twelveBit === 1 ? 12 : 10) : highBitDepth === 1 ? 10 : 8,
		monochrome: ((second >> 4) & 1) === 1,
		subsamplingX: (second >> 3) & 1,
		subsamplingY: (second >> 2) & 1,
		chromaSamplePosition: second & 0x03
	}
}

// av01.P.LLT.DD, with the optional .M.CCC.cp.tc.mc.F tail only when the colour is known and not the default.
export function av1CodecString({ av1C, colour }: { av1C: Uint8Array; colour?: ColourInfo | null }): string | null {
	const config = parseAv1C(av1C)

	if (config === null) {
		return null
	}

	const base = `av01.${String(config.profile)}.${dec(config.level, 2)}${config.tier === 1 ? "H" : "M"}.${dec(config.bitDepth, 2)}`
	const known = knownColour(colour)
	const chroma = `${String(config.subsamplingX)}${String(config.subsamplingY)}${String(config.chromaSamplePosition)}`
	const defaultChroma = !config.monochrome && config.subsamplingX === 1 && config.subsamplingY === 1

	if ((known === null || isDefaultColour(known)) && defaultChroma) {
		return base
	}

	const tail = known ?? { primaries: 1, transfer: 1, matrix: 1, fullRange: false }

	return [
		base,
		config.monochrome ? "1" : "0",
		chroma,
		dec(tail.primaries, 2),
		dec(tail.transfer, 2),
		dec(tail.matrix, 2),
		tail.fullRange ? "1" : "0"
	].join(".")
}

export interface Vp9Config {
	profile: number
	level: number
	bitDepth: number
	// 0/1 4:2:0 (vertical, colocated), 2 4:2:2, 3 4:4:4.
	chromaSubsampling: number
	colour: ColourInfo | null
}

// VPCodecConfigurationRecord: the vpcC box payload, version/flags included.
export function parseVpcC(vpcC: Uint8Array): Vp9Config | null {
	const version = vpcC[0]
	const profile = vpcC[4]
	const level = vpcC[5]
	const packed = vpcC[6]

	if (version === undefined || profile === undefined || level === undefined || packed === undefined) {
		return null
	}

	if (version === 0) {
		// The pre-standard layout: bitDepth(4) colorSpace(4), then chromaSubsampling(4) and the rest.
		return { profile, level, bitDepth: packed >> 4, chromaSubsampling: (vpcC[7] ?? 0x10) >> 4, colour: null }
	}

	const primaries = vpcC[7]
	const transfer = vpcC[8]
	const matrix = vpcC[9]

	return {
		profile,
		level,
		bitDepth: packed >> 4,
		chromaSubsampling: (packed >> 1) & 0x07,
		colour:
			primaries === undefined || transfer === undefined || matrix === undefined
				? null
				: { primaries, transfer, matrix, fullRange: (packed & 1) === 1 }
	}
}

// Matroska V_VP9 CodecPrivate: optional {id, length, value} features. Absent features stay undefined.
export function parseMatroskaVp9Private(bytes: Uint8Array): {
	profile?: number
	level?: number
	bitDepth?: number
	chromaSubsampling?: number
} {
	const features: { profile?: number; level?: number; bitDepth?: number; chromaSubsampling?: number } = {}
	let at = 0

	while (at + 2 <= bytes.length) {
		const id = bytes[at] ?? 0
		const length = bytes[at + 1] ?? 0
		const value = bytes[at + 2]

		if (length !== 1 || value === undefined) {
			break
		}

		if (id === 1) {
			features.profile = value
		} else if (id === 2) {
			features.level = value
		} else if (id === 3) {
			features.bitDepth = value
		} else if (id === 4) {
			features.chromaSubsampling = value
		}

		at += 2 + length
	}

	return features
}

// VP9 levels by the largest luma picture each allows, for a stream whose container does not say.
const VP9_LEVELS: readonly (readonly [number, number])[] = [
	[10, 36_864],
	[11, 73_728],
	[20, 122_880],
	[21, 245_760],
	[30, 552_960],
	[31, 983_040],
	[40, 2_228_224],
	[50, 8_912_896],
	[60, 35_651_584]
]

export function vp9LevelForSize({ width, height }: { width: number; height: number }): number {
	const area = width * height

	for (const [level, maxArea] of VP9_LEVELS) {
		if (area <= maxArea) {
			return level
		}
	}

	return 62
}

// vp09.PP.LL.DD, with the .CC.cp.tc.mc.FF tail only when the colour is known and not the default.
export function vp9CodecString(config: Vp9Config): string | null {
	const { profile, level, bitDepth, chromaSubsampling } = config

	if (profile > 3 || level === 0 || (bitDepth !== 8 && bitDepth !== 10 && bitDepth !== 12) || chromaSubsampling > 3) {
		return null
	}

	const base = `vp09.${dec(profile, 2)}.${dec(level, 2)}.${dec(bitDepth, 2)}`
	const known = knownColour(config.colour)

	if ((known === null || isDefaultColour(known)) && chromaSubsampling <= 1) {
		return base
	}

	const tail = known ?? { primaries: 1, transfer: 1, matrix: 1, fullRange: false }

	return [
		base,
		dec(chromaSubsampling, 2),
		dec(tail.primaries, 2),
		dec(tail.transfer, 2),
		dec(tail.matrix, 2),
		tail.fullRange ? "01" : "00"
	].join(".")
}

// ── Audio ────────────────────────────────────────────────────────────────────

export interface AacConfig {
	// The object type the codec string names: 5 (SBR) or 29 (PS) when signalled, else the core type.
	objectType: number
	sampleRate: number | null
	channels: number | null
}

const AAC_SAMPLE_RATES = [96000, 88200, 64000, 48000, 44100, 32000, 24000, 22050, 16000, 12000, 11025, 8000, 7350]
const AAC_CHANNELS: Readonly<Record<number, number>> = { 1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 8, 11: 7, 12: 8, 13: 24, 14: 8 }

function aacObjectType(bits: BitReader): number | null {
	const type = bits.read(5)

	if (type !== 31) {
		return type
	}

	const escaped = bits.read(6)

	return escaped === null ? null : 32 + escaped
}

function aacSampleRate(bits: BitReader): number | null {
	const index = bits.read(4)

	if (index === 0x0f) {
		return bits.read(24)
	}

	return index === null ? null : (AAC_SAMPLE_RATES[index] ?? null)
}

// Object types whose config is a GASpecificConfig.
const GA_OBJECT_TYPES = new Set([1, 2, 3, 4, 6, 7, 17, 19, 20, 21, 22, 23])

// Skips a GASpecificConfig; false when it holds a program config element this reader does not walk.
function skipGaSpecificConfig(bits: BitReader, objectType: number, channelConfiguration: number): boolean {
	if (channelConfiguration === 0) {
		return false
	}

	bits.read(1)

	if (bits.read(1) === 1) {
		bits.read(14)
	}

	const extension = bits.read(1)

	if (objectType === 6 || objectType === 20) {
		bits.read(3)
	}

	if (extension === 1) {
		if (objectType === 22) {
			bits.read(16)
		}

		if (objectType === 17 || objectType === 19 || objectType === 20 || objectType === 23) {
			bits.read(3)
		}

		bits.read(1)
	}

	return true
}

// AudioSpecificConfig: an esds DecoderSpecificInfo, or a Matroska A_AAC CodecPrivate.
export function parseAudioSpecificConfig(config: Uint8Array): AacConfig | null {
	const bits = new BitReader(config)
	let objectType = aacObjectType(bits)
	let sampleRate = aacSampleRate(bits)
	const channelConfiguration = bits.read(4)

	if (objectType === null || objectType === 0 || sampleRate === null || channelConfiguration === null) {
		return null
	}

	const channels = AAC_CHANNELS[channelConfiguration] ?? null

	// Explicit hierarchical signalling: the SBR/PS type first, then the output rate and the core type.
	if (objectType === 5 || objectType === 29) {
		const extensionRate = aacSampleRate(bits)
		const core = aacObjectType(bits)

		return core === null ? null : { objectType, sampleRate: extensionRate ?? sampleRate, channels: objectType === 29 ? 2 : channels }
	}

	// Backward-compatible signalling: a sync extension after the core config.
	if (GA_OBJECT_TYPES.has(objectType) && skipGaSpecificConfig(bits, objectType, channelConfiguration) && bits.left >= 16) {
		if (bits.read(11) === 0x2b7 && aacObjectType(bits) === 5 && bits.read(1) === 1) {
			sampleRate = aacSampleRate(bits) ?? sampleRate
			objectType = 5

			if (bits.left >= 12 && bits.read(11) === 0x548 && bits.read(1) === 1) {
				objectType = 29
			}
		}
	}

	return { objectType, sampleRate, channels: objectType === 29 ? 2 : channels }
}

// An MP4 audio track's string from its esds objectTypeIndication and, for MPEG-4 audio, its AudioSpecificConfig.
export function mp4aCodecString({
	objectTypeIndication,
	audioSpecificConfig
}: {
	objectTypeIndication: number
	audioSpecificConfig: Uint8Array | null
}): string | null {
	switch (objectTypeIndication) {
		case 0x40: {
			const config = audioSpecificConfig === null ? null : parseAudioSpecificConfig(audioSpecificConfig)

			return config === null ? null : `mp4a.40.${String(config.objectType)}`
		}

		case 0x66:
		case 0x67:
		case 0x68:
		case 0x69:
		case 0x6b:
			return `mp4a.${hex(objectTypeIndication, 2).toUpperCase()}`
		case 0xa5:
			return "ac-3"
		case 0xa6:
			return "ec-3"
		case 0xa9:
			return "dtsc"
		case 0xaa:
			return "dtsh"
		case 0xab:
			return "dtsl"
		case 0xac:
			return "dtse"
		case 0xad:
			return "opus"
		default:
			return null
	}
}

// Matroska's legacy A_AAC/<version>/<profile> IDs from before CodecPrivate carried the config.
const MATROSKA_AAC_PROFILES: Readonly<Record<string, number>> = { MAIN: 1, LC: 2, SSR: 3, LTP: 4, "LC/SBR": 5 }

export function matroskaAacCodecString({ codecId, codecPrivate }: { codecId: string; codecPrivate: Uint8Array | null }): string | null {
	if (codecPrivate !== null && codecPrivate.length >= 2) {
		return mp4aCodecString({ objectTypeIndication: 0x40, audioSpecificConfig: codecPrivate })
	}

	const profile = MATROSKA_AAC_PROFILES[codecId.replace(/^A_AAC\/MPEG[24]\//, "")]

	return profile === undefined ? null : `mp4a.40.${String(profile)}`
}

export type SimpleAudioCodec = "mp3" | "opus" | "flac" | "vorbis" | "ac3" | "eac3" | "truehd" | "alac"

const SIMPLE_AUDIO: Readonly<Record<SimpleAudioCodec, string>> = {
	mp3: "mp4a.6B",
	opus: "opus",
	flac: "flac",
	vorbis: "vorbis",
	ac3: "ac-3",
	eac3: "ec-3",
	truehd: "mlpa",
	alac: "alac"
}

// Codecs whose string carries nothing from their config.
export function simpleAudioCodecString(codec: SimpleAudioCodec): string {
	return SIMPLE_AUDIO[codec]
}

// DTS by its MP4 sample entry type or Matroska CodecID: core, Express (LBR), HD High Resolution, HD Master Audio.
export function dtsCodecString(name: string): string | null {
	switch (name) {
		case "dtsc":
		case "A_DTS":
			return "dtsc"
		case "dtse":
		case "A_DTS/EXPRESS":
			return "dtse"
		case "dtsh":
			return "dtsh"
		case "dtsl":
		case "A_DTS/LOSSLESS":
			return "dtsl"
		default:
			return null
	}
}

// ── Dolby Vision ─────────────────────────────────────────────────────────────

// Dolby Vision sample entries by the base codec's: hvc1 → dvh1, hev1 → dvhe, av01 → dav1, avc1 → dva1,
// avc3 → dvav. A Matroska track is written as a remux would write it.
const DOLBY_VISION_ENTRIES: Readonly<Record<string, string>> = {
	hvc1: "dvh1",
	hev1: "dvhe",
	dvh1: "dvh1",
	dvhe: "dvhe",
	av01: "dav1",
	dav1: "dav1",
	avc1: "dva1",
	avc3: "dvav",
	dva1: "dva1",
	dvav: "dvav"
}

// The Dolby Vision string (dvh1.PP.LL) a decoder must accept when the stream has no base layer that
// plays without Dolby Vision; null for a sample entry type that carries no Dolby Vision.
export function dolbyVisionCodecString({ fourcc, profile, level }: { fourcc: string; profile: number; level: number }): string | null {
	const entry = DOLBY_VISION_ENTRIES[fourcc]

	return entry === undefined ? null : `${entry}.${dec(profile, 2)}.${dec(level, 2)}`
}

// ── Equivalent spellings ─────────────────────────────────────────────────────

// Strings browsers answer differently for the same codec: Firefox accepts MP3 in MP4 only as "mp3",
// others only as an object type; FLAC and Opus have sample-entry spellings; AC-3/E-AC-3 have object types.
const EQUIVALENT_CODEC_STRINGS: readonly (readonly string[])[] = [
	["mp4a.6B", "mp4a.69", "mp3", "mp4a.6b"],
	["flac", "fLaC"],
	["opus", "Opus"],
	["ac-3", "mp4a.a5", "mp4a.A5"],
	["ec-3", "mp4a.a6", "mp4a.A6"]
]

// `codecString` first, then every spelling a capability query may also try for it.
export function codecStringAlternatives(codecString: string): string[] {
	const group = EQUIVALENT_CODEC_STRINGS.find(spellings => spellings.includes(codecString))

	return group === undefined ? [codecString] : [codecString, ...group.filter(spelling => spelling !== codecString)]
}
