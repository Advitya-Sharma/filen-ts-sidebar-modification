import { describe, expect, it } from "vitest"
import {
	av1CodecString,
	avcCodecString,
	codecStringAlternatives,
	dolbyVisionCodecString,
	dtsCodecString,
	hevcCodecString,
	matroskaAacCodecString,
	mp4aCodecString,
	parseAudioSpecificConfig,
	parseAvcC,
	parseAv1C,
	parseHvcC,
	parseMatroskaVp9Private,
	parseVpcC,
	simpleAudioCodecString,
	vp9CodecString,
	vp9LevelForSize
} from "@filen/shared/videos"

function fromHex(value: string): Uint8Array {
	const bytes = new Uint8Array(value.length / 2)

	for (let i = 0; i < bytes.length; i++) {
		bytes[i] = parseInt(value.slice(i * 2, i * 2 + 2), 16)
	}

	return bytes
}

// MSB-first bit packing for hand-built AudioSpecificConfigs.
function bits(...fields: (readonly [number, number])[]): Uint8Array {
	const out: number[] = []
	let current = 0
	let filled = 0

	for (const [value, width] of fields) {
		for (let i = width - 1; i >= 0; i--) {
			current = (current << 1) | (Math.floor(value / 2 ** i) % 2)
			filled++

			if (filled === 8) {
				out.push(current)
				current = 0
				filled = 0
			}
		}
	}

	if (filled > 0) {
		out.push(current << (8 - filled))
	}

	return Uint8Array.from(out)
}

// Config records as written by the muxers that made the local video fixtures (x264, x265, SVT-AV1, ffmpeg).
const X264_HIGH_L31_AVCC = fromHex("0164001fffe1001a6764001facd9405005bb011000000300100000030300f183196001000468ef89cbfdf8f800")
const X264_BASELINE_L12_AVCC = fromHex("0142c00cffe100176742c00cda05067e7c0440000003004000000783c50aa801000568ce013720")
const X265_MAIN10_L31_HVCC = fromHex("0102200000009000000000005df000fcfdfafa00000f04")
const SVT_AV1_MAIN_L31_AV1C = fromHex("81050c000a0c0200002d2a67fd9e017c0020")
const FFMPEG_AAC_LC_48K_STEREO_ASC = fromHex("119056e500")

describe("avcCodecString", () => {
	it("writes profile, compatibility and level as hex (RFC 6381 §3.3)", () => {
		expect(avcCodecString({ fourcc: "avc1", avcC: fromHex("0142e01effe000") })).toBe("avc1.42e01e")
		expect(avcCodecString({ fourcc: "avc3", avcC: fromHex("014d4028ffe000") })).toBe("avc3.4d4028")
	})

	it("matches the fixtures' x264 records", () => {
		expect(avcCodecString({ fourcc: "avc1", avcC: X264_HIGH_L31_AVCC })).toBe("avc1.64001f")
		expect(avcCodecString({ fourcc: "avc1", avcC: X264_BASELINE_L12_AVCC })).toBe("avc1.42c00c")
	})

	it("reads the bit depth from the High-profile extension (ISO/IEC 14496-15 §5.3.3.1)", () => {
		expect(parseAvcC(X264_HIGH_L31_AVCC)?.bitDepth).toBe(8)
		// High 10, one empty SPS and PPS, then chroma_format 1 and bit_depth_luma_minus8 2.
		expect(parseAvcC(fromHex("016e001fffe10000010000fdfaf800"))?.bitDepth).toBe(10)
		// High 10 without the extension still implies more than 8 bits.
		expect(parseAvcC(fromHex("016e001fffe00000"))?.bitDepth).toBe(10)
	})

	it("rejects a record of the wrong version or too short", () => {
		expect(avcCodecString({ fourcc: "avc1", avcC: fromHex("0264001f") })).toBeNull()
		expect(avcCodecString({ fourcc: "avc1", avcC: fromHex("0164") })).toBeNull()
		expect(avcCodecString({ fourcc: "avc1", avcC: new Uint8Array(0) })).toBeNull()
	})
})

describe("hevcCodecString", () => {
	it("encodes the Main profile example of ISO/IEC 14496-15 Annex E", () => {
		const hvcC = new Uint8Array(23)

		hvcC.set([1, 0x01, 0x60, 0, 0, 0, 0xb0, 0, 0, 0, 0, 0, 93])
		expect(hevcCodecString({ fourcc: "hev1", hvcC })).toBe("hev1.1.6.L93.B0")
	})

	it("matches the fixture's x265 Main 10 record", () => {
		expect(hevcCodecString({ fourcc: "hvc1", hvcC: X265_MAIN10_L31_HVCC })).toBe("hvc1.2.4.L93.90")
		expect(parseHvcC(X265_MAIN10_L31_HVCC)?.bitDepth).toBe(10)
	})

	it("writes the profile space letter, the high tier and every non-trailing constraint byte (ISO/IEC 14496-15 E.3)", () => {
		const hvcC = new Uint8Array(23)

		hvcC.set([1, (2 << 6) | (1 << 5) | 4, 0x00, 0x00, 0x00, 0x08, 0xb0, 0x00, 0x23, 0, 0, 0, 153])
		expect(hevcCodecString({ fourcc: "hvc1", hvcC })).toBe("hvc1.B4.10000000.H153.B0.0.23")
	})

	it("writes a remuxed track as hvc1 and rejects a short record", () => {
		expect(hevcCodecString({ fourcc: "V_MPEGH/ISO/HEVC", hvcC: X265_MAIN10_L31_HVCC })?.startsWith("hvc1.")).toBe(true)
		expect(hevcCodecString({ fourcc: "hvc1", hvcC: X265_MAIN10_L31_HVCC.subarray(0, 12) })).toBeNull()
	})
})

describe("av1CodecString", () => {
	it("matches the fixture's SVT-AV1 record", () => {
		expect(av1CodecString({ av1C: SVT_AV1_MAIN_L31_AV1C })).toBe("av01.0.05M.08")
	})

	it("adds the colour tail only for known, non-default colour (AV1 ISOBMFF binding, codecs parameter string)", () => {
		// Main, level 4, 10-bit, 4:2:0 colocated: the binding's HDR example.
		const av1C = fromHex("8104 4e 00".replaceAll(" ", ""))

		expect(av1CodecString({ av1C, colour: { primaries: 9, transfer: 16, matrix: 9, fullRange: false } })).toBe(
			"av01.0.04M.10.0.112.09.16.09.0"
		)
		expect(av1CodecString({ av1C: fromHex("81040c00"), colour: { primaries: 1, transfer: 1, matrix: 1, fullRange: false } })).toBe(
			"av01.0.04M.08"
		)
		// Unspecified (2) is not a colour worth writing.
		expect(av1CodecString({ av1C: fromHex("81040c00"), colour: { primaries: 2, transfer: 2, matrix: 2, fullRange: false } })).toBe(
			"av01.0.04M.08"
		)
	})

	it("reads 12-bit only for the Professional profile", () => {
		expect(parseAv1C(fromHex("8148e000"))?.bitDepth).toBe(12)
		expect(parseAv1C(fromHex("8108e000"))?.bitDepth).toBe(10)
		expect(av1CodecString({ av1C: fromHex("8248e000") })).toBeNull()
	})
})

describe("vp9CodecString", () => {
	it("writes the short form for default colour and the full form otherwise (VP codec ISOBMFF binding §5)", () => {
		expect(vp9CodecString({ profile: 0, level: 41, bitDepth: 8, chromaSubsampling: 1, colour: null })).toBe("vp09.00.41.08")
		expect(
			vp9CodecString({
				profile: 2,
				level: 10,
				bitDepth: 10,
				chromaSubsampling: 1,
				colour: { primaries: 9, transfer: 16, matrix: 9, fullRange: true }
			})
		).toBe("vp09.02.10.10.01.09.16.09.01")
	})

	it("parses a version 1 vpcC box payload", () => {
		const config = parseVpcC(fromHex("01000000" + "00" + "1f" + "82" + "010101" + "0000"))

		expect(config).toEqual({
			profile: 0,
			level: 31,
			bitDepth: 8,
			chromaSubsampling: 1,
			colour: { primaries: 1, transfer: 1, matrix: 1, fullRange: false }
		})
		expect(config === null ? null : vp9CodecString(config)).toBe("vp09.00.31.08")
	})

	it("reads Matroska CodecPrivate features and falls back to a level from the frame size", () => {
		expect(parseMatroskaVp9Private(fromHex("010102020129030108040101"))).toEqual({
			profile: 2,
			level: 41,
			bitDepth: 8,
			chromaSubsampling: 1
		})
		expect(parseMatroskaVp9Private(new Uint8Array(0))).toEqual({})
		expect(vp9LevelForSize({ width: 1280, height: 720 })).toBe(31)
		expect(vp9LevelForSize({ width: 1920, height: 1080 })).toBe(40)
		expect(vp9LevelForSize({ width: 16384, height: 16384 })).toBe(62)
	})

	it("rejects out-of-range fields", () => {
		expect(vp9CodecString({ profile: 4, level: 10, bitDepth: 8, chromaSubsampling: 1, colour: null })).toBeNull()
		expect(vp9CodecString({ profile: 0, level: 0, bitDepth: 8, chromaSubsampling: 1, colour: null })).toBeNull()
		expect(vp9CodecString({ profile: 0, level: 10, bitDepth: 9, chromaSubsampling: 1, colour: null })).toBeNull()
		expect(parseVpcC(fromHex("0100"))).toBeNull()
	})
})

describe("AAC", () => {
	it("reads the fixtures' ffmpeg LC config, whose sync extension says no SBR", () => {
		expect(parseAudioSpecificConfig(FFMPEG_AAC_LC_48K_STEREO_ASC)).toEqual({ objectType: 2, sampleRate: 48000, channels: 2 })
		expect(mp4aCodecString({ objectTypeIndication: 0x40, audioSpecificConfig: FFMPEG_AAC_LC_48K_STEREO_ASC })).toBe("mp4a.40.2")
	})

	it("names explicit SBR and PS signalling by the signalled type (ISO/IEC 14496-3 §1.6.2.1)", () => {
		// AOT 5, 24 kHz core, stereo, 48 kHz output, core AOT 2.
		expect(mp4aCodecString({ objectTypeIndication: 0x40, audioSpecificConfig: bits([5, 5], [6, 4], [2, 4], [3, 4], [2, 5]) })).toBe(
			"mp4a.40.5"
		)
		expect(parseAudioSpecificConfig(bits([29, 5], [6, 4], [1, 4], [3, 4], [2, 5]))).toEqual({
			objectType: 29,
			sampleRate: 48000,
			channels: 2
		})
	})

	it("names backward-compatible SBR and PS signalling from the sync extensions (ISO/IEC 14496-3 §1.6.5)", () => {
		const core = [
			[2, 5],
			[6, 4],
			[1, 4],
			[0, 3]
		] as const
		const sbr = bits(...core, [0x2b7, 11], [5, 5], [1, 1], [3, 4])
		const ps = bits(...core, [0x2b7, 11], [5, 5], [1, 1], [3, 4], [0x548, 11], [1, 1])

		expect(parseAudioSpecificConfig(sbr)).toEqual({ objectType: 5, sampleRate: 48000, channels: 1 })
		expect(mp4aCodecString({ objectTypeIndication: 0x40, audioSpecificConfig: ps })).toBe("mp4a.40.29")
	})

	it("reads an escaped object type and an explicit sampling frequency (ISO/IEC 14496-3 §1.6.2.1)", () => {
		// AOT 42 (USAC): 31 then 6 bits of 10; frequency index 15 with a 24-bit rate; channel config 2.
		expect(parseAudioSpecificConfig(bits([31, 5], [10, 6], [15, 4], [44056, 24], [2, 4]))).toEqual({
			objectType: 42,
			sampleRate: 44056,
			channels: 2
		})
	})

	it("returns null for a config it cannot read", () => {
		expect(parseAudioSpecificConfig(new Uint8Array(0))).toBeNull()
		expect(parseAudioSpecificConfig(fromHex("00"))).toBeNull()
		expect(mp4aCodecString({ objectTypeIndication: 0x40, audioSpecificConfig: null })).toBeNull()
	})

	it("maps Matroska CodecPrivate and the legacy profile IDs", () => {
		expect(matroskaAacCodecString({ codecId: "A_AAC", codecPrivate: FFMPEG_AAC_LC_48K_STEREO_ASC })).toBe("mp4a.40.2")
		expect(matroskaAacCodecString({ codecId: "A_AAC/MPEG4/LC/SBR", codecPrivate: null })).toBe("mp4a.40.5")
		expect(matroskaAacCodecString({ codecId: "A_AAC/MPEG2/MAIN", codecPrivate: null })).toBe("mp4a.40.1")
		expect(matroskaAacCodecString({ codecId: "A_AAC", codecPrivate: null })).toBeNull()
		expect(matroskaAacCodecString({ codecId: "A_AAC/MPEG4/toString", codecPrivate: null })).toBeNull()
		expect(matroskaAacCodecString({ codecId: "A_AAC/MPEG4/__proto__", codecPrivate: null })).toBeNull()
	})
})

describe("other audio", () => {
	it("maps esds object types (MP4RA object type registry)", () => {
		expect(mp4aCodecString({ objectTypeIndication: 0x6b, audioSpecificConfig: null })).toBe("mp4a.6B")
		expect(mp4aCodecString({ objectTypeIndication: 0x69, audioSpecificConfig: null })).toBe("mp4a.69")
		expect(mp4aCodecString({ objectTypeIndication: 0x67, audioSpecificConfig: null })).toBe("mp4a.67")
		expect(mp4aCodecString({ objectTypeIndication: 0xa5, audioSpecificConfig: null })).toBe("ac-3")
		expect(mp4aCodecString({ objectTypeIndication: 0xa6, audioSpecificConfig: null })).toBe("ec-3")
		expect(mp4aCodecString({ objectTypeIndication: 0xad, audioSpecificConfig: null })).toBe("opus")
		expect(mp4aCodecString({ objectTypeIndication: 0x20, audioSpecificConfig: null })).toBeNull()
	})

	it("names the codecs whose string carries no config", () => {
		expect(simpleAudioCodecString("mp3")).toBe("mp4a.6B")
		expect(simpleAudioCodecString("ac3")).toBe("ac-3")
		expect(simpleAudioCodecString("eac3")).toBe("ec-3")
		expect(simpleAudioCodecString("truehd")).toBe("mlpa")
		expect(simpleAudioCodecString("flac")).toBe("flac")
		expect(simpleAudioCodecString("opus")).toBe("opus")
		expect(simpleAudioCodecString("vorbis")).toBe("vorbis")
		expect(simpleAudioCodecString("alac")).toBe("alac")
	})

	it("names DTS variants by sample entry or Matroska CodecID (ETSI TS 102 114 Annex E)", () => {
		expect(dtsCodecString("A_DTS")).toBe("dtsc")
		expect(dtsCodecString("A_DTS/EXPRESS")).toBe("dtse")
		expect(dtsCodecString("A_DTS/LOSSLESS")).toBe("dtsl")
		expect(dtsCodecString("dtsh")).toBe("dtsh")
		expect(dtsCodecString("A_DTS/UNKNOWN")).toBeNull()
	})
})

describe("dolbyVisionCodecString", () => {
	it("names the Dolby Vision sample entry for the base entry, with two-digit profile and level (Dolby Vision in ISOBMFF §2.2)", () => {
		expect(dolbyVisionCodecString({ fourcc: "hvc1", profile: 5, level: 6 })).toBe("dvh1.05.06")
		expect(dolbyVisionCodecString({ fourcc: "toString", profile: 5, level: 6 })).toBeNull()
		expect(dolbyVisionCodecString({ fourcc: "hev1", profile: 5, level: 9 })).toBe("dvhe.05.09")
		expect(dolbyVisionCodecString({ fourcc: "av01", profile: 10, level: 1 })).toBe("dav1.10.01")
		expect(dolbyVisionCodecString({ fourcc: "avc3", profile: 9, level: 5 })).toBe("dvav.09.05")
		expect(dolbyVisionCodecString({ fourcc: "vp09", profile: 5, level: 6 })).toBeNull()
	})
})

describe("codecStringAlternatives", () => {
	it("lists the spellings browsers answer differently for MP3, FLAC, Opus, AC-3 and E-AC-3, the given one first", () => {
		expect(codecStringAlternatives("mp4a.6B")).toEqual(["mp4a.6B", "mp4a.69", "mp3", "mp4a.6b"])
		expect(codecStringAlternatives("mp3")).toEqual(["mp3", "mp4a.6B", "mp4a.69", "mp4a.6b"])
		expect(codecStringAlternatives("ec-3")).toEqual(["ec-3", "mp4a.a6", "mp4a.A6"])
		expect(codecStringAlternatives("fLaC")).toEqual(["fLaC", "flac"])
		expect(codecStringAlternatives("avc1.64001f")).toEqual(["avc1.64001f"])
	})
})
