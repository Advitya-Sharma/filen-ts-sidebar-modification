import { describe, it, expect, afterEach, vi } from "vitest"

// Each platform's real extension tables decide what an ambiguous extension resolves to: a sibling the
// platform cannot play is passed over for the usual meaning.

type Row = [name: string, mime: string | null, type: string, extension: string]

const BOTH: Row[] = [
	["voice.mp4", "audio/mp4", "audio", "m4a"],
	["clip.mp4", "video/mp4", "video", "mp4"],
	["clip.m4a", "video/mp4", "video", "mp4"],
	["song.m4a", "audio/mp4", "audio", "m4a"],
	["main.ts", "video/mp2t", "code", "ts"],
	["camera.mts", "video/mp2t", "unknown", "mts"],
	// None of an extension's meanings playable here: read like an unknown extension, by its type.
	["notes.mts", "text/plain", "text", "txt"],
	["module.mts", "application/typescript", "code", "ts"]
]

const IOS: Row[] = [
	...BOTH,
	// A transport stream: AVPlayer plays one only as HLS segments.
	["recording", "video/mp2t", "unknown", "m2ts"],
	["song.ogg", "audio/ogg", "unknown", "ogg"],
	["song.oga", "audio/ogg", "unknown", "oga"],
	["voice.ogg", "audio/opus", "audio", "opus"],
	["voice.oga", "audio/opus", "audio", "opus"],
	["talk.webm", "audio/webm", "unknown", "webm"],
	["film.mkv", "video/x-matroska", "unknown", "mkv"]
]

const ANDROID: Row[] = [
	...BOTH,
	// A stream by its type alone: no player until a read found one Media3 plays.
	["recording", "video/mp2t", "unknown", "m2ts"],
	["clip.m2ts", null as never, "unknown", "m2ts"],
	["song.ogg", "audio/ogg", "audio", "ogg"],
	["song.ogg", "video/ogg", "audio", "ogg"],
	["film.ogv", "audio/ogg", "audio", "oga"],
	["talk.webm", "audio/webm", "video", "webm"],
	["album.mka", "video/x-matroska", "video", "mkv"],
	["album.mka", "audio/x-matroska", "unknown", "mka"],
	["concert.mkv", "audio/x-matroska", "video", "mkv"],
	["voice.ogg", "audio/opus", "audio", "ogg"],
	["voice.weba", "audio/ogg", "audio", "ogg"]
]

// A fresh module graph per platform: @/constants reads Platform.OS once, at load.
async function loadFor(os: "ios" | "android") {
	vi.resetModules()

	const { Platform } = await import("react-native")

	Platform.OS = os

	return await import("@/lib/previewType")
}

describe.each([
	["ios", IOS],
	["android", ANDROID]
] as const)("ambiguous extensions on %s", (os, rows) => {
	afterEach(() => {
		vi.resetModules()
	})

	it.each(rows)("%s (%s) is %s by .%s", async (name, mime, type, extension) => {
		const { getPreviewType, fileTypeExtension } = await loadFor(os)

		expect(getPreviewType(name, mime)).toBe(type)
		expect(fileTypeExtension(name, mime)).toBe(extension)
	})
})

// What a content verdict makes a file on each platform: Media3 plays a progressive transport stream, AVPlayer
// only HLS segments; an .mp4 holding audio alone is audio on both.
type VerdictRow = [name: string, verdict: "transportStream" | "notTransportStream" | "audio" | "video", type: string, extension: string]

const VERDICTS_BOTH: VerdictRow[] = [
	["capture.ts", "notTransportStream", "code", "ts"],
	["voice.mp4", "audio", "audio", "m4a"],
	["clip.m4a", "video", "video", "mp4"]
]

describe.each([
	[
		"ios",
		[
			...VERDICTS_BOTH,
			["capture.ts", "transportStream", "unknown", "m2ts"],
			["camera.mts", "transportStream", "unknown", "m2ts"],
			["clip.m2ts", "transportStream", "unknown", "m2ts"]
		]
	],
	[
		"android",
		[
			...VERDICTS_BOTH,
			["capture.ts", "transportStream", "video", "m2ts"],
			["camera.mts", "transportStream", "video", "m2ts"],
			["clip.m2ts", "transportStream", "video", "m2ts"],
			["clip.m2ts", "notTransportStream", "unknown", "m2ts"]
		]
	]
] as const)("content verdicts on %s", (os, rows) => {
	afterEach(() => {
		vi.resetModules()
	})

	it.each(rows as readonly VerdictRow[])("%s with %s is %s by .%s", async (name, verdict, type, extension) => {
		const { getPreviewType, fileTypeExtension } = await loadFor(os)

		expect(getPreviewType(name, null, verdict)).toBe(type)
		expect(fileTypeExtension(name, null, verdict)).toBe(extension)
	})
})

describe.each([
	["ios", false],
	["android", true]
] as const)("a stream read on %s", (os, reads) => {
	afterEach(() => {
		vi.resetModules()
	})

	it(`${reads ? "is made" : "is skipped"} for an .mts, an .m2ts or a video/mp2t file`, async () => {
		const { isUncheckedTransportStream } = await loadFor(os)
		const MIB = 1024 * 1024

		expect(isUncheckedTransportStream("camera.mts", null, 2 * MIB, null)).toBe(reads)
		expect(isUncheckedTransportStream("clip.m2ts", null, 10, null)).toBe(reads)
		expect(isUncheckedTransportStream("recording", "video/mp2t", 2 * MIB, null)).toBe(reads)
		expect(isUncheckedTransportStream("clip.m2ts", null, 2 * MIB, "transportStream")).toBe(false)
		expect(isUncheckedTransportStream("capture.ts", null, 2 * MIB, null)).toBe(true)
	})

	it(`keeps what a read found in an .m2ts ${reads ? "" : "never"}`, async () => {
		const { recordContentVerdict } = await loadFor(os)
		const verdicts = (await import("@/lib/fileKindVerdicts")).default

		recordContentVerdict({ uuid: "m2ts", name: "clip.m2ts", mime: null, size: 10 }, "transportStream")
		recordContentVerdict({ uuid: "m2ts-not", name: "clip.m2ts", mime: null, size: 10 }, "notTransportStream")

		expect(verdicts.get("m2ts")).toBe(reads ? "transportStream" : undefined)
		// Kept where the gate reads an .m2ts, so it is not read again.
		expect(verdicts.get("m2ts-not")).toBe(reads ? "notTransportStream" : undefined)
	})
})
