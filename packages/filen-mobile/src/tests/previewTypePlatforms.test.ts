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
	["recording", "video/mp2t", "unknown", ""],
	// None of an extension's meanings playable here: read like an unknown extension, by its type.
	["notes.mts", "text/plain", "text", "txt"],
	["module.mts", "application/typescript", "code", "ts"]
]

const IOS: Row[] = [
	...BOTH,
	["song.ogg", "audio/ogg", "unknown", "ogg"],
	["song.oga", "audio/ogg", "unknown", "oga"],
	["voice.ogg", "audio/opus", "audio", "opus"],
	["voice.oga", "audio/opus", "audio", "opus"],
	["talk.webm", "audio/webm", "unknown", "webm"],
	["film.mkv", "video/x-matroska", "unknown", "mkv"]
]

const ANDROID: Row[] = [
	...BOTH,
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
