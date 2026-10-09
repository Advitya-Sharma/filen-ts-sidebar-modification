import { describe, expect, it } from "vitest"
import { resolveFileKind, verdictChangesKind, type ContentVerdict, type FileKindSource } from "@filen/shared"

// An app table as web's preview categories read it.
const KINDS: Readonly<Record<string, string>> = {
	ts: "code",
	py: "code",
	txt: "text",
	md: "markdown",
	sh: "code",
	json: "code",
	mp4: "video",
	m4v: "video",
	mov: "video",
	webm: "video",
	mkv: "video",
	m2ts: "video",
	mp3: "audio",
	m4a: "audio",
	ogg: "audio",
	oga: "audio",
	weba: "audio",
	gz: "archive",
	jpg: "image"
}

function kindOf(extension: string): string | null {
	return KINDS[extension] ?? null
}

const MIB = 1024 * 1024

type Row = [
	name: string,
	mime: string | null,
	size: number,
	verdict: ContentVerdict | null,
	kind: string | null,
	extension: string,
	source: FileKindSource,
	ambiguous: boolean,
	sniffable: boolean
]

const ROWS: Row[] = [
	// TypeScript uploads as video/mp2t: the MIME type never makes a .ts a video.
	["index.ts", "video/mp2t", 4_000, null, "code", "ts", "default", true, false],
	["bundle.ts", "video/mp2t", 4 * MIB, null, "code", "ts", "default", true, true],
	["bundle.ts", "video/mp2t", 4 * MIB, "notTransportStream", "code", "ts", "sniff", true, false],
	["lib.d.ts", "video/mp2t", 4 * MIB, null, "code", "ts", "default", true, false],
	// A real transport stream, once a read said so.
	["clip.ts", "video/mp2t", 40 * MIB, "transportStream", "video", "m2ts", "sniff", true, false],
	["CLIP.TS", null, 40 * MIB, "transportStream", "video", "m2ts", "sniff", true, false],
	["camera.mts", "video/mp2t", 40 * MIB, null, null, "mts", "default", true, true],
	["module.mts", null, 40 * MIB, "notTransportStream", null, "mts", "default", true, false],
	["camera.mts", "video/mp2t", 40 * MIB, "transportStream", "video", "m2ts", "sniff", true, false],
	["movie.m2ts", "video/mp2t", 40 * MIB, null, "video", "m2ts", "extension", false, false],
	// Audio or video containers: a verdict, else the MIME type's family, else the usual meaning.
	["talk.webm", "audio/webm", MIB, null, "audio", "weba", "mime", true, true],
	["talk.webm", "video/webm", MIB, null, "video", "webm", "mime", true, true],
	["talk.webm", null, MIB, null, "video", "webm", "default", true, true],
	["talk.webm", "video/webm", MIB, "audio", "audio", "weba", "sniff", true, false],
	["song.ogg", "audio/ogg", MIB, null, "audio", "oga", "mime", true, true],
	// The table has no .ogv: the MIME hint is passed over for the usual meaning.
	["film.ogg", "video/ogg", MIB, null, "audio", "ogg", "default", true, true],
	["song.ogg", null, MIB, null, "audio", "ogg", "default", true, true],
	["voice.mp4", "audio/mp4", MIB, null, "audio", "m4a", "mime", true, true],
	["voice.m4a", "video/mp4", MIB, "audio", "audio", "m4a", "sniff", true, false],
	["film.mkv", "video/x-matroska", MIB, null, "video", "mkv", "mime", true, true],
	["album.mka", "audio/x-matroska", MIB, null, null, "mka", "default", true, true],
	// An audio-only .mkv on a table without .mka stays the video it reads as.
	["concert.mkv", "video/x-matroska", MIB, "audio", "video", "mkv", "mime", true, false],
	// No or an unknown extension: a specific MIME type, then a well-known name, then a generic type.
	["recording", "video/mp2t", 40 * MIB, null, "video", "m2ts", "mime", false, false],
	// A well-known name wins over any MIME type.
	["Makefile", "video/mp4", 100, null, null, "makefile", "name", false, false],
	["recording", "application/octet-stream", MIB, null, null, "", "default", false, false],
	["recording", "", MIB, null, null, "", "default", false, false],
	["Makefile", "text/plain", 100, null, null, "makefile", "name", false, false],
	["LICENSE", "application/octet-stream", 100, null, "text", "txt", "name", false, false],
	["notes", "text/plain", 100, null, "text", "txt", "mime", false, false],
	["notes", "application/octet-stream", 100, null, null, "", "default", false, false],
	["data.weird", "application/json", 100, null, "code", "json", "mime", false, false],
	// A known extension that means one thing wins, whatever the type says; the last extension counts.
	["photo.JPG", "application/octet-stream", 100, null, "image", "jpg", "extension", false, false],
	["backup.tar.gz", "application/gzip", 100, null, "archive", "gz", "extension", false, false],
	["notes.md", "text/plain", 100, null, "markdown", "md", "extension", false, false]
]

describe("resolveFileKind", () => {
	it.each(ROWS)(
		"%s (%s, %d B, verdict %s) is %s by .%s from %s",
		(name, mime, size, verdict, kind, extension, source, ambiguous, sniffable) => {
			expect(resolveFileKind({ name, mime, size, verdict }, kindOf)).toEqual({ kind, extension, source, ambiguous, sniffable })
		}
	)

	it("says nothing is sniffable for a .ts of unknown size", () => {
		expect(resolveFileKind({ name: "clip.ts" }, kindOf).sniffable).toBe(false)
	})
})

// The mobile app's table today: no .weba, .oga, .ogv, .mka or .m2ts.
const MOBILE: Readonly<Record<string, string>> = {
	mp4: "video",
	mov: "video",
	webm: "video",
	mkv: "video",
	mp3: "audio",
	m4a: "audio",
	ogg: "audio",
	wav: "audio",
	ts: "code"
}

function mobileKindOf(extension: string): string | null {
	return MOBILE[extension] ?? null
}

describe("resolveFileKind on a table without the audio and video siblings", () => {
	it.each([
		["talk.webm", "audio/webm", null, "video", "webm"],
		["concert.mkv", "audio/x-matroska", "audio", "video", "mkv"],
		["film.ogg", "video/ogg", null, "audio", "ogg"],
		["voice.mp4", "audio/mp4", null, "audio", "m4a"],
		["clip.ts", "video/mp2t", "transportStream", "code", "ts"]
	] as const)("%s (%s, verdict %s) stays %s by .%s", (name, mime, verdict, kind, extension) => {
		expect(resolveFileKind({ name, mime, verdict }, mobileKindOf)).toMatchObject({ kind, extension })
	})
})

describe("verdictChangesKind", () => {
	it("keeps only a verdict that moves an ambiguous file to another kind or settles a .ts", () => {
		expect(verdictChangesKind({ name: "clip.ts", size: 40 * MIB }, "transportStream", kindOf)).toBe(true)
		expect(verdictChangesKind({ name: "server.ts", size: 40 * MIB }, "notTransportStream", kindOf)).toBe(true)
		expect(verdictChangesKind({ name: "index.ts", size: 4_000 }, "notTransportStream", kindOf)).toBe(false)
		expect(verdictChangesKind({ name: "lib.d.ts", size: 40 * MIB }, "notTransportStream", kindOf)).toBe(false)
		expect(verdictChangesKind({ name: "talk.webm", mime: "video/webm" }, "audio", kindOf)).toBe(true)
		expect(verdictChangesKind({ name: "talk.webm", mime: "video/webm" }, "video", kindOf)).toBe(false)
		expect(verdictChangesKind({ name: "film.mp4" }, "video", kindOf)).toBe(false)
		expect(verdictChangesKind({ name: "photo.jpg" }, "audio", kindOf)).toBe(false)
	})
})
