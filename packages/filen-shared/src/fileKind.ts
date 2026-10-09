import { extensionStart } from "./fileExtensions"
import { extensionForFileName, extensionForMime } from "./fileType"
import { MIN_TRANSPORT_STREAM_BYTES } from "./videos/transportStream"

export { MIN_TRANSPORT_STREAM_BYTES }

// What kind of file something is, decided in one place for both apps, from what is at hand without a
// read: its extension, a content verdict an earlier read left behind, its stored MIME type, a well-known
// file name. The kind itself is each app's (what it can show differs by platform): the app passes
// `kindOf`, its own extension table, and this decides which extension that table is asked about.
//
// 1. A known extension that means one thing wins.
// 2. An extension that means two (`.ts` TypeScript or an MPEG transport stream; `.webm`, `.ogg`, `.mp4`,
//    `.mkv` and their audio siblings) is settled by a content verdict for the file when one exists, else
//    by a MIME type that says audio or video, else by its usual meaning. A MIME type never makes a `.ts` a
//    video: browsers upload TypeScript as video/mp2t too.
// 3. No or an unknown extension: a well-known name (a Makefile is code whatever its type says), else its MIME
//    type, else unknown.
// Whatever a verdict or a MIME type points at, an extension the caller's table does not know is passed over
// for the next candidate: an app that has no audio Matroska keeps showing an `.mkv` as video.

// What a read of the file's content found.
export type ContentVerdict = "transportStream" | "notTransportStream" | "video" | "audio"

export type FileKindSource = "extension" | "sniff" | "mime" | "name" | "default"

export interface FileKindInput {
	name: string
	mime?: string | null | undefined
	// Bytes; decides whether a content read could still change the kind (`sniffable`).
	size?: number | bigint | undefined
	verdict?: ContentVerdict | null | undefined
}

export interface FileKind<K> {
	kind: K | null
	// The extension `kindOf` was asked about ("" when there was none).
	extension: string
	source: FileKindSource
	// The extension alone does not settle the kind: a content verdict may change it.
	ambiguous: boolean
	// Ambiguous, without a verdict yet, and a file a content read could still turn into another kind (a
	// `.ts` large enough to be a transport stream). A caller reading the file anyway may record one.
	sniffable: boolean
}

interface Ambiguity {
	// The usual meaning.
	fallback: string
	audio?: string
	video?: string
	transportStream?: string
	// Whether a MIME type may settle it.
	mimeHint: boolean
}

const TRANSPORT_STREAM = "m2ts"

const AMBIGUOUS: ReadonlyMap<string, Ambiguity> = new Map([
	["ts", { fallback: "ts", transportStream: TRANSPORT_STREAM, mimeHint: false }],
	["mts", { fallback: "mts", transportStream: TRANSPORT_STREAM, mimeHint: false }],
	["ogg", { fallback: "ogg", audio: "oga", video: "ogv", mimeHint: true }],
	["oga", { fallback: "oga", audio: "oga", video: "ogv", mimeHint: true }],
	["ogv", { fallback: "ogv", audio: "oga", video: "ogv", mimeHint: true }],
	["webm", { fallback: "webm", audio: "weba", video: "webm", mimeHint: true }],
	["weba", { fallback: "weba", audio: "weba", video: "webm", mimeHint: true }],
	["mp4", { fallback: "mp4", audio: "m4a", video: "mp4", mimeHint: true }],
	["m4a", { fallback: "m4a", audio: "m4a", video: "mp4", mimeHint: true }],
	["mkv", { fallback: "mkv", audio: "mka", video: "mkv", mimeHint: true }],
	["mka", { fallback: "mka", audio: "mka", video: "mkv", mimeHint: true }]
])

function lowerName(name: string): string {
	const trimmed = name.trim()

	return trimmed.slice(trimmed.lastIndexOf("/") + 1).toLowerCase()
}

export function isAmbiguousExtension(extension: string): boolean {
	return AMBIGUOUS.has(extension)
}

function verdictExtension(ambiguity: Ambiguity, verdict: ContentVerdict): string | undefined {
	switch (verdict) {
		case "transportStream":
			return ambiguity.transportStream
		case "notTransportStream":
			return ambiguity.transportStream === undefined ? undefined : ambiguity.fallback
		case "audio":
			return ambiguity.audio
		case "video":
			return ambiguity.video
	}
}

function mimeExtension(ambiguity: Ambiguity, mime: string | null | undefined): string | undefined {
	if (!ambiguity.mimeHint || mime === null || mime === undefined) {
		return undefined
	}

	const type = mime.trim().toLowerCase()

	return type.startsWith("audio/") ? ambiguity.audio : type.startsWith("video/") ? ambiguity.video : undefined
}

// Whether a content read could still turn the file into another kind.
function sniffable(ambiguity: Ambiguity, base: string, extension: string, size: number | bigint | undefined): boolean {
	if (ambiguity.transportStream === undefined) {
		return true
	}

	const stem = base.slice(0, base.length - extension.length - 1)

	return size !== undefined && Number(size) >= MIN_TRANSPORT_STREAM_BYTES && !stem.endsWith(".d")
}

export function resolveFileKind<K>(input: FileKindInput, kindOf: (extension: string) => K | null): FileKind<K> {
	const base = lowerName(input.name)
	const dot = extensionStart(base)
	const own = dot === -1 ? "" : base.slice(dot + 1)
	const ambiguity = AMBIGUOUS.get(own)

	if (ambiguity !== undefined) {
		const verdict = input.verdict ?? null
		const byVerdict = verdict === null ? undefined : verdictExtension(ambiguity, verdict)
		const byMime = mimeExtension(ambiguity, input.mime)
		const known = (extension: string | undefined): extension is string => extension !== undefined && kindOf(extension) !== null
		const extension = known(byVerdict) ? byVerdict : known(byMime) ? byMime : ambiguity.fallback

		return {
			kind: kindOf(extension),
			extension,
			source: known(byVerdict) ? "sniff" : known(byMime) ? "mime" : "default",
			ambiguous: true,
			sniffable: byVerdict === undefined && sniffable(ambiguity, base, own, input.size)
		}
	}

	const settled = (extension: string, source: FileKindSource): FileKind<K> => ({
		kind: kindOf(extension),
		extension,
		source,
		ambiguous: false,
		sniffable: false
	})

	if (own !== "" && kindOf(own) !== null) {
		return settled(own, "extension")
	}

	const byName = extensionForFileName(base)

	if (byName !== null) {
		return settled(byName, "name")
	}

	const byMime = extensionForMime(input.mime)

	if (byMime !== null && kindOf(byMime) !== null) {
		return settled(byMime, "mime")
	}

	return settled(own, "default")
}

// Whether `verdict` would change what the file resolves to, or settle a `.ts` a content read would otherwise
// still have to: worth keeping only then.
export function verdictChangesKind<K>(
	input: Omit<FileKindInput, "verdict">,
	verdict: ContentVerdict,
	kindOf: (extension: string) => K | null
): boolean {
	const without = resolveFileKind({ ...input, verdict: null }, kindOf)

	if (!without.ambiguous) {
		return false
	}

	if (verdict === "notTransportStream") {
		return without.sniffable
	}

	return resolveFileKind({ ...input, verdict }, kindOf).extension !== without.extension
}
