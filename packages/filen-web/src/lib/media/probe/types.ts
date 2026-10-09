// What a video file's container says about it, read from its headers without decoding anything. Kept
// JSON-serialisable and small (a few hundred bytes) so it can be cached per file uuid.

// card: badges (codecs, size, duration, audio list) in at most CARD_BLOCK_READS downloads.
// full: everything, for a detail page or playback. cover: only where an embedded cover sits.
export type ProbeTier = "card" | "full" | "cover"

export type MediaCodec =
	| "aac"
	| "mp3"
	| "ac3"
	| "eac3"
	| "dts"
	| "truehd"
	| "opus"
	| "vorbis"
	| "flac"
	| "alac"
	| "h264"
	| "hevc"
	| "av1"
	| "vp9"
	| "vp8"
	| "theora"
	| "mpeg4"

export type ProbeContainer = "mp4" | "mov" | "mkv" | "webm"

// Where a player finds keyframes: the MP4 sample tables, Matroska Cues, a fragmented MP4's mfra, or nowhere.
export type ProbeIndex = "moov" | "cues" | "mfra" | "none"

export type ProbeTransfer = "sdr" | "pq" | "hlg"

export interface ProbeVideo {
	codec: MediaCodec | null
	// The MP4 sample entry type; for Matroska the type a remux would write (avc1, hvc1, av01, vp09), else the CodecID.
	fourcc: string
	codecString: string | null
	w: number
	h: number
	// Pixel aspect ratio, width over height.
	par: number
	// Clockwise degrees the decoded picture is turned for display.
	rotation: number
	bitDepth: number | null
	transfer: ProbeTransfer
	dv?: { profile: number; compat: number }
}

export interface ProbeAudio {
	id: number
	codec: MediaCodec | null
	codecString: string | null
	ch: number | null
	rate: number | null
	// BCP 47 when the container has it, else its legacy ISO 639-2 code; null for undetermined.
	lang: string | null
	name: string | null
	def: boolean
	forced: boolean
	commentary: boolean
}

export interface ProbeSubtitle {
	id: number
	// Matroska CodecID or MP4 sample entry type.
	codecId: string
	lang: string | null
	name: string | null
	def: boolean
	forced: boolean
	hi: boolean
}

export interface ProbeChapter {
	t: number
	end?: number
	title: string | null
	skipType?: number
}

// The image's own bytes inside the file; never read by the probe.
export interface ProbeCover {
	offset: number
	size: number
	mime: string
}

export interface VideoProbe {
	v: 1
	container: ProbeContainer
	durationSec: number | null
	index: ProbeIndex | null
	video: ProbeVideo | null
	audio: ProbeAudio[]
	subs: ProbeSubtitle[]
	chapters: ProbeChapter[]
	cover: ProbeCover | null
	// The tier's read budget ran out or a structure was malformed before everything was found.
	partial: boolean
}

export type ParsedTrack =
	| { kind: "video"; enabled: boolean; video: ProbeVideo }
	| { kind: "audio"; enabled: boolean; audio: ProbeAudio }
	| { kind: "subtitle"; enabled: boolean; sub: ProbeSubtitle }

// A container walker's result, before a tier's record is cut from it.
export interface ParsedContainer {
	container: ProbeContainer
	durationSec: number | null
	index: ProbeIndex | null
	// In file order. `tracksFound` is false when the track list was never located.
	tracks: ParsedTrack[]
	tracksFound: boolean
	chapters: ProbeChapter[]
	cover: ProbeCover | null
	partial: boolean
}

export const MAX_CHAPTERS = 64
