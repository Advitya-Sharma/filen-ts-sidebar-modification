import type { ProbeIndex } from "@/lib/media/probe/types"

// The contract between the page's MSE controller and the remux worker. Plain data only: it crosses a
// worker boundary.

export type TrackKind = "video" | "audio"

// How an audio track reaches the player: its packets as they are, re-encoded, or not at all.
export type AudioMode = "copy" | "transcode" | "none"

// A re-encoded audio track's output codec.
export type AudioTarget = "aac" | "opus"

// Each target's WebCodecs encoder codec, which is also its fragmented-MP4 codec string.
export const AUDIO_TARGET_CODECS: Record<AudioTarget, string> = {
	aac: "mp4a.40.2",
	opus: "opus"
}

export interface RemuxOpenParams {
	// The reading end of a range port (lib/media/rangePort) serving the file.
	port: MessagePort
	size: number
	// From the file's probe, when the caller has it; otherwise the worker probes the head it reads anyway.
	index?: ProbeIndex
	// Audio output codecs the page's MediaSource accepts in fragmented MP4, best first.
	transcodeTargets: AudioTarget[]
	// The largest fragment worth producing per track: one the page's SourceBuffer could never hold fails
	// before the worker has gathered all of it.
	maxFragmentBytes: Record<TrackKind, number>
}

export interface RemuxVideoInfo {
	id: number
	codec: string
	// What a SourceBuffer is created with, for the output (not the source) sample entry.
	mime: string
	width: number
	height: number
}

export interface RemuxAudioInfo {
	id: number
	codec: string | null
	// The copied track's MIME; null when the codec cannot go into fragmented MP4 as is.
	copyMime: string | null
	channels: number
	sampleRate: number
	language: string
	name: string | null
	isDefault: boolean
}

export interface RemuxMediaInfo {
	duration: number
	// Bits per second over the whole file, for sizing the buffer window.
	bitrate: number
	index: ProbeIndex | null
	// No index to seek with: the demuxer finds a far position only by scanning every cluster before it.
	seekScansFile: boolean
	video: RemuxVideoInfo | null
	audio: RemuxAudioInfo[]
}

export interface RemuxStartParams {
	generation: number
	time: number
	// Media time up to which output is wanted before the pump pauses.
	horizon: number
}

export interface RemuxAudioStartParams extends RemuxStartParams {
	trackId: number
	mode: AudioMode
}

export interface RemuxSeekParams extends RemuxStartParams {
	audio: { trackId: number; mode: AudioMode } | null
}

export interface RemuxVideoStart {
	// The key packet's timestamp: where the output begins and the element should be seeked to.
	keyTime: number
	// The requested time lay past what a file without an index has been scanned to; keyTime is the
	// furthest reachable point instead.
	clamped: boolean
}

export interface RemuxAudioStart {
	startTime: number
	mime: string
}

export interface RemuxSeekResult {
	video: RemuxVideoStart
	audio: RemuxAudioStart | null
}

export type RemuxEvent =
	// `init`: the bytes begin with an initialization segment.
	| { type: "segment"; kind: TrackKind; generation: number; bytes: Uint8Array<ArrayBuffer>; init: boolean }
	| { type: "ended"; kind: TrackKind; generation: number }
	// `fragmentTooLarge`: a single key-packet interval holds more than a SourceBuffer can.
	| { type: "error"; kind: TrackKind; generation: number; message: string; fragmentTooLarge: boolean }

export interface RemuxStats {
	blockReads: number
	bytesRead: number
	rangeReads: number
	readAheadBlocks: number
	segmentsOut: number
	bytesOut: number
	// Seconds spent mixing and handing audio to the encoder, and waiting on the decoder, for its CPU share.
	transcodeSeconds: number
	decodeWaitSeconds: number
	transcodedMediaSeconds: number
}

// A seek or track switch that a newer one replaced before it started. Crosses Comlink by name.
export const SUPERSEDED_ERROR = "SupersededError"

export class SupersededError extends Error {
	public constructor() {
		super("superseded by a newer seek or track switch")
		this.name = SUPERSEDED_ERROR
	}
}

export function isSuperseded(error: unknown): boolean {
	return error instanceof Error && error.name === SUPERSEDED_ERROR
}
