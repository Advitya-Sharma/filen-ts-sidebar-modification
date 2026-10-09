import { normalizeCueText, type CueTextMode } from "@filen/shared/videos"
import { BLOCK_BYTES } from "@/lib/media/blockSource"
import { inflateCapped, parseElementHeader, type EbmlElement } from "@/lib/media/probe/ebml"

// Embedded Matroska text subtitles (S_TEXT/UTF8, S_TEXT/ASS and SSA, S_TEXT/WEBVTT), taken from the cluster bytes
// playback already downloaded: every block the block cache receives passes through here, and the clusters in it
// are walked for the text tracks' blocks. Nothing is ever read for subtitles, so cues exist only for the parts of
// the file playback has read (the playhead and what it buffered ahead); there is no full cue list.

const CLUSTER_ID = 0x1f43b675
const TIMESTAMP_ID = 0xe7
const SIMPLE_BLOCK_ID = 0xa3
const BLOCK_GROUP_ID = 0xa0
const BLOCK_ID = 0xa1
const BLOCK_DURATION_ID = 0x9b

// Top-level elements a cluster walk may meet between clusters, skipped by their size.
const LEVEL_ONE_IDS: ReadonlySet<number> = new Set([
	0x1c53bb6b, // Cues
	0x1254c367, // Tags
	0x1043a770, // Chapters
	0x1941a469, // Attachments
	0x114d9b74, // SeekHead
	0x1549a966, // Info
	0x1654ae6b, // Tracks
	0xec // Void
])

// Blocks kept for a cluster walk that stopped at a block boundary; the cache keeps the bytes anyway.
const WINDOW_BLOCKS = 8
// Walks kept at once (the playhead's, one left behind by a seek).
const MAX_CURSORS = 4
// Element headers read per block arrival at most, so a hostile file can't hold the worker.
const MAX_STEPS = 50_000
// A SimpleBlock carries no duration.
export const SIMPLE_BLOCK_CUE_SECONDS = 3
const MAX_CUE_BYTES = 64 * 1024
const MAX_SEEN = 200_000
// A cluster found by searching a block is believed only within this size.
const MAX_FOUND_CLUSTER_BYTES = 1024 * 1024 * 1024

export interface TextTapTrack {
	number: number
	codecId: string
	compression: { algorithm: number; settings: Uint8Array } | "unsupported" | null
}

export interface TextTapCue {
	track: number
	start: number
	end: number
	// Cue payload text (escaped, <i>/<b>/<u> only).
	text: string
	hint: number
}

const TEXT_MODES: Readonly<Record<string, CueTextMode>> = {
	"S_TEXT/UTF8": "srt",
	"S_TEXT/ASCII": "srt",
	"S_TEXT/ASS": "ass",
	"S_TEXT/SSA": "ass",
	"S_TEXT/WEBVTT": "vtt"
}

export function isTapTextCodec(codecId: string): boolean {
	return codecId in TEXT_MODES
}

// An ASS/SSA block is "ReadOrder,Layer,Style,Name,MarginL,MarginR,MarginV,Effect,Text": the text is what follows
// the eighth comma.
export function assBlockText(line: string): string {
	let at = 0

	for (let commas = 0; commas < 8; commas++) {
		const next = line.indexOf(",", at)

		if (next === -1) {
			return line
		}

		at = next + 1
	}

	return line.slice(at)
}

interface Cursor {
	pos: number
	cluster: { end: number; timestamp: number | null } | null
	used: number
}

interface PendingCue {
	track: TextTapTrack
	start: number
	end: number
	frame: Uint8Array
}

type Step = "more" | "stop" | "lost"

export class TextTap {
	private readonly tracks: ReadonlyMap<number, TextTapTrack>
	private readonly scale: number
	private readonly emit: (cues: TextTapCue[]) => void
	private readonly blockBytes: number
	private readonly window = new Map<number, Uint8Array>()
	private readonly cursors: Cursor[] = []
	private readonly seen = new Set<string>()
	private pending: PendingCue[] = []
	private uses = 0
	private closed = false

	public constructor(params: {
		tracks: readonly TextTapTrack[]
		// Nanoseconds per timestamp tick.
		timestampScale: number
		emit: (cues: TextTapCue[]) => void
		blockBytes?: number
	}) {
		this.tracks = new Map(
			params.tracks
				.filter(track => isTapTextCodec(track.codecId) && track.compression !== "unsupported")
				.map(track => [track.number, track])
		)
		this.scale = params.timestampScale > 0 ? params.timestampScale / 1e9 : 1e-3
		this.emit = params.emit
		this.blockBytes = params.blockBytes ?? BLOCK_BYTES
	}

	public get active(): boolean {
		return this.tracks.size > 0 && !this.closed
	}

	public close(): void {
		this.closed = true
		this.window.clear()
		this.cursors.length = 0
		this.pending = []
	}

	// A block landed in the cache. Never throws: a malformed file only ends the walks over it.
	public addBlock(index: number, bytes: Uint8Array): void {
		if (!this.active) {
			return
		}

		try {
			this.window.delete(index)
			this.window.set(index, bytes)

			while (this.window.size > WINDOW_BLOCKS) {
				const oldest = this.window.keys().next()

				if (oldest.done === true) {
					break
				}

				this.window.delete(oldest.value)
			}

			const start = index * this.blockBytes
			const end = start + bytes.length

			if (!this.cursors.some(cursor => cursor.pos >= start - 16 && cursor.pos < end)) {
				const found = this.findCluster(start, bytes)

				if (found !== null) {
					this.addCursor(found)
				}
			}

			let steps = 0

			for (const cursor of [...this.cursors]) {
				const result = this.advance(cursor, MAX_STEPS - steps)

				steps += result.steps

				if (result.state === "lost") {
					this.dropCursor(cursor)
				}
			}

			this.dropDuplicateCursors()
			this.drain()
		} catch {
			// A walk that throws on hostile bytes just stops; playback is not this module's concern.
		}
	}

	private addCursor(pos: number): void {
		this.cursors.push({ pos, cluster: null, used: ++this.uses })

		while (this.cursors.length > MAX_CURSORS) {
			let oldest: Cursor | undefined

			for (const cursor of this.cursors) {
				if (oldest === undefined || cursor.used < oldest.used) {
					oldest = cursor
				}
			}

			if (oldest === undefined) {
				break
			}

			this.dropCursor(oldest)
		}
	}

	private dropCursor(cursor: Cursor): void {
		const at = this.cursors.indexOf(cursor)

		if (at !== -1) {
			this.cursors.splice(at, 1)
		}
	}

	private dropDuplicateCursors(): void {
		const seen = new Set<number>()

		for (const cursor of [...this.cursors]) {
			if (seen.has(cursor.pos)) {
				this.dropCursor(cursor)
			} else {
				seen.add(cursor.pos)
			}
		}
	}

	// `length` bytes at `pos` when every block they lie in is in the window; null otherwise.
	private bytesAt(pos: number, length: number): Uint8Array | null {
		const first = Math.floor(pos / this.blockBytes)
		const last = Math.floor((pos + Math.max(length, 1) - 1) / this.blockBytes)
		const head = this.window.get(first)

		if (head === undefined) {
			return null
		}

		const offset = pos - first * this.blockBytes

		if (first === last) {
			return offset + length <= head.length ? head.subarray(offset, offset + length) : null
		}

		const out = new Uint8Array(length)
		let written = 0

		for (let index = first; index <= last; index++) {
			const block = this.window.get(index)

			if (block === undefined) {
				return null
			}

			const from = index === first ? offset : 0
			const take = Math.min(block.length - from, length - written)

			out.set(block.subarray(from, from + take), written)
			written += take
		}

		return written === length ? out : null
	}

	// The element header at `pos`: null when its bytes aren't here yet, "invalid" when none can start there.
	private header(pos: number): EbmlElement | "invalid" | null {
		for (const length of [12, 8, 4, 2]) {
			const bytes = this.bytesAt(pos, length)

			if (bytes === null) {
				continue
			}

			const parsed = parseElementHeader(bytes, 0)

			if (parsed === "invalid") {
				return "invalid"
			}

			if (parsed === "short") {
				return null
			}

			return {
				id: parsed.id,
				start: pos,
				dataStart: pos + parsed.dataStart,
				end: parsed.end === null ? null : pos + parsed.end
			}
		}

		return null
	}

	// A cluster's start within a freshly landed block nobody walks yet (playback jumped here): its ID followed by a
	// plausible size and a Timestamp as its first child.
	private findCluster(start: number, bytes: Uint8Array): number | null {
		for (let i = bytes.indexOf(0x1f); i !== -1 && i + 4 < bytes.length; i = bytes.indexOf(0x1f, i + 1)) {
			if (bytes[i + 1] !== 0x43 || bytes[i + 2] !== 0xb6 || bytes[i + 3] !== 0x75) {
				continue
			}

			const header = this.header(start + i)

			if (header === null || header === "invalid" || header.id !== CLUSTER_ID) {
				continue
			}

			if (header.end !== null && header.end - header.dataStart > MAX_FOUND_CLUSTER_BYTES) {
				continue
			}

			const child = this.header(header.dataStart)

			if (child !== null && child !== "invalid" && child.id === TIMESTAMP_ID) {
				return start + i
			}
		}

		return null
	}

	private advance(cursor: Cursor, budget: number): { state: Step; steps: number } {
		let steps = 0

		while (steps < budget) {
			steps++

			const cluster = cursor.cluster

			if (cluster !== null && cursor.pos >= cluster.end) {
				cursor.cluster = null

				continue
			}

			const header = this.header(cursor.pos)

			if (header === null) {
				return { state: "stop", steps }
			}

			if (header === "invalid") {
				return { state: "lost", steps }
			}

			if (header.id === CLUSTER_ID) {
				// A cluster of unknown size ends where the next one starts.
				cursor.cluster = { end: header.end ?? Number.POSITIVE_INFINITY, timestamp: null }
				cursor.pos = header.dataStart
				cursor.used = ++this.uses

				continue
			}

			if (cluster === null) {
				if (!LEVEL_ONE_IDS.has(header.id) || header.end === null) {
					return { state: "lost", steps }
				}

				cursor.pos = header.end

				continue
			}

			if (header.end === null) {
				return { state: "lost", steps }
			}

			const state = this.clusterChild(cluster, { id: header.id, start: header.start, dataStart: header.dataStart, end: header.end })

			if (state !== "more") {
				return { state, steps }
			}

			cursor.pos = header.end
			cursor.used = ++this.uses
		}

		return { state: "stop", steps }
	}

	private clusterChild(cluster: NonNullable<Cursor["cluster"]>, element: EbmlElement & { end: number }): Step {
		const size = element.end - element.dataStart

		switch (element.id) {
			case TIMESTAMP_ID: {
				const value = size > 0 && size <= 8 ? this.bytesAt(element.dataStart, size) : null

				if (value === null) {
					return size > 0 && size <= 8 ? "stop" : "lost"
				}

				let timestamp = 0

				for (const byte of value) {
					timestamp = timestamp * 256 + byte
				}

				cluster.timestamp = timestamp

				return "more"
			}
			case SIMPLE_BLOCK_ID:
				return this.block(cluster, element.dataStart, element.end, null)
			case BLOCK_GROUP_ID:
				return this.group(cluster, element)
			default:
				return "more"
		}
	}

	private group(cluster: NonNullable<Cursor["cluster"]>, group: EbmlElement & { end: number }): Step {
		let block: { start: number; end: number } | null = null
		let duration: number | null = null

		for (let pos = group.dataStart; pos < group.end;) {
			const child = this.header(pos)

			if (child === null) {
				return "stop"
			}

			if (child === "invalid" || child.end === null || child.end > group.end) {
				return "lost"
			}

			if (child.id === BLOCK_ID) {
				block = { start: child.dataStart, end: child.end }
			} else if (child.id === BLOCK_DURATION_ID) {
				const size = child.end - child.dataStart
				const value = size > 0 && size <= 8 ? this.bytesAt(child.dataStart, size) : null

				if (value === null) {
					return "stop"
				}

				duration = 0

				for (const byte of value) {
					duration = duration * 256 + byte
				}
			}

			pos = child.end
		}

		return block === null ? "more" : this.block(cluster, block.start, block.end, duration)
	}

	// A Block or SimpleBlock: track number (EBML vint), int16 relative timestamp, flags, then the frame.
	private block(cluster: NonNullable<Cursor["cluster"]>, start: number, end: number, duration: number | null): Step {
		const head = this.bytesAt(start, Math.min(end - start, 12))

		if (head === null) {
			return "stop"
		}

		const first = head[0] ?? 0
		let length = 0

		while (length < 8 && (first & (0x80 >> length)) === 0) {
			length++
		}

		length++

		if (length > 8 || head.length < length + 3) {
			return "more"
		}

		let number = first & (0xff >> length)

		for (let i = 1; i < length; i++) {
			number = number * 256 + (head[i] ?? 0)
		}

		const track = this.tracks.get(number)

		if (track === undefined || cluster.timestamp === null) {
			return "more"
		}

		const relative = (((head[length] ?? 0) << 24) >> 16) | (head[length + 1] ?? 0)
		const flags = head[length + 2] ?? 0
		const frameStart = start + length + 3
		const frameLength = end - frameStart

		// Laced subtitle blocks don't occur in practice; one is skipped rather than misread.
		if ((flags & 0x06) !== 0 || frameLength <= 0 || frameLength > MAX_CUE_BYTES) {
			return "more"
		}

		const frame = this.bytesAt(frameStart, frameLength)

		if (frame === null) {
			return "stop"
		}

		const startSeconds = (cluster.timestamp + relative) * this.scale
		const endSeconds = duration === null ? startSeconds + SIMPLE_BLOCK_CUE_SECONDS : startSeconds + duration * this.scale
		const key = `${String(number)}:${String(Math.round(startSeconds * 1000))}`

		if (!this.seen.has(key) && this.seen.size < MAX_SEEN && startSeconds >= 0) {
			this.seen.add(key)
			this.pending.push({ track, start: startSeconds, end: Math.max(startSeconds, endSeconds), frame: frame.slice() })
		}

		return "more"
	}

	// Decodes what was found (inflating where the track is compressed) and hands it on as one batch.
	private drain(): void {
		if (this.pending.length === 0) {
			return
		}

		const batch = this.pending

		this.pending = []

		void Promise.all(batch.map(cue => this.decode(cue))).then(cues => {
			const out = cues.filter(cue => cue !== null)

			if (out.length > 0 && !this.closed) {
				this.emit(out)
			}
		})
	}

	private async decode(cue: PendingCue): Promise<TextTapCue | null> {
		const compression = cue.track.compression
		let bytes: Uint8Array | null = cue.frame

		if (compression !== null && compression !== "unsupported") {
			if (compression.algorithm === 3) {
				const joined = new Uint8Array(compression.settings.length + bytes.length)

				joined.set(compression.settings)
				joined.set(bytes, compression.settings.length)
				bytes = joined
			} else {
				bytes = await inflateCapped(bytes, MAX_CUE_BYTES)
			}
		}

		if (bytes === null) {
			return null
		}

		const mode = TEXT_MODES[cue.track.codecId] ?? "srt"
		const raw = new TextDecoder().decode(bytes)
		const normalized = normalizeCueText(mode === "ass" ? assBlockText(raw) : raw, mode)

		return normalized.text.length === 0 || normalized.drawing
			? null
			: { track: cue.track.number, start: cue.start, end: cue.end, text: normalized.text, hint: normalized.hint }
	}
}
