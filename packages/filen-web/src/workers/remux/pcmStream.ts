import { isSourceReadError, type RemuxStats } from "@/lib/media/remuxTypes"
import { AudioUnsupportedError, type AudioInput, type AudioInputInfo } from "@/workers/remux/audioInput"
import type { SampleItem, SampleReader } from "@/workers/remux/audioTranscode"
import { LinearResampler } from "@/workers/remux/audioMix.logic"
import { layoutFor } from "@/workers/remux/channelLayout.logic"
import { decodedChannelMask } from "@/workers/remux/ffmpegDecoder"
import { outputPlanes, PcmAssembler } from "@/workers/remux/pcmAssembler.logic"
import { frameAtOffset, WAV_HEADER_BYTES, wavHeader, wavLayout, type WavLayout } from "@/workers/remux/wav.logic"

// An audio file decoded into a WAV stream of fixed size, served by byte range: a read continues the decode
// that ended where it starts, or restarts the decoder at the sample its offset names. The element's reads pace
// the decoding (the service worker reads such a stream with no window ahead), so nothing is decoded ahead of
// what it asks for beyond the block a read ends in.

// A decoder restarted at a seek begins this far before the requested sample: lossy codecs' first frame
// after a restart lacks the overlap of the one before it.
export const RESTART_PREROLL_SECONDS = 0.1
// A read this far past where a decode stands continues it, decoding through the gap, instead of restarting.
export const CONTINUE_WITHIN_SECONDS = 2
// Decodes kept at once: an element reading ahead while it probes elsewhere (WebKit) keeps both going.
const MAX_CURSORS = 2

interface Cursor {
	reader: SampleReader
	assembler: PcmAssembler
	// The block decoded last, null before the first.
	last: SampleItem | null
	// A block decoded already and not yet placed (the one the stream was opened with).
	pending: SampleItem | null
	// For blocks decoded at another rate than the stream's (a rate change mid-stream).
	resampler: { from: number; resampler: LinearResampler } | null
	ended: boolean
	usedAt: number
}

// Which decode serves a read from `offset`: one standing at or a little before it, else none (a restart).
export function pickCursor<T extends { nextOffset: number }>(cursors: readonly T[], offset: number, reachBytes: number): T | null {
	let best: T | null = null

	for (const cursor of cursors) {
		const gap = offset - cursor.nextOffset

		if (gap >= 0 && gap <= reachBytes && (best === null || cursor.nextOffset > best.nextOffset)) {
			best = cursor
		}
	}

	return best
}

export class PcmStream {
	public readonly layout: WavLayout
	// The decoded audio's format: what the decoder gives, which a container can misstate (HE-AAC's core rate,
	// a DTS core's rate under its lossless extension).
	public readonly info: AudioInputInfo
	private readonly input: AudioInput
	private readonly header: Uint8Array
	private readonly stats: RemuxStats
	private readonly reachBytes: number
	private cursors: Cursor[] = []
	private tail: Promise<unknown> = Promise.resolve()
	private uses = 0

	// The stream of `input`, laid out by its first decoded block's rate and channels; that decode becomes the
	// stream's first cursor. Throws AudioUnsupportedError when the audio cannot be decoded or laid out.
	public static async open(input: AudioInput, stats: RemuxStats): Promise<PcmStream> {
		const state = { closed: false }
		const reader = input.reader(stats, () => state.closed)
		let first: SampleItem | null

		try {
			first = await reader.first(0)
		} catch (error) {
			await reader.close().catch(() => undefined)

			throw isSourceReadError(error) ? error : new AudioUnsupportedError("the audio cannot be decoded")
		}

		const info: AudioInputInfo = {
			...input.info,
			sampleRate: first?.sample.sampleRate ?? input.info.sampleRate,
			// The container's count stands when the first block is narrower (a mono intro); a wider one wins (a
			// parametric-stereo stream the container calls mono).
			channels: Math.max(first?.sample.numberOfChannels ?? 0, input.info.channels)
		}
		const layout = wavLayout(info)

		if (layout === null) {
			first?.sample.close()
			await reader.close().catch(() => undefined)

			throw new AudioUnsupportedError("too long for a WAV stream")
		}

		return new PcmStream({ input, info, layout, stats, initial: { reader, first, state } })
	}

	public constructor(params: {
		input: AudioInput
		info?: AudioInputInfo
		layout: WavLayout
		stats: RemuxStats
		initial?: { reader: SampleReader; first: SampleItem | null; state: { closed: boolean } }
	}) {
		const { input, layout, initial } = params

		this.input = input
		this.info = params.info ?? input.info
		this.layout = layout
		this.header = wavHeader(layout)
		this.stats = params.stats
		this.reachBytes = Math.round(CONTINUE_WITHIN_SECONDS * layout.sampleRate) * layout.blockAlign

		if (initial !== undefined) {
			this.closedState = initial.state
			this.cursors.push({
				reader: initial.reader,
				assembler: new PcmAssembler(layout, 0),
				last: null,
				pending: initial.first,
				resampler: null,
				ended: initial.first === null,
				usedAt: 0
			})
		}
	}

	private closedState = { closed: false }

	private get closed(): boolean {
		return this.closedState.closed
	}

	// Bytes `start`..`end` (exclusive) of the stream. Reads run one at a time, in the order they came.
	public read(start: number, end: number, signal: AbortSignal): Promise<Uint8Array<ArrayBuffer>> {
		const run = this.tail.then(() => this.readNow(start, Math.min(end, this.layout.size), signal))

		this.tail = run.catch(() => undefined)

		return run
	}

	public async close(): Promise<void> {
		this.closedState.closed = true

		await Promise.all(this.cursors.splice(0).map(cursor => dropCursor(cursor)))
		await this.input.close()
	}

	private async readNow(start: number, end: number, signal: AbortSignal): Promise<Uint8Array<ArrayBuffer>> {
		signal.throwIfAborted()

		if (this.closed) {
			throw new Error("audio stream closed")
		}

		const out = new Uint8Array(Math.max(0, end - start))

		if (start < WAV_HEADER_BYTES) {
			out.set(this.header.subarray(start, Math.min(end, WAV_HEADER_BYTES)))
		}

		const from = Math.max(start, WAV_HEADER_BYTES)
		let position = from

		if (position >= end) {
			return out
		}

		const cursor = this.cursorFor(position)

		try {
			// A short way ahead of the decode: through the gap, discarding it.
			while (cursor.assembler.nextOffset < position) {
				if (cursor.assembler.queued === 0 && !(await this.produce(cursor, signal))) {
					break
				}

				cursor.assembler.skip(position - cursor.assembler.nextOffset)
			}

			while (position < end) {
				if (cursor.assembler.queued === 0 && !(await this.produce(cursor, signal))) {
					break
				}

				// What was decoded for a read given up on stays queued for the next one from here.
				signal.throwIfAborted()

				const bytes = cursor.assembler.take(end - position)

				out.set(bytes, position - start)
				position += bytes.length
			}
		} catch (error) {
			if (signal.aborted) {
				// What it took goes back, so a read again from where it started (WebKit asks again from where the
				// bytes it got ended) continues this decode instead of restarting it.
				cursor.assembler.giveBack(out.subarray(from - start, position - start))
			} else {
				// A decode that failed (a read of the file, the decoder) is not resumed: the next read starts afresh.
				this.cursors = this.cursors.filter(other => other !== cursor)
				void dropCursor(cursor)
			}

			throw error
		}

		return out
	}

	private cursorFor(offset: number): Cursor {
		const found = pickCursor(
			this.cursors.map(cursor => ({ cursor, nextOffset: cursor.assembler.nextOffset })),
			offset,
			this.reachBytes
		)

		this.uses++

		if (found !== null) {
			found.cursor.usedAt = this.uses

			return found.cursor
		}

		const { frame } = frameAtOffset(this.layout, offset)
		const cursor: Cursor = {
			reader: this.input.reader(this.stats, () => this.closed),
			assembler: new PcmAssembler(this.layout, frame),
			last: null,
			pending: null,
			resampler: null,
			ended: false,
			usedAt: this.uses
		}

		this.cursors.push(cursor)

		if (this.cursors.length > MAX_CURSORS) {
			this.cursors.sort((a, b) => a.usedAt - b.usedAt)

			for (const dropped of this.cursors.splice(0, this.cursors.length - MAX_CURSORS)) {
				void dropCursor(dropped)
			}
		}

		return cursor
	}

	// Decodes the next block onto the cursor's output; false once the stream is complete.
	private async produce(cursor: Cursor, signal: AbortSignal): Promise<boolean> {
		const { assembler } = cursor

		while (assembler.queued === 0) {
			signal.throwIfAborted()

			if (assembler.complete) {
				return false
			}

			if (cursor.ended) {
				assembler.finish()
				continue
			}

			const begun = performance.now()
			const item =
				cursor.pending ??
				(cursor.last === null
					? await cursor.reader.first(Math.max(0, offsetSeconds(this.layout, assembler.nextOffset) - RESTART_PREROLL_SECONDS))
					: await cursor.reader.next(cursor.last))

			cursor.pending = null

			if (item === null) {
				cursor.ended = true
				continue
			}

			cursor.last = item

			const { sample } = item

			try {
				const frames = sample.numberOfFrames
				const planes: Float32Array[] = []

				for (let channel = 0; channel < sample.numberOfChannels; channel++) {
					const plane = new Float32Array(frames)

					sample.copyTo(plane, { planeIndex: channel, format: "f32-planar" })
					planes.push(plane)
				}

				const speakers = layoutFor(this.input.info.header, planes.length, decodedChannelMask(sample))
				const rate = this.layout.sampleRate * this.layout.decimation
				const placed = this.atRate(cursor, outputPlanes(planes, this.layout.channels, speakers), sample.sampleRate, rate)

				assembler.push(placed.planes, Math.round((sample.timestamp + placed.lead / sample.sampleRate) * rate))
				this.stats.transcodedMediaSeconds += sample.duration
			} finally {
				sample.close()
				this.stats.transcodeSeconds += (performance.now() - begun) / 1000
			}
		}

		return true
	}

	// Planes decoded at `from` Hz as the stream's rate: a block the decoder gave at another rate than the
	// stream was laid out with goes through a resampler carried from block to block. `lead` is where the
	// first frame out lies against the block's own first, in input frames (the resampler's carried position).
	private atRate(cursor: Cursor, planes: Float32Array[], from: number, to: number): { planes: Float32Array[]; lead: number } {
		if (from === to) {
			cursor.resampler = null

			return { planes, lead: 0 }
		}

		if (cursor.resampler?.from !== from) {
			cursor.resampler = { from, resampler: new LinearResampler(from, to) }
		}

		const lead = cursor.resampler.resampler.nextPosition

		return { planes: cursor.resampler.resampler.process(planes), lead }
	}
}

// Lets go of a decode, the block it decoded and did not place included (its native audio memory).
function dropCursor(cursor: Cursor): Promise<void> {
	cursor.pending?.sample.close()
	cursor.pending = null

	return cursor.reader.close().catch(() => undefined)
}

function offsetSeconds(layout: WavLayout, offset: number): number {
	return frameAtOffset(layout, offset).frame / layout.sampleRate
}
