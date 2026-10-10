import { downmixToStereo, type Speaker } from "@/workers/remux/channelLayout.logic"
import { offsetOfFrame, type WavLayout } from "@/workers/remux/wav.logic"

// Decoded audio → the WAV stream's interleaved 16-bit frames, placed sample-exactly on the output timeline:
// output frame n is source frame n × decimation, whichever decode produced it. A decode restarted at a seek
// therefore yields the same bytes at the same offsets as one that ran there from the start, apart from the
// decoder's own warm-up, which lies before the requested frame and is dropped.

// A block whose first frame lies this close to where the previous one ended is taken as continuous: decoders
// round their timestamps, and a frame or two of drift is not a gap.
const CONTINUITY_SECONDS = 0.002

// The source's planes as the output's channels: stereo (a downmix of anything wider), or mono.
export function outputPlanes(planes: readonly Float32Array[], channels: number, layout: readonly Speaker[]): Float32Array[] {
	const frames = planes[0]?.length ?? 0

	if (planes.length === 0) {
		return Array.from({ length: channels }, () => new Float32Array(frames))
	}

	if (channels === 1 && planes.length === 1) {
		return [planes[0] ?? new Float32Array(frames)]
	}

	const [left, right] =
		planes.length === 1
			? [planes[0] ?? new Float32Array(frames), planes[0] ?? new Float32Array(frames)]
			: downmixToStereo(planes, layout)

	if (channels === 2) {
		return [left, right]
	}

	const mono = new Float32Array(frames)

	for (let i = 0; i < frames; i++) {
		mono[i] = ((left[i] ?? 0) + (right[i] ?? 0)) / 2
	}

	return [mono]
}

function toS16(value: number): number {
	const clamped = value > 1 ? 1 : value < -1 ? -1 : value

	return Math.round(clamped * 32767)
}

export class PcmAssembler {
	private readonly layout: WavLayout
	private readonly tolerance: number
	// The next source frame expected, and the output frame the next bytes belong to.
	private sourceFrame: number
	private outFrame: number
	private started = false
	// Source frames held back until a whole decimation group is in.
	private carry: Float32Array[] = []
	// The output not yet taken, in order: encoded bytes, or a count of silent bytes made only as they are
	// taken, so a long gap costs no memory until it is read.
	private readonly chunks: (Uint8Array<ArrayBuffer> | number)[] = []
	private chunkOffset = 0
	private queuedBytes = 0

	public constructor(layout: WavLayout, startFrame: number) {
		this.layout = layout
		this.outFrame = Math.min(startFrame, layout.frames)
		this.sourceFrame = this.outFrame * layout.decimation
		this.tolerance = Math.max(1, Math.round(CONTINUITY_SECONDS * layout.sampleRate * layout.decimation))
	}

	// The stream offset of the next byte taken.
	public get nextOffset(): number {
		return offsetOfFrame(this.layout, this.outFrame) - this.queuedBytes
	}

	public get queued(): number {
		return this.queuedBytes
	}

	// Every frame of the stream has been produced.
	public get complete(): boolean {
		return this.outFrame >= this.layout.frames
	}

	// Output-channel planes at the source rate whose first frame is source frame `first`. The first block
	// is cut to start exactly where this assembler does; later ones close a gap with silence and drop an
	// overlap, so the output never drifts from the source's timestamps.
	public push(planes: readonly Float32Array[], first: number): void {
		const frames = planes[0]?.length ?? 0
		let skip = 0

		if (frames === 0 || this.complete) {
			return
		}

		const offset = first - this.sourceFrame

		if (!this.started || Math.abs(offset) > this.tolerance) {
			if (offset > 0) {
				this.silence(offset)
			} else {
				skip = -offset
			}
		}

		// A block wholly before the start leaves it still to be cut exactly.
		if (skip >= frames) {
			return
		}

		this.started = true

		this.append(skip === 0 ? planes : planes.map(plane => plane.subarray(skip)))
	}

	// `frames` source frames of silence. Whole output frames of it are queued as a count, not as bytes.
	public silence(frames: number): void {
		const { decimation, blockAlign } = this.layout
		let left = frames

		this.started = true

		// A decimation group already begun is completed with real zeros.
		const carried = this.carry[0]?.length ?? 0

		if (carried > 0 && left > 0) {
			const fill = Math.min(left, decimation - carried)

			this.append(this.zeros(fill))
			left -= fill
		}

		const groups = Math.min(Math.floor(left / decimation), this.layout.frames - this.outFrame)

		if (groups > 0) {
			this.sourceFrame += groups * decimation
			this.outFrame += groups
			this.chunks.push(groups * blockAlign)
			this.queuedBytes += groups * blockAlign
			left -= groups * decimation
		}

		if (left > 0 && left < decimation && !this.complete) {
			this.append(this.zeros(left))
		}
	}

	// Silence up to the stream's end: the decoder stopped short of the duration the header promised.
	public finish(): void {
		const missing = (this.layout.frames - this.outFrame) * this.layout.decimation - (this.carry[0]?.length ?? 0)

		if (missing > 0) {
			this.silence(missing)
		}
	}

	// Up to `max` of the bytes produced so far.
	public take(max: number): Uint8Array<ArrayBuffer> {
		const out = new Uint8Array(Math.min(max, this.queuedBytes))

		this.consume(out.length, out)

		return out
	}

	// Puts back bytes taken last (from a read given up on), so they are the next taken.
	public giveBack(bytes: Uint8Array<ArrayBuffer>): void {
		if (bytes.length === 0) {
			return
		}

		const head = this.chunks[0]

		if (head !== undefined && this.chunkOffset > 0) {
			this.chunks[0] = typeof head === "number" ? head - this.chunkOffset : head.subarray(this.chunkOffset)
			this.chunkOffset = 0
		}

		this.chunks.unshift(bytes)
		this.queuedBytes += bytes.length
	}

	// Discards up to `max` of the bytes produced so far, without making them.
	public skip(max: number): void {
		this.consume(Math.min(max, this.queuedBytes), null)
	}

	private consume(size: number, out: Uint8Array | null): void {
		let written = 0

		while (written < size) {
			const chunk = this.chunks[0]

			if (chunk === undefined) {
				break
			}

			const length = typeof chunk === "number" ? chunk : chunk.length
			const count = Math.min(size - written, length - this.chunkOffset)

			// Silence is the zeros `out` starts as.
			if (out !== null && typeof chunk !== "number") {
				out.set(chunk.subarray(this.chunkOffset, this.chunkOffset + count), written)
			}

			written += count
			this.chunkOffset += count

			if (this.chunkOffset === length) {
				this.chunks.shift()
				this.chunkOffset = 0
			}
		}

		this.queuedBytes -= written
	}

	private zeros(frames: number): Float32Array[] {
		return Array.from({ length: this.layout.channels }, () => new Float32Array(frames))
	}

	private append(planes: readonly Float32Array[]): void {
		const { decimation, channels, blockAlign } = this.layout

		this.sourceFrame += planes[0]?.length ?? 0

		const joined = this.carry.length === 0 ? planes : planes.map((plane, channel) => concat(this.carry[channel], plane))
		const total = joined[0]?.length ?? 0
		const remaining = this.layout.frames - this.outFrame
		const groups = Math.min(Math.floor(total / decimation), remaining)
		const used = groups * decimation

		this.carry = groups === remaining || used === total ? [] : joined.map(plane => plane.slice(used))

		if (groups === 0) {
			return
		}

		const bytes = new Uint8Array(groups * blockAlign)
		const view = new DataView(bytes.buffer)

		for (let group = 0; group < groups; group++) {
			for (let channel = 0; channel < channels; channel++) {
				const plane = joined[channel] ?? joined[0]
				let sum = 0

				for (let k = 0; k < decimation; k++) {
					sum += plane?.[group * decimation + k] ?? 0
				}

				view.setInt16(group * blockAlign + channel * 2, toS16(sum / decimation), true)
			}
		}

		this.outFrame += groups
		this.chunks.push(bytes)
		this.queuedBytes += bytes.length
	}
}

function concat(head: Float32Array | undefined, tail: Float32Array): Float32Array {
	if (head === undefined || head.length === 0) {
		return tail
	}

	const out = new Float32Array(head.length + tail.length)

	out.set(head, 0)
	out.set(tail, head.length)

	return out
}
