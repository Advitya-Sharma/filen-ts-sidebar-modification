import { AUDIO_TARGET_CODECS, type AudioTarget } from "@/lib/media/remuxTypes"

// Re-encoded audio is always stereo at 48 kHz: every browser's encoder takes it (Opus encoders above two
// channels exist only in some engines), and it is what the source's surround is heard through on
// headphones and laptop speakers anyway.
export const TRANSCODE_SAMPLE_RATE = 48_000
export const TRANSCODE_CHANNELS = 2

// One of the four rates Chromium's AAC encoder accepts on Windows (96/128/160/192k); fine for Opus too.
export const TRANSCODE_BITRATE = 128_000

export interface EncoderTarget {
	target: AudioTarget
	// WebCodecs encoder codec, which is also the fragmented-MP4 codec string.
	codec: string
}

// The first of `preferred` (the page's MediaSource-accepted targets, best first) this engine can encode.
export async function pickEncoderTarget(
	preferred: readonly AudioTarget[],
	encodable: (codec: string) => Promise<boolean>
): Promise<EncoderTarget | null> {
	for (const target of preferred) {
		const candidate = { target, codec: AUDIO_TARGET_CODECS[target] }

		if (await encodable(candidate.codec)) {
			return candidate
		}
	}

	return null
}

// Where the next encoded block starts, in microseconds: continuous with the previous block (resampled
// blocks do not land on the source's own timestamps), unless the source jumped by more than a block (a
// gap or an overlap in the file), then on the source's timestamp again.
export function nextBlockTimestamp(params: { expectedUs: number | null; sourceUs: number; blockUs: number }): {
	timestampUs: number
	reanchored: boolean
} {
	const { expectedUs, sourceUs, blockUs } = params

	if (expectedUs === null || Math.abs(sourceUs - expectedUs) > Math.max(blockUs, 1)) {
		return { timestampUs: sourceUs, reanchored: true }
	}

	return { timestampUs: expectedUs, reanchored: false }
}

// Holds the decoder config an encoder reported with a chunk that was dropped (priming before the
// stream's start) until the first chunk that is written, which must carry it.
export class FirstChunkConfig<T> {
	private pending: T | null = null
	private delivered = false

	// The config to write with this chunk: its own, or one held from a dropped chunk.
	public take(config: T | undefined, dropped: boolean): T | undefined {
		if (dropped) {
			this.pending = config ?? this.pending

			return undefined
		}

		const out = this.delivered ? config : (config ?? this.pending ?? undefined)

		if (out !== undefined) {
			this.delivered = true
			this.pending = null
		}

		return out
	}
}

// Linear-interpolating rate conversion across consecutive blocks of one stream. Only for the rare
// source not at 48 kHz (AC-3 allows 32 and 44.1 kHz); quality is a step below a polyphase filter but
// well within what a 128 kb/s re-encode keeps.
export class LinearResampler {
	private readonly step: number
	// Position of the next output frame, in input frames, relative to the start of the next block.
	private position = 0
	// Each channel's last input frame of the previous block, for interpolating across the boundary.
	private previous: number[] = []

	public constructor(fromRate: number, toRate: number) {
		this.step = fromRate / toRate
	}

	// Where the next output frame lies, in input frames from the start of the next block (negative: between
	// the previous block's last frame and the next's first).
	public get nextPosition(): number {
		return this.position
	}

	public process(planes: readonly Float32Array[]): Float32Array[] {
		const frames = planes[0]?.length ?? 0

		if (this.previous.length !== planes.length) {
			this.previous = planes.map(plane => plane[0] ?? 0)
		}

		// Output frames whose position falls before this block's last frame.
		const count = frames === 0 ? 0 : Math.max(0, Math.ceil((frames - 1 - this.position) / this.step))
		const out = planes.map(() => new Float32Array(count))

		for (let channel = 0; channel < planes.length; channel++) {
			const plane = planes[channel]
			const target = out[channel]

			if (plane === undefined || target === undefined) {
				continue
			}

			const previous = this.previous[channel] ?? 0

			for (let index = 0; index < count; index++) {
				const at = this.position + index * this.step
				const base = Math.floor(at)
				const fraction = at - base
				const a = base < 0 ? previous : (plane[base] ?? 0)
				const b = plane[base + 1] ?? plane[base] ?? 0

				target[index] = a + (b - a) * fraction
			}

			this.previous[channel] = plane[frames - 1] ?? previous
		}

		this.position = this.position + count * this.step - frames

		return out
	}
}

// An MPEG-4 descriptor's tag, payload start and payload size; null when it does not fit.
function readDescriptor(bytes: Uint8Array, at: number): { tag: number; start: number; size: number } | null {
	const tag = bytes[at]
	let size = 0
	let position = at + 1

	// Up to four size bytes, seven bits each, continued while the top bit is set.
	for (let count = 0; count < 4; count++) {
		const byte = bytes[position++]

		if (tag === undefined || byte === undefined) {
			return null
		}

		size = (size << 7) | (byte & 0x7f)

		if ((byte & 0x80) === 0) {
			return position + size <= bytes.length ? { tag, start: position, size } : null
		}
	}

	return null
}

const ES_DESCRIPTOR = 0x03
const DECODER_CONFIG_DESCRIPTOR = 0x04
const DECODER_SPECIFIC_INFO = 0x05
// objectTypeIndication, streamType, bufferSizeDB (3), maxBitrate (4), avgBitrate (4).
const DECODER_CONFIG_FIXED_BYTES = 13

// The AudioSpecificConfig an AAC decoder config should carry as its description. WebKit's encoder
// hands out a whole ES_Descriptor (an `esds` payload) instead; written into an `esds` as is, the
// result is an init segment no MediaSource accepts.
export function audioSpecificConfig(description: Uint8Array): Uint8Array {
	if (description[0] !== ES_DESCRIPTOR) {
		return description
	}

	const es = readDescriptor(description, 0)

	if (es === null) {
		return description
	}

	let position = es.start + 2
	const flags = description[position++] ?? 0

	if ((flags & 0x80) !== 0) {
		position += 2
	}

	if ((flags & 0x40) !== 0) {
		position += 1 + (description[position] ?? 0)
	}

	if ((flags & 0x20) !== 0) {
		position += 2
	}

	const config = readDescriptor(description, position)

	if (config?.tag !== DECODER_CONFIG_DESCRIPTOR) {
		return description
	}

	const specific = readDescriptor(description, config.start + DECODER_CONFIG_FIXED_BYTES)

	return specific?.tag === DECODER_SPECIFIC_INFO ? description.subarray(specific.start, specific.start + specific.size) : description
}
