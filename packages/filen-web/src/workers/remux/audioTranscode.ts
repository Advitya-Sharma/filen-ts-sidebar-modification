import { AudioSampleSink, EncodedPacket, type AudioCodec, type AudioSample, type InputAudioTrack } from "mediabunny"
import { BlockAbortError } from "@/lib/media/blockCache"
import type { RemuxStats } from "@/lib/media/remuxTypes"
import {
	audioSpecificConfig,
	FirstChunkConfig,
	LinearResampler,
	nextBlockTimestamp,
	TRANSCODE_BITRATE,
	TRANSCODE_CHANNELS,
	TRANSCODE_SAMPLE_RATE,
	type EncoderTarget
} from "@/workers/remux/audioMix.logic"
import { downmixToStereo, type Speaker } from "@/workers/remux/channelLayout.logic"
import type { FragmentedTrackOutput } from "@/workers/remux/output"
import type { PumpItem, PumpReader, PumpWriter } from "@/workers/remux/pump"

function normalizedConfig(config: AudioDecoderConfig): AudioDecoderConfig {
	const { description } = config

	if (description === undefined || !config.codec.startsWith("mp4a.")) {
		return config
	}

	const bytes = ArrayBuffer.isView(description)
		? new Uint8Array(description.buffer, description.byteOffset, description.byteLength)
		: new Uint8Array(description)

	return { ...config, description: audioSpecificConfig(bytes) }
}

// Blocks the encoder may hold before the writer waits for it to catch up. Generous: WebKit's AAC
// encoder keeps a few inputs queued until more audio arrives, so a small cap would wait on it forever.
const MAX_ENCODE_QUEUE = 32

// The wait is also cut short by time, never resting on a dequeue event alone.
const ENCODE_WAIT_MS = 50

// The FFmpeg-based decoders are large (about 1-1.6 MB each, the codec running in a nested worker of
// their own), so each is loaded the first time a track needs it and registered with the demuxer once.
let ac3Decoder: Promise<void> | null = null
let dtsDecoder: Promise<void> | null = null

export async function ensureDecoder(codec: AudioCodec | null): Promise<void> {
	if (codec === "ac3" || codec === "eac3") {
		ac3Decoder ??= import("@mediabunny/ac3").then(module => {
			module.registerAc3Decoder()
		})
		ac3Decoder.catch(() => {
			ac3Decoder = null
		})
		await ac3Decoder
	} else if (codec === "dts") {
		dtsDecoder ??= import("@mediabunny/dts").then(module => {
			module.registerDtsDecoder()
		})
		dtsDecoder.catch(() => {
			dtsDecoder = null
		})
		await dtsDecoder
	}
}

export interface SampleItem extends PumpItem {
	sample: AudioSample
}

export interface SampleReader extends PumpReader<SampleItem> {
	close(): Promise<void>
}

// Decoded audio from `time` on, in order, as pump items. The wait counts towards the transcode's cost:
// the decoder works in a nested worker while this one waits. A read aborted by a seek that this pump
// survived (the demuxer can serve a new read through an old, aborted one) restarts decoding once where
// it left off.
export function sampleReader(track: InputAudioTrack, stats: RemuxStats, isCancelled: () => boolean): SampleReader {
	const sink = new AudioSampleSink(track)
	let samples: AsyncGenerator<AudioSample, void, unknown> | null = null
	// End of the last sample handed out; a restarted decode skips what lies before it.
	let resumeAt = 0

	const pull = async (retried: boolean): Promise<SampleItem | null> => {
		const begun = performance.now()

		try {
			for (;;) {
				const next = await samples?.next()

				if (next === undefined || next.done === true) {
					return null
				}

				const sample = next.value

				if (sample.timestamp + sample.duration <= resumeAt) {
					sample.close()

					continue
				}

				resumeAt = sample.timestamp + sample.duration

				return { timestamp: sample.timestamp, duration: sample.duration, type: "key", sample }
			}
		} catch (error) {
			if (retried || !(error instanceof BlockAbortError) || isCancelled()) {
				throw error
			}

			samples = sink.samples(resumeAt)

			return await pull(true)
		} finally {
			stats.decodeWaitSeconds += (performance.now() - begun) / 1000
		}
	}

	return {
		first: time => {
			resumeAt = time
			samples = sink.samples(time)

			return pull(false)
		},
		next: () => pull(false),
		close: async () => {
			await samples?.return(undefined)
		}
	}
}

// Decoded samples → stereo 48 kHz → the browser's own AudioEncoder → fragmented MP4.
export class TranscodeWriter implements PumpWriter<SampleItem> {
	private readonly encoder: AudioEncoder
	private readonly output: FragmentedTrackOutput
	private readonly stats: RemuxStats
	private resampler: LinearResampler | null = null
	private nextTimestampUs: number | null = null
	private pending: Promise<void> = Promise.resolve()
	private failure: unknown = null
	private readonly layout: (channels: number) => Speaker[]
	private readonly firstConfig = new FirstChunkConfig<AudioDecoderConfig>()

	// `layout` names the decoded planes' speakers for a given channel count.
	public constructor(params: {
		target: EncoderTarget
		output: FragmentedTrackOutput
		stats: RemuxStats
		layout: (channels: number) => Speaker[]
	}) {
		const { target, output, stats } = params

		this.output = output
		this.stats = stats
		this.layout = params.layout
		this.encoder = new AudioEncoder({
			output: (chunk, meta) => {
				// Priming output before the stream's start has nowhere to go on the source timeline; the
				// decoder config it may carry goes with the first chunk written instead.
				const dropped = chunk.timestamp < 0
				const decoderConfig = this.firstConfig.take(meta?.decoderConfig, dropped)

				if (dropped) {
					return
				}

				const packet = EncodedPacket.fromEncodedChunk(chunk)
				const config = decoderConfig === undefined ? undefined : { decoderConfig: normalizedConfig(decoderConfig) }

				this.pending = this.pending.then(() => this.output.write(packet, config))
				this.pending.catch((error: unknown) => {
					this.failure ??= error
				})
			},
			error: error => {
				this.failure ??= error
			}
		})
		this.encoder.configure({
			codec: target.codec,
			sampleRate: TRANSCODE_SAMPLE_RATE,
			numberOfChannels: TRANSCODE_CHANNELS,
			bitrate: TRANSCODE_BITRATE
		})
	}

	public async write(item: SampleItem): Promise<void> {
		this.throwIfFailed()

		const begun = performance.now()
		const { sample } = item

		try {
			const frames = sample.numberOfFrames
			const planes: Float32Array[] = []

			for (let channel = 0; channel < sample.numberOfChannels; channel++) {
				const plane = new Float32Array(frames)

				sample.copyTo(plane, { planeIndex: channel, format: "f32-planar" })
				planes.push(plane)
			}

			const anchor = nextBlockTimestamp({
				expectedUs: this.nextTimestampUs,
				sourceUs: sample.microsecondTimestamp,
				blockUs: sample.microsecondDuration
			})

			// A jump in the source (a gap or overlap): the resampler's carried state belongs to before it.
			if (anchor.reanchored) {
				this.resampler = null
			}

			const stereo = downmixToStereo(planes, this.layout(planes.length))
			const [left, right] = sample.sampleRate === TRANSCODE_SAMPLE_RATE ? stereo : this.resample(sample.sampleRate, stereo)
			const count = left.length

			if (count === 0) {
				return
			}

			const data = new Float32Array(count * 2)

			data.set(left, 0)
			data.set(right, count)

			const timestamp = anchor.timestampUs
			const audioData = new AudioData({
				format: "f32-planar",
				sampleRate: TRANSCODE_SAMPLE_RATE,
				numberOfFrames: count,
				numberOfChannels: TRANSCODE_CHANNELS,
				timestamp,
				data
			})

			this.nextTimestampUs = timestamp + (count * 1_000_000) / TRANSCODE_SAMPLE_RATE
			this.encoder.encode(audioData)
			audioData.close()
			this.stats.transcodedMediaSeconds += sample.duration
		} finally {
			sample.close()
			this.stats.transcodeSeconds += (performance.now() - begun) / 1000
		}

		while (this.encoder.encodeQueueSize > MAX_ENCODE_QUEUE && this.encoder.state === "configured") {
			await new Promise(resolve => {
				this.encoder.addEventListener("dequeue", resolve, { once: true })
				setTimeout(resolve, ENCODE_WAIT_MS)
			})
		}

		await this.pending
	}

	private resample(fromRate: number, stereo: [Float32Array, Float32Array]): [Float32Array, Float32Array] {
		this.resampler ??= new LinearResampler(fromRate, TRANSCODE_SAMPLE_RATE)

		const [left, right] = this.resampler.process(stereo)

		return [left ?? new Float32Array(), right ?? new Float32Array()]
	}

	public async end(): Promise<void> {
		this.throwIfFailed()
		await this.encoder.flush()
		await this.pending
		this.throwIfFailed()
		this.close()
		await this.output.end()
	}

	public close(): void {
		if (this.encoder.state !== "closed") {
			this.encoder.close()
		}
	}

	private throwIfFailed(): void {
		if (this.failure !== null) {
			throw this.failure instanceof Error ? this.failure : new Error("audio encoding failed", { cause: this.failure })
		}
	}
}
