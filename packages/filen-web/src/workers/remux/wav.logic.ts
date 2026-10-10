// The constant-bitrate stream an audio file the browser cannot play is decoded into: 16-bit PCM in a WAV
// file of known size, which every engine plays from an <audio> element and seeks in by arithmetic alone
// (byte = header + seconds × byte rate), so a Range request names the exact sample it wants and the decoder
// restarts there.

export const WAV_HEADER_BYTES = 44
export const WAV_CONTENT_TYPE = "audio/wav"

const BYTES_PER_SAMPLE = 2
// RIFF sizes are 32-bit: Firefox and WebKit stop at the 4 GiB they can address, whatever the header claims.
const MAX_RIFF_SIZE = 0xffff_ffff
const MAX_DATA_BYTES = MAX_RIFF_SIZE - (WAV_HEADER_BYTES - 8)
// Below this the output would not be worth listening to.
const MIN_SAMPLE_RATE = 8000

export interface WavLayout {
	sampleRate: number
	channels: number
	// Source frames per output frame: a power of two, 1 unless the output had to shrink to fit.
	decimation: number
	frames: number
	// Bytes per output frame.
	blockAlign: number
	size: number
}

// The output for a source of `duration` seconds at `sampleRate` with `channels`: at most stereo at the
// source's own rate, or, for one too long to fit a WAV file that way, mono and then halved rates. Null when
// even the smallest would not fit.
export function wavLayout(params: { duration: number; sampleRate: number; channels: number }): WavLayout | null {
	const { duration, sampleRate } = params

	if (!(duration > 0) || !(sampleRate > 0)) {
		return null
	}

	const candidates: { channels: number; decimation: number }[] = [{ channels: Math.min(2, Math.max(1, params.channels)), decimation: 1 }]

	for (let decimation = 1; sampleRate / decimation >= MIN_SAMPLE_RATE; decimation *= 2) {
		candidates.push({ channels: 1, decimation })
	}

	for (const { channels, decimation } of candidates) {
		const rate = sampleRate / decimation

		if (!Number.isInteger(rate)) {
			continue
		}

		const frames = Math.round(duration * rate)
		const blockAlign = channels * BYTES_PER_SAMPLE
		const dataBytes = frames * blockAlign

		if (dataBytes <= MAX_DATA_BYTES) {
			return { sampleRate: rate, channels, decimation, frames, blockAlign, size: WAV_HEADER_BYTES + dataBytes }
		}
	}

	return null
}

export function wavHeader(layout: WavLayout): Uint8Array<ArrayBuffer> {
	const bytes = new Uint8Array(WAV_HEADER_BYTES)
	const view = new DataView(bytes.buffer)
	const dataBytes = layout.size - WAV_HEADER_BYTES
	const ascii = (offset: number, text: string) => {
		for (let i = 0; i < text.length; i++) {
			bytes[offset + i] = text.charCodeAt(i)
		}
	}

	ascii(0, "RIFF")
	view.setUint32(4, layout.size - 8, true)
	ascii(8, "WAVE")
	ascii(12, "fmt ")
	view.setUint32(16, 16, true)
	// PCM
	view.setUint16(20, 1, true)
	view.setUint16(22, layout.channels, true)
	view.setUint32(24, layout.sampleRate, true)
	view.setUint32(28, layout.sampleRate * layout.blockAlign, true)
	view.setUint16(32, layout.blockAlign, true)
	view.setUint16(34, BYTES_PER_SAMPLE * 8, true)
	ascii(36, "data")
	view.setUint32(40, dataBytes, true)

	return bytes
}

// The output frame holding byte `offset` (at or past the header), and where in it the byte lies.
export function frameAtOffset(layout: WavLayout, offset: number): { frame: number; within: number } {
	const data = Math.max(0, offset - WAV_HEADER_BYTES)

	return { frame: Math.floor(data / layout.blockAlign), within: data % layout.blockAlign }
}

export function offsetOfFrame(layout: WavLayout, frame: number): number {
	return WAV_HEADER_BYTES + frame * layout.blockAlign
}
