import { hevcCodecString } from "@filen/shared/videos"

function bytesOf(source: AllowSharedBufferSource): Uint8Array {
	return ArrayBuffer.isView(source) ? new Uint8Array(source.buffer, source.byteOffset, source.byteLength) : new Uint8Array(source)
}

// The codec string of a video track as the fragmented-MP4 output carries it. The muxer always writes an
// `hvc1` sample entry for HEVC, so the string says `hvc1` whatever the source's own entry was, with or
// without a decoder configuration record to rebuild it from.
export function outputVideoCodecString(codec: string, config: { codec: string; description?: AllowSharedBufferSource }): string {
	if (codec !== "hevc") {
		return config.codec
	}

	const rebuilt = config.description === undefined ? null : hevcCodecString({ fourcc: "hvc1", hvcC: bytesOf(config.description) })

	return rebuilt ?? config.codec.replace(/^hev1\./, "hvc1.")
}
