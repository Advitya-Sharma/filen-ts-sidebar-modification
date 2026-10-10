import { Platform } from "react-native"
import { File } from "expo-file-system"
import { fetch } from "expo/fetch"
import { type ContentVerdict } from "@filen/shared"
import { TS_PACKET_BYTES, TS_SNIFF_PACKETS, TS_SYNC, type TsLayout } from "@filen/shared/videos/transportStreamHead"
import { normalizeFilePathForExpo } from "@/lib/paths"

// Whether a `.ts`/`.mts`/`.m2ts` file is an MPEG transport stream, and in which packet layout, from its first
// SNIFF_HEAD_BYTES through one bounded read. Decided as a player syncs, not as the web remux's probe reads: no
// PAT is required (Media3 plays the first program of a multiplex), and the first sync byte may sit anywhere in
// the first packet (a capture cut mid-packet), so "no stream" means no layout syncs at all.

// Plain 188-byte packets, BDAV (a 4-byte timestamp before each) and 204-byte (16 parity bytes after each).
const STRIDES = [TS_PACKET_BYTES, 192, 204] as const

export const SNIFF_HEAD_BYTES = Math.max(...STRIDES) * TS_SNIFF_PACKETS

// Media3's TsExtractor (androidx/media 1.9.0) peeks this much and gives up on a shorter file.
const MIN_STREAM_BYTES = TS_PACKET_BYTES * TS_SNIFF_PACKETS

// Whether `head` holds a C0 control other than whitespace, as a stream's packet headers and payload do and text
// does not: what keeps fixed-width TypeScript lines with a "G" in one column from syncing.
function holdsBinary(head: Uint8Array): boolean {
	for (const byte of head) {
		if (byte < 0x09 || (byte > 0x0d && byte < 0x20)) {
			return true
		}
	}

	return false
}

// The layout whose sync byte opens TS_SNIFF_PACKETS consecutive packets from some offset in the first one, as
// Media3's TsExtractor.sniff finds a 188-byte one; false for none. `prefix` is that offset.
export function transportStreamLayout(head: Uint8Array): TsLayout | false {
	for (const stride of STRIDES) {
		if (head.length < stride * TS_SNIFF_PACKETS) {
			continue
		}

		for (let offset = 0; offset < stride; offset++) {
			let packets = 0

			while (packets < TS_SNIFF_PACKETS && head[offset + packets * stride] === TS_SYNC) {
				packets++
			}

			if (packets === TS_SNIFF_PACKETS) {
				return holdsBinary(head) ? { stride, prefix: offset } : false
			}
		}
	}

	return false
}

// Reads `length` bytes at `start`; throws when they cannot be read.
export type ReadRange = (start: number, length: number) => Promise<Uint8Array>

// The stream's layout, or false for no stream, once the head was read; null when it could not be.
export type TransportStreamSniffer = (readRange: ReadRange, size: number) => Promise<TsLayout | false | null>

export const sniffTransportStream: TransportStreamSniffer = async (readRange, size) => {
	if (size < MIN_STREAM_BYTES) {
		return false
	}

	try {
		return transportStreamLayout(await readRange(0, Math.min(SNIFF_HEAD_BYTES, size)))
	} catch {
		return null
	}
}

// What a read that found `found` makes the file on `os`: no stream is notTransportStream. Media3's TsExtractor
// syncs on 188-byte packets only, from whatever offset in the first one they start at, so on Android a BDAV
// (192-byte) or parity (204-byte) stream is "unplayable": no verdict is kept and its page shows no preview.
// iOS plays no progressive stream at all (previewType.ts), so any layout is kept as the stream it is.
export function transportStreamOutcome(found: TsLayout | false, os: typeof Platform.OS = Platform.OS): ContentVerdict | "unplayable" {
	if (found === false) {
		return "notTransportStream"
	}

	return os !== "android" || found.stride === TS_PACKET_BYTES ? "transportStream" : "unplayable"
}

// At most `cap` bytes of `body`, then the rest is cancelled: nothing past them is kept.
async function readAtMost(body: ReadableStream<Uint8Array> | null, cap: number): Promise<Uint8Array> {
	const bytes = new Uint8Array(cap)
	let total = 0

	if (body === null) {
		return bytes.subarray(0, 0)
	}

	const reader = body.getReader()

	try {
		while (total < cap) {
			const { done, value } = await reader.read()

			if (done) {
				break
			}

			const take = Math.min(value.byteLength, cap - total)

			bytes.set(value.subarray(0, take), total)
			total += take
		}
	} finally {
		reader.cancel().catch(() => undefined)
	}

	return bytes.subarray(0, total)
}

// The first bytes of what `url` serves, without reading past them: a local copy through a file handle, the HTTP
// provider through one Range request. expo/fetch resolves at the headers and streams the body, so nothing past
// the bytes asked for is buffered. A 206 must carry exactly them; a 200 from the start is the whole file, and
// counts only when it is no longer than what was asked (a file that small); anything else fails the read.
export function readRangeFromUrl(url: string, signal: AbortSignal): ReadRange {
	return async (start, length) => {
		if (url.startsWith("file:")) {
			const handle = new File(normalizeFilePathForExpo(url)).open()

			try {
				handle.offset = start

				const bytes = handle.readBytes(length)

				if (bytes.byteLength !== length) {
					throw new Error("short read")
				}

				return bytes
			} finally {
				handle.close()
			}
		}

		const controller = new AbortController()
		const abort = () => controller.abort()

		signal.addEventListener("abort", abort, { once: true })

		try {
			const response = await fetch(url, {
				headers: {
					Range: `bytes=${start}-${start + length - 1}`
				},
				signal: controller.signal
			})

			if (response.status === 206) {
				const bytes = await readAtMost(response.body, length)

				if (bytes.byteLength !== length) {
					throw new Error("short read")
				}

				return bytes
			}

			if (response.status === 200 && start === 0) {
				// One byte past the request tells a file that small from one whose server ignored the range.
				const bytes = await readAtMost(response.body, length + 1)

				if (bytes.byteLength <= length) {
					return bytes
				}
			}

			throw new Error(`range not served: ${response.status}`)
		} finally {
			signal.removeEventListener("abort", abort)
			// Whatever the server still sends is not read.
			controller.abort()
		}
	}
}
