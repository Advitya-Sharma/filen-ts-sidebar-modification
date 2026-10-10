import { Platform } from "react-native"
import { File } from "expo-file-system"
import { fetch } from "expo/fetch"
import { type ContentVerdict } from "@filen/shared"
import { TS_SNIFF_BYTES, transportStreamOfHead, type TsLayout } from "@filen/shared/videos/transportStreamHead"
import { normalizeFilePathForExpo } from "@/lib/paths"

// Whether a `.ts`/`.mts`/`.m2ts` file is an MPEG transport stream, and in which packet layout, from its first
// TS_SNIFF_BYTES only, decided by the shared container probe (transportStreamOfHead) through one bounded read.

// Reads `length` bytes at `start`; throws when they cannot be read.
export type ReadRange = (start: number, length: number) => Promise<Uint8Array>

// The stream's layout, or false for no stream, once the head was read; null when it could not be.
export type TransportStreamSniffer = (readRange: ReadRange, size: number) => Promise<TsLayout | false | null>

export const sniffTransportStream: TransportStreamSniffer = async (readRange, size) => {
	try {
		return transportStreamOfHead(size < 8 ? new Uint8Array(0) : await readRange(0, Math.min(TS_SNIFF_BYTES, size)), size) ?? false
	} catch {
		return null
	}
}

// What a read that found `found` makes the file on `os`: no stream is notTransportStream. Media3's TsExtractor
// (androidx/media 1.9.0) syncs on 188-byte packets only, so on Android a BDAV (192-byte) or parity (204-byte)
// stream is "unplayable": no verdict is kept and its page shows no preview. iOS plays no progressive stream at
// all (previewType.ts), so any layout is kept as the stream it is.
export function transportStreamOutcome(found: TsLayout | false, os: typeof Platform.OS = Platform.OS): ContentVerdict | "unplayable" {
	if (found === false) {
		return "notTransportStream"
	}

	return os !== "android" || (found.stride === 188 && found.prefix === 0) ? "transportStream" : "unplayable"
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
