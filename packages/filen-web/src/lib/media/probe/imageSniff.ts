import { sniffImageMime } from "@/lib/media/probe/bytes"

// Bytes a container calls an image are passed on only once their first bytes say so: an MP4 `covr` may
// declare only `image/*`, and a Matroska attachment named cover.jpg can hold anything.

// Enough for every signature sniffImageMime knows (WebP's sits at 8-12).
const SNIFF_BYTES = 12

export interface ImageSniffStream {
	stream: TransformStream<Uint8Array, Uint8Array>
	// The bytes turned out not to be an image; the stream has errored.
	notImage(): boolean
}

export function imageSniffStream(): ImageSniffStream {
	let head: Uint8Array | null = new Uint8Array(0)
	let rejected = false
	const check = (bytes: Uint8Array): void => {
		if (sniffImageMime(bytes) === null) {
			rejected = true

			throw new Error("not an image")
		}
	}

	const stream = new TransformStream<Uint8Array, Uint8Array>({
		transform(chunk, controller) {
			if (head === null) {
				controller.enqueue(chunk)

				return
			}

			const joined = new Uint8Array(head.length + chunk.length)

			joined.set(head)
			joined.set(chunk, head.length)

			if (joined.length < SNIFF_BYTES) {
				head = joined

				return
			}

			check(joined)
			head = null
			controller.enqueue(joined)
		},
		flush(controller) {
			if (head !== null && head.length > 0) {
				check(head)
				controller.enqueue(head)
			}
		}
	})

	return { stream, notImage: () => rejected }
}

// The SDK's decoder reads with a BYOB reader, which only a byte stream offers; a TransformStream's readable is not one.
export function asByteStream(source: ReadableStream<Uint8Array>): ReadableStream<Uint8Array> {
	const reader = source.getReader()

	return new ReadableStream({
		type: "bytes",
		async pull(controller) {
			const { done, value } = await reader.read()

			if (done) {
				controller.close()
				// A BYOB read waiting on the end only settles once its request is answered with nothing.
				controller.byobRequest?.respond(0)

				return
			}

			// A copy: the byte stream takes ownership of what it is given, and a chunk may be a view of a larger buffer.
			controller.enqueue(new Uint8Array(value))
		},
		cancel(reason) {
			return reader.cancel(reason)
		}
	})
}
