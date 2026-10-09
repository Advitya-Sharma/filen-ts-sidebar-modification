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
