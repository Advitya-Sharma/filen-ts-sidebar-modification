// Cutting fragmented-MP4 bytes into SourceBuffer appends. A long-GOP fragment can be 50-100 MB; one
// appendBuffer of that size makes the browser copy and parse it in one go (a long main-thread task and a
// transient second copy), so it goes in pieces, each a view into the original bytes. Pieces end at box
// boundaries where the boxes allow it; a box larger than a piece is cut inside, which MSE's segment
// parser accepts since it consumes appends as a byte stream.

export const MAX_APPEND_BYTES = 16 * 1_048_576

export interface Fmp4Box {
	type: string
	start: number
	size: number
}

// Top-level boxes of `bytes`. Stops at the first header that does not fit (a truncated tail is not a box).
export function scanBoxes(bytes: Uint8Array): Fmp4Box[] {
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
	const boxes: Fmp4Box[] = []
	let position = 0

	while (position + 8 <= bytes.length) {
		let size = view.getUint32(position)
		const type = String.fromCharCode(
			view.getUint8(position + 4),
			view.getUint8(position + 5),
			view.getUint8(position + 6),
			view.getUint8(position + 7)
		)

		if (size === 1) {
			if (position + 16 > bytes.length) {
				break
			}

			size = Number(view.getBigUint64(position + 8))
		} else if (size === 0) {
			size = bytes.length - position
		}

		if (size < 8) {
			break
		}

		boxes.push({ type, start: position, size })
		position += size
	}

	return boxes
}

// Whether `bytes` begins an initialization segment (ftyp/moov) rather than a media segment (moof/mdat).
export function isInitSegment(bytes: Uint8Array): boolean {
	const [first] = scanBoxes(bytes.subarray(0, 16))

	return first?.type === "ftyp" || first?.type === "moov"
}

// Views over `bytes`, in order and together covering all of it, none longer than `maxPiece`.
export function splitForAppend<T extends ArrayBufferLike>(bytes: Uint8Array<T>, maxPiece: number = MAX_APPEND_BYTES): Uint8Array<T>[] {
	if (bytes.length <= maxPiece) {
		return [bytes]
	}

	const pieces: Uint8Array<T>[] = []
	let pieceStart = 0
	let cursor = 0

	const closePiece = () => {
		if (cursor > pieceStart) {
			pieces.push(bytes.subarray(pieceStart, cursor))
			pieceStart = cursor
		}
	}

	for (const box of scanBoxes(bytes)) {
		const end = Math.min(bytes.length, box.start + box.size)

		if (end - pieceStart <= maxPiece) {
			cursor = end

			continue
		}

		closePiece()

		if (end - box.start <= maxPiece) {
			cursor = end

			continue
		}

		cutEvenly(bytes, box.start, end, maxPiece, pieces)
		pieceStart = end
		cursor = end
	}

	closePiece()

	// Bytes after the last whole box header (never produced by a muxer, but kept rather than dropped).
	if (pieceStart < bytes.length) {
		cutEvenly(bytes, pieceStart, bytes.length, maxPiece, pieces)
	}

	return pieces
}

// Cuts [start, end) into the fewest pieces of at most `maxPiece`, of near-equal size, so no tiny last
// piece costs an append round trip of its own.
function cutEvenly<T extends ArrayBufferLike>(
	bytes: Uint8Array<T>,
	start: number,
	end: number,
	maxPiece: number,
	into: Uint8Array<T>[]
): void {
	const count = Math.ceil((end - start) / maxPiece)
	const size = Math.ceil((end - start) / count)

	for (let at = start; at < end; at += size) {
		into.push(bytes.subarray(at, Math.min(end, at + size)))
	}
}
