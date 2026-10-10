// Random-access input: mobile reads ranges of the file on disk, web wraps bytes it already holds.
export interface ByteSource {
	readonly size: number
	read(offset: number, length: number): Promise<Uint8Array>
}

// Streamed output. The writer never touches a chunk again once it is written, so a sink may keep it.
export interface ByteSink {
	write(chunk: Uint8Array): Promise<void>
	close(): Promise<void>
}

// A source over bytes in memory; reads are views, never copies.
export function bytesSource(bytes: Uint8Array): ByteSource {
	return {
		size: bytes.length,
		read(offset: number, length: number): Promise<Uint8Array> {
			if (offset < 0 || length < 0 || offset + length > bytes.length) {
				return Promise.reject(new RangeError("Read outside the source"))
			}

			return Promise.resolve(bytes.subarray(offset, offset + length))
		}
	}
}

export interface MemorySink extends ByteSink {
	// Everything written, as one array (the single chunk itself when only one was written).
	bytes(): Uint8Array
}

export function memorySink(): MemorySink {
	const chunks: Uint8Array[] = []
	let total = 0

	return {
		write(chunk: Uint8Array): Promise<void> {
			chunks.push(chunk)
			total += chunk.length

			return Promise.resolve()
		},
		close(): Promise<void> {
			return Promise.resolve()
		},
		bytes(): Uint8Array {
			if (chunks.length === 1 && chunks[0] !== undefined) return chunks[0]

			const out = new Uint8Array(total)
			let offset = 0

			for (const chunk of chunks) {
				out.set(chunk, offset)
				offset += chunk.length
			}

			chunks.length = 0
			chunks.push(out)

			return out
		}
	}
}
