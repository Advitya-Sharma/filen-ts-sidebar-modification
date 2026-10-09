// Random-access reads over a remote file, in whole blocks. The SDK downloads and decrypts a file chunk
// by chunk (1 MiB), so asking for an aligned chunk costs the same transfer as asking for three bytes of
// it; reading in chunk-sized blocks and keeping the last few turns a parser's many tiny reads (a tag
// header, a trailer probe, the atoms of an MP4 `moov`) into one download per distinct chunk.
export const BLOCK_BYTES = 1_048_576

// Resident blocks per parse. The usual working set is the head block and the tail block (every parse
// probes the file's end for ID3v1/APE trailers), plus one for a tag or cover crossing a boundary.
const CACHED_BLOCKS = 3

// Block downloads one parse may make. Covers a head tag with a multi-megabyte cover plus the tail, or a
// container's track list read from its head or tail; past it a parser is walking the file's body.
export const MAX_BLOCK_READS = 8

// A block download failed. The parse is abandoned but nothing is concluded about the file.
export class BlockReadError extends Error {
	public constructor(cause: unknown) {
		super("block read failed", { cause })
		this.name = "BlockReadError"
	}
}

// The parse asked for more of the file than it is allowed to read. A property of the file, not of the read.
export class ReadBudgetError extends Error {
	public constructor() {
		super("read budget exhausted")
		this.name = "ReadBudgetError"
	}
}

// A read asked for bytes outside the file: the parser followed a size or offset the file cannot hold.
export class OutOfFileReadError extends Error {
	public constructor(position: number, length: number, size: number) {
		super(`read of ${String(length)} bytes at ${String(position)} is outside a ${String(size)}-byte file`)
		this.name = "OutOfFileReadError"
	}
}

// `end` is exclusive.
export type ReadRange = (start: number, end: number) => Promise<Uint8Array>

export interface BlockSourceOptions {
	maxBlockReads?: number
	cachedBlocks?: number
}

// The block cache and budget one file's parse shares, possibly across more than one tokenizer (a
// content sniff that falls back to a mime hint re-reads the same blocks for free). `readError` keeps the
// first failed download even when a parser swallows the throw and resolves partial tags.
export class BlockSource {
	public readonly size: number
	public readError: BlockReadError | null = null
	private readonly readRange: ReadRange
	private readonly blocks = new Map<number, Promise<Uint8Array>>()
	private readonly cachedBlocks: number
	private maxBlockReads: number
	private reads = 0

	public constructor(size: number, readRange: ReadRange, options: BlockSourceOptions = {}) {
		this.size = size
		this.readRange = readRange
		this.maxBlockReads = options.maxBlockReads ?? MAX_BLOCK_READS
		this.cachedBlocks = options.cachedBlocks ?? CACHED_BLOCKS
	}

	// Block downloads made so far.
	public get blockReads(): number {
		return this.reads
	}

	// Raises the budget to `total` downloads once a parser knows how much of the file it has to walk
	// (an MP4 `moov`'s size); never lowers it.
	public allowBlockReads(total: number): void {
		this.maxBlockReads = Math.max(this.maxBlockReads, total)
	}

	// Whether reading at `position` costs no download.
	public isCached(position: number): boolean {
		return this.blocks.has(Math.floor(position / BLOCK_BYTES))
	}

	// The resident block holding `position`, shared rather than copied, and never downloaded: null when
	// it is not resident. For scans over bytes already paid for; the caller must not write to it.
	public async residentBlock(position: number): Promise<Uint8Array | null> {
		const block = this.blocks.get(Math.floor(position / BLOCK_BYTES))

		return block === undefined ? null : await block
	}

	private block(index: number): Promise<Uint8Array> {
		const cached = this.blocks.get(index)

		if (cached !== undefined) {
			// Re-inserted so the Map's insertion order stays least-recently-used first.
			this.blocks.delete(index)
			this.blocks.set(index, cached)

			return cached
		}

		if (this.reads >= this.maxBlockReads) {
			return Promise.reject(new ReadBudgetError())
		}

		this.reads++

		const start = index * BLOCK_BYTES
		const end = Math.min(this.size, start + BLOCK_BYTES)
		// A failed download stays cached as the rejection: the SDK has already retried it, so a parser that
		// swallows the throw and reads the block again fails fast instead of downloading it again.
		const guarded = this.readRange(start, end)
			.then(bytes => {
				if (bytes.length !== end - start) {
					throw new Error(`short block read: ${String(bytes.length)} of ${String(end - start)} bytes`)
				}

				return bytes
			})
			.catch((error: unknown) => {
				const readError = new BlockReadError(error)

				this.readError ??= readError

				throw readError
			})

		this.blocks.set(index, guarded)

		while (this.blocks.size > this.cachedBlocks) {
			const oldest = this.blocks.keys().next()

			if (oldest.done === true) {
				break
			}

			this.blocks.delete(oldest.value)
		}

		return guarded
	}

	// Fills `target[0, length)` from the file at `position`; a range outside the file rejects with
	// OutOfFileReadError rather than reading a block past the end, which would never fill the target.
	public async read(target: Uint8Array, position: number, length: number): Promise<void> {
		if (
			!Number.isSafeInteger(position) ||
			!Number.isSafeInteger(length) ||
			position < 0 ||
			length < 0 ||
			position + length > this.size
		) {
			throw new OutOfFileReadError(position, length, this.size)
		}

		let written = 0

		while (written < length) {
			const at = position + written
			const index = Math.floor(at / BLOCK_BYTES)
			const block = await this.block(index)
			const offset = at - index * BLOCK_BYTES
			const take = Math.min(length - written, block.length - offset)

			target.set(block.subarray(offset, offset + take), written)
			written += take
		}
	}
}

// Bytes already in memory, read without copying.
export function bytesReadRange(bytes: Uint8Array): ReadRange {
	return (start, end) => Promise.resolve(bytes.subarray(start, end))
}

export function blobReadRange(blob: Blob): ReadRange {
	return async (start, end) => new Uint8Array(await blob.slice(start, end).arrayBuffer())
}

// A Range request against a same-origin URL that answers 206 (the service worker's inline-preview
// route). Anything else is a failure: a 200 would be the whole file, so its body is never read.
export function urlReadRange(url: string): ReadRange {
	return async (start, end) => {
		const response = await fetch(url, { headers: { Range: `bytes=${String(start)}-${String(end - 1)}` } })

		if (response.status !== 206) {
			await response.body?.cancel().catch(() => undefined)

			throw new Error(`range read answered ${String(response.status)}`)
		}

		return new Uint8Array(await response.arrayBuffer())
	}
}
