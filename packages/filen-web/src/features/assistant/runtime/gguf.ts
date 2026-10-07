// Reads the special-token strings out of a GGUF header, for sanitizeUntrustedText: the engine renders the
// chat template to a string and tokenizes it with special-token parsing, so any of these strings inside
// a note or a chat message would turn into a real control token. wllama exposes no tokenizer API, and
// llama.cpp leaves array metadata out of what it reports, so the header is read directly.

const MAGIC = 0x46554747 // "GGUF", little endian

const TYPE_STRING = 8
const TYPE_ARRAY = 9

const FIXED_SIZE: Readonly<Record<number, number>> = {
	0: 1,
	1: 1,
	2: 2,
	3: 2,
	4: 4,
	5: 4,
	6: 4,
	7: 1,
	10: 8,
	11: 8,
	12: 8
}

// llama_token_type: CONTROL and USER_DEFINED are the ones the tokenizer matches inside plain text.
const TOKEN_CONTROL = 3
const TOKEN_USER_DEFINED = 4

const CHUNK = 1024 * 1024
// The tokenizer arrays sit near the front; a header past this is not one worth reading.
const MAX_HEADER = 64 * 1024 * 1024

export type GgufRead = (offset: number, length: number) => Promise<Uint8Array>

class Reader {
	private buffer = new Uint8Array(0)
	private start = 0
	private position = 0
	private readonly decoder = new TextDecoder()
	private readonly read: GgufRead
	private readonly size: number

	constructor({ read, size }: { read: GgufRead; size: number }) {
		this.read = read
		this.size = size
	}

	private async ensure(length: number): Promise<void> {
		const end = this.position + length

		if (end > Math.min(this.size, MAX_HEADER)) {
			throw new Error("GGUF header truncated or too large")
		}

		if (end <= this.start + this.buffer.length) {
			return
		}

		const keep = this.buffer.subarray(this.position - this.start)
		const more = await this.read(this.start + this.buffer.length, Math.max(CHUNK, end - this.start - this.buffer.length))
		const next = new Uint8Array(keep.length + more.length)

		next.set(keep)
		next.set(more, keep.length)
		this.buffer = next
		this.start = this.position

		if (end > this.start + this.buffer.length) {
			throw new Error("GGUF header truncated")
		}
	}

	private view(length: number): DataView {
		const offset = this.position - this.start

		this.position += length

		return new DataView(this.buffer.buffer, this.buffer.byteOffset + offset, length)
	}

	async u32(): Promise<number> {
		await this.ensure(4)

		return this.view(4).getUint32(0, true)
	}

	async i32(): Promise<number> {
		await this.ensure(4)

		return this.view(4).getInt32(0, true)
	}

	async u64(): Promise<number> {
		await this.ensure(8)

		const value = this.view(8).getBigUint64(0, true)

		if (value > BigInt(Number.MAX_SAFE_INTEGER)) {
			throw new Error("GGUF value out of range")
		}

		return Number(value)
	}

	async string(): Promise<string> {
		const length = await this.u64()

		await this.ensure(length)

		const offset = this.position - this.start

		this.position += length

		return this.decoder.decode(this.buffer.subarray(offset, offset + length))
	}

	async skip(length: number): Promise<void> {
		await this.ensure(length)
		this.position += length
	}
}

async function skipValue({ reader, type }: { reader: Reader; type: number }): Promise<void> {
	if (type === TYPE_STRING) {
		await reader.skip(await reader.u64())

		return
	}

	if (type === TYPE_ARRAY) {
		const itemType = await reader.u32()
		const count = await reader.u64()
		const fixed = FIXED_SIZE[itemType]

		if (fixed !== undefined) {
			await reader.skip(fixed * count)

			return
		}

		for (let index = 0; index < count; index++) {
			await skipValue({ reader, type: itemType })
		}

		return
	}

	const fixed = FIXED_SIZE[type]

	if (fixed === undefined) {
		throw new Error(`GGUF value type ${String(type)} unknown`)
	}

	await reader.skip(fixed)
}

async function arrayHeader({ reader, expected }: { reader: Reader; expected: number }): Promise<number> {
	const itemType = await reader.u32()

	if (itemType !== expected) {
		throw new Error("GGUF tokenizer array has an unexpected type")
	}

	return reader.u64()
}

// The model's CONTROL and USER_DEFINED token strings longer than one character.
export async function readGgufSpecialTokens({ read, size }: { read: GgufRead; size: number }): Promise<string[]> {
	const reader = new Reader({ read, size })

	if ((await reader.u32()) !== MAGIC) {
		throw new Error("Not a GGUF file")
	}

	const version = await reader.u32()

	if (version < 2) {
		throw new Error(`GGUF version ${String(version)} unsupported`)
	}

	await reader.u64() // tensor count

	const keys = await reader.u64()
	let tokens: string[] | null = null
	let types: number[] | null = null

	for (let index = 0; index < keys && (tokens === null || types === null); index++) {
		const key = await reader.string()
		const type = await reader.u32()

		if (key === "tokenizer.ggml.tokens" && type === TYPE_ARRAY) {
			const count = await arrayHeader({ reader, expected: TYPE_STRING })

			tokens = []

			for (let item = 0; item < count; item++) {
				tokens.push(await reader.string())
			}
		} else if (key === "tokenizer.ggml.token_type" && type === TYPE_ARRAY) {
			const count = await arrayHeader({ reader, expected: 5 })

			types = []

			for (let item = 0; item < count; item++) {
				types.push(await reader.i32())
			}
		} else {
			await skipValue({ reader, type })
		}
	}

	if (tokens === null || types === null) {
		throw new Error("GGUF has no tokenizer token list")
	}

	if (tokens.length !== types.length) {
		throw new Error("GGUF tokenizer arrays disagree")
	}

	const vocabulary = tokens

	return types.flatMap((type, index) => {
		const token = vocabulary[index]

		return (type === TOKEN_CONTROL || type === TOKEN_USER_DEFINED) && token !== undefined && token.length > 1 ? [token] : []
	})
}

export function blobGgufRead(blob: Blob): GgufRead {
	return async (offset, length) => new Uint8Array(await blob.slice(offset, offset + length).arrayBuffer())
}
