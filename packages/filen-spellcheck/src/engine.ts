import { WORD_SEPARATOR } from "./protocol"

// The spellchecker's wasm module (crate/, a C ABI over spellbook) from JavaScript. Platform-free: give it
// the compiled module or its bytes. The module owns its buffers and its dictionaries: input is encoded into
// its scratch region (fs_buffer gives the address, asked again before every write since the region moves
// when it grows), results are copied out of its output region, and a dictionary is a small id. Every length
// and id is checked on the Rust side; a bad one is an error code, never a trap.

interface Exports {
	readonly memory: WebAssembly.Memory
	fs_buffer(minLen: number): number
	fs_output(): number
	fs_new(affLen: number, dicLen: number): number
	fs_drop(id: number): number
	fs_check(id: number, len: number): number
	fs_suggest(id: number, len: number, ngram: number): number
	fs_add(id: number, len: number): number
}

const EXPORTS = ["memory", "fs_buffer", "fs_output", "fs_new", "fs_drop", "fs_check", "fs_suggest", "fs_add"] as const

// The module's error codes (crate/src/boundary.rs).
const ERRORS: Readonly<Record<number, string>> = {
	[-1]: "unknown dictionary id",
	[-2]: "length outside the scratch region",
	[-3]: "input is not UTF-8",
	[-4]: "the dictionary does not parse",
	[-5]: "out of memory in the spellcheck module"
}
const ERR_UTF8 = -3
const ERR_PARSE = -4
const ERR_MEMORY = -5

const encoder = new TextEncoder()
const decoder = new TextDecoder()

export class SpellcheckModuleError extends Error {
	readonly code: number

	constructor(code: number) {
		super(`Spellcheck module: ${ERRORS[code] ?? `error ${String(code)}`}`)
		this.name = "SpellcheckModuleError"
		this.code = code
	}
}

function checked(result: number): number {
	if (result < 0) throw new SpellcheckModuleError(result)

	return result
}

function exportsOf(instance: WebAssembly.Instance): Exports {
	const raw = instance.exports

	for (const name of EXPORTS) {
		if (!(name in raw)) throw new Error(`The spellcheck module lacks ${name}`)
	}

	return raw as unknown as Exports
}

export class SpellEngine {
	private readonly wasm: Exports

	private constructor(wasm: Exports) {
		this.wasm = wasm
	}

	static async instantiate(module: WebAssembly.Module | BufferSource): Promise<SpellEngine> {
		const compiled = module instanceof WebAssembly.Module ? module : await WebAssembly.compile(module)

		return new SpellEngine(exportsOf(await WebAssembly.instantiate(compiled, {})))
	}

	// Bytes currently in wasm memory (it only grows).
	get memoryBytes(): number {
		return this.wasm.memory.buffer.byteLength
	}

	// The scratch region with room for `len` bytes, as a view to write into (valid until the next call).
	private scratch(len: number): Uint8Array<ArrayBuffer> {
		const ptr = this.wasm.fs_buffer(len)

		if (ptr === 0) throw new SpellcheckModuleError(ERR_MEMORY)

		return new Uint8Array(this.wasm.memory.buffer, ptr, len)
	}

	// Parses a dictionary from its .aff and .dic bytes; undefined when they are not UTF-8 or do not parse.
	load(aff: Uint8Array, dic: Uint8Array): SpellDictionary | undefined {
		const region = this.scratch(aff.byteLength + dic.byteLength)

		region.set(aff)
		region.set(dic, aff.byteLength)

		const id = this.wasm.fs_new(aff.byteLength, dic.byteLength)

		return id === ERR_PARSE || id === ERR_UTF8 ? undefined : new SpellDictionary(this, checked(id))
	}

	// Encodes text into the scratch region; returns its UTF-8 length.
	private encode(text: string): number {
		// UTF-16 to UTF-8 takes at most 3 bytes per code unit.
		return encoder.encodeInto(text, this.scratch(text.length * 3)).written
	}

	private output(len: number): Uint8Array<ArrayBuffer> {
		return new Uint8Array(this.wasm.memory.buffer, this.wasm.fs_output(), len)
	}

	/** @internal */
	check(id: number, words: string, count: number): Uint8Array {
		const verdicts = checked(this.wasm.fs_check(id, this.encode(words)))

		if (verdicts !== count) throw new Error("The spellcheck module rejected a word list")

		return this.output(verdicts).slice()
	}

	/** @internal */
	suggest(id: number, word: string, ngram: boolean): string[] {
		const len = checked(this.wasm.fs_suggest(id, this.encode(word), ngram ? 1 : 0))

		return len === 0 ? [] : decoder.decode(this.output(len)).split(WORD_SEPARATOR)
	}

	/** @internal */
	add(id: number, word: string): boolean {
		return checked(this.wasm.fs_add(id, this.encode(word))) === 1
	}

	/** @internal */
	drop(id: number): void {
		checked(this.wasm.fs_drop(id))
	}
}

// One loaded dictionary. Its memory is reused by the module after dispose but never returned to the
// system: unloading a language for good means dropping the whole engine (its worker).
export class SpellDictionary {
	private readonly engine: SpellEngine
	private id: number

	constructor(engine: SpellEngine, id: number) {
		this.engine = engine
		this.id = id
	}

	private live(): number {
		if (this.id === 0) throw new Error("The dictionary was disposed")

		return this.id
	}

	// One verdict per word of `words` ("\n"-joined, `count` of them): 1 correct, 0 misspelled.
	checkJoined(words: string, count: number): Uint8Array {
		return count === 0 ? new Uint8Array(0) : this.engine.check(this.live(), words, count)
	}

	check(words: readonly string[]): Uint8Array {
		return this.checkJoined(words.join(WORD_SEPARATOR), words.length)
	}

	// Corrections, best first. Without `ngram` only the fast edit-based pass runs.
	suggest(word: string, ngram: boolean): string[] {
		return word.includes(WORD_SEPARATOR) ? [] : this.engine.suggest(this.live(), word, ngram)
	}

	// Accepts `word` from now on (with the casing variants a dictionary word gets).
	add(word: string): boolean {
		return this.engine.add(this.live(), word)
	}

	dispose(): void {
		if (this.id === 0) return

		this.engine.drop(this.id)
		this.id = 0
	}
}
