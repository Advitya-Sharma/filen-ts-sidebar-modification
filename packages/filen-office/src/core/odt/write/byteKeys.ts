// Values by a span of bytes (an element's markup) and a number (`salt`: what else the value depends on), found without decoding it: by an FNV-1a hash of the bytes, the
// bytes compared on a hit. Documents repeat a few property sets many times; a string key of each would be decoded
// and hashed again for every element. Bounded: it starts again past `limit` entries. The hash is seeded per map, and
// a bucket holds `CHAIN` entries at most (more are not kept), so crafted bytes that collide cost a few compares each.

const CHAIN = 8

export class ByteKeyMap<T> {
	private readonly buckets = new Map<number, { readonly bytes: Uint8Array; readonly salt: number; readonly value: T }[]>()
	private size = 0
	private readonly limit: number
	private readonly seed = randomSeed()

	constructor(limit = 4096) {
		this.limit = limit
	}

	get(bytes: Uint8Array, salt = 0): T | undefined {
		const bucket = this.buckets.get(hash(bytes, salt, this.seed))

		if (bucket === undefined) return undefined

		for (const entry of bucket) if (entry.salt === salt && equal(entry.bytes, bytes)) return entry.value

		return undefined
	}

	set(bytes: Uint8Array, value: T, salt = 0): void {
		if (this.size >= this.limit) {
			this.buckets.clear()
			this.size = 0
		}

		const key = hash(bytes, salt, this.seed)
		const bucket = this.buckets.get(key)

		if (bucket !== undefined && bucket.length >= CHAIN) return

		// A copy: the bytes are a view of a part that may be let go.
		const entry = { bytes: bytes.slice(), salt, value }

		if (bucket === undefined) this.buckets.set(key, [entry])
		else bucket.push(entry)

		this.size++
	}
}

function randomSeed(): number {
	const crypto: unknown = Reflect.get(globalThis, "crypto")
	const seed = new Uint32Array(1)

	// Hermes has no Web Crypto without a polyfill; a seed only has to be unknown to the file's author.
	if (typeof crypto === "object" && crypto !== null && "getRandomValues" in crypto && typeof crypto.getRandomValues === "function")
		crypto.getRandomValues(seed)
	else seed[0] = Math.floor(Math.random() * 0x1_0000_0000)

	return seed[0] ?? 0
}

function hash(bytes: Uint8Array, salt: number, seed: number): number {
	let h = Math.imul((0x811c9dc5 ^ seed) >>> 0, 0x01000193) ^ salt

	for (let k = 0; k < bytes.length; k++) h = Math.imul(h ^ (bytes[k] ?? 0), 0x01000193)

	return h >>> 0
}

function equal(a: Uint8Array, b: Uint8Array): boolean {
	if (a.length !== b.length) return false

	for (let k = 0; k < a.length; k++) if (a[k] !== b[k]) return false

	return true
}
