import { describe, expect, it } from "vitest"
import type { SpellcheckRequest, SpellcheckResponse } from "../protocol"
import { SPELLCHECK_CACHE, serveSpellcheckWorker } from "../worker"
import { BASE, TINY_AFF, TINY_DIC, fakeOrigin, memoryCaches, serveText, tinyAssets } from "./fixtures"

// Drives a worker directly: sends requests, collects responses by id.
function harness(options: Parameters<typeof serveSpellcheckWorker>[1]) {
	let deliver: ((request: SpellcheckRequest) => void) | undefined
	const waiting = new Map<number, (response: SpellcheckResponse) => void>()

	serveSpellcheckWorker(
		{
			post: response => {
				waiting.get(response.id)?.(response)
				waiting.delete(response.id)
			},
			listen: onMessage => {
				deliver = onMessage
			}
		},
		{ base: BASE, ...options }
	)

	const send = (request: SpellcheckRequest): void => deliver?.(request)
	const ask = (request: SpellcheckRequest & { id: number }): Promise<SpellcheckResponse> =>
		new Promise(resolve => {
			waiting.set(request.id, resolve)
			send(request)
		})

	return { send, ask }
}

describe("serveSpellcheckWorker", () => {
	it("loads, checks, suggests and adds words", async () => {
		const origin = fakeOrigin()
		const assets = tinyAssets(origin)
		const files = assets.dictionaries["en_US"]

		if (files === undefined) throw new Error("fixture")

		const { send, ask } = harness({ fetch: origin.fetch, caches: undefined })

		send({ type: "init", wasm: assets.wasm, current: [] })

		// Pipelined: the check waits for the load.
		const loaded = ask({ type: "load", id: 1, dictionary: "en_US", files, words: ["Filen"] })
		const checked = ask({ type: "check", id: 2, dictionary: "en_US", words: "hello\nhelo\nFilen\nworlds", count: 4 })

		expect(await loaded).toEqual({ type: "loaded", id: 1 })
		expect(await checked).toEqual({ type: "checked", id: 2, verdicts: new Uint8Array([1, 0, 1, 1]) })
		expect(await ask({ type: "suggest", id: 3, dictionary: "en_US", word: "helo", ngram: false })).toMatchObject({
			type: "suggested",
			suggestions: expect.arrayContaining(["hello"])
		})

		send({ type: "add", words: ["zorp"] })

		expect(await ask({ type: "check", id: 4, dictionary: "en_US", words: "zorp", count: 1 })).toMatchObject({
			verdicts: new Uint8Array([1])
		})

		send({ type: "unload", dictionary: "en_US" })

		expect(await ask({ type: "check", id: 5, dictionary: "en_US", words: "hello", count: 1 })).toMatchObject({ type: "failed" })
	})

	it("caches verified downloads and serves the next load from the cache", async () => {
		const origin = fakeOrigin()
		const assets = tinyAssets(origin)
		const files = assets.dictionaries["en_US"]
		const caches = memoryCaches()

		if (files === undefined) throw new Error("fixture")

		const first = harness({ fetch: origin.fetch, caches })

		first.send({ type: "init", wasm: assets.wasm, current: [files.aff.url, files.dic.url] })
		expect(await first.ask({ type: "load", id: 1, dictionary: "en_US", files, words: [] })).toEqual({ type: "loaded", id: 1 })
		expect([...(caches.stored.get(SPELLCHECK_CACHE)?.keys() ?? [])].sort()).toEqual([`${BASE}en.aff.gz`, `${BASE}en.dic.gz`])

		origin.requests.length = 0

		const second = harness({ fetch: origin.fetch, caches })

		second.send({ type: "init", wasm: assets.wasm, current: [files.aff.url, files.dic.url] })
		expect(await second.ask({ type: "load", id: 1, dictionary: "en_US", files, words: [] })).toEqual({ type: "loaded", id: 1 })
		expect(origin.requests).toEqual([`${BASE}spell.wasm`])
	})

	it("rejects a file that does not match the manifest and does not cache it", async () => {
		const origin = fakeOrigin()
		const assets = tinyAssets(origin)
		const caches = memoryCaches()
		const files = assets.dictionaries["en_US"]

		if (files === undefined) throw new Error("fixture")

		const tampered = { ...files, dic: { ...files.dic, sha256: "0".repeat(64) } }
		const short = { ...files, dic: { ...files.dic, bytes: files.dic.bytes + 1 } }
		const { send, ask } = harness({ fetch: origin.fetch, caches })

		send({ type: "init", wasm: assets.wasm, current: [] })

		expect(await ask({ type: "load", id: 1, dictionary: "en_US", files: tampered, words: [] })).toMatchObject({
			type: "failed",
			message: expect.stringContaining("SHA-256")
		})
		expect(caches.stored.get(SPELLCHECK_CACHE)?.has(`${BASE}en.dic.gz`)).toBe(false)
		expect(await ask({ type: "load", id: 2, dictionary: "en_US", files: short, words: [] })).toMatchObject({ type: "failed" })
		// A failed load is forgotten: a correct one succeeds afterwards.
		expect(await ask({ type: "load", id: 3, dictionary: "en_US", files, words: [] })).toEqual({ type: "loaded", id: 3 })
	})

	it("prunes cached files this build does not reference", async () => {
		const origin = fakeOrigin()
		const assets = tinyAssets(origin)
		const caches = memoryCaches()
		const stale = serveText(origin, "/old.dic.gz", "1\nold\n")
		const cache = await caches.open(SPELLCHECK_CACHE)

		await cache.put(stale.url, new Response("x"))
		await cache.put("/en.aff.gz", new Response("y"))

		const { send } = harness({ fetch: origin.fetch, caches })

		send({ type: "init", wasm: assets.wasm, current: ["/en.aff.gz"] })
		await new Promise(resolve => setTimeout(resolve, 10))

		expect([...(caches.stored.get(SPELLCHECK_CACHE)?.keys() ?? [])]).toEqual([`${BASE}en.aff.gz`])
	})

	it("drops a cached file that no longer reads back and fetches it again", async () => {
		const origin = fakeOrigin()
		const assets = tinyAssets(origin)
		const files = assets.dictionaries["en_US"]
		const caches = memoryCaches()

		if (files === undefined) throw new Error("fixture")

		const cache = await caches.open(SPELLCHECK_CACHE)

		await cache.put(files.dic.url, new Response(new Uint8Array([0x1f, 0x8b, 1, 2, 3])))

		const { send, ask } = harness({ fetch: origin.fetch, caches })

		send({ type: "init", wasm: assets.wasm, current: [] })
		expect(await ask({ type: "load", id: 1, dictionary: "en_US", files, words: [] })).toEqual({ type: "loaded", id: 1 })
		expect(origin.requests).toContain(`${BASE}en.dic.gz`)
		expect((await (await cache.match(files.dic.url))?.arrayBuffer())?.byteLength).toBeGreaterThan(5)
	})

	it("starts the engine again after a failed start", async () => {
		const origin = fakeOrigin()
		const assets = tinyAssets(origin)
		const files = assets.dictionaries["en_US"]
		const wasm = origin.files.get(`${BASE}spell.wasm`)

		if (files === undefined || wasm === undefined) throw new Error("fixture")

		origin.files.delete(`${BASE}spell.wasm`)

		const { send, ask } = harness({ fetch: origin.fetch, caches: undefined })

		send({ type: "init", wasm: assets.wasm, current: [] })
		expect(await ask({ type: "load", id: 1, dictionary: "en_US", files, words: [] })).toMatchObject({ type: "failed" })

		origin.files.set(`${BASE}spell.wasm`, wasm)
		expect(await ask({ type: "load", id: 2, dictionary: "en_US", files, words: [] })).toEqual({ type: "loaded", id: 2 })
	})

	it("takes dictionary files a host already decompressed", async () => {
		const origin = fakeOrigin()
		const assets = tinyAssets(origin)
		const files = { aff: serveText(origin, "/raw.aff", TINY_AFF, false), dic: serveText(origin, "/raw.dic", TINY_DIC, false) }
		const { send, ask } = harness({ fetch: origin.fetch, caches: undefined })

		send({ type: "init", wasm: assets.wasm, current: [] })
		expect(await ask({ type: "load", id: 1, dictionary: "raw", files, words: [] })).toEqual({ type: "loaded", id: 1 })
		expect(await ask({ type: "check", id: 2, dictionary: "raw", words: "hello\nhelo", count: 2 })).toMatchObject({
			verdicts: new Uint8Array([1, 0])
		})
	})

	it("reports a missing dictionary file", async () => {
		const origin = fakeOrigin()
		const assets = tinyAssets(origin)
		const { send, ask } = harness({ fetch: origin.fetch, caches: undefined })
		const files = assets.dictionaries["en_US"]

		if (files === undefined) throw new Error("fixture")

		send({ type: "init", wasm: assets.wasm, current: [] })
		expect(
			await ask({ type: "load", id: 1, dictionary: "en_US", files: { ...files, aff: { ...files.aff, url: "/nope.gz" } }, words: [] })
		).toMatchObject({ type: "failed", message: expect.stringContaining("404") })
	})
})
