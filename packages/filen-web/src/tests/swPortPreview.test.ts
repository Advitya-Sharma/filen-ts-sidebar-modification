import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { BLOCK_BYTES } from "@filen/shared/videos"
import { exposeRangePort } from "@/lib/media/rangePort"
import {
	SW_DOWNLOAD_PREFIX,
	SW_ERROR_SUPERSEDED,
	SW_MSG_PREVIEW_STATUS,
	SW_MSG_REGISTER_PORT_PREVIEW,
	SW_MSG_RELEASE_PREVIEW,
	SW_MSG_STREAM_GONE,
	SW_PREVIEW_BUCKET_PLAYER,
	SW_PROTOCOL_VERSION
} from "@/lib/sw/protocol"

// sw.ts's port-backed preview route (a public link's file, read by the page's SDK worker): the worker holds
// only the reading end of a range port, and answers Range requests by reading through it. `self` is faked
// as in swClientLifecycle.test.ts; the serving end is the real exposeRangePort over a real MessageChannel.

vi.mock("@filen/sdk-rs/service-worker/sdk-rs.js", () => ({
	default: () => Promise.resolve(),
	fromStringified: () => ({ free: () => undefined })
}))

type Listener = (event: unknown) => void

let listeners: Map<string, Listener>
// What the worker posted to its window clients.
let posted: unknown[] = []
const channels: MessageChannel[] = []

async function loadSw(): Promise<void> {
	listeners = new Map()
	posted = []
	vi.stubGlobal("self", {
		location: { origin: "https://app.test" },
		addEventListener: (type: string, listener: Listener) => {
			listeners.set(type, listener)
		},
		skipWaiting: () => Promise.resolve(),
		clients: {
			claim: () => Promise.resolve(),
			matchAll: () => Promise.resolve([{ postMessage: (data: unknown) => posted.push(data) }])
		}
	})
	vi.resetModules()
	await import("@/sw/sw")
}

function message(data: unknown): Promise<unknown> {
	return new Promise(resolve => {
		listeners.get("message")?.({ data, ports: [{ postMessage: resolve }] })
	})
}

// The byte at `offset` of the served file.
function byteAt(offset: number): number {
	return offset % 251
}

// Serves a file of `size` bytes on a fresh port and registers its reading end under `id`.
async function registerPort(id: string, size: number, contentType = "video/mp4", stall = false, generation?: number) {
	const channel = new MessageChannel()
	const lifetime = new AbortController()
	const reads: [number, number][] = []

	channels.push(channel)
	exposeRangePort(
		channel.port1,
		(start, end, signal) => {
			reads.push([start, end])

			if (stall) {
				return new Promise((_resolve, reject) => {
					signal.addEventListener("abort", () => {
						reject(new Error("aborted"))
					})
				})
			}

			return Promise.resolve(Uint8Array.from({ length: end - start }, (_, index) => byteAt(start + index)))
		},
		lifetime
	)

	const ack = await message({
		type: SW_MSG_REGISTER_PORT_PREVIEW,
		id,
		name: "clip.mp4",
		size,
		contentType,
		bucket: SW_PREVIEW_BUCKET_PLAYER,
		port: channel.port2,
		...(generation === undefined ? {} : { generation })
	})

	return { ack, lifetime, reads }
}

function get(id: string, range?: string): { response: Response; settled: Promise<unknown> } {
	const answered: { response?: Response; settled: Promise<unknown> } = { settled: Promise.resolve() }

	listeners.get("fetch")?.({
		request: {
			url: `https://app.test${SW_DOWNLOAD_PREFIX}${id}`,
			method: "GET",
			mode: "no-cors",
			headers: new Headers(range === undefined ? {} : { Range: range })
		},
		respondWith: (value: Response) => {
			answered.response = value
		},
		waitUntil: (promise: Promise<unknown>) => {
			answered.settled = promise
		}
	})

	if (answered.response === undefined) {
		throw new Error("not answered")
	}

	return { response: answered.response, settled: answered.settled }
}

function expected(start: number, end: number): number[] {
	return Array.from({ length: end - start + 1 }, (_, index) => byteAt(start + index))
}

async function bodyOf(response: Response): Promise<number[]> {
	return Array.from(new Uint8Array(await response.arrayBuffer()))
}

function tick(ms = 20): Promise<void> {
	return new Promise(resolve => setTimeout(resolve, ms))
}

beforeEach(async () => {
	await loadSw()
})

afterEach(() => {
	for (const channel of channels.splice(0)) {
		channel.port1.close()
		channel.port2.close()
	}

	vi.unstubAllGlobals()
})

describe("sw.ts port-backed preview", () => {
	it("registers without a session Client and answers a Range request from the port with an inline 206", async () => {
		const { ack } = await registerPort("p1", 1_000)

		expect(ack).toEqual({ ok: true })

		const { response, settled } = get("p1", "bytes=10-19")

		expect(response.status).toBe(206)
		expect(response.headers.get("Content-Range")).toBe("bytes 10-19/1000")
		expect(response.headers.get("Content-Length")).toBe("10")
		expect(response.headers.get("Content-Type")).toBe("video/mp4")
		expect(response.headers.get("Content-Disposition")).toBeNull()
		expect(response.headers.get("Content-Security-Policy")).toBe("sandbox")
		expect(await bodyOf(response)).toEqual(expected(10, 19))
		await settled
	})

	it("reads a range one SDK chunk at a time, each window ending on a chunk boundary", async () => {
		const size = 10 * BLOCK_BYTES
		const end = 3 * BLOCK_BYTES
		const { reads } = await registerPort("p1", size)
		const { response } = get("p1", `bytes=100-${String(end)}`)
		const body = await bodyOf(response)

		expect(response.headers.get("Content-Range")).toBe(`bytes 100-${String(end)}/${String(size)}`)
		expect(body.length).toBe(end - 100 + 1)
		expect(body.slice(0, 3)).toEqual(expected(100, 102))
		expect(body.at(-1)).toBe(byteAt(end))
		expect(reads).toEqual([
			[100, BLOCK_BYTES],
			[BLOCK_BYTES, 2 * BLOCK_BYTES],
			[2 * BLOCK_BYTES, 3 * BLOCK_BYTES],
			[3 * BLOCK_BYTES, end + 1]
		])
	})

	it("answers an open-ended range with one piece, read chunk by chunk", async () => {
		const size = 10 * BLOCK_BYTES
		const { reads } = await registerPort("p1", size)
		const { response } = get("p1", "bytes=0-")

		await response.arrayBuffer()

		expect(response.headers.get("Content-Range")).toBe(`bytes 0-${String(4 * BLOCK_BYTES - 1)}/${String(size)}`)
		expect(reads).toEqual([
			[0, BLOCK_BYTES],
			[BLOCK_BYTES, 2 * BLOCK_BYTES],
			[2 * BLOCK_BYTES, 3 * BLOCK_BYTES],
			[3 * BLOCK_BYTES, 4 * BLOCK_BYTES]
		])
	})

	it("sends each chunk as it arrives, and reads no further than the consumer has asked for", async () => {
		const { reads } = await registerPort("p1", 10 * BLOCK_BYTES)
		const { response, settled } = get("p1", "bytes=0-")
		const body = response.body?.getReader()
		const first = await body?.read()

		expect(first?.value?.length).toBe(BLOCK_BYTES)
		await tick()
		expect(reads).toEqual([[0, BLOCK_BYTES]])

		await body?.cancel()
		await settled
	})

	it("answers no Range header with the whole file and Accept-Ranges, and an unsatisfiable one with 416", async () => {
		await registerPort("p1", 300)

		const whole = get("p1").response

		expect(whole.status).toBe(200)
		expect(whole.headers.get("Accept-Ranges")).toBe("bytes")
		expect(await bodyOf(whole)).toEqual(expected(0, 299))
		expect(get("p1", "bytes=400-500").response.status).toBe(416)

		// A last byte past the end is clamped to it.
		const clamped = get("p1", "bytes=290-500").response

		expect(clamped.status).toBe(206)
		expect(clamped.headers.get("Content-Range")).toBe("bytes 290-299/300")
		expect(await bodyOf(clamped)).toEqual(expected(290, 299))
	})

	it("serves a contentType off the inline allowlist as an attachment", async () => {
		await registerPort("p1", 100, "text/html")

		const { response } = get("p1", "bytes=0-9")

		expect(response.headers.get("Content-Type")).toBe("application/octet-stream")
		expect(response.headers.get("Content-Disposition")).toContain("attachment")
		await response.arrayBuffer()
	})

	it("closes the port on release, and the id answers 404 after", async () => {
		const { lifetime } = await registerPort("p1", 100)

		listeners.get("message")?.({ data: { type: SW_MSG_RELEASE_PREVIEW, id: "p1" }, ports: [] })
		await tick()

		expect(lifetime.signal.aborted).toBe(true)
		expect(get("p1", "bytes=0-9").response.status).toBe(404)
	})

	it("closes the port of an id registered again", async () => {
		const first = await registerPort("p1", 100)

		await registerPort("p1", 100)
		await tick()

		expect(first.lifetime.signal.aborted).toBe(true)
		expect(get("p1", "bytes=0-9").response.status).toBe(206)
	})

	it("answers 404 once the port's serving side has gone, so the page registers again", async () => {
		const { lifetime } = await registerPort("p1", 100)

		lifetime.abort()
		await tick()

		expect(get("p1", "bytes=0-9").response.status).toBe(404)
	})

	it("tells its pages a stream is gone when it answers 404, once per id per interval", async () => {
		const { lifetime } = await registerPort("p1", 100)

		lifetime.abort()
		await tick()

		const first = get("p1", "bytes=0-9")

		await first.settled
		// WebKit's retry storm: the same 404, many times a second.
		for (let attempt = 0; attempt < 5; attempt++) {
			await get("p1", "bytes=0-9").settled
		}

		await get("unknown", "bytes=0-1").settled

		expect(first.response.status).toBe(404)
		expect(posted).toEqual([
			{ type: SW_MSG_STREAM_GONE, id: "p1" },
			{ type: SW_MSG_STREAM_GONE, id: "unknown" }
		])
	})

	it("errors a response whose port's serving side dies mid-stream", async () => {
		const { lifetime } = await registerPort("p1", 100, "video/mp4", true)
		const { response, settled } = get("p1", "bytes=0-9")
		const body = response.arrayBuffer()

		await tick()
		lifetime.abort()

		await expect(body).rejects.toThrow()
		await settled
		expect(get("p1", "bytes=0-9").response.status).toBe(404)
	})

	it("refuses a registration without a port or with a malformed size", async () => {
		expect(await message({ type: SW_MSG_REGISTER_PORT_PREVIEW, id: "p1", name: "a", size: 1, contentType: "video/mp4" })).toMatchObject(
			{
				ok: false
			}
		)

		const channel = new MessageChannel()

		channels.push(channel)
		expect(
			await message({
				type: SW_MSG_REGISTER_PORT_PREVIEW,
				id: "p1",
				name: "a",
				size: -1,
				contentType: "video/mp4",
				port: channel.port2
			})
		).toMatchObject({ ok: false })
		expect(get("p1", "bytes=0-0").response.status).toBe(404)
	})
})

describe("sw.ts port registration order", () => {
	it("refuses a registration older than the one it holds, and keeps serving the newer one", async () => {
		const newer = await registerPort("p1", 100, "video/mp4", false, 2)
		const older = await registerPort("p1", 100, "video/mp4", false, 1)

		expect(newer.ack).toEqual({ ok: true })
		expect(older.ack).toEqual({ ok: false, error: SW_ERROR_SUPERSEDED })

		const { response } = get("p1", "bytes=0-9")

		expect(await bodyOf(response)).toEqual(expected(0, 9))
		expect(newer.reads).toEqual([[0, 10]])
		expect(older.reads).toEqual([])
		expect(newer.lifetime.signal.aborted).toBe(false)
	})

	it("refuses a registration a release already covered, and admits a newer one", async () => {
		listeners.get("message")?.({ data: { type: SW_MSG_RELEASE_PREVIEW, id: "p1", generation: 3 }, ports: [] })

		expect((await registerPort("p1", 100, "video/mp4", false, 3)).ack).toEqual({ ok: false, error: SW_ERROR_SUPERSEDED })
		expect(get("p1", "bytes=0-9").response.status).toBe(404)
		expect((await registerPort("p1", 100, "video/mp4", false, 4)).ack).toEqual({ ok: true })
		expect(get("p1", "bytes=0-9").response.status).toBe(206)
	})

	it("spares a registration newer than the release, and drops one it covers", async () => {
		const held = await registerPort("p1", 100, "video/mp4", false, 5)

		listeners.get("message")?.({ data: { type: SW_MSG_RELEASE_PREVIEW, id: "p1", generation: 4 }, ports: [] })

		const spared = get("p1", "bytes=0-9")

		expect(spared.response.status).toBe(206)
		await bodyOf(spared.response)
		await spared.settled

		listeners.get("message")?.({ data: { type: SW_MSG_RELEASE_PREVIEW, id: "p1", generation: 5 }, ports: [] })
		await tick()
		expect(held.lifetime.signal.aborted).toBe(true)
		expect(get("p1", "bytes=0-9").response.status).toBe(404)
	})

	it("refuses a malformed generation", async () => {
		const channel = new MessageChannel()

		channels.push(channel)
		expect(
			await message({
				type: SW_MSG_REGISTER_PORT_PREVIEW,
				id: "p1",
				name: "a",
				size: 1,
				contentType: "video/mp4",
				generation: -1,
				port: channel.port2
			})
		).toMatchObject({ ok: false })
	})
})

describe("sw.ts preview status", () => {
	const status = (id: string) => message({ type: SW_MSG_PREVIEW_STATUS, id })

	it("answers present, lost and missing from the registry, without reading through the port", async () => {
		const { lifetime, reads } = await registerPort("p1", 100)

		expect(await status("p1")).toEqual({ status: "present" })
		expect(await status("unknown")).toEqual({ status: "missing" })
		expect(reads).toEqual([])

		lifetime.abort()
		await tick()

		expect(await status("p1")).toEqual({ status: "lost" })

		listeners.get("message")?.({ data: { type: SW_MSG_RELEASE_PREVIEW, id: "p1" }, ports: [] })
		expect(await status("p1")).toEqual({ status: "missing" })
	})
})

describe("sw.ts version", () => {
	it("reports its protocol version", async () => {
		expect(SW_PROTOCOL_VERSION).toBe(18)

		let response: Response | null = null

		listeners.get("fetch")?.({
			request: { url: "https://app.test/__sw/version", method: "GET", mode: "cors", headers: new Headers() },
			respondWith: (value: Response) => {
				response = value
			}
		})

		expect(await (response as Response | null)?.json()).toMatchObject({ v: 18 })
	})
})
