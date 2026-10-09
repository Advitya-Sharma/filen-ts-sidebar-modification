import * as Comlink from "comlink"
import { afterEach, describe, expect, it, vi } from "vitest"
import {
	exposeRangePort,
	RANGE_PORT_IDLE_MS,
	RANGE_READ_TIMEOUT_MS,
	RangePortClosedError,
	rangePortReader,
	type RangePortApi
} from "@/lib/media/rangePort"

const channels: MessageChannel[] = []

afterEach(() => {
	for (const channel of channels.splice(0)) {
		channel.port1.close()
		channel.port2.close()
	}

	vi.useRealTimers()
})

function setUp(read: (start: number, end: number, signal: AbortSignal) => Promise<Uint8Array>) {
	const channel = new MessageChannel()
	const lifetime = new AbortController()

	channels.push(channel)
	exposeRangePort(channel.port1, read, lifetime)

	return { reader: rangePortReader(channel.port2), lifetime }
}

function never(signal: AbortSignal, aborted?: () => void): Promise<Uint8Array> {
	return new Promise((_resolve, reject) => {
		signal.addEventListener("abort", () => {
			aborted?.()
			reject(new Error("aborted on the serving side"))
		})
	})
}

describe("range port", () => {
	it("serves byte ranges", async () => {
		const { reader } = setUp((start, end) => Promise.resolve(Uint8Array.from({ length: end - start }, (_, index) => start + index)))

		expect(Array.from(await reader.readRange(3, 6, new AbortController().signal))).toEqual([3, 4, 5])
	})

	it("rejects a read at once when its signal aborts, and cancels it on the serving side", async () => {
		const served = Promise.withResolvers<undefined>()
		const { reader } = setUp((_start, _end, signal) =>
			never(signal, () => {
				served.resolve(undefined)
			})
		)
		const controller = new AbortController()
		const read = reader.readRange(0, 10, controller.signal)

		controller.abort()

		await expect(read).rejects.toThrow()
		await served.promise
	})

	it("fails reads in flight and every later one once the serving side shuts down", async () => {
		const { reader, lifetime } = setUp((_start, _end, signal) => never(signal))
		const read = reader.readRange(0, 10, new AbortController().signal)

		await new Promise(resolve => setTimeout(resolve, 10))
		lifetime.abort()

		await expect(read).rejects.toBeInstanceOf(RangePortClosedError)
		await expect(reader.readRange(0, 10, new AbortController().signal)).rejects.toBeInstanceOf(RangePortClosedError)
	})

	it("tells the reader when the serving side goes away, but not after its own close", async () => {
		const lost = vi.fn()
		const { reader, lifetime } = setUp(() => Promise.resolve(new Uint8Array(1)))

		reader.onLost(lost)
		lifetime.abort()
		await new Promise(resolve => setTimeout(resolve, 20))

		expect(lost).toHaveBeenCalledTimes(1)

		const other = setUp(() => Promise.resolve(new Uint8Array(1)))
		const quiet = vi.fn()

		other.reader.onLost(quiet)
		other.reader.close()
		await new Promise(resolve => setTimeout(resolve, 20))

		expect(quiet).not.toHaveBeenCalled()
	})

	it("shuts the serving side down when the reader closes", async () => {
		const { reader, lifetime } = setUp(() => Promise.resolve(new Uint8Array(1)))

		reader.close()
		await new Promise(resolve => setTimeout(resolve, 20))

		expect(lifetime.signal.aborted).toBe(true)
	})

	it("shuts down a port nobody has read from for the idle period, while its reader is not known to hold on", () => {
		vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] })

		const { lifetime } = setUp(() => Promise.resolve(new Uint8Array(1)))

		vi.advanceTimersByTime(RANGE_PORT_IDLE_MS - 1)
		expect(lifetime.signal.aborted).toBe(false)
		vi.advanceTimersByTime(1)
		expect(lifetime.signal.aborted).toBe(true)
	})

	it("gives up on a read nobody answers", async () => {
		vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] })

		const { reader } = setUp((_start, _end, signal) => never(signal))
		const read = reader.readRange(0, 10, new AbortController().signal)

		vi.advanceTimersByTime(RANGE_READ_TIMEOUT_MS)

		await expect(read).rejects.toThrow("timed out")
	})

	it("learns at once that the serving worker died, though its port posts nothing", async () => {
		// A server whose life lock goes the way a terminated worker's does: released without any message.
		const channel = new MessageChannel()
		const died = Promise.withResolvers<undefined>()
		let lockName: string | null = null
		const server: RangePortApi = {
			read: () => new Promise<Uint8Array>(() => undefined),
			cancel: () => undefined,
			close: () => undefined,
			holdLife: name => {
				lockName = name
				const held = Promise.withResolvers<boolean>()

				void navigator.locks.request(name, () => {
					held.resolve(true)

					return died.promise
				})

				return held.promise
			}
		}

		channels.push(channel)
		Comlink.expose(server, channel.port1)

		const reader = rangePortReader(channel.port2)
		const lost = vi.fn()

		reader.onLost(lost)

		const read = reader.readRange(0, 10, new AbortController().signal)

		await new Promise(resolve => setTimeout(resolve, 20))
		expect(lockName).not.toBeNull()
		expect(lost).not.toHaveBeenCalled()

		died.resolve(undefined)

		await expect(read).rejects.toBeInstanceOf(RangePortClosedError)
		expect(lost).toHaveBeenCalledTimes(1)
	})

	it("holds the serving side's life lock until it shuts down", async () => {
		const names = async () => new Set(((await navigator.locks.query()).held ?? []).map(lock => lock.name))
		const before = await names()
		const { lifetime } = setUp(() => Promise.resolve(new Uint8Array(1)))

		await new Promise(resolve => setTimeout(resolve, 20))

		const added = [...(await names())].filter(name => name !== undefined && !before.has(name) && name.startsWith("filen-range-port-"))

		expect(added).toHaveLength(1)

		lifetime.abort()
		await new Promise(resolve => setTimeout(resolve, 20))

		expect((await names()).has(added[0])).toBe(false)
	})

	it("lives as long as a reader holding its lock, past the idle period, and ends when that reader goes", async () => {
		vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] })

		const { lifetime } = setUp(() => Promise.resolve(new Uint8Array(1)))
		const settle = async () => {
			for (let turn = 0; turn < 20; turn++) {
				await new Promise(resolve => setImmediate(resolve))
			}
		}

		await settle()
		vi.advanceTimersByTime(RANGE_PORT_IDLE_MS * 4)
		expect(lifetime.signal.aborted).toBe(false)

		// A reader whose worker dies releases its lock without closing anything.
		const channel = new MessageChannel()
		const served = new AbortController()
		const readerLock = `test-reader-${crypto.randomUUID()}`
		const gone = Promise.withResolvers<undefined>()
		const holding = Promise.withResolvers<undefined>()

		channels.push(channel)
		exposeRangePort(channel.port1, () => Promise.resolve(new Uint8Array(1)), served)
		void navigator.locks.request(readerLock, () => {
			holding.resolve(undefined)

			return gone.promise
		})
		await holding.promise

		const remote = Comlink.wrap<RangePortApi>(channel.port2)

		await expect(remote.holdLife(`test-server-${crypto.randomUUID()}`, readerLock)).resolves.toBe(true)
		vi.advanceTimersByTime(RANGE_PORT_IDLE_MS * 4)
		expect(served.signal.aborted).toBe(false)

		gone.resolve(undefined)
		await settle()
		expect(served.signal.aborted).toBe(true)
	})
})
