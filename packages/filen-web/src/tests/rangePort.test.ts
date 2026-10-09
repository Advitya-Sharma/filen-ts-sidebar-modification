import { afterEach, describe, expect, it, vi } from "vitest"
import { exposeRangePort, RANGE_PORT_IDLE_MS, RANGE_READ_TIMEOUT_MS, RangePortClosedError, rangePortReader } from "@/lib/media/rangePort"

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

	it("shuts down a port nobody has read from for the idle period", () => {
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
})
