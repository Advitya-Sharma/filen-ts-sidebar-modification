import { describe, expect, it } from "vitest"
import { keepItem, PumpGate, runPump, type PumpItem, type PumpReader, type PumpWriter } from "@/workers/remux/pump"

interface Packet extends PumpItem {
	index: number
}

// Packets in decode order, as a demuxer hands them out.
function track(packets: Omit<Packet, "index">[]): Packet[] {
	return packets.map((packet, index) => ({ ...packet, index }))
}

function reader(packets: Packet[], first: (time: number) => Packet | null): PumpReader<Packet> & { reads: number } {
	const state = {
		reads: 0,
		first: (time: number) => {
			state.reads++

			return Promise.resolve(first(time))
		},
		next: (packet: Packet) => {
			state.reads++

			return Promise.resolve(packets[packet.index + 1] ?? null)
		}
	}

	return state
}

// Cuts a fragment whenever a key packet arrives at least `fragment` seconds after the current one began,
// the way the fragmented-MP4 muxer does.
function writer(gate: PumpGate, fragment = 1): PumpWriter<Packet> & { written: Packet[]; ended: boolean } {
	let fragmentStart: number | null = null
	const state = {
		written: [] as Packet[],
		ended: false,
		write: (packet: Packet) => {
			if (fragmentStart === null) {
				expect(packet.type).toBe("key")
				fragmentStart = packet.timestamp
			} else if (packet.type === "key" && packet.timestamp - fragmentStart >= fragment) {
				gate.noteFlushed(packet.timestamp)
				fragmentStart = packet.timestamp
			}

			state.written.push(packet)

			return Promise.resolve()
		},
		end: () => {
			state.ended = true
			gate.noteFlushed(Number.POSITIVE_INFINITY)

			return Promise.resolve()
		}
	}

	return state
}

// An open GOP: the CRA key at 2.0 is followed in decode order by leading pictures presented before it.
const VIDEO = track([
	{ timestamp: 0, duration: 0.5, type: "key" },
	{ timestamp: 1, duration: 0.5, type: "delta" },
	{ timestamp: 0.5, duration: 0.5, type: "delta" },
	{ timestamp: 1.5, duration: 0.5, type: "delta" },
	{ timestamp: 2, duration: 0.5, type: "key" },
	{ timestamp: 1.75, duration: 0.25, type: "delta" },
	{ timestamp: 3, duration: 0.5, type: "delta" },
	{ timestamp: 2.5, duration: 0.5, type: "delta" },
	{ timestamp: 3.5, duration: 0.5, type: "delta" },
	{ timestamp: 4, duration: 0.5, type: "key" },
	{ timestamp: 5, duration: 0.5, type: "delta" },
	{ timestamp: 4.5, duration: 0.5, type: "delta" },
	{ timestamp: 5.5, duration: 0.5, type: "delta" }
])

const keyAtOrBefore = (time: number) => VIDEO.filter(packet => packet.type === "key" && packet.timestamp <= time).at(-1) ?? null

describe("keepItem", () => {
	it("keeps what is presented from the start on", () => {
		expect(keepItem({ timestamp: 2, duration: 0, type: "key" }, 2)).toBe(true)
		expect(keepItem({ timestamp: 1.75, duration: 0, type: "delta" }, 2)).toBe(false)
	})
})

describe("runPump", () => {
	it("starts at the key packet at or before the target and reports its timestamp", async () => {
		const gate = new PumpGate(Number.POSITIVE_INFINITY)
		const out = writer(gate)
		let started: number | null = null
		const result = await runPump(reader(VIDEO, keyAtOrBefore), out, gate, 2.7, {
			floor: Number.NEGATIVE_INFINITY,
			onStart: time => {
				started = time
			}
		})

		expect(started).toBe(2)
		expect(result).toEqual({ type: "ended", startTimestamp: 2 })
		expect(out.written[0]?.timestamp).toBe(2)
		expect(out.ended).toBe(true)
	})

	it("drops leading pictures presented before the key packet, keeps source timestamps", async () => {
		const gate = new PumpGate(Number.POSITIVE_INFINITY)
		const out = writer(gate)

		await runPump(reader(VIDEO, keyAtOrBefore), out, gate, 2, { floor: Number.NEGATIVE_INFINITY })

		expect(out.written.map(packet => packet.timestamp)).toEqual([2, 3, 2.5, 3.5, 4, 5, 4.5, 5.5])
		expect(Math.min(...out.written.map(packet => packet.timestamp))).toBe(2)
	})

	it("keeps B-frames reordered after the key packet", async () => {
		const gate = new PumpGate(Number.POSITIVE_INFINITY)
		const out = writer(gate)

		await runPump(reader(VIDEO, keyAtOrBefore), out, gate, 0, { floor: Number.NEGATIVE_INFINITY })

		expect(out.written.slice(0, 4).map(packet => packet.timestamp)).toEqual([0, 1, 0.5, 1.5])
	})

	it("drops audio priming before zero", async () => {
		const audio = track([
			{ timestamp: -0.02, duration: 0.02, type: "key" },
			{ timestamp: 0, duration: 0.02, type: "key" },
			{ timestamp: 0.02, duration: 0.02, type: "key" }
		])
		const gate = new PumpGate(Number.POSITIVE_INFINITY)
		const out = writer(gate)
		const result = await runPump(
			reader(audio, () => audio[0] ?? null),
			out,
			gate,
			0,
			{ floor: 0 }
		)

		expect(result).toEqual({ type: "ended", startTimestamp: 0 })
		expect(out.written.map(packet => packet.timestamp)).toEqual([0, 0.02])
	})

	it("stops reading once a fragment reaching the horizon is out, and goes on when the horizon moves", async () => {
		const gate = new PumpGate(2)
		const out = writer(gate)
		const source = reader(VIDEO, keyAtOrBefore)
		const pumping = runPump(source, out, gate, 0, { floor: Number.NEGATIVE_INFINITY })

		await new Promise(resolve => setTimeout(resolve, 0))

		// The key at 2.0 cut the first fragment; nothing past it was read.
		expect(gate.flushedUntil).toBe(2)
		expect(out.written.at(-1)?.timestamp).toBe(2)
		expect(source.reads).toBe(5)

		gate.setHorizon(Number.POSITIVE_INFINITY)

		await pumping

		expect(out.written).toHaveLength(VIDEO.length)
	})

	it("stops at once when cancelled while waiting for room", async () => {
		const gate = new PumpGate(1)
		const out = writer(gate)
		const pumping = runPump(reader(VIDEO, keyAtOrBefore), out, gate, 0, { floor: Number.NEGATIVE_INFINITY })

		await new Promise(resolve => setTimeout(resolve, 0))
		gate.cancel()

		expect(await pumping).toEqual({ type: "cancelled" })
		expect(out.ended).toBe(false)
	})

	it("ends an empty track without writing", async () => {
		const gate = new PumpGate(Number.POSITIVE_INFINITY)
		const out = writer(gate)

		expect(
			await runPump(
				reader([], () => null),
				out,
				gate,
				0,
				{ floor: 0 }
			)
		).toEqual({ type: "ended", startTimestamp: null })
		expect(out.written).toHaveLength(0)
		expect(out.ended).toBe(true)
	})

	it("hands every item read but not written to the discard hook", async () => {
		const gate = new PumpGate(Number.POSITIVE_INFINITY)
		const out = writer(gate)
		const discarded: number[] = []

		await runPump(reader(VIDEO, keyAtOrBefore), out, gate, 2, {
			floor: Number.NEGATIVE_INFINITY,
			discard: item => {
				discarded.push(item.timestamp)
			}
		})

		expect(discarded).toEqual([1.75])
	})

	it("discards the item it read when cancelled during the read", async () => {
		const gate = new PumpGate(Number.POSITIVE_INFINITY)
		const out = writer(gate)
		const discarded: number[] = []
		const source: PumpReader<Packet> = {
			first: time => Promise.resolve(keyAtOrBefore(time)),
			next: packet => {
				gate.cancel()

				return Promise.resolve(VIDEO[packet.index + 1] ?? null)
			}
		}

		expect(
			await runPump(source, out, gate, 0, {
				floor: Number.NEGATIVE_INFINITY,
				discard: item => {
					discarded.push(item.timestamp)
				}
			})
		).toEqual({ type: "cancelled" })
		expect(discarded).toEqual([1])
	})
})
