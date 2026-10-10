import { describe, expect, it, vi } from "vitest"
import { PumpGate } from "@/workers/remux/pump"

// Where video starts only at IDR pictures (Firefox with HEVC), the walk back to one is one read per key
// packet: a seek's canStartAt and the start that follows at the same time walk it once between them.

interface FakePacket {
	timestamp: number
	data: Uint8Array
}

// One length-prefixed NAL unit of `type`: 19 is an IDR picture, 21 a CRA.
const nal = (type: number): Uint8Array => new Uint8Array([0, 0, 0, 2, type << 1, 1])
// An IDR at 0, CRA keys every 2 s after it.
const KEYS: FakePacket[] = [0, 2, 4, 6, 8].map(timestamp => ({ timestamp, data: nal(timestamp === 0 ? 19 : 21) }))

const reads = vi.hoisted(() => ({ keyPackets: 0 }))

vi.mock("mediabunny", async importOriginal => ({
	...(await importOriginal<typeof import("mediabunny")>()),
	EncodedPacketSink: class {
		public getKeyPacket(timestamp: number): Promise<FakePacket | null> {
			reads.keyPackets++

			return Promise.resolve(KEYS.findLast(key => key.timestamp <= timestamp) ?? null)
		}

		public getFirstKeyPacket(): Promise<FakePacket | null> {
			return Promise.resolve(KEYS[0] ?? null)
		}
	}
}))

vi.mock("@/lib/media/rangePort", () => ({
	rangePortReader: () => ({ readRange: vi.fn(), onLost: vi.fn(), close: vi.fn() })
}))

const { RemuxEngine } = await import("@/workers/remux/engine")
const { EncodedPacketSink } = await import("mediabunny")

interface EngineInternals {
	video: { track: unknown; idrOnly: { lengthSize: number | null } | null }
	duration: number
	startKey(sink: unknown, at: number, video: EngineInternals["video"], gate: PumpGate): Promise<FakePacket | null>
}

describe("RemuxEngine IDR starts", () => {
	it("walks back to the IDR once for a canStartAt and the start at the same time", async () => {
		const engine = new RemuxEngine(
			{
				port: new MessageChannel().port1,
				size: 1000,
				transcodeTargets: [],
				maxFragmentBytes: { video: 1_000_000, audio: 1_000_000 },
				idrOnlyStart: true
			},
			() => undefined
		)
		const internals = engine as unknown as EngineInternals

		internals.video = { track: {}, idrOnly: { lengthSize: 4 } }
		internals.duration = 100

		expect(await engine.canStartAt(9)).toBe(true)

		const walk = reads.keyPackets

		expect(walk).toBe(5)

		const start = await internals.startKey(new EncodedPacketSink({} as never), 9, internals.video, new PumpGate(0))

		expect(start?.timestamp).toBe(0)
		expect(reads.keyPackets - walk).toBe(1)

		// Another time walks again.
		const other = await internals.startKey(new EncodedPacketSink({} as never), 7, internals.video, new PumpGate(0))

		expect(other?.timestamp).toBe(0)
		expect(reads.keyPackets - walk - 1).toBe(4)
	})
})
