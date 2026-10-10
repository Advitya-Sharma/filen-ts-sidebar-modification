import { describe, expect, it, vi } from "vitest"
import type { RemuxEvent, RemuxOpenParams } from "@/lib/media/remuxTypes"
import type { RemuxWorkerApi } from "@/workers/remux.worker"

// The remux worker's warm hand-over against a stub engine: a warm that lands after its session gave up on it
// builds nothing.
const { exposed, engines } = vi.hoisted(() => ({
	exposed: new Map<"api", unknown>(),
	engines: [] as unknown[]
}))

vi.mock("comlink", () => ({
	expose: (api: unknown) => {
		exposed.set("api", api)
	},
	transfer: (value: unknown) => value,
	releaseProxy: Symbol("releaseProxy")
}))

vi.mock("@/workers/remux/engine", () => ({
	RemuxEngine: class {
		public constructor() {
			engines.push(this)
		}

		public open(): Promise<unknown> {
			return Promise.resolve({})
		}

		public canStartAt(): Promise<boolean> {
			return Promise.resolve(true)
		}

		public close(): Promise<void> {
			return Promise.resolve()
		}
	}
}))

await import("@/workers/remux.worker")

const api = exposed.get("api") as RemuxWorkerApi

function params(): RemuxOpenParams & { port: MessagePort } {
	return {
		port: new MessageChannel().port1,
		size: 1000,
		transcodeTargets: [],
		maxFragmentBytes: { video: 1_000_000, audio: 1_000_000 }
	}
}

describe("remux worker warm", () => {
	it("drops a warm that lands after its session opened the file itself", async () => {
		await expect(api.open({ warmKey: "late" }, vi.fn<(event: RemuxEvent) => void>())).rejects.toMatchObject({ name: "WarmGoneError" })

		const before = engines.length
		const late = params()
		const close = vi.spyOn(late.port, "close")

		await api.warm(late, "late")

		expect(engines).toHaveLength(before)
		expect(close).toHaveBeenCalledOnce()
	})

	it("keeps a warm that lands first for the session to take over", async () => {
		const before = engines.length

		await api.warm(params(), "early")

		expect(engines).toHaveLength(before + 1)

		const taken = await api.open({ warmKey: "early" }, vi.fn<(event: RemuxEvent) => void>())

		expect(taken.id).toBeTypeOf("number")
		expect(engines).toHaveLength(before + 1)
		await api.close(taken.id)
	})
})
