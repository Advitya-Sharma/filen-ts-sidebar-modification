import { describe, expect, it, vi } from "vitest"
import * as z from "zod"
import { DEFAULT_MODEL_BASE_URL } from "@/features/assistant/runtime/manifest"

// Counts zod's eval probe (`new Function("")`), which the app's CSP reports as a violation even though zod
// catches the throw. The manifest module can evaluate before the shared tool schemas, so it sets jitless too.
const probes = vi.hoisted(() => {
	const seen = { count: 0 }

	vi.stubGlobal(
		"Function",
		new Proxy(Function, {
			construct(target, args, newTarget) {
				if (args.length === 1 && args[0] === "") {
					seen.count++
				}

				return Reflect.construct(target, args, newTarget) as object
			}
		})
	)

	return seen
})

describe("assistant model manifest zod config", () => {
	it("builds its schemas without zod's eval probe", () => {
		expect(DEFAULT_MODEL_BASE_URL).toMatch(/^https:/)
		expect(z.config().jitless).toBe(true)
		expect(z.strictObject({ a: z.string() }).parse({ a: "x" })).toEqual({ a: "x" })
		expect(probes.count).toBe(0)
	})
})
