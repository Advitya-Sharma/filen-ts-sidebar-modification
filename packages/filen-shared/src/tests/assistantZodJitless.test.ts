import { describe, expect, it, vi } from "vitest"
import * as z from "zod"
import { ASSISTANT_TOOLS } from "@filen/shared/assistant"

// Counts zod's eval probe (`new Function("")`), which a strict CSP reports as a violation even though zod
// catches the throw. Installed before the imports above evaluate.
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

describe("assistant zod config", () => {
	it("builds and parses the tool schemas without zod's eval probe", () => {
		expect(ASSISTANT_TOOLS.length).toBeGreaterThan(0)
		expect(z.config().jitless).toBe(true)
		expect(z.strictObject({ a: z.string() }).parse({ a: "x" })).toEqual({ a: "x" })
		expect(probes.count).toBe(0)
	})
})
