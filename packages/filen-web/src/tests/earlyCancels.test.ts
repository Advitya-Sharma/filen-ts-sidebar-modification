import { describe, expect, it } from "vitest"
import { EarlyCancels } from "@/workers/earlyCancels"

describe("EarlyCancels", () => {
	it("remembers a cancelled token for its read, once, and only for a minute", () => {
		let now = 0
		const cancels = new EarlyCancels(() => now)

		cancels.add("a")
		cancels.add("b")

		expect(cancels.take("a")).toBe(true)
		expect(cancels.take("a")).toBe(false)
		expect(cancels.take("never")).toBe(false)

		now = 61_000
		expect(cancels.take("b")).toBe(false)
	})

	it("forgets the oldest tokens past its bound", () => {
		const cancels = new EarlyCancels(() => 0)

		for (let i = 0; i <= 1024; i++) {
			cancels.add(String(i))
		}

		expect(cancels.take("0")).toBe(false)
		expect(cancels.take("1024")).toBe(true)
	})
})
