import { describe, expect, it } from "vitest"
import { GenerationGate } from "@/lib/sw/generationGate"

describe("GenerationGate", () => {
	it("admits only a generation newer than any seen for the id", () => {
		const gate = new GenerationGate(8)

		expect(gate.admit("a", 2)).toBe(true)
		expect(gate.admit("a", 1)).toBe(false)
		expect(gate.admit("a", 2)).toBe(false)
		expect(gate.admit("a", 3)).toBe(true)
		expect(gate.admit("b", 1)).toBe(true)
	})

	it("refuses what a release covered, and never lowers the mark", () => {
		const gate = new GenerationGate(8)

		gate.retire("a", 4)
		expect(gate.admit("a", 4)).toBe(false)
		expect(gate.admit("a", 6)).toBe(true)
		gate.retire("a", 5)
		expect(gate.admit("a", 6)).toBe(false)
		expect(gate.admit("a", 7)).toBe(true)
	})

	it("forgets the least recently touched id past its bound", () => {
		const gate = new GenerationGate(2)

		gate.admit("a", 5)
		gate.admit("b", 5)
		gate.admit("a", 6)
		gate.admit("c", 5)

		expect(gate.admit("a", 6)).toBe(false)
		expect(gate.admit("b", 1)).toBe(true)
	})
})
