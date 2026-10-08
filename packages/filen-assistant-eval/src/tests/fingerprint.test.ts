import { describe, expect, it } from "vitest"
import { caseFingerprint, parseRecordLine } from "../fingerprint"

describe("caseFingerprint", () => {
	it("changes with the case and the model, so a stale result is measured again", () => {
		const item = { id: "a", prompt: "Find my tax documents" }

		expect(caseFingerprint({ modelId: "m", item })).toBe(caseFingerprint({ modelId: "m", item: { ...item } }))
		expect(caseFingerprint({ modelId: "m", item: { ...item, prompt: "Find my taxes" } })).not.toBe(
			caseFingerprint({ modelId: "m", item })
		)
		expect(caseFingerprint({ modelId: "other", item })).not.toBe(caseFingerprint({ modelId: "m", item }))
	})

	it("skips a record line a crash cut short", () => {
		const line = JSON.stringify({ caseId: "a", transcript: "text" })

		expect(parseRecordLine(line)).toEqual({ caseId: "a", transcript: "text" })
		expect(parseRecordLine(line.slice(0, 20))).toBeNull()
	})
})
