import { describe, expect, it } from "vitest"
import {
	DEFAULT_OFFICE_AUTOSAVE,
	MAX_AUTHOR_NAME_LENGTH,
	OfficeSettingsVersionError,
	emptyOfficeSettings,
	mergeOfficeSettings,
	normalizeAuthorName,
	officeAuthorName,
	officeAutosave,
	parseOfficeSettings,
	sameOfficeSettings,
	serializeOfficeSettings,
	setAuthorName,
	setOfficeAutosave
} from "@filen/shared"

describe("office settings", () => {
	it("starts unset: no author name, autosave on", () => {
		const doc = emptyOfficeSettings()

		expect(officeAuthorName(doc)).toBeNull()
		expect(officeAutosave(doc)).toBe(DEFAULT_OFFICE_AUTOSAVE)
		expect(DEFAULT_OFFICE_AUTOSAVE).toBe(true)
	})

	it("normalises author names and refuses empty, overlong or control-character ones", () => {
		expect(normalizeAuthorName("  Jane \n  Doe ")).toBe("Jane Doe")
		expect(normalizeAuthorName("Café")).toBe("Café")
		expect(normalizeAuthorName("   ")).toBeUndefined()
		expect(normalizeAuthorName("x".repeat(MAX_AUTHOR_NAME_LENGTH + 1))).toBeUndefined()
		expect(normalizeAuthorName("a\u0000b")).toBeUndefined()
	})

	it("round-trips through its stable text", () => {
		const doc = setOfficeAutosave(setAuthorName(emptyOfficeSettings(), "Jane", 10), false, 20)
		const text = serializeOfficeSettings(doc)

		expect(parseOfficeSettings(text)).toEqual(doc)
		expect(serializeOfficeSettings(parseOfficeSettings(text))).toBe(text)
		expect(officeAuthorName(doc)).toBe("Jane")
		expect(officeAutosave(doc)).toBe(false)
	})

	it("drops invalid settings and refuses a newer format", () => {
		const doc = parseOfficeSettings(JSON.stringify({ v: 1, authorName: { value: 5, at: 1 }, autosave: { value: true, at: "x" } }))

		expect(doc).toEqual(emptyOfficeSettings())
		expect(parseOfficeSettings("[]")).toEqual(emptyOfficeSettings())
		expect(() => parseOfficeSettings(JSON.stringify({ v: 2 }))).toThrow(OfficeSettingsVersionError)
		expect(() => parseOfficeSettings("{")).toThrow(SyntaxError)
	})

	it("merges per setting, the later change winning, and the same way on every device", () => {
		const a = setAuthorName(emptyOfficeSettings(), "Ann", 100)
		const b = setOfficeAutosave(setAuthorName(emptyOfficeSettings(), "Bob", 200), false, 50)
		const merged = mergeOfficeSettings(a, b)

		expect(officeAuthorName(merged)).toBe("Bob")
		expect(officeAutosave(merged)).toBe(false)
		expect(sameOfficeSettings(merged, mergeOfficeSettings(b, a))).toBe(true)

		// A tie resolves by value, whichever side is local.
		const tieA = setAuthorName(emptyOfficeSettings(), "Ann", 5)
		const tieB = setAuthorName(emptyOfficeSettings(), "Bob", 5)

		expect(sameOfficeSettings(mergeOfficeSettings(tieA, tieB), mergeOfficeSettings(tieB, tieA))).toBe(true)
	})

	it("changes nothing for the same value, and outbids a register from a clock ahead", () => {
		const doc = setAuthorName(emptyOfficeSettings(), "Jane", 1000)

		expect(setAuthorName(doc, " Jane ", 2000)).toBe(doc)
		expect(setAuthorName(doc, "", 2000)).toBe(doc)

		const behind = setAuthorName(doc, "Joan", 10)

		expect(behind.authorName?.at).toBe(1001)
		expect(officeAuthorName(mergeOfficeSettings(doc, behind))).toBe("Joan")

		const autosave = setOfficeAutosave(doc, true, 5)

		expect(autosave.autosave).toEqual({ value: true, at: 5 })
		expect(setOfficeAutosave(autosave, true, 50)).toBe(autosave)
	})
})
