import { describe, expect, it } from "vitest"
import {
	UserDictionaryVersionError,
	addUserWord,
	emptyUserDictionary,
	hasUserWord,
	mergeUserDictionaries,
	normalizeUserWord,
	parseUserDictionary,
	removeUserWord,
	sameUserDictionary,
	serializeUserDictionary,
	userWords
} from "../userDictionary"

describe("user dictionary", () => {
	it("adds, removes and lists words", () => {
		let doc = addUserWord(emptyUserDictionary(), "Filen", 10)

		doc = addUserWord(doc, "zorp", 11)
		expect(userWords(doc)).toEqual(["Filen", "zorp"])
		expect(addUserWord(doc, "zorp", 50)).toBe(doc)

		doc = removeUserWord(doc, "zorp", 12)
		expect(userWords(doc)).toEqual(["Filen"])
		expect(removeUserWord(doc, "zorp", 13)).toBe(doc)

		// Re-added after a removal stamped later than this device's clock.
		doc = addUserWord(doc, "zorp", 5)
		expect(hasUserWord(doc, "zorp")).toBe(true)
	})

	it("normalizes and rejects words", () => {
		expect(normalizeUserWord("  Café ")).toBe("Café")
		expect(normalizeUserWord("")).toBeUndefined()
		expect(normalizeUserWord("two words")).toBeUndefined()
		expect(normalizeUserWord("x".repeat(101))).toBeUndefined()
		expect(addUserWord(emptyUserDictionary(), "a b", 1)).toEqual(emptyUserDictionary())
	})

	it("merges devices without losing additions or resurrecting removals", () => {
		const base = addUserWord(addUserWord(emptyUserDictionary(), "alpha", 1), "beta", 1)
		const phone = removeUserWord(addUserWord(base, "gamma", 5), "alpha", 6)
		const laptop = addUserWord(removeUserWord(base, "beta", 4), "delta", 7)
		const merged = mergeUserDictionaries(phone, laptop)

		expect(userWords(merged)).toEqual(["delta", "gamma"])
		expect(sameUserDictionary(merged, mergeUserDictionaries(laptop, phone))).toBe(true)
		expect(sameUserDictionary(mergeUserDictionaries(merged, phone), merged)).toBe(true)
	})

	it("serializes deterministically and parses back", () => {
		const doc = addUserWord(addUserWord(emptyUserDictionary(), "b", 2), "a", 1)
		const text = serializeUserDictionary(doc)

		expect(text).toBe('{"v":1,"words":{"a":1,"b":2},"removed":{}}')
		expect(sameUserDictionary(parseUserDictionary(text), doc)).toBe(true)
	})

	it("drops invalid entries and keeps hostile keys inert", () => {
		const doc = parseUserDictionary('{"v":1,"words":{"ok":1,"bad word":2,"__proto__":3,"nan":"x","inf":1e999},"removed":[]}')

		expect(userWords(doc)).toEqual(["__proto__", "ok"])
		expect(Object.getPrototypeOf(doc.words)).toBeNull()
		expect(serializeUserDictionary(doc)).toBe('{"v":1,"words":{"__proto__":3,"ok":1},"removed":{}}')
		expect(userWords(parseUserDictionary("[]"))).toEqual([])
	})

	it("clamps stamps to safe integers", () => {
		const doc = parseUserDictionary('{"v":1,"words":{"a":1e300,"b":-1e300,"c":1.5},"removed":{}}')

		expect(doc.words["a"]).toBe(Number.MAX_SAFE_INTEGER)
		expect(doc.words["b"]).toBe(-Number.MAX_SAFE_INTEGER)
		expect(doc.words["c"]).toBe(1)
		expect(removeUserWord(doc, "a", 5).removed["a"]).toBe(Number.MAX_SAFE_INTEGER)
	})

	it("refuses a newer format and invalid JSON", () => {
		expect(() => parseUserDictionary('{"v":2,"words":{}}')).toThrow(UserDictionaryVersionError)
		expect(() => parseUserDictionary("{")).toThrow(SyntaxError)
	})
})
